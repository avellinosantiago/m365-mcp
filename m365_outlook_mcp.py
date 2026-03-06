#!/usr/bin/env python3
"""
M365 Outlook MCP Server
Provides tools for managing Outlook emails, drafts, and folders
via Outlook COM (win32com). Requires Outlook desktop to be running.

Run: python m365_outlook_mcp.py
"""

import asyncio
import json
from functools import partial
from typing import Optional

import pythoncom
import win32com.client
from mcp.server.fastmcp import FastMCP
from pydantic import BaseModel, ConfigDict, Field

# =============================================================================
# MCP SERVER
# =============================================================================

mcp = FastMCP("m365_outlook_mcp")

# =============================================================================
# COM HELPERS
# =============================================================================

# Default folder constants (OlDefaultFolders enum)
WELL_KNOWN_FOLDERS = {
    "inbox": 6,
    "drafts": 16,
    "sentitems": 5,
    "sent": 5,
    "deleteditems": 3,
    "deleted": 3,
    "junkemail": 23,
    "junk": 23,
    "outbox": 4,
}

STORE_FILTER = "savellino"


def _get_outlook():
    """Get Outlook COM objects. Must be called from a COM-initialized thread."""
    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    return outlook, namespace


def _get_store(namespace):
    """Find the correct mail store.

    Searches all Outlook stores for one matching STORE_FILTER,
    excluding Public Folders (supports English and Spanish names).
    Prefers exact email match over partial display name match.
    """
    PUBLIC_KEYWORDS = ("public", "pública", "públicas", "publica", "publicas")
    available = []
    candidates = []
    for s in namespace.Stores:
        display = s.DisplayName
        available.append(display)
        display_lower = display.lower()
        if STORE_FILTER not in display_lower:
            continue
        if any(kw in display_lower for kw in PUBLIC_KEYWORDS):
            continue
        candidates.append(s)

    _get_store._last_available = available

    if not candidates:
        return None

    # Prefer exact email match (e.g. "savellino@megalabs.global")
    for s in candidates:
        if "@" in s.DisplayName:
            return s

    return candidates[0]


def _resolve_folder(namespace, folder_id: str):
    """Resolve a folder by well-known name, path, exact name, or partial match.

    Resolution order:
    1. Well-known names: inbox, drafts, sentitems, deleteditems, junkemail, outbox
    2. Path syntax: "parent/child" navigates the tree (e.g. "1. PROYECTOS/EVOCON")
    3. Exact name match: case-insensitive search in top-level and one-level-deep folders
    4. Partial match: if the search term is contained in the folder name (e.g. "EVOCON"
       matches "1. PROYECTOS" subfolder "EVOCON", or "REGULATORIO" matches "3. REGULATORIO")
    """
    lower = folder_id.lower().strip()

    # 1) Well-known folder
    if lower in WELL_KNOWN_FOLDERS:
        return namespace.GetDefaultFolder(WELL_KNOWN_FOLDERS[lower])

    store = _get_store(namespace)
    if not store:
        return None

    root = store.GetRootFolder()

    # 2) Path syntax: "parent/child" or "parent/child/grandchild"
    if "/" in folder_id:
        parts = [p.strip() for p in folder_id.split("/") if p.strip()]
        current = root
        for part in parts:
            part_lower = part.lower()
            found = None
            for f in current.Folders:
                if f.Name.lower() == part_lower:
                    found = f
                    break
            if not found:
                # Try partial match within this level
                for f in current.Folders:
                    if part_lower in f.Name.lower():
                        found = f
                        break
            if not found:
                return None
            current = found
        return current

    # 3) Exact name match (top-level + one level deep)
    for f in root.Folders:
        if f.Name.lower() == lower:
            return f
        for sub in f.Folders:
            if sub.Name.lower() == lower:
                return sub

    # 4) Partial/contains match (top-level first, then one level deep)
    for f in root.Folders:
        if lower in f.Name.lower():
            return f
    for f in root.Folders:
        for sub in f.Folders:
            if lower in sub.Name.lower():
                return sub

    return None


def _get_sender_email(item):
    """Get sender email, resolving Exchange DN if needed."""
    addr = item.SenderEmailAddress
    if addr and addr.startswith("/O="):
        # Exchange DN - try to resolve via SenderName or PropertyAccessor
        try:
            sender = item.Sender
            if sender:
                user = sender.GetExchangeUser()
                if user:
                    return user.PrimarySmtpAddress
        except Exception:
            pass
        return item.SenderName
    return addr or item.SenderName or ""


def _get_recipients(item, prop="To"):
    """Get recipient emails from a mail item."""
    result = []
    try:
        recipients = item.Recipients
        for i in range(1, recipients.Count + 1):
            r = recipients.Item(i)
            # Type 1=To, 2=CC, 3=BCC
            target_type = {"To": 1, "CC": 2, "BCC": 3}.get(prop, 1)
            if r.Type == target_type:
                addr = r.Address
                if addr and addr.startswith("/O="):
                    try:
                        user = r.AddressEntry.GetExchangeUser()
                        if user:
                            addr = user.PrimarySmtpAddress
                    except Exception:
                        addr = r.Name
                result.append(addr or r.Name)
    except Exception:
        # Fallback to simple To/CC property
        val = getattr(item, prop, "")
        if val:
            result = [x.strip() for x in val.split(";") if x.strip()]
    return result


def _format_item(item, index: int = 0) -> dict:
    """Format an Outlook mail item to a dict."""
    try:
        received = item.ReceivedTime.strftime("%Y-%m-%dT%H:%M:%S") if item.ReceivedTime else None
    except Exception:
        received = None

    return {
        "index": index,
        "entryID": item.EntryID,
        "subject": item.Subject or "",
        "from": _get_sender_email(item),
        "to": _get_recipients(item, "To"),
        "receivedDateTime": received,
        "isRead": not item.UnRead,
        "hasAttachments": item.Attachments.Count > 0,
        "importance": {0: "low", 1: "normal", 2: "high"}.get(item.Importance, "normal"),
    }


def _collect_folders(folder, depth: int = 0, max_depth: int = 3) -> list:
    """Recursively collect folder info."""
    result = []
    try:
        for f in folder.Folders:
            info = {
                "name": f.Name,
                "fullPath": f.FolderPath,
                "unreadCount": f.UnReadItemCount,
                "totalCount": f.Items.Count,
                "depth": depth,
            }
            result.append(info)
            if depth < max_depth:
                result.extend(_collect_folders(f, depth + 1, max_depth))
    except Exception:
        pass
    return result


# =============================================================================
# SYNC IMPLEMENTATIONS (run in thread via asyncio.to_thread)
# =============================================================================


def _sync_list_folders(parent_folder_name: str = None) -> str:
    try:
        outlook, namespace = _get_outlook()
        store = _get_store(namespace)
        if not store:
            available = getattr(_get_store, '_last_available', [])
            return json.dumps({
                "error": f"Mail store not found for filter '{STORE_FILTER}'",
                "available_stores": available,
                "hint": "Check STORE_FILTER value or Outlook profile configuration",
            }, ensure_ascii=False)

        if parent_folder_name:
            parent = _resolve_folder(namespace, parent_folder_name)
            if not parent:
                return json.dumps({"error": f"Folder not found: {parent_folder_name}"})
            folders = _collect_folders(parent, max_depth=2)
        else:
            root = store.GetRootFolder()
            folders = _collect_folders(root, max_depth=2)

        return json.dumps({"count": len(folders), "folders": folders}, indent=2, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": f"{type(e).__name__}: {str(e)}"})
    finally:
        pythoncom.CoUninitialize()


def _sync_list_messages(folder_name: str, top: int, is_read: bool = None, skip: int = 0) -> str:
    try:
        outlook, namespace = _get_outlook()
        folder = _resolve_folder(namespace, folder_name)
        if not folder:
            return json.dumps({"error": f"Folder not found: {folder_name}"})

        items = folder.Items
        items.Sort("[ReceivedTime]", True)

        # Apply read filter
        if is_read is not None:
            unread_val = "false" if is_read else "true"
            restriction = f"[UnRead] = {unread_val}"
            items = items.Restrict(restriction)

        result = []
        count = 0
        skipped = 0
        for item in items:
            if skipped < skip:
                skipped += 1
                continue
            if count >= top:
                break
            try:
                if item.Class == 43:  # olMail
                    result.append(_format_item(item, count + 1))
                    count += 1
            except Exception:
                continue

        return json.dumps({
            "count": len(result),
            "folder": folder_name,
            "messages": result,
        }, indent=2, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": f"{type(e).__name__}: {str(e)}"})
    finally:
        pythoncom.CoUninitialize()


def _sync_search_messages(query: str, folder_name: str = None, top: int = 25) -> str:
    try:
        outlook, namespace = _get_outlook()

        # Build search folders list
        folders_to_search = []
        if folder_name:
            f = _resolve_folder(namespace, folder_name)
            if f:
                folders_to_search.append(f)
            else:
                return json.dumps({"error": f"Folder not found: {folder_name}"})
        else:
            # Search inbox + sent by default
            folders_to_search.append(namespace.GetDefaultFolder(6))  # inbox
            folders_to_search.append(namespace.GetDefaultFolder(5))  # sent

        # Build DASL filter for subject/body search
        filter_str = (
            f'@SQL="urn:schemas:httpmail:subject" LIKE \'%{query}%\' OR '
            f'"urn:schemas:httpmail:textdescription" LIKE \'%{query}%\' OR '
            f'"urn:schemas:httpmail:fromemail" LIKE \'%{query}%\''
        )

        result = []
        for folder in folders_to_search:
            try:
                items = folder.Items.Restrict(filter_str)
                items.Sort("[ReceivedTime]", True)
                count = 0
                for item in items:
                    if count >= top:
                        break
                    try:
                        if item.Class == 43:
                            entry = _format_item(item, len(result) + 1)
                            entry["folder"] = folder.Name
                            # Add body preview
                            body = item.Body or ""
                            entry["bodyPreview"] = body[:200].replace("\r\n", " ").strip()
                            result.append(entry)
                            count += 1
                    except Exception:
                        continue
            except Exception:
                continue

        return json.dumps({"count": len(result), "query": query, "messages": result}, indent=2, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": f"{type(e).__name__}: {str(e)}"})
    finally:
        pythoncom.CoUninitialize()


def _sync_get_message(entry_id: str, include_body: bool = True) -> str:
    try:
        outlook, namespace = _get_outlook()
        item = namespace.GetItemFromID(entry_id)

        try:
            received = item.ReceivedTime.strftime("%Y-%m-%dT%H:%M:%S") if item.ReceivedTime else None
        except Exception:
            received = None
        try:
            sent = item.SentOn.strftime("%Y-%m-%dT%H:%M:%S") if item.SentOn else None
        except Exception:
            sent = None

        result = {
            "entryID": item.EntryID,
            "subject": item.Subject or "",
            "from": _get_sender_email(item),
            "to": _get_recipients(item, "To"),
            "cc": _get_recipients(item, "CC"),
            "receivedDateTime": received,
            "sentDateTime": sent,
            "isRead": not item.UnRead,
            "hasAttachments": item.Attachments.Count > 0,
            "importance": {0: "low", 1: "normal", 2: "high"}.get(item.Importance, "normal"),
            "conversationTopic": item.ConversationTopic or "",
        }

        if include_body:
            result["body"] = item.Body or ""
            result["htmlBody"] = item.HTMLBody or ""

        if item.Attachments.Count > 0:
            attachments = []
            for i in range(1, item.Attachments.Count + 1):
                att = item.Attachments.Item(i)
                attachments.append({
                    "name": att.FileName,
                    "size": att.Size,
                    "index": i,
                })
            result["attachments"] = attachments

        return json.dumps(result, indent=2, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": f"{type(e).__name__}: {str(e)}"})
    finally:
        pythoncom.CoUninitialize()


def _sync_create_draft(subject: str, body: str, to: list, cc: list = None, importance: str = None, display: bool = False) -> str:
    try:
        outlook, namespace = _get_outlook()
        mail = outlook.CreateItem(0)  # olMailItem

        mail.To = "; ".join(to)
        mail.Subject = subject
        mail.HTMLBody = body

        if cc:
            mail.CC = "; ".join(cc)

        if importance:
            imp_map = {"low": 0, "normal": 1, "high": 2}
            mail.Importance = imp_map.get(importance.lower(), 1)

        mail.Save()

        # Verify the draft persisted
        entry_id = mail.EntryID
        verified = False
        try:
            saved = namespace.GetItemFromID(entry_id)
            verified = saved is not None and saved.Subject == subject
        except Exception:
            pass

        # Open the draft in Outlook if requested
        if display:
            try:
                mail.Display()
            except Exception:
                pass

        return json.dumps({
            "status": "draft_created",
            "entryID": entry_id,
            "subject": mail.Subject,
            "to": to,
            "verified": verified,
            "displayed": display,
        }, indent=2, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": f"{type(e).__name__}: {str(e)}"})
    finally:
        pythoncom.CoUninitialize()


def _sync_send_draft(entry_id: str) -> str:
    try:
        outlook, namespace = _get_outlook()
        item = namespace.GetItemFromID(entry_id)
        item.Send()

        return json.dumps({
            "status": "sent",
            "entryID": entry_id,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": f"{type(e).__name__}: {str(e)}"})
    finally:
        pythoncom.CoUninitialize()


def _sync_reply_draft(entry_id: str, body: str, reply_all: bool = False) -> str:
    try:
        outlook, namespace = _get_outlook()
        item = namespace.GetItemFromID(entry_id)

        if reply_all:
            reply = item.ReplyAll()
        else:
            reply = item.Reply()

        reply.HTMLBody = body + reply.HTMLBody
        reply.Save()

        return json.dumps({
            "status": "reply_draft_created",
            "entryID": reply.EntryID,
            "reply_all": reply_all,
            "original_entryID": entry_id,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": f"{type(e).__name__}: {str(e)}"})
    finally:
        pythoncom.CoUninitialize()


def _sync_move_message(entry_id: str, destination_folder: str) -> str:
    try:
        outlook, namespace = _get_outlook()
        item = namespace.GetItemFromID(entry_id)

        dest = _resolve_folder(namespace, destination_folder)
        if not dest:
            return json.dumps({"error": f"Destination folder not found: {destination_folder}"})

        moved = item.Move(dest)

        return json.dumps({
            "status": "moved",
            "entryID": moved.EntryID,
            "destination": destination_folder,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": f"{type(e).__name__}: {str(e)}"})
    finally:
        pythoncom.CoUninitialize()


def _sync_mark_read(entry_id: str, is_read: bool = True) -> str:
    try:
        outlook, namespace = _get_outlook()
        item = namespace.GetItemFromID(entry_id)
        item.UnRead = not is_read
        item.Save()

        return json.dumps({
            "status": "updated",
            "entryID": entry_id,
            "isRead": is_read,
        }, indent=2)
    except Exception as e:
        return json.dumps({"error": f"{type(e).__name__}: {str(e)}"})
    finally:
        pythoncom.CoUninitialize()


def _sync_create_folder(display_name: str, parent_folder_name: str = None) -> str:
    try:
        outlook, namespace = _get_outlook()

        if parent_folder_name:
            parent = _resolve_folder(namespace, parent_folder_name)
            if not parent:
                return json.dumps({"error": f"Parent folder not found: {parent_folder_name}"})
        else:
            store = _get_store(namespace)
            if not store:
                return json.dumps({"error": "Mail store not found"})
            parent = store.GetRootFolder()

        new_folder = parent.Folders.Add(display_name)

        return json.dumps({
            "status": "created",
            "name": new_folder.Name,
            "fullPath": new_folder.FolderPath,
        }, indent=2, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": f"{type(e).__name__}: {str(e)}"})
    finally:
        pythoncom.CoUninitialize()


def _sync_search_gal(query: str, top: int = 10) -> str:
    try:
        outlook, namespace = _get_outlook()
        gal = namespace.GetGlobalAddressList()
        if not gal:
            return json.dumps({"error": "Global Address List not available"})

        entries = gal.AddressEntries
        query_lower = query.lower()
        results = []

        for i in range(1, entries.Count + 1):
            if len(results) >= top:
                break
            try:
                entry = entries.Item(i)
                name = entry.Name or ""
                if query_lower not in name.lower():
                    continue

                record = {"name": name, "type": "unknown"}

                # Try as Exchange User first
                try:
                    user = entry.GetExchangeUser()
                    if user:
                        record.update({
                            "type": "user",
                            "email": user.PrimarySmtpAddress or "",
                            "jobTitle": user.JobTitle or "",
                            "department": user.Department or "",
                            "company": user.CompanyName or "",
                        })
                        results.append(record)
                        continue
                except Exception:
                    pass

                # Try as Distribution List
                try:
                    dl = entry.GetExchangeDistributionList()
                    if dl:
                        record.update({
                            "type": "distribution_list",
                            "email": dl.PrimarySmtpAddress or "",
                        })
                        results.append(record)
                        continue
                except Exception:
                    pass

                # Fallback: include with address if available
                record["email"] = entry.Address or ""
                results.append(record)
            except Exception:
                continue

        return json.dumps({"count": len(results), "query": query, "results": results}, indent=2, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": f"{type(e).__name__}: {str(e)}"})
    finally:
        pythoncom.CoUninitialize()


def _sync_resolve_recipient(names: list) -> str:
    try:
        outlook, namespace = _get_outlook()
        results = []

        for name in names:
            record = {"name": name, "resolved": False, "email": None, "displayName": None}
            try:
                recipient = namespace.CreateRecipient(name)
                recipient.Resolve()
                if recipient.Resolved:
                    record["resolved"] = True
                    record["displayName"] = recipient.Name or name
                    entry = recipient.AddressEntry
                    if entry:
                        try:
                            user = entry.GetExchangeUser()
                            if user and user.PrimarySmtpAddress:
                                record["email"] = user.PrimarySmtpAddress
                            else:
                                record["email"] = entry.Address or ""
                        except Exception:
                            record["email"] = entry.Address or ""
            except Exception as e:
                record["error"] = str(e)
            results.append(record)

        return json.dumps({"count": len(results), "results": results}, indent=2, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": f"{type(e).__name__}: {str(e)}"})
    finally:
        pythoncom.CoUninitialize()


# =============================================================================
# PYDANTIC INPUT MODELS
# =============================================================================

_model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")


class ListFoldersInput(BaseModel):
    """Input for listing mail folders."""
    model_config = _model_config
    parent_folder: Optional[str] = Field(
        default=None,
        description="Parent folder name to list child folders. If not specified, lists all folders."
    )


class ListMessagesInput(BaseModel):
    """Input for listing messages in a folder."""
    model_config = _model_config
    folder: Optional[str] = Field(
        default=None,
        description="Folder name or well-known name (inbox, drafts, sentitems, sent, deleteditems, junkemail). Default: inbox"
    )
    top: int = Field(
        default=25,
        ge=1, le=100,
        description="Number of messages to return (1-100, default 25)"
    )
    is_read: Optional[bool] = Field(
        default=None,
        description="Filter by read status. None=all, True=read only, False=unread only"
    )
    skip: int = Field(
        default=0,
        ge=0,
        description="Number of messages to skip (for pagination)"
    )


class SearchMessagesInput(BaseModel):
    """Input for searching messages."""
    model_config = _model_config
    query: str = Field(
        ...,
        min_length=1,
        description="Search text. Searches subject, body, and sender email."
    )
    folder: Optional[str] = Field(
        default=None,
        description="Folder name to search in. If not specified, searches inbox and sent."
    )
    top: int = Field(
        default=25,
        ge=1, le=50,
        description="Max results to return (1-50, default 25)"
    )


class GetMessageInput(BaseModel):
    """Input for getting a full message."""
    model_config = _model_config
    entry_id: str = Field(..., description="Message EntryID (from list/search results)")
    include_body: bool = Field(
        default=True,
        description="Include the full message body (default: True)"
    )


class CreateDraftInput(BaseModel):
    """Input for creating a draft email."""
    model_config = _model_config
    subject: str = Field(..., min_length=1, max_length=512, description="Email subject")
    body: str = Field(..., min_length=1, description="Email body (HTML supported)")
    to: list[str] = Field(..., min_length=1, description="List of recipient email addresses")
    cc: Optional[list[str]] = Field(
        default=None,
        description="List of CC email addresses"
    )
    importance: Optional[str] = Field(
        default=None,
        description="Email importance: low, normal, high"
    )
    display: bool = Field(
        default=False,
        description="Open the draft in Outlook after creation for visual review (default: False)"
    )


class SendDraftInput(BaseModel):
    """Input for sending a draft."""
    model_config = _model_config
    entry_id: str = Field(..., description="Draft EntryID to send")


class ReplyDraftInput(BaseModel):
    """Input for creating a reply draft."""
    model_config = _model_config
    entry_id: str = Field(..., description="EntryID of the message to reply to")
    body: str = Field(..., min_length=1, description="Reply body (HTML supported)")
    reply_all: bool = Field(
        default=False,
        description="Reply to all recipients (default: False, reply to sender only)"
    )


class MoveMessageInput(BaseModel):
    """Input for moving a message to another folder."""
    model_config = _model_config
    entry_id: str = Field(..., description="EntryID of the message to move")
    destination_folder: str = Field(
        ...,
        description=(
            "Destination folder. Supports: (1) well-known names: inbox, drafts, sentitems, deleteditems, junkemail; "
            "(2) path syntax: 'PROYECTOS/EVOCON' or '1. PROYECTOS/EVOCON'; "
            "(3) exact folder name: '3. REGULATORIO'; "
            "(4) partial match: 'REGULATORIO' matches '3. REGULATORIO'"
        ),
    )


class MarkReadInput(BaseModel):
    """Input for marking a message as read/unread."""
    model_config = _model_config
    entry_id: str = Field(..., description="Message EntryID")
    is_read: bool = Field(
        default=True,
        description="True=mark as read, False=mark as unread"
    )


class CreateFolderInput(BaseModel):
    """Input for creating a mail folder."""
    model_config = _model_config
    display_name: str = Field(..., min_length=1, max_length=256, description="Folder name")
    parent_folder: Optional[str] = Field(
        default=None,
        description="Parent folder name to create subfolder. If not specified, creates at top level."
    )


class SearchGALInput(BaseModel):
    """Input for searching the Global Address List."""
    model_config = _model_config
    query: str = Field(
        ...,
        min_length=2,
        description="Search text to match against contact names (case-insensitive)"
    )
    top: int = Field(
        default=10,
        ge=1, le=50,
        description="Max results to return (1-50, default 10)"
    )


class ResolveRecipientInput(BaseModel):
    """Input for resolving names to SMTP email addresses."""
    model_config = _model_config
    names: list[str] = Field(
        ...,
        min_length=1,
        description="List of names, aliases, or partial emails to resolve to SMTP addresses"
    )


# =============================================================================
# TOOL IMPLEMENTATIONS
# =============================================================================


@mcp.tool(
    name="outlook_list_folders",
    annotations={
        "title": "List Mail Folders",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def outlook_list_folders(params: ListFoldersInput) -> str:
    """List mail folders with unread counts. Supports listing subfolders.

    Args:
        params: ListFoldersInput with optional parent_folder

    Returns:
        JSON list of folders with name, unreadCount, totalCount
    """
    return await asyncio.to_thread(_sync_list_folders, params.parent_folder)


@mcp.tool(
    name="outlook_list_messages",
    annotations={
        "title": "List Messages",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def outlook_list_messages(params: ListMessagesInput) -> str:
    """List messages in a mail folder (default: inbox).

    Supports pagination with top/skip and filtering by read status.

    Args:
        params: ListMessagesInput with optional folder, top, is_read, skip

    Returns:
        JSON list of messages with entryID, subject, from, to, receivedDateTime, isRead
    """
    return await asyncio.to_thread(
        _sync_list_messages,
        params.folder or "inbox",
        params.top,
        params.is_read,
        params.skip,
    )


@mcp.tool(
    name="outlook_search_messages",
    annotations={
        "title": "Search Messages",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def outlook_search_messages(params: SearchMessagesInput) -> str:
    """Search emails by text in subject, body, or sender.

    Searches inbox and sent folders by default, or a specific folder.

    Args:
        params: SearchMessagesInput with query, optional folder, top

    Returns:
        JSON list of matching messages with bodyPreview
    """
    return await asyncio.to_thread(
        _sync_search_messages,
        params.query,
        params.folder,
        params.top,
    )


@mcp.tool(
    name="outlook_get_message",
    annotations={
        "title": "Get Message",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def outlook_get_message(params: GetMessageInput) -> str:
    """Read a full email including body, headers, and attachment info.

    Args:
        params: GetMessageInput with entry_id, optional include_body

    Returns:
        JSON with full message details including body and attachments
    """
    return await asyncio.to_thread(
        _sync_get_message,
        params.entry_id,
        params.include_body,
    )


@mcp.tool(
    name="outlook_create_draft",
    annotations={
        "title": "Create Draft Email",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": False,
    },
)
async def outlook_create_draft(params: CreateDraftInput) -> str:
    """Create a draft email with To, CC, Subject, and HTML Body.

    The draft is saved in the Drafts folder for review before sending.
    Use outlook_send_draft to send it. Set display=true to open the draft
    in Outlook for visual review (recommended).

    Args:
        params: CreateDraftInput with subject, body, to, optional cc, importance, display

    Returns:
        JSON with the created draft's entryID, subject, verified status, and displayed flag
    """
    return await asyncio.to_thread(
        _sync_create_draft,
        params.subject,
        params.body,
        params.to,
        params.cc,
        params.importance,
        params.display,
    )


@mcp.tool(
    name="outlook_send_draft",
    annotations={
        "title": "Send Draft Email",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": False,
    },
)
async def outlook_send_draft(params: SendDraftInput) -> str:
    """Send an existing draft email.

    The draft must already exist (created via outlook_create_draft or in Outlook).

    Args:
        params: SendDraftInput with entry_id

    Returns:
        JSON with send status
    """
    return await asyncio.to_thread(_sync_send_draft, params.entry_id)


@mcp.tool(
    name="outlook_reply_draft",
    annotations={
        "title": "Create Reply Draft",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": False,
    },
)
async def outlook_reply_draft(params: ReplyDraftInput) -> str:
    """Create a draft reply to an existing email.

    Creates a reply draft that can be reviewed before sending with outlook_send_draft.
    The reply body is prepended to the original email thread.

    Args:
        params: ReplyDraftInput with entry_id, body, optional reply_all

    Returns:
        JSON with the reply draft's entryID
    """
    return await asyncio.to_thread(
        _sync_reply_draft,
        params.entry_id,
        params.body,
        params.reply_all,
    )


@mcp.tool(
    name="outlook_move_message",
    annotations={
        "title": "Move Message",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def outlook_move_message(params: MoveMessageInput) -> str:
    """Move an email to another folder.

    Supports: well-known names (inbox, drafts, sentitems, deleteditems, junkemail),
    path syntax ("1. PROYECTOS/EVOCON"), exact names ("3. REGULATORIO"),
    or partial match ("REGULATORIO" finds "3. REGULATORIO").

    Args:
        params: MoveMessageInput with entry_id and destination_folder

    Returns:
        JSON with move status and new entryID
    """
    return await asyncio.to_thread(
        _sync_move_message,
        params.entry_id,
        params.destination_folder,
    )


@mcp.tool(
    name="outlook_mark_read",
    annotations={
        "title": "Mark Read/Unread",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def outlook_mark_read(params: MarkReadInput) -> str:
    """Mark an email as read or unread.

    Args:
        params: MarkReadInput with entry_id and is_read flag

    Returns:
        JSON with update status
    """
    return await asyncio.to_thread(
        _sync_mark_read,
        params.entry_id,
        params.is_read,
    )


@mcp.tool(
    name="outlook_create_folder",
    annotations={
        "title": "Create Mail Folder",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": False,
    },
)
async def outlook_create_folder(params: CreateFolderInput) -> str:
    """Create a new mail folder. Supports creating subfolders.

    Args:
        params: CreateFolderInput with display_name and optional parent_folder

    Returns:
        JSON with the created folder's name and fullPath
    """
    return await asyncio.to_thread(
        _sync_create_folder,
        params.display_name,
        params.parent_folder,
    )


@mcp.tool(
    name="outlook_search_gal",
    annotations={
        "title": "Search Global Address List",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def outlook_search_gal(params: SearchGALInput) -> str:
    """Search the Global Address List (GAL) by name.

    Finds contacts, users, and distribution lists matching the query.
    Returns name, email, job title, department, and company for each match.

    Args:
        params: SearchGALInput with query (min 2 chars) and optional top

    Returns:
        JSON list of matching GAL entries with contact details
    """
    return await asyncio.to_thread(
        _sync_search_gal,
        params.query,
        params.top,
    )


@mcp.tool(
    name="outlook_resolve_recipient",
    annotations={
        "title": "Resolve Recipient",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def outlook_resolve_recipient(params: ResolveRecipientInput) -> str:
    """Resolve names or aliases to SMTP email addresses.

    Uses Outlook's built-in recipient resolution to find the best match
    for each name. Useful for verifying email addresses before creating drafts.

    Args:
        params: ResolveRecipientInput with list of names to resolve

    Returns:
        JSON list with resolved status, email, and display name for each input
    """
    return await asyncio.to_thread(
        _sync_resolve_recipient,
        params.names,
    )


# =============================================================================
# ENTRY POINT
# =============================================================================

if __name__ == "__main__":
    mcp.run()
