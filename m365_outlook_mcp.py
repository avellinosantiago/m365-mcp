#!/usr/bin/env python3
"""
M365 Outlook MCP Server
Provides tools for managing Outlook emails, drafts, and folders
via Microsoft Graph API. Uses Azure CLI authentication.

Run: python m365_outlook_mcp.py
"""

import json
import os
from pathlib import Path
from typing import Optional

import httpx
from azure.identity import DefaultAzureCredential
from dotenv import load_dotenv
from mcp.server.fastmcp import FastMCP
from pydantic import BaseModel, ConfigDict, Field

# =============================================================================
# CONFIGURATION
# =============================================================================

env_path = Path(__file__).parent / ".env"
load_dotenv(env_path)

GRAPH_API_BASE = "https://graph.microsoft.com/v1.0"
GRAPH_TIMEOUT = int(os.getenv("GRAPH_TIMEOUT", "30"))

# =============================================================================
# MCP SERVER
# =============================================================================

mcp = FastMCP("m365_outlook_mcp")

# =============================================================================
# AUTHENTICATION
# =============================================================================

_credential = DefaultAzureCredential()


def get_graph_token() -> str:
    """Get access token for Microsoft Graph API."""
    token = _credential.get_token("https://graph.microsoft.com/.default")
    return token.token


def get_headers() -> dict:
    """Get authorization headers for Graph API requests."""
    return {
        "Authorization": f"Bearer {get_graph_token()}",
        "Content-Type": "application/json",
    }


# =============================================================================
# HTTP HELPERS
# =============================================================================


async def graph_get(client: httpx.AsyncClient, path: str, params: dict = None) -> dict:
    """GET request to Graph API."""
    url = f"{GRAPH_API_BASE}{path}"
    response = await client.get(url, headers=get_headers(), params=params)
    response.raise_for_status()
    if response.status_code == 204:
        return {}
    return response.json()


async def graph_post(client: httpx.AsyncClient, path: str, body: dict = None) -> dict:
    """POST request to Graph API."""
    url = f"{GRAPH_API_BASE}{path}"
    kwargs = {"headers": get_headers()}
    if body is not None:
        kwargs["json"] = body
    response = await client.post(url, **kwargs)
    response.raise_for_status()
    if response.status_code == 202 or response.status_code == 204:
        return {"status": "ok"}
    return response.json()


async def graph_patch(client: httpx.AsyncClient, path: str, body: dict) -> dict:
    """PATCH request to Graph API."""
    url = f"{GRAPH_API_BASE}{path}"
    response = await client.patch(url, headers=get_headers(), json=body)
    response.raise_for_status()
    if response.status_code == 204:
        return {"status": "updated"}
    return response.json()


# =============================================================================
# HELPERS
# =============================================================================


def format_recipients(emails: list[str]) -> list[dict]:
    """Convert list of email strings to Graph API recipient format."""
    recipients = []
    for email in emails:
        email = email.strip()
        if email:
            recipients.append({"emailAddress": {"address": email}})
    return recipients


def format_message_summary(msg: dict) -> dict:
    """Extract key fields from a Graph API message object."""
    from_addr = ""
    if msg.get("from") and msg["from"].get("emailAddress"):
        from_addr = msg["from"]["emailAddress"].get("address", "")

    to_list = []
    for r in msg.get("toRecipients", []):
        if r.get("emailAddress"):
            to_list.append(r["emailAddress"].get("address", ""))

    return {
        "id": msg["id"],
        "subject": msg.get("subject", ""),
        "from": from_addr,
        "to": to_list,
        "receivedDateTime": msg.get("receivedDateTime"),
        "isRead": msg.get("isRead"),
        "hasAttachments": msg.get("hasAttachments", False),
        "importance": msg.get("importance", "normal"),
        "isDraft": msg.get("isDraft", False),
    }


# =============================================================================
# PYDANTIC INPUT MODELS
# =============================================================================

_model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")


class ListFoldersInput(BaseModel):
    """Input for listing mail folders."""
    model_config = _model_config
    parent_folder_id: Optional[str] = Field(
        default=None,
        description="Parent folder ID to list child folders. If not specified, lists top-level folders."
    )


class ListMessagesInput(BaseModel):
    """Input for listing messages in a folder."""
    model_config = _model_config
    folder_id: Optional[str] = Field(
        default=None,
        description="Folder ID or well-known name (inbox, drafts, sentitems, deleteditems, junkemail, archive). Default: inbox"
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
        description="Search query. Searches subject, body, and participants. Examples: 'budget report', 'from:user@domain.com', 'subject:quarterly'"
    )
    folder_id: Optional[str] = Field(
        default=None,
        description="Folder ID or well-known name to search in. If not specified, searches all folders."
    )
    top: int = Field(
        default=25,
        ge=1, le=50,
        description="Max results to return (1-50, default 25)"
    )


class GetMessageInput(BaseModel):
    """Input for getting a full message."""
    model_config = _model_config
    message_id: str = Field(..., description="Message ID")
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


class SendDraftInput(BaseModel):
    """Input for sending a draft."""
    model_config = _model_config
    message_id: str = Field(..., description="Draft message ID to send")


class ReplyDraftInput(BaseModel):
    """Input for creating a reply draft."""
    model_config = _model_config
    message_id: str = Field(..., description="Message ID to reply to")
    body: str = Field(..., min_length=1, description="Reply body (HTML supported)")
    reply_all: bool = Field(
        default=False,
        description="Reply to all recipients (default: False, reply to sender only)"
    )


class MoveMessageInput(BaseModel):
    """Input for moving a message to another folder."""
    model_config = _model_config
    message_id: str = Field(..., description="Message ID to move")
    destination_folder_id: str = Field(
        ...,
        description="Destination folder ID or well-known name (inbox, drafts, sentitems, deleteditems, junkemail, archive)"
    )


class MarkReadInput(BaseModel):
    """Input for marking a message as read/unread."""
    model_config = _model_config
    message_id: str = Field(..., description="Message ID")
    is_read: bool = Field(
        default=True,
        description="True=mark as read, False=mark as unread"
    )


class CreateFolderInput(BaseModel):
    """Input for creating a mail folder."""
    model_config = _model_config
    display_name: str = Field(..., min_length=1, max_length=256, description="Folder name")
    parent_folder_id: Optional[str] = Field(
        default=None,
        description="Parent folder ID to create subfolder. If not specified, creates at top level."
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
        "openWorldHint": True,
    },
)
async def outlook_list_folders(params: ListFoldersInput) -> str:
    """List mail folders with unread counts. Supports listing subfolders.

    Args:
        params: ListFoldersInput with optional parent_folder_id

    Returns:
        JSON list of folders with id, displayName, unreadItemCount, totalItemCount
    """
    try:
        async with httpx.AsyncClient(timeout=GRAPH_TIMEOUT) as client:
            if params.parent_folder_id:
                path = f"/me/mailFolders/{params.parent_folder_id}/childFolders"
            else:
                path = "/me/mailFolders"

            query = {"$top": "100"}
            data = await graph_get(client, path, params=query)
            folders = data.get("value", [])

            # Pagination
            while "@odata.nextLink" in data:
                next_url = data["@odata.nextLink"]
                response = await client.get(next_url, headers=get_headers())
                response.raise_for_status()
                data = response.json()
                folders.extend(data.get("value", []))

            result = []
            for f in folders:
                result.append({
                    "id": f["id"],
                    "displayName": f["displayName"],
                    "unreadItemCount": f.get("unreadItemCount", 0),
                    "totalItemCount": f.get("totalItemCount", 0),
                    "childFolderCount": f.get("childFolderCount", 0),
                })

            return json.dumps({"count": len(result), "folders": result}, indent=2)

    except httpx.HTTPStatusError as e:
        return json.dumps({"error": f"HTTP {e.response.status_code}: {e.response.text}"})
    except Exception as e:
        return json.dumps({"error": f"{type(e).__name__}: {str(e)}"})


@mcp.tool(
    name="outlook_list_messages",
    annotations={
        "title": "List Messages",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def outlook_list_messages(params: ListMessagesInput) -> str:
    """List messages in a mail folder (default: inbox).

    Supports pagination with top/skip and filtering by read status.

    Args:
        params: ListMessagesInput with optional folder_id, top, is_read, skip

    Returns:
        JSON list of messages with id, subject, from, to, receivedDateTime, isRead
    """
    try:
        folder = params.folder_id or "inbox"
        path = f"/me/mailFolders/{folder}/messages"

        query = {
            "$top": str(params.top),
            "$skip": str(params.skip),
            "$orderby": "receivedDateTime desc",
            "$select": "id,subject,from,toRecipients,receivedDateTime,isRead,hasAttachments,importance,isDraft",
        }

        if params.is_read is not None:
            query["$filter"] = f"isRead eq {str(params.is_read).lower()}"

        async with httpx.AsyncClient(timeout=GRAPH_TIMEOUT) as client:
            data = await graph_get(client, path, params=query)
            messages = data.get("value", [])

            result = [format_message_summary(m) for m in messages]
            output = {"count": len(result), "messages": result}

            if "@odata.nextLink" in data:
                output["hasMore"] = True

            return json.dumps(output, indent=2)

    except httpx.HTTPStatusError as e:
        return json.dumps({"error": f"HTTP {e.response.status_code}: {e.response.text}"})
    except Exception as e:
        return json.dumps({"error": f"{type(e).__name__}: {str(e)}"})


@mcp.tool(
    name="outlook_search_messages",
    annotations={
        "title": "Search Messages",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def outlook_search_messages(params: SearchMessagesInput) -> str:
    """Search emails by text, sender, or subject.

    Uses Graph API $search for full-text search across subject, body, and participants.
    For sender filtering, use 'from:user@domain.com' in the query.
    For subject filtering, use 'subject:keyword' in the query.

    Args:
        params: SearchMessagesInput with query, optional folder_id, top

    Returns:
        JSON list of matching messages
    """
    try:
        if params.folder_id:
            path = f"/me/mailFolders/{params.folder_id}/messages"
        else:
            path = "/me/messages"

        query = {
            "$search": f'"{params.query}"',
            "$top": str(params.top),
            "$select": "id,subject,from,toRecipients,receivedDateTime,isRead,hasAttachments,importance,isDraft,bodyPreview",
        }

        async with httpx.AsyncClient(timeout=GRAPH_TIMEOUT) as client:
            data = await graph_get(client, path, params=query)
            messages = data.get("value", [])

            result = []
            for m in messages:
                summary = format_message_summary(m)
                summary["bodyPreview"] = m.get("bodyPreview", "")[:200]
                result.append(summary)

            return json.dumps({"count": len(result), "messages": result}, indent=2)

    except httpx.HTTPStatusError as e:
        return json.dumps({"error": f"HTTP {e.response.status_code}: {e.response.text}"})
    except Exception as e:
        return json.dumps({"error": f"{type(e).__name__}: {str(e)}"})


@mcp.tool(
    name="outlook_get_message",
    annotations={
        "title": "Get Message",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def outlook_get_message(params: GetMessageInput) -> str:
    """Read a full email including body, headers, and attachment info.

    Args:
        params: GetMessageInput with message_id, optional include_body

    Returns:
        JSON with full message details
    """
    try:
        select_fields = "id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,isDraft,internetMessageHeaders,conversationId"
        if params.include_body:
            select_fields += ",body"

        path = f"/me/messages/{params.message_id}"
        query = {"$select": select_fields}

        async with httpx.AsyncClient(timeout=GRAPH_TIMEOUT) as client:
            msg = await graph_get(client, path, params=query)

            from_addr = ""
            from_name = ""
            if msg.get("from") and msg["from"].get("emailAddress"):
                from_addr = msg["from"]["emailAddress"].get("address", "")
                from_name = msg["from"]["emailAddress"].get("name", "")

            to_list = []
            for r in msg.get("toRecipients", []):
                if r.get("emailAddress"):
                    to_list.append({
                        "name": r["emailAddress"].get("name", ""),
                        "address": r["emailAddress"].get("address", ""),
                    })

            cc_list = []
            for r in msg.get("ccRecipients", []):
                if r.get("emailAddress"):
                    cc_list.append({
                        "name": r["emailAddress"].get("name", ""),
                        "address": r["emailAddress"].get("address", ""),
                    })

            result = {
                "id": msg["id"],
                "subject": msg.get("subject", ""),
                "from": {"name": from_name, "address": from_addr},
                "to": to_list,
                "cc": cc_list,
                "receivedDateTime": msg.get("receivedDateTime"),
                "sentDateTime": msg.get("sentDateTime"),
                "isRead": msg.get("isRead"),
                "hasAttachments": msg.get("hasAttachments", False),
                "importance": msg.get("importance", "normal"),
                "isDraft": msg.get("isDraft", False),
                "conversationId": msg.get("conversationId"),
            }

            if params.include_body and msg.get("body"):
                result["body"] = {
                    "contentType": msg["body"].get("contentType", ""),
                    "content": msg["body"].get("content", ""),
                }

            # Get attachments info if present
            if msg.get("hasAttachments"):
                att_data = await graph_get(
                    client,
                    f"/me/messages/{params.message_id}/attachments",
                    params={"$select": "id,name,contentType,size"},
                )
                result["attachments"] = [
                    {
                        "id": a["id"],
                        "name": a.get("name", ""),
                        "contentType": a.get("contentType", ""),
                        "size": a.get("size", 0),
                    }
                    for a in att_data.get("value", [])
                ]

            return json.dumps(result, indent=2)

    except httpx.HTTPStatusError as e:
        return json.dumps({"error": f"HTTP {e.response.status_code}: {e.response.text}"})
    except Exception as e:
        return json.dumps({"error": f"{type(e).__name__}: {str(e)}"})


@mcp.tool(
    name="outlook_create_draft",
    annotations={
        "title": "Create Draft Email",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True,
    },
)
async def outlook_create_draft(params: CreateDraftInput) -> str:
    """Create a draft email with To, CC, Subject, and HTML Body.

    The draft is saved in the Drafts folder for review before sending.
    Use outlook_send_draft to send it.

    Args:
        params: CreateDraftInput with subject, body, to, optional cc, importance

    Returns:
        JSON with the created draft's id and subject
    """
    try:
        body = {
            "subject": params.subject,
            "body": {
                "contentType": "HTML",
                "content": params.body,
            },
            "toRecipients": format_recipients(params.to),
        }

        if params.cc:
            body["ccRecipients"] = format_recipients(params.cc)

        if params.importance:
            body["importance"] = params.importance

        async with httpx.AsyncClient(timeout=GRAPH_TIMEOUT) as client:
            result = await graph_post(client, "/me/messages", body)

            return json.dumps({
                "status": "draft_created",
                "id": result["id"],
                "subject": result.get("subject", ""),
                "to": [r["emailAddress"]["address"] for r in result.get("toRecipients", [])],
            }, indent=2)

    except httpx.HTTPStatusError as e:
        return json.dumps({"error": f"HTTP {e.response.status_code}: {e.response.text}"})
    except Exception as e:
        return json.dumps({"error": f"{type(e).__name__}: {str(e)}"})


@mcp.tool(
    name="outlook_send_draft",
    annotations={
        "title": "Send Draft Email",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True,
    },
)
async def outlook_send_draft(params: SendDraftInput) -> str:
    """Send an existing draft email.

    The draft must already exist (created via outlook_create_draft or in Outlook).

    Args:
        params: SendDraftInput with message_id

    Returns:
        JSON with send status
    """
    try:
        async with httpx.AsyncClient(timeout=GRAPH_TIMEOUT) as client:
            await graph_post(client, f"/me/messages/{params.message_id}/send")

            return json.dumps({
                "status": "sent",
                "message_id": params.message_id,
            }, indent=2)

    except httpx.HTTPStatusError as e:
        return json.dumps({"error": f"HTTP {e.response.status_code}: {e.response.text}"})
    except Exception as e:
        return json.dumps({"error": f"{type(e).__name__}: {str(e)}"})


@mcp.tool(
    name="outlook_reply_draft",
    annotations={
        "title": "Create Reply Draft",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True,
    },
)
async def outlook_reply_draft(params: ReplyDraftInput) -> str:
    """Create a draft reply to an existing email.

    Creates a reply draft that can be reviewed before sending with outlook_send_draft.

    Args:
        params: ReplyDraftInput with message_id, body, optional reply_all

    Returns:
        JSON with the reply draft's id
    """
    try:
        action = "createReplyAll" if params.reply_all else "createReply"
        reply_body = {
            "comment": params.body,
        }

        async with httpx.AsyncClient(timeout=GRAPH_TIMEOUT) as client:
            # createReply returns a new draft message
            result = await graph_post(
                client,
                f"/me/messages/{params.message_id}/{action}",
                reply_body,
            )

            # Update the body with the user's HTML content
            draft_id = result["id"]
            await graph_patch(client, f"/me/messages/{draft_id}", {
                "body": {
                    "contentType": "HTML",
                    "content": params.body,
                },
            })

            return json.dumps({
                "status": "reply_draft_created",
                "id": draft_id,
                "reply_all": params.reply_all,
                "original_message_id": params.message_id,
            }, indent=2)

    except httpx.HTTPStatusError as e:
        return json.dumps({"error": f"HTTP {e.response.status_code}: {e.response.text}"})
    except Exception as e:
        return json.dumps({"error": f"{type(e).__name__}: {str(e)}"})


@mcp.tool(
    name="outlook_move_message",
    annotations={
        "title": "Move Message",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def outlook_move_message(params: MoveMessageInput) -> str:
    """Move an email to another folder.

    Supports well-known folder names (inbox, drafts, sentitems, deleteditems,
    junkemail, archive) or folder IDs.

    Args:
        params: MoveMessageInput with message_id and destination_folder_id

    Returns:
        JSON with move status and new message id
    """
    try:
        body = {"destinationId": params.destination_folder_id}

        async with httpx.AsyncClient(timeout=GRAPH_TIMEOUT) as client:
            result = await graph_post(
                client,
                f"/me/messages/{params.message_id}/move",
                body,
            )

            return json.dumps({
                "status": "moved",
                "id": result.get("id", params.message_id),
                "destination": params.destination_folder_id,
            }, indent=2)

    except httpx.HTTPStatusError as e:
        return json.dumps({"error": f"HTTP {e.response.status_code}: {e.response.text}"})
    except Exception as e:
        return json.dumps({"error": f"{type(e).__name__}: {str(e)}"})


@mcp.tool(
    name="outlook_mark_read",
    annotations={
        "title": "Mark Read/Unread",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def outlook_mark_read(params: MarkReadInput) -> str:
    """Mark an email as read or unread.

    Args:
        params: MarkReadInput with message_id and is_read flag

    Returns:
        JSON with update status
    """
    try:
        async with httpx.AsyncClient(timeout=GRAPH_TIMEOUT) as client:
            await graph_patch(
                client,
                f"/me/messages/{params.message_id}",
                {"isRead": params.is_read},
            )

            return json.dumps({
                "status": "updated",
                "message_id": params.message_id,
                "isRead": params.is_read,
            }, indent=2)

    except httpx.HTTPStatusError as e:
        return json.dumps({"error": f"HTTP {e.response.status_code}: {e.response.text}"})
    except Exception as e:
        return json.dumps({"error": f"{type(e).__name__}: {str(e)}"})


@mcp.tool(
    name="outlook_create_folder",
    annotations={
        "title": "Create Mail Folder",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True,
    },
)
async def outlook_create_folder(params: CreateFolderInput) -> str:
    """Create a new mail folder. Supports creating subfolders.

    Args:
        params: CreateFolderInput with display_name and optional parent_folder_id

    Returns:
        JSON with the created folder's id and displayName
    """
    try:
        body = {"displayName": params.display_name}

        if params.parent_folder_id:
            path = f"/me/mailFolders/{params.parent_folder_id}/childFolders"
        else:
            path = "/me/mailFolders"

        async with httpx.AsyncClient(timeout=GRAPH_TIMEOUT) as client:
            result = await graph_post(client, path, body)

            return json.dumps({
                "status": "created",
                "id": result["id"],
                "displayName": result["displayName"],
            }, indent=2)

    except httpx.HTTPStatusError as e:
        return json.dumps({"error": f"HTTP {e.response.status_code}: {e.response.text}"})
    except Exception as e:
        return json.dumps({"error": f"{type(e).__name__}: {str(e)}"})


# =============================================================================
# ENTRY POINT
# =============================================================================

if __name__ == "__main__":
    mcp.run()
