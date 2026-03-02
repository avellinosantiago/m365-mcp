#!/usr/bin/env python3
"""
M365 Planner MCP Server
Provides tools for managing Microsoft Planner tasks, buckets, and plans
via Microsoft Graph API. Uses Azure CLI authentication.

Run: python m365_planner_mcp.py
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
DEFAULT_PLAN_ID = os.getenv("PLANNER_DEFAULT_PLAN_ID", "")
GRAPH_TIMEOUT = int(os.getenv("GRAPH_TIMEOUT", "30"))

# =============================================================================
# MCP SERVER
# =============================================================================

mcp = FastMCP("m365_planner_mcp")

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


async def graph_post(client: httpx.AsyncClient, path: str, body: dict) -> dict:
    """POST request to Graph API."""
    url = f"{GRAPH_API_BASE}{path}"
    response = await client.post(url, headers=get_headers(), json=body)
    response.raise_for_status()
    return response.json()


async def graph_patch(client: httpx.AsyncClient, path: str, body: dict, etag: str) -> dict:
    """PATCH request to Graph API (requires If-Match header with etag)."""
    url = f"{GRAPH_API_BASE}{path}"
    headers = get_headers()
    headers["If-Match"] = etag
    response = await client.patch(url, headers=headers, json=body)
    response.raise_for_status()
    if response.status_code == 204:
        return {"status": "updated"}
    return response.json()


async def graph_delete(client: httpx.AsyncClient, path: str, etag: str) -> dict:
    """DELETE request to Graph API (requires If-Match header with etag)."""
    url = f"{GRAPH_API_BASE}{path}"
    headers = get_headers()
    headers["If-Match"] = etag
    response = await client.delete(url, headers=headers)
    response.raise_for_status()
    return {"status": "deleted"}


async def get_task_etag(client: httpx.AsyncClient, task_id: str) -> str:
    """Get the @odata.etag for a Planner task (required for updates)."""
    task = await graph_get(client, f"/planner/tasks/{task_id}")
    return task["@odata.etag"]


# =============================================================================
# PYDANTIC INPUT MODELS
# =============================================================================

_model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")


class ListTasksInput(BaseModel):
    """Input for listing tasks in a plan."""
    model_config = _model_config
    plan_id: Optional[str] = Field(
        default=None,
        description="Plan ID. Uses PLANNER_DEFAULT_PLAN_ID from .env if not specified"
    )
    bucket_id: Optional[str] = Field(
        default=None,
        description="Filter by bucket ID. If not specified, returns all tasks"
    )
    include_completed: bool = Field(
        default=False,
        description="Include completed tasks (default: only open tasks)"
    )


class MyTasksInput(BaseModel):
    """Input for listing my tasks across all plans."""
    model_config = _model_config
    include_completed: bool = Field(
        default=False,
        description="Include completed tasks (default: only open tasks)"
    )


class OverdueTasksInput(BaseModel):
    """Input for listing overdue tasks."""
    model_config = _model_config
    plan_id: Optional[str] = Field(
        default=None,
        description="Plan ID. Uses PLANNER_DEFAULT_PLAN_ID from .env if not specified"
    )


class GetTaskDetailsInput(BaseModel):
    """Input for getting task details including description and checklist."""
    model_config = _model_config
    task_id: str = Field(..., description="Task ID (GUID)")


class CreateTaskInput(BaseModel):
    """Input for creating a new task."""
    model_config = _model_config
    plan_id: Optional[str] = Field(
        default=None,
        description="Plan ID. Uses PLANNER_DEFAULT_PLAN_ID from .env if not specified"
    )
    bucket_id: str = Field(..., description="Bucket ID to place the task in")
    title: str = Field(..., min_length=1, max_length=256, description="Task title")
    due_date: Optional[str] = Field(
        default=None,
        description="Due date in ISO format (e.g., '2026-03-10'). Optional."
    )


class UpdateTaskInput(BaseModel):
    """Input for updating a task."""
    model_config = _model_config
    task_id: str = Field(..., description="Task ID (GUID)")
    title: Optional[str] = Field(default=None, max_length=256, description="New title")
    bucket_id: Optional[str] = Field(default=None, description="Move to different bucket")
    due_date: Optional[str] = Field(
        default=None,
        description="New due date in ISO format (e.g., '2026-03-10'), or 'clear' to remove"
    )
    percent_complete: Optional[int] = Field(
        default=None,
        ge=0, le=100,
        description="Completion percentage (0, 25, 50, 75, or 100)"
    )
    priority: Optional[int] = Field(
        default=None,
        ge=0, le=10,
        description="Priority: 0=No priority, 1=Urgent, 3=Important, 5=Medium, 9=Low"
    )


class CompleteTaskInput(BaseModel):
    """Input for marking a task as complete."""
    model_config = _model_config
    task_id: str = Field(..., description="Task ID (GUID)")


class UpdateTaskDetailsInput(BaseModel):
    """Input for updating task description and checklist."""
    model_config = _model_config
    task_id: str = Field(..., description="Task ID (GUID)")
    description: Optional[str] = Field(
        default=None,
        description="Task description (replaces existing)"
    )
    checklist_add: Optional[list[str]] = Field(
        default=None,
        description="List of checklist items to add"
    )


class ListBucketsInput(BaseModel):
    """Input for listing buckets in a plan."""
    model_config = _model_config
    plan_id: Optional[str] = Field(
        default=None,
        description="Plan ID. Uses PLANNER_DEFAULT_PLAN_ID from .env if not specified"
    )


class CreateBucketInput(BaseModel):
    """Input for creating a new bucket."""
    model_config = _model_config
    plan_id: Optional[str] = Field(
        default=None,
        description="Plan ID. Uses PLANNER_DEFAULT_PLAN_ID from .env if not specified"
    )
    name: str = Field(..., min_length=1, max_length=256, description="Bucket name")


# =============================================================================
# TOOL IMPLEMENTATIONS
# =============================================================================


@mcp.tool(
    name="planner_list_tasks",
    annotations={
        "title": "List Planner Tasks",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def planner_list_tasks(params: ListTasksInput) -> str:
    """List tasks in a Planner plan, optionally filtered by bucket.

    Returns tasks with title, status, due date, bucket, and priority.
    By default only returns open (incomplete) tasks.

    Args:
        params: ListTasksInput with optional plan_id, bucket_id, include_completed

    Returns:
        JSON list of tasks sorted by due date
    """
    plan_id = params.plan_id or DEFAULT_PLAN_ID
    if not plan_id:
        return json.dumps({"error": "No plan_id provided and PLANNER_DEFAULT_PLAN_ID not set in .env"})

    try:
        async with httpx.AsyncClient(timeout=GRAPH_TIMEOUT) as client:
            data = await graph_get(client, f"/planner/plans/{plan_id}/tasks")
            tasks = data.get("value", [])

            # Pagination
            while "@odata.nextLink" in data:
                next_url = data["@odata.nextLink"]
                response = await client.get(next_url, headers=get_headers())
                response.raise_for_status()
                data = response.json()
                tasks.extend(data.get("value", []))

            # Filter
            if not params.include_completed:
                tasks = [t for t in tasks if t["percentComplete"] < 100]
            if params.bucket_id:
                tasks = [t for t in tasks if t["bucketId"] == params.bucket_id]

            # Format output
            result = []
            for t in sorted(tasks, key=lambda x: x.get("dueDateTime") or "z"):
                result.append({
                    "id": t["id"],
                    "title": t["title"],
                    "percentComplete": t["percentComplete"],
                    "dueDateTime": t.get("dueDateTime"),
                    "bucketId": t["bucketId"],
                    "priority": t.get("priority", 5),
                    "createdDateTime": t.get("createdDateTime"),
                })

            return json.dumps({"count": len(result), "tasks": result}, indent=2)

    except httpx.HTTPStatusError as e:
        return json.dumps({"error": f"HTTP {e.response.status_code}: {e.response.text}"})
    except Exception as e:
        return json.dumps({"error": f"{type(e).__name__}: {str(e)}"})


@mcp.tool(
    name="planner_my_tasks",
    annotations={
        "title": "My Planner Tasks",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def planner_my_tasks(params: MyTasksInput) -> str:
    """List all my tasks across all plans.

    Returns tasks assigned to the current user across all Planner plans.

    Args:
        params: MyTasksInput with optional include_completed flag

    Returns:
        JSON list of tasks with plan context
    """
    try:
        async with httpx.AsyncClient(timeout=GRAPH_TIMEOUT) as client:
            data = await graph_get(client, "/me/planner/tasks")
            tasks = data.get("value", [])

            while "@odata.nextLink" in data:
                next_url = data["@odata.nextLink"]
                response = await client.get(next_url, headers=get_headers())
                response.raise_for_status()
                data = response.json()
                tasks.extend(data.get("value", []))

            if not params.include_completed:
                tasks = [t for t in tasks if t["percentComplete"] < 100]

            result = []
            for t in sorted(tasks, key=lambda x: x.get("dueDateTime") or "z"):
                result.append({
                    "id": t["id"],
                    "title": t["title"],
                    "percentComplete": t["percentComplete"],
                    "dueDateTime": t.get("dueDateTime"),
                    "planId": t["planId"],
                    "bucketId": t["bucketId"],
                    "priority": t.get("priority", 5),
                })

            return json.dumps({"count": len(result), "tasks": result}, indent=2)

    except httpx.HTTPStatusError as e:
        return json.dumps({"error": f"HTTP {e.response.status_code}: {e.response.text}"})
    except Exception as e:
        return json.dumps({"error": f"{type(e).__name__}: {str(e)}"})


@mcp.tool(
    name="planner_overdue",
    annotations={
        "title": "Overdue Planner Tasks",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def planner_overdue(params: OverdueTasksInput) -> str:
    """List overdue tasks (past due date, not completed).

    Args:
        params: OverdueTasksInput with optional plan_id

    Returns:
        JSON list of overdue tasks sorted by due date (oldest first)
    """
    plan_id = params.plan_id or DEFAULT_PLAN_ID
    if not plan_id:
        return json.dumps({"error": "No plan_id provided and PLANNER_DEFAULT_PLAN_ID not set in .env"})

    try:
        from datetime import datetime, timezone

        async with httpx.AsyncClient(timeout=GRAPH_TIMEOUT) as client:
            data = await graph_get(client, f"/planner/plans/{plan_id}/tasks")
            tasks = data.get("value", [])

            while "@odata.nextLink" in data:
                next_url = data["@odata.nextLink"]
                response = await client.get(next_url, headers=get_headers())
                response.raise_for_status()
                data = response.json()
                tasks.extend(data.get("value", []))

            now = datetime.now(timezone.utc).isoformat()
            overdue = [
                t for t in tasks
                if t["percentComplete"] < 100
                and t.get("dueDateTime")
                and t["dueDateTime"] < now
            ]

            result = []
            for t in sorted(overdue, key=lambda x: x["dueDateTime"]):
                result.append({
                    "id": t["id"],
                    "title": t["title"],
                    "dueDateTime": t["dueDateTime"],
                    "percentComplete": t["percentComplete"],
                    "bucketId": t["bucketId"],
                })

            return json.dumps({"count": len(result), "tasks": result}, indent=2)

    except httpx.HTTPStatusError as e:
        return json.dumps({"error": f"HTTP {e.response.status_code}: {e.response.text}"})
    except Exception as e:
        return json.dumps({"error": f"{type(e).__name__}: {str(e)}"})


@mcp.tool(
    name="planner_get_details",
    annotations={
        "title": "Get Task Details",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def planner_get_details(params: GetTaskDetailsInput) -> str:
    """Get full details of a task including description and checklist.

    Args:
        params: GetTaskDetailsInput with task_id

    Returns:
        JSON with task info, description, and checklist items
    """
    try:
        async with httpx.AsyncClient(timeout=GRAPH_TIMEOUT) as client:
            task = await graph_get(client, f"/planner/tasks/{params.task_id}")
            details = await graph_get(client, f"/planner/tasks/{params.task_id}/details")

            # Parse checklist
            checklist = []
            for item_id, item in details.get("checklist", {}).items():
                checklist.append({
                    "id": item_id,
                    "title": item.get("title", ""),
                    "isChecked": item.get("isChecked", False),
                })

            result = {
                "id": task["id"],
                "title": task["title"],
                "percentComplete": task["percentComplete"],
                "dueDateTime": task.get("dueDateTime"),
                "bucketId": task["bucketId"],
                "priority": task.get("priority", 5),
                "createdDateTime": task.get("createdDateTime"),
                "description": details.get("description", ""),
                "checklist": checklist,
                "referenceCount": task.get("referenceCount", 0),
            }

            return json.dumps(result, indent=2)

    except httpx.HTTPStatusError as e:
        return json.dumps({"error": f"HTTP {e.response.status_code}: {e.response.text}"})
    except Exception as e:
        return json.dumps({"error": f"{type(e).__name__}: {str(e)}"})


@mcp.tool(
    name="planner_create_task",
    annotations={
        "title": "Create Planner Task",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True,
    },
)
async def planner_create_task(params: CreateTaskInput) -> str:
    """Create a new task in a Planner plan.

    Args:
        params: CreateTaskInput with plan_id, bucket_id, title, optional due_date

    Returns:
        JSON with the created task's id, title, and bucket
    """
    plan_id = params.plan_id or DEFAULT_PLAN_ID
    if not plan_id:
        return json.dumps({"error": "No plan_id provided and PLANNER_DEFAULT_PLAN_ID not set in .env"})

    try:
        body = {
            "planId": plan_id,
            "bucketId": params.bucket_id,
            "title": params.title,
        }
        if params.due_date:
            body["dueDateTime"] = f"{params.due_date}T23:59:00Z"

        async with httpx.AsyncClient(timeout=GRAPH_TIMEOUT) as client:
            result = await graph_post(client, "/planner/tasks", body)

            return json.dumps({
                "status": "created",
                "id": result["id"],
                "title": result["title"],
                "bucketId": result["bucketId"],
                "dueDateTime": result.get("dueDateTime"),
            }, indent=2)

    except httpx.HTTPStatusError as e:
        return json.dumps({"error": f"HTTP {e.response.status_code}: {e.response.text}"})
    except Exception as e:
        return json.dumps({"error": f"{type(e).__name__}: {str(e)}"})


@mcp.tool(
    name="planner_update_task",
    annotations={
        "title": "Update Planner Task",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def planner_update_task(params: UpdateTaskInput) -> str:
    """Update a task's properties (title, bucket, due date, progress, priority).

    Handles etag retrieval automatically.

    Args:
        params: UpdateTaskInput with task_id and fields to update

    Returns:
        JSON with update status
    """
    try:
        body = {}
        if params.title is not None:
            body["title"] = params.title
        if params.bucket_id is not None:
            body["bucketId"] = params.bucket_id
        if params.due_date is not None:
            if params.due_date.lower() == "clear":
                body["dueDateTime"] = None
            else:
                body["dueDateTime"] = f"{params.due_date}T23:59:00Z"
        if params.percent_complete is not None:
            body["percentComplete"] = params.percent_complete
        if params.priority is not None:
            body["priority"] = params.priority

        if not body:
            return json.dumps({"error": "No fields to update. Provide at least one field."})

        async with httpx.AsyncClient(timeout=GRAPH_TIMEOUT) as client:
            etag = await get_task_etag(client, params.task_id)
            await graph_patch(client, f"/planner/tasks/{params.task_id}", body, etag)

            return json.dumps({
                "status": "updated",
                "task_id": params.task_id,
                "updated_fields": list(body.keys()),
            }, indent=2)

    except httpx.HTTPStatusError as e:
        return json.dumps({"error": f"HTTP {e.response.status_code}: {e.response.text}"})
    except Exception as e:
        return json.dumps({"error": f"{type(e).__name__}: {str(e)}"})


@mcp.tool(
    name="planner_complete_task",
    annotations={
        "title": "Complete Planner Task",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def planner_complete_task(params: CompleteTaskInput) -> str:
    """Mark a task as 100% complete.

    Handles etag retrieval automatically.

    Args:
        params: CompleteTaskInput with task_id

    Returns:
        JSON with completion status
    """
    try:
        async with httpx.AsyncClient(timeout=GRAPH_TIMEOUT) as client:
            etag = await get_task_etag(client, params.task_id)
            await graph_patch(
                client,
                f"/planner/tasks/{params.task_id}",
                {"percentComplete": 100},
                etag,
            )

            return json.dumps({
                "status": "completed",
                "task_id": params.task_id,
            }, indent=2)

    except httpx.HTTPStatusError as e:
        return json.dumps({"error": f"HTTP {e.response.status_code}: {e.response.text}"})
    except Exception as e:
        return json.dumps({"error": f"{type(e).__name__}: {str(e)}"})


@mcp.tool(
    name="planner_update_details",
    annotations={
        "title": "Update Task Details",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True,
    },
)
async def planner_update_details(params: UpdateTaskDetailsInput) -> str:
    """Update a task's description and/or add checklist items.

    Args:
        params: UpdateTaskDetailsInput with task_id, optional description and checklist_add

    Returns:
        JSON with update status
    """
    try:
        async with httpx.AsyncClient(timeout=GRAPH_TIMEOUT) as client:
            # Get current details for etag
            details = await graph_get(client, f"/planner/tasks/{params.task_id}/details")
            etag = details["@odata.etag"]

            body = {}
            if params.description is not None:
                body["description"] = params.description
                body["previewType"] = "description"

            if params.checklist_add:
                import uuid
                checklist = {}
                for item_text in params.checklist_add:
                    item_id = str(uuid.uuid4())
                    checklist[item_id] = {
                        "@odata.type": "#microsoft.graph.plannerChecklistItem",
                        "title": item_text,
                        "isChecked": False,
                    }
                body["checklist"] = checklist

            if not body:
                return json.dumps({"error": "No fields to update. Provide description or checklist_add."})

            await graph_patch(
                client,
                f"/planner/tasks/{params.task_id}/details",
                body,
                etag,
            )

            return json.dumps({
                "status": "updated",
                "task_id": params.task_id,
                "updated_fields": list(body.keys()),
            }, indent=2)

    except httpx.HTTPStatusError as e:
        return json.dumps({"error": f"HTTP {e.response.status_code}: {e.response.text}"})
    except Exception as e:
        return json.dumps({"error": f"{type(e).__name__}: {str(e)}"})


@mcp.tool(
    name="planner_list_buckets",
    annotations={
        "title": "List Planner Buckets",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def planner_list_buckets(params: ListBucketsInput) -> str:
    """List all buckets in a Planner plan.

    Args:
        params: ListBucketsInput with optional plan_id

    Returns:
        JSON list of buckets with id and name
    """
    plan_id = params.plan_id or DEFAULT_PLAN_ID
    if not plan_id:
        return json.dumps({"error": "No plan_id provided and PLANNER_DEFAULT_PLAN_ID not set in .env"})

    try:
        async with httpx.AsyncClient(timeout=GRAPH_TIMEOUT) as client:
            data = await graph_get(client, f"/planner/plans/{plan_id}/buckets")
            buckets = data.get("value", [])

            result = []
            for b in buckets:
                result.append({
                    "id": b["id"],
                    "name": b["name"],
                })

            return json.dumps({"count": len(result), "buckets": result}, indent=2)

    except httpx.HTTPStatusError as e:
        return json.dumps({"error": f"HTTP {e.response.status_code}: {e.response.text}"})
    except Exception as e:
        return json.dumps({"error": f"{type(e).__name__}: {str(e)}"})


@mcp.tool(
    name="planner_create_bucket",
    annotations={
        "title": "Create Planner Bucket",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True,
    },
)
async def planner_create_bucket(params: CreateBucketInput) -> str:
    """Create a new bucket in a Planner plan.

    Args:
        params: CreateBucketInput with plan_id and name

    Returns:
        JSON with the created bucket's id and name
    """
    plan_id = params.plan_id or DEFAULT_PLAN_ID
    if not plan_id:
        return json.dumps({"error": "No plan_id provided and PLANNER_DEFAULT_PLAN_ID not set in .env"})

    try:
        body = {
            "name": params.name,
            "planId": plan_id,
            "orderHint": " !",
        }

        async with httpx.AsyncClient(timeout=GRAPH_TIMEOUT) as client:
            result = await graph_post(client, "/planner/buckets", body)

            return json.dumps({
                "status": "created",
                "id": result["id"],
                "name": result["name"],
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
