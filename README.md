# M365 Planner MCP Server

MCP (Model Context Protocol) server for Microsoft 365 Planner integration with Claude Code and Claude Desktop.

Enables AI assistants to manage Planner tasks, buckets, and plans via Microsoft Graph API.

## Features

### Task Management
| Tool | Description |
|------|-------------|
| `planner_list_tasks` | List tasks in a plan (filter by bucket, completion status) |
| `planner_my_tasks` | List all my tasks across all plans |
| `planner_overdue` | List overdue tasks |
| `planner_get_details` | Get task details including description and checklist |
| `planner_create_task` | Create a new task |
| `planner_update_task` | Update task properties (title, bucket, due date, priority, progress) |
| `planner_complete_task` | Mark a task as complete |
| `planner_update_details` | Update task description and checklist items |

### Bucket Management
| Tool | Description |
|------|-------------|
| `planner_list_buckets` | List all buckets in a plan |
| `planner_create_bucket` | Create a new bucket |

## Prerequisites

### 1. Azure CLI
Install and authenticate:
```bash
# Install (Windows)
winget install Microsoft.AzureCLI

# Login
az login

# Verify
az account show
```

### 2. Python 3.10+
```bash
pip install -r requirements.txt
```

### 3. Microsoft 365 Access
You need a Microsoft 365 account with Planner access (Business Basic or higher).

## Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/avellinosantiago/m365-mcp.git
   cd m365-mcp
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Configure environment**
   ```bash
   cp .env.example .env
   # Edit .env with your default Planner Plan ID
   ```

4. **Configure Claude Code/Desktop** (see below)

## Configuration

### For Claude Code CLI

```bash
claude mcp add -s user -e "PATH=/path/to/azure/cli" m365-planner -- python /path/to/m365-mcp/m365_planner_mcp.py
```

Or add to your Claude settings (`~/.claude.json`):

```json
{
  "mcpServers": {
    "m365-planner": {
      "command": "python",
      "args": ["/path/to/m365-mcp/m365_planner_mcp.py"],
      "env": {}
    }
  }
}
```

### For Claude Desktop

**macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
**Windows**: `%APPDATA%\Claude\claude_desktop_config.json`

```json
{
  "mcpServers": {
    "m365-planner": {
      "command": "python",
      "args": ["/path/to/m365-mcp/m365_planner_mcp.py"],
      "env": {}
    }
  }
}
```

## Environment Variables

Create a `.env` file in the project root (see `.env.example`):

| Variable | Required | Description |
|----------|----------|-------------|
| `PLANNER_DEFAULT_PLAN_ID` | Yes | Your default Planner Plan ID |
| `GRAPH_TIMEOUT` | No | HTTP timeout in seconds (default: 30) |

## Usage Examples

Once configured, ask Claude to:

### Task management
> "Show me all my open tasks"
> "What tasks are overdue?"
> "Create a task called 'Review Q1 report' in the Marketing bucket"
> "Mark task XYZ as complete"
> "Move that task to the In Progress bucket"

### Task details
> "Show me the checklist for task XYZ"
> "Add a description to that task"
> "Add checklist items: review docs, send email, schedule meeting"

### Buckets
> "List all buckets in the plan"
> "Create a new bucket called 'Sprint 5'"

## Authentication

Uses Azure CLI for authentication via `DefaultAzureCredential`. Tokens are automatically managed.

If you get authentication errors:
1. Run `az login`
2. Select the correct tenant/subscription
3. Restart Claude Code/Desktop

## How It Works

- **Etag handling**: Planner API requires `If-Match` headers with etags for updates. This server handles etag retrieval automatically - you never need to worry about it.
- **Pagination**: Large task lists are automatically paginated.
- **Default plan**: Set `PLANNER_DEFAULT_PLAN_ID` to avoid passing plan_id on every call.

## Troubleshooting

### "PLANNER_DEFAULT_PLAN_ID not set"
Create a `.env` file with your Plan ID. See `.env.example`.

### "Failed to get Azure CLI token"
Run `az login` and authenticate with your Microsoft 365 account.

### "403 Forbidden"
Your account doesn't have Planner access. Check your M365 license.

### "404 Not Found"
- Verify the plan/task/bucket ID is correct
- Use `planner_list_*` tools to get valid IDs

## Project Structure

```
m365-mcp/
├── m365_planner_mcp.py  # Planner MCP server
├── requirements.txt     # Python dependencies
├── .env.example         # Environment template
├── .env                 # Your configuration (not committed)
├── .gitignore           # Git ignore rules
└── README.md            # This file
```

## Roadmap

Future MCP servers planned for this repo:
- **Outlook** - Email search, drafts, folder management
- **SharePoint** - File navigation, upload/download, list management

## License

MIT License
