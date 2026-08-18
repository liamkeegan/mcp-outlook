# Outlook MCP (Model Context Protocol) Server

A Model Context Protocol server that integrates with Microsoft Outlook through Microsoft Graph API, allowing Claude and other LLMs to check calendar events, schedule new ones, read emails, and send messages.

## Features

- 📅 **Calendar Integration**: View, list, create, update, and delete calendar events
- 📧 **Email Integration**: Read, send, draft, and manage emails from your Outlook account
- 🔁 **Model Context Protocol**: Follows MCP standards for LLM tool integration
- 🛡️ **Type Safety**: Full TypeScript implementation with Zod validation

## Prerequisites

- Node.js 18+
- Microsoft 365 account with appropriate permissions (commercial, GCC, GCC High, or DoD tenants are all supported)
- Microsoft Azure App Registration with Graph API permissions (Calendar and Mail)

## Setup


1. Register an application in Azure Active Directory:
   - Go to [Azure Portal](https://portal.azure.com)
   - Navigate to "App registrations"
   - Create a new registration with a redirect URI of type "Public client/native (mobile & desktop)"
     - Register `http://localhost` as the redirect URI

   - Configure API permissions:
     - Choose Microsoft Graph and type delegated, as we will act on the users behalf
     - For Calendar: "Calendars.Read" and "Calendars.ReadWrite"
     - For Email: "Mail.Read", "Mail.ReadWrite" and "Mail.Send"
     - For People: "People.Read"

   > **GCC High / DoD tenants**: the app registration is done at [https://portal.azure.us](https://portal.azure.us) (or the Entra admin center at [https://entra.microsoft.us](https://entra.microsoft.us)) instead of the commercial portal. The steps, the `http://localhost` public-client redirect URI, and the delegated Microsoft Graph permissions are all the same. Government tenants commonly disable user consent, so a tenant admin may need to grant admin consent for the app's delegated permissions before sign-in will succeed.

2. Note the values from your Azure app registration (Overview) to use for the MCP config as environment variables:
  - Client ID (Application (client) ID)
  - Authority ID (Directory (tenant) ID)

3. Register the MCP server
For Claude Desktop, create or update your Claude Desktop `claude_desktop_config.json` (macOS: `~/Library/Application Support/Claude/claude_desktop_config.json`, Windows: `%APPDATA%\Claude\claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "outlook": {
      "command": "npx",
      "args": [
        "mcp_outlook"
      ],
      "env": {
        "AUTHORITY": "your-authority-id",
        "CLIENT_ID": "your-client-id"
      }
    }
  }
}
```

For GCC High / DoD tenants add `"M365_CLOUD": "gcc-high"` (or `"dod"`) to the env block — see the US Government clouds section below. Make sure to replace the path and environment variables with your actual values.

## Microsoft 365 GCC High / DoD (US Government clouds)

Set `M365_CLOUD` to select which Microsoft cloud the server signs in to and calls Graph on. Values are case-insensitive and `_`/`-` are interchangeable.

| `M365_CLOUD` value | Aliases | Login endpoint | Graph endpoint |
| --- | --- | --- | --- |
| `global` (default when unset) | `commercial`, `public`, `gcc` | `https://login.microsoftonline.com` | `https://graph.microsoft.com` |
| `gcc-high` | `gcchigh`, `usgov`, `usgovhigh` | `https://login.microsoftonline.us` | `https://graph.microsoft.us` |
| `dod` | `usgovdod` | `https://login.microsoftonline.us` | `https://dod-graph.microsoft.us` |

> **GCC (Moderate) tenants use the GLOBAL endpoints.** Set nothing, or `M365_CLOUD=gcc` (which maps to `global`). Only GCC High and DoD use the `.us` endpoints.

Example Claude Desktop config for a GCC High tenant:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "npx",
      "args": [
        "mcp_outlook"
      ],
      "env": {
        "AUTHORITY": "your-authority-id",
        "CLIENT_ID": "your-client-id",
        "M365_CLOUD": "gcc-high"
      }
    }
  }
}
```

An invalid value causes the server to exit at startup with:
```
outlook-mcp: Invalid M365_CLOUD "<value>". Expected one of: global, gcc-high, dod (aliases: commercial, public, gcc, gcchigh, usgov, usgovhigh, usgovdod).
```

### Advanced overrides

For sovereign clouds not covered by the `M365_CLOUD` presets (e.g. 21Vianet China), or to point at a custom endpoint, use:

- `GRAPH_BASE_URL`: overrides the Graph origin on top of the selected preset (e.g. `https://microsoftgraph.chinacloudapi.cn`). Must be an https origin with no path; a trailing slash or missing scheme is tolerated.
- `LOGIN_AUTHORITY_HOST`: overrides the login origin (e.g. `https://login.chinacloudapi.cn` for China). Same rules as above.

`GRAPH_BASE_URL` is used both as the request origin and as the OAuth resource for the requested scopes, so it must be the real Microsoft Graph origin for that cloud — pointing it at a proxy will produce tokens with the wrong audience.

### Confirming which cloud is active

On startup the server prints one line to stderr, for example:
```
outlook-mcp: Microsoft cloud "gcc-high" (login https://login.microsoftonline.us, graph https://graph.microsoft.us)
```
In Claude Desktop, this line is visible in the MCP server logs.

### Known limitations / caveats

- People API (`searchPeople`/`getPerson`) results and `$search` relevance may be reduced in GCC High/DoD, since some relevance/insight signals are limited in government clouds.
- Admin consent is commonly required in government tenants before delegated permissions can be used.
- No HTTP(S) proxy support for Graph calls.
- The DoD preset is configured but not yet verified end-to-end by the maintainers.
- Feature availability differs between clouds — see [Microsoft Graph deployments](https://learn.microsoft.com/graph/deployments).

## Available Tools

### Calendar Tools
- **listCalendarEvents**: Lists the user's calendar events for a specified time range
- **createCalendarEvent**: Creates a new calendar event
- **getCalendarEvent**: Gets details of a specific calendar event
- **updateCalendarEvent**: Updates an existing calendar event
- **deleteCalendarEvent**: Deletes a calendar event

### Email Tools
- **listEmails**: Lists emails from a specified folder (inbox, sent, drafts, etc.)
- **getEmail**: Gets details of a specific email message
- **sendEmail**: Sends a new email message
- **createDraft**: Creates a draft email message without sending it
- **markEmailAsRead**: Marks an email message as read
- **markEmailAsUnread**: Marks an email message as unread
- **deleteEmail**: Deletes an email message

## People Tools
- **searchPeople** Find a person in the organisation or your recent contacts
- **getPerson** Get information on a person

### Resources
- **calendar**: Resource containing calendar events data
- **inbox**: Resource containing inbox messages data

## Development

Run in development mode with live reloading:
```
npm run dev
```

Run linting:
```
npm run lint
```

Run tests (vitest unit tests in `test/`):
```
npm test
```

Configure your MCP locally
```json
{
  "mcpServers": {
    "outlook": {
      "command": "node",
      "args": [
        "/ABSOLUTE/PATH/TO/outlook_mcp/build/index.js"
      ],
      "env": {
        "AUTHORITY": "your-authority-id",
        "CLIENT_ID": "your-client-id"
      }
    }
  }
}
```

For GCC High / DoD tenants add `"M365_CLOUD": "gcc-high"` (or `"dod"`) to the env block — see the US Government clouds section below.
