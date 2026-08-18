# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build Commands
- `npm run build` - Build the TypeScript project
- `npm test` - Run vitest unit tests in `test/`
- `npm run lint` - Lint the code (when implemented)
- `npm start` - Start the MCP server (when implemented)

## Code Style Guidelines
- **TypeScript**: Use strict typing with explicit return types (test files under `test/` are exempt from explicit return types — enforced via an ESLint override)
- **Formatting**: Follow 2-space indentation, trailing commas
- **Imports**: Group by external packages first, then internal modules
- **Naming**: camelCase for variables/functions, PascalCase for classes/types
- **Error Handling**: Use typed error responses when possible
- **Modules**: Use ES modules (type: "module" is set in package.json)
- **SDK Usage**: Follow @modelcontextprotocol/sdk patterns for tools and resources

## Project Structure
- `/src` - TypeScript source files
- `/build` - Compiled JavaScript output
- `/test` - Vitest unit tests

## Microsoft Cloud Configuration
- `src/cloudConfig.ts` is the single source of truth for Microsoft endpoints (login authority + Graph base URL), selected via the `M365_CLOUD` env var. Never hard-code `graph.microsoft.com` / `login.microsoftonline.com` elsewhere in the codebase.
- Never log to stdout in the server — stdio is the MCP transport. Use `console.error` for all logging/diagnostics.

This project is a model context protocol server for Microsoft Outlook. It allows Claude to:

1. **Calendar functionality**:
   - Check calendar events
   - Schedule new events
   - Update existing events
   - Delete events

2. **Email functionality**:
   - Read emails from inbox and other folders
   - Send new emails
   - Create draft emails
   - Mark emails as read/unread
   - Delete emails

The server uses the Microsoft Graph API to interact with Outlook's calendar and email systems.

