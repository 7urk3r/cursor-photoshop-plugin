# MCP Interface Overview

## Purpose
Allow external control over Photoshop via custom server.

## Capabilities
- `restartPlugin` – quits and relaunches Photoshop via AppleScript
- `sendLog(message)` – writes log to Cursor and internal log buffer
- `createPlugin`, `updatePluginFile` – allows code injection/rebuild

## Communication
- Via WebSocket between MCP and Cursor
- Supports logs, screenshots, OCR, auto-debug

## Example Command
```json
{
  "command": "restartPlugin"
}
```