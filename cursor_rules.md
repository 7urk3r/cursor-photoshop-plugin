# Cursor Plugin Development Rules

<CORE_PRINCIPLES>
- Never rush to conclusions — explore multiple angles first.
- Think in simple, natural language and show all reasoning.
- Break every idea into the smallest logical steps.
- Question everything, including assumptions and previous conclusions.
- Revise freely if needed, especially when uncertain.
</CORE_PRINCIPLES>

<DEVELOPMENT RULES>

✅ Build atomically — one feature per commit.
✅ Every feature must be verified in one or more ways:
  - Log confirmation
  - UI change visible in screenshot
  - Output file created

✅ After any code change:
  - Log what happened with prefix `[Cursor OK]`
  - If using MCP, sendLog + optionally screenshot

✅ Use the following files for context:
  - plugin-spec.md
  - file-io.md
  - mcp-protocol.md
  - cursor-feedback-loop.md

✅ Reuse logic when possible — share CSV loader, logger, etc.

✅ Layer naming:
  - Must match CSV headers like `text1`, `img1`
  - Never guess layer names — always confirm presence

✅ Logging:
  - Always update log panel AND memory buffer
  - Write `log.txt` at end of every process cycle

✅ UI:
  - Render each tab with its own `<div>` or section
  - Tabs should toggle via buttons or tabs UI
  - Display current file paths and settings on screen for debugging

<DEBUGGING STRATEGY>
If Photoshop is unresponsive or plugin does not load:
- Restart via MCP (`restartPlugin`)
- Confirm via MCP screenshot or OCR
- Retry logic once bug is fixed

<ERROR HANDLING>
- If CSV fails to load or is missing required columns:
  - Show error in UI
  - Log error to log.txt and MCP
- Always catch exceptions from file and layer operations