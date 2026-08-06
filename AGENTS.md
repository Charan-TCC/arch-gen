<!-- baton -->
## Project memory (Baton) — REQUIRED workflow
This project uses Baton MCP tools for cross-agent memory. You MUST use
them:

- At session start, call memory_load BEFORE exploring the codebase — it
  returns the trusted project digest; do not re-derive what it states.
- After completing ANY user-requested change (even small ones), call
  memory_checkpoint with the full updated digest (<=2000 tokens): current
  task, architecture, file map, decisions, tried-and-failed, next steps.
- Call memory_log immediately when you make a notable decision, hit a
  dead end, or leave something open (decision|blocker|todo).
- When the user says to hand off to another agent, call memory_handoff
  FIRST, then stop.

An out-of-date digest strands the next agent. Checkpointing is part of
completing the task, not optional cleanup.

NEVER read or write files under .agent-memory/ directly — always go
through the memory tools (direct edits skip the event log and break the
handoff trail). If a memory tool returns an error, report it to the user
instead of working around it.
<!-- /baton -->
