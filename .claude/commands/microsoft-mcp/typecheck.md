---
description: Type-check the codebase with pyright
allowed-tools: ["Bash(uv run:*)"]
argument-hint: "[path]"
---

Run pyright over the project (or the given path). Report and fix any new errors introduced by recent edits; do not chase pre-existing unrelated noise without asking.

Arguments passed: `$ARGUMENTS`

!`uv run pyright $ARGUMENTS`
