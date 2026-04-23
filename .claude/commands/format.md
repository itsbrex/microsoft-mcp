---
description: Format and auto-fix lint issues across the repo
allowed-tools: ["Bash(uvx ruff:*)"]
---

Format Python and auto-apply safe + unsafe lint fixes.

!`uvx ruff format .`
!`uvx ruff check --fix --unsafe-fixes .`
