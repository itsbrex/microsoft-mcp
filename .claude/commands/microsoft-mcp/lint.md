---
description: Type-check with pyright and lint with ruff
allowed-tools: ["Bash(uv run pyright:*)", "Bash(uvx ruff:*)"]
---

Run the same static checks CI would run.

!`uv run pyright`
!`uvx ruff check .`
