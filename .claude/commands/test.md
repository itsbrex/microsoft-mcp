---
description: Run pytest (all or a subset) with verbose output
argument-hint: "[pytest args, e.g. tests/test_auth.py -k msal]"
allowed-tools: ["Bash(uv run pytest:*)"]
---

Run the test suite. Pass any extra pytest args via `$ARGUMENTS`. Default is the full suite.

!`uv run pytest tests/ -v $ARGUMENTS`
