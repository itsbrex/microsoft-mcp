---
description: Verify saved MSAL tokens match their filenames (JWT check)
allowed-tools: ["Bash(uv run:*)"]
argument-hint: "[--live]"
---

Decode the JWT payload of every saved access token and compare its `upn` claim to the filename identifier. Catches mislabeled token files (e.g., authenticated while `MICROSOFT_MCP_ACCOUNT_ID` pointed at the wrong account).

- no arg → JWT-only check (fast, offline)
- `--live` → additionally call Graph `/me` for each token (throttled, opt-in)

Exit code: 0 if all match, 1 if any mismatch.

Arguments passed: `$ARGUMENTS`

!`uv run auth_refresh.py --verify $ARGUMENTS`
