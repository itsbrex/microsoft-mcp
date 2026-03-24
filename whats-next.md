# Handoff Document: Microsoft MCP "Logged in as: unknown" Fix

## Original Task

Debug and fix the issue where the Microsoft MCP installer displayed "Logged in as: unknown" after successful MSAL authentication, instead of showing the actual authenticated user's email/name.

## Work Completed

### Bug Investigation & Fix

1. **Root Cause Identified**: The `MSALRefreshTokenAuth.authenticate()` method in `src/microsoft_mcp/auth_msal.py` returned the raw MSAL response dictionary, which doesn't have a top-level `"username"` key. The installer script (`scripts/install.sh:806,819`) expected `result.get("username")` but the username was nested in `result["id_token_claims"]["preferred_username"]`.

2. **Fix Applied to `src/microsoft_mcp/auth_msal.py`**:
   - **Device code flow path** (line 442-443): Added `result["username"] = email or self.account_identifier` after extracting email from `id_token_claims`
   - **Silent auth path** (line 362-363): Added `result["username"] = cached_username or self.account_identifier` using the username from the cached MSAL account

3. **Tests Updated in `tests/test_auth_msal.py`**:
   - Added `assert result["username"] == "test@example.com"` to `test_authenticate_calls_msal` (line 468)
   - Added `assert result["username"] == "test@example.com"` to `test_authenticate_uses_cached_account` (line 493)

4. **All 81 tests pass** with Python 3.13

### Commits Created (5 total on `feature/alternative-auth-method` branch)

| Commit | Message |
|--------|---------|
| `da2d215` | fix: return username from MSAL authenticate() method |
| `d941fec` | feat: add multi-account management tools and fix logging |
| `9d9acd4` | feat: add multi-account setup to installer |
| `e080af7` | docs: add MCP configuration format examples |
| `6d3a810` | chore: update Python version to 3.13 and refresh lockfile |

### Additional Changes Committed (pre-existing uncommitted work)

- **Multi-account tools** in `tools.py`: `list_accounts()`, `set_active_account()`, `get_active_account()`
- **Logging fix** in `tools.py`: Removed file logging, log to stderr only (MCP protocol requirement)
- **Installer enhancements** in `install.sh`: Multi-account setup flow, full uv path resolution, defensive scripting
- **Documentation** in `CLAUDE.md`: MCP configuration format examples for MSAL and Azure SDK auth
- **Python version**: Updated from 3.12 to 3.13

## Work Remaining

### Immediate Next Steps

1. **Test the fix end-to-end**: Re-run the installer to verify "Logged in as: {email}" displays correctly
   ```bash
   ./scripts/install.sh
   ```

2. **Push commits to remote** (if ready for review):
   ```bash
   git push origin feature/alternative-auth-method
   ```

3. **Create PR** to merge `feature/alternative-auth-method` into `master`

### Optional Enhancements

- Consider adding the username to the Azure SDK auth path as well (for consistency)
- The `.cursorindexingignore` and `.specstory/` directories are untracked - decide if they should be gitignored

## Attempted Approaches

1. **Initial hypothesis**: Checked if the issue was in `install.sh` - confirmed the script was correctly using `result.get("username", "unknown")`

2. **Investigated MSAL response structure**: Found that MSAL's `acquire_token_by_device_flow()` returns:
   - `access_token`, `refresh_token`, `expires_in`, `scope` at top level
   - `id_token_claims.preferred_username` or `id_token_claims.email` for user identity
   - No `username` key at top level

3. **Considered two fix options**:
   - Option A: Modify `install.sh` to extract from `id_token_claims` (rejected - requires complex inline Python)
   - Option B: Modify `auth_msal.py` to add `username` to return dict (chosen - cleaner API)

## Critical Context

### Code Locations

- **Auth module**: `src/microsoft_mcp/auth_msal.py`
  - `authenticate()` method: lines 328-445
  - Silent auth path: lines 345-364
  - Device code flow: lines 366-445

- **Installer**: `scripts/install.sh`
  - Authentication loop: lines 789-826
  - Username display: lines 806, 819

- **Tests**: `tests/test_auth_msal.py`
  - Device code test: `test_authenticate_calls_msal` (line 440)
  - Silent auth test: `test_authenticate_uses_cached_account` (line 471)

### Environment Notes

- **Python 3.14 incompatibility**: `onnxruntime` doesn't have wheels for Python 3.14 yet; must use Python 3.13
- **Test command**: `uv run --python 3.13 pytest tests/ -v`
- **MCP protocol**: Servers must log to stderr (stdout is for JSON-RPC)

### Token Storage

- Location: `~/.config/microsoft-mcp/tokens/`
- Format: `{account_id}_access_token.json`, `{account_id}_refresh_only.txt`, `{account_id}_access_only.txt`

## Current State

### Deliverables Status

| Item | Status |
|------|--------|
| Bug fix (auth_msal.py) | ✅ Complete |
| Test updates | ✅ Complete |
| All tests passing | ✅ Complete (81/81) |
| Commits created | ✅ Complete (5 commits) |
| Pushed to remote | ⏳ Not done |
| PR created | ⏳ Not done |
| End-to-end verification | ⏳ Not done |

### Branch State

- **Current branch**: `feature/alternative-auth-method`
- **Ahead of master by**: 11 commits (6 previous + 5 new)
- **Untracked files**: `.cursorindexingignore`, `.specstory/` (editor artifacts)

### Open Questions

1. Should the Azure SDK auth path (`auth.py`) also return a `username` field for consistency?
2. Should the untracked editor files be added to `.gitignore`?
