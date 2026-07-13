# Lessons

## Lesson: Verify repository tool entry points before validation

### Trigger
Initial validation assumed `uv` was on PATH and referenced removed `tests/test_tools.py`.

### Root Pattern
Used documented commands and an old test filename without checking current shell and test inventory.

### Rule
Resolve required binaries and enumerate relevant test files before launching validation.

### Current Application
Use absolute `uv` path and current split test modules for this review.

## Lesson: Run type checker inside project environment

### Trigger
`uvx pyright` produced a large false missing-import report because isolated tool environment could not see project dependencies.

### Root Pattern
Tool executable existed, but dependency resolution context was wrong.

### Rule
Run Pyright through project environment (`uv run --with pyright pyright`) or declare it in the dev dependency group.

### Current Application
Run changed-file Pyright through `uv run --with pyright pyright`; track the repository-wide baseline cleanup and a permanent dev dependency separately.

## Lesson: Verify provider semantics before porting reference behavior

### Trigger
The imported Outlook refresh logic treated refresh tokens as resource-scoped and misclassified `AADSTS65002`.

### Root Pattern
Behavior from the reference repository was accepted without checking the Microsoft identity-platform contract.

### Rule
Verify OAuth token rotation and provider error codes against current official documentation before preserving reference-repository heuristics.

### Current Application
Persist the latest replacement refresh token from every successful resource refresh and direct `AADSTS65002` users to an appropriately authorized app registration.

## Lesson: Isolate authentication tests from saved accounts

### Trigger
The test suite inherited local Microsoft account and token-directory environment variables, while auth fixtures used a `cresa.com` identity.

### Root Pattern
Mocks prevented expected network calls, but test safety depended on every test patch remaining correct and allowed project imports to discover real local credential paths.

### Rule
Before project imports, force tests onto temporary token/cache paths, use non-production identities, disable interactive fallback, and hard-block protected account domains at the authentication boundary.

### Current Application
Pytest always blocks `cresa.com`, uses `cresa.email` fixtures, and runs with isolated MSAL, Azure, and outlook-creds storage paths.
