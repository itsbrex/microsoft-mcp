---
description: Commit current changes, push the branch, and open a PR
allowed-tools: ["Bash(git:*)", "Bash(gh:*)"]
---

Walk through the standard end-of-feature flow:

1. `git status` + `git diff` — show what's changing.
2. Review the diff for: secrets (tokens, .env keys), debug code, unrelated churn, missing tests, docs drift (CLAUDE.md / IMPLEMENTATION.md / README.md).
3. Draft a concise commit message focused on the *why*. Conventional commits (`feat:`, `fix:`, `refactor:`, `docs:`, `test:`). Never include "generated with" footers.
4. `git add` specific files (avoid `-A`/`.`) and commit via HEREDOC.
5. Push with `-u` if the upstream doesn't exist yet.
6. Open a PR with `gh pr create`, title under 70 chars, body in the format:

```
## Summary
- bullet
- bullet

## Test plan
- [ ] `/test`
- [ ] `/lint`
- [ ] manual verification of <golden path>
```

7. Return the PR URL.

Never run `git push --force` or `gh pr merge` without explicit user confirmation.
