#!/usr/bin/env bash
# PostToolUse hook: auto-format edited Python files with ruff.
# Receives the tool call JSON on stdin. Extracts file_path(s) and runs ruff.
# Non-zero exit would surface a hook failure to Claude; we swallow errors so a
# bad file never blocks the model.
set -u

payload="$(cat)"
project_dir="${CLAUDE_PROJECT_DIR:-$(pwd)}"

# Build a newline-separated list of candidate file paths from the tool input.
paths=$(
  printf '%s' "$payload" | jq -r '
    (.tool_input.file_path // empty),
    (.tool_input.edits[]?.file_path // empty)
  ' 2>/dev/null | sort -u
)

[ -z "$paths" ] && exit 0

while IFS= read -r f; do
  [ -z "$f" ] && continue
  case "$f" in
    *.py) ;;
    *) continue ;;
  esac
  # Only operate on files inside the project to avoid touching unrelated code.
  case "$f" in
    "$project_dir"/*) ;;
    *) continue ;;
  esac
  [ -f "$f" ] || continue

  uvx ruff format "$f" >/dev/null 2>&1 || true
  uvx ruff check --fix --unsafe-fixes "$f" >/dev/null 2>&1 || true
done <<< "$paths"

exit 0
