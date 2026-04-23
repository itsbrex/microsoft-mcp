#!/usr/bin/env bash
# Status line: <repo>  <branch> <model> <ctx%>  <cwd-rel>
# Reads Claude Code's JSON on stdin.
set -u

payload="$(cat)"

model=$(printf '%s' "$payload" | jq -r '.model.display_name // .model.id // "claude"')
cwd=$(printf '%s' "$payload" | jq -r '.workspace.current_dir // .cwd // "."')
project_dir=$(printf '%s' "$payload" | jq -r '.workspace.project_dir // ""')
session_name=$(printf '%s' "$payload" | jq -r '.session.name // empty')
tokens_used=$(printf '%s' "$payload" | jq -r '.cost.total_input_tokens // 0')
ctx_limit=$(printf '%s' "$payload" | jq -r '.model.context_limit // 200000')

repo="microsoft-mcp"
if [ -n "$project_dir" ]; then
  repo=$(basename "$project_dir")
fi

branch="?"
if git -C "${project_dir:-.}" rev-parse --abbrev-ref HEAD >/dev/null 2>&1; then
  branch=$(git -C "${project_dir:-.}" rev-parse --abbrev-ref HEAD 2>/dev/null)
fi

dirty=""
if git -C "${project_dir:-.}" status --porcelain 2>/dev/null | grep -q .; then
  dirty="*"
fi

rel_cwd="${cwd#$project_dir/}"
[ "$rel_cwd" = "$cwd" ] && rel_cwd="."

ctx_pct=0
if [ "$ctx_limit" -gt 0 ] 2>/dev/null; then
  ctx_pct=$(( tokens_used * 100 / ctx_limit ))
fi

# ANSI: dim=2, cyan=36, yellow=33, green=32, magenta=35, reset=0
printf '\033[2m%s\033[0m \033[36m⎇ %s%s\033[0m \033[35m%s\033[0m \033[33m%d%%\033[0m \033[2m· %s\033[0m' \
  "$repo" "$branch" "$dirty" "$model" "$ctx_pct" "$rel_cwd"

if [ -n "$session_name" ]; then
  printf ' \033[2m[%s]\033[0m' "$session_name"
fi
