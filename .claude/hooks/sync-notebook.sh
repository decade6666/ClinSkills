#!/bin/bash
# Auto-sync scripts/*.py -> notebooks/*.ipynb after Claude Code edits.
# Receives tool input JSON on stdin from PostToolUse hook.

TOOL_INPUT=$(cat)
FILE_PATH=$(echo "$TOOL_INPUT" | python -c "import sys,json; d=json.load(sys.stdin); print(d.get('tool_input',{}).get('file_path',''))" 2>/dev/null)

# Only act on scripts/*.py files
[[ -z "$FILE_PATH" || "$FILE_PATH" != *scripts* || "$FILE_PATH" != *.py ]] && exit 0

# Derive notebook path
SCRIPT_NAME=$(basename "$FILE_PATH" .py)
PROJECT_ROOT=$(cd "$(dirname "$FILE_PATH")/.." && pwd)
NB_PATH="$PROJECT_ROOT/notebooks/$SCRIPT_NAME.ipynb"

[[ ! -f "$NB_PATH" ]] && exit 0

echo "jupytext sync: scripts/$SCRIPT_NAME.py -> notebooks/$SCRIPT_NAME.ipynb"
cd "$PROJECT_ROOT"
jupytext --sync "notebooks/$SCRIPT_NAME.ipynb" 2>&1
