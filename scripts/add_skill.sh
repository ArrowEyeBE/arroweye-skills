#!/usr/bin/env bash
# Copy a skill folder (for example one from ~/.claude/skills/<name>) into skills/,
# leaving out OS junk and caches. Review the copy for private names, paths and
# client material before committing it.
# Usage: bash scripts/add_skill.sh /path/to/skill-folder
set -euo pipefail
[ $# -eq 1 ] && [ -f "$1/SKILL.md" ] || { echo "usage: $0 /path/to/skill-folder (must contain SKILL.md)"; exit 1; }
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
name="$(basename "$1")"
rsync -a --exclude '.DS_Store' --exclude '__pycache__/' --exclude '*.pyc' --exclude 'node_modules/' --exclude 'reference/' "$1/" "$ROOT/skills/$name/"
echo "copied to skills/$name. Now: check it for private content, then run python3 scripts/validate.py"
