#!/usr/bin/env bash
# Zip every skill in skills/ and every plugin in plugins/ into dist/, one zip each.
# Each zip unpacks to a single folder named after the skill or plugin, which is
# what claude.ai's skill upload and every tool's skills directory expect.
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
DIST="$ROOT/dist"
rm -rf "$DIST" && mkdir -p "$DIST"
for kind in skills plugins; do
  for dir in "$ROOT/$kind"/*/; do
    [ -d "$dir" ] || continue
    name="$(basename "$dir")"
    ( cd "$ROOT/$kind" && zip -qr "$DIST/$name.zip" "$name" \
        -x '*/.DS_Store' '*/__pycache__/*' '*.pyc' "$name/reference/*" "$name/tests/out/*" )
    echo "packaged $kind/$name -> dist/$name.zip"
  done
done
