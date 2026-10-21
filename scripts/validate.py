#!/usr/bin/env python3
"""Check the repository before a release: marketplace.json is valid and points at
existing plugins, every plugin has a plugin.json whose name matches its folder,
every standalone skill has a SKILL.md with name and description, and no private
material (reference/ folders, build outputs) is tracked.
Usage: python3 scripts/validate.py
"""
import json
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
errors = []

market_path = ROOT / ".claude-plugin/marketplace.json"
try:
    market = json.loads(market_path.read_text())
except Exception as e:  # noqa: BLE001
    errors.append(f"marketplace.json unreadable: {e}")
    market = {"plugins": []}
listed = set()
for entry in market.get("plugins", []):
    src = ROOT / entry.get("source", "")
    listed.add(src.name)
    if not (src / ".claude-plugin/plugin.json").exists():
        errors.append(f"marketplace entry '{entry.get('name')}' points at {entry.get('source')} which has no .claude-plugin/plugin.json")
    if "version" in entry:
        errors.append(f"marketplace entry '{entry.get('name')}' sets a version; versions live in the plugin's plugin.json")

for p in sorted(d for d in (ROOT / "plugins").iterdir() if d.is_dir()):
    manifest = p / ".claude-plugin/plugin.json"
    if not manifest.exists():
        errors.append(f"plugins/{p.name}: missing .claude-plugin/plugin.json")
        continue
    data = json.loads(manifest.read_text())
    if data.get("name") != p.name:
        errors.append(f"plugins/{p.name}: plugin.json name '{data.get('name')}' differs from folder name")
    if not re.match(r"^\d+\.\d+\.\d+$", str(data.get("version", ""))):
        errors.append(f"plugins/{p.name}: version '{data.get('version')}' is not semver")
    if p.name not in listed:
        errors.append(f"plugins/{p.name}: not listed in marketplace.json")


def frontmatter(path):
    text = path.read_text(errors="ignore")
    m = re.match(r"^---\n(.*?)\n---\n", text, re.S)
    return m.group(1) if m else None


skills = [d for d in (ROOT / "skills").iterdir() if d.is_dir()]
for s in sorted(skills):
    skill = s / "SKILL.md"
    if not skill.exists():
        errors.append(f"skills/{s.name}: missing SKILL.md")
        continue
    fm = frontmatter(skill)
    if fm is None:
        errors.append(f"skills/{s.name}: SKILL.md has no YAML frontmatter")
        continue
    for key in ("name", "description"):
        if not re.search(rf"^{key}:\s*\S", fm, re.M):
            errors.append(f"skills/{s.name}: frontmatter lacks '{key}'")
    name = re.search(r"^name:\s*(.+)$", fm, re.M)
    if name and name.group(1).strip().strip("\"'") != s.name:
        errors.append(f"skills/{s.name}: frontmatter name '{name.group(1).strip()}' differs from folder name")

tracked = subprocess.run(["git", "-C", str(ROOT), "ls-files"], capture_output=True, text=True).stdout.split("\n")
for f in tracked:
    parts = Path(f).parts
    if "reference" in parts or ("tests" in parts and "out" in parts) or f.endswith(".DS_Store"):
        errors.append(f"private or build file is tracked: {f}")

if errors:
    print("\n".join(f"ERROR: {e}" for e in errors))
    sys.exit(1)
print(f"ok: {len(listed)} plugin(s), {len(skills)} standalone skill(s)")
