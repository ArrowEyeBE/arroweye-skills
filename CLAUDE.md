# arroweye-skills

Public repository of ArrowEye Automations: agent skills and Claude Code plugins shared with our audience. Everything here is public and MIT-licensed; never commit client material, private names, prices or anything from a `reference/` folder.

- `skills/<name>/SKILL.md`: standalone skills in the Agent Skills format (work in Claude, Codex, Cursor, Copilot, ...). Folder name must equal the frontmatter `name`.
- `plugins/<name>/`: Claude Code plugins with their own `.claude-plugin/plugin.json`, listed in `.claude-plugin/marketplace.json` (marketplace name `arroweye-skills`). Bump the version in the plugin's `plugin.json`, never in the marketplace entry.
- `scripts/validate.py`: run before every commit. `scripts/package.sh`: builds one zip per skill and plugin in `dist/`. `scripts/add_skill.sh <folder>`: copies a skill in.
- Releasing: tag `vX.Y.Z` and push the tag; the workflow attaches the zips to a GitHub Release.
- Try a plugin locally: `claude --plugin-dir plugins/<name>`; validate the marketplace with `claude plugin validate .`.
