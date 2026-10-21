# ArrowEye skills

Free, public skills and plugins by [ArrowEye Automations](https://arroweye.be). Each one is a folder with a `SKILL.md` that tells an AI agent how to do one job well. The format is the open [Agent Skills](https://agentskills.io) standard, so the same folder works in Claude (claude.ai, Claude Code, Cowork), OpenAI Codex, Cursor, GitHub Copilot and other tools that read `SKILL.md`.

Two kinds of content live here:

| folder | what | install with |
|---|---|---|
| `skills/<name>/` | standalone skills: one `SKILL.md`, optional scripts and references | copy the folder, or download its zip from [Releases](../../releases) |
| `plugins/<name>/` | Claude Code plugins: skills, slash commands and agents bundled together | `/plugin install <name>@arroweye-skills` in Claude Code |

## Plugins (Claude Code)

| plugin | what it does |
|---|---|
| `financial-plan-builder` | Legal financial plan (art. 5:4 WVV / CSA) for incorporating a Belgian BV / SRL, in Dutch, French or English: interview, Excel model and Word document for the notary, with an independent review. |

Install in Claude Code (once for the marketplace, then per plugin):

```
/plugin marketplace add ArrowEyeBE/arroweye-skills
/plugin install financial-plan-builder@arroweye-skills
```

Updates arrive when Claude Code refreshes the marketplace, or with `/plugin marketplace update arroweye-skills`.

## Standalone skills

None published yet. Single-folder skills will appear in `skills/`; each one's `SKILL.md` says when to use it, and every release will carry a zip per skill.

## How to install a skill in your tool

Download the skill's zip from [Releases](../../releases) (or clone this repository) and put the skill folder where your tool looks for skills:

| tool | where the skill folder goes |
|---|---|
| **claude.ai / Claude desktop** | Settings → Capabilities → Skills → upload the zip |
| **Claude Code** | `~/.claude/skills/<name>/` (all projects) or `.claude/skills/<name>/` (one project) |
| **Claude Cowork** | Skills panel → add skill → choose the zip or folder |
| **OpenAI Codex** (CLI, IDE, app) | `~/.codex/skills/<name>/` (all projects) or `.agents/skills/<name>/` (one project) |
| **ChatGPT custom GPT** | paste the body of `SKILL.md` into the GPT's instructions and attach the files in `references/` as knowledge |
| **Cursor** | `~/.cursor/skills/<name>/` or `.agents/skills/<name>/` |
| **GitHub Copilot** | `.github/skills/<name>/` or `.agents/skills/<name>/` in the repository; `~/.copilot/skills/<name>/` for yourself |
| **other agents** | any tool that reads the Agent Skills standard: point it at the folder |

The folder name must stay the same as the `name` in the `SKILL.md` frontmatter. Skills with `scripts/` need Python 3; each skill's `SKILL.md` lists what it needs.

## Releases

Every release attaches one zip per skill and per plugin. A zip unpacks to a single folder named after the skill, ready to upload or copy.

## Contributing and support

Open an issue for a bug or a request. Pull requests are welcome for fixes; for new skills, open an issue first so we can agree on scope. Skills are provided as-is under the MIT licence (see `LICENSE`) and are not tax, legal or accounting advice.

## For maintainers

- `python3 scripts/validate.py` before every commit: checks manifests, frontmatter and that no private material is tracked.
- `bash scripts/add_skill.sh /path/to/skill` copies a skill in; review it for private names before committing.
- Bump the version in the plugin's `.claude-plugin/plugin.json`, then tag `vX.Y.Z` and push the tag: the workflow in `.github/workflows/release.yml` validates, zips and publishes the release.
