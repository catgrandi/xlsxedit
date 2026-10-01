# Agent instructions

These instructions apply to all work in this repository. `CLAUDE.md` imports
this file; keep shared guidance here so every agent reads one source. This file
only points to where each fact lives.

- Follow the [documentation-sync rule](.agents/skills/repository-documentation/assets/agent-instructions.md):
  when a change alters observable behavior, the documentation stating that
  behavior is part of the change, and every fact keeps one authoritative home.
- Follow [CONTRIBUTING.md](CONTRIBUTING.md) for development setup, tests, commit
  and merge conventions, the fork policy, where specs and working files go, and
  how installed skills are upgraded.

## Sources of truth

- [README.md](README.md): what xlsxedit is and how to install it.
- [docs/how-xlsxedit-works.md](docs/how-xlsxedit-works.md): architecture and
  the preservation model.
- [docs/features.md](docs/features.md): public API reference.
- [docs/excel-xlsx-structure.md](docs/excel-xlsx-structure.md): OOXML
  package reference.
- [docs/large-data-export.md](docs/large-data-export.md) and
  [docs/pandas.md](docs/pandas.md): bulk export and the pandas engine.
- [docs/design/](docs/design/README.md): specs and design documents.
- [.github/workflows/](.github/workflows/): CI and release automation.
- [Roadmap tracking issue](https://github.com/catgrandi/xlsxedit/issues/26):
  planned work.

## Committed skills

Skills live in [.agents/skills/](.agents/skills/), mirrored to
`.claude/skills/` as described in [CONTRIBUTING.md](CONTRIBUTING.md#agent-skills).

- [conventional-commits](.agents/skills/conventional-commits/SKILL.md): commit
  messages.
- [repository-documentation](.agents/skills/repository-documentation/SKILL.md):
  documentation work.
- [release-management](.agents/skills/release-management/SKILL.md): releases
  and versioning.
- [xlsxedit](.agents/skills/xlsxedit/SKILL.md): using the library.
