# Contributing to xlsxedit

Thanks for your interest in improving xlsxedit! Contributions of all kinds are
welcome: bug reports, documentation, tests, and code.

## Fork policy

This repository is a fork of
[jonas-kupferschmid/xlsxedit](https://github.com/jonas-kupferschmid/xlsxedit).
It does not send anything upstream: no pull requests, issues, or pushes to the
upstream repository. Changes land here only.

## Contributor License Agreement

[CLA.md](CLA.md) is the upstream project's agreement. It applies to
contributions submitted to the upstream repository, where you sign it by adding
this line to your pull request description, along with your full name and the
date:

> I have read the CLA document and I hereby sign the CLA.

Because this fork does not send changes upstream, pull requests to this fork
do not sign it.

## Development setup

```bash
git clone https://github.com/<your-username>/xlsxedit.git
cd xlsxedit
python -m venv .venv && source .venv/bin/activate
pip install -e ".[dev,pandas]"
```

The `pandas` extra is optional. Without it, the pandas tests are skipped.

## Running the tests

```bash
pytest
```

Please make sure the full test suite passes before opening a pull request, and
add tests for any new behavior.

## Commit messages

Write every commit message in the Conventional Commits format defined by the
[conventional-commits skill](.agents/skills/conventional-commits/SKILL.md).
That skill is the authoritative statement of the rules.

## Merging pull requests

Keep history linear. Merge pull requests with **rebase** or **squash** only;
never create a merge commit.

## Guidelines

- Keep changes focused; one logical change per pull request.
- Match the existing code style (type hints, no unnecessary comments).
- Preserve template fidelity: edits should change only what the API targets and
  leave unrelated package parts untouched.
- If you touch the OPC layer (`src/xlsxedit/opc/**`), remember parts of it are
  adapted from python-docx / python-pptx (MIT); keep the attribution intact.
- When a change alters observable behavior, update the documentation that
  states that behavior in the same change. See the
  [documentation-sync rule](.agents/skills/repository-documentation/assets/agent-instructions.md).

## Specs and working files

- Formal specs and design documents go in [`docs/design/`](docs/design/README.md)
  and are committed.
- Agent working files (plans, task briefs, reports, progress notes) go in
  `.agents.local/`, which is gitignored and never committed.

## Agent skills

Committed agent skills live in `.agents/skills/`. Each one is mirrored byte for
byte to `.claude/skills/`; change both copies together and confirm that
`diff -r .agents/skills .claude/skills` prints nothing.

The `conventional-commits`, `release-management`, and
`repository-documentation` skills are installed from
[catgrandi/agent-skills](https://github.com/catgrandi/agent-skills) at a pinned
release. Never edit them by hand. To upgrade one, reinstall it at the new
release tag, then copy it again to `.claude/skills/`:

```bash
gh skill install catgrandi/agent-skills <skill> --agent codex --scope project --pin agent-skills-vX.Y.Z --force
rm -rf .claude/skills/<skill> && cp -r .agents/skills/<skill> .claude/skills/<skill>
```

The `xlsxedit` skill is maintained in this repository.

## Reporting bugs

Open an issue with a minimal reproduction, the xlsxedit version, your Python
version, and (if possible) a sample `.xlsx` that triggers the problem.
