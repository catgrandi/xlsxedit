# Documentation-sync rule for coding agents

This skill loads only when documentation is the task, so it cannot influence an agent that is editing code. To apply the rule below during code changes, wire it into instructions your agents load on every task.

Default to one line in the repository's `AGENTS.md`, which most agent ecosystems read directly. Point agent-specific files at that same source instead of repeating it — for example, a `CLAUDE.md` whose content is `@AGENTS.md` plus any Claude-specific additions. Use a vendor-specific location such as `.github/instructions/` only when an agent in use reads nothing else. Link to this file's installed location; if a format cannot link, copy the rule with a comment naming this file as its source and update the copy when this file changes.

This file is the rule's authoritative home.

---

When a change alters observable behavior — commands, flags, configuration keys, APIs, defaults, output, errors, or documented workflows — the documentation that states that behavior is part of the change, not follow-up work.

- If you own the complete change, update the affected documentation in the same unit of work. The `repository-documentation` skill covers how.
- If you are delivering one part of a larger coordinated change, do not edit shared documentation mid-flight. List each affected document and what it must say once the change lands, and surface that list to whoever owns the complete change.
- Give every fact one authoritative home. Never restate a value or contract that already has one; link to it.
