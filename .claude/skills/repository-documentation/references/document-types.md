# Document types

Choose sections based on the repository, audience, and task. These patterns are defaults, not mandatory templates.

Use the four Diátaxis forms—tutorial, how-to guide, reference, and explanation—to classify a substantive page by its dominant reader need. Treat README files, landing pages, contribution guides, architecture decision records, policies, release notes, and similar files as entry points or specialized genres that can support one or more forms. Do not distort a useful repository convention merely to give every file a Diátaxis label.

See [audience and Diátaxis architecture](audience-and-diataxis.md) when choosing the hierarchy or handling content for multiple audiences.

## README

A README should let a new reader understand the project and reach a useful first success quickly.

Recommended order:

1. Project name and one-sentence purpose.
2. Status or maturity warning when it affects use.
3. Key capabilities or intended use cases.
4. Prerequisites.
5. Quickstart or installation.
6. Minimal usage example.
7. Links to deeper documentation.
8. Support, contributing, security, and license information when relevant.

Rules:

- Put the shortest useful path near the top.
- Explain who the project is for and what problem it solves.
- Distinguish a library, application, service, template, plugin, and internal tool clearly.
- Do not duplicate the full documentation site in the README.
- Do not add badges that do not provide useful status or navigation.
- Do not claim stability, compatibility, or performance without evidence.

## CONTRIBUTING.md

A contribution guide should remove uncertainty from preparing and submitting a change.

Include as relevant:

- Supported contribution types.
- Prerequisites and local setup.
- Exact development, build, test, lint, and documentation commands.
- Branch, commit, and pull-request conventions that the repository actually enforces.
- Issue-report expectations and reproduction requirements.
- Generated-file or dependency-update rules.
- Review expectations and required checks.
- Links to the code of conduct, security policy, architecture docs, and style guides.

Do not invent governance rules. Derive commands from the repository's automation and configuration.

## Architecture documentation

Architecture documentation should explain system boundaries, responsibilities, and decisions that are difficult to infer from individual files.

Include as relevant:

- Purpose, audience, and scope.
- System context and external dependencies.
- Constraints and quality attributes.
- Major components and their responsibilities.
- Data, control, event, or request flows.
- Public interfaces and integration boundaries.
- Persistence and state ownership.
- Security and trust boundaries.
- Failure modes, recovery, observability, and operational concerns.
- Deployment topology.
- Important decisions and links to ADRs.
- Known limitations and planned changes, clearly labeled.

Rules:

- Describe the current architecture first.
- Label proposed or target architecture explicitly.
- Use diagrams to clarify relationships, not replace prose. Follow [diagram guidance](diagrams.md) to choose a format and provide an accessible text equivalent.
- Keep component names consistent with the codebase.
- Record why boundaries exist when the reason is non-obvious.

## Architecture decision record

Follow the repository's ADR format when one exists. Otherwise, use:

- **Title:** The decision in specific terms.
- **Status:** Proposed, accepted, deprecated, or superseded.
- **Date:** The decision date.
- **Context:** The problem, forces, and constraints.
- **Decision:** What was chosen.
- **Consequences:** Positive, negative, and operational effects.
- **Alternatives considered:** Serious options and why they were not selected.
- **References:** Related issues, pull requests, experiments, or documents.

Keep accepted ADRs as historical records. Supersede them with a new ADR rather than rewriting the original decision to match current preferences.

## API reference and doc comments

Reference documentation should be complete, predictable, and optimized for lookup.

For each public type, member, endpoint, command, or operation, document as applicable:

- Purpose in the first sentence.
- Signature or syntax.
- Preconditions and required permissions.
- Parameters, fields, or inputs.
- Return value or output.
- Errors, exceptions, and failure behavior.
- Side effects and state changes.
- Defaults, limits, and accepted ranges.
- Threading, ordering, idempotency, or lifecycle behavior when relevant.
- A short realistic example.
- Related operations.
- Deprecation version, replacement, and migration action.

Rules:

- Cover every public element unless the repository explicitly defines a smaller supported surface.
- Do not repeat information that is obvious from a signature unless it clarifies semantics or constraints.
- Begin method descriptions with a specific present-tense verb such as **Creates**, **Returns**, **Updates**, **Deletes**, **Registers**, or **Validates**.
- For boolean parameters, state the behavior for both `true` and `false` when it is not obvious.
- Keep generated reference documentation in source comments or schemas, not in manually edited output.
- Follow the language's doc-comment conventions in addition to this skill.

## CLI reference

Document:

- Command purpose.
- Syntax.
- Required arguments.
- Options and defaults.
- Environment variables and configuration precedence.
- Working-directory requirements.
- Exit codes.
- Input and output formats.
- Examples for common tasks.
- Destructive behavior and confirmation rules.

Prefer one canonical syntax block and task-based examples. Use semantic placeholder markup in command syntax, and explain any placeholder whose expected value, format, or source is not evident from context.

## Configuration reference

Document:

- Configuration location and supported formats.
- Complete key hierarchy.
- Type, default, valid values, and required status for each key.
- Precedence among files, environment variables, CLI options, and built-in defaults.
- Reload or restart behavior.
- Security implications and secret-handling requirements.
- Version or migration requirements.
- Minimal and representative examples.

Do not present a sample configuration as complete unless it is complete.

## How-to guide

A how-to guide helps a reader complete one real task.

Include:

- Goal.
- Preconditions and prerequisites.
- Starting state.
- Ordered steps.
- Expected result.
- Verification.
- Troubleshooting or rollback only when likely to be needed.

Keep conceptual explanation brief and link to deeper background.

## Tutorial

A tutorial teaches through a guided, successful experience.

Include:

- What the reader will build or learn.
- Prerequisites.
- A sequence that builds understanding gradually.
- Explanations immediately after the actions they clarify.
- Checkpoints and expected output.
- A final result and suggested next steps.

Do not use a tutorial as exhaustive reference documentation.

## Concept or explanation document

Use a concept document to explain how or why something works.

Include as relevant:

- Definition and scope.
- Motivation or problem context.
- Mental model.
- Relationships among concepts.
- Tradeoffs and constraints.
- Examples and counterexamples.
- Links to related tasks and reference material.

Do not bury procedures inside long conceptual sections; link to a dedicated task when the procedure is substantial.

## Troubleshooting guide

Organize troubleshooting around observable symptoms, not internal component names alone.

For each issue, include:

- Symptom or error message.
- Likely causes.
- Diagnostic checks.
- Resolution steps.
- Verification.
- Escalation information or logs to collect when the issue persists.

Order issues by frequency, severity, or workflow stage. Preserve exact error text with `<samp>` when the target renderer supports it; otherwise, use code formatting.

## CHANGELOG and release notes

Describe user-visible impact, not the commit history.

- Follow the repository's existing release process and categories.
- State what changed and why it matters.
- Identify breaking changes and required migration actions prominently.
- Link to detailed migration guides for complex changes.
- Use exact versions and release dates.
- Keep unreleased changes clearly separate from released versions.
- Exclude internal refactors unless they affect users, contributors, operators, compatibility, security, or performance.

## SECURITY.md

A security policy should state:

- Supported versions.
- How to report a vulnerability privately.
- What information to include.
- Expected acknowledgment or response process only when the maintainers can uphold it.
- Disclosure expectations.
- Out-of-scope reports when needed.

Do not ask reporters to disclose vulnerabilities in public issues.

## Agent instruction files

Files such as `AGENTS.md`, `CLAUDE.md`, and vendor-specific instruction directories are documentation whose audience is a coding agent. Ambient rules belong in these files; reusable procedures belong in skills. Apply the same rules as for human documentation, with these specifics:

- Treat `AGENTS.md` as the shared, cross-agent home. Make agent-specific files pointers to it — for example, a `CLAUDE.md` containing `@AGENTS.md` plus only agent-specific additions — rather than parallel copies.
- Include only rules that must influence every task: conventions the repository enforces, facts agents repeatedly get wrong, and constraints such as files that must not be edited.
- Do not restate facts that live in `CONTRIBUTING.md`, configuration, or automation; link to them. Instruction files drift like any other copy.
- Keep the file short. Every line spends context on every task, so a rule that rarely matters costs more than it saves.

## Code comments and docstrings

Use comments and docstrings to preserve information that code alone cannot express.

Document:

- Contracts and public behavior.
- Non-obvious intent and rationale.
- Invariants and constraints.
- Side effects and failure behavior.
- Workarounds with issue or source references.
- Security-sensitive assumptions.

Do not narrate obvious syntax. Keep comments synchronized with the code. Follow the language's standard format and the repository's established conventions.
