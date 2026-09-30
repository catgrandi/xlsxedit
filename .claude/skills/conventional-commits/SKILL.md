---
description: Write, review, and validate Conventional Commit messages and organize changes into logical commits using explicit rules for types, scopes, headers, bodies, breaking changes, footers, and references. Use when drafting a commit message, choosing a commit type or scope, reviewing commit history, preparing a commit plan, committing changes, running commit-message validation, or configuring commitlint or gitlint enforcement in a repository that adopts these conventions.
metadata:
    github-path: skills/conventional-commits
    github-pinned: agent-skills-v1.5.0
    github-ref: refs/tags/agent-skills-v1.5.0
    github-repo: https://github.com/catgrandi/agent-skills
    github-tree-sha: eedfb7bba07b70b548a4ac8eed059883203496f9
    version: 0.4.0
name: conventional-commits
---
# Conventional Commits

Create concise commit messages that describe one logical change and follow the repository's Conventional Commits policy.

## Apply the rules

1. Inspect the complete change represented by the commit. Prefer the staged diff when files are staged; otherwise, use the change set identified by the user or repository workflow.
2. Separate unrelated changes into distinct commits. Do not hide multiple logical changes behind a broad description.
3. Select the first matching type from the ordered list below.
4. Add a scope only when it tells the reader something the surrounding context does not. See [choose a scope](#choose-a-scope).
5. Write the header, then add a body or footers only when they provide necessary context.
6. Validate the finished message with repository tooling when available, then review the semantic rules that tooling cannot determine.

Follow explicit user requirements and more specific repository policy when they conflict with this skill. Do not stage files, create commits, rewrite history, or split changes unless the user has authorized that action.

## Format the header

Use this format:

```text
<type>[(scope)][!]: <description>
```

Apply all of these constraints:

- Choose exactly one allowed type: `feat`, `fix`, `docs`, `style`, `refactor`, `perf`, `test`, `build`, `ci`, `chore`, or `revert`.
- Keep the complete header at 72 characters or fewer; target 50 characters or fewer.
- Write the description in the imperative mood.
- Do not start the description with a capital letter. A digit or a backticked identifier is correct when it is the description's natural first word.
- Do not end the description with a period.
- Use an optional scope that is a noun naming a module or component, written in lowercase letters and digits joined by single `-` or `_` separators. Do not name a file, an extension, or a path.
- Add `!` immediately before `:` when the commit contains a breaking change.

## Choose a scope

A scope is worth its cost when it answers a question the reader would otherwise have. Decide by where the description is read, not by whether the repository holds several packages.

Use a scope when the change history is shared:

- A repository with one changelog covering everything, including a monorepo released as a single unit. The scope is the only marker telling a reader which part a line refers to.
- A repository whose release tooling groups or routes entries by scope.
- A single package large enough that its areas are worth distinguishing, such as `parser`, `cli`, or `auth`.

Omit the scope when its answer is already on the page:

- Each package keeps its own changelog. A line in that file cannot be about any other package, so naming the package repeats its filename.
- Release tooling selects a package by the paths a commit touches rather than by the scope, which makes the scope decorative for the release and load-bearing only for a reader.

Weigh the cost. The header allows 72 characters and targets 50, and the scope is drawn from the same budget as the description that becomes the changelog entry. A scope such as `(repository-documentation)` spends 26 characters restating a filename, leaving too few to describe the change for someone deciding whether to upgrade. Prefer the description.

A scope naming an area inside the package stays useful even when the package name would not, because the reader cannot infer it from the changelog they are already reading.

## Select the type

Use the first category that matches the commit's primary purpose. Categories 2 through 7 name the material a commit touches, and they apply when that material is itself the point of the change. When a change to that material exists to correct or extend the behavior the product ships, continue to categories 8 and 9 instead: a dependency bump that stops a crash users hit is `fix`, not `build`, because its purpose is the fix and its release should carry one.

1. Revert a prior commit: `revert`.
2. Change only documentation: `docs`.
3. Change only formatting without affecting behavior: `style`.
4. Change only tests: `test`.
5. Change GitHub Actions, CI pipelines, or automation: `ci`.
6. Change build tooling, dependencies, lockfiles, or build scripts: `build`.
7. Change workspace, editor, administration, or housekeeping files: `chore`.
8. Correct a bug, regression, or other incorrect behavior: `fix`.
9. Add user-visible behavior or a public API: `feat`.
10. Improve performance without changing behavior: `perf`.
11. Restructure, rename, or reorganize code without changing behavior: `refactor`.
12. Use `chore` when no earlier category applies.

When uncertain, choose the more conservative applicable type, especially `ci`, `build`, or `chore` instead of `feat`. This tie-break settles changes that leave shipped behavior alone; it never redirects a change that corrects or extends that behavior away from `fix` or `feat`.

## Add a body

Omit the body when the header fully explains the change. Otherwise:

- Separate the body from the header with a blank line.
- Explain rationale, important implementation context, or impact that is not obvious from the header.
- Keep the body concise and technical.
- Wrap body lines at 72 characters. Leave a long URL or a pasted identifier whole on its own line rather than breaking it.
- Use bullets when they improve clarity.
- Do not repeat the header in prose.
- Do not open any body line with a single word followed by a colon. Commit parsers read such a line as the start of the footers, so a mid-paragraph `mechanically: ...` continuation breaks validation and a final `Before:` and `After:` paragraph is read as trailers. Rewrap so the colon sits mid-line, or fold the lines into a sentence or a bulleted list.

## Mark breaking changes

For every breaking change, use both markers:

1. Add `!` before the header's colon.
2. Add a `BREAKING CHANGE:` footer that explains what changed and how to migrate.

Example:

```text
feat(api)!: replace legacy authentication

BREAKING CHANGE: Replace API keys with OAuth access tokens.
```

## Add footers and references

Separate footers from the body, or from the header when there is no body, with a blank line. Use only these footer tokens:

- `BREAKING CHANGE`
- `BREAKING-CHANGE`
- `Fixes`
- `Closes`
- `Resolves`
- `Refs`
- `See`
- `Co-Authored-By`
- `Reviewed-By`
- `Signed-Off-By`
- `Acked-By`
- `Reverts`
- `Deprecated`

A repository may admit further trailers that its tooling writes, such as a session link an authoring tool stamps on a commit, through its own validator configuration. Treat a token that configuration admits as allowed, and do not write such a trailer by hand. [Validator integrations](references/validators.md) gives the form that configuration takes.

Write each trailer as `<Token>: <value>`. Match a token case-insensitively, so `Signed-off-by:` as git writes it and `Co-authored-by:` as GitHub writes it are both correct; write `BREAKING CHANGE` and `BREAKING-CHANGE` in uppercase, which the specification requires.

A footer value may run onto the following lines, so wrap a long migration explanation the way you wrap the body. The value continues until the next line that opens with a token, which is why a wrapped line must not begin with a word followed by a colon.

The specification also allows a `<Token> #<reference>` separator. Do not use it. Write `Fixes: #123` rather than `Fixes #123` so every footer in a message has one shape.

Keep footers in the trailing paragraphs of the message. A paragraph of prose after a footer paragraph turns those trailers back into body text, and tools that read trailers will not find them.

Format references as follows:

- Use `#123` for an issue or pull request in the same repository.
- Use `owner/repository#123` for an issue or pull request in another repository.
- Use `@username` for a user.
- Use at least 7 hexadecimal characters for a commit SHA.
- Use the full 40-character SHA in a `Reverts:` footer when it is available.
- Wrap file paths and code identifiers in backticks.

## Review the result

Before returning or using a message, confirm that:

- The commit contains one logical change.
- The selected type is the first applicable type in the ordered rules.
- The scope, when present, is useful and correctly formatted.
- The header is within the length limit and its description follows the grammar rules.
- The body adds information instead of restating the header.
- Every breaking change has both `!` and a migration footer.
- Every footer uses an allowed token and valid reference format, and the footers occupy the trailing paragraphs.

## Validate with repository tooling

Before creating a commit, inspect the repository for contributor instructions, documented validation commands, commitlint or gitlint configuration, development dependencies, and commit-message hooks.

1. Prefer the repository's documented validation command and existing validator configuration.
2. Validate the proposed message through standard input or a temporary message file before committing when the tooling supports it.
3. If validation fails, correct the message or report a material conflict between the repository policy and this skill.
4. Do not install dependencies, replace configuration, or create hooks unless the user asks you to configure enforcement.
5. When no validator exists, apply the review checklist in this skill manually.

Read [validator integrations](references/validators.md) when a repository uses commitlint or gitlint, when choosing between both configured tools, or when the user asks to configure either validator.

Treat validator output as a mechanical check, not a substitute for judgment. A linter cannot reliably determine whether a commit contains one logical change, uses the correct type for its meaning, has a genuinely imperative description, uses a useful scope, or explains a breaking migration adequately.

Return the proposed commit message in a code block when the user asks only for message text. When the user asks for a commit plan, list each logical commit with its included files or changes and its proposed message.

## Examples

```text
docs: explain selective skill installation
```

```text
fix(catalog): reject duplicate skill names

Report every conflicting entry so maintainers can correct the catalog in one pass.

Fixes: #123
```

```text
revert: restore previous catalog validation

Reverts: 0123456789abcdef0123456789abcdef01234567
```
