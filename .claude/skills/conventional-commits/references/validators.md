# Validator integrations

Use repository-owned commitlint or gitlint configuration before applying the bundled fallback policy. Do not install either tool, replace configuration, or add enforcement unless the user asks for that change.

## Select a validator

1. Run the validation command documented by the repository, even when it wraps the underlying linter.
2. Otherwise, use commitlint when the repository contains commitlint configuration and its required packages are available.
3. Otherwise, use gitlint when the repository contains `.gitlint` or contributor instructions that select gitlint and the command is available.
4. If both tools are configured and the repository does not establish precedence, run both when doing so is read-only. Report conflicting results instead of choosing a policy silently.
5. If no configured validator is available, use the manual checklist in `SKILL.md`.

Hook runners such as Husky and the pre-commit framework are not validators. Inspect their configuration to identify the commitlint or gitlint command, but prefer a direct, read-only validation command when checking a proposed message.

## Use commitlint

Recognize commitlint configuration in `.commitlintrc*`, `commitlint.config.*`, or a `commitlint` field in a package manifest. Prefer a repository package script when one exists because it can supply required arguments or environment settings.

Validate a proposed message by piping it to the repository's command or by passing a temporary message file with `--edit <message-file>`. Do not create a commit merely to test its message.

When the repository has `@commitlint/cli` but no commitlint policy, run the bundled validator from the repository root:

```text
node <skill-directory>/scripts/validate-commit-message.mjs
```

The validator reads standard input unless commitlint arguments such as `--edit <message-file>` are supplied. It uses the adjacent [bundled configuration](../scripts/commitlint.config.mjs) and resolves `@commitlint/cli` from the repository that contains the working directory, which it finds by locating the nearest `.git`.

The validator also exits with status 2, before running commitlint, when a line of the message read from standard input exceeds 2000 characters. Commitlint's parser does not finish on a line that long, so the check reports the line instead of hanging.

The validator exits with status 2 when it cannot find `@commitlint/cli` inside that repository, including when Node.js module resolution would otherwise reach an installation in a parent directory. A neighboring checkout's commitlint can hold a different version and a different policy, so reaching one silently would report a verdict the repository never established. Treat status 2 as a signal to perform the manual review instead. Status 0 means the message passed and status 1 means commitlint rejected it.

When the user asks to configure commitlint enforcement and the repository does not define a conflicting policy:

1. Copy the bundled configuration into the repository's established configuration location, or import and extend it as described below.
2. Install `@commitlint/cli` as a development dependency with the repository's existing package manager.
3. Integrate the validator with an existing commit-message hook or CI mechanism only when that enforcement surface is in scope.
4. Test representative accepted and rejected messages.

Do not copy the bundled configuration merely to validate one message.

### Extend the footer allowlist

A repository whose tooling writes a trailer the bundled list does not name, such as a session link an authoring tool stamps on each commit, admits it by importing the bundled configuration rather than editing a copy:

```javascript
import { createConfig } from "./.agents/skills/conventional-commits/scripts/commitlint.config.mjs";

export default createConfig({ footerTokens: ["Example-Session"] });
```

`createConfig` returns the bundled policy with the extra tokens allowed and named in its reports. A token that is not one word of letters, digits, and hyphens makes the configuration fail to load, so a mistake is reported rather than admitting nothing. Point the repository's commitlint command at that file; the bundled validator script reads the bundled configuration alone.

## Use gitlint

Treat `.gitlint` and repository contributor instructions as the source of truth. Validate a proposed message from a temporary file:

```text
gitlint --msg-filename <message-file>
```

Gitlint's general rules do not enable Conventional Commits automatically. Check whether `.gitlint` enables `contrib-title-conventional-commits` before claiming that it validates the Conventional Commits header. Even with that rule enabled, manually review this skill's semantic type selection, breaking-change migration footer, footer allowlist, and reference requirements.

When the user asks to configure gitlint enforcement and the repository does not define a conflicting policy:

1. Add gitlint through the repository's existing Python dependency workflow.
2. Create or update `.gitlint` and enable `contrib-title-conventional-commits` with the repository's allowed types.
3. Configure the title length and other mechanical rules to match the repository policy.
4. Integrate the validator with an existing commit-message hook or CI mechanism only when that enforcement surface is in scope.
5. Test representative accepted and rejected messages.

Do not synthesize custom gitlint rules merely to reproduce every semantic rule in this skill. Keep the manual review for checks that the configured linter cannot express reliably.
