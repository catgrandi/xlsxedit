# Release Please

What [Release Please](https://github.com/googleapis/release-please) adds to, narrows in, or makes concrete in the rules in `SKILL.md`. Read that file first. Nothing here replaces it, and a rule that would hold with the tool removed is stated there rather than repeated here.

Release Please parses Conventional Commit messages on a branch, proposes the next version in a pull request, and on merge writes the changelog, tags the commit, and creates a GitHub Release. It publishes to no package registry and implements no branching strategy for the project; the branches it creates are the ones its own pull requests ride on. Everything downstream of the tag — building, publishing, and confirming that consumers can resolve the result — belongs to the repository's own pipeline, and the core file governs it unchanged.

Two configurations stand behind the claims below: a grouped multi-component repository and a single-package Node repository. A claim that holds in both is stated plainly. A claim that depends on a configuration option names the option. Who may merge a proposal, which credentials the publish step uses, and where releases are announced are the repository's own process, and belong in its contribution guide.

## The proposal pull request is the release decision

Release Please does not release what lands on the branch. It maintains one open pull request that proposes the next release, and merging that pull request is what tags and releases.

- **The proposal accumulates.** It is recomputed and updated as further commits land, so leaving it open across several changes is the intended use rather than a backlog. By default it is only rewritten when the generated notes change; `always-update` rewrites it on every run.
- **Merging is the authorization.** The core file's rule that a release is an explicit decision has exactly one enactment here: a human merging the proposal. Nothing earlier in the pipeline commits to the release, and nothing later asks again.
- **A label records where the proposal is.** It opens carrying `autorelease: pending`; after the release is tagged that label is replaced with `autorelease: tagged`. Both are configurable, and publication tooling conventionally adds `autorelease: published` itself.
- **With one grouped proposal, every package in it releases together.** `separate-pull-requests` decides whether each package gets its own proposal or one proposal covers them all; under a grouped proposal, merging releases every package the body lists, at one commit.

## Do not edit the proposal body

The body is not documentation. It is both generated output and machine-read input, and editing it fails in both directions.

- **Before merge, the edit is lost.** Each run compares the body it would generate against the body that is there and pushes the generated one whenever they differ. A hand edit guarantees they differ, so the next run overwrites it.
- **After merge, the edit is what shipped.** Release Please parses the merged proposal's body to decide which components to release and to build each GitHub Release's notes. An edit that survives to the merge does not annotate the release; it becomes the release.

The core file says to fix the input a generator reads rather than its output. Release Please's input is the commit messages, and it offers one documented way to amend them after the fact: add a `BEGIN_COMMIT_OVERRIDE` / `END_COMMIT_OVERRIDE` block to the body of the merged pull request the commit came from — not the proposal — and the next run reads that block instead of the commit message. This works only where the source pull request was squash-merged, because a plain merge leaves Release Please unable to tell which commits the override applies to.

Once a version is released the input path is closed, and the correction is made where the record lives instead: edit the published GitHub Release, then land the matching changelog correction as its own commit, so the two records agree. Both are annotations attached to a published version, not a rewrite of what shipped.

## The changelog is prepend-only

The changelog updater finds the first version heading in the existing file, splices the new entry above it, and copies everything below it through unchanged. Released sections are never revisited.

That is why annotating a published section by hand is safe here, and it is the concrete instance of the core file's rule that a published record is corrected by addition. The first release's notes are the usual reason to use it: they describe whatever commits fall inside the bootstrap window rather than the product, which is rarely what a reader wants, and a hand-written replacement for that one section stays stable through every later release.

The same property is what makes a hand edit to *any* other release-managed file wrong. Versions, manifest entries, and extra-file annotations are recomputed; only the changelog's past is left alone.

## What causes a release

A release happens when the generated notes for a package are non-empty. Everything else is a consequence of that.

- **Only some commit types reach the notes.** `feat`, `fix`, `deps`, and breaking changes are releasable by default; `chore`, `build`, `ci`, `docs`, `refactor`, `style`, and `test` are not. Some release types add their own, and `changelog-sections` reassigns any of them, so the releasable set is a configuration fact and worth checking rather than assuming.
- **A run with nothing releasable is skipped,** and says so: `No user facing commits found since <sha> - skipping`. No proposal appears, and no proposal means no release.
- **`release-as` sets the version a release carries; it does not cause one to exist.** With `release-as` configured and nothing releasable in range, the run still skips. Remove or raise it once the proposal it was written for has merged, or every later run repeats that same version.
- **`bootstrap-sha` decides what is in range for the first proposal.** It is an exclusive lower bound, so it must sit one commit *before* the earliest commit to include, and it is ignored once Release Please has generated a release pull request of its own. Placing it after the first releasable commit produces no proposal at all.

## Pre-1.0 bumps are configuration

Under the default versioning strategy a breaking change bumps MAJOR at any version, including from `0.x`. Two options change that below `1.0.0`:

- `bump-minor-pre-major` makes a breaking change bump MINOR while the version is under `1.0.0`. The change still reaches the changelog with its breaking-change note.
- `bump-patch-for-minor-pre-major` makes a feature bump PATCH under the same condition.

The core file treats reaching `1.0.0` as a deliberate claim rather than a consequence. Here that claim has a mechanical form: while `bump-minor-pre-major` is set, no commit can promote the project past `0.x`, so removing it is the act that makes `1.0.0` reachable. Removing it is therefore the decision, and should be made when the stability claim is real rather than when a `feat!` happens to land.

## Merge style decides what reaches the changelog

Release Please derives entries from commit messages, so how a pull request is merged decides which messages exist to derive from. This constraint has no counterpart in the core file, because it exists only where the changelog is generated from history.

A squash merge collapses a branch's commits into one message, and the entries the collapsed messages would have produced are gone. Release Please recommends squash merging for exactly that reason: it keeps pull-request-local noise, such as a fix for a bug that never reached the branch, out of the notes. The same collapse is a loss where each commit was a separate consumer-visible change, and a rebase or fast-forward merge preserves them. Neither is correct in general; choose by whether the branch's messages are the entries you want.

Two things narrow the choice:

- **One squashed commit can still carry several entries.** Additional Conventional Commit messages placed at the bottom of the commit body are parsed as their own entries, each able to carry its own type, scope, and breaking-change footer.
- **In a multi-component repository, attribution follows the files a commit touches.** A squash merge attributes every file the branch touched to that single commit, so a branch spanning two components collapses into one entry and releases both from it. Where a branch carries more than one releasable commit across more than one component, rebase.

## Approve the proposal's workflow run

A proposal opened by `github-actions[bot]` can leave its checks unstarted at `action_required`, showing an approval control instead of check results. Approving it runs the workflow normally.

Treat this as a step in the release rather than a fault. Those checks are what covers the contents about to be released, so merging a proposal whose run never started releases something nothing validated, which the core file's readiness rule forbids.

## When a release is wanted and nothing qualifies

The core file forbids misdescribing a change to move a version number. Release Please provides levers that do the same job honestly, and the choice among them is the whole answer:

- **Wait.** Nothing releasable has landed, so nothing is owed to consumers yet. Cadence is a legitimate answer.
- **Set `release-as`** to name the version, provided something releasable is in range for the release to attach to.
- **Move `bootstrap-sha`,** before the first release only, so that a genuine user-facing commit already in history falls in range.

Retyping a `ci` or `build` change as `fix` also works, and that is the problem: it puts a line about repository plumbing into notes a consumer reads. A `0.1.0` shipped that way and was corrected afterwards by rewriting the section, when the same release could have been caused by the bootstrap move that commit already made.

## Components, tags, and path boundaries

In a manifest configuration each package is a path, and Release Please attributes a commit to a package by the files that commit touches. Scope in the commit message plays no part in that.

- **Tags carry the component by default:** `<component>-v<version>`. Setting `include-component-in-tag` to `false` gives `v<version>`. Where the repository's existing tags carry no component and the option is left true, Release Please looks for a tag nobody created and reports finding no releases, which reads as a missing release rather than as a tag-name mismatch.
- **The `.` package is the whole repository.** It releases when any change is made, subject to its own `exclude-paths`, which is what makes it usable as an aggregate over the other packages.
- **A package absent from the manifest is a package that has never been released.** Release Please fills in a missing entry itself, from `initial-version` if one is configured and from the release type's own default otherwise, and records the result on the first release. Hand-editing the manifest is appropriate only to seed a version Release Please could not know.
- **A release commit must change at least one file inside every package it releases.** Release Please bounds each package's history by finding the last release commit *within that package's own path-filtered commit list*; when the lookup misses, no cutoff is applied and every commit that path has ever had stays eligible. A package whose release artifacts all live outside its directory — a changelog, a version file, or a catalog held centrally — needs one file inside the directory that the release itself rewrites — an annotated version string updated through `extra-files` is the usual choice — or its released commits are proposed again indefinitely.
- **`exclude-paths` entries are literal directory prefixes, not globs.** Each entry is compared by testing whether a changed file starts with the entry plus a separator, so `docs` excludes `docs/`, `docs/**` matches nothing, and a file at the repository root cannot be excluded at all. A commit is skipped only when *every* relevant file it touches falls under an excluded path, so a commit spanning an excluded and a non-excluded file still counts.

## Failure modes

- **A proposal re-lists commits already under a package's tag.** This is the boundary failure above, seen from the outside. Recognize it by a package whose version moves although nothing has touched its path since its tag — confirm with `git merge-base --is-ancestor <commit> <tag>` and an empty `git log <tag>..<branch> -- <path>`. Close it and let the next run regenerate. Never merge it: merging spends a version number on no change and publishes notes that repeat a released entry. Closing does not stop it returning, because Release Please recomputes from history on every run; it clears when a real release for that package finally rewrites a file inside its directory. That clearing release is computed from the same missing cutoff, so its notes list the already-released commits beside the new ones, and merging it publishes them; correct the published record afterwards, as *Do not edit the proposal body* describes.
- **The release exists but the artifact never shipped.** Release Please tags and creates the GitHub Release, then hands off; a publish step that times out, is cancelled, or fails leaves a released version that reached no consumer. The core file's rule applies without modification — re-run distribution for the same ref, and do not bump the version because a step did not run. Guard the seam rather than the symptom: never cancel a run that has already tagged, and give the publish job more time than the validation inside it needs.
- **Draft releases hide their tags.** GitHub creates no Git tag for a draft release until it is published, so a configuration using `draft` leaves Release Please unable to find the previous release on the next run — which produces a changelog covering the entire history. `force-tag-creation` creates the tag immediately, which is what makes `draft` safe to use across runs.
- **A stale `autorelease: pending` label blocks the next proposal.** Where Release Please runs as the GitHub application, it will not open a new proposal while a pull request still carries that label. If a previous release left it behind, remove it and re-run.

## Sources

Each source was read at first hand on 2026-08-27, against release-please `v17.11.2`, the current 17.x and so the version the `^17.6.0` range in `release-please-action` v5.0.0 resolves to. Where behavior is observed rather than documented, the commit, pull request, or run that shows it is cited. Guidance that depends on one repository is not stated here at all.

Two configurations supply the observed evidence: [`catgrandi/agent-skills`](https://github.com/catgrandi/agent-skills), a grouped multi-component repository with component tags and an aggregate `.` package, and [`catgrandi/obsidian-compat`](https://github.com/catgrandi/obsidian-compat), a single-package Node repository. Neither repository's process is presented as the tool's behavior.

### The proposal pull request

- [Release Please README](https://github.com/googleapis/release-please/blob/v17.11.2/README.md), *What's a Release PR?*: release-please maintains Release PRs rather than continuously releasing what has landed; they are kept up-to-date as additional work is merged, and merging one updates the changelog, tags the commit, and creates a GitHub Release. The same section lists the `autorelease: pending`, `autorelease: tagged`, `autorelease: snapshot`, and `autorelease: published` states, and notes that release-please does not add the last of these itself.
- [Release Please README](https://github.com/googleapis/release-please/blob/v17.11.2/README.md), opening paragraph, states that release-please *does not handle publication to package managers or handle complex branch management*, which is the source for both exclusions in the opening paragraph above. The branches it does create carry its own pull requests: the proposal branch observed under *Bot-opened proposals* below, and the `release-please/bootstrap/` branch that [`src/bootstrapper.ts`](https://github.com/googleapis/release-please/blob/v17.11.2/src/bootstrapper.ts) names as `headBranchName` for the configuration pull request the `bootstrap` command opens, documented in [`cli.md`](https://github.com/googleapis/release-please/blob/v17.11.2/docs/cli.md), *Bootstrapping*.
- [Customizing releases](https://github.com/googleapis/release-please/blob/v17.11.2/docs/customizing.md), *Release Lifecycle Labels*, is the source for the pending-to-tagged transition and for both labels being configurable.
- [`manifest-releaser.md`](https://github.com/googleapis/release-please/blob/v17.11.2/docs/manifest-releaser.md) documents `separate-pull-requests`, defaulting to false so that one grouped pull request is raised, and `always-update`, which forces an update on every run instead of only when the release notes change.
- [`src/manifest.ts`](https://github.com/googleapis/release-please/blob/v17.11.2/src/manifest.ts), `maybeUpdateExistingPullRequest`, is the source for the default: it returns without pushing when `existing.body` equals the freshly generated body, and calls `updateExistingPullRequest` otherwise.
- obsidian-compat's `CONTRIBUTING.md` *Release* section records the accumulator as intended use rather than a backlog, first-hand.

### Editing the proposal body

- [`src/manifest.ts`](https://github.com/googleapis/release-please/blob/v17.11.2/src/manifest.ts), `createOrUpdatePullRequest` and `updateExistingPullRequest`, show that an existing proposal's body is replaced with the generated one whenever the two differ.
- [`src/strategies/base.ts`](https://github.com/googleapis/release-please/blob/v17.11.2/src/strategies/base.ts), `buildRelease`, parses the merged pull request's body and takes both the set of components to release and each release's notes from it: `const notes = releaseData?.notes`. This is why an edit surviving to the merge becomes the release rather than annotating it.
- [Release Please README](https://github.com/googleapis/release-please/blob/v17.11.2/README.md), *How can I fix release notes?*, documents the `BEGIN_COMMIT_OVERRIDE` block on the merged source pull request, and warns that it does not work with plain merges because release-please cannot tell which commits to apply the override to.

### Prepend-only changelogs

- [`src/updaters/changelog.ts`](https://github.com/googleapis/release-please/blob/v17.11.2/src/updaters/changelog.ts), `updateContent`, searches for the first version heading, and returns `before + entry + after` with `after` copied verbatim. Where no heading is found the new entry follows a fresh header and the previous content is appended below it. No path rewrites an existing section.
- obsidian-compat [`bf96602`](https://github.com/catgrandi/obsidian-compat/commit/bf96602622e9db64a511b3620221cf49a1052cf5) is the observed instance: the generated `0.1.0` section was replaced by hand with notes describing the package, on the stated reasoning that release-please prepends later releases above it and never rewrites a published one. The published `v0.1.0` release body carries the same text, so the two records agree.

### What causes a release

- [Release Please README](https://github.com/googleapis/release-please/blob/v17.11.2/README.md), *Release Please bot does not create a release PR. Why?*: a releasable unit is a commit with one of the prefixes `feat`, `fix`, or `deps`, a `chore` or `build` commit is not one, and some languages add their own — `docs` for Java and Python. [`manifest-releaser.md`](https://github.com/googleapis/release-please/blob/v17.11.2/docs/manifest-releaser.md) documents `changelog-sections` as the mapping that decides this.
- [`src/strategies/base.ts`](https://github.com/googleapis/release-please/blob/v17.11.2/src/strategies/base.ts) is the mechanism: the release is skipped when `changelogEmpty(releaseNotesBody)` holds, and the log line is `No user facing commits found since ${latestRelease ? latestRelease.sha : 'beginning of time'} - skipping`.
- [`manifest-releaser.md`](https://github.com/googleapis/release-please/blob/v17.11.2/docs/manifest-releaser.md) documents `release-as` as setting the next version while ignoring conventional commits, and warns that it must be removed or raised once the release pull request merges or subsequent runs will keep proposing the same version. It documents `bootstrap-sha` as an exclusive bound — choose one commit earlier than the first commit you want to include — that is ignored once release-please has generated at least one release pull request.
- obsidian-compat [`1ab8c56`](https://github.com/catgrandi/obsidian-compat/commit/1ab8c5609dd90da2b81bd60d52d00826169eeaa3) is the observed run showing that `release-as` does not cause a release: with `bootstrap-sha` pointing at the commit publishing was introduced at, leaving only `build` and `ci` commits after it, the log read `Setting version for . from release-as configuration` followed by `No user facing commits found since beginning of time - skipping`, and no proposal was opened. Moving `bootstrap-sha` one commit earlier, so a genuine fix fell in range, produced the release.

### Pre-1.0 bumps

- [`schemas/config.json`](https://github.com/googleapis/release-please/blob/v17.11.2/schemas/config.json): `bump-minor-pre-major` is *Breaking changes only bump semver minor if version < 1.0.0*; `bump-patch-for-minor-pre-major` is *Feature changes only bump semver patch if version < 1.0.0*. [Customizing releases](https://github.com/googleapis/release-please/blob/v17.11.2/docs/customizing.md) documents the `default` versioning strategy that both modify.
- obsidian-compat's `CONTRIBUTING.md` records removing the flag as the deliberate act rather than an accident of a `feat!` commit, first-hand, and its `release-please-config.json` carries `bump-minor-pre-major: true`.

### Levers when nothing qualifies

- `release-as` and `bootstrap-sha` are cited under *What causes a release* above; the levers section adds no behavior beyond them.
- obsidian-compat [`bf96602`](https://github.com/catgrandi/obsidian-compat/commit/bf96602622e9db64a511b3620221cf49a1052cf5) is the paid-for precedent. Its `0.1.0` notes carried an entry for [`1ab8c56`](https://github.com/catgrandi/obsidian-compat/commit/1ab8c5609dd90da2b81bd60d52d00826169eeaa3), a `bootstrap-sha` and `include-component-in-tag` change typed `fix(ci)` to make release-please see a user-facing commit. The correcting commit records that it did not need to be: the same commit moved `bootstrap-sha` to include a real fix, which was already enough to cause the release, so `ci` would have worked and kept the notes clean. Its `CONTRIBUTING.md` carries the rule that came out of it — when a release is wanted and nothing user-facing has landed, change the cadence or the bootstrap, not the commit type.

### Merge style

- [Release Please README](https://github.com/googleapis/release-please/blob/v17.11.2/README.md) states that both squash-merge and merge commits work with Release PRs, and recommends squash-merge — one of its stated reasons being control of the changelog, giving the example of a `fix` for a bug introduced earlier in the same pull request that is irrelevant to the release notes. The same file's *What if my PR contains multiple fixes or features?* documents additional Conventional Commit messages at the bottom of a commit body producing separate entries, and marks the position as important.
- The multi-component attribution consequence follows from path-based attribution, cited under components below. agent-skills' `CONTRIBUTING.md` records the resulting rule first-hand: squash a pull request making one logical change, rebase one carrying more than one releasable commit, because squashing would collapse them into one entry and attribute every file to one component.

### Bot-opened proposals

- Observed in obsidian-compat: CI runs on the `release-please--branches--main--components--obsidian-compat` branch completed with conclusion `action_required` rather than starting — runs [32141323493](https://github.com/catgrandi/obsidian-compat/actions/runs/32141323493) and [32066787259](https://github.com/catgrandi/obsidian-compat/actions/runs/32066787259). Its `CONTRIBUTING.md` records the same behavior and the approve-and-run control from the pull request side, and agent-skills' release procedure records it independently.

### Components, tags, and path boundaries

- [`manifest-releaser.md`](https://github.com/googleapis/release-please/blob/v17.11.2/docs/manifest-releaser.md), *Subsequent Versions*, gives the default search tag `<component-name>-v<release-version>` and the `⚠ Expected 1 releases, only found 0` symptom of leaving `include-component-in-tag` true where tags carry no component. *Releasing Root Path of Library (".")* states that `.` indicates a release should be created when any changes are made to the codebase. *Manifest* states that manual editing is appropriate only in the bootstrap case, and that release-please considers only version information it is missing for a configured package, which is what handles a new package. [`schemas/config.json`](https://github.com/googleapis/release-please/blob/v17.11.2/schemas/config.json) documents `initial-version`, and [`src/strategies/base.ts`](https://github.com/googleapis/release-please/blob/v17.11.2/src/strategies/base.ts) supplies the fallback default when none is configured.
- [Customizing releases](https://github.com/googleapis/release-please/blob/v17.11.2/docs/customizing.md), *Subdirectories (paths) in a repository*: a configured path makes release-please consider only commits that touch files on that path.
- The release-commit invariant is [`src/manifest.ts`](https://github.com/googleapis/release-please/blob/v17.11.2/src/manifest.ts), `commitsAfterSha`: it searches the package's own path-filtered commit list for the last release sha and returns the entire list when `findIndex` yields `-1`, so a package whose release commit touched nothing under its path is given no cutoff.
- [`schemas/config.json`](https://github.com/googleapis/release-please/blob/v17.11.2/schemas/config.json) states the `exclude-paths` contract: *Path of commits to be excluded from parsing. If all files from commit belong to one of the paths it will be skipped.* [`src/util/commit-exclude.ts`](https://github.com/googleapis/release-please/blob/v17.11.2/src/util/commit-exclude.ts) is the prefix test: `isRelevant` returns whether the changed file's index of the excluded path followed by a separator is zero, with no glob expansion, and `shouldInclude` applies the all-files condition.
- agent-skills [`3dd1a11`](https://github.com/catgrandi/agent-skills/commit/3dd1a11053d07595a541ab6091ccd15e71a771a3), [pull request 66](https://github.com/catgrandi/agent-skills/pull/66), is the paid-for precedent: every entry in that repository's exclude list was inert — trailing `/**` matched nothing and bare filenames could not match — and a dry run counted three commits for the aggregate that the configuration intended to exclude.

### Failure modes

- The re-listing proposal is agent-skills [pull request 62](https://github.com/catgrandi/agent-skills/pull/62), which proposed `release-management` 0.2.0 whose sole entry was [`d6025e8`](https://github.com/catgrandi/agent-skills/commit/d6025e805b98768688ea719ff45ae0af76e51b1d), a commit already released in 0.1.0 and cited in that changelog section. It was closed rather than merged. Its cause is the invariant above: the `0.1.0` release commit changed no file inside the skill's directory. The clearing release is [pull request 68](https://github.com/catgrandi/agent-skills/pull/68), merged as [`54eba99`](https://github.com/catgrandi/agent-skills/commit/54eba99ce961e75769091ef06a53e52e2d9843a2), whose `release-management` 0.2.0 section listed `d6025e8` a second time beside the commit that caused the release; [`ea17b95`](https://github.com/catgrandi/agent-skills/commit/ea17b950610e15fc46c721d6cafc1a78170ae340), [pull request 69](https://github.com/catgrandi/agent-skills/pull/69), is the correction landed after publication, with both GitHub Release bodies edited to match.
- The stranded release is documented in obsidian-compat's [`.github/workflows/release.yml`](https://github.com/catgrandi/obsidian-compat/blob/main/.github/workflows/release.yml), whose comments record both directions — cancelling a run that has tagged and created the release but not yet published strands a released version that never reaches the registry, and a publish-job timeout fires after the tag and release already exist, leaving the version recoverable only by re-running the job by hand.
- [`manifest-releaser.md`](https://github.com/googleapis/release-please/blob/v17.11.2/docs/manifest-releaser.md) documents `force-tag-creation` and the draft-release problem it solves: GitHub creates no Git tag for a draft release until it is published, which causes release-please to fail to find the previous release on subsequent runs and potentially generate changelogs including the entire commit history.
- [Release Please README](https://github.com/googleapis/release-please/blob/v17.11.2/README.md), *Step 2*, states that for GitHub application users release-please will not create a new pull request while an existing one carries the `autorelease: pending` label, and that the fix is to remove the stale label and re-run.
