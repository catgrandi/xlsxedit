---
description: Plan, cut, verify, and repair software releases for any release tooling, branching model, language, or distribution channel. Use when preparing or cutting a release, deciding which version number a change deserves, tagging a release ref, publishing or re-publishing an artifact, confirming that a publish reached consumers, writing or correcting the notes for an already-released version, deciding whether to yank, deprecate, unpublish, or delete a published version, handling a release that shipped broken or from the wrong ref, planning a hotfix, deprecating and later removing a feature, choosing pre-release identifiers, declaring the first stable version, deciding which released versions still receive fixes, or working with Release Please and its release pull request, generated changelog, or `release-please-config.json`. Not for drafting commit messages, which `conventional-commits` covers, or for changelog and documentation prose and structure, which `repository-documentation` covers.
metadata:
    github-path: skills/release-management
    github-pinned: agent-skills-v1.5.0
    github-ref: refs/tags/agent-skills-v1.5.0
    github-repo: https://github.com/catgrandi/agent-skills
    github-tree-sha: a6f975d28b022d890d4f9779c4cf194d425b62b6
    version: 0.2.1
name: release-management
---
# Release management

Publish releases consumers can rely on, and recover correctly when a published release turns out to be wrong.

## Scope and precedence

Use this skill for the decisions a release requires: what the next version should be, whether a release is warranted at all, what the notes say, whether the release actually reached consumers, and what to do about a version that should not have shipped.

Apply requirements in this order:

1. Follow the user's explicit requirements.
2. Follow the repository's own release process and any compatibility or support policy it publishes to consumers.
3. Apply the rules in this file.

A release is outward-facing and largely irreversible. Do not tag, publish, re-publish, edit a published release, deprecate, yank, unpublish, or delete anything unless the user has authorized that specific action. Otherwise describe the action, its consequence, and the alternative, and let the user decide.

Two adjacent concerns sit outside this skill. How a commit message is typed, including which type marks a breaking change, belongs to the `conventional-commits` skill. How changelog and release-note prose is written and structured belongs to the `repository-documentation` skill. Defer to those where the project uses them.

## A version is a promise

A version number names one exact set of contents. That promise is what every other rule here protects.

- **One version, one artifact.** Once a version is published, its contents never change. A correction is a new version. Semantic Versioning states this as a requirement, and the major registries enforce some form of it.
- **Never reuse a number,** including one that was withdrawn, unpublished, or published by mistake. Something may already hold the old contents under that number: a lockfile, a cache, a mirror, an offline copy, a build log. A number with two meanings makes every record of it ambiguous. Removal does not return the number: most registries that permit removal at all refuse to let that number be published again, and where one does allow it inside a short window, everything that already resolved the old contents still holds them.
- **Rebuilding under the same number is a rewrite,** even when the source ref is unchanged. Consumers cannot tell two builds apart, and cannot tell which one they hold.
- **The number is spent the moment anything could have reached a consumer,** not the moment you notice the mistake. Treat publication, not discovery, as the point of no return.
- **Undo forward.** A change that must be reversed is reversed by a new version that reverses it, never by altering the version that shipped. A version whose number said one thing and whose contents said another is corrected the same way: release the correction, and let the numbers tell consumers what happened.
- **A published record is corrected by addition.** A statement that turns out wrong gets a correction beside it that says what changed, rather than a silent replacement, so a reader who acted on the original can see that it moved. This is the discipline that keeps decision records and published standards trustworthy: the superseded record stays and names what replaced it, and corrections are attached rather than merged in.

## Choose a versioning scheme, then apply it

Choosing a scheme is a decision. Applying the chosen scheme consistently is a rule. Whatever the scheme, consumers must be able to order two versions and infer what an upgrade will cost them, and its meaning must be written where consumers look.

Semantic Versioning is the default expectation for anything with a versioned interface. Calendar versioning suits products whose value to a consumer is recency rather than compatibility. Another scheme is legitimate when the project states what its numbers mean.

Under Semantic Versioning:

- The bump is decided by the largest-impact change in the release, not by how many changes it contains.
- A **breaking** change — one that is incompatible with the declared public interface — bumps MAJOR and resets the lower numbers.
- A **feature** — behavior a previous consumer can take without changing anything — bumps MINOR.
- A **fix** — a correction that keeps the interface compatible — bumps PATCH.
- Compatibility is judged against a *declared* public interface. Where nobody has said what is public, no bump can be judged, and declaring the interface comes first.
- **`0.y.z` promises nothing.** Under the specification, anything may change at any time while the major version is zero. Some ecosystems give a `0.x` minor bump a stricter meaning in dependency resolution; where consumers rely on that, state the convention rather than assume it.
- **Reaching `1.0.0` is a deliberate claim, not a consequence.** It says the public interface is now defined and will be versioned accordingly. A breaking change during `0.y.z` does not promote a project to `1.0.0`; the maintainer decides that the interface is ready to be promised.
- **A pre-release identifier** marks a version that may not satisfy the compatibility its number would otherwise imply, and it sorts *before* the release it precedes. A pre-release is still published: it is immutable, its number is spent, and consumers who opt in hold it.
- **Build metadata does not distinguish releases.** Version comparison ignores it, so two artifacts that differ only there are one version to anything that resolves versions.

## Make the things that name a release agree

A release is not one object. It is a set of records that must all name the same thing:

- **The source ref** — a tag or commit. What exactly was released?
- **The version recorded in the project's own files.** What does the project call itself?
- **The release-notes entry.** What is in it, for a consumer?
- **The built artifact,** including any version it reports at run time. What do consumers actually run?
- **The distribution record** — a registry entry, release page, download index, or equivalent. What can a consumer resolve?

The release is complete when all of them agree, and most release defects are drift between two of them: a ref that points at contents the artifact was not built from, an artifact that reports a version the project never recorded, notes for a version that was never distributed.

Give the version exactly one authoritative home and derive every other occurrence from it. A version maintained by hand in two places is a mismatch waiting for the release that updates one and forgets the other.

## Check readiness before publishing

Confirm each of these before publishing. Each has produced released defects that could not be withdrawn:

- **Validation ran against the exact ref being released.** A passing result from an earlier revision says nothing about this one. Where the release ref is not the ref that was validated, validate the release ref.
- **The published file set is asserted, not assumed.** Build the artifact and inspect what it contains and what it omits. A missing file breaks consumers, and an accidentally included secret, credential, or private path cannot be recalled once distributed.
- **The artifact reports the version being released,** wherever it reports one.
- **Documentation describes the behavior this release ships,** not the behavior of an earlier revision or of the previous release.
- **No breaking change hides under a non-breaking label.** Ask what happens to a consumer who takes this version without reading anything. Removed or renamed public names, changed defaults, stricter validation, changed output formats, and raised minimum requirements are all breaking, whatever the change that caused them was called.
- **License and attribution files travel inside the artifact,** not only in the repository.
- **A release is actually wanted.** Whether to release is an explicit decision. Never misrepresent the nature of a change in order to cause a release or to prevent one: a type, label, or note chosen to move a version number rather than to describe the change corrupts the one record consumers rely on. When a release is wanted and no qualifying change exists, use the project's explicit mechanism for releasing deliberately, or say plainly that nothing warrants a release.

## Write release notes for the consumer

Release notes exist for someone deciding whether to take this version and what it will cost them.

- **Include what a consumer can observe:** added behavior, changed behavior, removed behavior, fixed behavior, security fixes, and deprecations.
- **Lead with breaking changes, and lead each one with the action.** State what a consumer must change, then why. A breaking change a reader has to reconstruct from a description is an outage waiting to happen.
- **Leave out what only a contributor can see.** Refactors, test changes, build and pipeline work, internal dependency bumps, and repository housekeeping do not belong in the notes. A dump of the project's internal change record is not release notes.
- **A published promise that turns out false is a consumer-facing fix,** even when no code changed. Documented behavior the product never had, a compatibility claim that was wrong, an example that cannot work: correcting these changes what a consumer may rely on, so the correction belongs in the notes.
- **Improve published notes freely; correct them by addition.** Adding a missing entry or clarifying a confusing one makes published notes more useful and is worth doing. Replacing a statement someone may have acted on is different: add the correction beside the original entry and mark it as a correction, so a reader who acted on the first wording can see that it moved.
- **Mark a version that should not be used, and leave its entry in place.** A version pulled for a serious bug or a security problem belongs in the notes precisely because it was pulled. The established changelog convention is a `[YANKED]` marker appended to that version's heading.
- **Where notes are generated, fix the input the generator reads.** Editing generated output either loses the edit at the next run or leaves the generator and the file permanently disagreeing. Where the input can no longer be changed because it is already published, annotate additively at the destination and record why.

## Verify after publishing

Publishing is not finishing. Verify from the consumer's side:

1. Obtain the release the way a consumer does — from the distribution consumers resolve, in a clean environment, without your working tree, credentials, or caches.
2. Confirm the artifact reports the intended version and contains the intended files.
3. Confirm the source ref, the project's recorded version, the notes entry, and the distribution record all name the same version.
4. Confirm the release is reachable by default resolution, not only by an exact pin.

Recover from what this finds by repairing the record rather than spending another number:

- **The release record exists but the artifact never reached distribution.** Re-run distribution for the same ref. Do not create a new ref or bump the version because a step did not run.
- **The artifact reached distribution with no ref or release record naming it.** Create the missing record, pointing it at the exact contents that were distributed. Never point it at different contents to make the record tidy.
- **The release was built from the wrong ref.** The number is spent. Mark the published version so consumers do not land on it, and release the intended contents as the next version. Do not move the ref: a ref consumers may already have resolved must keep meaning what it meant.
- **Part of a multi-part release published and the rest did not.** Complete the set at the same version. Do not bump the parts that succeeded so a retry matches.

## Triage a release that is wrong

Answer one question first: what is wrong with the published version?

### 1. It has an ordinary defect

Wrong behavior, a missing fix, or an inaccurate note. **Fix forward.** Release the correction as the next version and describe it in that version's notes. Nothing about the bad version is withdrawn, marked, or rewritten. Most wrong releases stop here.

A version whose number understated its impact belongs here too. When a breaking change went out under a feature or fix number, release a version that restores the compatibility the number promised, and say what happened. Do not relabel the published version: its number is part of what shipped, and changing it now would mean two different things by one name.

### 2. It should not be installed

The version is broken enough that landing on it is worse than not upgrading, or it was published by mistake. **Mark it and ship the fix.**

- Ship the corrected version first, so consumers who are steered away have somewhere to go. Marking without a fix tells people to leave without telling them where.
- Use the distribution channel's own mechanism for *still published, no longer to be chosen*. Channels name it differently — deprecate, yank, retract — and the mechanisms differ in what they cost. Some only print a warning and change nothing about resolution. Some exclude the version from new resolution while leaving existing pins working. At least one removes the artifact outright, which breaks anything that has to fetch it again. Read what the specific mechanism does before using it, because that difference decides whether existing consumers keep working.
- Annotate that version's notes entry.

Marking is not deletion, and where the channel keeps the artifact, that is the point: anything already depending on the version keeps resolving, while anyone arriving now is steered away.

### 3. The artifact itself is harmful

It carries a leaked credential, malicious code, or something that damages a consumer's system or data when used. **Remove the artifact from distribution and keep the ref and the release record,** with a note stating that the artifact was removed and why.

Removal limits further spread; it does not undo the exposure. A credential that was published is compromised from the moment it was published, so rotate it and treat the disclosure as an incident rather than as a release problem.

### Anything else

Deleting a published version is wrong outside case 3. Some registries also allow removal within a short window after publishing, or under narrow conditions such as having no dependents; using that window is defensible only when nothing can have resolved the version yet. Where a registry does hand the number back, treat it as spent anyway: something may already hold the old contents under it.

Consumers cannot be un-notified. Once a version has been announced, a follow-up naming the affected version and the action to take is part of the fix.

## Deprecate before removing

Removing something consumers use is a breaking change, and removing it without warning breaks the version's promise whatever number it carries.

1. **Deprecate.** The feature still works. Its documentation says it is deprecated, names the replacement, and states the migration. Under Semantic Versioning, marking public functionality as deprecated is itself a feature-level change and bumps MINOR.
2. **Announce.** The deprecation appears in the notes of the release that deprecates it, not only in the release that removes it. A consumer who reads release notes has to be able to see the removal coming.
3. **Remove,** as a breaking change, with the bump the scheme requires and a notes entry that names the replacement again.

At least one release has to carry the deprecation before the release that removes it, so that a consumer can upgrade to a version that warns, migrate, and then upgrade again. Compressing the two into one release is a removal with a note attached, not a deprecation. Do not shorten the sequence because a feature looks unused either: you can measure your own use of it, not your consumers'.

## State the support line

Which released versions still receive fixes is a decision, and one consumers must be able to read before they need it. Left unstated, consumers assume every version they run is supported and maintainers assume only the newest is.

State which versions receive fixes, which kinds of fix they receive — security only, or any — and until when. How a fix reaches an older line depends on the project's branching model and belongs with that model, not here.

## Layer this skill onto the project

This file holds only what is true of every release. Two layers sit on top of it:

- **Tooling and strategy.** A reference under `references/` may describe what a specific release tool or branching strategy adds to, or narrows in, these rules. Determine what the repository actually uses, and read the matching reference when one exists.
- **The repository's own process.** Pipeline steps, credentials, approvals, who may publish, and where releases are announced belong in the repository's contribution guide. Read them there; do not copy them into this skill.

Detect the tooling from what the repository contains:

| The repository contains | Read |
| --- | --- |
| `release-please-config.json` or `.release-please-manifest.json` | [Release Please](references/release-please.md) |

When none matches, apply this file and say that this skill carries no reference for that setup, rather than assuming another tool's behavior.

## Gotchas

- A green validation run proves something about the ref it ran on, and nothing about the ref you are releasing.
- A version that was never distributed still spent its number if the ref or the release record is public.
- Re-tagging looks like a fix and is a rewrite: the same name now resolves to different contents for anyone who fetches later.
- Deleting a bad version does not remove it from the lockfiles, caches, and mirrors that already hold it; it only removes the fixed version's audience.
- Marking a version without shipping a replacement tells consumers to leave without telling them where to go.
- Deprecation warnings that first appear in the release that removes the feature are removal notices, not deprecations.
- Notes assembled from an internal change record describe the project's work, not the consumer's upgrade.

## Provenance

See [sources](references/sources.md) for the primary source behind each rule, and for the rules this skill states that no primary source settles.
