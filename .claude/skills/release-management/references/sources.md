# Sources

The rules a primary source settles are cited below, grouped by the claim they support. Each source was read at first hand, at the URL given, on 2026-08-23. Where a source does not state what it is commonly said to state, this file says so rather than filling the gap.

The general rules come from specifications that prescribe no release tool, branching model, or language. Where a rule rests on registry policy instead, independent ecosystems are cited so that no single one carries it, and an ecosystem that departs from the common rule is named rather than generalized past.

Four things the skill states are not settled by any source here, and are not claimed to be: the pre-publish readiness checks, the consumer-side verification sequence, the fit of calendar versioning, and the support-line guidance. Those are practice, recorded as such so a reader can tell the difference.

## Immutability, and never reusing a number

- [Semantic Versioning 2.0.0](https://semver.org/spec/v2.0.0.html), item 3, is the general rule: once a versioned package has been released, the contents of that version must not be modified, and any modification must be released as a new version.
- [npm unpublish policy](https://docs.npmjs.com/policies/unpublish) states that a version number, once used, can never be used again, and that a new version must be published even if the old one was unpublished. It also carries the removal window and its conditions: within 72 hours of publishing if nothing in the registry depends on the package, and after that only if the package has no dependents, had fewer than 300 downloads in the last week, and has a single owner. The [`npm unpublish`](https://docs.npmjs.com/cli/v12/commands/npm-unpublish) command reference states the same reuse rule, and adds that unpublishing an entire package blocks new versions of that name for 24 hours.
- [PyPI help](https://pypi.org/help/#file-name-reuse) states that a filename cannot be reused even after a project has been deleted and recreated, and that a distribution filename is composed of the project name, version, and distribution type — which is what makes a deleted version unrepublishable. The same page states that deletion of a project, release, or file is permanent and irreversible, and that deleted files cannot be re-uploaded.
- [RubyGems yank policy](https://blog.rubygems.org/2015/04/13/permadelete-on-yank.html) states that a gem version cannot be published twice, so a yanked version still requires a version bump rather than a quick re-push. Neither [removing a published gem](https://guides.rubygems.org/removing-a-published-gem/) nor the [command reference](https://guides.rubygems.org/command-reference/#gem-yank) states this rule; the blog post is the source for it.
- [Publishing on crates.io](https://doc.rust-lang.org/cargo/reference/publishing.html) states that a published version can never be overwritten and the code cannot be deleted, and gives the rationale that crates.io is meant to act as a permanent archive.
- [Maven Central immutability](https://central.sonatype.org/publish/requirements/immutability/) states that components are not removed or modified once publicly available, and the FAQ [can I change a component](https://central.sonatype.org/faq/can-i-change-a-component/) answers no, because build tools cache release artifacts rather than re-checking them.
- [Go modules reference](https://go.dev/ref/mod) defines a version as an immutable snapshot of a module, and states that the checksum database ensures the bits associated with a specific version do not change from one day to the next even if the author later alters the tags in their repository. Its version-control guidance is the source for not moving a published ref.
- [Hex `mix hex.publish`](https://hex.hexdocs.pm/Mix.Tasks.Hex.Publish.html) is the documented counterexample, and the reason the skill states the reuse rule without claiming that every registry enforces it. Its `--replace` option overwrites an existing version, with public packages overwritable within one hour of first publication, and `--revert` removes a version, within one hour for a new version of an existing package and 24 hours for a new package. A registry can hand the number back; the consumers, caches, and records that already resolved it cannot be handed back with it, which is why the rule rests on ambiguity rather than on enforcement.

## Marking a version rather than deleting it

The mechanisms differ in what they cost the consumer, which is why this skill tells an agent to read the specific mechanism before using it.

- [`npm deprecate`](https://docs.npmjs.com/cli/v12/commands/npm-deprecate) updates the registry entry so that a deprecation warning is shown on install. It removes nothing, and resolution is unaffected. npm's own [unpublish](https://docs.npmjs.com/cli/v12/commands/npm-unpublish) page directs maintainers to deprecate instead when the intent is to stop people using a version. Deprecation accepts a single version or a version range.
- [PEP 592](https://peps.python.org/pep-0592/) defines yanking as letting an author effectively delete a file without breaking people who pinned to an exact version. Its one normative installer rule is that an installer must ignore yanked releases if the constraints can be satisfied without one. [PyPI's yanking guide](https://docs.pypi.org/project-management/yanking/) calls it a non-destructive alternative to deletion and names the cases for it: a broken or uninstallable release, a release that violates its own compatibility guarantees, and a release containing a security vulnerability. PEP 592 contains no statement of what yanking must *not* be used for; a claim of that shape cannot be sourced to it.
- [`cargo yank`](https://doc.rust-lang.org/cargo/commands/cargo-yank.html) removes a version from the index but explicitly deletes no data: the version stays downloadable, existing lock files keep working, and only new resolution is affected. The same page states that yanking does not stop a leaked credential from spreading, because existing lock files and direct downloads are unaffected.
- [Go modules reference](https://go.dev/ref/mod), retract directive: a retracted version should remain available so that builds depending on it keep working, and is excluded from upgrades and version queries rather than removed. The retraction is published as a *new* version carrying the directive, which is a concrete instance of correcting a published record by addition.
- [Removing a published gem](https://guides.rubygems.org/removing-a-published-gem/) is the counter-example that the skill's caution exists for: `gem yank` removes the gem file itself, so a yank on that registry is not a non-destructive marking. It also states that a yank does not undo distribution, and that accidentally published credentials must be reset immediately.
- [Keep a Changelog 1.1.0](https://keepachangelog.com/en/1.1.0/) supplies the notes-side marker: a version pulled for a serious bug or security issue should still appear, with `[YANKED]` appended to its heading, as in `## [0.0.5] - 2014-12-13 [YANKED]`.

## Correcting a published record by addition

- Michael Nygard, [Documenting architecture decisions](https://cognitect.com/blog/2011/11/15/documenting-architecture-decisions), defines the statuses proposed, accepted, deprecated, and superseded, and states the append-only rule: if a decision is reversed, the old record is kept and marked as superseded, with a reference to its replacement, because it stays relevant to know what the decision was. The article does not use the words *immutable* or *never delete*; the rule is stated affirmatively rather than as a prohibition.
- [The RFC series](https://www.ietf.org/process/rfcs/) states that with one exception, once an RFC is published it is never changed, which is why the series is described as archival in nature. The exception is format-only: a publication format that renders incorrectly may be replaced, but the content is not revised.
- [RFC errata](https://www.rfc-editor.org/series/rfc-errata/) is the correction mechanism that follows from that: verified errata are linked to the RFC and are not incorporated into its published formats.
- [RFC 7841](https://www.rfc-editor.org/rfc/rfc7841) section 3.1 defines the `Updates` and `Obsoletes` relations that a *new* document declares in its header, and [RFC 7322](https://www.rfc-editor.org/rfc/rfc7322) section 4.1.4 gives their syntax. Replacement is expressed by the replacement, not by editing what it replaces.
- Keep a Changelog's FAQ answers *should you ever rewrite a changelog?* with yes, giving adding missing releases as the example. This is why the skill separates improving published notes, which is encouraged, from silently replacing a statement a reader may have acted on, which is not.

## Versioning semantics

All from [Semantic Versioning 2.0.0](https://semver.org/spec/v2.0.0.html), by item number:

- Item 1: software using Semantic Versioning must declare a public API, in code or in documentation, and it should be precise and comprehensive. This is why the skill treats a declared public interface as the prerequisite for judging any bump.
- Items 6, 7, and 8: PATCH for backward compatible bug fixes, MINOR for new backward compatible functionality, MAJOR for backward incompatible changes to the public API, with lower numbers reset on each increment. Item 7 also requires a MINOR increment when public API functionality is marked as deprecated.
- Item 4: major version zero is for initial development, anything may change at any time, and the public API should not be considered stable.
- Item 5: version 1.0.0 defines the public API. The numbered specification does not say when to release it; the FAQ does, answering that software already used in production, or with a stable API users depend on, should probably already be 1.0.0. Treating the move to 1.0.0 as a decision rather than a side effect follows from the two together.
- Item 9: a pre-release version may be denoted by appending a hyphen and dot-separated identifiers, indicates that the version is unstable and might not satisfy the compatibility its associated normal version implies, and has lower precedence than that normal version.
- Items 10 and 11.1: build metadata is ignored when determining version precedence, so it cannot distinguish two releases.
- FAQ, on accidentally releasing a backward incompatible change under a compatible number: fix the problem and release a new version that restores compatibility. The advice is to fix forward, not to relabel what shipped.
- FAQ, on deprecating functionality: update the documentation and issue a new minor release with the deprecation in place, and ship at least one minor release containing the deprecation before removing the functionality in a major release.

## Notes written for the consumer

All from [Keep a Changelog 1.1.0](https://keepachangelog.com/en/1.1.0/):

- Guiding principles: changelogs are for humans, not machines; every version gets an entry; the same types of change are grouped; versions and sections are linkable; the latest version comes first; each version shows its release date; state whether the project follows Semantic Versioning.
- Types of change: Added, Changed, Deprecated, Removed, Fixed, and Security — a consumer-visible taxonomy, with no category for internal work.
- *Commit log diffs*: using commit log diffs as changelogs is a bad idea because they are full of noise, naming merge commits, obscure titles, and documentation changes. The distinction drawn is that a commit documents a step in the evolution of the source, while a changelog entry describes the noteworthy difference — often across many commits — for end users.
- *Ignoring deprecations*: when people upgrade from one version to another it should be painfully clear when something will break, and it should be possible to upgrade to a version that lists deprecations, remove what is deprecated, then upgrade to the version where the deprecations become removals.
- *Inconsistent changes*: a changelog that mentions only some changes can be as dangerous as no changelog, because readers trust it as a complete record.
- Dates are given in ISO 8601 form, `YYYY-MM-DD`.

## Scope of these sources

These sources fix what a release means and what a published version obliges. None of them prescribes a release tool, a branching model, or a pipeline, and neither does this skill. Guidance that depends on a particular tool or strategy belongs in a reference beside this file; guidance that depends on one repository belongs in that repository's contribution guide.
