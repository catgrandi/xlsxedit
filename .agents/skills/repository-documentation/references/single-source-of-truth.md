# Single source of truth

Use this reference when consolidating duplicated content, choosing where a fact should live, or updating documentation after a change to facts stated in several places.

## Choose the authoritative home

Order candidate homes by resistance to drift:

1. **The implementation itself.** Code, schemas, manifests, and configuration cannot disagree with themselves. Generate documentation from them when the toolchain supports it.
2. **Output generated from the implementation.** API schemas and rendered references produced by a generator inherit accuracy from their source. Point readers at them, or embed them through generation rather than retyping.

   Check whether the output is generated before relying on it. A usage or help string that a person edits by hand is prose that happens to live in a source file, and it drifts exactly like a README: a flag whose help text describes behavior the code never implements will mislead every document that repeats it. Confirm the claim against the code path that implements it, and report the mismatch instead of propagating it.
3. **One reference document.** When prose must state the fact, give it one location chosen by the audience that consults it most, and link from every other location.

A fact's home is where its precise value or contract lives. Other pages may name the fact and link to it; they should not restate the value.

A link points at the home only when the target states the fact. A page that names the fact and links onward is not the home, and three things disguise that: the link resolves, the heading matches the topic, and the value lives one page further on.

Before offering a link as a fact's home, open the target and read the passage the link lands on. When that passage links onward, follow the chain and link to the page at its end. When no page in the chain states the fact, the fact has no home yet. Give it one, or record the gap, rather than linking to the nearest page that mentions it. Link checking proves that a target exists, never that it is the right target.

## Distinguish framing from copying

Different audiences legitimately need different presentations of the same fact:

- A product user reads: raise `--concurrency` to check more links at once.
- A contributor reads: the pool caps in-flight requests at the configured concurrency.

Both sentences depend on one fact without duplicating its precise contract. Framing restates purpose in audience terms; copying restates the value or contract itself. The default value, the accepted range, and the failure behavior belong in exactly one home.

Summaries follow the same rule. A README options table may state what an option is for in one clause and link to the reference; it should not carry the full contract a second time.

See [audience and Diátaxis architecture](audience-and-diataxis.md) for exposing one canonical page through several navigation paths.

## Handle unavoidable copies

Some formats cannot link, and some facts must appear verbatim in more than one artifact.

- Identify the authoritative source in or beside each copy, such as a comment naming the file that owns the value.
- Update every copy in the same change as the source. A copy scheduled for later is a copy that will be missed.
- Prefer a test or generator that compares or produces the copies when the repository can maintain one.

## Consolidate existing duplication

When documentation already states a fact in several places:

1. Choose the home using the order in this reference.
2. Move the full statement there. When the copies disagree, verify the fact against the implementation before merging them.
3. Replace each former copy with a link, or with an audience-framed summary plus a link.
4. Remove resolved documentation debt, such as notes recording that the fact was undocumented or stale.

Do not leave the old statements in place beside the new home. Two statements that agree today are a future disagreement.
