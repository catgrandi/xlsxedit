---
description: Create, rewrite, review, and restructure repository documentation for product users, operators, integrators, contributors, and maintainers. Use for README files, contribution guides, architecture and decision records, API or CLI references, setup guides, tutorials, changelogs, code documentation, and documentation information architecture involving audience separation, navigation, shared content, or Diátaxis classification. Also use when a code change alters commands, flags, configuration, APIs, defaults, or output and documentation must be brought back into sync, when duplicated or drifting documentation needs consolidation into one authoritative source, and when existing documentation must be audited for accuracy, staleness, or gaps and reported on rather than changed. Ground documentation in the codebase and apply clear, concise, accessible, globally understandable technical-writing conventions.
metadata:
    github-path: skills/repository-documentation
    github-pinned: agent-skills-v1.5.0
    github-ref: refs/tags/agent-skills-v1.5.0
    github-repo: https://github.com/catgrandi/agent-skills
    github-tree-sha: a61780be1e7ecd44b7810f75dffbbd5e7af1a175
    version: 1.7.0
name: repository-documentation
---
# Repository documentation

Create documentation that is technically accurate, easy to scan, and useful to the people who must install, use, operate, maintain, or contribute to the repository.

## Priority order

Apply requirements in this order:

1. Follow the user's explicit requirements.
2. Preserve the repository's implemented behavior and literal identifiers.
3. Follow established repository conventions when they are consistent, usable, and within the user's requested change scope.
4. Apply this skill's documentation rules.

Never improve style by changing technical meaning. Preserve the exact spelling and casing of code identifiers, commands, flags, configuration keys, filenames, UI labels, API names, and trademarks.

## Single source of truth

Give every fact one authoritative home. A default, a flag, a version, a supported platform, or a procedure stated in several places will drift.

- State the precise fact once, in the location its primary audience consults, and link from everywhere else.
- Prefer homes that resist drift: code, schemas, manifests, and output generated from them are more authoritative than prose that describes them. A usage or help string maintained by hand is prose, however close to the code it sits; confirm it against the code path that implements the behavior before treating it as the home.
- Frame the same fact differently for different audiences when their needs differ, but keep one authoritative statement of the precise value or contract. A summary with a link is framing; a second full statement is a copy.
- When a copy is unavoidable, identify the authoritative source in or beside the copy, and update every copy in the same change as the source.

Read [single source of truth](references/single-source-of-truth.md) when consolidating duplicated content, choosing where a fact should live, or documenting a change that touches facts stated in several places.

## Workflow

### 1. Establish the source of truth

Before drafting substantial documentation:

1. Inspect the relevant repository files rather than relying on the request alone.
2. Read existing documentation that overlaps the requested scope.
3. Check implementation sources such as manifests, configuration, public interfaces, tests, examples, command help, and automation files.
4. Identify the reader's relationship to the product, task, and expected level of prior knowledge.
5. Separate current behavior from planned, experimental, deprecated, or proposed behavior.
6. Determine whether the task calls for preserving, incrementally evolving, or replacing the current documentation structure.
7. Identify the Markdown dialect, authoring environment, and rendered destinations before using renderer-specific syntax. Treat the editor and the published renderer as separate when they differ.

Do not invent commands, options, requirements, architecture, compatibility claims, examples, or project policies. When information cannot be verified, state the uncertainty or leave a clearly marked placeholder.

Ground every reference in content the repository tracks. Ignored directories, generated output, dependency trees, and local working files reach no other reader, so naming such a path is useful but describing, quoting, or linking to what it contains is not. When the reference would have been worth making, the fact behind it belongs in a tracked document: move it to the home its audience consults and reference that instead. When the content is unavailable and the fact cannot be recovered, remove the reference and record the gap where the repository tracks such gaps, rather than leaving a pointer no reader can follow. See [references the reader can follow](references/style-rules.md#references-the-reader-can-follow).

Ask a concise clarification before introducing renderer-specific syntax or converting links when repository evidence does not resolve the target renderer or link convention and the choice would materially affect compatibility. Do not ask when the established convention is clear or a localized edit can preserve it.

### 2. Classify the audience and purpose

Treat audience and documentation purpose as independent decisions:

1. Identify one primary audience and any secondary audiences. Use roles that describe the reader's relationship to the product, such as product user, operator, integrator or API consumer, contributor, or maintainer. Do not assume that **developer** means one audience.
2. Identify the reader's immediate need: learn through guided practice, complete a task, look up facts, or understand concepts and decisions.
3. Assign one dominant Diátaxis form to a substantive page: tutorial, how-to guide, reference, or explanation.
4. Split or cross-link material when different audiences need incompatible prerequisites, framing, workflows, or levels of detail.

Ask a concise clarification before creating or moving many pages when either decision remains materially uncertain after repository inspection:

- Whether audiences need separate navigation paths or can share the same framing.
- Whether to preserve, evolve, or replace the existing documentation structure.

Do not ask for routine, localized edits that do not affect information architecture. Do not assume preservation when the user rejects the current structure or explicitly requests a new framework.

Read [audience and Diátaxis architecture](references/audience-and-diataxis.md) when creating or changing documentation navigation, separating user and developer content, handling multi-audience material, or substantially restructuring a documentation set.

### 3. Choose the document pattern

Read [document types](references/document-types.md) when creating a new document, substantially restructuring one, or deciding which sections belong in it.

Use only sections that help the intended reader. Do not add empty or speculative boilerplate to satisfy a generic template.

### 4. Draft for the reader's task

- Put the purpose, outcome, or required action first.
- Use active voice and present tense for current behavior.
- Address the reader as **you** when direct address helps.
- Use imperative verbs for procedures.
- Put conditions and context before the action they govern.
- Use one consistent term for each concept.
- Prefer concrete language over jargon, metaphors, marketing language, or vague abstractions.
- Explain why only when it changes a decision, prevents an error, or clarifies a non-obvious constraint.

Read [style rules](references/style-rules.md) before a substantial rewrite or when wording, formatting, accessibility, or structure is uncertain.

### 5. Format technical content consistently

Read [Markdown variants](references/markdown-variants.md) before adding or changing alerts, callouts, internal links, embeds, or other renderer-specific Markdown.

Read [diagrams](references/diagrams.md) before adding or changing Mermaid diagrams, maps, 3D models, or static diagram assets.

- Use sentence case for titles and headings.
- Use one level-1 heading per Markdown file unless the repository format requires otherwise.
- Use task headings that begin with a base-form verb, such as **Install dependencies**.
- Use noun phrases for conceptual headings, such as **Configuration precedence**.
- Use numbered lists only when sequence or priority matters.
- Use bullets for unordered information and for a formatted single-step procedure.
- Use description lists for genuine term–description relationships when the target renderer supports them.
- Format placeholders, keyboard input, system output, UI labels, and actionable controls according to [Semantic HTML for technical text](references/style-rules.md#semantic-html-for-technical-text).
- Use inline code for literal technical content that does not require more specific semantic markup.
- Use fenced code blocks with a language identifier by default.
- When semantic elements must remain active inside a preformatted example, use the verified language-tagged `<pre><code>` structure documented in the style rules.
- Use alerts or callouts sparingly for exceptional information that materially affects success, safety, correctness, or compatibility. Keep normal instructions and explanations in the document flow.
- Use a diagram only when it makes an important relationship, flow, state change, hierarchy, topology, or spatial structure materially easier to understand than concise prose or a table. Keep an equivalent explanation in text.
- Use descriptive link text that remains meaningful out of context.
- Choose repository-local link syntax for the target Markdown dialect. When using standard Markdown, prefer relative links.

### 6. Make examples trustworthy

- Prefer the smallest realistic example that demonstrates the task.
- Make commands and code copyable when practical.
- Include prerequisites, dependencies, working directory, and required permissions when they are not obvious.
- Show expected output when it helps readers confirm success.
- Use safe values and secure patterns. Never place real secrets, credentials, personal data, or unsafe defaults in examples.
- Test examples when tools and dependencies are available.
- Do not claim that an example was tested unless it was actually run.
- Mark intentionally incomplete code with a language-appropriate comment, not an unexplained ellipsis.

### 7. Validate before finalizing

Read and apply the [review checklist](references/review-checklist.md) before finalizing a new document or a substantial documentation change, and when the task is to review documentation without changing it. Reviewing satisfies an item by reporting the problem where changing would correct it.

At minimum, verify that:

- Commands, paths, names, defaults, and links match the repository.
- The page has a clear primary audience and dominant reader need, or it is explicitly a landing or navigation page.
- Prerequisites appear before the steps that depend on them.
- Procedures have a clear starting point and completion state.
- Current, proposed, deprecated, and removed behavior are clearly distinguished.
- Examples do not contradict the implementation or each other.
- The document is understandable without relying only on images, color, screen position, or surrounding visual context.

### 8. Report unresolved issues

When returning the completed work, identify any material assumptions, unverified commands, broken references, missing repository information, or decisions that require a maintainer. Do not hide uncertainty behind polished prose.

## Core style rules

- Write like a knowledgeable teammate: direct, respectful, concise, and helpful.
- Front-load important words and decisions.
- Prefer short sentences and paragraphs, but do not make the prose abrupt.
- Use standard American English and the serial comma unless the repository explicitly uses another locale.
- Use natural contractions when they improve readability.
- Avoid **simply**, **just**, **easy**, **obvious**, and similar language that dismisses difficulty.
- Avoid institutional **we** when the reader or an explicit actor is clearer.
- Avoid vague pronouns with unclear antecedents.
- Avoid idioms, slang, culture-specific humor, and layout-dependent references such as **above** or **below**.
- Avoid exclusionary, gendered, violent, or ableist figurative language.
- Default to input-neutral UI verbs such as **select**, **open**, **go to**, **choose**, and **enter**. Use **click** or **tap** only when the input method matters.
- Do not use vague link text such as **click here**, **here**, or **this page**.
- Do not describe planned features as available or promise future behavior without an authoritative source.

## Edit existing documentation

- Preserve correct content and the author's intent; change only what improves accuracy, usability, consistency, or scope.
- Preserve, evolve, or replace the existing information architecture according to the user's intent; do not mistake an established structure for a required or effective one.
- Prefer focused edits over unnecessary full-file rewrites.
- Retain repository-specific terminology unless it is incorrect, unclear, or inconsistent.
- Update related links, examples, navigation, and references when a changed section affects them.
- Remove obsolete content rather than leaving contradictory alternatives.
- Do not normalize generated files manually. Update the source or generator instead.

## Documentation impact of code changes

When a change alters observable behavior — commands, flags, configuration keys, defaults, APIs, output formats, error messages, or workflows — the documentation that states that behavior is part of the change.

1. Find every document that states the old behavior. Identifiers, defaults, examples, diagrams, and recorded documentation debt go stale independently.
2. Update the affected documents in the same unit of work when you own the complete change.
3. When delivering one part of a larger coordinated change, do not edit shared documentation mid-flight. Record each affected document and the fact it must state once the change is complete, and give that list to whoever owns the complete change.
4. Resolve recorded documentation debt that the change satisfies, such as a tracked note that a feature is undocumented. Leaving the note contradicts the new documentation.

This skill loads only when documentation is the task. For this section to reach agents while they change code, the repository must wire [the documentation-sync rule](assets/agent-instructions.md) into instructions that load on every task — by default one line in `AGENTS.md` linking to the rule's installed location. When that wiring is missing, propose adding it rather than adding it silently; changing a repository's agent instructions is a maintainer decision.

## Gotchas

- A polished README is still wrong if its commands do not run.
- A help or usage string can be as stale as a README. Check the flag's implementation before repeating what its help text claims.
- A reference to an ignored working file describes a repository that only its author can see.
- Existing documentation is evidence, not necessarily the source of truth; confirm it against the repository.
- Existing navigation is evidence, not an obligation; do not preserve a framework that the user wants replaced.
- A document does not need every conventional section. Omit sections that have no useful content.
- Do not convert literal code or UI text to sentence case.
- Do not impose audience folders, four Diátaxis folders, a complete audience-by-type matrix, or a `shared/` folder without evidence that the navigation benefits readers.
- Do not use screenshots as the only explanation of a process.
- Do not turn every explanatory sequence into a numbered procedure; use numbers only when order matters.

## Provenance

See [sources and reconciliation](references/sources.md) for the primary guides and the decisions used to combine them.
