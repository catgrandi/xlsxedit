# Documentation review checklist

Use this checklist in two situations: before finalizing a new document or a substantial change, and when reviewing documentation you are not changing. Skip items that do not apply.

The items are written to be read either way. Where an item asks whether something was corrected, a review that produces no edits satisfies it by reporting the problem instead. Two items apply only to a change, and say so.

## Technical accuracy

- [ ] The document reflects the current repository, not assumptions or stale documentation.
- [ ] The Markdown dialect, authoring environment, and rendered destinations are identified when syntax varies by renderer.
- [ ] Commands, flags, paths, filenames, configuration keys, UI labels, API names, and defaults are exact.
- [ ] Current, experimental, planned, deprecated, and removed behavior are clearly distinguished.
- [ ] Diagrams match the documented system state and label current, proposed, or target architecture explicitly.
- [ ] Compatibility, security, performance, and stability claims have evidence.
- [ ] Behavior copied from a help or usage string was confirmed against the code that implements it.
- [ ] Generated documentation is changed only through its source or generator, never by hand.
- [ ] Examples use safe placeholders and contain no secrets or personal data.

## Synchronization and single source of truth

- [ ] Each fact in scope has one authoritative home; other mentions link to it or summarize it without restating the precise value or contract.
- [ ] Each link offered as a fact's home reaches a page that states the fact, rather than one that names it and links onward.
- [ ] Every document stating behavior the repository has superseded was found, and either corrected or listed for whoever owns the fix.
- [ ] Copies that cannot be links identify their authoritative source and agree with it.
- [ ] Documentation debt the repository has outgrown, such as a note that a feature is undocumented after it was documented, was resolved or reported.
- [ ] Every reference points at tracked content; no link, quotation, or description depends on an ignored directory, generated output, or a local working file.
- [ ] Facts the reader needs that existed only in untracked content were moved into a tracked home rather than dropped with the reference.
- [ ] Where untracked content was unavailable, the unfollowable reference was removed and the missing fact recorded as a gap.
- [ ] Paths are repository-relative rather than absolute paths from one machine.

## Audience and purpose

- [ ] The primary audience, reader task, and expected prior knowledge are clear from the title and opening section.
- [ ] Broad labels such as **developer** are split into useful roles when readers have different relationships to the product.
- [ ] Any secondary audiences can use the same page without incompatible prerequisites, framing, or workflow.
- [ ] The page has one dominant Diátaxis purpose—tutorial, how-to, reference, or explanation—or is intentionally an entry or navigation page.
- [ ] Prerequisite knowledge is appropriate and stated when necessary.
- [ ] The document answers the reader's likely first questions early.
- [ ] Content outside the document's purpose is removed or linked elsewhere.

## Structure and navigation

- [ ] The choice to preserve, evolve, or replace the existing documentation structure matches the user's stated or reasonably inferred intent.
- [ ] An established but ineffective framework was not preserved solely because it already exists.
- [ ] **Change only:** material uncertainty about audience separation or restructuring scope was resolved before broad file moves or duplication.
- [ ] The first navigation layer reflects the distinction readers most need to make, whether audience, documentation form, topic, or a deliberate hybrid.
- [ ] The structure does not contain empty sections created only to complete an audience-by-type matrix.
- [ ] Multi-audience content has one canonical source linked from every relevant reader path, unless different framing requires separate pages.
- [ ] Landing pages explain the audience and contents instead of presenting an unexplained link dump.
- [ ] The file has one clear level-1 heading when the format expects one.
- [ ] Heading levels are hierarchical and not skipped.
- [ ] Headings are unique, descriptive, concise, and sentence-cased.
- [ ] Task headings begin with base-form verbs; concept headings use noun phrases.
- [ ] Sections appear in the order readers need them.
- [ ] Lists and tables are used only when they improve understanding.
- [ ] Repository-local links are relative when practical.
- [ ] Incoming anchors and table-of-contents entries resolve, including any a changed heading would break.

## Procedures and examples

- [ ] The goal and starting context are clear.
- [ ] Prerequisites appear before dependent steps.
- [ ] Numbered steps are used only when order matters.
- [ ] Each step starts with an imperative verb or a necessary context phrase.
- [ ] Each step has one principal action.
- [ ] The procedure includes the action that completes or saves the task.
- [ ] The expected result or verification method is included when useful.
- [ ] Commands identify the working directory, permissions, and dependencies when needed.
- [ ] Examples are minimal, realistic, copyable, and internally consistent.
- [ ] Tested examples are identified accurately; untested examples are not described as verified.

## Language and terminology

- [ ] The prose is direct, concise, respectful, and free of marketing filler.
- [ ] Active voice and present tense are used when appropriate.
- [ ] Conditions and context appear before the instructions they govern.
- [ ] One consistent term is used for each concept.
- [ ] Abbreviations are defined when the audience might not know them.
- [ ] Vague pronouns, jargon, idioms, slang, and culture-specific humor are removed.
- [ ] Words such as **simply**, **just**, **easy**, **obvious**, and **of course** are removed unless strictly literal.
- [ ] Inclusive language is used; literal legacy identifiers are preserved only where required.

## Formatting

- [ ] Titles and headings use sentence case.
- [ ] Literal technical content uses code formatting when no more specific semantic element applies.
- [ ] Placeholders, keyboard input, system output, UI labels, and actionable controls use the appropriate semantic HTML when the target renderer supports it.
- [ ] Ordinary code blocks use fenced syntax with the correct language identifier.
- [ ] Code blocks containing active semantic elements use the verified language-tagged `<pre><code>` structure.
- [ ] The same language identifier appears in the `lang` attribute and both `language-*` classes.
- [ ] Characters that HTML would interpret are escaped inside raw HTML code blocks.
- [ ] List items are parallel and punctuated consistently.
- [ ] Alerts or callouts are reserved for exceptional information, use a type that matches the consequence, and use syntax supported by the target renderer.
- [ ] Each diagram materially improves understanding, uses an appropriate format, and renders in every required destination.
- [ ] Link text is descriptive and meaningful out of context.
- [ ] Internal links and embeds use the repository's intended syntax and resolve in every required renderer.
- [ ] Dates, times, versions, and units are unambiguous.

## Accessibility and global use

- [ ] Essential information is available as text, not only through images, color, sound, or position.
- [ ] Each informative diagram is introduced and has an adjacent text explanation that conveys its relationships, sequence, or conclusion.
- [ ] Informative images have meaningful alt text or an adjacent text equivalent.
- [ ] Link text and headings make sense to screen-reader users.
- [ ] UI instructions use input-neutral verbs unless the input method matters.
- [ ] Layout-dependent references such as **above**, **below**, **left**, and **right** are avoided.
- [ ] Sentences avoid ambiguous word order and unclear antecedents.
- [ ] Idioms, colloquialisms, seasonal assumptions, and local cultural references are avoided.

## Final verification

- [ ] Relevant documentation linting or formatting checks pass.
- [ ] Links and anchors were checked.
- [ ] Commands or code samples were run when feasible.
- [ ] Renderer-native diagrams, maps, and 3D models were previewed in every required renderer when feasible.
- [ ] **Change only:** the diff contains no unrelated rewriting or accidental technical changes.
- [ ] Material assumptions, unverified details, and missing information are reported to the maintainer.
