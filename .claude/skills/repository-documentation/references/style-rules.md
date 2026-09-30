# Style rules

Use these rules for repository prose written for product users, operators, integrators, contributors, or maintainers. Apply judgment when a rule conflicts with technical accuracy or an established repository convention.

## Voice and tone

- Sound like a knowledgeable teammate, not a marketer, legal notice, or classroom lecturer.
- Be conversational, friendly, and respectful without slang or excessive casualness.
- Lead with the result, action, constraint, or decision that matters most.
- Give readers enough context to act confidently, then stop.
- Avoid exaggerated claims such as **best**, **seamless**, **powerful**, **revolutionary**, or **production-ready** unless they are objectively supported and relevant.
- Avoid pre-announcing later content. State the information where readers need it.

## Person, voice, tense, and mood

- Use second person for reader actions: **You can configure...**
- Use imperative mood for procedures: **Run the tests.**
- Use active voice unless passive voice is useful because the actor is unknown or irrelevant.
- Use present tense for current behavior.
- Name the actor when responsibility matters.
- Avoid anthropomorphism. Software can **return**, **calculate**, or **reject**; it does not **want**, **know**, or **decide** unless those terms describe an actual system design.

## Sentence structure

- Put conditions before instructions: **If the cache is stale, delete it.**
- Put context before an action: **In the repository root, run `npm test`.**
- Keep the main action near the beginning of the sentence.
- Break sentences that contain several exceptions, parenthetical clauses, or semicolons.
- Use parallel structure for related headings, steps, and list items.
- Make pronoun references explicit, especially across sentences.

## Word choice and terminology

- Use one term for one concept. Do not alternate among synonyms for variety.
- Prefer common, precise verbs: **use**, **run**, **create**, **update**, **remove**, **return**, **send**.
- Replace vague verbs such as **handle**, **process**, **manage**, or **support** when a more exact action is known.
- Define unfamiliar abbreviations on first use unless the intended audience will unquestionably know them.
- Avoid Latin abbreviations such as **e.g.** and **i.e.** in normal prose; use **for example** and **that is**.
- Avoid **etc.**, **and so on**, and similar endings when the omitted items matter. Give representative examples or describe the category.
- Avoid idioms, puns, regional phrases, and culture-specific references.
- Avoid **simply**, **just**, **easy**, **obvious**, **clearly**, and **of course** when they judge the reader's experience.

## Inclusive and accessible language

- Refer to people by relevant roles or actions, not stereotypes or unnecessary personal characteristics.
- Use people-first or identity-first disability language according to the preference of the community or person being discussed. When no preference is known, use respectful neutral wording.
- Avoid gendered defaults such as **guys**, **man-hours**, or generic **he**.
- Prefer precise alternatives to exclusionary legacy terms, such as **allowlist** and **blocklist**, unless the legacy term is a literal identifier that must be reproduced exactly.
- Avoid violent or ableist metaphors when literal language works better.
- Do not rely on color, position, shape, sound, or an image alone to communicate required information.
- Give informative images meaningful alt text. Use empty alt text only for purely decorative images when the format supports it.

## Capitalization and punctuation

- Use sentence case for document titles, headings, captions, table headings, and prose labels.
- Preserve official casing for products, technologies, code, filenames, commands, flags, configuration keys, and literal UI text.
- Do not end headings with a period. Use a question mark only when the heading is genuinely a question.
- Use standard American punctuation unless the repository specifies another locale.
- Use the serial comma in lists of three or more items.
- Use one space after sentence-ending punctuation.
- Prefer commas, periods, or separate sentences over semicolons.
- Use em dashes sparingly. Do not use repeated hyphens as a substitute.

## Headings and organization

- Use one level-1 heading per standalone Markdown page.
- Do not skip heading levels.
- Make headings unique, descriptive, and understandable when viewed in a table of contents.
- Start task headings with a base-form verb: **Configure authentication**.
- Use noun phrases for concepts: **Authentication configuration**.
- Avoid headings that begin with an **-ing** form when a clearer task or noun phrase is available.
- Keep headings short and place distinguishing words early.
- Avoid two consecutive headings with no meaningful content between them unless the document system explicitly uses such structure.
- Break long sections into scannable units, but do not create headings for one-sentence fragments without a structural reason.

## Paragraphs

- Keep each paragraph focused on one idea.
- Put the topic sentence first when practical.
- Prefer short paragraphs for online reading.
- Use a list only when it improves scanning or shows a relationship among items.

## Lists

- Introduce a list with a complete sentence when context is needed.
- Use bullets for unordered items, including labeled items that are not genuine term–description relationships.
- Use numbers for sequences, priorities, or items referenced by number.
- Use `<dl>`, `<dt>`, and `<dd>` for genuine term–description relationships when the target renderer supports description lists.
- Use parallel grammatical structure.
- Begin each item with a capital letter unless the item is a literal that requires different casing.
- End sentence-like items with periods.
- Omit periods for single words, short noun phrases, pure code items, or link-title lists.
- Do not make list items complete an introductory sentence fragment in documentation intended for localization.
- Do not end list items with semicolons or coordinating conjunctions.

## Procedures

- State the goal or necessary context before the steps.
- Use a numbered list for a multi-step procedure.
- Use a bullet or a normal sentence for a single-step procedure.
- Start each step with an imperative verb unless a short context phrase must come first.
- Use one principal action per step. Combine only small actions that occur together and cannot reasonably be separated.
- Tell readers where to act before telling them what to do.
- Include the final action required to complete or save the task.
- Mark optional steps with **Optional:** at the beginning.
- Describe the expected result when readers need confirmation.
- Prefer the shortest method that is accessible to the full audience.
- Link to a shared procedure instead of duplicating it.

## Links and cross-references

- Use short, descriptive link text that identifies the destination or purpose.
- Make link text meaningful when read without surrounding prose.
- Prefer the destination title when it fits naturally.
- Avoid **click here**, **here**, **learn more**, **this**, and raw URLs as link text.
- Use **see** for optional or supporting cross-references.
- Choose internal link syntax for the target Markdown dialect and established repository convention. See [Markdown variants](markdown-variants.md).
- When using standard Markdown, prefer relative links for files in the same repository.
- Do not use **above**, **below**, **left**, or **right** as the only way to identify referenced content.
- Check anchors after changing headings.

## References the reader can follow

Documentation travels with the repository, so a reference resolves only for a reader who also receives its target. Content the repository does not track never reaches them: ignored working directories, generated output, dependency trees, build artifacts, and local scratch files exist only on the machine that produced them.

- Do not link to, quote, paraphrase, or describe the contents of an untracked file. The author's copy is usually the only one that ever existed.
- Naming an untracked path is useful when readers need to know that it exists, must create it, or should not commit it. Stating what they will find inside it is not.
- Use repository-relative paths. An absolute path from one machine, such as a home directory, resolves for nobody else.
- Treat the wish to cite untracked content as evidence that something inside it belongs in the repository. Extract the fact the reader needs — a decision, a constraint, a procedure, a value — into the tracked document whose audience consults it, then reference that home.
- Move the fact, not the file. Progress notes, task state, and abandoned approaches do not become useful to readers by being committed, and content kept out of the repository deliberately, such as secrets and generated output, stays out.
- When the untracked content is unavailable, the fact cannot be extracted. Remove the reference and record what is missing where the repository tracks documentation gaps. A citation nobody can follow disguises the loss; a recorded gap makes it something a maintainer can restore.

Useful:

- Implementation plans live in `.agents.local/`, which the repository ignores.
- Copy `.env.example` to `.env` and set <code><var>API_TOKEN</var></code>. The repository ignores `.env`.
- The limit is 8 because the upstream API rejects more concurrent requests.
- The benchmark that set this limit is not recorded anywhere in the repository. (Written where the repository tracks documentation gaps, after the citation to an unavailable file was removed.)

Not useful:

- See the migration plan in `.agents.local/plans/migration.md`.
- The rationale for this limit is recorded in the working notes.

This applies with particular force to agent working files. An agent that writes a plan, a report, or a progress ledger into an ignored directory and then cites it from committed documentation has described a repository that only it can see.

## Alerts and callouts

Follow [Markdown variants](markdown-variants.md) to decide whether information warrants an alert or callout, select a type by meaning, and use syntax supported by every required renderer.

## Code, commands, and literals

- Format placeholders according to [Semantic HTML for technical text](#semantic-html-for-technical-text).
- Use inline code for literal technical content that does not require more specific semantic markup.
- Use fenced code blocks with the correct language identifier by default.
- Use the language-tagged `<pre><code>` structure from [Semantic HTML for technical text](#semantic-html-for-technical-text) when nested semantic elements must remain active within a preformatted example.
- Preserve the language's formatting conventions in code samples.
- Keep examples focused on realistic tasks.
- Explain non-obvious prerequisites and output, and explain placeholders whose expected value, format, or source is not evident from context.
- Do not use real credentials, personal data, private domains, or unsafe defaults.
- Do not use `...` to hide required code. Use a language-appropriate comment to identify omitted sections.
- Comment why something is necessary, not what an obvious line already states.
- Show error handling when it is part of the behavior being taught, not as unrelated noise.

## Semantic HTML for technical text

When the target renderer supports raw HTML, choose elements according to the role of the content.

- **`<var>PLACEHOLDER</var>`**: Use when a replaceable value is discussed as a variable in prose rather than presented as literal technical syntax. Example: Replace <var>PROJECT_NAME</var> with the name of the project.
- **`<code><var>PLACEHOLDER</var></code>`**: Use when the placeholder represents a literal technical value, such as a command argument, identifier, path, filename, configuration value, or environment variable.
- **`<code>...<var>PLACEHOLDER</var>...</code>`**: Use when a placeholder appears within a larger inline command, expression, path, or other technical syntax.
- **`<code>INPUT</code>`**: Use for literal technical text that the reader should enter, excluding individual keyboard keys and shortcuts.
- **`<kbd>KEY</kbd>`**: Use for an individual keyboard key or keyboard shortcut that the reader should press.
- **`<samp>OUTPUT</samp>`**: Use for text produced or displayed by a system, application, command, or process, including a nonactionable UI label mentioned descriptively.
- **`<kbd><samp>CONTROL</samp></kbd>`**: Use for a system-presented control or choice that the reader should activate, such as a button, tab, menu item, or dropdown option.

Use uppercase placeholder names with underscores unless literal syntax or repository conventions require another form.

Examples:

- Replace <var>PROJECT_NAME</var> with a descriptive name.
- Save the file as <code><var>FILENAME</var></code>.
- Run <code>tool --output <var>DIRECTORY</var></code>.
- Enter <code>npm install</code>.
- Press <kbd>Ctrl</kbd>+<kbd>C</kbd>.
- The command returns <samp>Installation complete.</samp>.
- Select <kbd><samp>Save</samp></kbd>.

Use fenced code blocks with language identifiers when the block does not contain active semantic HTML. When semantic elements such as `<var>` must remain active inside a preformatted example, use this compatibility structure:

```html
<pre lang="LANGUAGE_ID" class="language-LANGUAGE_ID"><code class="language-LANGUAGE_ID">COMMAND <var>PLACEHOLDER</var></code></pre>
```

Replace every `LANGUAGE_ID` placeholder with the same supported language identifier. For example:

```html
<pre lang="powershell" class="language-powershell"><code class="language-powershell">Get-ChildItem -LiteralPath <var>DIRECTORY</var></code></pre>
```

This structure is a repository house convention verified in GitHub-rendered Markdown and Obsidian Reading view.

The `lang` attribute is a GitHub renderer compatibility hint in this structure; it is not standards-based HTML language metadata. The `language-*` classes provide the language hint used by compatible syntax highlighters.

Escape characters that HTML would interpret inside raw HTML code blocks. In particular, write `&` as `&amp;` and `<` as `&lt;` when those characters are part of the displayed code rather than markup. A literal `>` generally does not require escaping, but `&gt;` may be used when needed to avoid ambiguity.

Do not use semantic HTML elements for visual styling alone.

## UI instructions

- Focus on the task rather than naming every UI control.
- Default to input-neutral verbs: **select**, **open**, **go to**, **choose**, **enter**, **clear**, and **turn on**.
- Use **click** or **tap** only when the input method is relevant.
- Do not use **click on**.
- Format UI labels and actionable controls according to [Semantic HTML for technical text](#semantic-html-for-technical-text).
- Match the displayed UI label exactly unless repository conventions deliberately normalize it.
- Put the application, page, dialog, or field before the action when needed. Example: In <samp>Settings</samp>, select <kbd><samp>Updates</samp></kbd>.
- Use menu paths such as <kbd><samp>File</samp></kbd> &gt; <kbd><samp>New</samp></kbd> &gt; <kbd><samp>Project</samp></kbd> only when the format is accessible to the intended audience; otherwise, describe the navigation in words.

## Dates, versions, and future behavior

- Use exact, unambiguous dates when a date matters.
- Include a time zone when time-sensitive behavior crosses regions.
- Tie version-specific behavior to an explicit version or range.
- Distinguish current, deprecated, removed, experimental, and planned behavior.
- Do not promise unreleased features or dates without an authoritative source.
- In deprecation notices, state the replacement and the required migration action.

## Tables, diagrams, and images

- Use a table only when readers need to compare values across consistent attributes.
- Use a list when the information is primarily sequential or descriptive.
- Give each column a clear heading and keep cell structure consistent.
- Do not place long procedures or large prose blocks in tables.
- Follow [diagram guidance](diagrams.md) to decide whether a diagram helps, choose syntax supported by the target renderer, and keep the visual maintainable.
- Introduce diagrams and images in the text and explain the information they add.
- Provide a text equivalent for information that is essential to the task.
