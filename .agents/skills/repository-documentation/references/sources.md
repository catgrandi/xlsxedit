# Sources and reconciliation

This skill is an original, condensed, and paraphrased operational synthesis of the following primary sources:

- [Microsoft Writing Style Guide](https://learn.microsoft.com/en-us/style-guide/welcome/), including [graphics, design, and media](https://learn.microsoft.com/en-us/style-guide/accessibility/graphics-design-media)
- [Google developer documentation style guide](https://developers.google.com/style), including [diagrams, figures, and other images](https://developers.google.com/style/images)
- [Agent Skills specification and authoring guidance](https://agentskills.io/home)
- [Diátaxis](https://diataxis.fr/), including its guidance on [complex hierarchies](https://diataxis.fr/complex-hierarchies/) and [using Diátaxis as a guide rather than a plan](https://diataxis.fr/how-to-use-diataxis/)
- [GitHub basic writing and formatting syntax](https://docs.github.com/en/get-started/writing-on-github/getting-started-with-writing-and-formatting-on-github/basic-writing-and-formatting-syntax)
- [GitHub creating diagrams](https://docs.github.com/en/get-started/writing-on-github/working-with-advanced-formatting/creating-diagrams)
- [Obsidian callouts](https://help.obsidian.md/callouts)
- [Obsidian advanced formatting syntax](https://help.obsidian.md/advanced-syntax)
- [Obsidian internal links](https://help.obsidian.md/links)

The source guides remain authoritative for their own organizations. This skill adapts their overlapping principles for repository documentation that serves product users, operators, integrators, contributors, and maintainers.

## Reconciliation decisions

When the guides differ, this skill uses the following defaults:

- **Technical-document structure:** Prefer Google's task, heading, procedure, code, and reference-document patterns because its guide is specifically aimed at developer documentation.
- **Audience and purpose:** Treat audience roles and Diátaxis forms as independent dimensions. Use product-specific roles, and classify each substantive page by one dominant reader need.
- **Information architecture:** Choose audience-first, Diátaxis-first, or hybrid navigation according to the distinction readers need to make first. Do not impose four top-level boxes or a complete audience-by-form matrix.
- **Multi-audience content:** Keep one canonical page when wording and prerequisites work across audiences. Use a primary audience location or a clearly defined shared location and expose the page through each relevant navigation path. Split pages when audiences need different framing or workflows.
- **Entry points:** Allow README files and landing or navigation pages to route readers across multiple Diátaxis forms rather than forcing those pages into one form.
- **Input methods:** Prefer Microsoft's input-neutral UI language. Use **select**, **open**, or **go to** unless a mouse, touch screen, or other modality is materially relevant.
- **List punctuation:** End sentence-like items with periods. Omit periods for single words, short phrases, pure code, or link-title lists.
- **Task headings:** Use a base-form verb without **To**, such as **Create a project**.
- **Semantic technical markup:** Use the repository house style for `<var>`, `<code>`, `<pre>`, `<kbd>`, and `<samp>`. This convention intentionally differs from portions of the Microsoft and Google guidance and takes precedence when the target renderer supports the markup.
- **Description lists:** Use `<dl>`, `<dt>`, and `<dd>` for genuine term–description relationships when the target renderer supports them; use bullets for other unordered or labeled items.
- **Code blocks:** Use fenced blocks with language identifiers by default. When nested semantic elements must remain active, use `<pre lang="LANGUAGE" class="language-LANGUAGE"><code class="language-LANGUAGE">...</code></pre>`.
- **Markdown dialects:** Inspect the repository's authoring and rendering environments before using dialect-specific callouts, links, or embeds. Preserve a clear local convention, use a verified common subset for multi-renderer documents, and ask the user when evidence cannot resolve a materially incompatible choice.
- **Diagrams:** Use a diagram only when it communicates a relationship or spatial structure more effectively than prose or a table. Prefer maintainable text-based source when renderers support it, and always provide an equivalent text explanation.
- **Duplication and synchronization:** Give each fact one authoritative home; link, summarize with a link, or generate everywhere else. Treat audience-specific framing as legitimate and copied values or contracts as defects. This extends the source guides' reuse guidance into an explicit repository policy, including the rule that documentation stating changed behavior is part of the change that alters it.
- **Repository truth:** Source code, schemas, configuration, tests, generated help, and explicit repository policies take precedence over generic editorial rules.

## Attribution and licensing

This package does not reproduce the source guides verbatim. Source content is subject to the licensing and terms published by its respective owner. Preserve this sources file when redistributing the skill so maintainers can review the underlying guidance and future changes.
