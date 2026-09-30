# Markdown variants

Use this reference before adding or changing alerts, callouts, internal links, embeds, or other syntax that differs across Markdown renderers.

## Identify the target

1. Identify every place the source must render, not only the application used to edit it.
2. Inspect repository evidence such as contribution rules, site-generator configuration, Markdown linting, `.obsidian/` settings, and nearby documents.
3. Preserve a consistent, working convention when the repository already has one.
4. Verify the result in each required renderer or with its documented syntax.

Do not assume that every file hosted on GitHub targets GitHub Flavored Markdown (GFM), or that every Markdown file in an Obsidian vault should use Wikilinks.

Ask the user which convention to use when repository evidence remains ambiguous and the choice changes rendering, portability, or link behavior. Ask before a broad link conversion, introducing a custom callout taxonomy, or choosing between incompatible renderers. Do not interrupt a localized edit that can preserve a clear existing convention.

When one source must work in several renderers, use their verified common subset. If that subset cannot express a required behavior, document the compatibility tradeoff and ask which destination takes priority.

## Decide whether to use a callout

Default to normal prose. Use a callout only when brief, exceptional information needs extra prominence because overlooking it could change the reader's result.

Use a callout for:

- A prerequisite, constraint, or compatibility fact that readers must notice before acting.
- A material risk such as data loss, security exposure, unexpected cost, difficult recovery, or an incorrect result.
- A non-obvious condition that commonly causes a task to fail.
- Short supplemental context or an optional technique that is valuable to readers who are scanning.

Use a heading, paragraph, list, or procedure instead when the content belongs to the normal reading flow, contains required steps, needs a stable link target, or is too substantial to remain an aside. Never rely on a callout's color or icon to communicate meaning.

If the repository does not define its own type semantics, use these defaults:

| Type | Use |
| --- | --- |
| `NOTE` | Supplemental context that changes how readers interpret nearby content. |
| `TIP` | An optional technique that improves the task without changing its requirements. |
| `IMPORTANT` | A prerequisite, constraint, or action required to achieve the intended result. |
| `WARNING` | An urgent condition readers must address to avoid failure or an incorrect result. |
| `CAUTION` | A risk with material negative consequences, such as data loss, security exposure, or difficult recovery. |

Choose a type for its meaning, not its color or icon. Avoid consecutive callouts and repeated callouts that train readers to ignore them.

## Write GitHub alerts

GitHub-rendered Markdown supports `NOTE`, `TIP`, `IMPORTANT`, `WARNING`, and `CAUTION` alerts:

```markdown
> [!WARNING]
> Back up the configuration before replacing it.
```

Keep the type marker on its own line. Do not add an Obsidian-style custom title, fold marker, or nested callout. GitHub recommends limiting alerts to one or two per article and does not support alerts nested inside other elements.

## Write Obsidian callouts

Obsidian supports the shared types plus additional types, aliases, custom titles, folding, nesting, and vault-defined custom types:

```markdown
> [!warning] Back up the vault
> Create a recoverable copy before changing the plugin configuration.
```

- Use additional or custom types only when the vault defines or consistently uses their meaning.
- Ask before introducing a custom type that depends on a CSS snippet or plugin.
- Use foldable callouts only for optional detail. Do not collapse prerequisites, required actions, or safety-critical information.
- Avoid nesting unless the document is Obsidian-only and the hierarchy materially improves understanding.

For a document that must render in both GitHub and Obsidian, use one of the five shared type markers on its own line, without a custom title, folding, or nesting. Verify the rendered result because the two applications can style the same type differently.

## Choose internal link syntax

Preserve the repository or vault's established syntax. Do not convert links merely to make their source look uniform; link syntax can encode authoring, navigation, graph, embed, and portability decisions.

| Target | File link | Heading link |
| --- | --- | --- |
| GitHub-rendered Markdown | `[Contributing](../CONTRIBUTING.md)` | `[Configuration](guide.md#configuration)` |
| Obsidian Wikilink | `[[Guides/Contributing|Contributing]]` | `[[Guide#Configuration|Configuration]]` |
| Obsidian Markdown link | `[Contributing](Guides/Contributing.md)` | `[Configuration](Guide.md#Configuration)` |

For GitHub-rendered Markdown, resolve relative paths from the current file and verify generated heading anchors after heading changes.

For Obsidian, inspect the vault's **Files and links** settings and nearby notes before choosing Wikilinks or Markdown links. Wikilink folder paths are vault-relative; use an alias after `|` for descriptive display text. In Markdown links, URL-encode destinations, including spaces as `%20`, and follow the vault's configured path convention. Obsidian block references, embeds such as `![[image.png]]`, and some heading behaviors are not portable Markdown.

For a source that must work in both GFM and Obsidian, prefer standard Markdown links with explicit file extensions and paths verified in both destinations. Use standard Markdown image syntax with useful alt text instead of Obsidian embeds. Test links and anchors in every required renderer.
