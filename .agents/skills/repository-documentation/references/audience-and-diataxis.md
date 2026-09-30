# Audience and Diátaxis architecture

Organize documentation around the reader's needs. Treat the audience as **who the reader is in relation to the product** and the Diátaxis form as **what the reader needs now**. These are independent dimensions.

## Define useful audiences

Name roles according to the product rather than adopting a fixed taxonomy. Common roles include:

- **Product user:** Uses the product to achieve a domain goal.
- **Operator or administrator:** Deploys, configures, secures, monitors, or supports it.
- **Integrator or API consumer:** Uses an API, SDK, plugin interface, or file format to build another product.
- **Contributor:** Changes the repository itself.
- **Maintainer:** Reviews, governs, releases, or supports the project.

Do not use **developer** as a catch-all when integrators and contributors have different goals. A developer who consumes an API is a user of that API; a contributor who changes the API is working on the product.

Identify one primary audience for each substantive page. Record secondary audiences only when the same page is genuinely useful to them without compromising its primary framing.

Ask the user whether separate audience paths are wanted only when the distinction would materially change navigation or content and the repository does not provide enough evidence to decide.

## Classify the reader's need

Choose one dominant form for each substantive page:

| Reader's immediate need | Diátaxis form | Writing focus |
| --- | --- | --- |
| Learn through a guided, successful experience | Tutorial | Skill acquisition |
| Complete a specific real-world task | How-to guide | Goal completion |
| Look up accurate facts while working | Reference | Predictable information retrieval |
| Understand concepts, reasons, or tradeoffs | Explanation | Mental models and context |

Do not classify by filename alone. For example, contributor setup can be a tutorial when it teaches a newcomer through a first success, or a how-to guide when it gets an experienced contributor operational.

Keep one dominant purpose per substantive page. Link to another page when a digression would interrupt that purpose. README files, landing pages, and navigation pages may intentionally introduce and route readers to several forms; do not force these entry points into one category.

## Decide whether to keep the existing structure

Treat the current documentation framework as evidence, not a requirement. Preserve it for localized changes when it is usable and the user has not put information architecture in scope. Evolve it when incremental improvement reduces disruption without preserving known problems. Replace it when the user rejects it or requests a fundamentally different framework and the task authorizes broad restructuring.

If a substantial restructuring request does not reveal the user's preference, ask whether to preserve, evolve, or replace the current structure before moving or duplicating many pages. Do not ask when the request or repository evidence already makes the intended scope clear.

## Choose the navigation hierarchy

For a new or substantially reorganized documentation set, choose the first navigation layer according to the distinction readers need to make first.

Use **audience first, then Diátaxis form** when audiences have substantially different goals, prerequisites, vocabulary, or journeys. This is often effective when product users, integrators, and contributors experience the product differently:

```text
docs/
├── users/
│   ├── tutorials/
│   ├── how-to/
│   └── reference/
├── integrators/
│   ├── tutorials/
│   ├── how-to/
│   └── reference/
└── contributors/
    ├── how-to/
    └── explanation/
```

Use **Diátaxis form first, then audience or topic** when audiences share most terminology and material, regularly move between roles, or benefit more from choosing a mode of help first.

Use a **hybrid** when learning and task content needs audience-specific framing but reference or explanation is genuinely shared. A documentation system may expose one canonical page from several navigation paths; physical location does not need to mirror every reader path.

Treat these trees as conceptual models, not required folders. Do not create empty sections to complete a four-part matrix. Add a section only when useful content and navigation justify it. Use landing pages to explain who a section serves and what its links contain; avoid long undifferentiated link lists.

## Handle multiple audiences

Choose among these patterns:

1. **Same wording and prerequisites:** Keep one canonical page and link to it from every relevant audience path.
2. **Clear primary audience:** Keep the page with that audience's material, identify secondary audiences when the documentation system supports it, and cross-link from their navigation.
3. **Different framing or workflow:** Create audience-specific pages and share only the narrower facts or concepts that remain identical.

Shared schemas, file formats, protocol specifications, glossaries, compatibility matrices, and some conceptual explanations often work as canonical multi-audience material. Tutorials and how-to guides more often need separate framing because prerequisites, context, and desired outcomes differ.

Do not require a top-level `shared/` directory. Use one only when it matches the site's information architecture and has a clear, narrow definition. A concept-based or reference-based canonical location is often easier to understand and maintain.

Avoid duplicating a page solely to make it appear under several audience directories. If the publishing system cannot reuse one page in several navigation locations, link to the canonical page from each audience landing page.

## Use metadata when it helps

Follow the repository's existing frontmatter and navigation schema. Do not introduce metadata merely to mirror the conceptual model.

When structured metadata is supported and a new schema is in scope, keep audience and form separate. For example:

```yaml
audience:
  primary: integrator
  secondary:
    - contributor
diataxis_type: reference
```

Define allowed audience values for the product. Use one `diataxis_type` value for a substantive page. Generate audience-specific navigation from metadata only when the documentation toolchain supports and validates that behavior.
