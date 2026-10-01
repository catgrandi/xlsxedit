# Design documents

This directory holds the formal specs and design documents for xlsxedit. Write
one here before implementing a change that needs a recorded decision, such as a
new public API, a change to the preservation model, or a cross-cutting
refactor. Most roadmap work is tracked in the
[roadmap issue](https://github.com/catgrandi/xlsxedit/issues/26).

Agent working files, such as plans and progress notes, do not belong here; see
[CONTRIBUTING.md](../../CONTRIBUTING.md#specs-and-working-files).

## Naming

Name each spec after the issue it serves: `NNN-short-slug.md`, where `NNN` is
the GitHub issue number padded to three digits and `short-slug` is a few
lowercase words joined by hyphens. For example, the spec for issue #42 would be
`042-short-slug.md`.

## Template

Start the spec with a title and a link to its issue, then cover four sections.
**Context** describes the problem, the constraints, and the current behavior.
**Decision** states what will be built and how it behaves. **Alternatives**
lists the options considered and why each was rejected. **Open questions**
records what is still undecided and who decides it.
