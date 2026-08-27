# PPT Index Generation Workflow

## Candidate selection

Search `MASTER_INDEX.md` first, then open relevant per-deck indexes. Derive related terms from the user's wording and vocabulary already present in the indexed decks. Do not rely on a fixed industry taxonomy.

Classify candidates:

- `strong`: the theme appears in the title or summary, or the slide directly addresses the request.
- `supporting`: the slide provides useful context, evidence, metrics, examples, or framing.
- `optional`: adjacent material that may help but may be redundant.
- `duplicate-risk`: similar titles, structures, or repeated content appear elsewhere.

## Material collection rules

Make the first collection intentionally generous. Include likely context slides and repeated versions when they may help the user choose. Extract original pages unchanged; do not edit text, layout, styling, or slide elements.

Include these fields in candidate notes:

- output slide number
- source PPTX path
- source slide number
- detected title
- short reason for inclusion
- candidate class
- duplication or uncertainty note

Use this packaging rule:

- Default: create one candidate pack per source deck and one cross-deck Markdown summary.
- One source deck: return one material pack plus its notes.
- Multiple source decks: group all selected pages from the same source deck into that deck's material pack; never create one file per slide.
- Combined PPTX: create one only when the user explicitly requests it and the source decks have compatible slide sizes and formats.
- Incompatible sources: keep separate packs even if the user initially asks for one file, explain the incompatibility, and preserve the original pages.

When the user asks only for a content summary, return a concise Markdown inventory with source deck and slide references. When the user asks for pages to be collected, also produce extracted PPTX packs.

## Optional confirmation

When duplicate or uncertain pages materially affect the result, ask the user to review the candidate collection. Accept instructions to keep, remove, or reorder existing pages. Confirmation is for selection only, not for creating new material.

## Output boundaries

- Preserve each selected slide as a complete original page.
- Do not rewrite, summarize onto new slides, redesign, or restyle source pages.
- Do not add covers, directories, section pages, transitions, or closing pages.
- Do not merge elements from different slides into a newly composed page.
- Keep the default per-source-deck packs and cross-deck summary unless an explicitly requested combined pack is compatible.
- Keep a page map that points every extracted page back to its source deck and slide number.

## Safety and privacy checks

Before reporting completion:

- reopen the generated PPTX and verify slide count and order
- compare source file size and modification time with the pre-run values
- confirm no output path resolves to a source path
- confirm extracted slides were not edited or supplemented
- keep a page map for traceability
- remind the user that indexes include absolute paths and extracted text
- remove or replace identifying examples before sharing any repository or demo
