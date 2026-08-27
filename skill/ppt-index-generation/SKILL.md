---
name: ppt-index-generation
description: Build and use AI-readable indexes for user-provided PowerPoint (.pptx) files. Use when the user supplies PPTX paths and wants Codex to index existing decks, search their contents by topic, identify relevant source slides, or extract selected original slides into a focused material pack without modifying or rewriting slide content.
---

# PPT Index Generation

## Core rule

Treat every source PPTX as read-only. Never move, rename, overwrite, or save over a source file. Write indexes, extracted slide packs, and notes to a separate output folder.

This Skill retrieves and collects existing material. Do not use it to rewrite slide content, redesign pages, generate new slides, or add covers, directories, transitions, or closing pages.

## Workflow

1. **Build the index**
   - Accept only user-provided PPTX paths.
   - Generate a master index and one page-level index per deck.
   - Record absolute source paths and slide numbers so future work can return to the original material.

2. **Search and collect relevant material**
   - Search the indexes for the user's theme.
   - Prefer recall over precision: include plausible pages even when some may be redundant.
   - Produce a topic summary listing relevant decks and source slide numbers.
   - By default, create one extracted material pack per source deck and one cross-deck Markdown summary.
   - Combine slides from different source decks into one PPTX only when the user explicitly requests it and the source decks have compatible slide sizes and formats.
   - Extract selected source slides unchanged.
   - Write notes listing source deck, source slide number, title, relevance, and possible duplication.

3. **Refine when requested**
   - Let the user confirm which existing slides to keep, remove, or reorder when the first collection contains duplicates or uncertain matches.
   - Apply the confirmed selection by extracting the original pages again without editing their contents.
   - Keep a page map from every output slide to its source.

## Output structure

Use a task-specific folder unless the user specifies another location:

```text
PPT_Index/
  MASTER_INDEX.md
  indexes/
    <deck-name>--<deck-id>.md

outputs/
  <theme-slug>/
    <theme>_候选页总说明.md
    <source-deck>_候选页材料包.pptx
    <source-deck>_候选页说明.md
    <source-deck>_确认版材料包.pptx
    <source-deck>_确认版说明.md
```

## Bundled resources

- Use `scripts/build_ppt_index.py` to create `MASTER_INDEX.md` and per-deck indexes.
- Use `scripts/extract_candidate_slides.py` to extract and reorder selected pages from one source deck and write page-map notes.
- Read `references/workflow.md` before selecting, extracting, or summarizing relevant pages.

## Quality checks

- Verify output PPTX files open and contain the expected slides in the requested order.
- Verify extracted pages preserve their original content and appearance.
- Verify each source deck has its own material pack unless the user explicitly requested a compatible combined pack.
- Verify source file size and modification time remain unchanged.
- Confirm outputs are outside source files and preferably outside source directories.
- Scan shareable outputs for sensitive paths, names, and extracted content before publishing.
- In final responses, link to the output PPTX and notes and state that source files were not modified.
