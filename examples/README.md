# Generic examples

These examples use placeholders only. Keep real presentations and generated indexes outside the repository.

## Build an index

```bash
python3 skill/ppt-index-generation/scripts/build_ppt_index.py \
  "/path/to/overview.pptx" \
  "/path/to/case-studies.pptx" \
  --output-dir "/private/output/PPT_Index"
```

## Extract a candidate pack

```bash
python3 skill/ppt-index-generation/scripts/extract_candidate_slides.py \
  --source "/path/to/overview.pptx" \
  --slides "1,4,7-9" \
  --output "/private/output/customer-outcomes-candidates.pptx" \
  --notes "/private/output/customer-outcomes-candidates.md" \
  --title "Customer outcomes candidate pack"
```

The extraction command preserves the requested order. For example, `--slides "3,1"` creates a two-slide output containing source slide 3 followed by source slide 1.
