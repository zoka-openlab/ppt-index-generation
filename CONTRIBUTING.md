# Contributing

Thank you for helping improve PPT Index Generation.

## Development setup

```bash
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install -r requirements-dev.txt
python3 -m unittest discover -s tests -v
```

## Contribution rules

- Never commit real presentations, generated indexes, customer information, or absolute user paths.
- Use fictional names and synthetic data in tests and examples.
- Preserve the read-only guarantee for source presentations.
- Add or update tests for changes to indexing, slide selection, path handling, or output behavior.
- Keep `SKILL.md` concise. Put user-facing project documentation at the repository root, not inside the installable Skill folder.

Before opening a pull request, run:

```bash
python3 -m unittest discover -s tests -v
python3 /path/to/skill-creator/scripts/quick_validate.py skill/ppt-index-generation
```
