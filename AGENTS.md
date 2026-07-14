# Repository instructions

## Purpose

Maintain a public, synthetic demonstration of a safety-first workbook replacement workflow.
Never add company names, personal paths, production filenames, credentials, logs, or real workbooks.

## Working agreements

- Keep dry-run as the default. Applying changes must require both `--apply` and `--confirm REPLACE`.
- Preserve rollback behavior whenever file movement changes.
- Resolve relative paths against the configuration file, not the current shell directory.
- Keep the core tool usable on Windows with Python 3.11 or later.
- Add or update a regression test for every defect correction.
- Prefer actionable error messages that tell an operator how to recover safely.

## Validation

Run these commands after modifying Python files:

```powershell
python -m unittest discover -s tests -v
python -m py_compile file_automation.py create_demo.py
```

For an end-to-end check, generate the synthetic files and run dry-run mode:

```powershell
python create_demo.py
python file_automation.py --config config.example.json
```

Do not run apply mode against any directory other than a disposable test or demo directory.

## Definition of done

- Tests pass.
- Dry-run does not modify workbooks.
- No private or production information appears in the diff.
- README behavior and commands match the implementation.
- Safety limitations are documented rather than hidden.
