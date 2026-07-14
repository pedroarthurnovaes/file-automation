# File Automation

A safety-first Python utility for validating, archiving, and replacing recurring Excel exports.

This repository is a public, synthetic case study of a real operational automation. It demonstrates how human requirements and review were translated into a guarded workflow with Codex-assisted implementation, regression testing, and documented oversight. It contains no production workbooks or company-specific paths.

## Why it exists

Replacing a recurring workbook sounds simple until a wrong download, stale extract, locked file, or partial move leaves the active inventory in an unreliable state. File Automation turns that manual sequence into a visible plan with explicit safety gates.

## Safety controls

- Dry-run is the default.
- Apply mode requires `--apply --confirm REPLACE`.
- Active and incoming files must contain a usable configured date column.
- Incoming data cannot be older than active data by default.
- Existing archives are not overwritten by default.
- The active workbook is restored if replacement fails after archiving.
- Paths and business terms live in configuration, not source code.

## Quick start

Requires Python 3.11 or later.

```powershell
python -m pip install -r requirements.txt
python create_demo.py
python file_automation.py --config config.example.json
```

The last command is a dry run. It prints the proposed archive and replacement paths without changing either workbook.

To exercise apply mode against the synthetic demo only:

```powershell
python file_automation.py --config config.example.json --apply --confirm REPLACE
```

Windows users can also double-click `run_dry_run.bat` and `run_apply.bat` after installing Python and the requirements.

## Configuration

Copy `config.example.json` to `config.json` and adapt the values:

| Setting | Meaning |
|---|---|
| `active_dir` | Folder containing exactly one active workbook |
| `incoming_dir` | Folder receiving new exports |
| `archive_dir` | Destination for the dated previous version |
| `active_stem` | Active filename without its extension |
| `incoming_stem` | Expected export filename without its extension |
| `date_column` | Column used to compare data freshness and name the archive |
| `require_incoming_not_older` | Reject stale incoming data when `true` |
| `overwrite_existing_archive` | Allow replacement of a matching archive when `true` |

Relative paths are resolved from the configuration file's directory, which makes configurations portable.

## Validation

```powershell
python -m unittest discover -s tests -v
python -m py_compile file_automation.py create_demo.py
```

The test suite uses temporary folders and generated workbooks. It never touches production files.

## Project evidence

- [Case study](docs/case-study.md): problem, iteration history, human oversight, risks, and measurement plan.
- [Repository instructions](AGENTS.md): durable privacy, safety, and verification rules for Codex.
- [Automated tests](tests/test_file_automation.py): regression and failure-recovery coverage.
- [CI workflow](.github/workflows/tests.yml): validation on pushes and pull requests.

## Responsible use

Review dry-run output before applying changes. Test any customized configuration against disposable copies first. This utility is an operational guard, not a substitute for access controls, backups, or business-specific data-quality validation.

## License

MIT License. See [LICENSE](LICENSE).
