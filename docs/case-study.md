# File Automation case study: from a fragile manual update to a guarded workflow

## Business problem

A recurring workbook export had to replace an active inventory file while retaining a dated archive of the previous version. The manual sequence was easy to perform incorrectly: select the wrong download, derive the wrong archive date, overwrite an existing archive, or leave the active folder without a usable file after a partial failure.

## Human and Codex collaboration

The human owner defined the operational contract, reviewed file-selection assumptions, reproduced failures from the Windows and Excel environment, and decided which safeguards were acceptable. Codex helped translate those requirements into a repeatable Python workflow, launch scripts, regression tests, and operator-facing diagnostics.

The workflow improved iteratively after real review findings:

1. A downloaded filename containing whitespace before the extension was not detected. Filename normalization and a regression test were added.
2. An Excel or synchronization lock produced an unclear failure. The error now gives a specific recovery action.
3. Console execution did not make progress sufficiently visible. Operational status and a persistent log were added.
4. The public version added freshness validation, archive collision protection, explicit apply confirmation, and rollback after partial failure.

## Human oversight and risk controls

- Dry-run is the default and performs validation without moving files.
- Apply mode requires an explicit confirmation phrase.
- Both workbooks are checked for the configured date column.
- An older incoming extract is rejected by default.
- Existing archives are preserved unless the operator explicitly opts into replacement.
- If the incoming move fails after archiving, the active workbook is restored.
- Public examples contain synthetic data only.

## Evidence

- Automated tests cover filename anomalies, schema errors, stale incoming data, archive collisions, successful replacement, and rollback.
- GitHub Actions runs the tests on every push and pull request.
- `AGENTS.md` makes privacy, safety, and validation expectations durable for future Codex tasks.

## Outcome measurement

Before publishing a time-saved claim, record at least five comparable manual and automated runs. Report the median duration and the observed exception rate. Do not substitute an estimate for measured evidence.

Suggested measurement table:

| Measure | Manual baseline | Guarded workflow | Evidence source |
|---|---:|---:|---|
| Median operator time | To measure | To measure | Timestamped run log |
| Validation checks | To measure | 5 automated gates | Test and run output |
| Recoverable failure scenarios tested | 0 documented | 6+ | Automated tests |

## Limitations

- `.xls` is intentionally unsupported because `openpyxl` cannot safely read it.
- The tool validates the configured date field, not every business rule in a workbook.
- Network drives and synchronization clients can still introduce external locking behavior.
- Operators remain responsible for reviewing dry-run output before applying changes.
