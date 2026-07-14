import json
import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

import file_automation as guard


def make_workbook(path: Path, values: list[object], column: str = "Last Updated") -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["Asset", column])
    for index, value in enumerate(values, start=1):
        sheet.append([f"SAMPLE-{index}", value])
    workbook.save(path)


def make_config(root: Path, **overrides: object) -> guard.Config:
    values = {
        "active_dir": root / "active",
        "incoming_dir": root / "incoming",
        "archive_dir": root / "archive",
        "active_stem": "Current_Inventory",
        "incoming_stem": "New Inventory Export",
        "date_column": "Last Updated",
        "require_incoming_not_older": True,
        "overwrite_existing_archive": False,
    }
    values.update(overrides)
    return guard.Config(**values)


class FileAutomationTests(unittest.TestCase):
    def test_load_config_resolves_relative_paths(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            config_path = root / "config.json"
            config_path.write_text(
                json.dumps(
                    {
                        "active_dir": "active",
                        "incoming_dir": "incoming",
                        "archive_dir": "archive",
                        "active_stem": "Current_Inventory",
                        "incoming_stem": "New Inventory Export",
                        "date_column": "Last Updated",
                    }
                ),
                encoding="utf-8",
            )

            config = guard.load_config(config_path)

            self.assertEqual(config.active_dir, root / "active")
            self.assertEqual(config.incoming_dir, root / "incoming")

    def test_find_incoming_allows_duplicate_suffix_and_trailing_space(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            config = make_config(root)
            config.incoming_dir.mkdir()
            candidate = config.incoming_dir / "New Inventory Export (2) .xlsx"
            make_workbook(candidate, [date(2026, 7, 14)])

            self.assertEqual(guard.find_incoming_file(config), candidate)

    def test_build_plan_uses_workbook_dates(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            config = make_config(root)
            active = config.active_dir / "Current_Inventory.xlsx"
            incoming = config.incoming_dir / "New Inventory Export.xlsx"
            make_workbook(active, [date(2026, 6, 1), date(2026, 6, 30)])
            make_workbook(incoming, [date(2026, 7, 14)])

            plan = guard.build_plan(config)

            self.assertEqual(plan.active_date, date(2026, 6, 30))
            self.assertEqual(plan.incoming_date, date(2026, 7, 14))
            self.assertEqual(
                plan.archive_target,
                config.archive_dir / "Current_Inventory_20260630.xlsx",
            )

    def test_rejects_an_older_incoming_workbook(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            config = make_config(root)
            make_workbook(
                config.active_dir / "Current_Inventory.xlsx", [date(2026, 7, 14)]
            )
            make_workbook(
                config.incoming_dir / "New Inventory Export.xlsx", [date(2026, 7, 1)]
            )

            with self.assertRaisesRegex(guard.GuardError, "Incoming workbook is older"):
                guard.build_plan(config)

    def test_rejects_missing_date_column(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            workbook = Path(temporary) / "sample.xlsx"
            make_workbook(workbook, [date(2026, 7, 14)], column="Wrong Column")

            with self.assertRaisesRegex(guard.GuardError, "was not found"):
                guard.max_column_date(workbook, "Last Updated")

    def test_refuses_to_overwrite_an_archive_by_default(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            config = make_config(root)
            make_workbook(
                config.active_dir / "Current_Inventory.xlsx", [date(2026, 6, 30)]
            )
            make_workbook(
                config.incoming_dir / "New Inventory Export.xlsx", [date(2026, 7, 14)]
            )
            make_workbook(
                config.archive_dir / "Current_Inventory_20260630.xlsx",
                [date(2026, 6, 30)],
            )

            with self.assertRaisesRegex(guard.GuardError, "Archive already exists"):
                guard.build_plan(config)

    def test_apply_moves_both_files(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            config = make_config(root)
            active = config.active_dir / "Current_Inventory.xlsx"
            incoming = config.incoming_dir / "New Inventory Export.xlsx"
            make_workbook(active, [date(2026, 6, 30)])
            make_workbook(incoming, [date(2026, 7, 14)])
            plan = guard.build_plan(config)

            guard.apply_plan(plan)

            self.assertFalse(incoming.exists())
            self.assertTrue(plan.archive_target.exists())
            self.assertTrue(plan.replacement_target.exists())
            self.assertEqual(
                guard.max_column_date(plan.replacement_target, "Last Updated"),
                date(2026, 7, 14),
            )

    def test_apply_rolls_back_when_replacement_fails(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            config = make_config(root)
            active = config.active_dir / "Current_Inventory.xlsx"
            incoming = config.incoming_dir / "New Inventory Export.xlsx"
            make_workbook(active, [date(2026, 6, 30)])
            make_workbook(incoming, [date(2026, 7, 14)])
            plan = guard.build_plan(config)
            real_move = guard.shutil.move
            calls = 0

            def fail_second_move(source: str, target: str) -> str:
                nonlocal calls
                calls += 1
                if calls == 2:
                    raise OSError("simulated replacement failure")
                return real_move(source, target)

            with patch.object(guard.shutil, "move", side_effect=fail_second_move):
                with self.assertRaisesRegex(guard.GuardError, "No replacement was completed"):
                    guard.apply_plan(plan)

            self.assertTrue(active.exists())
            self.assertTrue(incoming.exists())
            self.assertFalse(plan.archive_target.exists())


if __name__ == "__main__":
    unittest.main()
