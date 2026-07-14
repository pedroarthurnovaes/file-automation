from datetime import date
from pathlib import Path

from openpyxl import Workbook


ROOT = Path(__file__).resolve().parent


def create_workbook(path: Path, dates: list[date]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Inventory"
    sheet.append(["Asset", "Last Updated", "Status"])
    for index, updated in enumerate(dates, start=1):
        sheet.append([f"SAMPLE-{index:03d}", updated, "Active"])
    workbook.save(path)


def main() -> None:
    create_workbook(
        ROOT / "demo" / "active" / "Current_Inventory.xlsx",
        [date(2026, 6, 1), date(2026, 6, 30)],
    )
    create_workbook(
        ROOT / "demo" / "incoming" / "New Inventory Export.xlsx",
        [date(2026, 7, 1), date(2026, 7, 14)],
    )
    print("Synthetic demo workbooks created.")


if __name__ == "__main__":
    main()

