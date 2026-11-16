"""Check sheet.py's guards on a small workbook built in a temporary folder. Run: python test_sheet.py"""

import json
import subprocess
import sys
import tempfile
from datetime import date
from pathlib import Path

import openpyxl
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

SCRIPT = Path(__file__).with_name("sheet.py")
CONFIG = {
    "workbook": "test",
    "output_dir": "out",
    "id_prefix": "ID_",
    "sheet": {
        "tab": "Feedback Log",
        "columns": {
            "id": "ID", "feedback": "Feedback", "status": "Status", "response": "Response",
            "priority": "Priority", "owner": "Assigned to", "issue": "Tracker issue",
        },
    },
    "statuses": {"open": "Open", "in_progress": "In progress", "closed": "Closed"},
    "priorities": ["1", "2", "3"],
    "owners": {"B": {"name": "Sam Example", "initials": "SE", "layer": "back end"}},
    "tracker": {"project": "Example", "queue_label": "uat-queue"},
}


def build(folder: Path) -> Path:
    """Write a workbook with a title row, a header row, a dropdown, a conditional format and four rows (one blank)."""
    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = "Feedback Log"
    book.create_sheet("Names")
    sheet.append(["UAT feedback"])
    sheet.append(["Feedback", "ID", "Status", "Response", "Priority", "Assigned to", "Tracker issue"])
    sheet.append(["Price wrong", "ID_4", "Open", "Earlier note", 1, "B", None])
    sheet.append(["Map missing", None, None, None, 2, "B", "https://tracker/ENG-1"])
    sheet.append([None, None, None, None, None, None, None])
    sheet.append(["Logo blurry", None, "Open", None, 3, None, None])
    dropdown = DataValidation(type="list", formula1='"Open,In progress,Closed"')
    sheet.add_data_validation(dropdown)
    dropdown.add("C3:C100")
    sheet.conditional_formatting.add("E3:E100", CellIsRule(operator="equal", formula=["1"], fill=PatternFill(bgColor="FFC7CE")))
    path = folder / "feedback.xlsx"
    book.save(path)
    (folder / "config.json").write_text(json.dumps(CONFIG))
    return path


def run(folder: Path, *arguments: str) -> subprocess.CompletedProcess:
    """Run sheet.py from the folder, as an agent would from the project root."""
    return subprocess.run([sys.executable, str(SCRIPT), *arguments], cwd=folder, capture_output=True, text=True)


def cell(path: Path, coordinate: str) -> object:
    """Read one cell of the feedback tab."""
    return openpyxl.load_workbook(path)["Feedback Log"][coordinate].value


def main() -> None:
    """Run each guard against a fresh workbook."""
    with tempfile.TemporaryDirectory() as temp:
        folder = Path(temp)
        workbook = build(folder)
        base = ["config.json", "feedback.xlsx"]

        # A copy run changes only the copy
        result = run(folder, "ids", *base, "--copy")
        assert result.returncode == 0, result.stderr
        assert cell(workbook, "B4") is None and "B4: None -> 'ID_5'" in result.stdout

        # UIDs continue from the largest, with the prefix, skipping the blank row
        result = run(folder, "ids", *base)
        assert result.returncode == 0, result.stderr
        assert (cell(workbook, "B4"), cell(workbook, "B6")) == ("ID_5", "ID_6")

        # A response goes on top, signed and dated, with the earlier note kept below
        result = run(folder, "write", *base, "--id", "ID_4", "--status", "In progress", "--response", "Found the cause", "--by", "SE")
        assert result.returncode == 0, result.stderr
        today = date.today()
        assert cell(workbook, "D3") == f"- {today.day} {today:%b}, SE - Found the cause\nEarlier note"
        assert cell(workbook, "C3") == "In progress"

        # An unknown status is refused and nothing is saved
        result = run(folder, "write", *base, "--id", "ID_4", "--status", "Done")
        assert result.returncode != 0 and cell(workbook, "C3") == "In progress"

        # A different issue link is refused; the same one is accepted
        assert run(folder, "write", *base, "--id", "ID_5", "--issue", "https://tracker/ENG-2").returncode != 0
        assert run(folder, "write", *base, "--id", "ID_5", "--issue", "https://tracker/ENG-1").returncode == 0

        # Duplicate UIDs, or a UID without the prefix, stop the ID run
        book = openpyxl.load_workbook(workbook)
        book["Feedback Log"]["B6"] = "ID_4"
        book.save(workbook)
        assert "Duplicate UIDs ['ID_4']" in run(folder, "ids", *base).stderr
        book["Feedback Log"]["B6"] = 6
        book.save(workbook)
        assert "does not start with 'ID_'" in run(folder, "ids", *base).stderr
        book["Feedback Log"]["B6"] = "ID_6"
        book.save(workbook)

        # rows --open leaves out the closed row and keeps the in-progress and open rows
        assert run(folder, "write", *base, "--id", "ID_5", "--status", "Closed").returncode == 0
        open_rows = [json.loads(line)["feedback"] for line in run(folder, "rows", *base, "--open").stdout.splitlines()]
        assert open_rows == ["Price wrong", "Logo blurry"], open_rows

        # Excel's lock file stops a real write but not a copy run
        (folder / "~$feedback.xlsx").write_text("")
        assert "open in Excel" in run(folder, "write", *base, "--id", "ID_6", "--status", "Closed").stderr
        assert run(folder, "write", *base, "--id", "ID_6", "--status", "Closed", "--copy").returncode == 0
        assert cell(workbook, "C6") == "Open"

        # The layout survived every write
        sheet = openpyxl.load_workbook(workbook)["Feedback Log"]
        assert len(sheet.data_validations.dataValidation) == 1 and len(sheet.conditional_formatting) == 1
    print("sheet.py: all checks passed")


if __name__ == "__main__":
    main()
