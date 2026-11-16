"""Read and write a UAT feedback sheet described by a project's uat-tracker.json.

Run from the project root:
  new   <config> <workbook> [--legacy] build an empty feedback sheet from the config, refusing if the file exists
  rows  <config> <workbook> [--open]   print every row with feedback as one JSON line; --open leaves out closed rows
  ids   <config> <workbook> [--copy]   give each row with feedback and no UID the prefix and the next number
  write <config> <workbook> --id UID [--id UID ...] [--status S] [--priority P] [--owner CODE] [--response TEXT --by INITIALS] [--issue URL] [--copy]
A write backs the workbook up first and prints every cell it changed; --copy writes to a copy instead.
"""

import argparse
import json
import shutil
from datetime import date, datetime
from pathlib import Path

import openpyxl
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.worksheet import Worksheet

# The new sheet's rows run from the header to here; Excel extends the table's dropdowns as rows are added
LAST_ROW = 300


def header(sheet: Worksheet, columns: dict[str, str]) -> tuple[int, dict[str, int]]:
    """Return the header row's number and each column key's column number; the header row is the first holding every configured header."""
    wanted = set(columns.values())
    for row in sheet.iter_rows():
        found = {cell.value: cell.column for cell in row if isinstance(cell.value, str) and cell.value in wanted}
        if len(found) == len(wanted):
            return row[0].row, {key: found[text] for key, text in columns.items()}
    raise SystemExit(f"No row holds every configured header: {sorted(wanted)}")


def read_rows(sheet: Worksheet, config: dict) -> tuple[dict[str, int], list[dict]]:
    """Return the column map and every row below the header that has feedback, each with its sheet row number."""
    header_row, columns = header(sheet, config["sheet"]["columns"])
    rows = []
    for number in range(header_row + 1, sheet.max_row + 1):
        values = {key: sheet.cell(number, column).value for key, column in columns.items()}
        # A row without feedback is an empty or formatting-only row
        if values["feedback"] in (None, ""):
            continue
        rows.append({"row": number, **values})
    return columns, rows


def uid_number(uid: object, prefix: str) -> int:
    """Return the number after the configured prefix in a UID cell, such as 12 from "ID_12"."""
    text = str(uid)
    if not text.startswith(prefix):
        raise SystemExit(f"UID {text!r} does not start with {prefix!r}; fix it in the sheet first. Nothing saved.")
    return int(text[len(prefix):])


def assign_ids(sheet: Worksheet, config: dict) -> list[str]:
    """Give each row with no UID the prefix and the next number after the largest in the column; refuse duplicates."""
    prefix = config["id_prefix"]
    columns, rows = read_rows(sheet, config)
    existing = [uid_number(row["id"], prefix) for row in rows if row["id"] not in (None, "")]
    duplicates = sorted({f"{prefix}{number}" for number in existing if existing.count(number) > 1})
    if duplicates:
        raise SystemExit(f"Duplicate UIDs {duplicates}; fix them in the sheet first. Nothing saved.")
    next_number = max(existing, default=0) + 1
    messages = []
    for row in rows:
        if row["id"] in (None, ""):
            uid = f"{prefix}{next_number}"
            sheet.cell(row["row"], columns["id"]).value = uid
            messages.append(f"sheet row {row['row']}: {uid}")
            next_number += 1
    return messages


def find_row(rows: list[dict], uid: str) -> dict:
    """Return the one row whose UID cell reads exactly this text."""
    matches = [row for row in rows if str(row["id"]) == uid]
    if len(matches) != 1:
        raise SystemExit(f"UID {uid} matches {len(matches)} rows, not one. Nothing saved.")
    return matches[0]


def write_rows(sheet: Worksheet, config: dict, args: argparse.Namespace) -> list[str]:
    """Set status, priority and owner, add a dated response line on top, and link an issue, on each row named by ID."""
    if not (args.status or args.priority or args.owner or args.response or args.issue):
        raise SystemExit("Give --status, --priority, --owner, --response or --issue.")
    if args.status and args.status not in config["statuses"].values():
        raise SystemExit(f"Status {args.status!r} is not one of {list(config['statuses'].values())}.")
    if args.priority and args.priority not in config["priorities"]:
        raise SystemExit(f"Priority {args.priority!r} is not one of {list(config['priorities'])}.")
    if args.owner and args.owner not in config["owners"]:
        raise SystemExit(f"Owner {args.owner!r} is not one of {list(config['owners'])}.")
    if args.response and not args.by:
        raise SystemExit("A response needs --by with the writer's initials.")
    columns, rows = read_rows(sheet, config)
    today = date.today()
    messages = []
    for uid in args.id:
        number = find_row(rows, uid)["row"]
        if args.status:
            sheet.cell(number, columns["status"]).value = args.status
        if args.priority:
            # A priority column usually takes whole numbers, so a digits-only level is written as a number
            sheet.cell(number, columns["priority"]).value = int(args.priority) if args.priority.isdigit() else args.priority
        if args.owner:
            sheet.cell(number, columns["owner"]).value = args.owner
        if args.response:
            cell = sheet.cell(number, columns["response"])
            line = f"- {today.day} {today:%b}, {args.by} - {args.response}"
            # Earlier entries stay below the new line, read from the cell as it is now
            earlier = str(cell.value).strip() if cell.value is not None else ""
            cell.value = f"{line}\n{earlier}" if earlier else line
        if args.issue:
            cell = sheet.cell(number, columns["issue"])
            if cell.value not in (None, "", args.issue):
                raise SystemExit(f"{uid} already links {cell.value!r}. Nothing saved.")
            cell.value = args.issue
        messages.append(f"{uid} (sheet row {number}) written")
    return messages


def layout(path: Path, tab: str) -> tuple[int, int, list[str]]:
    """Count the tab's dropdowns and conditional formats, and list the workbook's tabs."""
    book = openpyxl.load_workbook(path)
    sheet = book[tab]
    return len(sheet.data_validations.dataValidation), len(sheet.conditional_formatting), book.sheetnames


def changed_cells(before: Path, after: Path, tab: str) -> list[str]:
    """List every cell of the tab whose value differs between two copies of the workbook."""
    old = openpyxl.load_workbook(before)[tab]
    new = openpyxl.load_workbook(after)[tab]
    changes = []
    for row in range(1, max(old.max_row, new.max_row) + 1):
        for column in range(1, max(old.max_column, new.max_column) + 1):
            was, now = old.cell(row, column).value, new.cell(row, column).value
            if was != now:
                changes.append(f"{new.cell(row, column).coordinate}: {was!r} -> {now!r}")
    return changes


def guarded_write(config: dict, workbook: Path, args: argparse.Namespace) -> None:
    """Back the workbook up, apply the command to it or to a copy, then report the changed cells and check the layout survived."""
    if not args.copy and (workbook.parent / f"~${workbook.name}").exists():
        raise SystemExit("The workbook is open in Excel on this machine; close it first.")
    output_dir = Path(config["output_dir"])
    output_dir.mkdir(parents=True, exist_ok=True)
    stamp = f"{datetime.now():%Y%m%d-%H%M%S}"
    backup = output_dir / f"backup-{stamp}-{workbook.name}"
    shutil.copy2(workbook, backup)
    target = workbook
    if args.copy:
        target = output_dir / f"copy-{stamp}-{workbook.name}"
        shutil.copy2(workbook, target)

    tab = config["sheet"]["tab"]
    before = layout(target, tab)
    book = openpyxl.load_workbook(target)
    sheet = book[tab]
    messages = assign_ids(sheet, config) if args.command == "ids" else write_rows(sheet, config, args)
    book.save(target)

    print(f"{'copy' if args.copy else 'workbook'}: {target}\nbackup: {backup}")
    for line in messages + changed_cells(backup, target, tab):
        print(line)
    after = layout(target, tab)
    if before != after:
        raise SystemExit(f"Layout changed {before} -> {after}; restore {backup}.")


def new_workbook(config: dict, path: Path, legacy: bool) -> None:
    """Build an empty feedback sheet from the config: UID first and locked, dropdowns from the config, a guide tab and a names tab; legacy adds a column comparing with the system being replaced."""
    if path.exists():
        raise SystemExit(f"{path} already exists. Nothing saved.")
    columns = config["sheet"]["columns"]
    statuses = config["statuses"]
    legacy_header = "Same issue in legacy?"
    headers = [
        columns["id"], "Date", "Tester", "Type", "Reference", columns["feedback"], *([legacy_header] if legacy else []), "Links",
        columns["status"], columns["response"], columns["priority"], columns["owner"], columns["issue"],
    ]
    # Each header's column letter, such as {"Status": "H"}
    letter = {text: get_column_letter(number) for number, text in enumerate(headers, start=1)}

    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = config["sheet"]["tab"]
    sheet["A1"] = f"{config['tracker']['project']} | UAT feedback"
    sheet["A1"].font = Font(size=14, bold=True)
    for number, text in enumerate(headers, start=1):
        sheet.cell(3, number, text)
        wide = text in (columns["feedback"], columns["response"])
        sheet.column_dimensions[get_column_letter(number)].width = 60 if wide else 16
    style = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    sheet.add_table(Table(displayName="FeedbackLog", ref=f"A3:{letter[headers[-1]]}{LAST_ROW}", tableStyleInfo=style))
    sheet.freeze_panes = "B4"

    # The UID column is grey so it reads as a label, and long text wraps
    for number in range(4, LAST_ROW + 1):
        sheet[f"{letter[columns['id']]}{number}"].fill = PatternFill("solid", fgColor="E7E6E6")
        for text in (columns["feedback"], columns["response"]):
            sheet[f"{letter[text]}{number}"].alignment = Alignment(wrap_text=True, vertical="top")

    # A custom rule of FALSE makes Excel refuse anything typed; the script writes the file directly and is unaffected
    lock = DataValidation(
        type="custom", formula1="FALSE", showInputMessage=True, promptTitle=columns["id"],
        prompt="Filled in automatically. Quote it when you refer to a row.",
        errorStyle="stop", errorTitle=columns["id"], error="This is filled in automatically. Leave it blank.",
    )
    rules = {
        columns["id"]: lock,
        "Tester": DataValidation(type="list", formula1="Testers"),
        "Type": DataValidation(type="list", formula1='"Bug,Feature request,Question"'),
        columns["status"]: DataValidation(type="list", formula1=f'"{",".join(statuses.values())}"'),
        columns["priority"]: DataValidation(type="list", formula1=f'"{",".join(config["priorities"])}"'),
        columns["owner"]: DataValidation(type="list", formula1=f'"{",".join(config["owners"])}"'),
    }
    if legacy:
        rules[legacy_header] = DataValidation(type="list", formula1='"Yes,No,Not applicable"')
    for text, rule in rules.items():
        rule.showErrorMessage = True
        sheet.add_data_validation(rule)
        rule.add(f"{letter[text]}4:{letter[text]}{LAST_ROW}")
    status_range = f"{letter[columns['status']]}4:{letter[columns['status']]}{LAST_ROW}"
    for key, colour in (("open", "FFC7CE"), ("in_progress", "FFEB9C"), ("closed", "C6EFCE")):
        rule = CellIsRule(operator="equal", formula=[f'"{statuses[key]}"'], fill=PatternFill(bgColor=colour))
        sheet.conditional_formatting.add(status_range, rule)

    guide = book.create_sheet("How to use")
    guide["A1"] = "How to use"
    guide["A1"].font = Font(size=14, bold=True)
    lines = [
        ("Where to write", f"Use the next empty row in {sheet.title}. Add one bug, request or question per row."),
        (columns["id"], "Filled in automatically. Quote it when you refer to a row; Excel refuses anything typed there."),
        ("Date, Tester, Type", "Enter the date and choose your name and the type from the dropdowns."),
        ("Reference", "The record the feedback is about, such as an order or booking number, when there is one."),
        (columns["feedback"], "For a bug, say what you did, what happened and what you expected. For a request, say what you would change."),
        ("Links", "A sharing link to screenshots or further information. Make sure the reviewing team can open it."),
        (columns["status"], f"{statuses['open']} (red): not yet being worked on. {statuses['in_progress']} (amber): work is under way. {statuses['closed']} (green): done or no further action, and the response says why."),
        ("Reviewing team", f"{columns['status']}, {columns['response']}, {columns['priority']}, {columns['owner']} and {columns['issue']} are filled in by the reviewing team. Testers can leave them blank."),
        (columns["priority"], " ".join(f"{level}: {meaning}." for level, meaning in config["priorities"].items()) + " A request or question takes the priority of the problem it solves."),
        ("Names", "Add testers on the Names tab, directly below the list without gaps."),
    ]
    if legacy:
        lines.insert(5, (legacy_header, "For a bug, try the same thing in the old system and choose Yes or No. Choose Not applicable if it has no equivalent, or for a request or question."))
    for number, (label, text) in enumerate(lines, start=3):
        guide.cell(number, 1, label).font = Font(bold=True)
        guide.cell(number, 2, text).alignment = Alignment(wrap_text=True, vertical="top")
    guide.column_dimensions["A"].width = 22
    guide.column_dimensions["B"].width = 100

    names = book.create_sheet("Names")
    names["A1"] = "Tester names"
    names["A2"] = "Add each tester below, without blank rows; the Tester dropdown reads this list."
    names["A3"], names["C3"], names["D3"] = "Tester", "Owner code", "Owner"
    for number, (code, owner) in enumerate(config["owners"].items(), start=4):
        names.cell(number, 3, code)
        names.cell(number, 4, owner["name"])
    # The dropdown grows with the list: from A4 down as many cells as hold a name, never fewer than one
    book.defined_names["Testers"] = DefinedName("Testers", attr_text="OFFSET(Names!$A$4,0,0,MAX(1,COUNTA(Names!$A$4:$A$1000)),1)")
    book.save(path)
    print(f"workbook: {path}")


def main() -> None:
    """Parse the command line and run one command."""
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    commands = parser.add_subparsers(dest="command", required=True)
    for name in ("new", "rows", "ids", "write"):
        command = commands.add_parser(name)
        command.add_argument("config", type=Path)
        command.add_argument("workbook", type=Path)
        if name == "new":
            command.add_argument("--legacy", action="store_true", help="add a column comparing each bug with the system being replaced")
            continue
        if name == "rows":
            command.add_argument("--open", action="store_true", help="leave out rows whose status is the closed word")
        else:
            command.add_argument("--copy", action="store_true", help="write to a copy in output_dir")
    write = commands.choices["write"]
    write.add_argument("--id", action="append", required=True, help="a row's UID as the sheet shows it; repeat for several rows")
    write.add_argument("--status")
    write.add_argument("--priority", help="one of the config's priorities")
    write.add_argument("--owner", help="one of the config's owner codes")
    write.add_argument("--response")
    write.add_argument("--by", help="the writer's initials, signed on the response")
    write.add_argument("--issue", help="the tracker issue's address")
    args = parser.parse_args()

    config = json.loads(args.config.read_text(encoding="utf-8"))
    if args.command == "new":
        new_workbook(config, args.workbook, args.legacy)
        return
    if args.command == "rows":
        sheet = openpyxl.load_workbook(args.workbook, data_only=True)[config["sheet"]["tab"]]
        _columns, rows = read_rows(sheet, config)
        for row in rows:
            # A blank status counts as open, so only the closed word is left out
            if args.open and row["status"] == config["statuses"]["closed"]:
                continue
            print(json.dumps(row, default=str, ensure_ascii=False))
        return
    guarded_write(config, args.workbook, args)


if __name__ == "__main__":
    main()
