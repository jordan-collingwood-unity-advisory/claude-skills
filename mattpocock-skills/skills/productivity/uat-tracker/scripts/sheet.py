"""Read and write a UAT feedback sheet described by a project's uat-tracker.json.

Run from the project root:
  rows  <config> <workbook>            print every row with feedback as one JSON line
  ids   <config> <workbook> [--copy]   give each row with feedback and no ID the next ID
  write <config> <workbook> --id N [--id M ...] [--status S] [--response TEXT --by INITIALS] [--issue URL] [--copy]
A write backs the workbook up first and prints every cell it changed; --copy writes to a copy instead.
"""

import argparse
import json
import shutil
from datetime import date, datetime
from pathlib import Path

import openpyxl
from openpyxl.worksheet.worksheet import Worksheet


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


def assign_ids(sheet: Worksheet, config: dict) -> list[str]:
    """Give each row with no ID the next number after the largest ID in the column; refuse duplicate IDs."""
    columns, rows = read_rows(sheet, config)
    existing = [int(row["id"]) for row in rows if row["id"] not in (None, "")]
    duplicates = sorted({row_id for row_id in existing if existing.count(row_id) > 1})
    if duplicates:
        raise SystemExit(f"Duplicate IDs {duplicates}; fix them in the sheet first. Nothing saved.")
    next_id = max(existing, default=0) + 1
    messages = []
    for row in rows:
        if row["id"] in (None, ""):
            sheet.cell(row["row"], columns["id"]).value = next_id
            messages.append(f"sheet row {row['row']}: ID {next_id}")
            next_id += 1
    return messages


def find_row(rows: list[dict], row_id: int) -> dict:
    """Return the one row carrying this ID."""
    matches = [row for row in rows if row["id"] not in (None, "") and int(row["id"]) == row_id]
    if len(matches) != 1:
        raise SystemExit(f"ID {row_id} matches {len(matches)} rows, not one. Nothing saved.")
    return matches[0]


def write_rows(sheet: Worksheet, config: dict, args: argparse.Namespace) -> list[str]:
    """Set status, add a dated response line on top, and link an issue, on each row named by ID."""
    if not (args.status or args.response or args.issue):
        raise SystemExit("Give --status, --response or --issue.")
    if args.status and args.status not in config["statuses"].values():
        raise SystemExit(f"Status {args.status!r} is not one of {list(config['statuses'].values())}.")
    if args.response and not args.by:
        raise SystemExit("A response needs --by with the writer's initials.")
    columns, rows = read_rows(sheet, config)
    today = date.today()
    messages = []
    for row_id in args.id:
        number = find_row(rows, row_id)["row"]
        if args.status:
            sheet.cell(number, columns["status"]).value = args.status
        if args.response:
            cell = sheet.cell(number, columns["response"])
            line = f"- {today.day} {today:%b}, {args.by} - {args.response}"
            # Earlier entries stay below the new line, read from the cell as it is now
            earlier = str(cell.value).strip() if cell.value is not None else ""
            cell.value = f"{line}\n{earlier}" if earlier else line
        if args.issue:
            cell = sheet.cell(number, columns["issue"])
            if cell.value not in (None, "", args.issue):
                raise SystemExit(f"ID {row_id} already links {cell.value!r}. Nothing saved.")
            cell.value = args.issue
        messages.append(f"ID {row_id} (sheet row {number}) written")
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


def main() -> None:
    """Parse the command line and run one command."""
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    commands = parser.add_subparsers(dest="command", required=True)
    for name in ("rows", "ids", "write"):
        command = commands.add_parser(name)
        command.add_argument("config", type=Path)
        command.add_argument("workbook", type=Path)
        if name != "rows":
            command.add_argument("--copy", action="store_true", help="write to a copy in output_dir")
    write = commands.choices["write"]
    write.add_argument("--id", type=int, action="append", required=True)
    write.add_argument("--status")
    write.add_argument("--response")
    write.add_argument("--by", help="the writer's initials, signed on the response")
    write.add_argument("--issue", help="the tracker issue's address")
    args = parser.parse_args()

    config = json.loads(args.config.read_text(encoding="utf-8"))
    if args.command == "rows":
        sheet = openpyxl.load_workbook(args.workbook, data_only=True)[config["sheet"]["tab"]]
        _columns, rows = read_rows(sheet, config)
        for row in rows:
            print(json.dumps(row, default=str, ensure_ascii=False))
        return
    guarded_write(config, args.workbook, args)


if __name__ == "__main__":
    main()
