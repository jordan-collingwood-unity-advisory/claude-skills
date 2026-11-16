"""Read and write a UAT feedback sheet described by a project's uat-tracker.json.

Run from the project root:
  rows  <config> <workbook> [--open]   print every row with feedback as one JSON line; --open leaves out closed rows
  ids   <config> <workbook> [--copy]   give each row with feedback and no UID the prefix and the next number
  write <config> <workbook> --id UID [--id UID ...] [--status S] [--response TEXT --by INITIALS] [--issue URL] [--copy]
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
    for uid in args.id:
        number = find_row(rows, uid)["row"]
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


def main() -> None:
    """Parse the command line and run one command."""
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    commands = parser.add_subparsers(dest="command", required=True)
    for name in ("rows", "ids", "write"):
        command = commands.add_parser(name)
        command.add_argument("config", type=Path)
        command.add_argument("workbook", type=Path)
        if name == "rows":
            command.add_argument("--open", action="store_true", help="leave out rows whose status is the closed word")
        else:
            command.add_argument("--copy", action="store_true", help="write to a copy in output_dir")
    write = commands.choices["write"]
    write.add_argument("--id", action="append", required=True, help="a row's UID as the sheet shows it; repeat for several rows")
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
            # A blank status counts as open, so only the closed word is left out
            if args.open and row["status"] == config["statuses"]["closed"]:
                continue
            print(json.dumps(row, default=str, ensure_ascii=False))
        return
    guarded_write(config, args.workbook, args)


if __name__ == "__main__":
    main()
