"""
CSV to Excel Desktop Converter
- Local Tkinter GUI; no browser or network access.
- Uses pandas for CSV parsing and openpyxl for .xlsx creation.
- Imports every field as text to preserve values such as leading zeros,
  long identifiers, date-like strings, and formula-like text.
"""

from __future__ import annotations

import csv
import sys
import traceback
from pathlib import Path
from typing import Iterable

import pandas as pd
from charset_normalizer import from_bytes
from openpyxl import Workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from tkinter import Tk, StringVar, filedialog, messagebox, Listbox, END, MULTIPLE
from tkinter import ttk

EXCEL_MAX_ROWS = 1_048_576
EXCEL_MAX_COLUMNS = 16_384
EXCEL_MAX_CELL_LENGTH = 32_767


def detect_encoding(path: Path) -> tuple[str, bytes]:
    """Detect a file encoding, honoring common Unicode BOMs first."""
    raw = path.read_bytes()
    if raw.startswith(b"\xef\xbb\xbf"):
        return "utf-8-sig", raw
    if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        return "utf-16", raw

    match = from_bytes(raw).best()
    if match is None:
        raise UnicodeError("Could not determine the file encoding.")
    return match.encoding or "utf-8", raw


def detect_delimiter(sample: str) -> str:
    """Detect comma, semicolon, or tab separators, with a safe fallback."""
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
        if dialect.delimiter in {",", ";", "\t"}:
            return dialect.delimiter
    except csv.Error:
        pass

    # Fallback: choose the supported delimiter occurring most often
    # outside quoted fields. Sniffer handles the normal case.
    counts = {d: sample.count(d) for d in (",", ";", "\t")}
    delimiter = max(counts, key=counts.get)
    if counts[delimiter] == 0:
        return ","
    return delimiter


def read_csv_as_text(path: Path) -> pd.DataFrame:
    """Parse a CSV into a DataFrame without converting values to other types."""
    encoding, raw = detect_encoding(path)
    sample = raw[:100_000].decode(encoding, errors="replace")
    delimiter = detect_delimiter(sample)

    # header=None prevents pandas from renaming duplicate column headers.
    # Every record is retained as a row; no header/index is added in Excel.
    return pd.read_csv(
        path,
        sep=delimiter,
        header=None,
        dtype=str,
        keep_default_na=False,
        na_filter=False,
        skip_blank_lines=False,
        encoding=encoding,
        engine="python",
        on_bad_lines="error",
    )


def safe_sheet_name(name: str, used: set[str]) -> str:
    """Create a valid, unique Excel worksheet name."""
    invalid = set(r'[]:*?/\\')
    cleaned = "".join("_" if ch in invalid else ch for ch in name).strip("'")
    cleaned = cleaned or "Sheet"
    cleaned = cleaned[:31]
    candidate = cleaned
    counter = 2
    while candidate.casefold() in used:
        suffix = f" ({counter})"
        candidate = cleaned[: 31 - len(suffix)] + suffix
        counter += 1
    used.add(candidate.casefold())
    return candidate


def write_workbook(csv_paths: Iterable[Path], output_path: Path) -> int:
    """Write each CSV to its own worksheet and return the number of sheets."""
    paths = list(csv_paths)
    if not paths:
        raise ValueError("Select at least one CSV file.")

    workbook = Workbook()
    workbook.remove(workbook.active)
    used_names: set[str] = set()

    for path in paths:
        frame = read_csv_as_text(path)
        rows, columns = frame.shape
        if rows > EXCEL_MAX_ROWS or columns > EXCEL_MAX_COLUMNS:
            raise ValueError(
                f"{path.name} exceeds Excel's worksheet limit "
                f"({rows:,} rows × {columns:,} columns)."
            )

        sheet = workbook.create_sheet(safe_sheet_name(path.stem, used_names))
        for row_index, row in enumerate(frame.itertuples(index=False, name=None), start=1):
            for column_index, value in enumerate(row, start=1):
                text_value = "" if value is None else str(value)
                if len(text_value) > EXCEL_MAX_CELL_LENGTH:
                    raise ValueError(
                        f"{path.name}: cell at row {row_index}, column "
                        f"{column_index} exceeds Excel's 32,767-character limit."
                    )
                if ILLEGAL_CHARACTERS_RE.search(text_value):
                    raise ValueError(
                        f"{path.name}: cell at row {row_index}, column "
                        f"{column_index} contains a character Excel cannot store."
                    )
                cell = sheet.cell(row=row_index, column=column_index)
                cell.value = text_value
                # Force a literal string, including values beginning with "=".
                cell.data_type = "s"

        # Make the worksheet easier to inspect without changing cell contents.
        sheet.freeze_panes = "A1"
        sheet.sheet_view.showGridLines = True

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    return len(paths)


class ConverterApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title("CSV to Excel Converter")
        self.root.geometry("650x420")
        self.root.minsize(540, 340)
        self.files: list[Path] = []
        self.status = StringVar(value="Select one or more CSV files to begin.")

        self._build_ui()

    def _build_ui(self) -> None:
        outer = ttk.Frame(self.root, padding=18)
        outer.pack(fill="both", expand=True)

        ttk.Label(
            outer, text="CSV to Excel Converter", font=("TkDefaultFont", 16, "bold")
        ).pack(anchor="w")
        ttk.Label(
            outer,
            text="Convert CSV files into one .xlsx workbook. Each CSV becomes a worksheet.",
            wraplength=590,
        ).pack(anchor="w", pady=(5, 14))

        actions = ttk.Frame(outer)
        actions.pack(fill="x", pady=(0, 8))
        ttk.Button(actions, text="Select CSV", command=self.select_files).pack(side="left")
        ttk.Button(actions, text="Remove selected", command=self.remove_selected).pack(
            side="left", padx=8
        )
        ttk.Button(actions, text="Clear", command=self.clear_files).pack(side="left")

        list_frame = ttk.Frame(outer)
        list_frame.pack(fill="both", expand=True)
        self.file_list = Listbox(list_frame, selectmode=MULTIPLE, height=10)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.file_list.yview)
        self.file_list.configure(yscrollcommand=scrollbar.set)
        self.file_list.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        ttk.Label(outer, textvariable=self.status, wraplength=590).pack(
            anchor="w", pady=(12, 8)
        )
        bottom = ttk.Frame(outer)
        bottom.pack(fill="x")
        ttk.Button(bottom, text="Convert", command=self.convert, style="Accent.TButton").pack(
            side="right"
        )

        style = ttk.Style()
        try:
            style.configure("Accent.TButton", padding=(18, 8))
        except Exception:
            pass

    def select_files(self) -> None:
        selected = filedialog.askopenfilenames(
            title="Select CSV files",
            filetypes=[("CSV and text files", "*.csv *.txt"), ("All files", "*.*")],
        )
        for item in selected:
            path = Path(item)
            if path not in self.files:
                self.files.append(path)
                self.file_list.insert(END, str(path))
        if self.files:
            self.status.set(f"{len(self.files)} file(s) ready to convert.")

    def remove_selected(self) -> None:
        indices = list(self.file_list.curselection())
        for index in reversed(indices):
            self.file_list.delete(index)
            del self.files[index]
        self.status.set(f"{len(self.files)} file(s) selected.")

    def clear_files(self) -> None:
        self.files.clear()
        self.file_list.delete(0, END)
        self.status.set("Select one or more CSV files to begin.")

    def convert(self) -> None:
        if not self.files:
            messagebox.showwarning("No files selected", "Select at least one CSV file.")
            return

        default_name = (
            f"{self.files[0].stem}.xlsx" if len(self.files) == 1 else "converted_workbook.xlsx"
        )
        output = filedialog.asksaveasfilename(
            title="Save Excel workbook",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if not output:
            return

        try:
            self.status.set("Converting…")
            self.root.update_idletasks()
            count = write_workbook(self.files, Path(output))
            self.status.set(f"Done. Created {output}")
            messagebox.showinfo(
                "Conversion complete",
                f"Successfully converted {len(self.files)} CSV file(s) "
                f"into {count} worksheet(s).\\n\\nSaved to:\\n{output}",
            )
        except Exception as exc:
            self.status.set("Conversion failed. See the error message for details.")
            messagebox.showerror(
                "Conversion failed",
                f"{exc}\\n\\nNo successful workbook was reported. "
                "Check the file and try again.",
            )


def main() -> None:
    root = Tk()
    try:
        ttk.Style().theme_use("clam")
    except Exception:
        pass
    ConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
