# Project Documentation: Excel/CSV to Notepad Converter

> **Full analysis of the project — libraries, functions, architecture, commands, and workflow**

---

## Table of Contents

1. [Project Overview](#1-project-overview)
2. [System Architecture](#2-system-architecture)
3. [Directory Structure](#3-directory-structure)
4. [Dependencies & Libraries](#4-dependencies--libraries)
5. [Function Reference](#5-function-reference)
6. [Execution Flow](#6-execution-flow)
7. [Mode Comparison: Layout vs Table](#7-mode-comparison-layout-vs-table)
8. [CLI Commands](#8-cli-commands)
9. [Complete Usage Examples](#9-complete-usage-examples)
10. [Output Format Specifications](#10-output-format-specifications)
11. [Edge Cases Handled](#11-edge-cases-handled)

---

## 1. Project Overview

### What It Does

A **robust Python script** that reads **Excel (.xlsx/.xlsm/.xltx/.xltm/.xls) workbooks** or **CSV files** and converts each sheet into **well-formatted plain-text (.txt) files** and **high-fidelity PDF (.pdf) files**. The output is optimized for viewing in **Notepad** or any **monospace text editor**.

### Key Capabilities

| Capability | Description |
|---|---|
| **Automatic table detection** | Finds tabular data even in sheets with titles, notes, empty rows, or mixed content |
| **Two export modes** | `layout` (preserves Excel visual grid) and `table` (extracts structured data blocks) |
| **Fixed-width formatting** | Aligned columns using spaces — perfect for Notepad |
| **Smart text wrapping** | Long cell values wrap within configurable column width limits |
| **Merged cell support** | Respects merged cells in layout mode; marks covered cells as empty |
| **Duplicate-safe output** | Avoids overwriting existing files with numbered suffixes (`--no-overwrite`) |
| **PDF export** | Generates landscape A4 PDFs with cell borders, Unicode support, and auto-scaling |
| **CSV support** | CSV files processed identically with a single TXT/PDF output |

---

## 2. System Architecture

```
┌─────────────────────────────────────────────────────────────────────────┐
│                           CLI Entry Point                               │
│                      python main.py [input] [options]                   │
└────────────────────────────┬────────────────────────────────────────────┘
                             │
                             ▼
┌─────────────────────────────────────────────────────────────────────────┐
│                         main() — Argparse Parser                        │
│                                                                         │
│  Positional: input                                                      │
│  Optional:   --output, --mode, --max-column-width,                      │
│              --no-overwrite, --verbose                                  │
└────────────────────────────┬────────────────────────────────────────────┘
                             │
                             ▼
┌─────────────────────────────────────────────────────────────────────────┐
│                     process_input() — File Dispatcher                   │
│                                                                         │
│  Validates: file exists, supported extension                            │
│  Checks:    .csv → process_csv()                                        │
│             .xlsx/.xlsm/etc → process_excel()                           │
└────────────────────────────┬────────────────────────────────────────────┘
                             │
              ┌──────────────┴──────────────┐
              │                              │
              ▼                              ▼
┌──────────────────────────┐  ┌──────────────────────────────────────────┐
│      process_csv()       │  │          process_excel()                 │
│                          │  │                                          │
│  1. Read CSV via pandas  │  │  1. Load workbook with openpyxl          │
│  2. Format with          │  │  2. For each sheet:                      │
│     format_tables()      │  │     if mode=="table":                    │
│  3. Write .txt & .pdf    │  │       pandas read + format_tables()      │
│                          │  │     if mode=="layout":                   │
│                          │  │       openpyxl raw + format_layout_sheet()│
│                          │  │     Export PDF via export_to_pdf()       │
│                          │  │  3. Write .txt & .pdf per sheet          │
└──────────────────────────┘  └──────────────────────────────────────────┘
```

### Data Flow Diagram (Table Mode)

```
Excel/CSV File
     │
     ▼
┌─────────────┐     ┌──────────────────┐     ┌───────────────────┐
│ pd.read_excel│────▶│ get_non_empty_    │────▶│ extract_table_from_│
│ (header=None)│     │ blocks()          │     │ block()            │
└─────────────┘     └──────────────────┘     └───────────────────┘
                                                        │
                                                        ▼
                                              ┌────────────────────┐
                                              │ format_fixed_width()│
                                              │   (wrapping +       │
                                              │    alignment)       │
                                              └─────────┬──────────┘
                                                        │
                                                        ▼
                                              ┌────────────────────┐
                                              │ write_text_file()  │
                                              │     → .txt file    │
                                              └────────────────────┘
```

### Data Flow Diagram (Layout Mode)

```
Excel File (.xlsx .xlsm .xltx .xltm)
     │
     ▼
┌──────────────────┐     ┌───────────────────┐     ┌──────────────────────┐
│ load_workbook()   │────▶│ get_covered_      │────▶│ get_used_bounds()    │
│ (openpyxl)        │     │ merged_cells()    │     │ (visible data bbox)  │
└──────────────────┘     └───────────────────┘     └──────────┬───────────┘
                                                              │
                                                              ▼
                                                  ┌──────────────────────┐
                                                  │ format_layout_sheet()│
                                                  │  • read each cell    │
                                                  │  • use Excel widths  │
                                                  │  • render_wrapped_row│
                                                  └──────────┬───────────┘
                                                             │
                                                             ▼
                                                  ┌──────────────────────┐
                                                  │ write_text_file()    │
                                                  │      → .txt file     │
                                                  └──────────────────────┘
```

### PDF Export Flow

```
┌──────────────┐    ┌─────────────────┐    ┌──────────────────────┐
│ For each      │───▶│ For each row in  │───▶│ multi_cell() with    │
│ worksheet     │    │ bounding box     │    │ borders + alignment  │
│ (openpyxl)    │    │                  │    │ + line height calc   │
└──────────────┘    └─────────────────┘    └──────────┬───────────┘
                                                       │
                                                       ▼
                                             ┌──────────────────────┐
                                             │ setup_pdf_font()     │
                                             │  • Try Arial/Liberation│
                                             │  • Fallback helvetica│
                                             └──────────────────────┘
```

---

## 3. Directory Structure

```
excel to notepad/
│
├── main.py                          # Main application script (925 lines)
│
├── README.md                        # User-facing documentation
│
├── requirement.md                   # Original project requirements specification
│
├── requirements.txt                 # Python package dependencies
│
├── PROJECT_DOCUMENTATION.md         # This file
│
└── data/
    │
    ├── Input/                       # ⬅ Place your source files here
    │   ├── Requirements Traceability Matrix (RTM) - Workforce Development.xlsx
    │   ├── Requirements_HighLevel.xlsx
    │   └── RTM Sheet ACTO.xlsx
    │
    └── Output/                      # ⬅ Generated output appears here
        │
        ├── Requirements Traceability Matrix (RTM) - Workforce Development/
        │   ├── TXT/                    # Text output files (8 sheets)
        │   │   ├── Acceptance Criteria.txt
        │   │   ├── Document Version Control.txt
        │   │   ├── Dropdown List(s).txt
        │   │   ├── Instructions.txt
        │   │   ├── OFS Maximo Use Case Scenarios.txt
        │   │   ├── Project Overview.txt
        │   │   ├── Requirements (obsoleteFeb06).txt
        │   │   └── Requirements.txt
        │   └── PDF/                    # PDF output files (same number)
        │
        ├── Requirements_HighLevel/
        │   ├── TXT/
        │   │   ├── Business Requirements.txt
        │   │   ├── Legends and Lookups.txt
        │   │   ├── Non-Functional Requirements.txt
        │   │   ├── Project Overview.txt
        │   │   ├── Technical Requirements NM.txt
        │   │   ├── Technical Requirements.txt
        │   │   └── Use Cases.txt
        │   └── PDF/
        │
        └── RTM Sheet ACTO/
            ├── TXT/
            │   └── Sheet1.txt
            └── PDF/
```

### Output Naming Convention

```
data/Output/
  └── [sanitized_source_filename]/         # Folder named after input file
        ├── TXT/
        │   └── [sanitized_sheet_name].txt # One TXT per Excel sheet
        └── PDF/
             └── [sanitized_sheet_name].pdf # One PDF per Excel sheet
```

- **Sanitization**: Characters `\ / : * ? " < > |` are replaced with `_`
- **Duplicate handling** (with `--no-overwrite`): `Sheet.txt` → `Sheet_1.txt` → `Sheet_2.txt` ...

---

## 4. Dependencies & Libraries

### External Packages

| Package | Version | Purpose |
|---|---|---|
| [**pandas**](https://pandas.pydata.org/) | ≥ 2.0.0 | Primary data manipulation. Reads Excel/CSV into DataFrames, handles NaN detection, provides `.dropna()`, `.iterrows()`, `.iloc[]`, `.map()`. |
| [**openpyxl**](https://openpyxl.readthedocs.io/) | ≥ 3.1.0 | Advanced Excel reading. Accesses raw worksheet objects, merged cells, column dimensions, cell values. Required for `layout` mode. |
| [**xlrd**](https://xlrd.readthedocs.io/) | ≥ 2.0.1 | Support for legacy `.xls` (Excel 97-2003) format files via pandas engine. |
| [**fpdf2**](https://pyfpdf.github.io/fpdf2/) | ≥ 2.7.4 | PDF generation. Creates landscape A4 PDFs with `multi_cell()` (text with borders), `cell()`, `rect()`, font management. |

### Python Standard Library Modules

| Module | Purpose in this Project |
|---|---|
| **`argparse`** | Parses CLI arguments: `input`, `-o`, `--mode`, `--max-column-width`, `--no-overwrite`, `-v` |
| **`logging`** | Console-based logging (`INFO` level by default, `DEBUG` with `-v`). Format: `LEVEL: message` |
| **`re`** | Regex for filename sanitization (`[\\/*?:"<>|]`) and collapsing underscores |
| **`textwrap`** | `textwrap.wrap()` — splits long text lines at word boundaries to fit column widths |
| **`pathlib.Path`** | Cross-platform file path operations: `exists()`, `stem`, `suffix`, `mkdir()`, `read_text()`/`write_text()` |
| **`typing`** | Type hints: `Dict`, `List`, `Tuple` |
| **`dataclasses`/`os`** | (Not used — pure pathlib) |

---

## 5. Function Reference

### 1. `sanitize_filename(filename, max_length=200) → str`

**Purpose**: Convert any string into a valid filename for Windows/Linux/macOS.

**Logic**:
1. Replace invalid chars (`\ / : * ? " < > |`) with `_`
2. Collapse multiple `_` into one
3. Strip leading/trailing `_`
4. Truncate to `max_length`
5. Return `"unnamed_sheet"` if result is empty

**Called by**: `process_csv()`, `process_excel()`, `get_output_file()`

---

### 2. `clean_cell_value(value, preserve_newlines=False) → str`

**Purpose**: Convert any cell value to a clean string representation.

**Logic**:
1. `pd.isna(value)` → return `""`
2. Datetime objects (`value.strftime` exists) → format as `"YYYY-MM-DD HH:MM:SS"`
3. Convert to string, strip whitespace
4. Normalize `\r\n` → `\n`, `\r` → `\n`
5. Replace tabs with spaces
6. If `preserve_newlines=False`, collapse all whitespace (including newlines) to single spaces

**Called by**: `wrap_text()`, `format_fixed_width()`, `get_layout_cell_value()`, `extract_table_from_block()`

---

### 3. `is_non_empty(value) → bool`

**Purpose**: Check if a cell has visible content.

**Logic**: Returns `True` when `pd.notna(value) AND str(value).strip() != ""`

**Called by**: `is_meaningful_row()`, `detect_header_row()`

---

### 4. `is_meaningful_row(row, threshold=1) → bool`

**Purpose**: Check if a row contains enough data to be considered meaningful.

**Logic**: Counts non-empty cells → returns `True` if count ≥ `threshold`

**Called by**: `get_non_empty_blocks()`, `extract_table_from_block()`

---

### 5. `make_unique_headers(headers) → List[str]`

**Purpose**: Convert a list of header names into unique, non-empty column names.

**Logic**:
1. Empty/None names → `"Column_{n}"`
2. Duplicate names → `"Name_2"`, `"Name_3"`, ...

**Called by**: `extract_table_from_block()`, `format_fixed_width()`

---

### 6. `get_non_empty_blocks(df) → List[pd.DataFrame]`

**Purpose**: Split a messy DataFrame into clean blocks separated by fully empty rows.

**Flow**:
1. Drop all-empty columns
2. Scan row by row
3. If row has meaningful data → add to current block
4. If row is empty AND we have accumulated rows → save block, reset
5. Return all non-empty blocks

**Called by**: `extract_tabular_blocks()`

---

### 7. `detect_header_row(block) → int`

**Purpose**: Find the most likely header row (the one with the most non-empty cells in the top 8 rows).

**Logic**:
1. Scan first 8 rows (or fewer if block is smaller)
2. Count non-empty cells per row
3. Return the position of the row with the highest count

**Called by**: `extract_table_from_block()`

---

### 8. `extract_table_from_block(block) → pd.DataFrame`

**Purpose**: Convert one non-empty block into a structured, clean DataFrame with proper headers.

**Flow**:
1. Detect header row
2. Extract from header position onward
3. Clean row 0 values → use as column names via `make_unique_headers()`
4. Remove the header row from data
5. Drop empty rows and columns
6. Handle single-row edge case

**Called by**: `extract_tabular_blocks()`

---

### 9. `extract_tabular_blocks(df) → List[pd.DataFrame]`

**Purpose**: Extract ALL tables from a messy sheet.

**Flow**: For each block from `get_non_empty_blocks()` → `extract_table_from_block()` → collect non-empty results

**Called by**: `format_tables()`

---

### 10. `wrap_text(value, width) → List[str]`

**Purpose**: Wrap cell text to fit within a specified column width.

**Logic**:
1. Clean with `preserve_newlines=True` (keeps \n)
2. Split by `\n` (handle multi-paragraph)
3. Use `textwrap.wrap()` for each paragraph
4. Return list of wrapped lines

**Called by**: `render_wrapped_row()`

---

### 11. `render_wrapped_row(values, widths, separator="  ") → List[str]`

**Purpose**: Render one logical table row as potentially multiple physical text lines.

**Flow**:
1. Wrap each cell value to its column width
2. Determine `row_height` = max lines needed
3. For each line index, assemble text segments from each cell
4. Left-align each segment within its column width

**Called by**: `format_fixed_width()`, `format_layout_sheet()`

---

### 12. `format_fixed_width(df, max_column_width=60) → str`

**Purpose**: Convert a clean DataFrame into a fixed-width aligned text table.

**Output format**:
```
Header1    Header2    Header3
--------   --------   --------
value1     value2     value3
long text  more data  wrapped
here...               properly
```

**Logic**:
1. Clean all cell values (preserving newlines)
2. Ensure unique headers
3. Calculate column widths (clamped between 2 and `max_column_width`)
4. Render header row
5. Render dashes separator
6. Render each data row (with wrapping)

**Called by**: `format_tables()`

---

### 13. `format_tables(df_raw, sheet_name, max_column_width) → str`

**Purpose**: Orchestrate table extraction and formatting for a raw sheet.

**Logic**:
1. Call `extract_tabular_blocks()`
2. If no tables → fallback: raw export with `df.to_string(index=False)`
3. For each table → `format_fixed_width()`
4. Prepend `"BLOCK N"` titles when multiple tables exist

**Called by**: `process_csv()`, `process_excel()`

---

### 14. `get_covered_merged_cells(worksheet) → Dict[Tuple[int,int], bool]`

**Purpose**: Identify cells that are visually hidden by merged cell ranges.

**Logic**: For each merged range, mark all cells EXCEPT the top-left cell as covered.

**Called by**: `format_layout_sheet()`, `export_to_pdf()`

---

### 15. `get_layout_cell_value(worksheet, row, col, covered_cells) → str`

**Purpose**: Return visible text for one cell, respecting merged cells.

**Logic**: If cell is covered → return `""`; otherwise → `clean_cell_value(cell.value, preserve_newlines=True)`

**Called by**: `format_layout_sheet()`, `export_to_pdf()`, `get_used_bounds()`

---

### 16. `get_used_bounds(worksheet, covered_cells) → Tuple | None`

**Purpose**: Find the bounding box (min/max row/col) of all visible data.

**Logic**: Scan entire worksheet → record positions of cells with non-empty content → return `(min_row, max_row, min_col, max_col)` or `None` if empty.

**Called by**: `format_layout_sheet()`, `export_to_pdf()`

---

### 17. `excel_column_width_to_chars(width) → int`

**Purpose**: Convert Excel's column width units to character counts.

**Logic**: `None` → 12; otherwise `max(MIN_COLUMN_WIDTH, int(round(float(width))))`

**Called by**: `format_layout_sheet()`

---

### 18. `format_layout_sheet(worksheet, max_column_width) → str`

**Purpose**: Approximate the Excel grid as text, preserving column positions and widths.

**Flow**:
1. Identify merged cells
2. Find bounding box of visible data
3. Calculate each column's width (max of Excel width and content width)
4. Iterate through each row → `render_wrapped_row()` → collect lines

**Called by**: `process_excel()`

---

### 19. `get_output_file(directory, base_name, extension, overwrite) → Path`

**Purpose**: Generate an output file path, avoiding existing files when needed.

**Logic**:
- `overwrite=True` → return `dir / "name.ext"`
- `overwrite=False` → while file exists, append `_1`, `_2`, etc.

**Called by**: `process_csv()`, `process_excel()`

---

### 20. `setup_pdf_font(pdf)`

**Purpose**: Load a Unicode-capable font for PDF (supports special characters).

**Tries** (in order):
1. `C:\Windows\Fonts\arial.ttf` (Windows)
2. `/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf` (Linux)
3. `/System/Library/Fonts/Supplemental/Arial.ttf` (macOS)
4. Falls back to built-in `helvetica` (Latin-1 only)

**Called by**: `export_to_pdf()`

---

### 21. `export_to_pdf(worksheet, output_path, title)`

**Purpose**: Generate a high-fidelity PDF with cell borders and proper alignment.

**Details**:
- Landscape A4, 15mm bottom margin
- Title centered at top
- Column widths from Excel (scaled ×2, min 10mm)
- Auto-scales if total width exceeds page width
- Auto page breaks
- Each cell rendered with `multi_cell()` (borders, left-aligned)
- Merged cells → empty rectangles
- Font: Unicode (Arial/Liberation) or helvetica fallback

---

### 22. `write_text_file(path, content)`

**Purpose**: Write UTF-8 encoded text to disk.

**Implementation**: `path.write_text(content, encoding="utf-8")`

---

### 23. `process_csv(input_path, output_folder, max_column_width, overwrite) → bool`

**Purpose**: Full pipeline for CSV files.

**Flow**:
1. Create `[file_root]/TXT/` and `[file_root]/PDF/` directories
2. Read CSV via `pd.read_csv(header=None)`
3. Format via `format_tables()`
4. Write `.txt` file
5. Generate PDF (simple line-by-line text rendering)
6. Return success/failure

---

### 24. `process_excel(input_path, output_folder, mode, max_column_width, overwrite) → bool`

**Purpose**: Full pipeline for Excel files.

**Flow**:
1. Create `[file_root]/TXT/` and `[file_root]/PDF/` directories
2. Load workbook with `openpyxl.load_workbook(data_only=True)`
3. For each sheet:
   - `mode == "layout"` → `format_layout_sheet(worksheet, ...)`
   - `mode == "table"` → `pd.read_excel()` + `format_tables()`
   - Write `.txt` file
   - Generate PDF via `export_to_pdf()`
4. Log processing summary (processed/skipped counts)
5. Return `True` if at least one sheet processed

---

### 25. `process_input(input_file, output_folder, mode, max_column_width, overwrite) → bool`

**Purpose**: Validate input and dispatch to CSV or Excel processor.

**Checks**:
- File exists
- Extension is in `SUPPORTED_INPUT_EXTENSIONS` (`.xlsx`, `.xlsm`, `.xltx`, `.xltm`, `.xls`, `.csv`)
- Creates output directory if needed

---

### 26. `main()`

**Purpose**: CLI entry point. Parse arguments, configure logging, call `process_input()`.

**Returns**: Exit code 0 (success) or 1 (failure).

---

## 6. Execution Flow

### Step-by-step for `python main.py data/Input/report.xlsx`

```
1. main()
   │
   ├── argparse parses: input="data/Input/report.xlsx",
   │                   output="data/Output" (default),
   │                   mode="layout" (default),
   │                   max_column_width=60 (default),
   │                   overwrite=True, verbose=False
   │
   └── process_input("data/Input/report.xlsx", "data/Output", "layout", 60, True)
        │
        ├── Check: input_path.exists() → True
        ├── Check: .xlsx → SUPPORTED_INPUT_EXTENSIONS → True
        ├── Create: data/Output (if not exists)
        └── → process_excel("data/Input/report.xlsx", "data/Output", "layout", 60, True)
              │
              ├── mkdir: data/Output/report/TXT/
              ├── mkdir: data/Output/report/PDF/
              ├── openpyxl.load_workbook("data/Input/report.xlsx", data_only=True)
              │
              ├── For each sheet_name in workbook.sheetnames:
              │    │
              │    ├── LOG: "Processing sheet: 'Sheet1'..."
              │    │
              │    ├── mode == "layout":
              │    │   └── content = format_layout_sheet(worksheet, 60)
              │    │        ├── covered = get_covered_merged_cells(worksheet)
              │    │        ├── bounds = get_used_bounds(worksheet, covered)
              │    │        └── For each row → render_wrapped_row(values, widths)
              │    │
              │    ├── write_text_file(txt_path, content)
              │    │
              │    └── export_to_pdf(worksheet, pdf_path, sheet_name)
              │         ├── setup_pdf_font(pdf) → try Arial, fallback helvetica
              │         ├── For each row → multi_cell(col_width, 5, value, border=1)
              │         └── pdf.output(str(pdf_path))
              │
              └── LOG: Processing complete! Processed: N sheet(s), Skipped: 0 sheet(s)
```

---

## 7. Mode Comparison: Layout vs Table

| Aspect | **Layout Mode** (default) | **Table Mode** |
|---|---|---|
| **Best for** | Worksheets with visual structure, forms, or formatted reports | Clean tabular data, databases, spreadsheets |
| **Data source** | `openpyxl` (raw worksheet, merged cells, column widths) | `pandas` DataFrames |
| **Merged cells** | ✅ Preserved — covered cells shown as empty | ❌ Not applicable (pandas flattens structure) |
| **Column widths** | Uses Excel's defined widths + content measurement | Auto-calculated from longest values |
| **Empty rows/cols** | Respects position within bounding box | Removed during block extraction |
| **Multiple tables** | Renders as one continuous grid | Detects and separates with "BLOCK N" headers |
| **Fallback** | N/A — renders whatever is in the bounding box | "Raw content for sheet" if no table detected |
| **Example output** | Like viewing Excel in a terminal | Like a clean database query result |
| **Row height** | Single line per cell height | Auto-wrapped to fit content |
| **Speed** | Slower (cell-by-cell processing) | Faster (vectorized pandas operations) |

### Visual Examples

**Layout mode output** (preserves Excel's visual arrangement):
```
Name          | Age | Department
--------------------------------
John Smith    | 34  | Engineering
Jane Doe      | 28  | Marketing
Note: These are active employees (this note appears in the same grid)
```

**Table mode output** (extracts clean tables, ignores notes):
```
Name          Age   Department
--------      ----  ----------
John Smith    34    Engineering
Jane Doe      28    Marketing
```

---

## 8. CLI Commands

### Installation

```bash
# Install all required dependencies
pip install -r requirements.txt

# Or install them individually
pip install pandas>=2.0.0 openpyxl>=3.1.0 xlrd>=2.0.1 fpdf2>=2.7.4
```

### Basic Usage

```bash
# Convert an Excel file (default layout mode)
python main.py "data/Input/Your File.xlsx"

# Convert a CSV file
python main.py "data/Input/data.csv"
```

### Output Options

```bash
# Specify custom output directory
python main.py report.xlsx -o "my_output"

# Prevent overwriting existing files (creates numbered copies)
python main.py report.xlsx --no-overwrite
```

### Mode Selection

```bash
# Layout mode (default — preserves Excel grid)
python main.py report.xlsx --mode layout

# Table mode (extracts clean data blocks)
python main.py report.xlsx --mode table
```

### Formatting Control

```bash
# Set maximum column width (in characters) before wrapping
python main.py report.xlsx --max-column-width 40
```

### Verbose / Debug

```bash
# Enable detailed debug logging
python main.py report.xlsx -v

# Verbose + custom output
python main.py report.xlsx -o "data/Output" -v
```

### Help

```bash
python main.py --help
```

**Output**:
```
usage: main.py [-h] [-o OUTPUT] [--mode {table,layout}] [--max-column-width MAX_COLUMN_WIDTH]
               [--no-overwrite] [-v]
               input

Convert Excel sheets or CSV files to formatted .txt files.

positional arguments:
  input                 Path to the source Excel or CSV file

options:
  -h, --help            show this help message and exit
  -o, --output OUTPUT   Base folder for output (default: data/Output)
  --mode {table,layout}
                        Export style: table extracts data blocks; layout approximates the
                        Excel grid (default: layout)
  --max-column-width MAX_COLUMN_WIDTH
                        Maximum text width per column before wrapping (default: 60)
  --no-overwrite        Keep existing files and create numbered names instead of overwriting
  -v, --verbose         Enable verbose logging
```

---

## 9. Complete Usage Examples

### Example 1: Basic Conversion

```bash
# Input file
python main.py "data/Input/Sales Report.xlsx"

# Output:
# data/Output/Sales Report/TXT/Sales.txt
# data/Output/Sales Report/TXT/Summary.txt
# data/Output/Sales Report/PDF/Sales.pdf
# data/Output/Sales Report/PDF/Summary.pdf
```

### Example 2: Table Extraction Mode

```bash
# Extract clean tables, ignore titles/notes
python main.py "data/Input/Sales Report.xlsx" --mode table
```

### Example 3: Custom Output + No Overwrite

```bash
# Save to "results/" folder, preserve existing files
python main.py "data/Input/report.xlsx" -o results --no-overwrite

# If "results/report/TXT/Sheet1.txt" already exists:
# → "results/report/TXT/Sheet1_1.txt"
```

### Example 4: Narrow Columns for Better Readability

```bash
# Wrap long text at 30 characters
python main.py "data/Input/wide_report.xlsx" --max-column-width 30
```

### Example 5: CSV File

```bash
python main.py "data/Input/export.csv"
# → data/Output/export/TXT/export.txt
# → data/Output/export/PDF/export.pdf
```

### Example 6: Debug Mode

```bash
python main.py "data/Input/problem_file.xlsx" -v
# Shows detailed logs of block detection, row counts, column widths
```

### Example 7: Full Pipeline with All Options

```bash
python main.py "data/Input/Annual Report.xlsx" ^
    -o "data/Output" ^
    --mode table ^
    --max-column-width 50 ^
    --no-overwrite ^
    -v
```
*(Use `^` for line continuation in Windows CMD; use `\` in bash/zsh)*

---

## 10. Output Format Specifications

### TXT File Format (Fixed-Width Tables)

```
Column1        Column2        Column3
----------     ----------     ----------
Value 1        Row 1 Col 2    Row 1 Col 3
Long text      Short          Medium
wrapped                          text
here                         Wrapped
```

- **Alignment**: Left-aligned
- **Separator**: Two spaces between columns
- **Header separator**: Dashes (`-`) matching each column's width
- **Wrapping**: Long text wraps to multiple lines within the column
- **Encoding**: UTF-8
- **Newlines**: Unix-style (`\n`)

### TXT File Format (Layout Mode)

```
Title Text
Column A       Column B       Column C
----------     ----------     ----------
Cell Value     Cell Value     Cell Value
Empty Cell     Filled Cell    
Note: This is a note-like row outside the table
```

- Merged cells: covered cells rendered as empty strings
- Column widths: based on Excel's defined widths (not auto-calculated)
- Row heights: each cell rendered on one physical line (no multi-line wrapping in layout)

### PDF File Format

- **Page size**: A4 Landscape (297mm × 210mm)
- **Orientation**: Landscape
- **Font**: Unicode (Arial/LiberationSans) with helvetica fallback
- **Font size**: 8pt
- **Borders**: All cells have visible borders
- **Title**: Sheet name centered at top in 14pt bold
- **Scaling**: Columns auto-scaled if total width exceeds page width
- **Margins**: 15mm bottom margin for auto page breaks

---

## 11. Edge Cases Handled

| Edge Case | How It's Handled |
|---|---|
| **Empty file/sheet** | Returns `""` content, logs warning, skips file creation |
| **No clear table** | Falls back to raw `df.to_string(index=False)` export |
| **Merged cells** | Layout mode: covered cells return `""`. Table mode: processed as individual cells |
| **Duplicate headers** | `make_unique_headers()` appends `_2`, `_3` suffixes |
| **Empty headers** | Auto-named as `Column_1`, `Column_2`, ... |
| **NaN values** | `clean_cell_value()` + `pd.isna()` → `""` |
| **Datetime objects** | Formatted as `"YYYY-MM-DD HH:MM:SS"` |
| **Special characters** | Filenames: sanitized with regex. PDF: Unicode font fallback |
| **Newlines in cells** | `preserve_newlines=False` (table mode) → collapsed. `True` (layout) → preserved |
| **Tabs in cells** | Replaced with single space |
| **Invalid filenames** | `sanitize_filename()` replaces `\ / : * ? " < > \|` with `_` |
| **Duplicate filenames** | `get_output_file()` with `--no-overwrite` → `_1`, `_2`, ... |
| **Very long text** | `textwrap.wrap()` splits at word boundaries, `break_long_words=True` for edge cases |
| **.xls files** | Table mode only (openpyxl doesn't support .xls layout) |
| **Multiple tables/sheet** | `get_non_empty_blocks()` separates blocks, labeled as `BLOCK 1`, `BLOCK 2` |
| **Large datasets** | Uses pandas vectorized operations; no recursion or memory-intensive loops |
| **Non-UTF8 characters in PDF** | Helvetica fallback: `encode("latin-1", "replace")` |
| **Unicode in PDF** | Tries Arial → LiberationSans → helvetica fallback |
| **Windows/macOS line endings** | Normalized to Unix `\n` in `clean_cell_value()` |

---

## Quick Reference Card

```bash
# ┌──────────────────────────────────────────────────────────────────┐
# │                     QUICK REFERENCE                              │
# ├──────────────────────────────────────────────────────────────────┤
# │                                                                  │
# │  INSTALL                                                         │
# │    pip install -r requirements.txt                               │
# │                                                                  │
# │  PLACE INPUT FILES                                               │
# │    data/Input/                                                   │
# │                                                                  │
# │  RUN (basic)                                                     │
# │    python main.py "data\Input\Your File.xlsx"                    │
# │                                                                  │
# │  OUTPUT LOCATION                                                 │
# │    data/Output/Your File/                                        │
# │    ├── TXT/                                                      │
# │    └── PDF/                                                      │
# │                                                                  │
# │  USEFUL FLAGS                                                    │
# │    --mode table     Switch to table extraction                   │
# │    -o my_folder     Custom output directory                      │
# │    --max-column-width 40  Narrower columns for readability       │
# │    --no-overwrite   Keep existing files                          │
# │    -v               Verbose logging                              │
# │                                                                  │
# └──────────────────────────────────────────────────────────────────┘
