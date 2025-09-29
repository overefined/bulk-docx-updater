# DOCX Bulk Updater

Command-line tool to bulk update DOCX files.

## Usage

Basic form

- `python main.py PATH [options]`

Options

- `-c, --config PATH`  JSON file with `operations` array
- `-s, --search TEXT`  Search text (with `--replace`)
- `-r, --replace TEXT` Replacement text (with `--search`)
- `--xml-search-file PATH`  File containing raw WordprocessingML XML to search
- `--xml-replace-file PATH` File containing raw WordprocessingML XML to replace with
- `--set-table-headers`  Set first row (or rows matching `--header-pattern`) to repeat as table headers
- `--header-pattern TEXT` Pattern to identify header rows (used with `--set-table-headers`)
- `--standardize-margins`  Enable margin standardization for all documents
- `--margins VALUE`  Comma-separated margins in inches `top,bottom,left,right` or preset `letter|legal|a4`
- `--margin-top FLOAT`    Override top margin (inches)
- `--margin-bottom FLOAT` Override bottom margin (inches)
- `--margin-left FLOAT`   Override left margin (inches)
- `--margin-right FLOAT`  Override right margin (inches)
- `--recursive`  Recurse into subdirectories
- `--pattern GLOB`  File pattern (default: `*.docx`)
- `--no-format`  Do not preserve formatting during text replacement
- `--dry-run`  Show diffs without modifying files
- `--xml-diff` Include XML-level diffs with `--dry-run`
- `--diff-context INT`  Unified diff context lines (default: 3)
- `--inspect-xml`  Inspect document XML (no modifications)
- `--xml-pattern TEXT`  Filter for XML inspection mode
- `--show-xml`  Print full formatted XML during inspection
- `--verbose`  Enable verbose logging

Examples

- Config file: `python main.py ./docs --config replace.json`
- Single replace: `python main.py ./docs --search "old" --replace "new"`
- XML replace from files: `python main.py ./docs --xml-search-file in.xml --xml-replace-file out.xml`
- Dry run with XML diff: `python main.py ./docs --config replace.json --dry-run --xml-diff`
- Recursive with pattern: `python main.py ./docs --config replace.json --recursive --pattern "*.docx"`
- Standardize margins: `python main.py ./docs --config replace.json --standardize-margins --margins 1.0,1.0,1.0,1.0`
- Set table header rows: `python main.py ./docs --set-table-headers --header-pattern "Phase, Time, O2 %"`

## Config JSON

Use a JSON file with an `operations` array. Each item is one operation.

Minimal structure

```json
{ "operations": [ /* one or more operations */ ] }
```

Supported operations

- Replace text
```json
{ "op": "replace", "search": "Old Text", "replace": "New Text" }
```

- Replace XML (WordprocessingML)
```json
{ "op": "xml_replace", "search": "<w:t>old</w:t>", "replace": "<w:t>new</w:t>" }
```

- Replace XML from files
```json
{ "op": "xml_replace", "search_file": "search.xml", "replace_file": "replace.xml" }
```

- Repeat table header rows
```json
{ "op": "table_header_repeat", "pattern": "Phase, Time, O2 %", "enabled": true }
```

- Change font sizes
```json
{ "op": "font_size", "from": 8, "to": 10 }
```

- Set table column widths (inches)
```json
{ "op": "set_table_column_widths", "table_header": "Phase, Time, O2 %", "column_widths": [1.5, 2.0, 1.2] }
```

- Replace a specific table cell
```json
{ "op": "replace_table_cell", "table_header": "Phase, Time, O2 %", "row": 0, "column": 1, "replace": "Time" }
```

- Cleanup empty paragraph after a pattern
```json
{ "op": "cleanup_empty_after", "pattern": "SOME HEADER" }
```
