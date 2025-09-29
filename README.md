# DOCX Bulk Updater

Bulk find/replace in DOCX files with formatting and table support.

## Usage

```bash
# Basic usage with config file
python main.py /path/to/docs --config replace.json

# Single replacement
python main.py /path/to/docs --search "old" --replace "new"

# Preview changes only
python main.py /path/to/docs --config replace.json --dry-run

# Process subfolders
python main.py /path/to/docs --config replace.json --recursive
```

## Operations

**Text replacement:**
```json
{ "op": "replace", "search": "Old Text", "replace": "New Text" }
```

**Table cell replacement:**
```json
{ "op": "replace_table_cell", "table_header": "Phase, Time, O2 %", "row": 0, "column": 0, "search": "Phase", "replace": "Time" }
```

**Formatting:**
```json
{ "op": "replace", "search": "Title", "replace": "{format:bold,center}TITLE{/format}" }
```

**Font size changes:**
```json
{ "op": "font_size", "from": 8, "to": 10 }
```

## Config Example

```json
{
  "operations": [
    { "op": "replace", "search": "{{ old_var }}", "replace": "{{ new_var }}" },
    { "op": "replace_table_cell", "table_header": "Phase, Time, O2 %", "row": 0, "column": 0, "replace": "Time" },
    { "op": "replace_table_cell", "table_header": "Time, Phase, O2 %", "row": 0, "column": 1, "replace": "Phase" }
  ]
}
```
