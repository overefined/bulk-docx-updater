# DOCX Bulk Updater

Command-line tool to bulk update DOCX files.

## Usage

Basic form

- `python src/main.py PATH [options]`

Options

- `-c, --config PATH`  JSON config file (array of operations)
- `-s, --search TEXT`  Search text
- `-r, --replace TEXT` Replacement text
- `--recursive`  Recurse into subdirectories
- `--pattern GLOB`  File pattern (default: `*.docx`)
- `--dry-run`  Show diffs without modifying files
- `--verbose`  Enable verbose logging

Examples

- Config file: `python src/main.py ./docs --config replace.json`
- Single replace: `python src/main.py ./docs --search "old" --replace "new"`
- Dry run: `python src/main.py ./docs --config replace.json --dry-run`
- Recursive: `python src/main.py ./docs --config replace.json --recursive`

## Config Format

JSON array of operations:

```json
[
  { "search": "old", "replace": "new" },
  { "search": "<w:t>xml</w:t>", "replace": "<w:t>new</w:t>", "xml_mode": true },
  { "cleanup_empty_after": "HEADER" }
]
```

### Operations

**Replace text**
```json
{ "search": "Old Text", "replace": "New Text" }
```

**Replace XML**
```json
{ "search": "<w:t>old</w:t>", "replace": "<w:t>new</w:t>", "xml_mode": true }
```

**XML from files**
```json
{ "search_file": "search.xml", "replace_file": "replace.xml", "xml_mode": true }
```

**Table operations**
```json
{ "replace_table_cell": { "table_header": "Phase, Time", "row": 0, "column": 1, "replace": "Time" } }
{ "set_table_column_widths": { "table_header": "Phase, Time", "column_widths": [1.5, 2.0] } }
{ "table_header_repeat": { "pattern": "Phase, Time", "enabled": true } }
```

**Other**
```json
{ "cleanup_empty_after": "HEADER" }
{ "font_size": { "from": 8, "to": 10 } }
```

## Formatting

Use `{format:options}text{/format}` for inline formatting:

```json
{ "search": "Title", "replace": "{format:bold,center,size16}New Title{/format}" }
```

**Options:** `bold`, `italic`, `underline`, `left`, `center`, `right`, `justify`, `size8`-`size24`, `spacing0`-`spacing24`

**Tokens:** `pagebreak`, `linebreak`, `paragraphbreak`

```json
{ "search": "PHOTOS", "replace": "Photo1paragraphbreakPhoto2paragraphbreakPhoto3" }
```
