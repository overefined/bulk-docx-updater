# Bulk DOCX Updater

Command-line tool to bulk update DOCX files.

## Usage

Basic form

- `python main.py PATH [options]`

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

JSON array of operations with `op` field:

```json
[
  { "op": "replace", "search": "old", "replace": "new" },
  { "op": "xml_replace", "search": "<w:t>xml</w:t>", "replace": "<w:t>new</w:t>" },
  { "op": "set_comments", "value": "Template: {{FILENAME}}" },
  { "op": "clear_properties", "properties": ["author", "company", "title"] }
]
```

### Operations

**Replace text**
```json
{ "op": "replace", "search": "Old Text", "replace": "New Text" }
```

**Replace XML**
```json
{ "op": "xml_replace", "search": "<w:t>old</w:t>", "replace": "<w:t>new</w:t>" }
{ "op": "xml_replace", "search_file": "search.xml", "replace_file": "replace.xml" }
```

**Table operations**
```json
{ "op": "replace_table_cell", "table_header": "Phase, Time", "row": 0, "column": 1, "replace": "Time" }
{ "op": "set_table_column_widths", "table_header": "Phase, Time", "column_widths": [1.5, 2.0] }
{ "op": "table_header_repeat", "pattern": "Phase, Time", "enabled": true }
{ "op": "align_table_cells", "patterns": ["text1", "text2"], "alignment": "left" }
```

Aligns table cells containing specific text patterns. Supported alignments: `left`, `center`, `right`, `justify` (defaults to `left`).

**Image replacement**
```json
{ "op": "replace_image", "image_path": "path/to/logo.png" }
{ "op": "replace_image", "image_path": "logo.png", "scale": 0.5, "center": true }
```

Replaces first image, maintaining aspect ratio. Optional: `scale` (0.5 = 50%), `center` (true/false). Advanced: `name`/`alt_text`/`index` to target specific images.

**Set comments**
```json
{ "op": "set_comments", "value": "{{FILENAME}}" }
```

Sets Comments field (File → Info → Properties). Placeholders: `{{FILENAME}}`, `{{BASENAME}}`, `{{EXTENSION}}`, `{{PARENT_DIR}}`.

**Clear properties**
```json
{ "op": "clear_properties", "properties": ["author", "company"] }
{ "op": "clear_properties", "properties": true }
```

Clears document properties. Use `true` to clear all common properties.

**Supported:** `title`, `subject`, `author`, `keywords`, `comments`, `last_modified_by`, `category`, `content_status`, `company`

**Other**
```json
{ "op": "cleanup_empty_after", "pattern": "HEADER" }
{ "op": "font_size", "from": 8, "to": 10 }
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
