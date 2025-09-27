# DOCX Bulk Updater

Run simple, reliable bulk find/replace in DOCX files while preserving structure and formatting.

## Install

Requires Python 3.8+.

```bash
# From the project root
python -m venv .venv
source .venv/bin/activate   # macOS/Linux
# or
.venv\Scripts\activate      # Windows

pip install -r requirements.txt
```

## Quick Start

```bash
# Use a JSON config
python main.py /path/to/docs --config replace.json

# Single replacement from CLI
python main.py /path/to/docs --search "old" --replace "new"

# Recurse into subfolders
python main.py /path/to/docs --config replace.json --recursive

# Preview only (no writes) with unified diff
python main.py /path/to/docs --config replace.json --dry-run
```

## Common Tasks

- Basic text replace (JSON config):
  ```json
  {
    "replacements": [
      { "search": "Old", "replace": "New" }
    ]
  }
  ```

- Insert content after a match:
  ```json
  {
    "replacements": [
      { "search": "SECTION HEADER", "insert_after": "pagebreak{format:center}More{/format}" }
    ]
  }
  ```

- Clean up the next paragraph after a match:
  ```json
  {
    "replacements": [
      { "search": "PATTERN", "replace": "NEW", "remove_empty_paragraphs_after": true }
    ]
  }
  ```

- Repeat table headers:
  - JSON (all tables):
    ```json
    { "replacements": [ { "set_table_header_repeat": true } ] }
    ```
  - JSON (specific header pattern):
    ```json
    { "replacements": [ { "set_table_header_repeat": "Spectrum    Time" } ] }
    ```
  - CLI (all tables): `python main.py <path> --set-table-headers`
  - CLI (specific header pattern): `python main.py <path> --set-table-headers --header-pattern "Spectrum    Time"`

- Change font sizes across a document:
  ```json
  {
    "replacements": [
      { "change_font_size": { "from": 8, "to": 10 } }
    ]
  }
  ```

## XML Mode

Replace raw WordprocessingML XML using files.

```bash
python main.py /path/to/docs \
  --xml-search-file patterns/search.xml \
  --xml-replace-file patterns/replace.xml
```

JSON config:
```json
{
  "replacements": [
    { "xml_mode": true, "search_file": "patterns/search.xml", "replace_file": "patterns/replace.xml" }
  ]
}
```

## Formatting

- Global tokens: `pagebreak`, `linebreak`, `paragraphbreak`
- Inline block: `{format:options}text{/format}`
  - Options supported:
    - Text: `bold`, `italic`, `underline`
    - Alignment: `left`, `center`, `right`, `justify`
    - Font size: `size12`, `size14`, `size16`, … (`sizeNN`)
    - Spacing: `spacebefore6`, `spaceafter6`, … (`spacebeforeNN`, `spaceafterNN`)
    - Font family: `font:Arial Narrow` (any font name)

Example:
```json
{ "search": "Title", "replace": "{format:bold,center,size16,spaceafter6}TITLE{/format}" }
```

## Regex

Use Python-style regex for text replacements.

```json
{
  "replacements": [
    { "search": "ACME\\s+Corp(oration)?", "replace": "TechCorp", "regex": true }
  ]
}
```

## Useful Flags

- `--recursive` process subfolders
- `--pattern "*.docx"` change file match pattern
- `--no-format` disable formatting processing
- `--dry-run` preview changes only
- `--xml-diff` include XML-only diffs in dry-run
- Margin helpers: `--standardize-margins`, `--margins "1,1,1,1"`, `--margin-top 1.25` …

## Inspect XML

```bash
python main.py /path/to/document.docx --inspect-xml  # show structure
python main.py /path/to/document.docx --show-xml     # print full XML
```

## Examples

```bash
python main.py "/path/to/templates" -c replace.json
python main.py "/path/to/test" -c replace.json --dry-run --xml-diff
python main.py "/path/to/document.docx" --inspect-xml
```

---
