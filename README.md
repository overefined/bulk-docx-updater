# DOCX Bulk Updater

A powerful tool for performing bulk find & replace operations in DOCX files with advanced formatting control and structure preservation.

## Installation

Requires Python 3.7+ and the `python-docx` library:

```bash
# Clone or download the project
cd bulk-docx-updater

# Activate virtual environment (recommended)
source .venv/bin/activate  # Linux/Mac
# or
.venv\Scripts\activate     # Windows

# Install dependencies (if needed)
pip install python-docx
```

## Quick Start

### Basic Usage

```bash
# Use JSON config file
python main.py /path/to/documents --config replace.json

# Single replacement
python main.py /path/to/documents --search "old text" --replace "new text"

# Process entire directory recursively
python main.py /path/to/documents --config replace.json --recursive

# Preview changes without modifying files
python main.py /path/to/documents --config replace.json --dry-run
```

### Sample Configuration

Create a `replace.json` file:

```json
{
  "replacements": [
    {
      "search": "CALIBRATION GAS CERTIFICATES",
      "replace": "CALIBRATION CERTIFICATES",
      "remove_empty_paragraphs_after": true
    },
    {
      "search": "CALIBRATION CERTIFICATES",
      "insert_after": "{format:center,size12}{% if cylinder_certs != none %}{% for cert in cylinder_certs %}{{ cert }}paragraphbreak{% endfor %}{% endif %}{/format}pagebreak"
    }
  ]
}
```

## Configuration Reference

### Operation Types

#### 1. Replace Operation
Replaces all occurrences of the search text:

```json
{
  "search": "Old Text",
  "replace": "New Text with {format:bold}formatting{/format}"
}
```

#### 2. Insert After Operation  
Keeps original text and inserts new content after it:

```json
{
  "search": "SECTION HEADER",
  "insert_after": "pagebreak{format:center}Additional Content{/format}"
}
```

#### 3. Cleanup Operation
Removes leading whitespace from paragraphs following a pattern:

```json
{
  "search": "PATTERN TEXT",
  "replace": "NEW TEXT",
  "remove_empty_paragraphs_after": true
}
```

### Formatting Options

#### Global Formatting Tokens
- `pagebreak` - Inserts a page break
- `linebreak` - Inserts a line break  
- `paragraphbreak` - Creates a new paragraph

#### Inline Formatting Blocks
Use `{format:options}text{/format}` syntax:

```json
{
  "search": "Title",
  "replace": "{format:bold,center,size16}TITLE{/format}"
}
```

**Available options:**
- `bold`, `italic` - Text styling
- `size12`, `size14`, `size16` - Font sizes
- `center`, `left`, `right` - Text alignment
- `spacing_before6`, `spacing_after6` - Paragraph spacing

## Advanced Features

### Hyperlink Text Replacement

The tool intelligently handles hyperlinked text by:
- Preserving XML structure including tab elements
- Maintaining hyperlink functionality
- Processing only `replace` operations on hyperlinks (not `insert_after`)
- Working with any text pattern without hardcoded rules

### Column Break Cleanup

The `remove_empty_paragraphs_after` feature:
- Runs before text replacements to clean document structure
- Removes leading whitespace runs (`\n`, spaces, tabs) from the next paragraph
- Simple preprocessing approach that avoids complex post-processing

### Structure Preservation

Text replacement preserves DOCX structure by:
- Handling text split across multiple runs
- Maintaining paragraph formatting
- Preserving table cell structure
- Processing headers and footers

## Command Line Options

### Basic Options
```bash
python main.py <path> --config <config.json>         # Use JSON configuration
python main.py <path> --search "old" --replace "new" # Single replacement
python main.py <path> --recursive                    # Process subdirectories
python main.py <path> --dry-run                      # Preview changes only
python main.py <path> --dry-run --diff-context 1     # Fewer surrounding lines in unified diff
python main.py <path> --dry-run --xml-diff --xml-diff-sections Body(XML) Headers/Footers(XML)
python main.py <path> --verbose                      # Enable debug logging (internal)
```

### File Processing
```bash
python main.py <path> --pattern "*.docx"             # Custom file pattern
python main.py <path> --no-format                    # Disable formatting processing
```
### Margin Control
```bash
python main.py <path> --standardize-margins          # 1-inch margins
python main.py <path> --margins "1.25,1.0,1.0,1.0"   # Custom margins (T,B,L,R)
python main.py <path> --margin-top 1.5               # Individual margin settings
```

## Examples

### Example 1: Simple Text Replacement
Replace company names across all documents:

```json
{
  "replacements": [
    {
      "search": "ACME Corporation", 
      "replace": "{format:bold}TechCorp Industries{/format}"
    }
  ]
}
```

### Example 2: Section Header with Content
Replace section headers and add formatted content:

```json
{
  "replacements": [
    {
      "search": "TECHNICAL SPECIFICATIONS",
      "replace": "{format:bold,center,size14}TECHNICAL SPECIFICATIONS{/format}"
    },
    {
      "search": "TECHNICAL SPECIFICATIONS", 
      "insert_after": "paragraphbreak{format:italic}Updated specifications as of {{ current_date }}{/format}paragraphbreak"
    }
  ]
}
```

### Example 3: Cleanup and Formatting
Clean up document structure while replacing text:

```json
{
  "replacements": [
    {
      "search": "CALIBRATION DATA",
      "replace": "{format:bold,size12}CALIBRATION RESULTS{/format}",
      "remove_empty_paragraphs_after": true
    }
  ]
}
```

### Example Commands

```bash
python main.py "/path/to/templates" -c replace.json
python main.py "/path/to/test" -c replace.json --dry-run --xml-diff
python main.py "/path/to/document.docx" --inspect-xml
```

## Running Tests

```bash
# Run all tests
python run_tests.py

# Run with coverage report
python run_tests.py --coverage

# Run specific test modules
python run_tests.py --file tests/test_formatting.py
```

