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

Create a `replace.json` file for text replacements:

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

#### 4. Table Header Repeat Operation
Set table rows to repeat as headers on each page:

```json
{
  "set_table_header_repeat": "Spectrum    Time    Phase"
}
```

#### 5. Font Size Change Operation
Change all text with a specific font size to a new size:

```json
{
  "change_font_size": {
    "from": 5,
    "to": 6
  }
}
```

#### 6. XML Mode Operation
Direct manipulation of WordprocessingML XML:

```json
{
  "search": "<w:t>Old</w:t>",
  "replace": "<w:t>New</w:t>",
  "xml_mode": true,
  "regex": false,
  "ignore_case": false
}
```

#### 7. File-Based XML Operation
For large XML blocks with embedded quotes:

```json
{
  "search_file": "patterns/search_pattern.xml",
  "replace_file": "patterns/replace_pattern.xml",
  "xml_mode": true,
  "regex": true
}
```

**XML Mode Configuration Options:**
- `xml_mode`: Enable XML replacement mode (boolean, required)
- `regex`: Use regular expressions for pattern matching (boolean, optional)
- `ignore_case`: Case-insensitive matching (boolean, optional)
- `search_file`: External file containing XML search pattern (string, optional)
- `replace_file`: External file containing XML replacement pattern (string, optional)

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

### XML Replacement Mode

For advanced users who need direct XML manipulation, the tool supports raw XML find and replace operations:

#### Inline XML Configuration
```json
{
  "replacements": [
    {
      "search": "<w:t>Old Text</w:t>",
      "replace": "<w:t>New Text</w:t>",
      "xml_mode": true
    },
    {
      "search": "w:val=\"[^\"]*\"",
      "replace": "w:val=\"new_value\"",
      "xml_mode": true,
      "regex": true
    }
  ]
}
```

#### File-Based XML Configuration (Recommended for Large XML)
For large XML blocks (2000+ lines) with embedded quotes, use external files:

```json
{
  "replacements": [
    {
      "search_file": "xml_patterns/complex_table_search.xml",
      "replace_file": "xml_patterns/complex_table_replace.xml",
      "xml_mode": true
    },
    {
      "search_file": "patterns/header_search.xml",
      "replace": "<w:t>Direct replacement text</w:t>",
      "xml_mode": true,
      "regex": true
    }
  ]
}
```

#### Command-Line XML Operations
```bash
# Direct XML replacement with files
python main.py /path/to/documents --xml-search-file search.xml --xml-replace-file replace.xml

# Enable XML mode for command-line text replacement
python main.py /path/to/documents --search "<w:t>old</w:t>" --replace "<w:t>new</w:t>" --xml-mode
```

**XML Mode Features:**
- **Direct XML manipulation**: Search and replace raw WordprocessingML XML
- **Regex support**: Use regular expressions with `"regex": true`
- **Case-insensitive matching**: Use `"ignore_case": true`
- **XML validation**: Automatically validates resulting XML structure
- **Takes precedence**: XML replacements are processed before text replacements

**Important Notes:**
- XML mode only works with `search`/`replace` operations (not `insert_after`)
- Malformed XML replacements are automatically rejected
- File-based XML configuration handles embedded quotes automatically
- Requires understanding of WordprocessingML XML structure
- Use with caution - incorrect XML can corrupt documents
- Always test with `--dry-run` first when working with large XML replacements

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
python main.py <path> --dry-run --xml-diff           # Include XML-level diffs
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

### Table Header Repeat
```bash
python main.py <path> --set-table-headers            # Set first row of all tables to repeat
python main.py <path> --set-table-headers --header-pattern "Spectrum    Time" # Only tables with pattern
```

### Font Size Changes
Change all text with specific font sizes (requires JSON configuration):
```bash
python main.py <path> --config font-size-config.json # Use JSON config for font size changes
```

### XML Replacement
```bash
# File-based XML replacement
python main.py <path> --xml-search-file search.xml --xml-replace-file replace.xml

# Command-line XML mode
python main.py <path> --search "<w:t>old</w:t>" --replace "<w:t>new</w:t>" --xml-mode
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

### Example 4: Font Size Changes
Change all 5-point fonts to 6-point across all documents:

```json
{
  "replacements": [
    {
      "change_font_size": {
        "from": 5,
        "to": 6
      }
    }
  ]
}
```

### Example 5: Combined Operations
Multiple operations in a single configuration:

```json
{
  "replacements": [
    {
      "search": "Company Name",
      "replace": "New Company Name"
    },
    {
      "change_font_size": {
        "from": 8,
        "to": 10
      }
    },
    {
      "set_table_header_repeat": "Column1    Column2    Column3"
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

## Performance Profiling

```bash
# Activate virtual environment first
source .venv/bin/activate

# Run performance profiling (requires profile_test_templates/ directory with DOCX files)
python profile_performance.py
```

### Profiling Features

- **CPU Profiling**: Analyzes function call times and identifies bottlenecks
- **Memory Analysis**: Tracks memory allocation during document processing
- **Module-Specific Analysis**: Focuses on custom modules (document_processor, text_replacement, formatting)
- **Interactive Profile**: Saves detailed profile data for further analysis

### Profiling Output

The profiler generates:
- Top 20 functions by cumulative time
- Module-specific performance breakdown
- Memory usage analysis with top allocations
- Detailed profile file (`profile_results.prof`) for interactive analysis

```bash
# View interactive profile after running profiler
python -m pstats profile_results.prof
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

