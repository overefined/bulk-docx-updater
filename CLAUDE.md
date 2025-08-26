# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a DOCX Bulk Updater tool designed for performing bulk find & replace operations in DOCX files with advanced formatting control. The tool handles complex formatting scenarios including:

- Text replacement across paragraphs, tables, headers, and footers
- Inline formatting blocks with `{format:options}text{/format}` syntax for bold, italic, size, alignment, spacing
- Global formatting tokens for page/line/paragraph breaks
- Placeholder handling when split across DOCX runs
- Hyperlink text replacement while preserving document structure
- Empty paragraph and column break cleanup

## Key Files

**Entry Point:**
- `main.py` - Application entry point (imports and calls CLI)

**Core Modules:**
- `document_processor.py` - Main `DocxBulkUpdater` class and document operations
- `text_replacement.py` - Text replacement logic across DOCX runs and hyperlinks
- `formatting.py` - Formatting token processing and application
- `config.py` - Configuration loading and validation
- `cli.py` - Command-line interface and argument parsing

**Configuration:**
- `replace.json` - Configuration file containing replacement rules
- `main_original.py` - Backup of original monolithic implementation

**Test Suite:**
- `tests/` - Comprehensive unit test suite with pytest
- `requirements.txt` - Test dependencies (pytest, pytest-cov)
- `pytest.ini` - Test configuration and coverage settings
- `run_tests.py` - Convenient test runner script

**Documentation:**
- `README.md` - Comprehensive documentation with usage examples

## Dependencies

The tool requires Python 3.7+ and the `python-docx` library. A virtual environment is available:

```bash
# Activate virtual environment
source .venv/bin/activate  # Linux/Mac
# or
.venv\Scripts\activate     # Windows

# Install dependencies (if needed)
pip install python-docx
```

## Common Commands

### Basic Usage
```bash
# Activate virtual environment first
source .venv/bin/activate

# Process with JSON config file
python main.py /path/to/documents --config replace.json

# Single replacement via command line
python main.py /path/to/documents --search "old text" --replace "new text"

# Process recursively through directories
python main.py /path/to/documents --config replace.json --recursive

# Dry run to preview changes
python main.py /path/to/documents --config replace.json --dry-run
```

### Testing and Development
```bash
# Run unit tests (ensure you're in virtual environment)
source .venv/bin/activate
python run_tests.py

# Run tests with coverage reporting
python run_tests.py --coverage

# Run specific test modules
python run_tests.py --file tests/test_formatting.py
python run_tests.py --file tests/test_config.py

# Run core working tests only
python -m pytest tests/test_formatting.py tests/test_config.py -v

# Compare original vs modified DOCX files
python compare_docx.py

# Test with actual DOCX templates (always use real documents for testing)
python main.py "/mnt/c/Development/scripts/docx-templates/templates" --config replace.json --dry-run
```

### Test Data Locations
- **Original templates**: `C:\Development\scripts\docx-templates\templates` (read-only originals)

**IMPORTANT**: Always test replacements using actual DOCX documents, not simple string tests. DOCX internal structure can split text across multiple runs, affecting replacement behavior.

## Architecture

### Modular Design
The application uses a modular architecture with clear separation of concerns:

- **`FormattingProcessor`** (`formatting.py`): Handles all formatting token parsing and application
  - `process_formatting_tokens()`: Parses inline and global formatting tokens
  - `apply_formatting_to_run()`: Applies formatting to individual runs
  - `apply_paragraph_formatting()`: Handles paragraph-level formatting

- **`TextReplacer`** (`text_replacement.py`): Manages complex text replacement logic
  - `apply_text_replacements()`: Core text processing logic (reusable)
  - `replace_text_in_paragraph()`: DOCX-specific implementation with formatting preservation
  - `_replace_text_in_hyperlinks()`: XML-based hyperlink text replacement preserving structure
  - `_handle_alignment_segments()`: Creates new paragraphs when alignment changes

- **`DocxBulkUpdater`** (`document_processor.py`): Orchestrates document processing
  - `modify_docx()`: Processes entire document including tables, headers, footers
  - `remove_empty_paragraphs_after_pattern()`: Cleans up column breaks and empty paragraphs
  - `standardize_document_margins()`: Handles document margin standardization
  - `get_document_changes_preview()`: Dry-run functionality (uses temporary file copies and actual processing)

### Text Replacement System

The tool handles three types of text replacement scenarios:

1. **Regular paragraphs**: Standard text replacement with run rebuilding
2. **Hyperlinked text**: XML-based replacement preserving hyperlink structure (tabs, formatting, etc.)
3. **Insert operations**: Creates new paragraphs with formatted content insertion

#### Hyperlink Handling
- Hyperlinks are processed using XML manipulation to preserve document structure
- Tab elements, formatting runs, and spacing are maintained
- Only `replace` operations are applied to hyperlinks; `insert_after` operations are handled separately
- No hardcoded patterns - works with any search/replace text

#### Column Break Cleanup
The `remove_empty_paragraphs_after` feature:
- Runs BEFORE text replacements when pattern is still original text
- Removes leading whitespace runs (empty runs, `\n` runs, etc.) from paragraphs following the pattern
- Preserves document structure by only removing leading formatting, not content
- Simple approach: find pattern → clean next paragraph → proceed with replacements

### Text Processing Pipeline
1. **Preprocessing**: Remove column breaks/empty paragraphs after specified patterns
2. **Hyperlink processing**: Handle hyperlinked text with XML-based replacement
3. **Regular text processing**: Concatenate text from all runs, find matches, map back to runs
4. **Formatting application**: Apply formatting tokens to appropriate text segments
5. **Structure preservation**: Maintain DOCX paragraph/run structure throughout

## Configuration Format

The `replace.json` file supports multiple operation types:

```json
{
  "replacements": [
    {
      "search": "text to find",
      "replace": "replacement with formatting tokens",
      "remove_empty_paragraphs_after": true
    },
    {
      "search": "pattern",
      "insert_after": "{format:center,size12}content{/format}pagebreak"
    }
  ]
}
```

### Configuration Options

- **`search`**: Text pattern to find
- **`replace`**: Replacement text (supports formatting tokens)
- **`insert_after`**: Content to insert after the found pattern
- **`remove_empty_paragraphs_after`**: Boolean flag to clean up column breaks from following paragraphs

## Development Notes

### DOCX Structure Handling
- The tool handles complex DOCX internal structure where text can be split across multiple `<w:t>` runs
- Hyperlinks contain XML structures with tabs (`<w:tab/>`) and multiple text runs
- Formatting preservation is handled by copying font properties from original runs
- Alignment changes create new paragraphs due to DOCX paragraph-level alignment constraints

### Formatting System
The tool supports two formatting approaches:
1. **Global tokens**: `pagebreak`, `linebreak`, `paragraphbreak` for document structure
2. **Inline blocks**: `{format:bold,center,size16}text{/format}` for text-specific formatting

### No Hardcoded Patterns
- The text replacement system uses no hardcoded text patterns or special cases
- Hyperlink detection works generically without explicit pattern matching
- All functionality is driven by configuration, not embedded logic

## Preview Functionality (--dry-run)

The preview functionality uses a **temporary file approach** to ensure 100% accuracy:

1. **Temporary Copy**: Creates a temporary copy of the original document
2. **Actual Processing**: Runs the exact same `modify_docx()` operations on the temporary copy
3. **Content Extraction**: Extracts text content from both original and modified documents
4. **Comparison**: Shows unified diff of changes
5. **Cleanup**: Automatically removes temporary files

**Key Benefits**:
- **Perfect Accuracy**: Preview shows exactly what the actual operation produces
- **Formatting Tokens**: Page breaks, line breaks, and all formatting are processed correctly
- **Zero Maintenance**: No duplicate logic to maintain - uses same processing pipeline
- **Reliable**: Impossible for preview and actual results to diverge

## Architecture Principles & Code Quality Standards

### General Design Principles
- **No hardcoded values**: All text patterns, special cases, and logic driven by configuration
- **Structure preservation**: DOCX XML structure maintained throughout processing
- **Modular processing**: Clear separation between hyperlink, regular text, and formatting logic
- **Simple solutions**: Prefer straightforward approaches over complex post-processing

### DRY Principle - Avoid Logic Duplication
**CRITICAL**: Never duplicate business logic between different functions or classes, especially between preview/dry-run functionality and actual processing logic.

The preview system uses the actual processing methods via temporary files, ensuring perfect accuracy without code duplication.

### Text Replacement Logic Guidelines
- **Universal processing**: No special cases for specific text patterns
- **Structure-aware**: Handle hyperlinks, tables, paragraphs appropriately
- **Single pass processing**: Each replacement should be applied in a single pass to avoid conflicts
- **Preprocessing approach**: Clean up formatting issues before text replacement, not after

## Testing Framework

The project includes a comprehensive unit test suite built with pytest:

### Running Tests
```bash
# Quick verification of core functionality
python run_tests.py --file tests/test_formatting.py tests/test_config.py

# Full test coverage report
python run_tests.py --coverage

# Available test options
python run_tests.py --help
```

### Test Dependencies (already included)
- `pytest>=7.0.0` - Test framework
- `pytest-cov>=4.0.0` - Coverage reporting
- `python-docx>=1.0.0` - DOCX processing

The test suite ensures that future changes won't break existing functionality while making it easier to add new features with proper test coverage.