"""
Configuration loading and validation for DOCX bulk updater.

Handles loading replacement rules from JSON files and validating
configuration structure.
"""
from __future__ import annotations
import json
import sys
from pathlib import Path
from typing import List, Dict, Any
import logging


def load_replacements_from_json(config_file: Path) -> List[Dict[str, Any]]:
    """Load replacements from a JSON configuration file."""
    try:
        with open(config_file, 'r', encoding='utf-8') as f:
            data = json.load(f)

        # Only support the new 'operations' schema
        if not (isinstance(data, dict) and isinstance(data.get('operations'), list)):
            raise ValueError("JSON must contain an 'operations' list")

        replacements: List[Dict[str, Any]] = _operations_to_replacements(data['operations'])

        # Process file references for large XML content
        config_dir = config_file.parent
        for replacement in replacements:
            # Load file references and remove the file keys after loading
            replacement = _process_file_references(replacement, config_dir)

        return replacements

    except Exception as e:
        logging.getLogger(__name__).error("Error loading config file %s: %s", config_file, e)
        sys.exit(1)


def _operations_to_replacements(operations: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Translate new-style operations into the existing replacements model.

    Supported operations:
      - { op: 'replace', search, replace, regex? }
      - { op: 'cleanup_empty_after', pattern }
      - { op: 'table_header_repeat', pattern?, enabled? }
      - { op: 'font_size', from, to }
      - { op: 'xml_replace', search_file|search, replace_file|replace }
      - { op: 'replace_table_cell', table_index?, row, column, search?, replace }
    """
    out: List[Dict[str, Any]] = []
    for i, op in enumerate(operations):
        if not isinstance(op, dict) or 'op' not in op:
            logging.getLogger(__name__).error("Invalid operation at index %s: expected object with 'op'", i)
            sys.exit(1)

        kind = op.get('op')
        if kind == 'replace':
            search = op.get('search')
            replace = op.get('replace')
            if search is None or replace is None:
                logging.getLogger(__name__).error("Operation %s: 'replace' requires 'search' and 'replace'", i)
                sys.exit(1)
            item: Dict[str, Any] = {'search': search, 'replace': replace}
            if 'regex' in op:
                item['regex'] = bool(op['regex'])
            out.append(item)
        elif kind == 'cleanup_empty_after':
            pattern = op.get('pattern')
            if not pattern:
                logging.getLogger(__name__).error("Operation %s: 'cleanup_empty_after' requires 'pattern'", i)
                sys.exit(1)
            out.append({'remove_empty_paragraphs_after': pattern})
        elif kind == 'table_header_repeat':
            enabled = op.get('enabled', True)
            pattern = op.get('pattern')
            # Represent using existing key with dict payload
            payload: Dict[str, Any] = {'enabled': bool(enabled)}
            if pattern is not None:
                payload['pattern'] = pattern
            out.append({'set_table_header_repeat': payload})
        elif kind == 'font_size':
            from_size = op.get('from')
            to_size = op.get('to')
            if from_size is None or to_size is None:
                logging.getLogger(__name__).error("Operation %s: 'font_size' requires 'from' and 'to'", i)
                sys.exit(1)
            out.append({'change_font_size': {'from': from_size, 'to': to_size}})
        elif kind == 'xml_replace':
            # Support both file-based and inline xml
            repl: Dict[str, Any] = {'xml_mode': True}
            if 'search_file' in op:
                repl['search_file'] = op['search_file']
            if 'replace_file' in op:
                repl['replace_file'] = op['replace_file']
            if 'search' in op:
                repl['search'] = op['search']
            if 'replace' in op:
                repl['replace'] = op['replace']
            out.append(repl)
        elif kind == 'replace_table_cell':
            # Required: row, column, replace
            # Optional: table_index, table_header (for table selection), search (for validation)
            row = op.get('row')
            column = op.get('column')
            replace = op.get('replace')
            if row is None or column is None or replace is None:
                logging.getLogger(__name__).error("Operation %s: 'replace_table_cell' requires 'row', 'column', and 'replace'", i)
                sys.exit(1)

            item: Dict[str, Any] = {
                'replace_table_cell': {
                    'row': row,
                    'column': column,
                    'replace': replace
                }
            }

            # Optional parameters
            if 'table_index' in op:
                item['replace_table_cell']['table_index'] = op['table_index']
            if 'table_header' in op:
                item['replace_table_cell']['table_header'] = op['table_header']
            if 'search' in op:
                item['replace_table_cell']['search'] = op['search']

            out.append(item)
        else:
            logging.getLogger(__name__).error("Unsupported operation kind '%s'", kind)
            sys.exit(1)

    return out


def _process_file_references(replacement: Dict, config_dir: Path) -> Dict:
    """Process file references in replacement configuration."""
    # Handle search_file and replace_file for external XML content
    if 'search_file' in replacement:
        search_file_path = config_dir / replacement['search_file']
        try:
            with open(search_file_path, 'r', encoding='utf-8') as f:
                replacement['search'] = f.read().strip()
            # Remove the file reference after loading
            del replacement['search_file']
        except Exception as e:
            logging.getLogger(__name__).error("Error loading search file %s: %s", search_file_path, e)
            sys.exit(1)

    if 'replace_file' in replacement:
        replace_file_path = config_dir / replacement['replace_file']
        try:
            with open(replace_file_path, 'r', encoding='utf-8') as f:
                replacement['replace'] = f.read().strip()
            # Remove the file reference after loading
            del replacement['replace_file']
        except Exception as e:
            logging.getLogger(__name__).error("Error loading replace file %s: %s", replace_file_path, e)
            sys.exit(1)

    return replacement


def validate_replacements(replacements: List[Dict[str, Any]]) -> None:
    """Validate replacement configuration structure."""
    for i, repl in enumerate(replacements):
        if not isinstance(repl, dict):
            logging.getLogger(__name__).error("Error: Replacement %s must be a dictionary", i)
            sys.exit(1)

        # Must have either search/replace pair OR standalone remove_empty_paragraphs_after OR set_table_header_repeat
        # Also support file-based references before content is loaded
        has_search_replace = ('search' in repl and 'replace' in repl) or ('search_file' in repl and 'replace_file' in repl) or ('search_file' in repl and 'replace' in repl) or ('search' in repl and 'replace_file' in repl)
        has_standalone_cleanup_action = 'remove_empty_paragraphs_after' in repl and 'search' not in repl and 'search_file' not in repl
        has_table_header_repeat = 'set_table_header_repeat' in repl
        has_font_size_change = 'change_font_size' in repl
        has_table_cell_replace = 'replace_table_cell' in repl

        if not (has_search_replace or has_standalone_cleanup_action or has_table_header_repeat or has_font_size_change or has_table_cell_replace):
            logging.getLogger(__name__).error("Error: Replacement %s must have either 'search'+'replace' keys, 'search_file'+'replace_file' keys, standalone 'remove_empty_paragraphs_after' key, 'set_table_header_repeat' key, 'change_font_size' key, or 'replace_table_cell' key", i)
            sys.exit(1)

        # No special-case checks for unsupported keys

        # Validate XML mode options
        if 'xml_mode' in repl:
            xml_mode_value = repl['xml_mode']
            if not isinstance(xml_mode_value, bool):
                logging.getLogger(__name__).error("Error: 'xml_mode' in replacement %s must be a boolean", i)
                sys.exit(1)

            # XML mode is only compatible with search/replace operations
            if xml_mode_value and not has_search_replace:
                logging.getLogger(__name__).error("Error: 'xml_mode' in replacement %s can only be used with 'search'+'replace' keys", i)
                sys.exit(1)
            # Do not allow regex/ignore_case in XML mode
            if xml_mode_value and ('regex' in repl or 'ignore_case' in repl):
                logging.getLogger(__name__).error("Error: 'regex' and 'ignore_case' are not supported in XML mode (replacement %s)", i)
                sys.exit(1)

        # Validate regex option (applies to text mode only)
        if 'regex' in repl:
            regex_value = repl['regex']
            if not isinstance(regex_value, bool):
                logging.getLogger(__name__).error("Error: 'regex' in replacement %s must be a boolean", i)
                sys.exit(1)
        
        # The 'ignore_case' option is no longer supported
        if 'ignore_case' in repl:
            logging.getLogger(__name__).error("Error: 'ignore_case' is not supported. Remove it from replacement %s", i)
            sys.exit(1)

        # Validate remove_empty_paragraphs_after value
        if 'remove_empty_paragraphs_after' in repl:
            cleanup_value = repl['remove_empty_paragraphs_after']
            # Allow boolean true for search/replace operations, or string pattern for standalone cleanup
            if has_standalone_cleanup_action:
                if not isinstance(cleanup_value, str):
                    logging.getLogger(__name__).error("Error: Standalone 'remove_empty_paragraphs_after' in replacement %s must be a string pattern", i)
                    sys.exit(1)
            else:
                if not isinstance(cleanup_value, bool) or cleanup_value is not True:
                    logging.getLogger(__name__).error("Error: 'remove_empty_paragraphs_after' in replacement %s with search/replace must be boolean true", i)
                    sys.exit(1)
        # Validate set_table_header_repeat payloads
        if 'set_table_header_repeat' in repl:
            payload = repl['set_table_header_repeat']
            if isinstance(payload, dict):
                if 'enabled' in payload and not isinstance(payload['enabled'], bool):
                    logging.getLogger(__name__).error("Error: 'enabled' in set_table_header_repeat must be boolean")
                    sys.exit(1)
                if 'pattern' in payload and not isinstance(payload['pattern'], str):
                    logging.getLogger(__name__).error("Error: 'pattern' in set_table_header_repeat must be string")
                    sys.exit(1)

        # Validate replace_table_cell payloads
        if 'replace_table_cell' in repl:
            payload = repl['replace_table_cell']
            if not isinstance(payload, dict):
                logging.getLogger(__name__).error("Error: 'replace_table_cell' must be a dictionary")
                sys.exit(1)

            # Required fields
            if 'row' not in payload or not isinstance(payload['row'], int) or payload['row'] < 0:
                logging.getLogger(__name__).error("Error: 'row' in replace_table_cell must be a non-negative integer")
                sys.exit(1)
            if 'column' not in payload or not isinstance(payload['column'], int) or payload['column'] < 0:
                logging.getLogger(__name__).error("Error: 'column' in replace_table_cell must be a non-negative integer")
                sys.exit(1)
            if 'replace' not in payload or not isinstance(payload['replace'], str):
                logging.getLogger(__name__).error("Error: 'replace' in replace_table_cell must be a string")
                sys.exit(1)

            # Optional fields
            if 'table_index' in payload and (not isinstance(payload['table_index'], int) or payload['table_index'] < 0):
                logging.getLogger(__name__).error("Error: 'table_index' in replace_table_cell must be a non-negative integer")
                sys.exit(1)
            if 'table_header' in payload and not isinstance(payload['table_header'], str):
                logging.getLogger(__name__).error("Error: 'table_header' in replace_table_cell must be a string")
                sys.exit(1)
            if 'search' in payload and not isinstance(payload['search'], str):
                logging.getLogger(__name__).error("Error: 'search' in replace_table_cell must be a string")
                sys.exit(1)

            # Validate that table_index and table_header are not both specified
            if 'table_index' in payload and 'table_header' in payload:
                logging.getLogger(__name__).error("Error: 'table_index' and 'table_header' cannot both be specified in replace_table_cell")
                sys.exit(1)



def parse_margin_settings(args) -> Dict[str, float]:
    """Parse margin settings from command line arguments."""
    margins = {
        'top': 1.0,
        'bottom': 1.0,
        'left': 1.0,
        'right': 1.0
    }
    
    # Handle preset margin configurations
    if args.margins:
        if args.margins.lower() == 'letter':
            margins = {'top': 1.0, 'bottom': 1.0, 'left': 1.0, 'right': 1.0}
        elif args.margins.lower() == 'legal':
            margins = {'top': 1.0, 'bottom': 1.0, 'left': 1.0, 'right': 1.0}
        elif args.margins.lower() == 'a4':
            margins = {'top': 1.0, 'bottom': 1.0, 'left': 1.0, 'right': 1.0}
        else:
            # Parse comma-separated values
            try:
                parts = [float(x.strip()) for x in args.margins.split(',')]
                if len(parts) == 4:
                    margins = {
                        'top': parts[0],
                        'bottom': parts[1],
                        'left': parts[2],
                        'right': parts[3]
                    }
                else:
                    logging.getLogger(__name__).error("Error: --margins must have exactly 4 comma-separated values (top,bottom,left,right)")
                    sys.exit(1)
            except ValueError:
                logging.getLogger(__name__).error("Error: --margins values must be numbers")
                sys.exit(1)
    
    # Override with individual margin settings if provided
    if args.margin_top is not None:
        margins['top'] = args.margin_top
    if args.margin_bottom is not None:
        margins['bottom'] = args.margin_bottom
    if args.margin_left is not None:
        margins['left'] = args.margin_left
    if args.margin_right is not None:
        margins['right'] = args.margin_right
    
    return margins
