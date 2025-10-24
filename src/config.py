"""Configuration loading and validation for DOCX bulk updater operations.

This module supports a public "operations" JSON schema while providing
helpers to convert those operations into the internal replacements structure
consumed by the processing engine. It also validates both layers.
"""

import json
import sys
from pathlib import Path
from typing import List, Dict, Any
import logging


def load_operations_from_json(config_file: Path) -> List[Dict[str, Any]]:
    """Load operations from a JSON configuration file.

    Format: [{"search": "old", "replace": "new"}, ...]
    """
    try:
        with open(config_file, 'r', encoding='utf-8') as f:
            data = json.load(f)

        # Only support array format
        if not isinstance(data, list):
            raise ValueError("JSON must be an array of operations")

        operations: List[Dict[str, Any]] = data

        # Process file references for large XML content
        config_dir = config_file.parent
        for i, operation in enumerate(operations):
            # Normalize and load file references
            operations[i] = _normalize_operation(operation, config_dir)

        return operations

    except Exception as e:
        logging.getLogger(__name__).error("Error loading config file %s: %s", config_file, e)
        sys.exit(1)


def _normalize_operation(operation: Dict, config_dir: Path) -> Dict:
    """Normalize operation format and process file references.

    Converts simplified format to normalized format with 'op' field:
    - {"search": "x", "replace": "y"} -> {"op": "replace", "search": "x", "replace": "y"}
    - {"replace_table_cell": {...}} -> {"op": "replace_table_cell", ...}
    """
    # If already has 'op' field, just process file references
    if 'op' in operation:
        return _process_file_references(operation, config_dir)

    # Infer operation type from fields present
    # Check for file references first (they need to be loaded before we can determine the op type)
    if 'search_file' in operation or 'replace_file' in operation:
        # Process file references first
        operation = _process_file_references(operation, config_dir)
        # After loading files, fall through to determine op type

    if 'search' in operation and 'replace' in operation:
        # Text replacement operation
        op_type = 'xml_replace' if operation.get('xml_mode') else 'replace'
        operation['op'] = op_type
        return operation

    # Single-key operation formats
    single_key_ops = [
        'replace_table_cell',
        'set_table_column_widths',
        'cleanup_empty_after',
        'table_header_repeat',
        'font_size',
        'remove_title'
    ]

    for op_name in single_key_ops:
        if op_name in operation:
            # Flatten: {"replace_table_cell": {...}} -> {"op": "replace_table_cell", ...}
            op_config = operation[op_name]
            if isinstance(op_config, dict):
                normalized = {'op': op_name, **op_config}
            elif isinstance(op_config, (str, bool, int, float)):
                # Handle simple values like {"cleanup_empty_after": "HEADER"}
                normalized = {'op': op_name, 'value': op_config}
            else:
                normalized = {'op': op_name}
            return _process_file_references(normalized, config_dir)

    # If we can't infer, return as-is and let validation catch it
    return operation


def _process_file_references(operation: Dict, config_dir: Path) -> Dict:
    """Process file references in operation configuration."""
    # Handle search_file and replace_file for external XML content
    if 'search_file' in operation:
        search_file_path = config_dir / operation['search_file']
        try:
            with open(search_file_path, 'r', encoding='utf-8') as f:
                operation['search'] = f.read().strip()
            # Remove the file reference after loading
            del operation['search_file']
        except Exception as e:
            logging.getLogger(__name__).error("Error loading search file %s: %s", search_file_path, e)
            sys.exit(1)

    if 'replace_file' in operation:
        replace_file_path = config_dir / operation['replace_file']
        try:
            with open(replace_file_path, 'r', encoding='utf-8') as f:
                operation['replace'] = f.read().strip()
            # Remove the file reference after loading
            del operation['replace_file']
        except Exception as e:
            logging.getLogger(__name__).error("Error loading replace file %s: %s", replace_file_path, e)
            sys.exit(1)

    return operation


def validate_operations(operations: List[Dict[str, Any]]) -> None:
    """Validate operations configuration structure."""
    for i, op in enumerate(operations):
        if not isinstance(op, dict):
            logging.getLogger(__name__).error("Error: Operation %s must be a dictionary", i)
            sys.exit(1)

        # Must have 'op' field
        if 'op' not in op:
            logging.getLogger(__name__).error("Error: Operation %s must have 'op' field", i)
            sys.exit(1)

        op_type = op['op']

        # Validate based on operation type
        if op_type in ('replace', 'xml_replace'):
            # Validate replace operation
            # Either inline search/replace or file references must exist; file refs are resolved earlier
            if not (('search' in op and 'replace' in op) or ('search_file' in op or 'replace_file' in op)):
                logging.getLogger(__name__).error("Error: Operation %s: 'replace' requires 'search' and 'replace' fields", i)
                sys.exit(1)

            if 'search' in op and not isinstance(op['search'], str):
                logging.getLogger(__name__).error("Error: Operation %s: 'search' must be a string", i)
                sys.exit(1)
            if 'replace' in op and not isinstance(op['replace'], str):
                logging.getLogger(__name__).error("Error: Operation %s: 'replace' must be a string", i)
                sys.exit(1)

            # Validate optional regex field
            if 'regex' in op and not isinstance(op['regex'], bool):
                logging.getLogger(__name__).error("Error: Operation %s: 'regex' must be boolean", i)
                sys.exit(1)

            # xml_replace implies xml_mode
            if op_type == 'xml_replace':
                op['xml_mode'] = True

        elif op_type == 'cleanup_empty_after':
            # Validate cleanup operation
            # Support both {"op": "cleanup_empty_after", "pattern": "X"} and simplified {"cleanup_empty_after": "X"}
            if 'pattern' not in op and 'value' not in op:
                logging.getLogger(__name__).error("Error: Operation %s: 'cleanup_empty_after' requires 'pattern' or 'value' field", i)
                sys.exit(1)

            pattern = op.get('pattern') or op.get('value')
            if not isinstance(pattern, str):
                logging.getLogger(__name__).error("Error: Operation %s: 'pattern' must be string", i)
                sys.exit(1)

            # Normalize to use 'pattern'
            if 'value' in op:
                op['pattern'] = op.pop('value')

        elif op_type == 'table_header_repeat':
            # Validate table header repeat operation
            # Support {"table_header_repeat": true/false} or {"table_header_repeat": {"pattern": "X"}}
            if 'value' in op:
                # Simplified format: {"table_header_repeat": true} or {"table_header_repeat": {"pattern": "X"}}
                value = op.pop('value')
                if isinstance(value, bool):
                    op['enabled'] = value
                elif isinstance(value, dict):
                    op.update(value)
                elif isinstance(value, str):
                    op['pattern'] = value

            if 'pattern' in op and not isinstance(op['pattern'], str):
                logging.getLogger(__name__).error("Error: Operation %s: 'pattern' must be string", i)
                sys.exit(1)

            if 'enabled' in op and not isinstance(op['enabled'], bool):
                logging.getLogger(__name__).error("Error: Operation %s: 'enabled' must be boolean", i)
                sys.exit(1)

        elif op_type == 'font_size':
            # Validate font size operation
            # Support {"font_size": {"from": 8, "to": 10}}
            if 'value' in op and isinstance(op['value'], dict):
                op.update(op.pop('value'))

            if 'from' not in op or 'to' not in op:
                logging.getLogger(__name__).error("Error: Operation %s: 'font_size' requires 'from' and 'to' fields", i)
                sys.exit(1)

            if not isinstance(op['from'], (int, float)) or not isinstance(op['to'], (int, float)):
                logging.getLogger(__name__).error("Error: Operation %s: 'from' and 'to' must be numbers", i)
                sys.exit(1)

        elif op_type == 'replace_table_cell':
            # Validate table cell replacement operation
            required_fields = ['row', 'column', 'replace']
            for field in required_fields:
                if field not in op:
                    logging.getLogger(__name__).error("Error: Operation %s: 'replace_table_cell' requires '%s' field", i, field)
                    sys.exit(1)

            if not isinstance(op['row'], int) or op['row'] < 0:
                logging.getLogger(__name__).error("Error: Operation %s: 'row' must be non-negative integer", i)
                sys.exit(1)

            if not isinstance(op['column'], int) or op['column'] < 0:
                logging.getLogger(__name__).error("Error: Operation %s: 'column' must be non-negative integer", i)
                sys.exit(1)

            if not isinstance(op['replace'], str):
                logging.getLogger(__name__).error("Error: Operation %s: 'replace' must be string", i)
                sys.exit(1)

            # Optional fields
            if 'table_index' in op and not isinstance(op['table_index'], int):
                logging.getLogger(__name__).error("Error: Operation %s: 'table_index' must be integer", i)
                sys.exit(1)

            if 'table_header' in op and not isinstance(op['table_header'], str):
                logging.getLogger(__name__).error("Error: Operation %s: 'table_header' must be string", i)
                sys.exit(1)

            if 'search' in op and not isinstance(op['search'], str):
                logging.getLogger(__name__).error("Error: Operation %s: 'search' must be string", i)
                sys.exit(1)

            # Cannot specify both table_index and table_header
            if 'table_index' in op and 'table_header' in op:
                logging.getLogger(__name__).error("Error: Operation %s: cannot specify both 'table_index' and 'table_header'", i)
                sys.exit(1)

        elif op_type == 'set_table_column_widths':
            # Validate table column widths operation
            if 'column_widths' not in op:
                logging.getLogger(__name__).error("Error: Operation %s: 'set_table_column_widths' requires 'column_widths' field", i)
                sys.exit(1)

            column_widths = op['column_widths']
            if not isinstance(column_widths, list):
                logging.getLogger(__name__).error("Error: Operation %s: 'column_widths' must be a list", i)
                sys.exit(1)

            # Validate that all widths are numbers
            for j, width in enumerate(column_widths):
                if not isinstance(width, (int, float)) or width < 0:
                    logging.getLogger(__name__).error("Error: Operation %s: column_widths[%d] must be non-negative number", i, j)
                    sys.exit(1)

            # Optional fields
            if 'table_index' in op and not isinstance(op['table_index'], int):
                logging.getLogger(__name__).error("Error: Operation %s: 'table_index' must be integer", i)
                sys.exit(1)

            if 'table_header' in op and not isinstance(op['table_header'], str):
                logging.getLogger(__name__).error("Error: Operation %s: 'table_header' must be string", i)
                sys.exit(1)

            # Cannot specify both table_index and table_header
            if 'table_index' in op and 'table_header' in op:
                logging.getLogger(__name__).error("Error: Operation %s: cannot specify both 'table_index' and 'table_header'", i)
                sys.exit(1)

        elif op_type == 'remove_title':
            # Validate remove_title operation - it doesn't require any additional fields
            # Support both {"remove_title": true} and {"op": "remove_title"}
            pass

        else:
            logging.getLogger(__name__).error("Error: Operation %s: unsupported operation type '%s'", i, op_type)
            sys.exit(1)


def parse_margin_settings(args) -> Dict[str, float]:
    """Parse margin settings from command line arguments."""
    margins = {
        'top': 1.0,
        'bottom': 1.0,
        'left': 1.0,
        'right': 1.0
    }

    if hasattr(args, 'margins') and args.margins:
        try:
            preset = str(args.margins).strip().lower()
            if preset in ('letter', 'legal', 'a4'):
                # For now all presets map to defaults; individual overrides still apply below
                pass
            else:
                parts = [p.strip() for p in str(args.margins).split(',')]
                if len(parts) != 4:
                    raise ValueError("Expected 4 comma-separated values for margins")
                margins['top'] = float(parts[0])
                margins['bottom'] = float(parts[1])
                margins['left'] = float(parts[2])
                margins['right'] = float(parts[3])
        except ValueError as e:
            logging.getLogger(__name__).error("Error parsing margin settings: %s", e)
            sys.exit(1)

    # Apply individual margin overrides
    if hasattr(args, 'margin_top') and args.margin_top is not None:
        margins['top'] = args.margin_top
    if hasattr(args, 'margin_bottom') and args.margin_bottom is not None:
        margins['bottom'] = args.margin_bottom
    if hasattr(args, 'margin_left') and args.margin_left is not None:
        margins['left'] = args.margin_left
    if hasattr(args, 'margin_right') and args.margin_right is not None:
        margins['right'] = args.margin_right

    return margins


# --- Conversion and compatibility helpers ---

def _operations_to_replacements(operations: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Convert validated operations into internal replacements structure."""
    replacements: List[Dict[str, Any]] = []

    for op in operations:
        op_type = op.get('op')

        if op_type in ('replace', 'xml_replace'):
            item: Dict[str, Any] = {
                'search': op.get('search', ''),
                'replace': op.get('replace', ''),
            }
            if op.get('regex') is not None:
                item['regex'] = bool(op['regex'])
            # xml_replace or explicit xml_mode flag
            if op_type == 'xml_replace' or op.get('xml_mode'):
                item['xml_mode'] = True
            replacements.append(item)

        elif op_type == 'cleanup_empty_after':
            replacements.append({'remove_empty_paragraphs_after': op['pattern']})

        elif op_type == 'table_header_repeat':
            pattern = op.get('pattern')
            enabled = op.get('enabled', True)
            if pattern is None and enabled:
                # Enable header repeat on first row of all tables
                replacements.append({'set_table_header_repeat': True})
            elif pattern is None and not enabled:
                replacements.append({'set_table_header_repeat': {'pattern': None, 'enabled': False}})
            else:
                replacements.append({'set_table_header_repeat': {'pattern': pattern, 'enabled': bool(enabled)}})

        elif op_type == 'font_size':
            replacements.append({'change_font_size': {'from': op['from'], 'to': op['to']}})

        elif op_type == 'replace_table_cell':
            cfg: Dict[str, Any] = {
                'row': op['row'],
                'column': op['column'],
                'replace': op['replace'],
            }
            if 'table_index' in op:
                cfg['table_index'] = op['table_index']
            if 'table_header' in op:
                cfg['table_header'] = op['table_header']
            if 'search' in op:
                cfg['search'] = op['search']
            replacements.append({'replace_table_cell': cfg})

        elif op_type == 'set_table_column_widths':
            cfg2: Dict[str, Any] = {
                'column_widths': op['column_widths']
            }
            if 'table_index' in op:
                cfg2['table_index'] = op['table_index']
            if 'table_header' in op:
                cfg2['table_header'] = op['table_header']
            replacements.append({'set_table_column_widths': cfg2})

        else:
            logging.getLogger(__name__).error("Unsupported operation type during conversion: %s", op_type)
            sys.exit(1)

    return replacements


def load_replacements_from_json(config_file: Path) -> List[Dict[str, Any]]:
    """Load operations from JSON and convert to internal replacements."""
    operations = load_operations_from_json(config_file)
    validate_operations(operations)
    return _operations_to_replacements(operations)


def validate_replacements(replacements: List[Dict[str, Any]]) -> None:
    """Validate internal replacements structure for runtime safety."""
    for i, repl in enumerate(replacements):
        if not isinstance(repl, dict):
            logging.getLogger(__name__).error("Error: Replacement %s must be a dictionary", i)
            sys.exit(1)

        # Recognized shapes
        if 'search' in repl or 'replace' in repl:
            if 'search' not in repl or 'replace' not in repl:
                logging.getLogger(__name__).error("Error: Replacement %s must have both 'search' and 'replace'", i)
                sys.exit(1)
            if not isinstance(repl['search'], str) or not isinstance(repl['replace'], str):
                logging.getLogger(__name__).error("Error: Replacement %s 'search' and 'replace' must be strings", i)
                sys.exit(1)
            if 'regex' in repl and not isinstance(repl['regex'], bool):
                logging.getLogger(__name__).error("Error: Replacement %s 'regex' must be boolean", i)
                sys.exit(1)
            if 'xml_mode' in repl and not isinstance(repl['xml_mode'], bool):
                logging.getLogger(__name__).error("Error: Replacement %s 'xml_mode' must be boolean", i)
                sys.exit(1)
            if repl.get('xml_mode') and (('search' not in repl) or ('replace' not in repl)):
                logging.getLogger(__name__).error("Error: Replacement %s xml_mode requires 'search' and 'replace'", i)
                sys.exit(1)
            if 'remove_empty_paragraphs_after' in repl and repl['remove_empty_paragraphs_after'] is not True:
                logging.getLogger(__name__).error("Error: Replacement %s cleanup flag on search/replace must be True", i)
                sys.exit(1)
            continue

        if 'remove_empty_paragraphs_after' in repl:
            if not isinstance(repl['remove_empty_paragraphs_after'], str):
                logging.getLogger(__name__).error("Error: Replacement %s standalone cleanup requires string pattern", i)
                sys.exit(1)
            continue

        if 'set_table_header_repeat' in repl:
            cfg = repl['set_table_header_repeat']
            if isinstance(cfg, bool):
                pass
            elif isinstance(cfg, str):
                pass
            elif isinstance(cfg, dict):
                if 'enabled' in cfg and not isinstance(cfg['enabled'], bool):
                    logging.getLogger(__name__).error("Error: Replacement %s 'enabled' must be boolean", i)
                    sys.exit(1)
                if 'pattern' in cfg and cfg['pattern'] is not None and not isinstance(cfg['pattern'], str):
                    logging.getLogger(__name__).error("Error: Replacement %s 'pattern' must be string or None", i)
                    sys.exit(1)
            else:
                logging.getLogger(__name__).error("Error: Replacement %s invalid 'set_table_header_repeat' value", i)
                sys.exit(1)
            continue

        if 'change_font_size' in repl:
            cfg = repl['change_font_size']
            if not isinstance(cfg, dict) or not isinstance(cfg.get('from'), (int, float)) or not isinstance(cfg.get('to'), (int, float)):
                logging.getLogger(__name__).error("Error: Replacement %s invalid 'change_font_size' config", i)
                sys.exit(1)
            continue

        if 'set_table_column_widths' in repl:
            cfg = repl['set_table_column_widths']
            if not isinstance(cfg, dict) or 'column_widths' not in cfg or not isinstance(cfg['column_widths'], list):
                logging.getLogger(__name__).error("Error: Replacement %s invalid 'set_table_column_widths' config", i)
                sys.exit(1)
            for j, width in enumerate(cfg['column_widths']):
                if not isinstance(width, (int, float)) or width < 0:
                    logging.getLogger(__name__).error("Error: Replacement %s column_widths[%d] must be non-negative number", i, j)
                    sys.exit(1)
            if 'table_index' in cfg and not isinstance(cfg['table_index'], int):
                logging.getLogger(__name__).error("Error: Replacement %s 'table_index' must be integer", i)
                sys.exit(1)
            if 'table_header' in cfg and not isinstance(cfg['table_header'], str):
                logging.getLogger(__name__).error("Error: Replacement %s 'table_header' must be string", i)
                sys.exit(1)
            if 'table_index' in cfg and 'table_header' in cfg:
                logging.getLogger(__name__).error("Error: Replacement %s cannot specify both 'table_index' and 'table_header'", i)
                sys.exit(1)
            continue

        if 'replace_table_cell' in repl:
            cfg = repl['replace_table_cell']
            if not isinstance(cfg, dict):
                logging.getLogger(__name__).error("Error: Replacement %s invalid 'replace_table_cell' config", i)
                sys.exit(1)
            for field in ('row', 'column', 'replace'):
                if field not in cfg:
                    logging.getLogger(__name__).error("Error: Replacement %s missing '%s' in 'replace_table_cell'", i, field)
                    sys.exit(1)
            if not isinstance(cfg['row'], int) or not isinstance(cfg['column'], int):
                logging.getLogger(__name__).error("Error: Replacement %s 'row' and 'column' must be integers", i)
                sys.exit(1)
            if not isinstance(cfg['replace'], str):
                logging.getLogger(__name__).error("Error: Replacement %s 'replace' must be string", i)
                sys.exit(1)
            if 'table_index' in cfg and not isinstance(cfg['table_index'], int):
                logging.getLogger(__name__).error("Error: Replacement %s 'table_index' must be integer", i)
                sys.exit(1)
            if 'table_header' in cfg and not isinstance(cfg['table_header'], str):
                logging.getLogger(__name__).error("Error: Replacement %s 'table_header' must be string", i)
                sys.exit(1)
            if 'search' in cfg and not isinstance(cfg['search'], str):
                logging.getLogger(__name__).error("Error: Replacement %s 'search' must be string", i)
                sys.exit(1)
            if 'table_index' in cfg and 'table_header' in cfg:
                logging.getLogger(__name__).error("Error: Replacement %s cannot specify both 'table_index' and 'table_header'", i)
                sys.exit(1)
            continue

        # If we got here, no recognized action keys were present
        logging.getLogger(__name__).error("Error: Replacement %s has no valid action keys", i)
        sys.exit(1)
