"""Configuration loading and validation for DOCX bulk updater operations.

Supports a dict-based JSON config format that gets expanded into the internal
operations array consumed by the processing engine.
"""

import json
import sys
from pathlib import Path
from typing import List, Dict, Any, Tuple
import logging


# Valid operation keys in the dict config
_OPERATION_KEYS = {
    'replace', 'xml_replace', 'font_size', 'clear_properties',
    'set_comments', 'table_header_repeat', 'cleanup_empty_after',
    'replace_image', 'align_table_cells', 'replace_table_cell',
    'set_table_column_widths', 'replace_in_table',
}

# Settings keys (not operations)
_SETTINGS_KEYS = {'margins', 'preserve_formatting'}

# All valid top-level keys
_VALID_KEYS = _OPERATION_KEYS | _SETTINGS_KEYS


def load_operations_from_json(config_file: Path) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """Load operations from a dict-format JSON configuration file.

    Returns:
        Tuple of (operations list, settings dict).
        Settings may contain 'margins' and 'preserve_formatting'.
    """
    try:
        with open(config_file, 'r', encoding='utf-8') as f:
            data = json.load(f)

        if not isinstance(data, dict):
            raise ValueError("JSON config must be a dict (see README for format)")

        # Check for unknown keys
        unknown = set(data.keys()) - _VALID_KEYS
        if unknown:
            raise ValueError(f"Unknown config keys: {', '.join(sorted(unknown))}")

        config_dir = config_file.parent
        operations, settings = _expand_dict_config(data, config_dir)

        # Validate the expanded operations
        validate_operations(operations)

        # Resolve file references after validation (handles paths created during validation)
        for i, op in enumerate(operations):
            operations[i] = _process_file_references(op, config_dir)

        return operations, settings

    except Exception as e:
        logging.getLogger(__name__).error("Error loading config file %s: %s", config_file, e)
        sys.exit(1)


def _expand_dict_config(data: Dict[str, Any], config_dir: Path) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """Expand a dict config into (operations list, settings dict)."""
    operations = []
    settings = {}

    for key, value in data.items():
        if key in _SETTINGS_KEYS:
            if key == 'margins':
                settings['margins'] = _parse_margins_value(value)
                settings['standardize_margins'] = True
            elif key == 'preserve_formatting':
                if not isinstance(value, bool):
                    raise ValueError("'preserve_formatting' must be boolean")
                settings['preserve_formatting'] = value
            continue

        if key == 'replace':
            if not isinstance(value, list):
                raise ValueError("'replace' must be a list of [search, replace] pairs")
            for i, entry in enumerate(value):
                if isinstance(entry, list):
                    if len(entry) < 2 or len(entry) > 3:
                        raise ValueError(f"replace[{i}]: expected [search, replace] or [search, replace, options]")
                    op = {"op": "replace", "search": entry[0], "replace": entry[1]}
                    if len(entry) == 3 and isinstance(entry[2], dict):
                        op.update(entry[2])
                    operations.append(op)
                elif isinstance(entry, dict):
                    op = {"op": "replace"}
                    op.update(entry)
                    operations.append(op)
                else:
                    raise ValueError(f"replace[{i}]: must be a list or dict")

        elif key == 'xml_replace':
            if not isinstance(value, list):
                raise ValueError("'xml_replace' must be a list of {search, replace} dicts")
            for i, entry in enumerate(value):
                if not isinstance(entry, dict):
                    raise ValueError(f"xml_replace[{i}]: must be a dict")
                op = {"op": "xml_replace"}
                op.update(_process_file_references(entry, config_dir))
                operations.append(op)

        elif key == 'font_size':
            if not isinstance(value, dict):
                raise ValueError("'font_size' must be a dict with 'from' and 'to'")
            op = {"op": "font_size"}
            op.update(value)
            operations.append(op)

        elif key == 'clear_properties':
            op = {"op": "clear_properties"}
            if isinstance(value, bool):
                if not value:
                    raise ValueError("'clear_properties': false doesn't make sense")
                op['properties'] = ['author', 'company', 'title', 'subject', 'keywords', 'category']
            elif isinstance(value, list):
                op['properties'] = value
            elif isinstance(value, str):
                op['properties'] = [value]
            else:
                raise ValueError("'clear_properties' must be true, a string, or a list")
            operations.append(op)

        elif key == 'set_comments':
            if not isinstance(value, str):
                raise ValueError("'set_comments' must be a string")
            operations.append({"op": "set_comments", "value": value})

        elif key == 'table_header_repeat':
            op = {"op": "table_header_repeat"}
            if isinstance(value, bool):
                op['enabled'] = value
            elif isinstance(value, str):
                op['pattern'] = value
                op['enabled'] = True
            elif isinstance(value, dict):
                op.update(value)
            else:
                raise ValueError("'table_header_repeat' must be bool, string, or dict")
            operations.append(op)

        elif key == 'cleanup_empty_after':
            if isinstance(value, str):
                value = [value]
            if not isinstance(value, list):
                raise ValueError("'cleanup_empty_after' must be a string or list of strings")
            for pattern in value:
                if not isinstance(pattern, str):
                    raise ValueError("'cleanup_empty_after' patterns must be strings")
                operations.append({"op": "cleanup_empty_after", "pattern": pattern})

        elif key == 'replace_image':
            entries = value if isinstance(value, list) else [value]
            for i, entry in enumerate(entries):
                if not isinstance(entry, dict):
                    raise ValueError(f"replace_image[{i}]: must be a dict")
                op = {"op": "replace_image"}
                op.update(entry)
                operations.append(op)

        elif key == 'align_table_cells':
            entries = value if isinstance(value, list) else [value]
            for i, entry in enumerate(entries):
                if not isinstance(entry, dict):
                    raise ValueError(f"align_table_cells[{i}]: must be a dict")
                op = {"op": "align_table_cells"}
                op.update(entry)
                operations.append(op)

        elif key == 'replace_table_cell':
            entries = value if isinstance(value, list) else [value]
            for i, entry in enumerate(entries):
                if not isinstance(entry, dict):
                    raise ValueError(f"replace_table_cell[{i}]: must be a dict")
                op = {"op": "replace_table_cell"}
                op.update(entry)
                operations.append(op)

        elif key == 'set_table_column_widths':
            entries = value if isinstance(value, list) else [value]
            for i, entry in enumerate(entries):
                if not isinstance(entry, dict):
                    raise ValueError(f"set_table_column_widths[{i}]: must be a dict")
                op = {"op": "set_table_column_widths"}
                op.update(entry)
                operations.append(op)

        elif key == 'replace_in_table':
            entries = value if isinstance(value, list) else [value]
            for i, entry in enumerate(entries):
                if not isinstance(entry, dict):
                    raise ValueError(f"replace_in_table[{i}]: must be a dict")
                op = {"op": "replace_in_table"}
                op.update(entry)
                operations.append(op)

    return operations, settings


def _parse_margins_value(value) -> Dict[str, float]:
    """Parse margins from config value.

    Accepts:
        - String: "1,1,1.5,1.5" (top,bottom,left,right) or preset name
        - Dict: {"top": 1.0, "bottom": 1.0, "left": 1.5, "right": 1.5}
    """
    defaults = {'top': 1.0, 'bottom': 1.0, 'left': 1.0, 'right': 1.0}

    if isinstance(value, dict):
        margins = dict(defaults)
        for k in ('top', 'bottom', 'left', 'right'):
            if k in value:
                if not isinstance(value[k], (int, float)):
                    raise ValueError(f"margins.{k} must be a number")
                margins[k] = float(value[k])
        return margins

    if isinstance(value, str):
        preset = value.strip().lower()
        if preset in ('letter', 'legal', 'a4'):
            return dict(defaults)
        parts = [p.strip() for p in value.split(',')]
        if len(parts) != 4:
            raise ValueError("margins string must be 'top,bottom,left,right' or a preset name")
        try:
            return {
                'top': float(parts[0]),
                'bottom': float(parts[1]),
                'left': float(parts[2]),
                'right': float(parts[3]),
            }
        except ValueError:
            raise ValueError("margins values must be numbers")

    raise ValueError("'margins' must be a string or dict")


def _process_file_references(operation: Dict, config_dir: Path) -> Dict:
    """Process file references in operation configuration."""
    if 'search_file' in operation:
        search_file_path = config_dir / operation['search_file']
        try:
            with open(search_file_path, 'r', encoding='utf-8') as f:
                operation['search'] = f.read().strip()
            del operation['search_file']
        except Exception as e:
            logging.getLogger(__name__).error("Error loading search file %s: %s", search_file_path, e)
            sys.exit(1)

    if 'replace_file' in operation:
        replace_file_path = config_dir / operation['replace_file']
        try:
            with open(replace_file_path, 'r', encoding='utf-8') as f:
                operation['replace'] = f.read().strip()
            del operation['replace_file']
        except Exception as e:
            logging.getLogger(__name__).error("Error loading replace file %s: %s", replace_file_path, e)
            sys.exit(1)

    if 'image_path' in operation:
        image_path_str = operation['image_path']
        if not Path(image_path_str).is_absolute():
            operation['image_path'] = str(config_dir / image_path_str)

    return operation


def validate_operations(operations: List[Dict[str, Any]]) -> None:
    """Validate operations configuration structure."""
    for i, op in enumerate(operations):
        if not isinstance(op, dict):
            logging.getLogger(__name__).error("Error: Operation %s must be a dictionary", i)
            sys.exit(1)

        if 'op' not in op:
            logging.getLogger(__name__).error("Error: Operation %s must have 'op' field", i)
            sys.exit(1)

        op_type = op['op']

        if op_type in ('replace', 'xml_replace'):
            if not (('search' in op and 'replace' in op) or ('search_file' in op or 'replace_file' in op)):
                logging.getLogger(__name__).error("Error: Operation %s: 'replace' requires 'search' and 'replace' fields", i)
                sys.exit(1)

            if 'search' in op and not isinstance(op['search'], str):
                logging.getLogger(__name__).error("Error: Operation %s: 'search' must be a string", i)
                sys.exit(1)
            if 'replace' in op and not isinstance(op['replace'], str):
                logging.getLogger(__name__).error("Error: Operation %s: 'replace' must be a string", i)
                sys.exit(1)
            if 'regex' in op and not isinstance(op['regex'], bool):
                logging.getLogger(__name__).error("Error: Operation %s: 'regex' must be boolean", i)
                sys.exit(1)
            if 'count' in op:
                if not isinstance(op['count'], int) or op['count'] < 0:
                    logging.getLogger(__name__).error("Error: Operation %s: 'count' must be non-negative integer", i)
                    sys.exit(1)
            if 'occurrence' in op:
                if not isinstance(op['occurrence'], int) or op['occurrence'] < 1:
                    logging.getLogger(__name__).error("Error: Operation %s: 'occurrence' must be positive integer (1-based)", i)
                    sys.exit(1)
            if op_type == 'xml_replace':
                op['xml_mode'] = True

        elif op_type == 'cleanup_empty_after':
            if 'pattern' not in op:
                logging.getLogger(__name__).error("Error: Operation %s: 'cleanup_empty_after' requires 'pattern' field", i)
                sys.exit(1)
            if not isinstance(op['pattern'], str):
                logging.getLogger(__name__).error("Error: Operation %s: 'pattern' must be string", i)
                sys.exit(1)

        elif op_type == 'table_header_repeat':
            if 'pattern' in op and not isinstance(op['pattern'], str):
                logging.getLogger(__name__).error("Error: Operation %s: 'pattern' must be string", i)
                sys.exit(1)
            if 'enabled' in op and not isinstance(op['enabled'], bool):
                logging.getLogger(__name__).error("Error: Operation %s: 'enabled' must be boolean", i)
                sys.exit(1)

        elif op_type == 'font_size':
            if 'from' not in op or 'to' not in op:
                logging.getLogger(__name__).error("Error: Operation %s: 'font_size' requires 'from' and 'to' fields", i)
                sys.exit(1)
            if not isinstance(op['from'], (int, float)) or not isinstance(op['to'], (int, float)):
                logging.getLogger(__name__).error("Error: Operation %s: 'from' and 'to' must be numbers", i)
                sys.exit(1)

        elif op_type == 'replace_table_cell':
            for field in ('row', 'column', 'replace'):
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
            if 'table_index' in op and not isinstance(op['table_index'], int):
                logging.getLogger(__name__).error("Error: Operation %s: 'table_index' must be integer", i)
                sys.exit(1)
            if 'table_header' in op and not isinstance(op['table_header'], str):
                logging.getLogger(__name__).error("Error: Operation %s: 'table_header' must be string", i)
                sys.exit(1)
            if 'search' in op and not isinstance(op['search'], str):
                logging.getLogger(__name__).error("Error: Operation %s: 'search' must be string", i)
                sys.exit(1)
            if 'table_index' in op and 'table_header' in op:
                logging.getLogger(__name__).error("Error: Operation %s: cannot specify both 'table_index' and 'table_header'", i)
                sys.exit(1)

        elif op_type == 'set_table_column_widths':
            if 'column_widths' not in op:
                logging.getLogger(__name__).error("Error: Operation %s: 'set_table_column_widths' requires 'column_widths' field", i)
                sys.exit(1)
            column_widths = op['column_widths']
            if not isinstance(column_widths, list):
                logging.getLogger(__name__).error("Error: Operation %s: 'column_widths' must be a list", i)
                sys.exit(1)
            for j, width in enumerate(column_widths):
                if not isinstance(width, (int, float)) or width < 0:
                    logging.getLogger(__name__).error("Error: Operation %s: column_widths[%d] must be non-negative number", i, j)
                    sys.exit(1)
            if 'table_index' in op and not isinstance(op['table_index'], int):
                logging.getLogger(__name__).error("Error: Operation %s: 'table_index' must be integer", i)
                sys.exit(1)
            if 'table_header' in op and not isinstance(op['table_header'], str):
                logging.getLogger(__name__).error("Error: Operation %s: 'table_header' must be string", i)
                sys.exit(1)
            if 'table_index' in op and 'table_header' in op:
                logging.getLogger(__name__).error("Error: Operation %s: cannot specify both 'table_index' and 'table_header'", i)
                sys.exit(1)

        elif op_type == 'replace_image':
            if 'image_path' not in op:
                logging.getLogger(__name__).error("Error: Operation %s: 'replace_image' requires 'image_path' field", i)
                sys.exit(1)
            if not isinstance(op['image_path'], str):
                logging.getLogger(__name__).error("Error: Operation %s: 'image_path' must be string", i)
                sys.exit(1)
            if 'name' in op and not isinstance(op['name'], str):
                logging.getLogger(__name__).error("Error: Operation %s: 'name' must be string", i)
                sys.exit(1)
            if 'alt_text' in op and not isinstance(op['alt_text'], str):
                logging.getLogger(__name__).error("Error: Operation %s: 'alt_text' must be string", i)
                sys.exit(1)
            if 'index' in op and not isinstance(op['index'], int):
                logging.getLogger(__name__).error("Error: Operation %s: 'index' must be integer", i)
                sys.exit(1)
            if 'scale' in op:
                if not isinstance(op['scale'], (int, float)):
                    logging.getLogger(__name__).error("Error: Operation %s: 'scale' must be a number", i)
                    sys.exit(1)
                if op['scale'] <= 0:
                    logging.getLogger(__name__).error("Error: Operation %s: 'scale' must be greater than 0", i)
                    sys.exit(1)
            if 'center' in op and not isinstance(op['center'], bool):
                logging.getLogger(__name__).error("Error: Operation %s: 'center' must be boolean", i)
                sys.exit(1)
            identifier_count = sum([1 for key in ['name', 'alt_text', 'index'] if key in op])
            if identifier_count > 1:
                logging.getLogger(__name__).error("Error: Operation %s: can only specify one of 'name', 'alt_text', or 'index'", i)
                sys.exit(1)

        elif op_type == 'set_comments':
            if 'value' not in op:
                logging.getLogger(__name__).error("Error: Operation %s: 'set_comments' requires 'value' field", i)
                sys.exit(1)
            if not isinstance(op['value'], str):
                logging.getLogger(__name__).error("Error: Operation %s: 'value' must be string", i)
                sys.exit(1)

        elif op_type == 'clear_properties':
            if 'properties' not in op:
                logging.getLogger(__name__).error("Error: Operation %s: 'clear_properties' requires 'properties' field", i)
                sys.exit(1)
            if not isinstance(op['properties'], list):
                logging.getLogger(__name__).error("Error: Operation %s: 'properties' must be a list", i)
                sys.exit(1)
            valid_properties = [
                'title', 'subject', 'author', 'keywords', 'comments',
                'last_modified_by', 'category', 'content_status', 'company'
            ]
            for prop in op['properties']:
                if not isinstance(prop, str):
                    logging.getLogger(__name__).error("Error: Operation %s: property name must be string", i)
                    sys.exit(1)
                if prop not in valid_properties:
                    logging.getLogger(__name__).error("Error: Operation %s: invalid property '%s'. Valid: %s", i, prop, ', '.join(valid_properties))
                    sys.exit(1)

        elif op_type == 'align_table_cells':
            if 'patterns' not in op:
                logging.getLogger(__name__).error("Error: Operation %s: 'align_table_cells' requires 'patterns' field", i)
                sys.exit(1)
            if not isinstance(op['patterns'], list) or len(op['patterns']) == 0:
                logging.getLogger(__name__).error("Error: Operation %s: 'patterns' must be a non-empty list", i)
                sys.exit(1)
            for j, pattern in enumerate(op['patterns']):
                if not isinstance(pattern, str):
                    logging.getLogger(__name__).error("Error: Operation %s: patterns[%d] must be string", i, j)
                    sys.exit(1)
            if 'alignment' not in op:
                op['alignment'] = 'left'
            valid_alignments = ['left', 'center', 'right', 'justify']
            if op['alignment'] not in valid_alignments:
                logging.getLogger(__name__).error("Error: Operation %s: 'alignment' must be one of: %s", i, ', '.join(valid_alignments))
                sys.exit(1)

        elif op_type == 'replace_in_table':
            if 'table_heading' not in op:
                logging.getLogger(__name__).error("Error: Operation %s: 'replace_in_table' requires 'table_heading' field", i)
                sys.exit(1)
            if not isinstance(op['table_heading'], str):
                logging.getLogger(__name__).error("Error: Operation %s: 'table_heading' must be string", i)
                sys.exit(1)
            if 'search' not in op:
                logging.getLogger(__name__).error("Error: Operation %s: 'replace_in_table' requires 'search' field", i)
                sys.exit(1)
            if not isinstance(op['search'], str):
                logging.getLogger(__name__).error("Error: Operation %s: 'search' must be string", i)
                sys.exit(1)
            if 'replace' not in op:
                logging.getLogger(__name__).error("Error: Operation %s: 'replace_in_table' requires 'replace' field", i)
                sys.exit(1)
            if not isinstance(op['replace'], str):
                logging.getLogger(__name__).error("Error: Operation %s: 'replace' must be string", i)
                sys.exit(1)
            if 'regex' in op and not isinstance(op['regex'], bool):
                logging.getLogger(__name__).error("Error: Operation %s: 'regex' must be boolean", i)
                sys.exit(1)
            if 'table_index' in op and not isinstance(op['table_index'], int):
                logging.getLogger(__name__).error("Error: Operation %s: 'table_index' must be integer", i)
                sys.exit(1)

        else:
            logging.getLogger(__name__).error("Error: Operation %s: unsupported operation type '%s'", i, op_type)
            sys.exit(1)
