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
    'set_table_column_widths', 'replace_in_table', 'replace_table',
    'merge_tables', 'landscape_table', 'format_table', 'section_break_before',
    'divider', 'insert_block', 'remove_page_break', 'replace_block',
}

# Op keys whose value is a single dict, or a list of dicts, each expanded
# verbatim into an operation (just stamped with its 'op' name). Ops needing extra
# handling — replace, xml_replace, font_size, clear_properties, set_comments,
# table_header_repeat, cleanup_empty_after, insert_block, remove_page_break — are
# handled explicitly in _expand_dict_config.
_SIMPLE_DICT_OPS = (
    'replace_image', 'align_table_cells', 'replace_table_cell',
    'set_table_column_widths', 'replace_in_table', 'replace_table',
    'merge_tables', 'landscape_table', 'format_table', 'section_break_before',
    'divider',
)

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

        elif key in _SIMPLE_DICT_OPS:
            entries = value if isinstance(value, list) else [value]
            for i, entry in enumerate(entries):
                if not isinstance(entry, dict):
                    raise ValueError(f"{key}[{i}]: must be a dict")
                op = {"op": key}
                op.update(entry)
                operations.append(op)

        elif key == 'insert_block':
            entries = value if isinstance(value, list) else [value]
            for i, entry in enumerate(entries):
                if not isinstance(entry, dict):
                    raise ValueError(f"insert_block[{i}]: must be a dict")
                op = {"op": "insert_block"}
                op.update(_process_file_references(entry, config_dir))
                operations.append(op)

        elif key == 'remove_page_break':
            entries = value if isinstance(value, list) else [value]
            for i, entry in enumerate(entries):
                if isinstance(entry, str):
                    entry = {"in_paragraph": entry}
                if not isinstance(entry, dict):
                    raise ValueError(f"remove_page_break[{i}]: must be a string or dict")
                op = {"op": "remove_page_break"}
                op.update(entry)
                operations.append(op)

        elif key == 'replace_block':
            entries = value if isinstance(value, list) else [value]
            for i, entry in enumerate(entries):
                if not isinstance(entry, dict):
                    raise ValueError(f"replace_block[{i}]: must be a dict")
                op = {"op": "replace_block"}
                op.update(_process_file_references(entry, config_dir))
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


def _fail(index: int, message: str, *args) -> None:
    """Log a config validation error for operation `index` and exit."""
    logging.getLogger(__name__).error("Error: Operation %s: " + message, index, *args)
    sys.exit(1)


# --- Per-op validators. Each takes (op, index) and calls _fail on bad config. ---

def _v_replace(op, i):
    if not (('search' in op and 'replace' in op) or ('search_file' in op or 'replace_file' in op)):
        _fail(i, "'replace' requires 'search' and 'replace' fields")
    if 'search' in op and not isinstance(op['search'], str):
        _fail(i, "'search' must be a string")
    if 'replace' in op and not isinstance(op['replace'], str):
        _fail(i, "'replace' must be a string")
    if 'regex' in op and not isinstance(op['regex'], bool):
        _fail(i, "'regex' must be boolean")
    if 'count' in op and (not isinstance(op['count'], int) or op['count'] < 0):
        _fail(i, "'count' must be non-negative integer")
    if 'occurrence' in op and (not isinstance(op['occurrence'], int) or op['occurrence'] < 1):
        _fail(i, "'occurrence' must be positive integer (1-based)")
    if op['op'] == 'xml_replace':
        op['xml_mode'] = True


def _v_cleanup_empty_after(op, i):
    if 'pattern' not in op:
        _fail(i, "'cleanup_empty_after' requires 'pattern' field")
    if not isinstance(op['pattern'], str):
        _fail(i, "'pattern' must be string")


def _v_table_header_repeat(op, i):
    if 'pattern' in op and not isinstance(op['pattern'], str):
        _fail(i, "'pattern' must be string")
    if 'enabled' in op and not isinstance(op['enabled'], bool):
        _fail(i, "'enabled' must be boolean")


def _v_font_size(op, i):
    if 'from' not in op or 'to' not in op:
        _fail(i, "'font_size' requires 'from' and 'to' fields")
    if not isinstance(op['from'], (int, float)) or not isinstance(op['to'], (int, float)):
        _fail(i, "'from' and 'to' must be numbers")


def _v_replace_table_cell(op, i):
    for field in ('row', 'column', 'replace'):
        if field not in op:
            _fail(i, "'replace_table_cell' requires '%s' field", field)
    if not isinstance(op['row'], int) or op['row'] < 0:
        _fail(i, "'row' must be non-negative integer")
    if not isinstance(op['column'], int) or op['column'] < 0:
        _fail(i, "'column' must be non-negative integer")
    if not isinstance(op['replace'], str):
        _fail(i, "'replace' must be string")
    if 'table_index' in op and not isinstance(op['table_index'], int):
        _fail(i, "'table_index' must be integer")
    if 'table_header' in op and not isinstance(op['table_header'], str):
        _fail(i, "'table_header' must be string")
    if 'search' in op and not isinstance(op['search'], str):
        _fail(i, "'search' must be string")
    if 'table_index' in op and 'table_header' in op:
        _fail(i, "cannot specify both 'table_index' and 'table_header'")


def _v_replace_table(op, i):
    # 'replace' may be supplied inline or resolved later from 'replace_file'
    if not ('replace' in op or 'replace_file' in op):
        _fail(i, "'replace_table' requires 'replace' or 'replace_file' field")
    if 'replace' in op and not isinstance(op['replace'], str):
        _fail(i, "'replace' must be a string")
    if not any(k in op for k in ('table_index', 'table_header', 'match')):
        _fail(i, "'replace_table' requires one of 'table_index', 'table_header', or 'match' to locate the table")


def _v_merge_tables(op, i):
    if not any(k in op for k in ('table_header', 'match')):
        _fail(i, "'merge_tables' requires 'table_header' or 'match' to locate the tables")
    if 'skip_rows' in op and (not isinstance(op['skip_rows'], int) or op['skip_rows'] < 0):
        _fail(i, "'skip_rows' must be a non-negative integer")
    if 'header_row' in op and (not isinstance(op['header_row'], int) or op['header_row'] < 0):
        _fail(i, "'header_row' must be a non-negative integer")


def _v_landscape_table(op, i):
    if not any(k in op for k in ('table_index', 'table_header', 'match')):
        _fail(i, "'landscape_table' requires one of 'table_index', 'table_header', or 'match' to locate the table")
    if 'margins' in op and not isinstance(op['margins'], (str, dict)):
        _fail(i, "'margins' must be a string or dict")


def _v_format_table(op, i):
    if not any(k in op for k in ('table_index', 'table_header', 'match')):
        _fail(i, "'format_table' requires one of 'table_index', 'table_header', or 'match' to locate the table")
    if 'cell_margins' not in op and 'align' not in op:
        _fail(i, "'format_table' requires 'cell_margins' and/or 'align'")
    if 'cell_margins' in op and not isinstance(op['cell_margins'], (int, str)):
        _fail(i, "'cell_margins' must be an int or string")
    if 'align' in op and op['align'] not in ('left', 'center', 'right', 'justify'):
        _fail(i, "'align' must be left, center, right, or justify")


def _v_section_break_before(op, i):
    if 'match' not in op or not isinstance(op['match'], str):
        _fail(i, "'section_break_before' requires a string 'match' field")
    if 'table_index' in op and not isinstance(op['table_index'], int):
        _fail(i, "'table_index' must be integer")
    if 'table_header' in op and not isinstance(op['table_header'], str):
        _fail(i, "'table_header' must be string")
    if 'table_index' in op and 'table_header' in op:
        _fail(i, "cannot specify both 'table_index' and 'table_header'")


def _v_set_table_column_widths(op, i):
    if 'column_widths' not in op:
        _fail(i, "'set_table_column_widths' requires 'column_widths' field")
    if not isinstance(op['column_widths'], list):
        _fail(i, "'column_widths' must be a list")
    for j, width in enumerate(op['column_widths']):
        if not isinstance(width, (int, float)) or width < 0:
            _fail(i, "column_widths[%d] must be non-negative number", j)
    if 'table_index' in op and not isinstance(op['table_index'], int):
        _fail(i, "'table_index' must be integer")
    if 'table_header' in op and not isinstance(op['table_header'], str):
        _fail(i, "'table_header' must be string")
    if 'table_index' in op and 'table_header' in op:
        _fail(i, "cannot specify both 'table_index' and 'table_header'")


def _v_replace_image(op, i):
    if 'image_path' not in op:
        _fail(i, "'replace_image' requires 'image_path' field")
    if not isinstance(op['image_path'], str):
        _fail(i, "'image_path' must be string")
    if 'name' in op and not isinstance(op['name'], str):
        _fail(i, "'name' must be string")
    if 'alt_text' in op and not isinstance(op['alt_text'], str):
        _fail(i, "'alt_text' must be string")
    if 'index' in op and not isinstance(op['index'], int):
        _fail(i, "'index' must be integer")
    if 'scale' in op:
        if not isinstance(op['scale'], (int, float)):
            _fail(i, "'scale' must be a number")
        if op['scale'] <= 0:
            _fail(i, "'scale' must be greater than 0")
    if 'center' in op and not isinstance(op['center'], bool):
        _fail(i, "'center' must be boolean")
    if sum(1 for key in ('name', 'alt_text', 'index') if key in op) > 1:
        _fail(i, "can only specify one of 'name', 'alt_text', or 'index'")


def _v_set_comments(op, i):
    if 'value' not in op:
        _fail(i, "'set_comments' requires 'value' field")
    if not isinstance(op['value'], str):
        _fail(i, "'value' must be string")


def _v_clear_properties(op, i):
    if 'properties' not in op:
        _fail(i, "'clear_properties' requires 'properties' field")
    if not isinstance(op['properties'], list):
        _fail(i, "'properties' must be a list")
    valid_properties = ['title', 'subject', 'author', 'keywords', 'comments',
                        'last_modified_by', 'category', 'content_status', 'company']
    for prop in op['properties']:
        if not isinstance(prop, str):
            _fail(i, "property name must be string")
        if prop not in valid_properties:
            _fail(i, "invalid property '%s'. Valid: %s", prop, ', '.join(valid_properties))


def _v_align_table_cells(op, i):
    if 'patterns' not in op:
        _fail(i, "'align_table_cells' requires 'patterns' field")
    if not isinstance(op['patterns'], list) or len(op['patterns']) == 0:
        _fail(i, "'patterns' must be a non-empty list")
    for j, pattern in enumerate(op['patterns']):
        if not isinstance(pattern, str):
            _fail(i, "patterns[%d] must be string", j)
    if 'alignment' not in op:
        op['alignment'] = 'left'
    if op['alignment'] not in ('left', 'center', 'right', 'justify'):
        _fail(i, "'alignment' must be one of: left, center, right, justify")


def _v_replace_in_table(op, i):
    if 'table_heading' not in op:
        _fail(i, "'replace_in_table' requires 'table_heading' field")
    if not isinstance(op['table_heading'], str):
        _fail(i, "'table_heading' must be string")
    if 'search' not in op:
        _fail(i, "'replace_in_table' requires 'search' field")
    if not isinstance(op['search'], str):
        _fail(i, "'search' must be string")
    if 'replace' not in op:
        _fail(i, "'replace_in_table' requires 'replace' field")
    if not isinstance(op['replace'], str):
        _fail(i, "'replace' must be string")
    if 'regex' in op and not isinstance(op['regex'], bool):
        _fail(i, "'regex' must be boolean")
    if 'table_index' in op and not isinstance(op['table_index'], int):
        _fail(i, "'table_index' must be integer")


def _v_divider(op, i):
    if 'match' not in op or not isinstance(op['match'], str):
        _fail(i, "'divider' requires a string 'match' field")


def _v_insert_block(op, i):
    if not (('before' in op) ^ ('after' in op)):
        _fail(i, "'insert_block' requires exactly one of 'before' or 'after'")
    anchor = op.get('before', op.get('after'))
    if not isinstance(anchor, str) or not anchor.strip():
        _fail(i, "'insert_block' anchor ('before'/'after') must be a non-empty string")
    if not ('replace' in op or 'replace_file' in op):
        _fail(i, "'insert_block' requires 'replace' or 'replace_file' field")
    if 'replace' in op and not isinstance(op['replace'], str):
        _fail(i, "'replace' must be a string")
    if 'skip_if_present' in op and not isinstance(op['skip_if_present'], str):
        _fail(i, "'skip_if_present' must be a string")


def _v_remove_page_break(op, i):
    if 'in_paragraph' not in op or not isinstance(op['in_paragraph'], str) or not op['in_paragraph'].strip():
        _fail(i, "'remove_page_break' requires a non-empty string 'in_paragraph' field")


def _v_replace_block(op, i):
    for anchor in ('from', 'to'):
        if anchor not in op or not isinstance(op[anchor], str) or not op[anchor].strip():
            _fail(i, "'replace_block' requires a non-empty string '%s' field", anchor)
    if 'replace' in op and not isinstance(op['replace'], str):
        _fail(i, "'replace' must be a string")
    for flag in ('keep_from', 'keep_to'):
        if flag in op and not isinstance(op[flag], bool):
            _fail(i, "'%s' must be boolean", flag)
    if 'skip_if_present' in op and not isinstance(op['skip_if_present'], str):
        _fail(i, "'skip_if_present' must be a string")


# op name -> validator. 'replace' and 'xml_replace' share one validator.
_OP_VALIDATORS = {
    'replace': _v_replace,
    'xml_replace': _v_replace,
    'cleanup_empty_after': _v_cleanup_empty_after,
    'table_header_repeat': _v_table_header_repeat,
    'font_size': _v_font_size,
    'replace_table_cell': _v_replace_table_cell,
    'replace_table': _v_replace_table,
    'merge_tables': _v_merge_tables,
    'landscape_table': _v_landscape_table,
    'format_table': _v_format_table,
    'section_break_before': _v_section_break_before,
    'set_table_column_widths': _v_set_table_column_widths,
    'replace_image': _v_replace_image,
    'set_comments': _v_set_comments,
    'clear_properties': _v_clear_properties,
    'align_table_cells': _v_align_table_cells,
    'replace_in_table': _v_replace_in_table,
    'divider': _v_divider,
    'insert_block': _v_insert_block,
    'remove_page_break': _v_remove_page_break,
    'replace_block': _v_replace_block,
}


def validate_operations(operations: List[Dict[str, Any]]) -> None:
    """Validate operations configuration structure."""
    for i, op in enumerate(operations):
        if not isinstance(op, dict):
            logging.getLogger(__name__).error("Error: Operation %s must be a dictionary", i)
            sys.exit(1)
        if 'op' not in op:
            logging.getLogger(__name__).error("Error: Operation %s must have 'op' field", i)
            sys.exit(1)

        validator = _OP_VALIDATORS.get(op['op'])
        if validator is None:
            _fail(i, "unsupported operation type '%s'", op['op'])
        validator(op, i)
