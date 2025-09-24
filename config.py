"""
Configuration loading and validation for DOCX bulk updater.

Handles loading replacement rules from JSON files and validating
configuration structure.
"""
from __future__ import annotations
import json
import sys
from pathlib import Path
from typing import List, Dict
import logging


def load_replacements_from_json(config_file: Path) -> List[Dict[str, str]]:
    """Load replacements from a JSON configuration file."""
    try:
        with open(config_file, 'r', encoding='utf-8') as f:
            data = json.load(f)

        if isinstance(data, list):
            replacements = data
        elif isinstance(data, dict) and 'replacements' in data:
            replacements = data['replacements']
        else:
            raise ValueError("JSON must be a list of replacements or contain a 'replacements' key")

        # Process file references for large XML content
        config_dir = config_file.parent
        for replacement in replacements:
            replacement = _process_file_references(replacement, config_dir)

        return replacements

    except Exception as e:
        logging.getLogger(__name__).error("Error loading config file %s: %s", config_file, e)
        sys.exit(1)


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


def validate_replacements(replacements: List[Dict[str, str]]) -> None:
    """Validate replacement configuration structure."""
    for i, repl in enumerate(replacements):
        if not isinstance(repl, dict):
            logging.getLogger(__name__).error("Error: Replacement %s must be a dictionary", i)
            sys.exit(1)

        # Must have either search/replace pair OR search/insert_after pair OR standalone remove_empty_paragraphs_after OR set_table_header_repeat
        # Also support file-based references before content is loaded
        has_search_replace = ('search' in repl and 'replace' in repl) or ('search_file' in repl and 'replace_file' in repl) or ('search_file' in repl and 'replace' in repl) or ('search' in repl and 'replace_file' in repl)
        has_search_insert_after = ('search' in repl and 'insert_after' in repl) or ('search_file' in repl and 'insert_after' in repl)
        has_standalone_cleanup_action = 'remove_empty_paragraphs_after' in repl and 'search' not in repl and 'search_file' not in repl
        has_table_header_repeat = 'set_table_header_repeat' in repl
        has_font_size_change = 'change_font_size' in repl

        if not (has_search_replace or has_search_insert_after or has_standalone_cleanup_action or has_table_header_repeat or has_font_size_change):
            logging.getLogger(__name__).error("Error: Replacement %s must have either 'search'+'replace' keys, 'search_file'+'replace_file' keys, 'search'+'insert_after' keys, standalone 'remove_empty_paragraphs_after' key, 'set_table_header_repeat' key, or 'change_font_size' key", i)
            sys.exit(1)

        # Cannot have both replace and insert_after
        if 'replace' in repl and 'insert_after' in repl:
            logging.getLogger(__name__).error("Error: Replacement %s cannot have both 'replace' and 'insert_after' keys", i)
            sys.exit(1)

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

        # Validate regex option
        if 'regex' in repl:
            regex_value = repl['regex']
            if not isinstance(regex_value, bool):
                logging.getLogger(__name__).error("Error: 'regex' in replacement %s must be a boolean", i)
                sys.exit(1)

        # Validate ignore_case option
        if 'ignore_case' in repl:
            ignore_case_value = repl['ignore_case']
            if not isinstance(ignore_case_value, bool):
                logging.getLogger(__name__).error("Error: 'ignore_case' in replacement %s must be a boolean", i)
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