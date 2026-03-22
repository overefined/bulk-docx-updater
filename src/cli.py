"""
Command-line interface for the DOCX bulk updater.

Handles argument parsing and orchestrates the document processing workflow.
"""
from __future__ import annotations
import sys
import argparse
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path
import logging

from src.config import load_operations_from_json, validate_operations
from src.document_processor import DocxBulkUpdater
from src.xml_inspector import inspect_docx_xml


def _process_single_file(operations, preserve_formatting, standardize_margins, margins, file_path):
    """Process a single DOCX file. Top-level function for multiprocessing compatibility."""
    try:
        updater = DocxBulkUpdater(
            operations,
            preserve_formatting=preserve_formatting,
            standardize_margins=standardize_margins,
            margins=margins,
        )
        changed = updater.modify_docx(file_path)
        return (str(file_path), True, changed, None)
    except Exception as e:
        return (str(file_path), False, False, str(e))


def main():
    """Main CLI entry point."""
    parser = argparse.ArgumentParser(description="Bulk find & replace in DOCX files")
    parser.add_argument("path", help="Directory or file path to process")
    parser.add_argument("-c", "--config", type=Path, help="JSON config file (dict format, see README)")
    parser.add_argument("-s", "--search", help="Text to search for (quick one-off mode)")
    parser.add_argument("-r", "--replace", help="Text to replace with (quick one-off mode)")
    parser.add_argument("--recursive", action="store_true", help="Process directories recursively")
    parser.add_argument("--pattern", default="*.docx", help="File pattern to match (default: *.docx)")
    parser.add_argument("-j", "--workers", type=int, default=1, help="Number of parallel workers (default: 1)")
    parser.add_argument("--dry-run", action="store_true", help="Show what would be changed without making changes")
    parser.add_argument("--xml-diff", action="store_true", help="Include XML-level diffs in dry-run output")
    parser.add_argument("--diff-context", type=int, default=3, help="Unified diff context lines (default: 3)")
    parser.add_argument("--verbose", action="store_true", help="Enable verbose logging for debugging")
    parser.add_argument("--inspect-xml", action="store_true", help="Inspect XML structure of DOCX files")
    parser.add_argument("--xml-pattern", help="Text pattern to search for in XML structure")
    parser.add_argument("--show-xml", action="store_true", help="Display full formatted XML content")

    args = parser.parse_args()
    # Logging setup
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.WARNING, format='%(levelname)s: %(message)s')

    # Handle XML inspection mode
    if args.inspect_xml or args.show_xml:
        path = Path(args.path)
        if path.is_file():
            files = [path]
        elif path.is_dir():
            if args.recursive:
                files = list(path.rglob(args.pattern))
            else:
                files = list(path.glob(args.pattern))
        else:
            print(f"Error: Path {path} does not exist", file=sys.stderr)
            sys.exit(1)

        if not files:
            print(f"No files matching pattern '{args.pattern}' found in {path}", file=sys.stderr)
            return

        for file_path in files:
            print(f"\n{'='*80}")
            try:
                inspect_docx_xml(str(file_path), args.xml_pattern, show_full_xml=args.show_xml)
            except Exception as e:
                print(f"Error inspecting {file_path}: {e}", file=sys.stderr)
        return

    # Determine operations source
    settings = {}
    if args.config:
        operations, settings = load_operations_from_json(args.config)
    elif args.search and args.replace:
        operations = [{"op": "replace", "search": args.search, "replace": args.replace}]
    else:
        print("Error: Must provide either --config file or --search/--replace pair", file=sys.stderr)
        sys.exit(1)

    # Validate operations
    validate_operations(operations)

    # Extract settings
    preserve_formatting = settings.get('preserve_formatting', True)
    standardize_margins = settings.get('standardize_margins', False)
    margins = settings.get('margins', {'top': 1.0, 'bottom': 1.0, 'left': 1.0, 'right': 1.0})

    # Find files to process
    path = Path(args.path)
    if path.is_file():
        files = [path]
    elif path.is_dir():
        if args.recursive:
            files = list(path.rglob(args.pattern))
        else:
            files = list(path.glob(args.pattern))
    else:
        print(f"Error: Path {path} does not exist", file=sys.stderr)
        sys.exit(1)

    if not files:
        print(f"No files matching pattern '{args.pattern}' found in {path}", file=sys.stderr)
        return

    print(f"Processing {len(files)} file(s) with {len(operations)} operation(s)...")
    if args.dry_run:
        print("DRY RUN - No files will be modified")

    if args.dry_run:
        # Dry-run stays sequential for readable output
        updater = DocxBulkUpdater(
            operations,
            preserve_formatting=preserve_formatting,
            standardize_margins=standardize_margins,
            margins=margins,
            diff_context=args.diff_context
        )
        for file_path in files:
            try:
                changes = updater.get_document_changes_preview(file_path)
                if changes:
                    print(f"\n{'='*60}")
                    print(f"CHANGES FOR: {file_path}")
                    print('='*60)
                    for section_name, (original_lines, modified_lines) in changes.items():
                        print(f"\n--- {section_name} ---")
                        diff_output = updater.format_diff(original_lines, modified_lines, section_name)
                        print(diff_output)
                    if args.xml_diff:
                        xml_changes = updater.get_document_xml_changes_preview(file_path)
                        for section_name, (original_lines, modified_lines) in xml_changes.items():
                            print(f"\n--- {section_name} ---")
                            diff_output = updater.format_diff(original_lines, modified_lines, section_name)
                            print(diff_output)
                else:
                    print(f"no changes: {file_path}")
            except Exception as e:
                print(f"[ERROR] {file_path}: {e}", file=sys.stderr)
    elif args.workers > 1 and len(files) > 1:
        # Parallel processing
        num_workers = min(args.workers, len(files))
        print(f"Using {num_workers} workers...")
        with ProcessPoolExecutor(max_workers=num_workers) as executor:
            futures = {
                executor.submit(
                    _process_single_file, operations, preserve_formatting,
                    standardize_margins, margins, file_path
                ): file_path
                for file_path in files
            }
            for future in as_completed(futures):
                file_path_str, success, changed, error = future.result()
                if not success:
                    print(f"[ERROR] {file_path_str}: {error}", file=sys.stderr)
                elif changed:
                    print(f"[OK] {file_path_str}")
                else:
                    print(f"- {file_path_str} (no changes)")
    else:
        # Sequential processing
        updater = DocxBulkUpdater(
            operations,
            preserve_formatting=preserve_formatting,
            standardize_margins=standardize_margins,
            margins=margins,
            diff_context=args.diff_context
        )
        for file_path in files:
            try:
                if updater.modify_docx(file_path):
                    print(f"[OK] {file_path}")
                else:
                    print(f"- {file_path} (no changes)")
            except Exception as e:
                print(f"[ERROR] {file_path}: {e}", file=sys.stderr)
