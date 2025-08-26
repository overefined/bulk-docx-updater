"""
Command-line interface for the DOCX bulk updater.

Handles argument parsing and orchestrates the document processing workflow.
"""
from __future__ import annotations
import sys
import argparse
from pathlib import Path
import logging

from config import load_replacements_from_json, validate_replacements, parse_margin_settings
from document_processor import DocxBulkUpdater
from xml_inspector import inspect_docx_xml


def main():
    """Main CLI entry point."""
    parser = argparse.ArgumentParser(description="Bulk find & replace in DOCX files")
    parser.add_argument("path", help="Directory or file path to process")
    parser.add_argument("-c", "--config", type=Path, help="JSON config file with replacements")
    parser.add_argument("-s", "--search", help="Text to search for")
    parser.add_argument("-r", "--replace", help="Text to replace with")
    parser.add_argument("--recursive", action="store_true", help="Process directories recursively")
    parser.add_argument("--pattern", default="*.docx", help="File pattern to match (default: *.docx)")
    parser.add_argument("--no-format", action="store_true", help="Don't preserve formatting")
    parser.add_argument("--dry-run", action="store_true", help="Show what would be changed without making changes")
    parser.add_argument("--xml-diff", action="store_true", help="Include XML-level diffs in dry-run output")
    parser.add_argument("--diff-context", type=int, default=3, help="Unified diff context lines (default: 3)")
    parser.add_argument("--xml-diff-sections", nargs="*", choices=["Body(XML)", "Tables(XML)", "Headers/Footers(XML)"], help="Restrict XML diff to specific sections")
    parser.add_argument("--verbose", action="store_true", help="Enable verbose logging for debugging")
    parser.add_argument("--standardize-margins", action="store_true", help="Standardize margins across all documents")
    parser.add_argument("--margins", help="Comma-separated margins in inches (top,bottom,left,right) or preset name (letter,legal,a4)")
    parser.add_argument("--margin-top", type=float, help="Top margin in inches")
    parser.add_argument("--margin-bottom", type=float, help="Bottom margin in inches")
    parser.add_argument("--margin-left", type=float, help="Left margin in inches")
    parser.add_argument("--margin-right", type=float, help="Right margin in inches")
    parser.add_argument("--inspect-xml", action="store_true", help="Inspect XML structure of DOCX files")
    parser.add_argument("--xml-pattern", help="Text pattern to search for in XML structure")
    
    args = parser.parse_args()
    # Logging setup
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.WARNING, format='%(levelname)s: %(message)s')
    
    # Handle XML inspection mode
    if args.inspect_xml:
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
                inspect_docx_xml(str(file_path), args.xml_pattern, show_full_xml=False)
            except Exception as e:
                print(f"Error inspecting {file_path}: {e}", file=sys.stderr)
        return
    
    # Parse margin settings
    margins = None
    if args.standardize_margins:
        margins = parse_margin_settings(args)
    
    # Determine replacements source
    if args.config:
        replacements = load_replacements_from_json(args.config)
    elif args.search and args.replace:
        replacements = [{"search": args.search, "replace": args.replace}]
    else:
        print("Error: Must provide either --config file or --search/--replace pair", file=sys.stderr)
        sys.exit(1)
    
    # Validate replacements
    validate_replacements(replacements)
    
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
    
    # Process files
    updater = DocxBulkUpdater(
        replacements, 
        preserve_formatting=not args.no_format,
        standardize_margins=args.standardize_margins,
        margins=margins,
        diff_context=args.diff_context
    )
    
    print(f"Processing {len(files)} file(s) with {len(replacements)} replacement(s)...")
    if args.dry_run:
        print("DRY RUN - No files will be modified")
    
    for file_path in files:
        try:
            if args.dry_run:
                # Get preview of changes with diff
                changes = updater.get_document_changes_preview(file_path)
                
                if changes:
                    print(f"\n{'='*60}")
                    print(f"CHANGES FOR: {file_path}")
                    print('='*60)
                    
                    for section_name, (original_lines, modified_lines) in changes.items():
                        print(f"\n--- {section_name} ---")
                        diff_output = updater.format_diff(original_lines, modified_lines, section_name)
                        print(diff_output)
                    # Optionally include XML-level diffs
                    if args.xml_diff:
                        xml_changes = updater.get_document_xml_changes_preview(file_path)
                        # Filter sections if requested
                        items = xml_changes.items()
                        if args.xml_diff_sections:
                            items = [(k, v) for k, v in items if k in set(args.xml_diff_sections)]
                        for section_name, (original_lines, modified_lines) in items:
                            print(f"\n--- {section_name} ---")
                            diff_output = updater.format_diff(original_lines, modified_lines, section_name)
                            print(diff_output)
                else:
                    print(f"no changes: {file_path}")
            else:
                if updater.modify_docx(file_path):
                    print(f"✓ {file_path}")
                else:
                    print(f"- {file_path} (no changes)")
        except Exception as e:
            print(f"✗ {file_path}: {e}", file=sys.stderr)