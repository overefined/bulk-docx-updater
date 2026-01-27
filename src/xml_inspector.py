"""
XML Inspector for DOCX files - helps debug text matching issues by showing raw XML content
"""

import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
import re
from typing import Dict, List, Optional


class DocxXmlInspector:
    """Inspector for examining raw XML content in DOCX files"""
    
    def __init__(self, docx_path: str):
        self.docx_path = Path(docx_path)
        if not self.docx_path.exists():
            raise FileNotFoundError(f"DOCX file not found: {docx_path}")
    
    def extract_document_xml(self) -> str:
        """Extract the main document XML content"""
        with zipfile.ZipFile(self.docx_path, 'r') as docx_zip:
            try:
                return docx_zip.read('word/document.xml').decode('utf-8')
            except KeyError:
                raise ValueError(f"Invalid DOCX file: {self.docx_path}")
    
    def format_xml_pretty(self, xml_content: str) -> str:
        """Format XML content for readable display"""
        try:
            root = ET.fromstring(xml_content)
            ET.indent(root, space="  ")
            return ET.tostring(root, encoding='unicode')
        except ET.ParseError:
            return xml_content
    
    def find_text_in_xml(self, search_text: str, context_lines: int = 3) -> List[Dict]:
        """Find occurrences of text in XML and return with context"""
        xml_content = self.extract_document_xml()
        lines = xml_content.split('\n')
        matches = []
        
        for i, line in enumerate(lines):
            if search_text in line:
                start = max(0, i - context_lines)
                end = min(len(lines), i + context_lines + 1)
                context = lines[start:end]
                
                matches.append({
                    'line_number': i + 1,
                    'context': context,
                    'context_start_line': start + 1
                })
        
        return matches
    
    def show_paragraph_structure_around_text(self, search_text: str) -> str:
        """Show paragraph XML structure around matching text"""
        xml_content = self.extract_document_xml()
        
        # Find paragraphs containing the text
        paragraph_pattern = r'<w:p[^>]*>.*?</w:p>'
        paragraphs = re.findall(paragraph_pattern, xml_content, re.DOTALL)
        
        matching_paragraphs = []
        for i, para in enumerate(paragraphs):
            if search_text in para:
                matching_paragraphs.append({
                    'paragraph_index': i,
                    'xml_content': para
                })
        
        return matching_paragraphs
    
    def extract_all_text_runs(self) -> List[Dict]:
        """Extract all text runs with their XML context"""
        xml_content = self.extract_document_xml()
        
        # Find all w:t elements (text runs)
        text_run_pattern = r'<w:t[^>]*>(.*?)</w:t>'
        matches = re.finditer(text_run_pattern, xml_content, re.DOTALL)
        
        runs = []
        for match in matches:
            runs.append({
                'text_content': match.group(1),
                'xml_snippet': match.group(0),
                'start_pos': match.start(),
                'end_pos': match.end()
            })
        
        return runs
    
    def inspect_text_pattern(self, pattern: str, show_context: bool = True) -> Dict:
        """Comprehensive inspection of how a text pattern appears in the XML"""
        result = {
            'pattern': pattern,
            'xml_matches': self.find_text_in_xml(pattern),
            'paragraph_matches': self.show_paragraph_structure_around_text(pattern),
            'text_runs': []
        }
        
        # Check if pattern spans across multiple text runs
        all_runs = self.extract_all_text_runs()
        consecutive_text = ""
        run_group = []
        
        for run in all_runs:
            consecutive_text += run['text_content']
            run_group.append(run)
            
            if pattern in consecutive_text:
                result['text_runs'].append({
                    'matching_runs': run_group.copy(),
                    'combined_text': consecutive_text
                })
                consecutive_text = ""
                run_group = []
            elif len(consecutive_text) > len(pattern) * 2:
                # Reset if we've gone too far without a match
                consecutive_text = run['text_content']
                run_group = [run]
        
        return result


def inspect_docx_xml(docx_path: str, pattern: Optional[str] = None, 
                     show_full_xml: bool = False, context_lines: int = 5):
    """Main function to inspect DOCX XML content"""
    inspector = DocxXmlInspector(docx_path)
    
    print(f"Inspecting DOCX file: {docx_path}")
    print("=" * 60)
    
    if show_full_xml:
        xml_content = inspector.extract_document_xml()
        pretty_xml = inspector.format_xml_pretty(xml_content)
        print("Full XML Content:")
        print(pretty_xml)
        return
    
    if pattern:
        print(f"Searching for pattern: '{pattern}'")
        print("-" * 40)
        
        inspection = inspector.inspect_text_pattern(pattern)
        
        if inspection['xml_matches']:
            print(f"\nFound {len(inspection['xml_matches'])} direct XML matches:")
            for i, match in enumerate(inspection['xml_matches']):
                print(f"\nMatch {i+1} at line {match['line_number']}:")
                for j, line in enumerate(match['context']):
                    line_num = match['context_start_line'] + j
                    marker = ">>>" if pattern in line else "   "
                    print(f"{marker} {line_num:4d}: {line}")
        
        if inspection['paragraph_matches']:
            print(f"\nFound in {len(inspection['paragraph_matches'])} paragraphs:")
            for i, para in enumerate(inspection['paragraph_matches']):
                print(f"\nParagraph {para['paragraph_index']}:")
                formatted_xml = inspector.format_xml_pretty(para['xml_content'])
                print(formatted_xml)
        
        if inspection['text_runs']:
            print(f"\nFound in {len(inspection['text_runs'])} text run sequences:")
            for i, run_seq in enumerate(inspection['text_runs']):
                print(f"\nRun sequence {i+1}:")
                print(f"Combined text: '{run_seq['combined_text']}'")
                for j, run in enumerate(run_seq['matching_runs']):
                    print(f"  Run {j+1}: '{run['text_content']}' -> {run['xml_snippet']}")
        
        if not any([inspection['xml_matches'], inspection['paragraph_matches'], inspection['text_runs']]):
            print("No matches found for the specified pattern.")
            print("\nTip: The pattern might be split across multiple XML runs.")
            print("Try searching for smaller parts of the pattern.")
    
    else:
        print("No pattern specified. Use --pattern to search for specific text.")
        print("Use --show-xml to display the full XML content.")


if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description="Inspect DOCX XML content for debugging text replacement issues")
    parser.add_argument("docx_file", help="Path to DOCX file to inspect")
    parser.add_argument("--pattern", "-p", help="Text pattern to search for in XML")
    parser.add_argument("--show-xml", action="store_true", help="Display full formatted XML content")
    parser.add_argument("--context", "-c", type=int, default=5, help="Number of context lines around matches")
    
    args = parser.parse_args()
    
    try:
        inspect_docx_xml(args.docx_file, args.pattern, args.show_xml, args.context)
    except Exception as e:
        print(f"Error: {e}")