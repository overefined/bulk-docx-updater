"""
Text replacement logic for DOCX documents.

Handles complex text replacement across DOCX runs while preserving formatting,
including alignment changes that require creating new paragraphs.
"""
from __future__ import annotations
import re
from typing import List, Dict, Optional, Tuple
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph

from formatting import FormattingProcessor


class TextReplacer:
    """Handles text replacement operations in DOCX paragraphs."""
    
    def __init__(self, replacements: List[Dict[str, str]], formatting_processor: FormattingProcessor):
        self.replacements = replacements
        self.formatter = formatting_processor
    
    def replace_text_across_paragraphs(self, paragraphs: List[Paragraph]) -> bool:
        """Handle text replacement across multiple consecutive paragraphs."""
        if not paragraphs:
            return False
        
        # Find which paragraphs actually contain parts of the search patterns
        for replacement in self.replacements:
            if 'search' not in replacement:
                continue
            if not ('replace' in replacement or 'insert_after' in replacement):
                continue
                
            search_text = replacement['search']
            
            # Find paragraphs that together contain the complete search text
            combined_text = "".join(para.text for para in paragraphs)
            
            if search_text not in combined_text:
                continue
            
            # Check if this pattern actually spans paragraphs
            spans_paragraphs = True
            for para in paragraphs:
                if search_text in para.text:
                    spans_paragraphs = False
                    break
            
            if not spans_paragraphs:
                continue  # Let single-paragraph processing handle this
            
            # Find the exact paragraphs involved in this pattern
            # Strategy: Find the first paragraph that starts the pattern and the last that completes it
            
            # Find where the pattern starts
            start_para_idx = None
            for i, para in enumerate(paragraphs):
                # Check if this paragraph contains the beginning of our search text
                para_text = para.text
                if para_text and search_text.startswith(para_text[:50]):  # Check first 50 chars
                    start_para_idx = i
                    break
                # Also check for partial matches at the end of paragraph
                for j in range(1, min(len(para_text), len(search_text)) + 1):
                    if search_text.startswith(para_text[-j:]):
                        start_para_idx = i
                        break
                if start_para_idx is not None:
                    break
            
            if start_para_idx is None:
                continue  # Couldn't find start of pattern
            
            # Now find consecutive paragraphs until we have the complete pattern
            affected_paragraphs = []
            accumulated_text = ""
            
            for i in range(start_para_idx, len(paragraphs)):
                accumulated_text += paragraphs[i].text
                affected_paragraphs.append(i)
                
                # Check if we now have the complete search pattern
                if search_text in accumulated_text:
                    break
            
            if not affected_paragraphs:
                continue
            
            # Combine text only from affected paragraphs
            affected_combined_text = "".join(paragraphs[i].text for i in affected_paragraphs)
            
            # Apply the replacement
            new_text, modified = self.apply_text_replacements(affected_combined_text)
            
            if not modified:
                continue
            
            # Put the new text in the first affected paragraph
            first_para_idx = affected_paragraphs[0]
            first_paragraph = paragraphs[first_para_idx]
            self._rebuild_paragraph_with_text(first_paragraph, new_text)
            
            # Clear the remaining affected paragraphs
            for para_idx in affected_paragraphs[1:]:
                self._clear_paragraph(paragraphs[para_idx])
            
            return True
        
        return False
    
    def _rebuild_paragraph_with_text(self, paragraph: Paragraph, new_text: str, preserve_advanced_formatting: bool = False):
        """Rebuild paragraph with new text while preserving formatting.
        
        Args:
            paragraph: The paragraph to rebuild
            new_text: The new text content
            preserve_advanced_formatting: If True, uses advanced formatting preservation (for single-paragraph replacements)
        """
        if preserve_advanced_formatting:
            self._rebuild_paragraph_advanced(paragraph, new_text)
        else:
            self._rebuild_paragraph_basic(paragraph, new_text)
    
    def _rebuild_paragraph_basic(self, paragraph: Paragraph, new_text: str):
        """Basic paragraph rebuilding for cross-paragraph replacements."""
        # Store original formatting from first run if available
        original_font_formatting = {}
        if paragraph.runs:
            first_run = paragraph.runs[0]
            if hasattr(first_run, 'font'):
                original_font_formatting['font_name'] = first_run.font.name
                original_font_formatting['font_size'] = first_run.font.size
                original_font_formatting['bold'] = first_run.bold
                original_font_formatting['italic'] = first_run.italic
        
        # Clear all runs
        self._clear_paragraph(paragraph)
        
        # Process formatting tokens in the new text
        text_segments = self.formatter.process_formatting_tokens(new_text, paragraph)
        
        # Add runs with the new text and formatting
        for text, formatting in text_segments:
            if text:  # Only create runs for non-empty text
                run = paragraph.add_run(text)
                
                # Apply original formatting as base
                if original_font_formatting.get('font_name'):
                    run.font.name = original_font_formatting['font_name']
                if original_font_formatting.get('font_size'):
                    run.font.size = original_font_formatting['font_size']
                if original_font_formatting.get('bold'):
                    run.bold = original_font_formatting['bold']
                if original_font_formatting.get('italic'):
                    run.italic = original_font_formatting['italic']
                
                # Apply new formatting from tokens
                self.formatter.apply_formatting_to_run(run, formatting, paragraph)
    
    def _rebuild_paragraph_advanced(self, paragraph: Paragraph, new_text: str):
        """Advanced paragraph rebuilding with sophisticated formatting preservation."""
        # Store original run formatting and detect leading whitespace BEFORE clearing runs
        original_runs = list(paragraph.runs)
        original_formatting = []
        leading_whitespace_runs = []
        
        for run in original_runs:
            formatting = {
                'font_name': run.font.name,
                'font_size': run.font.size,
                'bold': run.font.bold,
                'italic': run.font.italic,
                'underline': run.font.underline
            }
            original_formatting.append(formatting)
        
        # Check if the original text had leading newlines/whitespace that we need to preserve
        for run in original_runs:
            # Check if this run contains only whitespace characters
            if run.text and all(c in '\n \t' for c in run.text):
                # This run contains only whitespace - preserve it
                leading_whitespace_runs.append(run.text)
            else:
                # Found first non-whitespace run, stop looking for leading whitespace
                break
        
        # Clear all runs
        for run in original_runs:
            run.text = ''
        
        # Remove all but the first run
        while len(paragraph.runs) > 1:
            last_run = paragraph.runs[-1]
            last_run._element.getparent().remove(last_run._element)
        
        # Process the new text for formatting tokens and create new runs
        text_segments = self.formatter.process_formatting_tokens(new_text, paragraph)
        
        # Check if any segment has alignment formatting or paragraph breaks
        has_alignment_segments = any(seg_formatting.get('alignment') for _, seg_formatting in text_segments)
        has_paragraph_breaks = any(seg_formatting.get('paragraph_break_after') for _, seg_formatting in text_segments)
        
        if has_alignment_segments or has_paragraph_breaks:
            # Handle alignment and paragraph breaks by creating separate paragraphs
            self._handle_alignment_segments(paragraph, text_segments, original_runs[0] if original_runs else None, leading_whitespace_runs, original_formatting)
        else:
            # No alignment, rebuild runs normally
            self._apply_text_segments_to_paragraph(paragraph, text_segments, original_formatting, leading_whitespace_runs)
    
    def _apply_text_segments_to_paragraph(self, paragraph: Paragraph, text_segments, original_formatting, leading_whitespace_runs):
        """Apply text segments to paragraph with sophisticated formatting preservation."""
        # Get the first run to work with
        first_run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
        
        # First, add back any leading whitespace runs
        current_run = first_run
        for i, whitespace_text in enumerate(leading_whitespace_runs):
            if i == 0:
                # Use the first run for the first whitespace
                current_run.text = whitespace_text
            else:
                # Create new runs for additional whitespace
                current_run = paragraph.add_run(whitespace_text)
            
            # Apply original formatting if available
            if i < len(original_formatting):
                base_formatting = original_formatting[i]
                if base_formatting['font_name']:
                    current_run.font.name = base_formatting['font_name']
                if base_formatting['font_size']:
                    current_run.font.size = base_formatting['font_size']
                current_run.font.bold = base_formatting['bold']
                current_run.font.italic = base_formatting['italic']
                current_run.font.underline = base_formatting['underline']
        
        # Apply the text segments after the whitespace runs
        if text_segments:
            # Determine which run to use for the first text segment
            if leading_whitespace_runs:
                # If we have whitespace runs, create a new run for the first text segment
                first_text_run = paragraph.add_run()
                current_run = first_text_run
            else:
                # No whitespace runs, use the first run
                first_text_run = first_run
                current_run = first_text_run
            
            first_text, first_formatting = text_segments[0]
            first_text_run.text = first_text
            
            # Apply original formatting as base (use formatting from after whitespace)
            base_formatting_idx = len(leading_whitespace_runs)
            if base_formatting_idx < len(original_formatting):
                base_formatting = original_formatting[base_formatting_idx]
                if base_formatting['font_name']:
                    first_text_run.font.name = base_formatting['font_name']
                if base_formatting['font_size']:
                    first_text_run.font.size = base_formatting['font_size']
                first_text_run.font.bold = base_formatting['bold']
                first_text_run.font.italic = base_formatting['italic']
                first_text_run.font.underline = base_formatting['underline']
            
            # Apply segment-specific formatting
            if first_formatting:
                self.formatter.apply_formatting_to_run(first_text_run, first_formatting, paragraph)
            
            # Create additional runs for remaining segments
            for i, (segment_text, segment_formatting) in enumerate(text_segments[1:], 1):
                if segment_text:
                    new_run = paragraph.add_run(segment_text)
                    
                    # Apply base formatting from original runs if available
                    base_idx = min(base_formatting_idx + i, len(original_formatting) - 1)
                    if base_idx < len(original_formatting):
                        base_formatting = original_formatting[base_idx]
                        if base_formatting['font_name']:
                            new_run.font.name = base_formatting['font_name']
                        if base_formatting['font_size']:
                            new_run.font.size = base_formatting['font_size']
                        new_run.font.bold = base_formatting['bold']
                        new_run.font.italic = base_formatting['italic']
                        new_run.font.underline = base_formatting['underline']
                    
                    # Apply segment-specific formatting
                    if segment_formatting:
                        self.formatter.apply_formatting_to_run(new_run, segment_formatting, paragraph)
    
    def _clear_paragraph(self, paragraph: Paragraph):
        """Clear all runs from a paragraph."""
        for run in paragraph.runs:
            run._element.getparent().remove(run._element)

    
    def _has_page_break_in_run(self, run) -> bool:
        """Check if a run contains a page break."""
        # Check for hard page breaks (manual page breaks) in the XML
        return 'w:br' in run._element.xml and 'type="page"' in run._element.xml
    
    def _detect_page_breaks_after_text(self, paragraph, search_text: str) -> bool:
        """Check if there's a page break after specific text in a paragraph."""
        full_text = paragraph.text
        if search_text not in full_text:
            return False
        
        # Find the position of the search text
        search_pos = full_text.find(search_text)
        search_end = search_pos + len(search_text)
        
        # Map character positions to runs to find which runs come after the search text
        char_pos = 0
        for i, run in enumerate(paragraph.runs):
            run_start = char_pos
            run_end = char_pos + len(run.text)
            
            # If the search text ends within this run or before next run
            if run_start <= search_end <= run_end:
                # Check this run and subsequent runs for page breaks
                for j in range(i, len(paragraph.runs)):
                    if self._has_page_break_in_run(paragraph.runs[j]):
                        return True
                break
            
            char_pos = run_end
        
        return False
    

    def _is_text_in_hyperlink(self, paragraph, search_text: str) -> bool:
        """Check if the search text is within a hyperlink in the paragraph (coarse check)."""
        if paragraph is None:
            return False
        if 'hyperlink' not in paragraph._p.xml.lower():
            return False
        try:
            import xml.etree.ElementTree as ET
            xml_str = paragraph._p.xml
            root = ET.fromstring(xml_str)
            hyperlinks = root.findall('.//w:hyperlink', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
            for hyperlink in hyperlinks:
                text_elements = hyperlink.findall('.//w:t', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                hyperlink_text = ''.join(elem.text or '' for elem in text_elements)
                if search_text in hyperlink_text:
                    return True
        except Exception:
            pass
        return False

    def _compute_hyperlink_ranges(self, paragraph: Paragraph) -> List[Tuple[int, int]]:
        """Compute [start, end) character ranges of paragraph.text that are inside hyperlinks."""
        ranges: List[Tuple[int, int]] = []
        try:
            p_el = paragraph._p
            if 'hyperlink' not in p_el.xml.lower():
                return ranges
            ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

            def sum_text_len(el) -> int:
                total_len = 0
                # XPath over lxml element to gather all w:t descendants
                for t in el.xpath('.//w:t', namespaces=ns):
                    total_len += len(t.text or '')
                return total_len

            current_index = 0
            for child in p_el.iterchildren():
                tag = str(child.tag)
                if tag.endswith('}hyperlink'):
                    hl_len = sum_text_len(child)
                    if hl_len > 0:
                        ranges.append((current_index, current_index + hl_len))
                    current_index += hl_len
                else:
                    current_index += sum_text_len(child)
        except Exception:
            return []
        return ranges

    def _match_overlaps_hyperlink(self, paragraph: Optional[Paragraph], pattern_text: str, match_start: int, match_end: int) -> bool:
        """Check if a specific match span overlaps a hyperlink span in paragraph."""
        if paragraph is None:
            return False
        ranges = self._compute_hyperlink_ranges(paragraph)
        if not ranges:
            return False
        for hs, he in ranges:
            if match_start < he and match_end > hs:
                return True
        return False

    def _replace_text_in_hyperlinks(self, paragraph) -> bool:
        """Replace text within hyperlink elements while preserving XML structure (tabs, formatting, etc)."""
        import xml.etree.ElementTree as ET
        
        modified = False
        xml_str = paragraph._p.xml
        
        # Create a filtered replacements list with only 'replace' operations
        replace_only_replacements = [r for r in self.replacements 
                                   if 'search' in r and 'replace' in r and 'insert_after' not in r]
        
        if not replace_only_replacements:
            return False
        
        try:
            root = ET.fromstring(xml_str)
            
            # Find hyperlink elements
            hyperlinks = root.findall('.//w:hyperlink', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
            
            for hyperlink in hyperlinks:
                # Reconstruct text from all text elements to get full hyperlink content
                text_elements = hyperlink.findall('.//w:t', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                full_text = ''.join(elem.text or '' for elem in text_elements)
                
                # Apply replacements to the full text
                original_replacements = self.replacements
                self.replacements = replace_only_replacements
                new_full_text, text_modified = self.apply_text_replacements(full_text, None)
                self.replacements = original_replacements
                
                if text_modified:
                    # Replace text content while preserving XML structure
                    # Find which text elements need to be updated
                    old_text_parts = [elem.text or '' for elem in text_elements]
                    
                    # Apply the same replacements to each text part individually
                    new_text_parts = []
                    for part in old_text_parts:
                        if part:  # Only process non-empty parts
                            original_replacements = self.replacements
                            self.replacements = replace_only_replacements
                            new_part, _ = self.apply_text_replacements(part, None)
                            self.replacements = original_replacements
                            new_text_parts.append(new_part)
                        else:
                            new_text_parts.append(part)
                    
                    # Update the text elements with the new text parts
                    for elem, new_part in zip(text_elements, new_text_parts):
                        elem.text = new_part
                    
                    modified = True
            
            if modified:
                # Replace paragraph XML with updated XML
                new_xml_str = ET.tostring(root, encoding='unicode')
                # Parse and replace the paragraph element
                from docx.oxml import parse_xml
                new_p_element = parse_xml(new_xml_str)
                old_p_element = paragraph._p
                old_p_element.getparent().replace(old_p_element, new_p_element)
                
        except Exception as e:
            # If XML processing fails, fall back to normal text replacement
            pass
            
        return modified

    def apply_text_replacements(self, text: str, paragraph=None) -> tuple[str, bool]:
        """Apply text replacements to a string. Returns (new_text, modified)."""
        new_text = text
        modified = False
        
        # Apply all replacements to the full text
        for replacement in self.replacements:
            # Skip replacements that are cleanup actions (remove_empty_paragraphs_after)
            if 'search' not in replacement:
                continue
            if not ('replace' in replacement or 'insert_after' in replacement):
                continue
                
            search_text = replacement['search']
            use_regex = bool(replacement.get('regex'))
            ignore_case = bool(replacement.get('ignore_case'))
            flags = re.IGNORECASE if ignore_case else 0
            pattern = re.compile(search_text if use_regex else re.escape(search_text), flags)
            
            # Handle insert_after operation
            if 'insert_after' in replacement:
                insert_text = replacement['insert_after']
                
                # Find all matches and choose the first one that's not in a hyperlink
                matches = list(pattern.finditer(new_text))
                for match in matches:
                    start, end = match.start(), match.end()
                    # Check if this match is in a hyperlink (if paragraph available)
                    if paragraph and self._match_overlaps_hyperlink(paragraph, search_text, start, end):
                        continue  # Skip this match, try the next one
                    
                    # Process this match - add line break before inserted content for proper separation
                    new_text = new_text[:end] + '\n' + insert_text + new_text[end:]
                    modified = True
                    break  # Only process the first non-appendix-list match
                    
                continue
            
            # Handle regular replace operation
            replace_text = replacement['replace']
            
            # Check if there's an existing page break after this search text (if paragraph provided)
            if paragraph is not None:
                has_existing_pagebreak = self._detect_page_breaks_after_text(paragraph, search_text)
                
                # If there's an existing page break and the replacement doesn't already include "pagebreak",
                # automatically add it to preserve the existing break
                if has_existing_pagebreak and 'pagebreak' not in replace_text.lower():
                    replace_text = replace_text + 'pagebreak'
            
            # Replace all non-hyperlink matches
            matches = list(pattern.finditer(new_text))
            replacements_made = 0
            
            # Process matches in reverse order to avoid offset issues
            for match in reversed(matches):
                start, end = match.start(), match.end()
                # Check if this match is in a hyperlink (if paragraph available)
                if paragraph and self._match_overlaps_hyperlink(paragraph, search_text, start, end):
                    continue  # Skip this match, try the next one
                
                # Replace this match
                new_text = new_text[:start] + replace_text + new_text[end:]
                replacements_made += 1
            
            if replacements_made > 0:
                modified = True
                
        return new_text, modified

    def _handle_insert_after_in_paragraph(self, paragraph, full_text: str) -> bool:
        """Handle insert_after operations while preserving existing paragraph structure."""
        modified = False
        
        # Process each insert_after replacement
        for replacement in self.replacements:
            if 'search' not in replacement or 'insert_after' not in replacement:
                continue
                
            search_text = replacement['search']
            insert_text = replacement['insert_after']
            
            if search_text not in full_text:
                continue
                
            # Find first non-hyperlink match for insert_after operation
            matches = list(re.finditer(re.escape(search_text), full_text))
            if not matches:
                continue
            
            # Find the first match that's not in a hyperlink
            target_match = None
            for match in matches:
                start, end = match.start(), match.end()
                if not self._is_text_in_hyperlink(paragraph, search_text):
                    target_match = match
                    break
            
            if target_match is None:
                continue  # All matches were in hyperlinks
                
            # Create a new paragraph after the current one for the inserted content
            parent = paragraph._element.getparent()
            next_paragraph = paragraph._element.getnext()
            
            # Process formatting tokens in the insert text
            temp_paragraph = paragraph  # Use current paragraph for context
            text_segments = self.formatter.process_formatting_tokens(insert_text, temp_paragraph)
            
            # Create new paragraph element
            from docx.oxml import parse_xml
            from docx.oxml.ns import nsdecls, qn
            new_p_xml = f'<w:p {nsdecls("w")}></w:p>'
            new_p_element = parse_xml(new_p_xml)
            
            # Insert the new paragraph after current paragraph
            if next_paragraph is not None:
                parent.insert(parent.index(paragraph._element) + 1, new_p_element)
            else:
                parent.append(new_p_element)
                
            # Create paragraph object from the element
            from docx.text.paragraph import Paragraph
            new_paragraph = Paragraph(new_p_element, parent)
            
            # Add runs to the new paragraph based on formatted segments, creating new paragraphs for paragraph breaks
            current_paragraph = new_paragraph
            
            # Get the most common font from the document
            original_font_formatting = {}
            try:
                # Find the most commonly used font in the document
                doc = paragraph._parent
                font_counter = {}
                for para in doc.paragraphs:
                    for run in para.runs:
                        if run.text.strip() and run.font.name is not None:
                            font_name = run.font.name
                            font_counter[font_name] = font_counter.get(font_name, 0) + 1
                
                # Use the most common font
                if font_counter:
                    most_common_font = max(font_counter, key=font_counter.get)
                    original_font_formatting['font_name'] = most_common_font
            except:
                pass  # If we can't determine the font, just continue without it
            
            for i, (text, formatting) in enumerate(text_segments):
                if text or formatting.get('page_break_after') or formatting.get('line_break_after') or formatting.get('paragraph_break_after'):
                    # Create run with text
                    run = current_paragraph.add_run(text)
                    
                    # Apply the most common font from the document (unless a specific font is specified)
                    if 'font_name' in original_font_formatting and not formatting.get('font_name'):
                        run.font.name = original_font_formatting['font_name']
                    
                    # Then apply any specific formatting from the replacement text
                    self.formatter.apply_formatting_to_run(run, formatting, current_paragraph)
                    
                    # Apply paragraph-level formatting
                    if formatting:
                        self.formatter.apply_paragraph_formatting(current_paragraph, formatting)
                    
                    # If this segment has a paragraph break, create a new paragraph for the next segment
                    if formatting.get('paragraph_break_after') and i < len(text_segments) - 1:
                        # Create another new paragraph element
                        from docx.oxml import parse_xml
                        from docx.oxml.ns import nsdecls
                        new_p_xml = f'<w:p {nsdecls("w")}></w:p>'
                        next_p_element = parse_xml(new_p_xml)
                        
                        # Insert after the current paragraph
                        parent.insert(parent.index(current_paragraph._element) + 1, next_p_element)
                        
                        # Update current paragraph reference
                        from docx.text.paragraph import Paragraph
                        current_paragraph = Paragraph(next_p_element, parent)
                    
            modified = True
            break  # Only process the first insert_after match
            
        return modified

    def replace_text_in_paragraph(self, paragraph) -> bool:
        """Replace text in a paragraph, handling splits across runs while preserving formatting."""
        # Check if paragraph contains hyperlinks that need text replacement
        has_hyperlinks = 'hyperlink' in paragraph._p.xml.lower()
        if has_hyperlinks:
            # Handle hyperlink text replacement directly in XML
            modified = self._replace_text_in_hyperlinks(paragraph)
            if modified:
                return True
        
        # Get the original full text
        full_text = paragraph.text
        
        # Check if any replacement for this paragraph is insert_after
        has_insert_after = any('insert_after' in repl for repl in self.replacements if 'search' in repl and repl['search'] in full_text)
        
        if has_insert_after:
            # Handle insert_after operations without rebuilding paragraph structure
            return self._handle_insert_after_in_paragraph(paragraph, full_text)
        
        # Apply text replacements using the extracted method
        new_text, modified = self.apply_text_replacements(full_text, paragraph)
        
        if not modified:
            return False
        
        # Now we need to rebuild the paragraph with the new text
        # but preserve formatting where possible
        
        # Use unified paragraph rebuilding with advanced formatting preservation
        self._rebuild_paragraph_with_text(paragraph, new_text, preserve_advanced_formatting=True)
        
        return True

    def _handle_alignment_segments(self, paragraph, new_run_data: List[Tuple[str, Dict]], 
                                 original_run, leading_whitespace_runs: Optional[List[str]] = None, 
                                 original_formatting: Optional[List[Dict]] = None):
        """Handle segments with different alignments or paragraph breaks by creating separate paragraphs if needed."""
        # Group segments by alignment and paragraph breaks
        segment_groups = []
        current_group = []
        current_alignment = None
        
        for text_segment, segment_formatting in new_run_data:
            seg_alignment = segment_formatting.get('alignment')
            
            # Add segment to current group
            current_group.append((text_segment, segment_formatting))
            
            # Check if this segment forces a new paragraph (alignment change or paragraph break)
            alignment_changed = seg_alignment != current_alignment and seg_alignment is not None
            has_paragraph_break = segment_formatting.get('paragraph_break_after', False)
            
            if alignment_changed or has_paragraph_break:
                # Finalize current group
                segment_groups.append((current_alignment, current_group))
                current_group = []
                current_alignment = seg_alignment
            elif seg_alignment is not None:
                # Update current alignment without breaking
                current_alignment = seg_alignment
        
        if current_group:
            segment_groups.append((current_alignment, current_group))
        
        # Handle the first group in the existing paragraph
        if segment_groups:
            first_alignment, first_group = segment_groups[0]
            
            # First, add any leading whitespace runs
            if leading_whitespace_runs and original_formatting:
                for i, whitespace_text in enumerate(leading_whitespace_runs):
                    if i == 0 and original_run is not None:
                        # Use the original run for the first whitespace
                        original_run.text = whitespace_text
                        # Apply original formatting
                        if i < len(original_formatting):
                            base_formatting = original_formatting[i]
                            if base_formatting['font_name']:
                                original_run.font.name = base_formatting['font_name']
                            if base_formatting['font_size']:
                                original_run.font.size = base_formatting['font_size']
                            original_run.font.bold = base_formatting['bold']
                            original_run.font.italic = base_formatting['italic']
                            original_run.font.underline = base_formatting['underline']
                    else:
                        # Create new runs for additional whitespace
                        ws_run = paragraph.add_run(whitespace_text)
                        if i < len(original_formatting):
                            base_formatting = original_formatting[i]
                            if base_formatting['font_name']:
                                ws_run.font.name = base_formatting['font_name']
                            if base_formatting['font_size']:
                                ws_run.font.size = base_formatting['font_size']
                            ws_run.font.bold = base_formatting['bold']
                            ws_run.font.italic = base_formatting['italic']
                            ws_run.font.underline = base_formatting['underline']
            else:
                # No leading whitespace, clear the original run
                if original_run is not None:
                    original_run.text = ''
            
            # Add the text segments
            for i, (text_segment, segment_formatting) in enumerate(first_group):
                if text_segment:  # Only add non-empty text
                    if not leading_whitespace_runs and i == 0 and original_run is not None:
                        # No whitespace runs, use original run for first text segment
                        original_run.text = text_segment
                        if segment_formatting:
                            self.formatter.apply_formatting_to_run(original_run, segment_formatting, paragraph)
                    else:
                        # Create new run for text segment (either because we have whitespace runs, not first segment, or original_run is None)
                        new_run = paragraph.add_run(text_segment) 
                        # Apply base formatting
                        base_idx = len(leading_whitespace_runs or []) + i
                        if original_formatting and base_idx < len(original_formatting):
                            base_formatting = original_formatting[base_idx]
                            if base_formatting['font_name']:
                                new_run.font.name = base_formatting['font_name']
                            if base_formatting['font_size']:
                                new_run.font.size = base_formatting['font_size']
                            new_run.font.bold = base_formatting['bold']
                            new_run.font.italic = base_formatting['italic']
                            new_run.font.underline = base_formatting['underline']
                        
                        if segment_formatting:
                            self.formatter.apply_formatting_to_run(new_run, segment_formatting, paragraph)
            
            # Apply alignment to first paragraph
            if first_alignment:
                paragraph.alignment = first_alignment
            
            # Create new paragraphs for remaining groups
            for alignment, group in segment_groups[1:]:
                # Create new paragraph after current one
                new_p_el = OxmlElement("w:p")
                paragraph._p.addnext(new_p_el)
                new_paragraph = Paragraph(new_p_el, paragraph._parent)
                
                # Set alignment for new paragraph
                if alignment:
                    new_paragraph.alignment = alignment
                
                # Add runs to new paragraph
                for text_segment, segment_formatting in group:
                    if text_segment:
                        new_run = new_paragraph.add_run(text_segment)
                        # Copy base formatting if original_run exists
                        if original_run is not None:
                            new_run.font.name = original_run.font.name
                            new_run.font.size = original_run.font.size
                            new_run.font.bold = original_run.font.bold
                            new_run.font.italic = original_run.font.italic
                            new_run.font.underline = original_run.font.underline
                        
                        if segment_formatting:
                            self.formatter.apply_formatting_to_run(new_run, segment_formatting, new_paragraph)