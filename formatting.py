"""
Formatting token processing and application for DOCX documents.

Handles parsing and application of:
- Inline formatting blocks: {format:bold,center,size16}text{/format}
- Global formatting tokens: pagebreak, linebreak, paragraphbreak
- Font properties, alignment, spacing
"""
from __future__ import annotations
import re
from typing import List, Dict, Tuple
from docx.shared import Pt
from docx.enum.text import WD_BREAK, WD_ALIGN_PARAGRAPH


class FormattingProcessor:
    """Handles all formatting token processing and application."""
    
    def process_formatting_tokens(self, text: str, paragraph) -> List[Tuple[str, Dict]]:
        """Process special formatting tokens and return list of (text, formatting) segments."""
        # First, handle inline formatting blocks like {format:bold,size14}text{/format}
        segments = self._parse_inline_formatting(text)
        
        # Then process each segment for other tokens like pagebreak and paragraphbreak
        final_segments = []
        for segment_text, segment_formatting in segments:
            # Handle multiple types of breaks in order of precedence
            if any(token in segment_text.lower() for token in ['pagebreak', 'paragraphbreak', 'linebreak']):
                # Split on all break types, preserving the break tokens
                parts = re.split(r'(pagebreak|paragraphbreak|linebreak)', segment_text, flags=re.IGNORECASE)
                
                for i, part in enumerate(parts):
                    part_lower = part.lower()
                    if part_lower == 'pagebreak':
                        # Handle pagebreak - add to previous segment or create empty segment
                        if final_segments:
                            final_segments[-1][1]['page_break_after'] = True
                        else:
                            final_segments.append(["", {'page_break_after': True}])
                    elif part_lower == 'paragraphbreak':
                        # Handle paragraphbreak - add to previous segment or create empty segment
                        if final_segments:
                            final_segments[-1][1]['paragraph_break_after'] = True
                        else:
                            final_segments.append(["", {'paragraph_break_after': True}])
                    elif part_lower == 'linebreak':
                        # Handle linebreak - add to previous segment or create empty segment
                        if final_segments:
                            final_segments[-1][1]['line_break_after'] = True
                        else:
                            final_segments.append(["", {'line_break_after': True}])
                    elif part.strip():  # Non-empty text
                        part_formatting = self._extract_formatting(part)
                        # Merge with segment formatting (segment formatting takes precedence)
                        merged_formatting = {**part_formatting, **segment_formatting}
                        clean_text = self._clean_formatting_tokens(part)
                        if clean_text.strip():
                            final_segments.append([clean_text, merged_formatting])
            else:
                # No breaks, process normally
                part_formatting = self._extract_formatting(segment_text)
                # Merge with segment formatting (segment formatting takes precedence)
                merged_formatting = {**part_formatting, **segment_formatting}
                clean_text = self._clean_formatting_tokens(segment_text)
                if clean_text.strip():
                    final_segments.append([clean_text, merged_formatting])
        
        return final_segments if final_segments else [["", {}]]
    
    def _parse_inline_formatting(self, text: str) -> List[Tuple[str, Dict]]:
        """Parse inline formatting blocks like {format:bold,center}text{/format}."""
        segments = []
        
        # Pattern to match {format:options}text{/format}
        pattern = r'\{format:([^}]+)\}(.*?)\{/format\}'
        
        last_end = 0
        for match in re.finditer(pattern, text):
            # Add text before this formatted block
            before_text = text[last_end:match.start()]
            if before_text:
                segments.append([before_text, {}])
            
            # Parse formatting options
            format_options = match.group(1)
            formatted_text = match.group(2)
            
            formatting = self._parse_format_options(format_options)
            segments.append([formatted_text, formatting])
            
            last_end = match.end()
        
        # Add remaining text after last match
        if last_end < len(text):
            remaining = text[last_end:]
            if remaining:
                segments.append([remaining, {}])
        
        # If no inline formatting found, return the whole text
        if not segments:
            segments = [[text, {}]]
            
        return segments
    
    def _parse_format_options(self, options_str: str) -> Dict:
        """Parse format options string like 'bold,center,size14' into formatting dict."""
        formatting = {
            'line_break_after': False,
            'paragraph_break_after': False,
            'page_break_after': False,
            # Explicitly set all formatting properties to defaults - inline formatting overrides everything
            'bold': False,
            'italic': False,
            'underline': False,
            'font_name': None
        }
        
        # Split options and preserve case for font names
        raw_options = [opt.strip() for opt in options_str.split(',')]
        options = []
        for opt in raw_options:
            if opt.lower().startswith('font:'):
                # Preserve case for font names
                options.append(opt)
            else:
                # Convert to lowercase for other options
                options.append(opt.lower())
        
        for option in options:
            if option == 'bold':
                formatting['bold'] = True
            elif option == 'italic':
                formatting['italic'] = True
            elif option == 'underline':
                formatting['underline'] = True
            elif option == 'center':
                formatting['center'] = True
                formatting['alignment'] = WD_ALIGN_PARAGRAPH.CENTER
            elif option == 'left':
                formatting['alignment'] = WD_ALIGN_PARAGRAPH.LEFT
            elif option == 'right':
                formatting['alignment'] = WD_ALIGN_PARAGRAPH.RIGHT
            elif option == 'justify':
                formatting['alignment'] = WD_ALIGN_PARAGRAPH.JUSTIFY
            elif option.startswith('size'):
                size_match = re.search(r'size(\d+)', option)
                if size_match:
                    formatting['font_size'] = int(size_match.group(1))
            elif option.startswith('spaceafter'):
                space_match = re.search(r'spaceafter(\d+)', option)
                if space_match:
                    formatting['space_after'] = int(space_match.group(1))
            elif option.startswith('spacebefore'):
                space_match = re.search(r'spacebefore(\d+)', option)
                if space_match:
                    formatting['space_before'] = int(space_match.group(1))
            elif option.startswith('font'):
                # Extract font name after 'font:' - e.g., 'font:Arial Narrow'
                font_match = re.search(r'font:(.+)', option)
                if font_match:
                    formatting['font_name'] = font_match.group(1).strip()
        
        return formatting
    
    def _extract_formatting(self, text: str) -> Dict:
        """Extract formatting information from text (global tokens only)."""
        formatting = {
            'line_break_after': False,
            'paragraph_break_after': False,
            'alignment': None,
            'font_size': None,
            'bold': None,
            'italic': None,
            'underline': None,
            'center': False,
            'space_after': None,
            'space_before': None,
            'page_break_after': False
        }
        
        text_lower = text.lower()
        
        # Only handle global break tokens - no legacy formatting
        if "linebreak" in text_lower:
            formatting['line_break_after'] = True
            
        if "paragraphbreak" in text_lower:
            formatting['paragraph_break_after'] = True
        
        if "remove_empty_paragraphs" in text_lower:
            formatting['remove_empty_paragraphs'] = True
        
        return formatting
    
    def _clean_formatting_tokens(self, text: str) -> str:
        """Remove global formatting tokens from text."""
        # Remove only global break tokens - no legacy formatting tokens
        text = re.sub(r'pagebreak', '', text, flags=re.IGNORECASE)
        text = re.sub(r'linebreak', '', text, flags=re.IGNORECASE)
        text = re.sub(r'paragraphbreak', '', text, flags=re.IGNORECASE)
        
        # Clean up extra spaces
        text = re.sub(r'\s+', ' ', text).strip()
        
        return text

    def apply_formatting_to_run(self, run, formatting: Dict, paragraph):
        """Apply formatting options to a run and paragraph."""
        # Apply font formatting to the run
        if formatting.get('font_name'):
            run.font.name = formatting['font_name']
        if formatting.get('font_size'):
            run.font.size = Pt(formatting['font_size'])
        if formatting.get('bold') is not None:
            run.font.bold = formatting['bold']
        if formatting.get('italic') is not None:
            run.font.italic = formatting['italic']
        if formatting.get('underline') is not None:
            run.font.underline = formatting['underline']
        
        # Apply breaks after the run - these need to be in separate runs
        if formatting.get('line_break_after'):
            br_run = paragraph.add_run()
            br_run.add_break(WD_BREAK.LINE)
        if formatting.get('paragraph_break_after'):
            # For paragraph breaks, we need to signal to create a new paragraph
            # This will be handled at a higher level in text_replacement.py
            pass
        if formatting.get('page_break_after'):
            br_run = paragraph.add_run()
            br_run.add_break(WD_BREAK.PAGE)
    
    def apply_paragraph_formatting(self, paragraph, formatting: Dict):
        """Apply paragraph-level formatting like alignment and spacing."""
        if formatting.get('alignment'):
            paragraph.alignment = formatting['alignment']
        if formatting.get('space_after'):
            paragraph.paragraph_format.space_after = Pt(formatting['space_after'])
        if formatting.get('space_before'):
            paragraph.paragraph_format.space_before = Pt(formatting['space_before'])
    
