"""
Font formatting utilities for DOCX documents.

Provides common font operations to reduce code duplication across modules.
"""
from __future__ import annotations
from typing import Dict, Optional, Any
from docx.shared import Pt


class FontFormatter:
    """Utility class for applying font formatting to DOCX runs."""
    
    @staticmethod
    def extract_font_properties(run) -> Dict[str, Any]:
        """Extract font properties from a run."""
        if not hasattr(run, 'font'):
            return {}
        
        return {
            'font_name': run.font.name,
            'font_size': run.font.size,
            'bold': run.font.bold,
            'italic': run.font.italic,
            'underline': run.font.underline
        }
    
    @staticmethod
    def apply_font_properties(run, properties: Dict[str, Any]):
        """Apply font properties to a run."""
        if not hasattr(run, 'font') or not properties:
            return
        
        if properties.get('font_name'):
            run.font.name = properties['font_name']
        if properties.get('font_size'):
            run.font.size = properties['font_size']
        if properties.get('bold') is not None:
            run.font.bold = properties['bold']
        if properties.get('italic') is not None:
            run.font.italic = properties['italic']
        if properties.get('underline') is not None:
            run.font.underline = properties['underline']
    
    @staticmethod
    def copy_font_formatting(source_run, target_run):
        """Copy font formatting from source run to target run."""
        properties = FontFormatter.extract_font_properties(source_run)
        FontFormatter.apply_font_properties(target_run, properties)
    
    @staticmethod
    def get_base_font_formatting(runs) -> Dict[str, Any]:
        """Get base font formatting from the first available run."""
        for run in runs:
            if hasattr(run, 'font'):
                return FontFormatter.extract_font_properties(run)
        return {}
    
    @staticmethod
    def find_most_common_font(document) -> Optional[str]:
        """Find the most commonly used font in a document."""
        try:
            font_counter = {}
            for para in document.paragraphs:
                for run in para.runs:
                    if run.text.strip() and run.font.name is not None:
                        font_name = run.font.name
                        font_counter[font_name] = font_counter.get(font_name, 0) + 1
            
            if font_counter:
                return max(font_counter, key=font_counter.get)
        except:
            pass
        return None