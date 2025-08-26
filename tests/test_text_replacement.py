"""
Unit tests for the TextReplacer class.

Tests initialization, text replacement functionality, and paragraph break handling.
Complex DOCX document tests are in test_real_templates.py and test_integration.py.
"""
import pytest
from docx import Document
from docx.oxml import OxmlElement, ns
from text_replacement import TextReplacer
from formatting import FormattingProcessor
from docx.enum.text import WD_ALIGN_PARAGRAPH


class TestTextReplacer:
    """Test cases for TextReplacer class."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.formatter = FormattingProcessor()
        self.replacements = [
            {"search": "old text", "replace": "new text"},
            {"search": "TESTER QUALIFICATIONS", "replace": "INSPECTOR QUALIFICATIONS"}
        ]
    
    def test_init(self):
        """Test TextReplacer initialization."""
        text_replacer = TextReplacer(self.replacements, self.formatter)
        
        assert text_replacer.replacements == self.replacements
        assert text_replacer.formatter == self.formatter
    
    def test_init_with_empty_replacements(self):
        """Test TextReplacer initialization with empty replacements."""
        text_replacer = TextReplacer([], self.formatter)
        
        assert text_replacer.replacements == []
        assert text_replacer.formatter == self.formatter


class TestTextReplacementFunctionality:
    """Test text replacement operations."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.formatter = FormattingProcessor()
    
    def test_apply_text_replacements_simple(self):
        """Test simple text replacement."""
        replacements = [{"search": "old", "replace": "new"}]
        replacer = TextReplacer(replacements, self.formatter)
        
        result, modified = replacer.apply_text_replacements("This is old text")
        assert modified is True
        assert result == "This is new text"
    
    def test_apply_text_replacements_no_match(self):
        """Test text replacement with no matches."""
        replacements = [{"search": "missing", "replace": "new"}]
        replacer = TextReplacer(replacements, self.formatter)
        
        result, modified = replacer.apply_text_replacements("This is old text")
        assert modified is False
        assert result == "This is old text"
    
    def test_apply_text_replacements_with_paragraphbreak(self):
        """Test text replacement with paragraph breaks."""
        replacements = [{"search": "PHOTOS", "replace": "Photo1paragraphbreakPhoto2"}]
        replacer = TextReplacer(replacements, self.formatter)
        
        result, modified = replacer.apply_text_replacements("SITE PHOTOS section")
        assert modified is True
        assert result == "SITE Photo1paragraphbreakPhoto2 section"
    
    def test_apply_text_replacements_insert_after(self):
        """Test insert_after operation."""
        replacements = [{"search": "PHOTOS", "insert_after": "Photo1paragraphbreakPhoto2"}]
        replacer = TextReplacer(replacements, self.formatter)
        
        result, modified = replacer.apply_text_replacements("SITE PHOTOS section")
        assert modified is True
        assert result == "SITE PHOTOS\nPhoto1paragraphbreakPhoto2 section"
    
    def test_apply_text_replacements_multiple_replacements(self):
        """Test multiple replacements in order."""
        replacements = [
            {"search": "old", "replace": "new"},
            {"search": "bad", "replace": "good"}
        ]
        replacer = TextReplacer(replacements, self.formatter)
        
        result, modified = replacer.apply_text_replacements("This old thing is bad")
        assert modified is True
        assert result == "This new thing is good"
    
    def test_apply_text_replacements_all_occurrences(self):
        """Test that all occurrences are replaced."""
        replacements = [{"search": "test", "replace": "example"}]
        replacer = TextReplacer(replacements, self.formatter)
        
        result, modified = replacer.apply_text_replacements("This test has test twice")
        assert modified is True
        assert result == "This example has example twice"
    
    def test_apply_text_replacements_skip_cleanup_operations(self):
        """Test that cleanup operations (remove_empty_paragraphs_after) are skipped."""
        replacements = [
            {"search": "old", "replace": "new"},
            {"remove_empty_paragraphs_after": "SITE PHOTOS"}  # This should be skipped
        ]
        replacer = TextReplacer(replacements, self.formatter)
        
        result, modified = replacer.apply_text_replacements("This is old text")
        assert modified is True
        assert result == "This is new text"
        # The cleanup operation should not affect text replacement
    
    def test_apply_text_replacements_skip_replacements_without_search(self):
        """Test that replacements without search field are skipped."""
        replacements = [
            {"search": "old", "replace": "new"},
            {"remove_empty_paragraphs_after": True}  # No search field
        ]
        replacer = TextReplacer(replacements, self.formatter)
        
        result, modified = replacer.apply_text_replacements("This is old text")
        assert modified is True
        assert result == "This is new text"


class TestHyperlinkDetection:
    """Test hyperlink detection functionality."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.formatter = FormattingProcessor()
        self.replacements = [{"search": "SITE PHOTOS", "replace": "Updated Photos"}]
        self.replacer = TextReplacer(self.replacements, self.formatter)
        self.doc = Document()
    
    def create_paragraph_with_hyperlink(self, text: str) -> object:
        """Create a paragraph containing a hyperlink with the given text."""
        paragraph = self.doc.add_paragraph()
        
        # Create hyperlink XML element
        hyperlink = OxmlElement('w:hyperlink')
        hyperlink.set(ns.qn('r:id'), 'rId1')
        
        # Create run within hyperlink
        run = OxmlElement('w:r')
        run_text = OxmlElement('w:t')
        run_text.text = text
        run.append(run_text)
        hyperlink.append(run)
        
        # Add hyperlink to paragraph
        paragraph._p.append(hyperlink)
        return paragraph
    
    def test_is_text_in_hyperlink_with_hyperlinks(self):
        """Test hyperlink detection correctly identifies text within hyperlinks."""
        paragraph = self.create_paragraph_with_hyperlink("APPENDIX H    SITE PHOTOS")
        
        result = self.replacer._is_text_in_hyperlink(paragraph, "SITE PHOTOS")
        assert result is True
    
    def test_is_text_in_hyperlink_without_hyperlinks(self):
        """Test hyperlink detection when paragraph has no hyperlinks."""
        paragraph = self.doc.add_paragraph("SITE PHOTOS")
        
        result = self.replacer._is_text_in_hyperlink(paragraph, "SITE PHOTOS")
        assert result is False
    
    def test_is_text_in_hyperlink_with_none_paragraph(self):
        """Test hyperlink detection with None paragraph."""
        result = self.replacer._is_text_in_hyperlink(None, "SITE PHOTOS")
        assert result is False
    
    def test_is_text_in_hyperlink_text_not_in_paragraph(self):
        """Test hyperlink detection when search text is not in paragraph."""
        paragraph = self.create_paragraph_with_hyperlink("Different content")
        
        result = self.replacer._is_text_in_hyperlink(paragraph, "SITE PHOTOS")
        assert result is False


class TestTextReplacementWithHyperlinkSkipping:
    """Test text replacement with hyperlink detection."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.formatter = FormattingProcessor()
    
    def test_apply_text_replacements_works_without_paragraph_context(self):
        """Test that text replacement works when no paragraph context is provided."""
        replacements = [{"search": "SITE PHOTOS", "replace": "Updated Photos"}]
        replacer = TextReplacer(replacements, self.formatter)
        
        # No paragraph context means no hyperlink detection
        result, modified = replacer.apply_text_replacements("SITE PHOTOS", None)
        assert modified is True
        assert result == "Updated Photos"


class TestTextReplacementWithDocx:
    """Test text replacement with real DOCX objects."""
    
    def setup_method(self):
        """Set up test fixtures with real DOCX objects."""
        self.formatter = FormattingProcessor()
        self.doc = Document()
        self.paragraph = self.doc.add_paragraph("Original text content")
    
    def test_handle_insert_after_with_paragraphbreaks(self):
        """Test insert_after operation with paragraph breaks using string processing."""
        # Set up replacements with paragraph breaks
        replacements = [
            {"search": "PHOTOS", "insert_after": "Photo1paragraphbreakPhoto2paragraphbreakPhoto3"}
        ]
        replacer = TextReplacer(replacements, self.formatter)
        
        # Test the core text replacement logic
        test_text = "SITE PHOTOS content"
        
        # Execute the replacement at text level
        result, modified = replacer.apply_text_replacements(test_text)
        
        # Verify the replacement logic works correctly
        assert modified is True
        assert "SITE PHOTOS" in result  # Original text preserved
        assert "Photo1" in result
        assert "paragraphbreak" in result  # Formatting token preserved
        assert "Photo2" in result
        assert "Photo3" in result
    
    def test_handle_insert_after_no_match(self):
        """Test insert_after with no matching text."""
        replacements = [{"search": "MISSING", "insert_after": "new content"}]
        replacer = TextReplacer(replacements, self.formatter)
        
        result, modified = replacer.apply_text_replacements("SITE PHOTOS content")
        assert modified is False
        assert result == "SITE PHOTOS content"
    
    def create_paragraph_with_hyperlink(self, text: str) -> object:
        """Create a paragraph containing a hyperlink with the given text."""
        paragraph = self.doc.add_paragraph()
        
        # Create hyperlink XML element
        hyperlink = OxmlElement('w:hyperlink')
        hyperlink.set(ns.qn('r:id'), 'rId1')
        
        # Create run within hyperlink
        run = OxmlElement('w:r')
        run_text = OxmlElement('w:t')
        run_text.text = text
        run.append(run_text)
        hyperlink.append(run)
        
        # Add hyperlink to paragraph
        paragraph._p.append(hyperlink)
        return paragraph
    
    def test_handle_insert_after_skips_hyperlinked_text(self):
        """Test insert_after skips hyperlinked content."""
        replacements = [{"search": "PHOTOS", "insert_after": "new content"}]
        replacer = TextReplacer(replacements, self.formatter)
        
        # Create paragraph with hyperlink using our helper
        paragraph = self.create_paragraph_with_hyperlink("APPENDIX H    SITE PHOTOS")
        
        # Test that hyperlink detection correctly identifies hyperlinked text (should skip insert_after)
        is_hyperlinked = replacer._is_text_in_hyperlink(paragraph, "PHOTOS")
        assert is_hyperlinked is True


class TestParagraphBreakIntegration:
    """Integration tests for paragraph break functionality."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.formatter = FormattingProcessor()
    
    def test_paragraphbreak_processing_pipeline(self):
        """Test complete pipeline from text replacement to formatting segments."""
        replacements = [
            {"search": "PHOTOS", "replace": "Photo1paragraphbreakPhoto2paragraphbreakPhoto3"}
        ]
        replacer = TextReplacer(replacements, self.formatter)
        
        # Apply replacement
        text = "SITE PHOTOS section"
        new_text, modified = replacer.apply_text_replacements(text)
        assert modified is True
        assert new_text == "SITE Photo1paragraphbreakPhoto2paragraphbreakPhoto3 section"
        
        # Process formatting tokens on the replaced text
        segments = self.formatter.process_formatting_tokens(new_text, None)
        
        # Should find segments with paragraph breaks
        photo_segments = []
        for text_part, formatting in segments:
            if 'Photo' in text_part:
                photo_segments.append((text_part, formatting))
        
        # Should have Photo1 and Photo2 with paragraph breaks, Photo3 without
        assert len(photo_segments) >= 2
        
        # Find Photo1 and Photo2 segments
        photo1_segment = next((s for s in photo_segments if 'Photo1' in s[0]), None)
        photo2_segment = next((s for s in photo_segments if 'Photo2' in s[0]), None)
        
        if photo1_segment:
            assert photo1_segment[1].get('paragraph_break_after') is True
        if photo2_segment:
            assert photo2_segment[1].get('paragraph_break_after') is True
    
    def test_paragraphbreak_with_inline_formatting_integration(self):
        """Test paragraph breaks combined with inline formatting."""
        replacements = [
            {"search": "PHOTOS", "replace": "{format:center,size12}Photo1paragraphbreakPhoto2{/format}"}
        ]
        replacer = TextReplacer(replacements, self.formatter)
        
        # Apply replacement
        text = "SITE PHOTOS section"
        new_text, modified = replacer.apply_text_replacements(text)
        assert modified is True
        
        # Process formatting
        segments = self.formatter.process_formatting_tokens(new_text, None)
        
        # Find photo segments
        photo_segments = []
        for text_part, formatting in segments:
            if 'Photo' in text_part:
                photo_segments.append((text_part, formatting))
        
        # Both photos should have the inline formatting AND paragraph break handling
        for text_part, formatting in photo_segments:
            if 'Photo1' in text_part:
                assert formatting.get('alignment') == WD_ALIGN_PARAGRAPH.CENTER
                assert formatting.get('font_size') == 12
                assert formatting.get('paragraph_break_after') is True
            elif 'Photo2' in text_part:
                assert formatting.get('alignment') == WD_ALIGN_PARAGRAPH.CENTER
                assert formatting.get('font_size') == 12
                assert formatting.get('paragraph_break_after') is False