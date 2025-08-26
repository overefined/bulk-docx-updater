"""
Integration tests specifically for paragraph break functionality.

Tests the complete pipeline from configuration to DOCX processing
with paragraph break tokens.
"""
import tempfile
import json
from pathlib import Path

from document_processor import DocxBulkUpdater
from config import load_replacements_from_json, validate_replacements


class TestParagraphBreakEndToEnd:
    """End-to-end tests for paragraph break functionality."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.temp_dir = Path(tempfile.mkdtemp())
        
    def teardown_method(self):
        """Clean up test fixtures."""
        import shutil
        shutil.rmtree(self.temp_dir)
    
    def test_paragraphbreak_config_loading_and_validation(self):
        """Test that paragraph break configurations load and validate correctly."""
        config_data = {
            "replacements": [
                {
                    "search": "SITE PHOTOS",
                    "insert_after": "pagebreak{format:center,size12}Photo1paragraphbreakPhoto2paragraphbreakPhoto3{/format}"
                },
                {
                    "search": "TEST RESULTS",
                    "replace": "Result1paragraphbreakResult2paragraphbreakResult3"
                }
            ]
        }
        
        # Write config file
        config_file = self.temp_dir / "test_config.json"
        with open(config_file, 'w') as f:
            json.dump(config_data, f)
        
        # Load and validate
        replacements = load_replacements_from_json(config_file)
        validate_replacements(replacements)  # Should not raise
        
        assert len(replacements) == 2
        assert replacements[0]["search"] == "SITE PHOTOS"
        assert "paragraphbreak" in replacements[0]["insert_after"]
        assert replacements[1]["search"] == "TEST RESULTS"
        assert "paragraphbreak" in replacements[1]["replace"]
    
    def test_docx_bulk_updater_with_paragraphbreaks(self):
        """Test DocxBulkUpdater initialization with paragraph break replacements."""
        replacements = [
            {
                "search": "PHOTOS",
                "insert_after": "Photo1paragraphbreakPhoto2"
            }
        ]
        
        updater = DocxBulkUpdater(replacements)
        assert updater.text_replacer is not None
        assert updater.text_replacer.replacements == replacements
        assert updater.text_replacer.formatter is not None
    
    def test_paragraph_break_token_processing_integration(self):
        """Test integration between text replacement and formatting for paragraph breaks."""
        replacements = [
            {
                "search": "PHOTOS",
                "replace": "Photo1paragraphbreakPhoto2paragraphbreakPhoto3"
            }
        ]
        
        updater = DocxBulkUpdater(replacements)
        
        # Test text replacement
        original_text = "SITE PHOTOS content"
        new_text, modified = updater.text_replacer.apply_text_replacements(original_text)
        assert modified is True
        assert "paragraphbreak" in new_text
        
        # Test formatting processing
        segments = updater.text_replacer.formatter.process_formatting_tokens(new_text, None)
        
        # Should have segments for "SITE", "Photo1", "Photo2", "Photo3", "content"
        photo_segments = [seg for seg in segments if 'Photo' in seg[0]]
        assert len(photo_segments) >= 2  # At least Photo1 and Photo2
        
        # Photo1 and Photo2 should have paragraph breaks
        for text_part, formatting in photo_segments:
            if 'Photo1' in text_part or 'Photo2' in text_part:
                assert formatting.get('paragraph_break_after') is True
            elif 'Photo3' in text_part:
                assert formatting.get('paragraph_break_after') is False
    
    def test_multiple_break_types_integration(self):
        """Test integration of multiple break types (paragraph, line, page)."""
        replacements = [
            {
                "search": "CONTENT",
                "replace": "Line1linebreakLine2paragraphbreakPage1pagebreakPage2"
            }
        ]
        
        updater = DocxBulkUpdater(replacements)
        
        # Test text replacement
        original_text = "TEST CONTENT here"
        new_text, modified = updater.text_replacer.apply_text_replacements(original_text)
        assert modified is True
        
        # Test formatting processing
        segments = updater.text_replacer.formatter.process_formatting_tokens(new_text, None)
        
        # Check that different break types are processed correctly
        line_break_found = False
        paragraph_break_found = False
        page_break_found = False
        
        for text_part, formatting in segments:
            if formatting.get('line_break_after'):
                line_break_found = True
            if formatting.get('paragraph_break_after'):
                paragraph_break_found = True
            if formatting.get('page_break_after'):
                page_break_found = True
        
        assert line_break_found, "Line break should be found"
        assert paragraph_break_found, "Paragraph break should be found"  
        assert page_break_found, "Page break should be found"
    
    def test_paragraph_breaks_with_inline_formatting_integration(self):
        """Test paragraph breaks combined with inline formatting in full pipeline."""
        replacements = [
            {
                "search": "PHOTOS",
                "insert_after": "{format:center,bold,size14}Photo1paragraphbreakPhoto2{/format}paragraphbreakPhoto3"
            }
        ]
        
        updater = DocxBulkUpdater(replacements)
        
        # Test text replacement
        original_text = "SITE PHOTOS section"
        new_text, modified = updater.text_replacer.apply_text_replacements(original_text)
        assert modified is True
        
        # Test formatting processing  
        segments = updater.text_replacer.formatter.process_formatting_tokens(new_text, None)
        
        # Find the photo segments
        photo_segments = []
        for text_part, formatting in segments:
            if 'Photo' in text_part:
                photo_segments.append((text_part, formatting))
        
        # Verify formatting and breaks are correctly applied
        for text_part, formatting in photo_segments:
            if 'Photo1' in text_part:
                # Photo1 should have inline formatting AND paragraph break
                assert formatting.get('alignment') is not None  # center alignment
                assert formatting.get('bold') is True
                assert formatting.get('font_size') == 14
                assert formatting.get('paragraph_break_after') is True
            elif 'Photo2' in text_part:
                # Photo2 should have inline formatting AND paragraph break
                assert formatting.get('alignment') is not None
                assert formatting.get('bold') is True
                assert formatting.get('font_size') == 14
                assert formatting.get('paragraph_break_after') is True
            elif 'Photo3' in text_part:
                # Photo3 should have paragraph break but no inline formatting
                assert formatting.get('bold') is None or formatting.get('bold') is False
                assert formatting.get('paragraph_break_after') is False
    
    def test_configuration_edge_cases(self):
        """Test edge cases in configuration handling."""
        # Test empty paragraph break
        replacements = [{"search": "TEST", "replace": "contentparagraphbreak"}]
        updater = DocxBulkUpdater(replacements)
        
        text = "This is TEST content"
        new_text, modified = updater.text_replacer.apply_text_replacements(text)
        assert modified is True
        
        segments = updater.text_replacer.formatter.process_formatting_tokens(new_text, None)
        
        # Should handle trailing paragraph break correctly
        found_break = False
        for text_part, formatting in segments:
            if formatting.get('paragraph_break_after'):
                found_break = True
                break
        
        assert found_break, "Should handle trailing paragraph break"
    
    def test_jinja_template_with_paragraphbreaks(self):
        """Test configuration that matches actual use case with Jinja templates."""
        replacements = [
            {
                "search": "SITE PHOTOS",
                "insert_after": "pagebreak{format:center,size12}{% if ecom_photos != none %}{% for site_photo in ecom_photos.site_photos %}{{ site_photo }}paragraphbreak{% endfor %}{% endif %}{/format}"
            }
        ]
        
        updater = DocxBulkUpdater(replacements)
        
        # Test the Jinja template content processing
        text = "APPENDIX H	SITE PHOTOS"
        new_text, modified = updater.text_replacer.apply_text_replacements(text)
        assert modified is True
        assert "pagebreak" in new_text
        assert "paragraphbreak" in new_text
        assert "{format:" in new_text
        assert "{/format}" in new_text
        
        # Test formatting processing of the template
        segments = updater.text_replacer.formatter.process_formatting_tokens(new_text, None)
        
        # Should have page break and inline formatting
        page_break_found = False
        paragraph_break_found = False
        inline_formatting_found = False
        
        for text_part, formatting in segments:
            if formatting.get('page_break_after'):
                page_break_found = True
            if formatting.get('paragraph_break_after'):
                paragraph_break_found = True
            if formatting.get('alignment') is not None or formatting.get('font_size') is not None:
                inline_formatting_found = True
        
        assert page_break_found, "Should find page break"
        assert paragraph_break_found, "Should find paragraph break"
        assert inline_formatting_found, "Should find inline formatting"