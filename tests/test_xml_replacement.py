"""
Test XML replacement functionality
"""
import pytest
import json
import tempfile
import os
from unittest.mock import Mock, MagicMock
from pathlib import Path

from src.text_replacement import TextReplacer
from src.formatting import FormattingProcessor
from src.config import validate_replacements, validate_operations, load_operations_from_json, _process_file_references


class TestXMLReplacementConfig:
    """Test XML replacement configuration validation."""

    def test_xml_mode_boolean_validation(self):
        """Test that xml_mode must be a boolean."""
        operations = [
            {
                "op": "xml_replace",
                "search": "test",
                "replace": "result",
                "xml_mode": "true"  # Should be boolean, not string
            }
        ]

        with pytest.raises(SystemExit):
            validate_replacements(operations)

    def test_xml_mode_requires_search_replace(self):
        """Test that xml_mode can only be used with search/replace operations."""
        operations = [
            {
                "xml_mode": True  # Should fail - requires search and replace
            }
        ]

        with pytest.raises(SystemExit):
            validate_replacements(operations)

    def test_valid_xml_mode_config(self):
        """Test valid XML mode configuration (inline XML allowed by validator)."""
        operations = [
            {
                "op": "xml_replace",
                "search": "<w:t>old</w:t>",
                "replace": "<w:t>new</w:t>"
            },
            {
                "op": "xml_replace",
                "search": "normal text",
                "replace": "new text"
            }
        ]

        # Should not raise any exceptions
        validate_replacements(operations)

    def test_regex_validation(self):
        """Test regex option validation."""
        operations = [
            {
                "op": "xml_replace",
                "search": "test",
                "replace": "result",
                "regex": "true"  # Should be boolean
            }
        ]

        with pytest.raises(SystemExit):
            validate_replacements(operations)


class TestXMLReplacement:
    """Test XML replacement functionality in TextReplacer."""

    def setup_method(self):
        """Setup test fixtures."""
        self.formatting_processor = Mock()
        self.formatting_processor.process_formatting_tokens.return_value = [("test", {})]

    def test_xml_mode_filtering(self):
        """Test that XML mode replacements are filtered from regular text processing."""
        operations = [
            {
                "op": "replace",
                "search": "text_search",
                "replace": "text_replace"
            },
            {
                "op": "xml_replace",
                "search": "<w:t>xml_search</w:t>",
                "replace": "<w:t>xml_replace</w:t>"
            }
        ]

        replacer = TextReplacer(operations, self.formatting_processor)

        # Test regular text replacement - should not process XML mode replacements
        result, modified = replacer.apply_text_replacements("text_search and <w:t>xml_search</w:t>")

        # Only the text replacement should be applied
        assert "text_replace" in result
        assert "<w:t>xml_search</w:t>" in result  # XML search should remain unchanged
        assert modified

    def test_xml_replacement_basic(self):
        """Test basic XML replacement functionality."""
        operations = [
            {
                "op": "xml_replace",
                "search": "<w:t>old</w:t>",
                "replace": "<w:t>new</w:t>"
            }
        ]

        replacer = TextReplacer(operations, self.formatting_processor)

        # Create mock paragraph with proper namespace XML
        mock_paragraph = Mock()
        mock_p_element = Mock()
        # Use proper DOCX namespace XML
        mock_p_element.xml = '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>old</w:t></w:r></w:p>'
        mock_paragraph._p = mock_p_element

        # Mock getparent and replace methods
        mock_parent = Mock()
        mock_p_element.getparent.return_value = mock_parent
        mock_parent.replace = Mock()

        # Test XML replacement
        result = replacer._replace_xml_in_paragraph(mock_paragraph)

        # Should have been modified
        assert result is True
        # Replace should have been called
        mock_parent.replace.assert_called_once()

    def test_xml_replacement_literal_attribute(self):
        """Test XML replacement with literal attribute patterns (no regex)."""
        operations = [
            {
                "op": "xml_replace",
                "search": 'w:val="240"',
                "replace": 'w:val="new_value"'
            }
        ]

        replacer = TextReplacer(operations, self.formatting_processor)

        mock_paragraph = Mock()
        mock_p_element = Mock()
        mock_p_element.xml = '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:spacing w:val="240"/><w:sz w:val="12"/></w:p>'
        mock_paragraph._p = mock_p_element

        mock_parent = Mock()
        mock_p_element.getparent.return_value = mock_parent
        mock_parent.replace = Mock()

        result = replacer._replace_xml_in_paragraph(mock_paragraph)
        assert result is True

    def test_xml_replacement_literal_case_sensitive(self):
        """Test XML replacement with exact literal casing (no ignore_case)."""
        operations = [
            {
                "op": "xml_replace",
                "search": "<w:t>test</w:t>",
                "replace": "<w:t>replaced</w:t>"
            }
        ]

        replacer = TextReplacer(operations, self.formatting_processor)

        mock_paragraph = Mock()
        mock_p_element = Mock()
        mock_p_element.xml = '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>test</w:t></w:r></w:p>'
        mock_paragraph._p = mock_p_element

        mock_parent = Mock()
        mock_p_element.getparent.return_value = mock_parent
        mock_parent.replace = Mock()

        result = replacer._replace_xml_in_paragraph(mock_paragraph)
        assert result is True

    def test_xml_replacement_malformed_handling(self):
        """Test that malformed XML replacements are rejected."""
        operations = [
            {
                "op": "xml_replace",
                "search": "<w:t>",
                "replace": "malformed",  # This would break XML structure
                "xml_mode": True
            }
        ]

        replacer = TextReplacer(operations, self.formatting_processor)

        # Create mock paragraph
        mock_paragraph = Mock()
        mock_p_element = Mock()
        mock_p_element.xml = '<w:p><w:r><w:t>content</w:t></w:r></w:p>'
        mock_paragraph._p = mock_p_element

        # Test XML replacement - should not modify due to malformed result
        result = replacer._replace_xml_in_paragraph(mock_paragraph)

        # Should not have been modified due to malformed XML
        assert result is False

    def test_xml_replacement_no_matches(self):
        """Test XML replacement when no patterns match."""
        operations = [
            {
                "op": "xml_replace",
                "search": "<w:nonexistent>",
                "replace": "<w:replacement>"
            }
        ]

        replacer = TextReplacer(operations, self.formatting_processor)

        # Create mock paragraph
        mock_paragraph = Mock()
        mock_p_element = Mock()
        mock_p_element.xml = '<w:p><w:r><w:t>content</w:t></w:r></w:p>'
        mock_paragraph._p = mock_p_element

        # Test XML replacement
        result = replacer._replace_xml_in_paragraph(mock_paragraph)

        # Should not have been modified
        assert result is False


class TestXMLReplacementIntegration:
    """Test XML replacement integration with paragraph processing."""

    def setup_method(self):
        """Setup test fixtures."""
        self.formatting_processor = Mock()
        self.formatting_processor.process_formatting_tokens.return_value = [("test", {})]

    def test_xml_replacement_precedence(self):
        """Test that XML replacements take precedence over text replacements."""
        operations = [
            {
                "op": "xml_replace",
                "search": "content",
                "replace": "text_replaced"
            },
            {
                "op": "xml_replace",
                "search": "<w:t>content</w:t>",
                "replace": "<w:t>xml_replaced</w:t>"
            }
        ]

        replacer = TextReplacer(operations, self.formatting_processor)

        # Create mock paragraph with proper XML namespace
        mock_paragraph = Mock()
        mock_p_element = Mock()
        mock_p_element.xml = '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>content</w:t></w:r></w:p>'
        mock_paragraph._p = mock_p_element
        mock_paragraph.text = "content"
        mock_paragraph.runs = []  # Mock empty runs list to avoid iteration error

        # Mock getparent and replace methods
        mock_parent = Mock()
        mock_p_element.getparent.return_value = mock_parent
        mock_parent.replace = Mock()

        # Mock cache behavior
        replacer._text_cache = {}
        replacer._xml_cache = {}
        replacer._paragraph_has_page_breaks_cache = {}

        # Test paragraph replacement - XML should take precedence
        result = replacer.replace_text_in_paragraph(mock_paragraph)

        # Should have been modified by XML replacement
        assert result is True
        # XML replacement should have been called
        mock_parent.replace.assert_called_once()


class TestXMLFileReferences:
    """Test file-based XML replacement configuration."""

    def test_process_file_references_search_and_replace_files(self):
        """Test loading XML content from search_file and replace_file."""
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_dir = Path(temp_dir)

            # Create test XML files
            search_file = temp_dir / "search.xml"
            replace_file = temp_dir / "replace.xml"

            search_content = '<w:t xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">old text</w:t>'
            replace_content = '<w:t xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">new text</w:t>'

            search_file.write_text(search_content, encoding='utf-8')
            replace_file.write_text(replace_content, encoding='utf-8')

            # Test processing file references
            replacement = {
                "search_file": "search.xml",
                "replace_file": "replace.xml"
            }

            result = _process_file_references(replacement, temp_dir)

            assert "search" in result
            assert "replace" in result
            assert result["search"] == search_content
            assert result["replace"] == replace_content
            assert "search_file" not in result
            assert "replace_file" not in result

    def test_process_file_references_mixed_file_and_direct(self):
        """Test mixing file references and direct content."""
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_dir = Path(temp_dir)

            # Create test XML file
            search_file = temp_dir / "search.xml"
            search_content = '<w:t>file content</w:t>'
            search_file.write_text(search_content, encoding='utf-8')

            # Test mixing file reference with direct content
            replacement = {
                "search_file": "search.xml",
                "replace": "<w:t>direct content</w:t>"
            }

            result = _process_file_references(replacement, temp_dir)

            assert result["search"] == search_content
            assert result["replace"] == "<w:t>direct content</w:t>"
            assert "search_file" not in result

    def test_load_replacements_with_file_references(self):
        """Test loading replacements config with file references."""
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_dir = Path(temp_dir)

            # Create XML content files
            search_file = temp_dir / "search_pattern.xml"
            replace_file = temp_dir / "replace_pattern.xml"

            search_content = """<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                <w:r>
                    <w:t>Large XML block with "quotes" and complex structure</w:t>
                </w:r>
            </w:p>"""

            replace_content = """<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                <w:r>
                    <w:rPr><w:b/></w:rPr>
                    <w:t>REPLACED XML with "preserved quotes" and bold formatting</w:t>
                </w:r>
            </w:p>"""

            search_file.write_text(search_content, encoding='utf-8')
            replace_file.write_text(replace_content, encoding='utf-8')

            # Create config file
            config_file = temp_dir / "config.json"
            config_data = [
                {
                    "search_file": "search_pattern.xml",
                    "replace_file": "replace_pattern.xml",
                    "xml_mode": True
                },
                {
                    "search": "regular text",
                    "replace": "replacement text"
                }
            ]

            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(config_data, f)

            # Load and test
            replacements = load_operations_from_json(config_file)

            assert len(replacements) == 2

            # First replacement should have loaded XML content
            xml_replacement = replacements[0]
            assert "search" in xml_replacement
            assert "replace" in xml_replacement
            assert xml_replacement["op"] == "xml_replace"
            assert "Large XML block with" in xml_replacement["search"]
            assert "REPLACED XML with" in xml_replacement["replace"]
            assert "search_file" not in xml_replacement
            assert "replace_file" not in xml_replacement

            # Second replacement should be unchanged
            text_replacement = replacements[1]
            assert text_replacement["search"] == "regular text"
            assert text_replacement["replace"] == "replacement text"

    def test_file_reference_validation(self):
        """Test validation of file-based replacement configurations."""
        # Valid file-based config
        valid_operations = [
            {
                "op": "xml_replace",
                "search_file": "search.xml",
                "replace_file": "replace.xml"
            }
        ]

        # Should not raise any exceptions
        validate_operations(valid_operations)

        # Mixed file/direct config
        mixed_operations = [
            {
                "op": "xml_replace",
                "search_file": "search.xml",
                "replace": "direct replacement"
            }
        ]

        # Should not raise any exceptions
        validate_operations(mixed_operations)

    def test_file_not_found_error_handling(self):
        """Test error handling when referenced files don't exist."""
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_dir = Path(temp_dir)

            replacement = {
                "search_file": "nonexistent.xml",
                "replace_file": "also_nonexistent.xml"
            }

            # Should raise SystemExit due to file not found
            with pytest.raises(SystemExit):
                _process_file_references(replacement, temp_dir)
