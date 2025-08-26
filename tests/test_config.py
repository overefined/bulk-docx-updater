"""
Unit tests for configuration loading and validation functions.

Tests JSON configuration loading, replacement validation,
and margin settings parsing.
"""
import pytest
import json
import sys
from unittest.mock import patch, mock_open, Mock
from pathlib import Path

from config import (
    load_replacements_from_json,
    validate_replacements,
    parse_margin_settings
)


class TestConfigLoading:
    """Test cases for configuration loading functions."""
    
    def test_load_replacements_from_json_list_format(self):
        """Test loading replacements from JSON list format."""
        json_data = [
            {"search": "old1", "replace": "new1"},
            {"search": "old2", "replace": "new2"}
        ]
        json_content = json.dumps(json_data)
        
        with patch("builtins.open", mock_open(read_data=json_content)):
            result = load_replacements_from_json(Path("test.json"))
            
            assert result == json_data
    
    def test_load_replacements_from_json_dict_format(self):
        """Test loading replacements from JSON dict format with 'replacements' key."""
        replacements_data = [
            {"search": "old1", "replace": "new1"},
            {"search": "old2", "replace": "new2"}
        ]
        json_data = {"replacements": replacements_data}
        json_content = json.dumps(json_data)
        
        with patch("builtins.open", mock_open(read_data=json_content)):
            result = load_replacements_from_json(Path("test.json"))
            
            assert result == replacements_data
    
    def test_load_replacements_from_json_invalid_format(self):
        """Test loading replacements from invalid JSON format."""
        json_data = {"invalid": "format"}
        json_content = json.dumps(json_data)
        
        with patch("builtins.open", mock_open(read_data=json_content)):
            with pytest.raises(SystemExit):
                load_replacements_from_json(Path("test.json"))
    
    def test_load_replacements_from_json_file_not_found(self):
        """Test loading replacements when file doesn't exist."""
        with patch("builtins.open", side_effect=FileNotFoundError("File not found")):
            with pytest.raises(SystemExit):
                load_replacements_from_json(Path("nonexistent.json"))
    
    def test_load_replacements_from_json_invalid_json(self):
        """Test loading replacements with malformed JSON."""
        invalid_json = '{"invalid": json syntax'
        
        with patch("builtins.open", mock_open(read_data=invalid_json)):
            with pytest.raises(SystemExit):
                load_replacements_from_json(Path("test.json"))
    
    def test_load_replacements_from_json_utf8_encoding(self):
        """Test that JSON files are loaded with UTF-8 encoding."""
        json_data = [{"search": "café", "replace": "coffee"}]
        json_content = json.dumps(json_data, ensure_ascii=False)
        
        mock_file = mock_open(read_data=json_content)
        with patch("builtins.open", mock_file):
            result = load_replacements_from_json(Path("test.json"))
            
            # Verify UTF-8 encoding was used
            mock_file.assert_called_once_with(Path("test.json"), 'r', encoding='utf-8')
            assert result == json_data


class TestReplacementValidation:
    """Test cases for replacement validation."""
    
    def test_validate_replacements_valid_search_replace(self):
        """Test validation of valid search/replace pairs."""
        replacements = [
            {"search": "old1", "replace": "new1"},
            {"search": "old2", "replace": "new2"}
        ]
        
        # Should not raise any exception
        validate_replacements(replacements)
    
    def test_validate_replacements_valid_cleanup_action(self):
        """Test validation of valid cleanup actions."""
        replacements = [
            {"remove_empty_paragraphs_after": "PATTERN1"},
            {"remove_empty_paragraphs_after": "PATTERN2"}
        ]
        
        # Should not raise any exception
        validate_replacements(replacements)
    
    def test_validate_replacements_valid_insert_after(self):
        """Test validation of valid search/insert_after pairs."""
        replacements = [
            {"search": "SITE PHOTOS", "insert_after": "Photo1paragraphbreakPhoto2"},
            {"search": "APPENDIX", "insert_after": "pagebreak{format:center}New Content{/format}"}
        ]
        
        # Should not raise any exception
        validate_replacements(replacements)
    
    def test_validate_replacements_mixed_valid_types(self):
        """Test validation of mixed valid replacement types."""
        replacements = [
            {"search": "old", "replace": "new"},
            {"search": "PHOTOS", "insert_after": "Photo content"},
            {"remove_empty_paragraphs_after": "PATTERN"}
        ]
        
        # Should not raise any exception
        validate_replacements(replacements)
    
    def test_validate_replacements_with_cleanup_flag_replace(self):
        """Test validation of replace operations with cleanup flag."""
        replacements = [
            {"search": "old", "replace": "new", "remove_empty_paragraphs_after": True},
            {"search": "test", "replace": "result"}
        ]
        
        # Should not raise any exception
        validate_replacements(replacements)
    
    def test_validate_replacements_with_cleanup_flag_insert_after(self):
        """Test validation of insert_after operations with cleanup flag."""
        replacements = [
            {"search": "PHOTOS", "insert_after": "Photo content", "remove_empty_paragraphs_after": True},
            {"search": "HEADER", "insert_after": "New content"}
        ]
        
        # Should not raise any exception
        validate_replacements(replacements)
    
    def test_validate_replacements_invalid_dict_type(self):
        """Test validation fails with non-dict replacement."""
        replacements = ["invalid", {"search": "old", "replace": "new"}]
        
        with pytest.raises(SystemExit):
            validate_replacements(replacements)
    
    def test_validate_replacements_missing_search(self):
        """Test validation fails when search key is missing."""
        replacements = [{"replace": "new"}]
        
        with pytest.raises(SystemExit):
            validate_replacements(replacements)
    
    def test_validate_replacements_missing_replace(self):
        """Test validation fails when replace key is missing."""
        replacements = [{"search": "old"}]
        
        with pytest.raises(SystemExit):
            validate_replacements(replacements)
    
    def test_validate_replacements_missing_insert_after(self):
        """Test validation fails when insert_after key is missing from search/insert_after pair."""
        replacements = [{"search": "PHOTOS"}]
        
        with pytest.raises(SystemExit):
            validate_replacements(replacements)
    
    def test_validate_replacements_conflicting_replace_and_insert_after(self):
        """Test validation fails when both replace and insert_after are specified."""
        replacements = [{"search": "PHOTOS", "replace": "replacement", "insert_after": "insertion"}]
        
        with pytest.raises(SystemExit):
            validate_replacements(replacements)
    
    def test_validate_replacements_no_valid_keys(self):
        """Test validation fails when replacement has no valid action keys."""
        replacements = [{"invalid": "keys"}]
        
        with pytest.raises(SystemExit):
            validate_replacements(replacements)
    
    def test_validate_replacements_empty_list(self):
        """Test validation of empty replacement list."""
        replacements = []
        
        # Should not raise any exception
        validate_replacements(replacements)
    
    def test_validate_replacements_invalid_cleanup_flag_type(self):
        """Test validation fails when cleanup flag is not boolean true for search/replace operations."""
        replacements = [{"search": "old", "replace": "new", "remove_empty_paragraphs_after": False}]
        
        with pytest.raises(SystemExit):
            validate_replacements(replacements)
    
    def test_validate_replacements_invalid_cleanup_flag_string(self):
        """Test validation fails when cleanup flag is string for search/replace operations."""
        replacements = [{"search": "old", "replace": "new", "remove_empty_paragraphs_after": "pattern"}]
        
        with pytest.raises(SystemExit):
            validate_replacements(replacements)
    
    def test_validate_replacements_invalid_standalone_cleanup_type(self):
        """Test validation fails when standalone cleanup is not string."""
        replacements = [{"remove_empty_paragraphs_after": True}]
        
        with pytest.raises(SystemExit):
            validate_replacements(replacements)
    
    def test_validate_replacements_complex_valid_structure(self):
        """Test validation of complex but valid replacement structure."""
        replacements = [
            {"search": "pattern1", "replace": "replacement1"},
            {"search": "pattern2", "replace": "{format:bold}replacement2{/format}"},
            {"remove_empty_paragraphs_after": "CLEANUP_PATTERN"},
            {"search": "TESTER QUALIFICATIONS", "replace": "INSPECTOR QUALIFICATIONS"}
        ]
        
        # Should not raise any exception
        validate_replacements(replacements)


class TestMarginSettingsParsing:
    """Test cases for margin settings parsing."""
    
    def create_mock_args(self, **kwargs):
        """Create mock arguments object with specified attributes."""
        args = Mock()
        
        # Set default values
        args.margins = None
        args.margin_top = None
        args.margin_bottom = None
        args.margin_left = None
        args.margin_right = None
        
        # Override with provided kwargs
        for key, value in kwargs.items():
            setattr(args, key, value)
            
        return args
    
    def test_parse_margin_settings_defaults(self):
        """Test parsing margin settings with default values."""
        args = self.create_mock_args()
        
        result = parse_margin_settings(args)
        
        expected = {'top': 1.0, 'bottom': 1.0, 'left': 1.0, 'right': 1.0}
        assert result == expected
    
    def test_parse_margin_settings_preset_letter(self):
        """Test parsing margin settings with letter preset."""
        args = self.create_mock_args(margins="letter")
        
        result = parse_margin_settings(args)
        
        expected = {'top': 1.0, 'bottom': 1.0, 'left': 1.0, 'right': 1.0}
        assert result == expected
    
    def test_parse_margin_settings_preset_legal(self):
        """Test parsing margin settings with legal preset."""
        args = self.create_mock_args(margins="legal")
        
        result = parse_margin_settings(args)
        
        expected = {'top': 1.0, 'bottom': 1.0, 'left': 1.0, 'right': 1.0}
        assert result == expected
    
    def test_parse_margin_settings_preset_a4(self):
        """Test parsing margin settings with A4 preset."""
        args = self.create_mock_args(margins="A4")
        
        result = parse_margin_settings(args)
        
        expected = {'top': 1.0, 'bottom': 1.0, 'left': 1.0, 'right': 1.0}
        assert result == expected
    
    def test_parse_margin_settings_custom_comma_separated(self):
        """Test parsing margin settings with custom comma-separated values."""
        args = self.create_mock_args(margins="0.5,1.5,0.75,1.25")
        
        result = parse_margin_settings(args)
        
        expected = {'top': 0.5, 'bottom': 1.5, 'left': 0.75, 'right': 1.25}
        assert result == expected
    
    def test_parse_margin_settings_individual_overrides(self):
        """Test parsing margin settings with individual margin overrides."""
        args = self.create_mock_args(
            margin_top=0.5,
            margin_bottom=2.0,
            margin_left=0.25,
            margin_right=1.75
        )
        
        result = parse_margin_settings(args)
        
        expected = {'top': 0.5, 'bottom': 2.0, 'left': 0.25, 'right': 1.75}
        assert result == expected
    
    def test_parse_margin_settings_preset_with_individual_override(self):
        """Test that individual margin settings override preset values."""
        args = self.create_mock_args(
            margins="letter",
            margin_top=0.5,
            margin_right=2.0
        )
        
        result = parse_margin_settings(args)
        
        expected = {'top': 0.5, 'bottom': 1.0, 'left': 1.0, 'right': 2.0}
        assert result == expected
    
    def test_parse_margin_settings_custom_with_individual_override(self):
        """Test that individual margins override custom comma-separated values."""
        args = self.create_mock_args(
            margins="0.5,0.5,0.5,0.5",
            margin_bottom=2.0,
            margin_left=1.5
        )
        
        result = parse_margin_settings(args)
        
        expected = {'top': 0.5, 'bottom': 2.0, 'left': 1.5, 'right': 0.5}
        assert result == expected
    
    def test_parse_margin_settings_invalid_comma_count(self):
        """Test parsing margin settings with wrong number of comma-separated values."""
        args = self.create_mock_args(margins="1.0,2.0,3.0")  # Only 3 values
        
        with pytest.raises(SystemExit):
            parse_margin_settings(args)
    
    def test_parse_margin_settings_invalid_number_format(self):
        """Test parsing margin settings with non-numeric values."""
        args = self.create_mock_args(margins="1.0,invalid,3.0,4.0")
        
        with pytest.raises(SystemExit):
            parse_margin_settings(args)
    
    def test_parse_margin_settings_empty_values(self):
        """Test parsing margin settings with empty values in comma-separated string."""
        args = self.create_mock_args(margins="1.0,,3.0,4.0")
        
        with pytest.raises((SystemExit, ValueError)):
            parse_margin_settings(args)
    
    def test_parse_margin_settings_case_insensitive_presets(self):
        """Test that preset names are case insensitive."""
        test_cases = ["LETTER", "Legal", "a4", "LeTtEr"]
        
        for preset in test_cases:
            args = self.create_mock_args(margins=preset)
            result = parse_margin_settings(args)
            
            expected = {'top': 1.0, 'bottom': 1.0, 'left': 1.0, 'right': 1.0}
            assert result == expected
    
    def test_parse_margin_settings_float_precision(self):
        """Test parsing margin settings with high precision floats."""
        args = self.create_mock_args(margins="0.125,0.375,0.625,0.875")
        
        result = parse_margin_settings(args)
        
        expected = {'top': 0.125, 'bottom': 0.375, 'left': 0.625, 'right': 0.875}
        assert result == expected
    
    def test_parse_margin_settings_negative_values(self):
        """Test parsing margin settings with negative values (should work)."""
        args = self.create_mock_args(margins="-0.5,1.0,0.0,2.5")
        
        result = parse_margin_settings(args)
        
        expected = {'top': -0.5, 'bottom': 1.0, 'left': 0.0, 'right': 2.5}
        assert result == expected
    
    def test_parse_margin_settings_zero_values(self):
        """Test parsing margin settings with zero values."""
        args = self.create_mock_args(margins="0,0,0,0")
        
        result = parse_margin_settings(args)
        
        expected = {'top': 0.0, 'bottom': 0.0, 'left': 0.0, 'right': 0.0}
        assert result == expected