"""
Unit tests for configuration loading and validation functions.

Tests the dict-based JSON config format, expansion to internal operations,
and margin settings parsing.
"""
import pytest
import json
import tempfile
from pathlib import Path

from src.config import load_operations_from_json, validate_operations, _expand_dict_config, _parse_margins_value


class TestDictConfigExpansion:
    """Test expanding dict config into internal operations array."""

    def test_replace_tuple_format(self):
        """Test replace with [search, replace] tuples."""
        data = {"replace": [["old", "new"], ["foo", "bar"]]}
        ops, settings = _expand_dict_config(data, Path("."))
        assert len(ops) == 2
        assert ops[0] == {"op": "replace", "search": "old", "replace": "new"}
        assert ops[1] == {"op": "replace", "search": "foo", "replace": "bar"}

    def test_replace_with_options(self):
        """Test replace with [search, replace, {options}] tuples."""
        data = {"replace": [["pat.*", "repl", {"regex": True}]]}
        ops, _ = _expand_dict_config(data, Path("."))
        assert ops[0] == {"op": "replace", "search": "pat.*", "replace": "repl", "regex": True}

    def test_replace_dict_format(self):
        """Test replace with dict entries."""
        data = {"replace": [{"search": "old", "replace": "new", "regex": True}]}
        ops, _ = _expand_dict_config(data, Path("."))
        assert ops[0] == {"op": "replace", "search": "old", "replace": "new", "regex": True}

    def test_xml_replace(self):
        """Test xml_replace expansion."""
        data = {"xml_replace": [{"search": "<a>", "replace": "<b>"}]}
        ops, _ = _expand_dict_config(data, Path("."))
        assert ops[0]["op"] == "xml_replace"
        assert ops[0]["search"] == "<a>"

    def test_font_size(self):
        """Test font_size expansion."""
        data = {"font_size": {"from": 8, "to": 10}}
        ops, _ = _expand_dict_config(data, Path("."))
        assert ops[0] == {"op": "font_size", "from": 8, "to": 10}

    def test_clear_properties_true(self):
        """Test clear_properties with true value."""
        data = {"clear_properties": True}
        ops, _ = _expand_dict_config(data, Path("."))
        assert ops[0]["op"] == "clear_properties"
        assert "author" in ops[0]["properties"]
        assert "company" in ops[0]["properties"]

    def test_clear_properties_list(self):
        """Test clear_properties with list of property names."""
        data = {"clear_properties": ["author", "title"]}
        ops, _ = _expand_dict_config(data, Path("."))
        assert ops[0]["properties"] == ["author", "title"]

    def test_clear_properties_string(self):
        """Test clear_properties with single string."""
        data = {"clear_properties": "author"}
        ops, _ = _expand_dict_config(data, Path("."))
        assert ops[0]["properties"] == ["author"]

    def test_set_comments(self):
        """Test set_comments expansion."""
        data = {"set_comments": "{{FILENAME}}"}
        ops, _ = _expand_dict_config(data, Path("."))
        assert ops[0] == {"op": "set_comments", "value": "{{FILENAME}}"}

    def test_table_header_repeat_bool(self):
        """Test table_header_repeat with bool."""
        data = {"table_header_repeat": True}
        ops, _ = _expand_dict_config(data, Path("."))
        assert ops[0] == {"op": "table_header_repeat", "enabled": True}

    def test_table_header_repeat_string(self):
        """Test table_header_repeat with pattern string."""
        data = {"table_header_repeat": "Phase, Time"}
        ops, _ = _expand_dict_config(data, Path("."))
        assert ops[0] == {"op": "table_header_repeat", "pattern": "Phase, Time", "enabled": True}

    def test_cleanup_empty_after_string(self):
        """Test cleanup_empty_after with single string."""
        data = {"cleanup_empty_after": "HEADER"}
        ops, _ = _expand_dict_config(data, Path("."))
        assert ops[0] == {"op": "cleanup_empty_after", "pattern": "HEADER"}

    def test_cleanup_empty_after_list(self):
        """Test cleanup_empty_after with list."""
        data = {"cleanup_empty_after": ["PAT1", "PAT2"]}
        ops, _ = _expand_dict_config(data, Path("."))
        assert len(ops) == 2
        assert ops[0]["pattern"] == "PAT1"
        assert ops[1]["pattern"] == "PAT2"

    def test_replace_image(self):
        """Test replace_image expansion."""
        data = {"replace_image": [{"image_path": "logo.png", "scale": 0.5}]}
        ops, _ = _expand_dict_config(data, Path("."))
        assert ops[0]["op"] == "replace_image"
        assert ops[0]["image_path"] == "logo.png"
        assert ops[0]["scale"] == 0.5

    def test_align_table_cells(self):
        """Test align_table_cells expansion."""
        data = {"align_table_cells": [{"patterns": ["x"], "alignment": "center"}]}
        ops, _ = _expand_dict_config(data, Path("."))
        assert ops[0]["op"] == "align_table_cells"
        assert ops[0]["patterns"] == ["x"]

    def test_replace_table_cell(self):
        """Test replace_table_cell expansion."""
        data = {"replace_table_cell": [{"row": 0, "column": 1, "replace": "text"}]}
        ops, _ = _expand_dict_config(data, Path("."))
        assert ops[0]["op"] == "replace_table_cell"
        assert ops[0]["row"] == 0

    def test_set_table_column_widths(self):
        """Test set_table_column_widths expansion."""
        data = {"set_table_column_widths": [{"column_widths": [1.5, 2.0]}]}
        ops, _ = _expand_dict_config(data, Path("."))
        assert ops[0]["op"] == "set_table_column_widths"
        assert ops[0]["column_widths"] == [1.5, 2.0]

    def test_replace_in_table(self):
        """Test replace_in_table expansion."""
        data = {"replace_in_table": [{"table_heading": "O2", "search": "a", "replace": "b"}]}
        ops, _ = _expand_dict_config(data, Path("."))
        assert ops[0]["op"] == "replace_in_table"
        assert ops[0]["table_heading"] == "O2"

    def test_single_dict_auto_wraps_in_list(self):
        """Test that single dict values for list ops get auto-wrapped."""
        data = {"replace_image": {"image_path": "logo.png"}}
        ops, _ = _expand_dict_config(data, Path("."))
        assert ops[0]["op"] == "replace_image"


class TestDictConfigSettings:
    """Test settings extraction from dict config."""

    def test_margins_string(self):
        """Test margins as comma-separated string."""
        data = {"margins": "1,1,1.5,1.5"}
        _, settings = _expand_dict_config(data, Path("."))
        assert settings['standardize_margins'] is True
        assert settings['margins'] == {'top': 1.0, 'bottom': 1.0, 'left': 1.5, 'right': 1.5}

    def test_margins_dict(self):
        """Test margins as dict."""
        data = {"margins": {"top": 0.5, "bottom": 1.0, "left": 0.75, "right": 1.25}}
        _, settings = _expand_dict_config(data, Path("."))
        assert settings['margins'] == {'top': 0.5, 'bottom': 1.0, 'left': 0.75, 'right': 1.25}

    def test_margins_preset(self):
        """Test margins as preset name."""
        data = {"margins": "letter"}
        _, settings = _expand_dict_config(data, Path("."))
        assert settings['margins'] == {'top': 1.0, 'bottom': 1.0, 'left': 1.0, 'right': 1.0}

    def test_preserve_formatting(self):
        """Test preserve_formatting setting."""
        data = {"preserve_formatting": False}
        _, settings = _expand_dict_config(data, Path("."))
        assert settings['preserve_formatting'] is False

    def test_no_settings_returns_empty(self):
        """Test that no settings keys returns empty dict."""
        data = {"replace": [["a", "b"]]}
        _, settings = _expand_dict_config(data, Path("."))
        assert settings == {}


class TestDictConfigValidation:
    """Test validation of dict config."""

    def test_unknown_key_raises(self):
        """Test that unknown config keys cause an error."""
        data = {"unknown_key": "value"}
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
            json.dump(data, f)
            f.flush()
            with pytest.raises(SystemExit):
                load_operations_from_json(Path(f.name))

    def test_array_format_rejected(self):
        """Test that old array format is rejected."""
        data = [{"op": "replace", "search": "a", "replace": "b"}]
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
            json.dump(data, f)
            f.flush()
            with pytest.raises(SystemExit):
                load_operations_from_json(Path(f.name))

    def test_replace_invalid_entry(self):
        """Test replace with invalid entry type."""
        data = {"replace": ["not_a_list_or_dict"]}
        with pytest.raises(ValueError):
            _expand_dict_config(data, Path("."))

    def test_replace_too_few_elements(self):
        """Test replace with too few elements in tuple."""
        data = {"replace": [["only_one"]]}
        with pytest.raises(ValueError):
            _expand_dict_config(data, Path("."))

    def test_clear_properties_false_raises(self):
        """Test clear_properties with false value."""
        data = {"clear_properties": False}
        with pytest.raises(ValueError):
            _expand_dict_config(data, Path("."))

    def test_preserve_formatting_non_bool_raises(self):
        """Test preserve_formatting with non-bool value."""
        data = {"preserve_formatting": "yes"}
        with pytest.raises(ValueError):
            _expand_dict_config(data, Path("."))


class TestMarginsValueParsing:
    """Test margin value parsing."""

    def test_string_format(self):
        """Test comma-separated string."""
        result = _parse_margins_value("0.5,1.5,0.75,1.25")
        assert result == {'top': 0.5, 'bottom': 1.5, 'left': 0.75, 'right': 1.25}

    def test_dict_format(self):
        """Test dict format with partial overrides."""
        result = _parse_margins_value({"top": 0.5, "left": 2.0})
        assert result['top'] == 0.5
        assert result['left'] == 2.0
        assert result['bottom'] == 1.0  # default
        assert result['right'] == 1.0  # default

    def test_preset_letter(self):
        result = _parse_margins_value("letter")
        assert result == {'top': 1.0, 'bottom': 1.0, 'left': 1.0, 'right': 1.0}

    def test_preset_case_insensitive(self):
        result = _parse_margins_value("LETTER")
        assert result == {'top': 1.0, 'bottom': 1.0, 'left': 1.0, 'right': 1.0}

    def test_wrong_count_raises(self):
        with pytest.raises(ValueError):
            _parse_margins_value("1.0,2.0,3.0")

    def test_non_numeric_raises(self):
        with pytest.raises(ValueError):
            _parse_margins_value("1.0,invalid,3.0,4.0")

    def test_invalid_type_raises(self):
        with pytest.raises(ValueError):
            _parse_margins_value(42)


class TestLoadFromFile:
    """Test loading config from actual files."""

    def test_load_simple_replace(self):
        """Test loading a simple replace config from file."""
        data = {"replace": [["old", "new"]]}
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
            json.dump(data, f)
            f.flush()
            ops, settings = load_operations_from_json(Path(f.name))
            assert len(ops) == 1
            assert ops[0]["op"] == "replace"
            assert ops[0]["search"] == "old"

    def test_load_with_settings(self):
        """Test loading config with settings."""
        data = {
            "replace": [["old", "new"]],
            "margins": "1,1,1.5,1.5",
            "preserve_formatting": False
        }
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
            json.dump(data, f)
            f.flush()
            ops, settings = load_operations_from_json(Path(f.name))
            assert settings['standardize_margins'] is True
            assert settings['preserve_formatting'] is False

    def test_load_complex_config(self):
        """Test loading a complex config with multiple operation types."""
        data = {
            "replace": [["old", "new"]],
            "font_size": {"from": 8, "to": 10},
            "clear_properties": True,
            "set_comments": "{{FILENAME}}",
            "table_header_repeat": True,
            "cleanup_empty_after": "HEADER"
        }
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
            json.dump(data, f)
            f.flush()
            ops, _ = load_operations_from_json(Path(f.name))
            op_types = [op["op"] for op in ops]
            assert "replace" in op_types
            assert "font_size" in op_types
            assert "clear_properties" in op_types
            assert "set_comments" in op_types
            assert "table_header_repeat" in op_types
            assert "cleanup_empty_after" in op_types

    def test_load_file_not_found(self):
        """Test loading from nonexistent file."""
        with pytest.raises(SystemExit):
            load_operations_from_json(Path("nonexistent.json"))

    def test_load_invalid_json(self):
        """Test loading invalid JSON."""
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
            f.write("{invalid json")
            f.flush()
            with pytest.raises(SystemExit):
                load_operations_from_json(Path(f.name))
