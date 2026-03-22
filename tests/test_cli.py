"""
Unit tests for the CLI module.

Tests command-line argument parsing, file discovery,
and CLI workflow orchestration.
"""
import pytest
from unittest.mock import patch, Mock, MagicMock
from pathlib import Path
import sys
from argparse import Namespace

from src.cli import main


class TestCLIArgumentParsing:
    """Test cases for CLI argument parsing."""

    @patch('sys.argv', ['main.py', '/test/path', '--config', 'config.json'])
    def test_basic_config_arguments(self):
        """Test parsing basic arguments with config file."""
        test_ops = [{"op": "replace", "search": "old", "replace": "new"}]
        with patch('src.cli.load_operations_from_json') as mock_load:
            with patch('src.cli.validate_operations'):
                with patch('pathlib.Path.is_file', return_value=True):
                    with patch('src.cli.DocxBulkUpdater') as mock_updater:
                        mock_updater.return_value.modify_docx.return_value = True
                        mock_load.return_value = (test_ops, {})

                        try:
                            main()
                        except SystemExit:
                            pass

                        mock_load.assert_called_once()

    @patch('sys.argv', ['main.py', '/test/path', '--search', 'old', '--replace', 'new'])
    def test_search_replace_arguments(self):
        """Test parsing search/replace command line arguments."""
        with patch('src.cli.validate_operations'):
            with patch('pathlib.Path.is_file', return_value=True):
                with patch('src.cli.DocxBulkUpdater') as mock_updater:
                    mock_updater.return_value.modify_docx.return_value = True

                    try:
                        main()
                    except SystemExit:
                        pass

                    # Verify DocxBulkUpdater was called with correct operations
                    args, kwargs = mock_updater.call_args
                    operations = args[0]
                    assert operations == [{"op": "replace", "search": "old", "replace": "new"}]

    @patch('sys.argv', ['main.py', '/test/path'])
    def test_missing_required_arguments(self):
        """Test error handling when required arguments are missing."""
        with pytest.raises(SystemExit):
            main()

    @patch('sys.argv', ['main.py', '/test/path', '--recursive', '--config', 'config.json'])
    def test_recursive_flag(self):
        """Test recursive directory processing flag."""
        with patch('src.cli.load_operations_from_json', return_value=([], {})):
            with patch('src.cli.validate_operations'):
                with patch('pathlib.Path.is_dir', return_value=True):
                    with patch('pathlib.Path.rglob') as mock_rglob:
                        with patch('src.cli.DocxBulkUpdater') as mock_updater:
                            mock_rglob.return_value = [Path("test.docx")]
                            mock_updater.return_value.modify_docx.return_value = False

                            try:
                                main()
                            except SystemExit:
                                pass

                            mock_rglob.assert_called_once_with("*.docx")

    @patch('sys.argv', ['main.py', '/test/path', '--config', 'config.json', '--pattern', '*.doc'])
    def test_custom_file_pattern(self):
        """Test custom file pattern argument."""
        with patch('src.cli.load_operations_from_json', return_value=([], {})):
            with patch('src.cli.validate_operations'):
                with patch('pathlib.Path.is_dir', return_value=True):
                    with patch('pathlib.Path.glob') as mock_glob:
                        with patch('src.cli.DocxBulkUpdater') as mock_updater:
                            mock_glob.return_value = [Path("test.doc")]
                            mock_updater.return_value.modify_docx.return_value = False

                            try:
                                main()
                            except SystemExit:
                                pass

                            mock_glob.assert_called_once_with("*.doc")


class TestCLIFileDiscovery:
    """Test cases for file discovery logic."""

    @patch('sys.argv', ['main.py', 'single_file.docx', '--search', 'old', '--replace', 'new'])
    def test_single_file_processing(self):
        """Test processing a single file."""
        mock_file = Mock()
        mock_file.is_file.return_value = True
        mock_file.is_dir.return_value = False

        with patch('src.cli.Path', return_value=mock_file):
            with patch('src.cli.DocxBulkUpdater') as mock_updater:
                mock_updater.return_value.modify_docx.return_value = True

                try:
                    main()
                except SystemExit:
                    pass

                mock_updater.return_value.modify_docx.assert_called_once_with(mock_file)

    @patch('sys.argv', ['main.py', '/test/directory', '--config', 'config.json'])
    def test_directory_processing(self):
        """Test processing files in a directory."""
        mock_path = Mock()
        mock_path.is_file.return_value = False
        mock_path.is_dir.return_value = True
        mock_path.glob.return_value = [Path("file1.docx"), Path("file2.docx")]

        with patch('src.cli.Path', return_value=mock_path):
            with patch('src.cli.load_operations_from_json', return_value=([], {})):
                with patch('src.cli.validate_operations'):
                    with patch('src.cli.DocxBulkUpdater') as mock_updater:
                        mock_updater.return_value.modify_docx.return_value = False

                        try:
                            main()
                        except SystemExit:
                            pass

                        mock_path.glob.assert_called_once_with("*.docx")

    @patch('sys.argv', ['main.py', '/nonexistent/path', '--search', 'old', '--replace', 'new'])
    def test_nonexistent_path_error(self):
        """Test error handling for nonexistent paths."""
        mock_path = Mock()
        mock_path.is_file.return_value = False
        mock_path.is_dir.return_value = False

        with patch('src.cli.Path', return_value=mock_path):
            with pytest.raises(SystemExit):
                main()

    @patch('sys.argv', ['main.py', '/empty/directory', '--config', 'config.json'])
    def test_empty_directory_handling(self):
        """Test handling of directory with no matching files."""
        mock_path = Mock()
        mock_path.is_file.return_value = False
        mock_path.is_dir.return_value = True
        mock_path.glob.return_value = []

        with patch('src.cli.Path', return_value=mock_path):
            with patch('src.cli.load_operations_from_json', return_value=([], {})):
                with patch('src.cli.validate_operations'):
                    with patch('builtins.print') as mock_print:
                        try:
                            main()
                        except SystemExit:
                            pass

                        assert any("No files" in str(call) for call in mock_print.call_args_list)


class TestCLIWorkflow:
    """Test cases for CLI workflow orchestration."""

    @patch('sys.argv', ['main.py', 'test.docx', '--config', 'config.json'])
    def test_successful_processing_workflow(self):
        """Test complete successful processing workflow."""
        test_operations = [{"op": "replace", "search": "old", "replace": "new"}]

        with patch('src.cli.load_operations_from_json', return_value=(test_operations, {})):
            with patch('src.cli.validate_operations') as mock_validate:
                with patch('pathlib.Path.is_file', return_value=True):
                    with patch('src.cli.DocxBulkUpdater') as mock_updater_class:
                        mock_updater = Mock()
                        mock_updater.modify_docx.return_value = True
                        mock_updater_class.return_value = mock_updater

                        with patch('builtins.print') as mock_print:
                            try:
                                main()
                            except SystemExit:
                                pass

                        mock_validate.assert_called_once_with(test_operations)
                        mock_updater_class.assert_called_once()
                        mock_updater.modify_docx.assert_called_once()

                        success_printed = any("[OK]" in str(call) for call in mock_print.call_args_list)
                        assert success_printed

    @patch('sys.argv', ['main.py', 'test.docx', '--config', 'config.json', '--dry-run'])
    def test_dry_run_workflow(self):
        """Test dry run workflow."""
        test_operations = [{"op": "replace", "search": "old", "replace": "new"}]
        test_changes = {"Body": (["original line"], ["modified line"])}

        with patch('src.cli.load_operations_from_json', return_value=(test_operations, {})):
            with patch('src.cli.validate_operations'):
                with patch('pathlib.Path.is_file', return_value=True):
                    with patch('src.cli.DocxBulkUpdater') as mock_updater_class:
                        mock_updater = Mock()
                        mock_updater.get_document_changes_preview.return_value = test_changes
                        mock_updater.format_diff.return_value = "--- diff output ---"
                        mock_updater_class.return_value = mock_updater

                        with patch('builtins.print') as mock_print:
                            try:
                                main()
                            except SystemExit:
                                pass

                        mock_updater.get_document_changes_preview.assert_called_once()
                        mock_updater.modify_docx.assert_not_called()

                        dry_run_printed = any("DRY RUN" in str(call) for call in mock_print.call_args_list)
                        assert dry_run_printed

    @patch('sys.argv', ['main.py', 'test.docx', '--config', 'config.json', '--dry-run', '--xml-diff'])
    def test_dry_run_with_xml_diff(self):
        """Test dry run workflow with XML diff enabled."""
        test_operations = [{"op": "replace", "search": "old", "replace": "new"}]
        text_changes = {"Body": (["a"], ["b"])}
        xml_changes = {"Body(XML)": (["<p>a</p>"], ["<p>b</p>"])}

        with patch('src.cli.load_operations_from_json', return_value=(test_operations, {})):
            with patch('src.cli.validate_operations'):
                with patch('pathlib.Path.is_file', return_value=True):
                    with patch('src.cli.DocxBulkUpdater') as mock_updater_class:
                        mock_updater = Mock()
                        mock_updater.get_document_changes_preview.return_value = text_changes
                        mock_updater.get_document_xml_changes_preview.return_value = xml_changes
                        mock_updater.format_diff.side_effect = ["--- text diff ---", "--- xml diff ---"]
                        mock_updater_class.return_value = mock_updater

                        with patch('builtins.print') as mock_print:
                            try:
                                main()
                            except SystemExit:
                                pass

                        mock_updater.get_document_changes_preview.assert_called_once()
                        mock_updater.get_document_xml_changes_preview.assert_called_once()
                        printed_xml = any("Body(XML)" in str(call) for call in mock_print.call_args_list)
                        assert printed_xml

    @patch('sys.argv', ['main.py', 'test.docx', '--search', 'old', '--replace', 'new'])
    def test_processing_with_no_changes(self):
        """Test processing workflow when no changes are made."""
        with patch('pathlib.Path.is_file', return_value=True):
            with patch('src.cli.DocxBulkUpdater') as mock_updater_class:
                mock_updater = Mock()
                mock_updater.modify_docx.return_value = False
                mock_updater_class.return_value = mock_updater

                with patch('builtins.print') as mock_print:
                    try:
                        main()
                    except SystemExit:
                        pass

                no_changes_printed = any("no changes" in str(call) for call in mock_print.call_args_list)
                assert no_changes_printed

    @patch('sys.argv', ['main.py', 'test.docx', '--search', 'old', '--replace', 'new'])
    def test_processing_with_exception(self):
        """Test processing workflow when exception occurs."""
        with patch('pathlib.Path.is_file', return_value=True):
            with patch('src.cli.DocxBulkUpdater') as mock_updater_class:
                mock_updater = Mock()
                mock_updater.modify_docx.side_effect = Exception("Test error")
                mock_updater_class.return_value = mock_updater

                with patch('builtins.print') as mock_print:
                    try:
                        main()
                    except SystemExit:
                        pass

                error_printed = any("[ERROR]" in str(call) for call in mock_print.call_args_list)
                assert error_printed

    @patch('sys.argv', ['main.py', '/test/dir', '--config', 'config.json'])
    def test_multiple_file_processing(self):
        """Test processing workflow with multiple files."""
        test_files = [Path("file1.docx"), Path("file2.docx"), Path("file3.docx")]
        test_operations = [{"op": "replace", "search": "old", "replace": "new"}]

        with patch('src.cli.load_operations_from_json', return_value=(test_operations, {})):
            with patch('src.cli.validate_operations'):
                with patch('pathlib.Path.is_dir', return_value=True):
                    with patch('pathlib.Path.glob', return_value=test_files):
                        with patch('src.cli.DocxBulkUpdater') as mock_updater_class:
                            mock_updater = Mock()
                            mock_updater.modify_docx.side_effect = [True, False, True]
                            mock_updater_class.return_value = mock_updater

                            with patch('builtins.print') as mock_print:
                                try:
                                    main()
                                except SystemExit:
                                    pass

                            assert mock_updater.modify_docx.call_count == 3

                            count_printed = any("3 file(s)" in str(call) for call in mock_print.call_args_list)
                            assert count_printed


class TestCLISettingsFromConfig:
    """Test that config settings are passed through to the updater."""

    @patch('sys.argv', ['main.py', 'test.docx', '--config', 'config.json'])
    def test_margins_from_config(self):
        """Test margins from config are passed to updater."""
        margins = {'top': 0.5, 'bottom': 1.0, 'left': 0.75, 'right': 1.25}
        settings = {'standardize_margins': True, 'margins': margins}
        with patch('src.cli.load_operations_from_json', return_value=([], settings)):
            with patch('src.cli.validate_operations'):
                with patch('pathlib.Path.is_file', return_value=True):
                    with patch('src.cli.DocxBulkUpdater') as mock_updater:
                        mock_updater.return_value.modify_docx.return_value = False
                        try:
                            main()
                        except SystemExit:
                            pass
                        _, kwargs = mock_updater.call_args
                        assert kwargs['standardize_margins'] is True
                        assert kwargs['margins'] == margins

    @patch('sys.argv', ['main.py', 'test.docx', '--config', 'config.json'])
    def test_preserve_formatting_from_config(self):
        """Test preserve_formatting from config is passed to updater."""
        settings = {'preserve_formatting': False}
        with patch('src.cli.load_operations_from_json', return_value=([], settings)):
            with patch('src.cli.validate_operations'):
                with patch('pathlib.Path.is_file', return_value=True):
                    with patch('src.cli.DocxBulkUpdater') as mock_updater:
                        mock_updater.return_value.modify_docx.return_value = False
                        try:
                            main()
                        except SystemExit:
                            pass
                        _, kwargs = mock_updater.call_args
                        assert kwargs['preserve_formatting'] is False
