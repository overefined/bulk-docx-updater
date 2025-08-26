"""
Test package for the DOCX bulk updater.

This package contains comprehensive unit tests, integration tests,
and test fixtures for all components of the DOCX bulk updater tool.

Test Modules:
- test_formatting.py: Tests for FormattingProcessor class
- test_text_replacement.py: Tests for TextReplacer class  
- test_document_processor.py: Tests for DocxBulkUpdater class
- test_config.py: Tests for configuration loading and validation
- test_cli.py: Tests for CLI interface
- test_integration.py: End-to-end integration tests
- conftest.py: Shared fixtures and configuration

Usage:
    Run all tests: pytest
    Run with coverage: pytest --cov
    Run specific test: pytest tests/test_formatting.py
    Run integration tests: pytest -m integration
    Run unit tests only: pytest -m unit
"""