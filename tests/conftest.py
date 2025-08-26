"""
Shared test configuration for pytest.

Provides pytest configuration and markers for the test suite.
Individual test files manage their own fixtures and setup.
"""
import pytest


# Pytest configuration
def pytest_configure(config):
    """Configure pytest with custom markers."""
    config.addinivalue_line(
        "markers", "integration: mark test as an integration test"
    )
    config.addinivalue_line(
        "markers", "slow: mark test as slow running"
    )
    config.addinivalue_line(
        "markers", "unit: mark test as a unit test"
    )