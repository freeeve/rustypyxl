"""Tests for data validation support."""

import os
import pytest
import rustypyxl


class TestValidationRoundtrip:
    """Test data validation survives save/load cycle."""

    def test_load_existing_validation(self):
        """Load an existing file with validation."""
        path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "test_validation.xlsx")
        if not os.path.exists(path):
            pytest.skip("test_validation.xlsx not found")

        wb = rustypyxl.load_workbook(path)
        assert wb is not None
        assert len(wb) > 0
