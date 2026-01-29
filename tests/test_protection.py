"""Tests for sheet protection support."""

import os
import pytest
import rustypyxl


class TestProtectionRoundtrip:
    """Test protection survives save/load cycle."""

    def test_load_existing_protection(self):
        """Load an existing file with protection."""
        path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "test_protection.xlsx")
        if not os.path.exists(path):
            pytest.skip("test_protection.xlsx not found")

        wb = rustypyxl.load_workbook(path)
        assert wb is not None
        assert len(wb) > 0
