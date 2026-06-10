# Python tests silently skip in CI (fixtures are gitignored root files)

`conftest.py` points `sample_xlsx_path` at the repo-root `test_simple.xlsx`,
and test_hyperlinks/test_comments/test_formulas/test_named_ranges/
test_protection/test_validation/test_roundtrip `pytest.skip` when root
`test_*.xlsx` files are absent — which they always are in CI (`*.xlsx` is
gitignored). test_protection.py and test_validation.py have exactly one test
each, so protection and validation have zero effective CI coverage.

Fix: generate or commit minimal fixtures under `tests/fixtures/`
(`.gitignore` already whitelists `!tests/fixtures/*.xlsx`) and repoint the
tests. Prefer generating with openpyxl in a session-scoped fixture so the
load path is tested against independently-produced files.

Also: register the `slow` pytest mark (test_openpyxl_compat.py:745 warns).
