"""The formula engine computes =SUM(A1:A10) and friends -- a capability openpyxl
lacks (it stores formulas and reads Excel's cached result, but does not evaluate).

Scope is a common subset; unsupported syntax returns an Excel-style error string.
"""

import rustypyxl


def _wb():
    wb = rustypyxl.Workbook()
    ws = wb.create_sheet("S")
    wb.write_rows("S", [[10], [20], [30]])  # A1..A3
    return wb


def test_evaluate_formula_arithmetic_and_aggregates():
    wb = _wb()
    assert wb.evaluate_formula("S", "=1+2*3") == 7
    assert wb.evaluate_formula("S", "=SUM(A1:A3)") == 60
    assert wb.evaluate_formula("S", "=AVERAGE(A1:A3)") == 20
    assert wb.evaluate_formula("S", "=A1*A2") == 200


def test_evaluate_formula_logic_and_text():
    wb = _wb()
    assert wb.evaluate_formula("S", '=IF(A1>5,"big","small")') == "big"
    assert wb.evaluate_formula("S", '=UPPER("abc")') == "ABC"
    assert wb.evaluate_formula("S", '="a"&"b"&"c"') == "abc"
    assert wb.evaluate_formula("S", "=COUNTIF(A1:A3,\">15\")") == 2


def test_evaluate_cell_computes_formula_cells():
    wb = _wb()
    ws = wb["S"]
    ws["B1"] = "=SUM(A1:A3)"
    ws["B2"] = "=B1*2"
    assert wb.evaluate_cell("S", 1, 2) == 60
    assert wb.evaluate_cell("S", 2, 2) == 120
    # a plain cell returns its value
    assert wb.evaluate_cell("S", 1, 1) == 10


def test_errors_are_returned_as_strings():
    wb = _wb()
    assert wb.evaluate_formula("S", "=1/0") == "#DIV/0!"
    assert wb.evaluate_formula("S", "=NOSUCHFN(1)") == "#NAME?"
