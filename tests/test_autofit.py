"""auto_fit_column / auto_fit_all size columns to their content. The width is a
character-count estimate of the displayed string, so tests assert a sensible
band and relative ordering, not exact pixels.
"""

import rustypyxl


def test_auto_fit_column_returns_width_and_persists(tmp_path):
    wb = rustypyxl.Workbook()
    ws = wb.create_sheet("S")
    ws["A1"] = "hi"
    ws["A2"] = "a much longer piece of text"
    width = ws.auto_fit_column(1)
    assert width is not None
    assert 28 < width < 30  # 27 chars + padding

    # The width survives a save/load round-trip.
    out = str(tmp_path / "w.xlsx")
    wb.save(out)
    reloaded = rustypyxl.load_workbook(out)["S"]
    # Re-fitting the reloaded sheet yields the same estimate.
    assert abs(reloaded.auto_fit_column(1) - width) < 0.001


def test_auto_fit_column_empty_returns_none():
    wb = rustypyxl.Workbook()
    ws = wb.create_sheet("S")
    ws["A1"] = "x"
    assert ws.auto_fit_column(5) is None


def test_auto_fit_measures_formatted_value():
    wb = rustypyxl.Workbook()
    ws = wb.create_sheet("S")
    ws["A1"] = 0.1
    ws["A1"].number_format = "0.00%"  # displays "10.00%"
    width = ws.auto_fit_column(1)
    assert 7 < width < 9  # 6 chars + padding, not len("0.1")


def test_auto_fit_all_sizes_every_populated_column():
    wb = rustypyxl.Workbook()
    ws = wb.create_sheet("S")
    ws["A1"] = "short"
    ws["C1"] = "a longer heading here"
    ws.auto_fit_all()
    a = ws.auto_fit_column(1)
    c = ws.auto_fit_column(3)
    assert c > a
