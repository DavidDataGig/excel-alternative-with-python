"""
Integration test: simulate the exact workflow a user would do in the notebook.
Load the sample data, perform edits, save, reload, verify.
"""
import sys, os, tempfile
sys.path.insert(0, "/home/claude")
from spreadsheet_engine import Spreadsheet, rc_to_a1


def banner(msg):
    print(f"\n{'=' * 60}\n{msg}\n{'=' * 60}")


def show_grid(s, max_rows=None, max_cols=None):
    rows = max_rows or s.rows
    cols = max_cols or s.cols
    # header
    print("    " + "".join(f"{c:>10}" for c in [chr(ord('A') + i) for i in range(cols)]))
    for r in range(rows):
        line = f"{r+1:>3} "
        for c in range(cols):
            v = s.get_display(r, c)
            line += f"{v:>10}"
        print(line)


# The same sample as in the notebook
SAMPLE = [
    ['Product',  'Q1',   'Q2',   'Q3',   'Q4',   'Total',        'Avg'],
    ['Widgets',  '1200', '1450', '1600', '1800', '=SUM(B2:E2)',  '=AVERAGE(B2:E2)'],
    ['Gadgets',  '800',  '950',  '1100', '1300', '=SUM(B3:E3)',  '=AVERAGE(B3:E3)'],
    ['Gizmos',   '500',  '600',  '750',  '900',  '=SUM(B4:E4)',  '=AVERAGE(B4:E4)'],
    ['Doodads',  '300',  '350',  '400',  '450',  '=SUM(B5:E5)',  '=AVERAGE(B5:E5)'],
    ['Total',    '=SUM(B2:B5)', '=SUM(C2:C5)', '=SUM(D2:D5)', '=SUM(E2:E5)',
                 '=SUM(F2:F5)', '=AVERAGE(G2:G5)'],
    [],
    ['Stats', 'Value'],
    ['Best quarter Q4 total', '=E6'],
    ['Growth Q1->Q4 (Widgets)', '=(E2-B2)/B2'],
    ['High performer?', '=IF(F2>5000, "Yes", "No")'],
]


banner("Step 1: Load sample data")
s = Spreadsheet(14, 8)
s.load_from_2d(SAMPLE)
show_grid(s, max_rows=12, max_cols=7)

banner("Step 2: Verify all calculated values")
checks = [
    ("Widgets Total (F2)",       (1, 5), 6050),
    ("Gadgets Total (F3)",       (2, 5), 4150),
    ("Gizmos Total (F4)",        (3, 5), 2750),
    ("Doodads Total (F5)",       (4, 5), 1500),
    ("Q1 total (B6)",            (5, 1), 2800),
    ("Q2 total (C6)",            (5, 2), 3350),
    ("Q3 total (D6)",            (5, 3), 3850),
    ("Q4 total (E6)",            (5, 4), 4450),
    ("Grand total (F6)",         (5, 5), 14450),
    ("Widgets average (G2)",     (1, 6), 1512.5),
    ("Best quarter (B9)",        (8, 1), 4450),
    ("Widget growth Q1->Q4 (B10)", (9, 1), 0.5),
    ("High performer (B11)",     (10, 1), "Yes"),
]
all_ok = True
for name, (r, c), expected in checks:
    got = s.get_value(r, c)
    if isinstance(expected, float):
        ok = abs(got - expected) < 1e-6
    else:
        ok = got == expected
    mark = "OK  " if ok else "FAIL"
    print(f"  [{mark}] {name}: expected {expected!r}, got {got!r}")
    if not ok:
        all_ok = False

banner("Step 3: Edit Widgets Q1 from 1200 to 3000 and check cascading recalc")
s.set_cell(1, 1, "3000")
print("After edit:")
show_grid(s, max_rows=12, max_cols=7)
cascade_checks = [
    ("Widgets total (F2) should be 6050+1800=7850", (1, 5), 7850),
    ("Q1 total (B6) should be 2800+1800=4600",      (5, 1), 4600),
    ("Grand total (F6) should be 14450+1800=16250", (5, 5), 16250),
    ("Best quarter (B9) unchanged = 4450",          (8, 1), 4450),
    ("Growth Q1->Q4 widgets = (1800-3000)/3000 = -0.4", (9, 1), -0.4),
    ("High performer now Yes (F2=7850 > 5000)",     (10, 1), "Yes"),
]
for name, (r, c), expected in cascade_checks:
    got = s.get_value(r, c)
    if isinstance(expected, float):
        ok = abs(got - expected) < 1e-6
    else:
        ok = got == expected
    mark = "OK  " if ok else "FAIL"
    print(f"  [{mark}] {name}: got {got!r}")
    if not ok:
        all_ok = False


banner("Step 4: Undo and verify")
s.undo()
print(f"After undo, Widgets Q1 = {s.get_value(1, 1)} (should be 1200)")
print(f"After undo, Widgets Total = {s.get_value(1, 5)} (should be 6050)")
if s.get_value(1, 1) != 1200 or s.get_value(1, 5) != 6050:
    all_ok = False


banner("Step 5: Redo and verify")
s.redo()
print(f"After redo, Widgets Q1 = {s.get_value(1, 1)} (should be 3000)")
print(f"After redo, Widgets Total = {s.get_value(1, 5)} (should be 7850)")
if s.get_value(1, 1) != 3000 or s.get_value(1, 5) != 7850:
    all_ok = False


banner("Step 6: Save to XLSX and reload")
with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
    xlsx_path = f.name
s.save_xlsx(xlsx_path)
print(f"Saved to {xlsx_path} ({os.path.getsize(xlsx_path)} bytes)")

s2 = Spreadsheet()
s2.load_xlsx(xlsx_path)
print(f"Reloaded. Widgets Total: {s2.get_value(1, 5)} (should be 7850)")
print(f"Reloaded. Grand total: {s2.get_value(5, 5)} (should be 16250)")
if s2.get_value(1, 5) != 7850 or s2.get_value(5, 5) != 16250:
    all_ok = False
os.unlink(xlsx_path)


banner("Step 7: Insert a new row at row 2, verify data shifts")
s.insert_row(1)
# Old row 1 (Widgets) should now be at row 2
w_row = s.get_value(2, 0)
print(f"After insert_row(1), row 3 col A = {w_row!r} (was 'Widgets')")
# NOTE: formulas don't auto-adjust, so B6 (Q1 total) still references B2:B5
# which now points to different cells. This is a documented limitation.


banner("Step 8: Stress test — 100 dependent cells")
ss = Spreadsheet(120, 3)
ss.set_cell(0, 0, "1")
for i in range(1, 100):
    ss.set_cell(i, 0, f"=A{i}+1")
val = ss.get_value(99, 0)
print(f"A100 after chain of 100 = {val} (should be 100)")
if val != 100:
    all_ok = False
# Change A1 and verify propagation
ss.set_cell(0, 0, "500")
val = ss.get_value(99, 0)
print(f"After A1=500, A100 = {val} (should be 599)")
if val != 599:
    all_ok = False


banner("Step 9: Error handling")
se = Spreadsheet(5, 5)
se.set_cell(0, 0, "=1/0")
se.set_cell(0, 1, "=UNKNOWN(1,2)")
se.set_cell(0, 2, "=SUM(")
print(f"Div by zero: {se.get_value(0, 0)!r}")
print(f"Unknown function: {se.get_value(0, 1)!r}")
print(f"Syntax error: {se.get_value(0, 2)!r}")

banner("FINAL RESULT")
print("ALL INTEGRATION CHECKS PASSED" if all_ok else "SOME CHECKS FAILED")
sys.exit(0 if all_ok else 1)
