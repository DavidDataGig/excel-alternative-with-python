"""
Test suite for the spreadsheet engine.
Covers: A1 conversion, formula parsing, arithmetic, functions,
dependency tracking, undo/redo, structural ops, and I/O.
"""
import os
import sys
import tempfile
import traceback

sys.path.insert(0, "/home/claude")
from spreadsheet_engine import Spreadsheet, a1_to_rc, rc_to_a1, col_letter, col_index, Evaluator


# Simple test framework
class TestRunner:
    def __init__(self):
        self.passed = 0
        self.failed = 0
        self.errors = []

    def check(self, name, got, expected, tol=1e-9):
        ok = False
        if isinstance(expected, float) and isinstance(got, (int, float)):
            ok = abs(got - expected) < tol
        else:
            ok = got == expected
        if ok:
            self.passed += 1
            print(f"  PASS  {name}")
        else:
            self.failed += 1
            msg = f"  FAIL  {name}: got {got!r}, expected {expected!r}"
            self.errors.append(msg)
            print(msg)

    def check_raises(self, name, fn, *args, **kwargs):
        try:
            fn(*args, **kwargs)
            self.failed += 1
            print(f"  FAIL  {name}: did not raise")
            self.errors.append(f"{name}: did not raise")
        except Exception:
            self.passed += 1
            print(f"  PASS  {name}")

    def summary(self):
        total = self.passed + self.failed
        print(f"\n{'=' * 60}")
        print(f"RESULTS: {self.passed}/{total} passed, {self.failed} failed")
        print("=" * 60)
        return self.failed == 0


t = TestRunner()


# ---- Section 1: A1 notation ----
print("\n[Section 1] A1 notation helpers")
t.check("col_letter(0)", col_letter(0), "A")
t.check("col_letter(25)", col_letter(25), "Z")
t.check("col_letter(26)", col_letter(26), "AA")
t.check("col_letter(51)", col_letter(51), "AZ")
t.check("col_letter(52)", col_letter(52), "BA")
t.check("col_index('A')", col_index("A"), 0)
t.check("col_index('Z')", col_index("Z"), 25)
t.check("col_index('AA')", col_index("AA"), 26)
t.check("col_index('BA')", col_index("BA"), 52)
t.check("a1_to_rc('A1')", a1_to_rc("A1"), (0, 0))
t.check("a1_to_rc('B3')", a1_to_rc("B3"), (2, 1))
t.check("a1_to_rc('AA10')", a1_to_rc("AA10"), (9, 26))
t.check("rc_to_a1(0,0)", rc_to_a1(0, 0), "A1")
t.check("rc_to_a1(9,26)", rc_to_a1(9, 26), "AA10")


# ---- Section 2: Basic cell operations ----
print("\n[Section 2] Basic cell set/get")
s = Spreadsheet(10, 10)
s.set_cell(0, 0, "Hello")
t.check("string stored", s.get_value(0, 0), "Hello")
s.set_cell(0, 1, "42")
t.check("int parsed", s.get_value(0, 1), 42)
s.set_cell(0, 2, "3.14")
t.check("float parsed", s.get_value(0, 2), 3.14)
s.set_cell(0, 3, "")
t.check("empty clears", s.get_value(0, 3), "")


# ---- Section 3: Arithmetic formulas ----
print("\n[Section 3] Arithmetic formulas")
s = Spreadsheet(10, 10)
s.set_cell(0, 0, "10")
s.set_cell(0, 1, "20")
s.set_cell(0, 2, "=A1+B1")
t.check("A1+B1 = 30", s.get_value(0, 2), 30)
s.set_cell(0, 3, "=A1*B1")
t.check("A1*B1 = 200", s.get_value(0, 3), 200)
s.set_cell(0, 4, "=B1/A1")
t.check("B1/A1 = 2", s.get_value(0, 4), 2.0)
s.set_cell(0, 5, "=B1-A1")
t.check("B1-A1 = 10", s.get_value(0, 5), 10)
s.set_cell(0, 6, "=2^3")
t.check("2^3 = 8", s.get_value(0, 6), 8)
s.set_cell(0, 7, "=-A1")
t.check("-A1 = -10", s.get_value(0, 7), -10)
s.set_cell(0, 8, "=(A1+B1)*2")
t.check("(A1+B1)*2 = 60", s.get_value(0, 8), 60)
s.set_cell(1, 0, "=2+3*4")
t.check("operator precedence 2+3*4 = 14", s.get_value(1, 0), 14)


# ---- Section 4: Division by zero ----
print("\n[Section 4] Division by zero")
s = Spreadsheet(5, 5)
s.set_cell(0, 0, "=1/0")
t.check("#DIV/0!", s.get_value(0, 0), "#DIV/0!")
s.set_cell(0, 1, "0")
s.set_cell(0, 2, "=5/B1")
t.check("5/B1 where B1=0", s.get_value(0, 2), "#DIV/0!")


# ---- Section 5: Built-in functions ----
print("\n[Section 5] Built-in functions")
s = Spreadsheet(10, 10)
for i, v in enumerate([1, 2, 3, 4, 5]):
    s.set_cell(i, 0, str(v))
s.set_cell(0, 1, "=SUM(A1:A5)")
t.check("SUM(A1:A5) = 15", s.get_value(0, 1), 15)
s.set_cell(0, 2, "=AVERAGE(A1:A5)")
t.check("AVERAGE = 3", s.get_value(0, 2), 3.0)
s.set_cell(0, 3, "=MIN(A1:A5)")
t.check("MIN = 1", s.get_value(0, 3), 1)
s.set_cell(0, 4, "=MAX(A1:A5)")
t.check("MAX = 5", s.get_value(0, 4), 5)
s.set_cell(0, 5, "=COUNT(A1:A5)")
t.check("COUNT = 5", s.get_value(0, 5), 5)
s.set_cell(0, 6, "=PRODUCT(A1:A5)")
t.check("PRODUCT = 120", s.get_value(0, 6), 120)
s.set_cell(0, 7, "=ABS(-42)")
t.check("ABS(-42) = 42", s.get_value(0, 7), 42)
s.set_cell(0, 8, "=ROUND(3.14159, 2)")
t.check("ROUND(3.14159, 2) = 3.14", s.get_value(0, 8), 3.14)


# ---- Section 6: IF and comparisons ----
print("\n[Section 6] IF and comparisons")
s = Spreadsheet(5, 5)
s.set_cell(0, 0, "10")
s.set_cell(0, 1, '=IF(A1>5, "big", "small")')
t.check("IF(A1>5) -> big", s.get_value(0, 1), "big")
s.set_cell(0, 2, '=IF(A1<5, "big", "small")')
t.check("IF(A1<5) -> small", s.get_value(0, 2), "small")
s.set_cell(0, 3, "=IF(A1=10, 1, 0)")
t.check("IF equality", s.get_value(0, 3), 1)
s.set_cell(0, 4, "=IF(A1<>10, 1, 0)")
t.check("IF not equal", s.get_value(0, 4), 0)


# ---- Section 7: Mixed ranges with text ----
print("\n[Section 7] SUM skips text cells")
s = Spreadsheet(5, 5)
s.set_cell(0, 0, "5")
s.set_cell(1, 0, "hello")  # text should be skipped
s.set_cell(2, 0, "10")
s.set_cell(3, 0, "=SUM(A1:A3)")
t.check("SUM skips text", s.get_value(3, 0), 15)


# ---- Section 8: Dependency tracking (cascading recalc) ----
print("\n[Section 8] Dependency tracking")
s = Spreadsheet(5, 5)
s.set_cell(0, 0, "10")
s.set_cell(0, 1, "=A1*2")
s.set_cell(0, 2, "=B1+5")
t.check("initial B1 = 20", s.get_value(0, 1), 20)
t.check("initial C1 = 25", s.get_value(0, 2), 25)
# Change A1, B1 and C1 should both update
s.set_cell(0, 0, "100")
t.check("A1->100, B1 recalcs to 200", s.get_value(0, 1), 200)
t.check("A1->100, C1 recalcs to 205", s.get_value(0, 2), 205)


# ---- Section 9: Undo / redo ----
print("\n[Section 9] Undo / redo")
s = Spreadsheet(5, 5)
s.set_cell(0, 0, "first")
s.set_cell(0, 0, "second")
s.set_cell(0, 0, "third")
t.check("current = third", s.get_value(0, 0), "third")
s.undo()
t.check("after 1 undo = second", s.get_value(0, 0), "second")
s.undo()
t.check("after 2 undo = first", s.get_value(0, 0), "first")
s.redo()
t.check("after redo = second", s.get_value(0, 0), "second")
s.undo()
s.undo()
t.check("undo past start = empty", s.get_value(0, 0), "")


# ---- Section 10: Insert/delete rows and cols ----
print("\n[Section 10] Insert/delete rows & cols")
s = Spreadsheet(5, 5)
s.set_cell(0, 0, "row0")
s.set_cell(1, 0, "row1")
s.set_cell(2, 0, "row2")
s.insert_row(1)
t.check("after insert row at 1, row 0 unchanged", s.get_value(0, 0), "row0")
t.check("after insert, old row1 now at row2", s.get_value(2, 0), "row1")
t.check("newly inserted row1 is empty", s.get_value(1, 0), "")

s = Spreadsheet(5, 5)
s.set_cell(0, 0, "a")
s.set_cell(0, 1, "b")
s.set_cell(0, 2, "c")
s.delete_col(1)
t.check("after delete col 1, col 0 = a", s.get_value(0, 0), "a")
t.check("after delete col 1, col 1 = c (was col 2)", s.get_value(0, 1), "c")


# ---- Section 11: CSV I/O ----
print("\n[Section 11] CSV round-trip")
s = Spreadsheet(5, 5)
s.set_cell(0, 0, "Name")
s.set_cell(0, 1, "Score")
s.set_cell(1, 0, "Alice")
s.set_cell(1, 1, "95")
s.set_cell(2, 0, "Bob")
s.set_cell(2, 1, "87")
s.set_cell(3, 1, "=SUM(B2:B3)")
with tempfile.NamedTemporaryFile(suffix=".csv", delete=False, mode="w") as f:
    csv_path = f.name
s.save_csv(csv_path)
s2 = Spreadsheet(5, 5)
s2.load_csv(csv_path)
t.check("CSV reload A1", s2.get_value(0, 0), "Name")
t.check("CSV reload A2", s2.get_value(1, 0), "Alice")
t.check("CSV reload sum cell stored as number", s2.get_value(3, 1), 182)
os.unlink(csv_path)


# ---- Section 12: XLSX I/O ----
print("\n[Section 12] XLSX round-trip")
s = Spreadsheet(5, 5)
s.set_cell(0, 0, "Item")
s.set_cell(0, 1, "Qty")
s.set_cell(1, 0, "Widget")
s.set_cell(1, 1, "10")
s.set_cell(2, 0, "Gadget")
s.set_cell(2, 1, "5")
s.set_cell(3, 0, "Total")
s.set_cell(3, 1, "=SUM(B2:B3)")
with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
    xlsx_path = f.name
s.save_xlsx(xlsx_path)
s3 = Spreadsheet(5, 5)
s3.load_xlsx(xlsx_path)
t.check("XLSX reload A1", s3.get_value(0, 0), "Item")
t.check("XLSX reload preserves formula (recomputed)", s3.get_value(3, 1), 15)
os.unlink(xlsx_path)


# ---- Section 13: Complex financial example ----
print("\n[Section 13] Complex scenario: mini P&L")
s = Spreadsheet(10, 10)
# Headers
s.set_cell(0, 0, "Item")
s.set_cell(0, 1, "Q1")
s.set_cell(0, 2, "Q2")
s.set_cell(0, 3, "Q3")
s.set_cell(0, 4, "Q4")
s.set_cell(0, 5, "Total")
# Revenue row
s.set_cell(1, 0, "Revenue")
s.set_cell(1, 1, "1000")
s.set_cell(1, 2, "1200")
s.set_cell(1, 3, "1500")
s.set_cell(1, 4, "1800")
s.set_cell(1, 5, "=SUM(B2:E2)")
# Cost row
s.set_cell(2, 0, "Cost")
s.set_cell(2, 1, "600")
s.set_cell(2, 2, "700")
s.set_cell(2, 3, "850")
s.set_cell(2, 4, "1000")
s.set_cell(2, 5, "=SUM(B3:E3)")
# Profit row (formula-based)
s.set_cell(3, 0, "Profit")
s.set_cell(3, 1, "=B2-B3")
s.set_cell(3, 2, "=C2-C3")
s.set_cell(3, 3, "=D2-D3")
s.set_cell(3, 4, "=E2-E3")
s.set_cell(3, 5, "=SUM(B4:E4)")
# Margin row
s.set_cell(4, 0, "Margin %")
s.set_cell(4, 1, "=B4/B2")
s.set_cell(4, 5, "=F4/F2")

t.check("Revenue total", s.get_value(1, 5), 5500)
t.check("Cost total", s.get_value(2, 5), 3150)
t.check("Q1 Profit", s.get_value(3, 1), 400)
t.check("Total Profit", s.get_value(3, 5), 2350)
t.check("Q1 Margin", round(s.get_value(4, 1), 2), 0.4)

# Now change a revenue value and ensure everything cascades
s.set_cell(1, 1, "2000")  # Q1 revenue doubled
t.check("after Q1 rev=2000, Q1 profit", s.get_value(3, 1), 1400)
t.check("after Q1 rev=2000, total rev", s.get_value(1, 5), 6500)
t.check("after Q1 rev=2000, total profit", s.get_value(3, 5), 3350)


# ---- Section 14: Edge cases ----
print("\n[Section 14] Edge cases")
s = Spreadsheet(5, 5)
# Empty range
s.set_cell(0, 0, "=SUM(B1:B5)")
t.check("SUM of empty range = 0", s.get_value(0, 0), 0)

# Self-reference (direct circular) - should not crash, may yield 0 or #ERROR
s = Spreadsheet(5, 5)
s.set_cell(0, 0, "=A1+1")
# Just ensure it didn't infinite loop (we got here)
t.check("self-reference does not hang", True, True)

# Deep dependency chain
s = Spreadsheet(20, 5)
s.set_cell(0, 0, "1")
for i in range(1, 10):
    s.set_cell(i, 0, f"=A{i}+1")
t.check("deep chain A10 = 10", s.get_value(9, 0), 10)

# Now change A1 and check propagation through the chain
s.set_cell(0, 0, "100")
t.check("after A1=100, A10 = 109", s.get_value(9, 0), 109)


# ---- Done ----
ok = t.summary()
if not ok:
    print("\nErrors:")
    for e in t.errors:
        print(" ", e)
sys.exit(0 if ok else 1)
