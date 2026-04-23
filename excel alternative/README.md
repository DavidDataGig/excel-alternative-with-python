# MiniExcel — ipywidgets Excel alternative

A Jupyter-native spreadsheet with a real formula engine, built on `ipywidgets`.

## Files

| File | Purpose |
|---|---|
| `MiniExcel.ipynb` | **Main deliverable.** Self-contained notebook. Open in Jupyter, run both cells, you get an interactive spreadsheet pre-loaded with sample sales data. |
| `spreadsheet_engine.py` | Standalone engine module. The notebook embeds a copy, but this file is what the tests import. |
| `test_engine.py` | Unit tests (72 assertions across 14 sections). |
| `test_integration.py` | End-to-end scenario test — loads sample, edits, undoes, saves, reloads, stress-tests a 100-cell dependency chain. |

## How to run the notebook

```bash
pip install ipywidgets pandas openpyxl
jupyter notebook MiniExcel.ipynb
```

In JupyterLab ≥ 3, ipywidgets works out of the box. In classic Jupyter you may need `jupyter nbextension enable --py widgetsnbextension`.

Then: Cell → Run All. The UI appears below the second code cell.

## Features implemented

**Grid**
- Editable `Text` widget per cell (14 rows × 8 columns by default; configurable)
- Row numbers (1, 2, 3…) and column letters (A, B, C… AA, AB…) as headers
- Formula bar that shows the raw formula for the selected cell

**Formula engine** (recursive-descent parser, no `eval`)
- Arithmetic: `+ - * / ^`, unary minus, parentheses, correct operator precedence
- Comparisons: `= <> < > <= >=`
- Cell references: `A1`, `AA10`
- Ranges: `A1:B5`
- Functions: `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `PRODUCT`, `ABS`, `ROUND`, `IF`, `CONCAT` / `CONCATENATE`
- Excel-style error values: `#DIV/0!`, `#VALUE!`, `#NAME?`, `#SYNTAX`
- Strings in double quotes: `=IF(A1>5, "big", "small")`

**Recalculation**
- Dependency graph tracked per formula
- Editing a cell automatically recomputes every cell that transitively depends on it (BFS over the dependents graph)

**Editing**
- Undo / Redo (up to 50 steps)
- Insert row / delete row (shifts data; formula refs are NOT rewritten — see Limitations)
- Insert column / delete column
- Clear All

**I/O**
- Save as CSV (writes computed values)
- Save as XLSX (writes raw cells including formulas, so Excel/LibreOffice will evaluate them on open)
- Load CSV / XLSX (available programmatically; UI file picker not wired up)

**Status bar** at the bottom shows what just happened ("Set B2 = '3000' -> 3000", "Saved to foo.xlsx", etc.).

## Test results

### Unit tests — `python3 test_engine.py`

```
RESULTS: 72/72 passed, 0 failed
```

Covered:
- A1 ↔ (row, col) conversion, including 2-letter columns (AA, BA)
- Basic cell set/get with type inference
- Arithmetic, precedence, parentheses, unary minus, exponentiation
- Division by zero (both `=1/0` and via cell reference)
- All 8 built-in functions + IF + CONCAT
- Comparisons used inside IF
- SUM correctly skips non-numeric cells in a range
- Dependency cascading (change a source, dependents update)
- Undo / redo, including past the history floor
- Insert / delete rows and columns shift data correctly
- CSV round-trip
- XLSX round-trip (formulas preserved, reloaded, and re-evaluated)
- A 10-row sales scenario with totals, margins, and a full cascade test
- Edge cases: empty range sums to 0, self-reference doesn't hang, 10-deep dependency chain

### Integration test — `python3 test_integration.py`

```
ALL INTEGRATION CHECKS PASSED
```

Simulates the real user flow:
1. Loads the 11-row sample P&L dataset
2. Verifies 13 calculated values (row totals, grand total, averages, growth %, IF-based label)
3. Changes Widgets Q1 from 1200 → 3000 and verifies 6 cascading recalculations
4. Undoes, verifies rollback
5. Redoes, verifies forward
6. Saves XLSX, reloads from disk, verifies values survived (~5 KB file)
7. Inserts a row, verifies data shift
8. Stress test: 100-cell dependency chain. Setting `A1=500` correctly propagates to `A100=599`
9. Error cases render correctly: `#DIV/0!`, `#NAME?: UNKNOWN`, `#SYNTAX: unexpected token`

## Known limitations and additional feedback

These are things I'd fix before calling this "production":

1. **Formula references don't auto-adjust on insert/delete.** If B6 contains `=SUM(B2:B5)` and you insert a row at position 2, the data shifts down but the formula still points at `B2:B5` — so it now sums the wrong cells. Real Excel rewrites references when rows/cols are inserted or deleted. Fixing this means parsing every formula's AST, shifting each reference, and re-serializing. Doable, just more work.

2. **No focus event on Text widgets.** `ipywidgets.Text` doesn't expose a focus callback, so the formula bar updates only after you finish typing in a cell, not the moment you click it. A workaround would be wrapping each cell in a `Button` with `on_click` to set selection before editing, but that adds a lot of visual noise. The pragmatic fix: add a small "select cell" dropdown next to the formula bar.

3. **`ipywidgets.Text` doesn't support cell navigation with arrow keys.** Each cell is an independent input box. Tab moves focus across widgets but the order depends on the notebook's widget tree. Proper arrow-key navigation would need a custom `DOMWidget` or a switch to a single HTML `<table>` rendered via `ipywidgets.HTML` plus JS callbacks — much bigger lift.

4. **No cell formatting** (bold, color, number format, borders). Everything renders as plain text. Could be added with CSS classes on `Text` widgets but the ipywidgets styling surface is limited.

5. **Self-referencing formulas (`A1 = =A1+1`) are detected but handled only by not infinite-looping** — the result is "0" or the stale value. Real Excel shows `#CIRCULAR!` or uses iterative calc. I chose the safe minimum.

6. **Performance**: fine up to ~1000 cells. Each cell edit does a full dependency BFS, and rendering ~14×8 Text widgets is already slow to initialize in Jupyter (1–2 seconds). A grid of 1000×100 via raw ipywidgets would be painful — that's where you'd want a proper JS frontend.

7. **The "sample data" loads fine but row 7 is blank on purpose.** When saving to CSV this becomes an empty line, which some CSV readers treat as end-of-file. If you plan to share the CSVs, add a header or remove the blank row.

8. **File save location**: files land in the notebook's current working directory (whichever folder you launched Jupyter from). Consider passing absolute paths in the filename field if you want them elsewhere.

9. **If you want a more polished grid**, `ipydatagrid` (`pip install ipydatagrid`) gives you a real virtualized grid widget. You could plug the `Spreadsheet` engine from `spreadsheet_engine.py` straight into it — the engine is UI-agnostic.

## Quick verification without Jupyter

You can exercise the engine from a plain Python shell:

```python
from spreadsheet_engine import Spreadsheet
s = Spreadsheet(10, 5)
s.set_cell(0, 0, "10")
s.set_cell(0, 1, "=A1*2")
print(s.get_value(0, 1))    # 20
s.set_cell(0, 0, "50")
print(s.get_value(0, 1))    # 100
s.save_xlsx("out.xlsx")
```
