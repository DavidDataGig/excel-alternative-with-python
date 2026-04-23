"""
Spreadsheet Engine - the brain behind the ipywidgets Excel alternative.
Separated from the UI so it can be unit-tested without a Jupyter kernel.

Supports:
- Cell storage with formulas and values
- Formula parser for =SUM, =AVERAGE, =MIN, =MAX, =COUNT, =IF, arithmetic, cell references, ranges
- A1 <-> (row,col) conversion
- Load/save CSV and XLSX
- Dependency tracking for recalc
- Undo/redo history
"""
from __future__ import annotations
import re
import csv
import copy
from typing import Any

# ---------- A1 notation helpers ----------

def col_letter(col: int) -> str:
    """0-indexed column number -> Excel letter (0='A', 25='Z', 26='AA')."""
    s = ""
    n = col
    while True:
        s = chr(ord("A") + n % 26) + s
        n = n // 26 - 1
        if n < 0:
            break
    return s

def col_index(letters: str) -> int:
    """'A' -> 0, 'Z' -> 25, 'AA' -> 26."""
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n - 1

def a1_to_rc(ref: str) -> tuple[int, int]:
    """'B3' -> (2, 1) as (row, col) 0-indexed."""
    m = re.fullmatch(r"([A-Za-z]+)(\d+)", ref.strip())
    if not m:
        raise ValueError(f"Bad cell reference: {ref}")
    letters, digits = m.groups()
    return int(digits) - 1, col_index(letters)

def rc_to_a1(row: int, col: int) -> str:
    return f"{col_letter(col)}{row + 1}"


# ---------- Formula evaluator ----------

class FormulaError(Exception):
    pass

# Functions we support. Operate on an already-flattened list of numeric values.
FUNCTIONS: dict[str, Any] = {
    "SUM":     lambda xs: sum(xs),
    "AVERAGE": lambda xs: (sum(xs) / len(xs)) if xs else FormulaError("#DIV/0!"),
    "MIN":     lambda xs: min(xs) if xs else FormulaError("#VALUE!"),
    "MAX":     lambda xs: max(xs) if xs else FormulaError("#VALUE!"),
    "COUNT":   lambda xs: len(xs),
    "PRODUCT": lambda xs: _product(xs),
    "ABS":     lambda xs: abs(xs[0]) if len(xs) == 1 else FormulaError("#VALUE!"),
    "ROUND":   lambda xs: round(xs[0], int(xs[1])) if len(xs) == 2 else FormulaError("#VALUE!"),
}

def _product(xs):
    p = 1
    for x in xs:
        p *= x
    return p


class Evaluator:
    """
    Recursive-descent evaluator for a small subset of Excel formula grammar.

    Grammar (informal):
        expr       := term (('+'|'-') term)*
        term       := power (('*'|'/') power)*
        power      := factor ('^' factor)?
        factor     := NUMBER | STRING | CELLREF | RANGE | FUNC '(' args ')' | '(' expr ')' | '-' factor
        args       := expr (',' expr)*   (for IF we allow bool-like first arg)
        CELLREF    := [A-Z]+[0-9]+
        RANGE      := CELLREF ':' CELLREF
    """

    def __init__(self, get_cell):
        # get_cell(row, col) -> value (number, string, or None)
        self.get_cell = get_cell
        self.deps: set[tuple[int, int]] = set()

    def evaluate(self, text: str) -> Any:
        self.deps = set()
        self.tokens = self._tokenize(text)
        self.pos = 0
        result = self._expr()
        if self.pos != len(self.tokens):
            raise FormulaError("#SYNTAX")
        return result

    # Tokenizer
    def _tokenize(self, s: str):
        tokens = []
        i = 0
        while i < len(s):
            c = s[i]
            if c.isspace():
                i += 1
            elif c.isdigit() or (c == "." and i + 1 < len(s) and s[i + 1].isdigit()):
                j = i
                while j < len(s) and (s[j].isdigit() or s[j] == "."):
                    j += 1
                tokens.append(("NUM", float(s[i:j])))
                i = j
            elif c == '"':
                j = i + 1
                while j < len(s) and s[j] != '"':
                    j += 1
                tokens.append(("STR", s[i + 1:j]))
                i = j + 1
            elif c.isalpha():
                j = i
                while j < len(s) and (s[j].isalnum() or s[j] == "_"):
                    j += 1
                word = s[i:j]
                # Is it a cell ref like A1, range start, or a function name?
                if re.fullmatch(r"[A-Za-z]+\d+", word):
                    tokens.append(("REF", word.upper()))
                elif word.upper() in ("TRUE", "FALSE"):
                    tokens.append(("BOOL", word.upper() == "TRUE"))
                else:
                    tokens.append(("FUNC", word.upper()))
                i = j
            elif c in "+-*/^(),:<>=":
                # Handle 2-char operators: <=, >=, <>
                if c in "<>" and i + 1 < len(s) and s[i + 1] == "=":
                    tokens.append(("OP", c + "="))
                    i += 2
                elif c == "<" and i + 1 < len(s) and s[i + 1] == ">":
                    tokens.append(("OP", "<>"))
                    i += 2
                else:
                    tokens.append(("OP", c))
                    i += 1
            else:
                raise FormulaError(f"#SYNTAX: unexpected '{c}'")
        return tokens

    def _peek(self, offset=0):
        if self.pos + offset < len(self.tokens):
            return self.tokens[self.pos + offset]
        return (None, None)

    def _eat(self):
        tok = self.tokens[self.pos]
        self.pos += 1
        return tok

    def _match(self, kind, value=None):
        t = self._peek()
        if t[0] == kind and (value is None or t[1] == value):
            return self._eat()
        return None

    # expr := compare (for IF-like comparisons)
    def _expr(self):
        left = self._additive()
        while self._peek()[0] == "OP" and self._peek()[1] in ("=", "<>", "<", ">", "<=", ">="):
            op = self._eat()[1]
            right = self._additive()
            left = self._compare(left, op, right)
        return left

    def _compare(self, a, op, b):
        if op == "=":  return a == b
        if op == "<>": return a != b
        if op == "<":  return a < b
        if op == ">":  return a > b
        if op == "<=": return a <= b
        if op == ">=": return a >= b

    def _additive(self):
        left = self._term()
        while self._peek()[0] == "OP" and self._peek()[1] in ("+", "-"):
            op = self._eat()[1]
            right = self._term()
            left = (self._num(left) + self._num(right)) if op == "+" else (self._num(left) - self._num(right))
        return left

    def _term(self):
        left = self._power()
        while self._peek()[0] == "OP" and self._peek()[1] in ("*", "/"):
            op = self._eat()[1]
            right = self._power()
            if op == "*":
                left = self._num(left) * self._num(right)
            else:
                r = self._num(right)
                if r == 0:
                    raise FormulaError("#DIV/0!")
                left = self._num(left) / r
        return left

    def _power(self):
        left = self._factor()
        if self._peek()[0] == "OP" and self._peek()[1] == "^":
            self._eat()
            right = self._factor()
            return self._num(left) ** self._num(right)
        return left

    def _factor(self):
        t = self._peek()
        if t[0] == "OP" and t[1] == "-":
            self._eat()
            return -self._num(self._factor())
        if t[0] == "OP" and t[1] == "+":
            self._eat()
            return self._num(self._factor())
        if t[0] == "NUM":
            return self._eat()[1]
        if t[0] == "STR":
            return self._eat()[1]
        if t[0] == "BOOL":
            return self._eat()[1]
        if t[0] == "OP" and t[1] == "(":
            self._eat()
            v = self._expr()
            if not self._match("OP", ")"):
                raise FormulaError("#SYNTAX: missing )")
            return v
        if t[0] == "REF":
            # Could be single ref OR start of range
            ref1 = self._eat()[1]
            if self._peek()[0] == "OP" and self._peek()[1] == ":":
                self._eat()
                if self._peek()[0] != "REF":
                    raise FormulaError("#SYNTAX: expected cell after ':'")
                ref2 = self._eat()[1]
                return self._resolve_range(ref1, ref2)
            return self._resolve_cell(ref1)
        if t[0] == "FUNC":
            name = self._eat()[1]
            if not self._match("OP", "("):
                raise FormulaError(f"#NAME?: {name}")
            args = []
            if not (self._peek()[0] == "OP" and self._peek()[1] == ")"):
                args.append(self._expr())
                while self._peek()[0] == "OP" and self._peek()[1] == ",":
                    self._eat()
                    args.append(self._expr())
            if not self._match("OP", ")"):
                raise FormulaError(f"#SYNTAX: missing ) in {name}")
            return self._call(name, args)
        raise FormulaError("#SYNTAX: unexpected token")

    def _resolve_cell(self, ref):
        row, col = a1_to_rc(ref)
        self.deps.add((row, col))
        v = self.get_cell(row, col)
        if v is None or v == "":
            return 0  # Excel treats empty cells as 0 in arithmetic
        if isinstance(v, str):
            # Try to parse a number from the string
            try:
                return float(v)
            except ValueError:
                return v
        return v

    def _resolve_range(self, r1, r2):
        rr1, cc1 = a1_to_rc(r1)
        rr2, cc2 = a1_to_rc(r2)
        r_start, r_end = min(rr1, rr2), max(rr1, rr2)
        c_start, c_end = min(cc1, cc2), max(cc1, cc2)
        values = []
        for r in range(r_start, r_end + 1):
            for c in range(c_start, c_end + 1):
                self.deps.add((r, c))
                v = self.get_cell(r, c)
                if v is None or v == "":
                    continue
                if isinstance(v, (int, float)):
                    values.append(float(v))
                elif isinstance(v, str):
                    try:
                        values.append(float(v))
                    except ValueError:
                        pass
        return values  # ranges return lists; functions will flatten

    def _call(self, name, args):
        # Special-case IF because it expects exactly 3 args and the first is a bool
        if name == "IF":
            if len(args) != 3:
                raise FormulaError("#VALUE!: IF needs 3 args")
            cond = args[0]
            if isinstance(cond, list):  # a range used as a bool is weird; take truthy length
                cond = bool(cond)
            return args[1] if cond else args[2]
        if name == "CONCAT" or name == "CONCATENATE":
            return "".join(str(a) for a in args)
        if name not in FUNCTIONS:
            raise FormulaError(f"#NAME?: {name}")
        # Flatten args: ranges are lists; singletons become single-element lists
        flat = []
        for a in args:
            if isinstance(a, list):
                flat.extend(a)
            elif isinstance(a, (int, float)):
                flat.append(float(a))
            elif isinstance(a, str):
                try:
                    flat.append(float(a))
                except ValueError:
                    pass  # skip non-numeric strings, matches Excel SUM behavior
            elif isinstance(a, bool):
                flat.append(1.0 if a else 0.0)
        result = FUNCTIONS[name](flat)
        if isinstance(result, FormulaError):
            raise result
        return result

    def _num(self, v):
        if isinstance(v, bool):
            return 1 if v else 0
        if isinstance(v, (int, float)):
            return v
        if isinstance(v, str):
            try:
                return float(v)
            except ValueError:
                raise FormulaError("#VALUE!")
        if isinstance(v, list):
            raise FormulaError("#VALUE!: range in arithmetic")
        raise FormulaError("#VALUE!")


# ---------- The Spreadsheet itself ----------

class Spreadsheet:
    def __init__(self, rows: int = 20, cols: int = 10):
        self.rows = rows
        self.cols = cols
        # Two parallel maps: raw (formulas as typed) and computed (evaluated values)
        self.raw: dict[tuple[int, int], str] = {}
        self.values: dict[tuple[int, int], Any] = {}
        self.deps_of: dict[tuple[int, int], set[tuple[int, int]]] = {}  # cell -> cells it depends on
        self.dependents: dict[tuple[int, int], set[tuple[int, int]]] = {}  # cell -> cells that depend on it
        self._history: list[dict] = []
        self._future: list[dict] = []
        self._history_limit = 50

    # ----- snapshot helpers for undo/redo -----
    def _snapshot(self):
        return {
            "raw": dict(self.raw),
            "values": dict(self.values),
            "deps_of": {k: set(v) for k, v in self.deps_of.items()},
            "dependents": {k: set(v) for k, v in self.dependents.items()},
            "rows": self.rows,
            "cols": self.cols,
        }

    def _restore(self, snap):
        self.raw = dict(snap["raw"])
        self.values = dict(snap["values"])
        self.deps_of = {k: set(v) for k, v in snap["deps_of"].items()}
        self.dependents = {k: set(v) for k, v in snap["dependents"].items()}
        self.rows = snap["rows"]
        self.cols = snap["cols"]

    def _push_history(self):
        self._history.append(self._snapshot())
        if len(self._history) > self._history_limit:
            self._history.pop(0)
        self._future.clear()

    def undo(self) -> bool:
        if not self._history:
            return False
        self._future.append(self._snapshot())
        self._restore(self._history.pop())
        return True

    def redo(self) -> bool:
        if not self._future:
            return False
        self._history.append(self._snapshot())
        self._restore(self._future.pop())
        return True

    # ----- core cell operations -----
    def set_cell(self, row: int, col: int, text: str, record_history: bool = True):
        if record_history:
            self._push_history()
        key = (row, col)
        # Clear old dependency edges
        if key in self.deps_of:
            for dep in self.deps_of[key]:
                self.dependents.get(dep, set()).discard(key)
            del self.deps_of[key]

        if text is None or text == "":
            self.raw.pop(key, None)
            self.values.pop(key, None)
        else:
            self.raw[key] = text

        self._recalc_cell(row, col)
        # Recalc everything that depends on this cell (transitively)
        self._recalc_dependents(key)

    def _recalc_cell(self, row: int, col: int):
        key = (row, col)
        text = self.raw.get(key, "")
        if text == "":
            self.values.pop(key, None)
            return
        if isinstance(text, str) and text.startswith("="):
            evaluator = Evaluator(lambda r, c: self.values.get((r, c)))
            try:
                result = evaluator.evaluate(text[1:])
                if isinstance(result, list):
                    # A formula that returns a raw range - show first element or error
                    result = result[0] if result else ""
                self.values[key] = result
                # Save new dependency edges
                self.deps_of[key] = evaluator.deps
                for dep in evaluator.deps:
                    self.dependents.setdefault(dep, set()).add(key)
            except FormulaError as e:
                self.values[key] = str(e)
            except Exception:
                self.values[key] = "#ERROR!"
        else:
            # Try numeric parse; otherwise keep as string
            try:
                self.values[key] = float(text) if "." in text or "e" in text.lower() else int(text)
            except (ValueError, TypeError):
                self.values[key] = text

    def _recalc_dependents(self, key):
        # BFS over dependents, guarding against cycles
        visited = set()
        queue = list(self.dependents.get(key, set()))
        while queue:
            cell = queue.pop(0)
            if cell in visited:
                continue
            visited.add(cell)
            self._recalc_cell(*cell)
            queue.extend(self.dependents.get(cell, set()))

    def recalculate_all(self):
        """Recompute every formula in dependency order (approximate via iteration)."""
        # Simple fixed-point: iterate up to N passes until values stop changing.
        for _ in range(5):
            changed = False
            for key in list(self.raw.keys()):
                before = self.values.get(key)
                self._recalc_cell(*key)
                if self.values.get(key) != before:
                    changed = True
            if not changed:
                break

    # ----- reading -----
    def get_raw(self, row: int, col: int) -> str:
        return self.raw.get((row, col), "")

    def get_value(self, row: int, col: int):
        return self.values.get((row, col), "")

    def get_display(self, row: int, col: int) -> str:
        v = self.values.get((row, col), "")
        if isinstance(v, float):
            # Tidy float formatting: drop trailing .0 for whole numbers
            if v.is_integer():
                return str(int(v))
            return f"{v:g}"
        return str(v)

    # ----- structural ops -----
    def insert_row(self, at: int):
        self._push_history()
        # Shift rows >= at down by 1
        new_raw = {}
        for (r, c), v in self.raw.items():
            new_raw[(r + 1 if r >= at else r, c)] = v
        self.raw = new_raw
        self.rows += 1
        self._rebuild_values()

    def delete_row(self, at: int):
        self._push_history()
        new_raw = {}
        for (r, c), v in self.raw.items():
            if r == at:
                continue
            new_raw[(r - 1 if r > at else r, c)] = v
        self.raw = new_raw
        self.rows = max(1, self.rows - 1)
        self._rebuild_values()

    def insert_col(self, at: int):
        self._push_history()
        new_raw = {}
        for (r, c), v in self.raw.items():
            new_raw[(r, c + 1 if c >= at else c)] = v
        self.raw = new_raw
        self.cols += 1
        self._rebuild_values()

    def delete_col(self, at: int):
        self._push_history()
        new_raw = {}
        for (r, c), v in self.raw.items():
            if c == at:
                continue
            new_raw[(r, c - 1 if c > at else c)] = v
        self.raw = new_raw
        self.cols = max(1, self.cols - 1)
        self._rebuild_values()

    def _rebuild_values(self):
        self.values.clear()
        self.deps_of.clear()
        self.dependents.clear()
        for key in list(self.raw.keys()):
            self._recalc_cell(*key)
        self.recalculate_all()

    # ----- I/O -----
    def load_from_2d(self, data: list[list]):
        self.raw.clear()
        self.values.clear()
        self.deps_of.clear()
        self.dependents.clear()
        if data:
            self.rows = max(self.rows, len(data))
            self.cols = max(self.cols, max(len(row) for row in data))
        for r, row in enumerate(data):
            for c, val in enumerate(row):
                if val is None or val == "":
                    continue
                self.set_cell(r, c, str(val), record_history=False)
        self.recalculate_all()
        self._history.clear()
        self._future.clear()

    def to_2d(self, use_values: bool = True) -> list[list]:
        out = [["" for _ in range(self.cols)] for _ in range(self.rows)]
        source = self.values if use_values else self.raw
        for (r, c), v in source.items():
            if r < self.rows and c < self.cols:
                out[r][c] = v
        return out

    def save_csv(self, path: str):
        data = self.to_2d(use_values=True)
        with open(path, "w", newline="") as f:
            writer = csv.writer(f)
            writer.writerows(data)

    def load_csv(self, path: str):
        with open(path, newline="") as f:
            rows = list(csv.reader(f))
        self.load_from_2d(rows)

    def save_xlsx(self, path: str):
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        for (r, c), text in self.raw.items():
            # openpyxl uses 1-indexed
            ws.cell(row=r + 1, column=c + 1, value=text)
        wb.save(path)

    def load_xlsx(self, path: str):
        from openpyxl import load_workbook
        wb = load_workbook(path)
        ws = wb.active
        data = []
        for row in ws.iter_rows(values_only=True):
            data.append(list(row))
        self.load_from_2d(data)
