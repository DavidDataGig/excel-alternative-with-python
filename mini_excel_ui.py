"""
MiniExcel UI — high-performance ipywidgets spreadsheet.

Key design: a fixed *widget pool* is instantiated once per `_render_grid` call
(triggered only on structural changes such as config edits or data loads).
Page navigation calls the much cheaper `_update_page_data`, which only mutates
`.value` / `.options` / `layout.display` on existing widgets — zero new objects
are created, so the kernel↔frontend Comm channel stays quiet.
"""

from __future__ import annotations

import csv
import io
from typing import Any

import ipywidgets as widgets
from IPython.display import display

from spreadsheet_engine import (
    FormulaError,  # noqa: F401 — re-exported for convenience
    Spreadsheet,
    col_index,
    col_letter,
    rc_to_a1,
)

# (HBox row container, row-number label, ordered list of cell widgets for each column)
_PoolRow = tuple[widgets.HBox, widgets.Label, list[widgets.Widget]]


class MiniExcelUI:
    """
    Spreadsheet UI backed by a static widget pool.

    Structural operations (load, config change, col/row insert/delete, page-size
    change) call `_render_grid`, which rebuilds the pool once.  Pagination calls
    `_update_page_data`, which runs in O(page_size × cols) attribute assignments
    with no widget allocation.
    """

    _FREEZE_CSS = """
<style>
.mini-excel-freeze-header {
    position: sticky !important;
    top: 0;
    z-index: 101;
    background: #f0f0f0 !important;
}
.mini-excel-freeze-header .widget-label { background: #f0f0f0; }
.mini-excel-freeze-row1 {
    position: sticky !important;
    top: 28px;
    z-index: 100;
    background: #fffde7 !important;
}
</style>
"""

    def __init__(self, rows: int = 12, cols: int = 8) -> None:
        self.sheet = Spreadsheet(rows, cols)
        self.selected: tuple[int, int] = (0, 0)
        self._programmatic_update = False

        self.col_dropdowns: dict[int, list[str]] = {}
        self.col_widths: dict[int, int] = {}
        self.hidden_cols: set[int] = set()
        self.editable_cols: set[int] = set()

        self.page = 0
        self.page_size = 50
        self.freeze_top = False
        self.grid_height = 400
        self.auto_save = False

        # Widget pool state — rebuilt by _render_grid, reused by _update_page_data
        self._pool_rows: list[_PoolRow] = []
        self.col_header_widgets: dict[int, widgets.Label] = {}
        self._frozen_row_labels: dict[int, widgets.Label] = {}
        self._frozen_row_hbox: widgets.HBox | None = None

        self._build()

    # ── Layout helpers ─────────────────────────────────────────────────────

    def _col_w(self, c: int) -> str:
        return f"{self.col_widths.get(c, 90)}px"

    def _cell_layout(self, c: int) -> widgets.Layout:
        w = self._col_w(c)
        return widgets.Layout(width=w, min_width=w, flex="0 0 auto")

    @staticmethod
    def _fixed40() -> widgets.Layout:
        return widgets.Layout(
            width="40px", min_width="40px", flex="0 0 auto", border="1px solid #ccc"
        )

    def _apply_col_width(self, c: int, px: int) -> None:
        """Update column width on all live pool widgets without a full rebuild."""
        w_str = f"{px}px"
        if c in self.col_header_widgets:
            h = self.col_header_widgets[c]
            h.layout.width = h.layout.min_width = w_str
        if c in self._frozen_row_labels:
            lbl = self._frozen_row_labels[c]
            lbl.layout.width = lbl.layout.min_width = w_str
        for _, _, cell_pool in self._pool_rows:
            if c < len(cell_pool):
                w = cell_pool[c]
                w.layout.width = w.layout.min_width = w_str

    def _apply_col_visibility(self, c: int, hidden: bool) -> None:
        """Show or hide a column across header, frozen row, and pool."""
        val = "none" if hidden else ""
        if c in self.col_header_widgets:
            self.col_header_widgets[c].layout.display = val
        if c in self._frozen_row_labels:
            self._frozen_row_labels[c].layout.display = val
        for _, _, cell_pool in self._pool_rows:
            if c < len(cell_pool):
                cell_pool[c].layout.display = val

    # ── One-time UI construction ───────────────────────────────────────────

    def _build(self) -> None:
        """Create all toolbar widgets and scroll wrapper. Called once at init."""

        # Row 0 — paste / load data
        self.input_area = widgets.Textarea(
            placeholder="Paste CSV or tab-separated data here, then click Load…",
            layout=widgets.Layout(width="560px", height="60px"),
        )
        self.load_btn = widgets.Button(
            description="Load", button_style="info", layout=widgets.Layout(width="70px")
        )
        self.load_btn.on_click(self._on_load_data)
        toolbar0 = widgets.HBox(
            [widgets.Label("Input data:", layout=widgets.Layout(width="80px")),
             self.input_area, self.load_btn]
        )

        # Row 1 — save
        self.path_input = widgets.Text(
            value="miniexcel_output",
            placeholder="filename (no extension)",
            layout=widgets.Layout(width="220px"),
        )
        self.save_csv_btn  = widgets.Button(description="Save CSV",  layout=widgets.Layout(width="90px"))
        self.save_xlsx_btn = widgets.Button(description="Save XLSX", layout=widgets.Layout(width="90px"))
        self.load_csv_btn  = widgets.Button(description="Load CSV",  button_style="info", layout=widgets.Layout(width="90px"))
        self.load_xlsx_btn = widgets.Button(description="Load XLSX", button_style="info", layout=widgets.Layout(width="90px"))
        self.save_csv_btn.on_click(self._on_save_csv)
        self.save_xlsx_btn.on_click(self._on_save_xlsx)
        self.load_csv_btn.on_click(self._on_load_csv)
        self.load_xlsx_btn.on_click(self._on_load_xlsx)
        toolbar1 = widgets.HBox(
            [widgets.Label("File:", layout=widgets.Layout(width="35px")),
             self.path_input,
             self.save_csv_btn, self.save_xlsx_btn,
             self.load_csv_btn, self.load_xlsx_btn]
        )

        # Row 2 — formula bar
        self.cell_label = widgets.Label(value="A1", layout=widgets.Layout(width="50px"))
        self.formula_bar = widgets.Text(
            value="",
            continuous_update=False,
            placeholder="Enter value or formula (e.g. =SUM(A1:A5))",
            layout=widgets.Layout(width="500px"),
        )
        self.formula_bar.observe(self._on_formula_bar_change, names="value")
        toolbar2 = widgets.HBox([self.cell_label, self.formula_bar])

        # Row 3 — structural edit buttons
        def _btn(desc: str, **kw) -> widgets.Button:
            return widgets.Button(description=desc, **kw)

        self.undo_btn    = _btn("Undo", icon="undo",    layout=widgets.Layout(width="90px"))
        self.redo_btn    = _btn("Redo", icon="redo",    layout=widgets.Layout(width="90px"))
        self.ins_row_btn = _btn("+Row", tooltip="Insert row above selected",   layout=widgets.Layout(width="70px"))
        self.del_row_btn = _btn("-Row", tooltip="Delete selected row",         layout=widgets.Layout(width="70px"))
        self.ins_col_btn = _btn("+Col", tooltip="Insert column before selected", layout=widgets.Layout(width="70px"))
        self.del_col_btn = _btn("-Col", tooltip="Delete selected column",      layout=widgets.Layout(width="70px"))
        self.clear_btn   = _btn("Clear All", button_style="warning",           layout=widgets.Layout(width="100px"))
        self.undo_btn.on_click(self._on_undo)
        self.redo_btn.on_click(self._on_redo)
        self.ins_row_btn.on_click(self._on_ins_row)
        self.del_row_btn.on_click(self._on_del_row)
        self.ins_col_btn.on_click(self._on_ins_col)
        self.del_col_btn.on_click(self._on_del_col)
        self.clear_btn.on_click(self._on_clear)
        toolbar3 = widgets.HBox(
            [self.undo_btn, self.redo_btn,
             self.ins_row_btn, self.del_row_btn,
             self.ins_col_btn, self.del_col_btn,
             self.clear_btn]
        )

        # Row 4 — dropdown config
        self.add_dd_btn = widgets.Button(
            description="Add Dropdown", icon="list", layout=widgets.Layout(width="130px")
        )
        self.add_dd_btn.on_click(self._on_toggle_dd_config)
        self.dd_col_input   = widgets.Text(placeholder="Column (e.g. A)", layout=widgets.Layout(width="100px"))
        self.dd_items_input = widgets.Text(placeholder="Items: Yes, No, Maybe", layout=widgets.Layout(width="260px"))
        self.dd_apply_btn   = widgets.Button(description="Apply",  button_style="success", layout=widgets.Layout(width="70px"))
        self.dd_remove_btn  = widgets.Button(description="Remove", button_style="danger",  layout=widgets.Layout(width="80px"))
        self.dd_apply_btn.on_click(self._on_apply_dropdown)
        self.dd_remove_btn.on_click(self._on_remove_dropdown)
        self.dd_config_row = widgets.HBox(
            [widgets.Label("Column:", layout=widgets.Layout(width="58px")),
             self.dd_col_input,
             widgets.Label("Items:", layout=widgets.Layout(width="43px")),
             self.dd_items_input, self.dd_apply_btn, self.dd_remove_btn],
            layout=widgets.Layout(display="none"),
        )
        toolbar4 = widgets.HBox([self.add_dd_btn, self.dd_config_row])

        # Row 5 — column width
        self.col_w_btn = widgets.Button(
            description="Column Width", icon="arrows-h", layout=widgets.Layout(width="130px")
        )
        self.col_w_btn.on_click(self._on_toggle_cw_config)
        self.cw_col_input   = widgets.Text(placeholder="Column (e.g. A)", layout=widgets.Layout(width="100px"))
        self.cw_width_input = widgets.BoundedIntText(value=90, min=30, max=600, step=10, layout=widgets.Layout(width="75px"))
        self.cw_apply_btn   = widgets.Button(description="Apply", button_style="success", layout=widgets.Layout(width="70px"))
        self.cw_reset_btn   = widgets.Button(description="Reset",                         layout=widgets.Layout(width="70px"))
        self.cw_apply_btn.on_click(self._on_apply_col_width)
        self.cw_reset_btn.on_click(self._on_reset_col_width)
        self.cw_config_row = widgets.HBox(
            [widgets.Label("Column:", layout=widgets.Layout(width="58px")),
             self.cw_col_input,
             widgets.Label("Width px:", layout=widgets.Layout(width="65px")),
             self.cw_width_input, self.cw_apply_btn, self.cw_reset_btn],
            layout=widgets.Layout(display="none"),
        )
        toolbar5 = widgets.HBox([self.col_w_btn, self.cw_config_row])

        # Row 6 — hide / unhide column
        self.hide_col_btn = widgets.Button(
            description="Hide/Unhide Col", icon="eye-slash", layout=widgets.Layout(width="150px")
        )
        self.hide_col_btn.on_click(self._on_toggle_hide_config)
        self.hc_col_input  = widgets.Text(placeholder="Column (e.g. A)", layout=widgets.Layout(width="100px"))
        self.hc_hide_btn   = widgets.Button(description="Hide",   button_style="warning", layout=widgets.Layout(width="70px"))
        self.hc_unhide_btn = widgets.Button(description="Unhide", button_style="success", layout=widgets.Layout(width="80px"))
        self.hc_hide_btn.on_click(self._on_hide_col)
        self.hc_unhide_btn.on_click(self._on_unhide_col)
        self.hc_config_row = widgets.HBox(
            [widgets.Label("Column:", layout=widgets.Layout(width="58px")),
             self.hc_col_input, self.hc_hide_btn, self.hc_unhide_btn],
            layout=widgets.Layout(display="none"),
        )
        toolbar6 = widgets.HBox([self.hide_col_btn, self.hc_config_row])

        # Row 7 — pagination
        self.page_size_dd = widgets.Dropdown(
            options=[10, 20, 30, 40, 50], value=self.page_size, layout=widgets.Layout(width="65px")
        )
        self.page_size_dd.observe(self._on_page_size_change, names="value")
        self.prev_btn   = widgets.Button(description="◀ Prev", layout=widgets.Layout(width="85px"))
        self.next_btn   = widgets.Button(description="Next ▶", layout=widgets.Layout(width="85px"))
        self.page_label = widgets.Label(value="Page 1 of 1", layout=widgets.Layout(width="110px"))
        self.prev_btn.on_click(self._on_prev_page)
        self.next_btn.on_click(self._on_next_page)
        toolbar7 = widgets.HBox(
            [widgets.Label("Rows/page:", layout=widgets.Layout(width="75px")),
             self.page_size_dd, self.prev_btn, self.page_label, self.next_btn]
        )

        # Row 8 — freeze top row
        self.freeze_btn = widgets.ToggleButton(
            value=False, description="Freeze Top Row", icon="lock",
            layout=widgets.Layout(width="145px"),
        )
        self.freeze_btn.observe(self._on_freeze_toggle, names="value")
        self.freeze_height_input = widgets.BoundedIntText(
            value=self.grid_height, min=100, max=800, step=50, layout=widgets.Layout(width="75px")
        )
        self.freeze_height_apply = widgets.Button(
            description="Apply", button_style="success", layout=widgets.Layout(width="70px")
        )
        self.freeze_height_apply.on_click(self._on_apply_grid_height)
        self.freeze_config_row = widgets.HBox(
            [widgets.Label("Height px:", layout=widgets.Layout(width="72px")),
             self.freeze_height_input, self.freeze_height_apply],
            layout=widgets.Layout(display="none"),
        )
        toolbar8 = widgets.HBox([self.freeze_btn, self.freeze_config_row])

        # Row 9 — column edit whitelist
        self.col_edit_btn = widgets.Button(
            description="Column Edit", icon="pencil", layout=widgets.Layout(width="130px")
        )
        self.col_edit_btn.on_click(self._on_toggle_col_edit_config)
        self.ce_col_input       = widgets.Text(placeholder="Column (e.g. A)", layout=widgets.Layout(width="100px"))
        self.ce_enable_btn      = widgets.Button(description="Enable",      button_style="success", layout=widgets.Layout(width="75px"))
        self.ce_disable_btn     = widgets.Button(description="Disable",     button_style="warning", layout=widgets.Layout(width="75px"))
        self.ce_enable_all_btn  = widgets.Button(description="Enable All",  button_style="info",    layout=widgets.Layout(width="90px"))
        self.ce_disable_all_btn = widgets.Button(description="Disable All", button_style="danger",  layout=widgets.Layout(width="95px"))
        self.ce_enable_btn.on_click(self._on_enable_col_edit)
        self.ce_disable_btn.on_click(self._on_disable_col_edit)
        self.ce_enable_all_btn.on_click(self._on_enable_all_cols)
        self.ce_disable_all_btn.on_click(self._on_disable_all_cols)
        self.ce_config_row = widgets.HBox(
            [widgets.Label("Column:", layout=widgets.Layout(width="58px")),
             self.ce_col_input,
             self.ce_enable_btn, self.ce_disable_btn,
             self.ce_enable_all_btn, self.ce_disable_all_btn],
            layout=widgets.Layout(display="none"),
        )
        toolbar9 = widgets.HBox([self.col_edit_btn, self.ce_config_row])

        # Row 10 — auto-save
        self.auto_save_btn = widgets.ToggleButton(
            value=False, description="Auto Save", icon="floppy-o",
            layout=widgets.Layout(width="120px"),
        )
        self.auto_save_btn.observe(self._on_auto_save_toggle, names="value")
        self.auto_save_fmt    = widgets.Dropdown(options=["CSV", "XLSX"], value="CSV", layout=widgets.Layout(width="75px"))
        self.auto_save_status = widgets.Label(value="Off", layout=widgets.Layout(width="260px"))
        toolbar10 = widgets.HBox(
            [self.auto_save_btn,
             widgets.Label("Format:", layout=widgets.Layout(width="55px")),
             self.auto_save_fmt, self.auto_save_status]
        )

        self.freeze_css = widgets.HTML(value=self._FREEZE_CSS)
        self.status = widgets.HTML(
            value='<span style="color:#888">Ready. All columns locked — use Column Edit to enable editing.</span>'
        )
        self.grid_box = widgets.VBox([])
        self.scroll_wrap = widgets.Box(
            [self.grid_box],
            layout=widgets.Layout(width="100%", overflow_x="auto", display="block"),
        )

        self._render_grid()

        self.container = widgets.VBox(
            [self.freeze_css,
             toolbar0, toolbar1, toolbar2, toolbar3,
             toolbar4, toolbar5, toolbar6, toolbar7, toolbar8, toolbar9,
             toolbar10,
             self.scroll_wrap, self.status]
        )

    # ── Pagination ─────────────────────────────────────────────────────────

    def _total_pages(self) -> int:
        return max(1, -(-self.sheet.rows // self.page_size))

    def _update_page_indicator(self) -> None:
        total = self._total_pages()
        self.page_label.value = f"Page {self.page + 1} of {total}"
        self.prev_btn.disabled = self.page == 0
        self.next_btn.disabled = self.page >= total - 1

    def _on_prev_page(self, _: Any) -> None:
        if self.page > 0:
            self._auto_save()
            self.page -= 1
            self._update_page_data()   # ← fast path: values only

    def _on_next_page(self, _: Any) -> None:
        if self.page < self._total_pages() - 1:
            self._auto_save()
            self.page += 1
            self._update_page_data()   # ← fast path: values only

    def _on_page_size_change(self, change: dict) -> None:
        self.page_size = change["new"]
        self.page = 0
        self._render_grid()            # pool size changes → full rebuild

    # ── Auto-save ──────────────────────────────────────────────────────────

    def _on_auto_save_toggle(self, change: dict) -> None:
        self.auto_save = change["new"]
        if self.auto_save:
            self.auto_save_btn.button_style = "success"
            self.auto_save_status.value = "On — will save on every page turn"
        else:
            self.auto_save_btn.button_style = ""
            self.auto_save_status.value = "Off"

    def _auto_save(self) -> None:
        if not self.auto_save:
            return
        fmt  = self.auto_save_fmt.value
        path = (self.path_input.value or "miniexcel_output") + "." + fmt.lower()
        self.sheet.save_csv(path) if fmt == "CSV" else self.sheet.save_xlsx(path)
        self.auto_save_status.value = f"Auto-saved → {path}"
        self._set_status(f"Auto-saved → {path}")

    # ── Freeze helpers ─────────────────────────────────────────────────────

    def _on_freeze_toggle(self, change: dict) -> None:
        self.freeze_top = change["new"]
        if self.freeze_top:
            self.scroll_wrap.layout.max_height = f"{self.grid_height}px"
            self.scroll_wrap.layout.overflow_y = "auto"
            self.freeze_config_row.layout.display = ""
            self.freeze_btn.button_style = "info"
        else:
            self.scroll_wrap.layout.max_height = ""
            self.scroll_wrap.layout.overflow_y = ""
            self.freeze_config_row.layout.display = "none"
            self.freeze_btn.button_style = ""
        self._render_grid()

    def _on_apply_grid_height(self, _: Any) -> None:
        self.grid_height = self.freeze_height_input.value
        if self.freeze_top:
            self.scroll_wrap.layout.max_height = f"{self.grid_height}px"
        self._set_status(f"Grid height set to {self.grid_height}px.")

    # ── Column-edit whitelist ──────────────────────────────────────────────

    def _on_toggle_col_edit_config(self, _: Any) -> None:
        cur = self.ce_config_row.layout.display
        self.ce_config_row.layout.display = "none" if cur != "none" else ""

    def _on_enable_col_edit(self, _: Any) -> None:
        c = self._parse_col_input(self.ce_col_input)
        if c is None: return
        self.editable_cols.add(c)
        self._render_grid()
        self._set_status(f"Column {col_letter(c)} is now editable.")

    def _on_disable_col_edit(self, _: Any) -> None:
        c = self._parse_col_input(self.ce_col_input)
        if c is None: return
        self.editable_cols.discard(c)
        self._render_grid()
        self._set_status(f"Column {col_letter(c)} is now read-only.")

    def _on_enable_all_cols(self, _: Any) -> None:
        self.editable_cols = set(range(self.sheet.cols))
        self._render_grid()
        self._set_status("All columns are now editable.")

    def _on_disable_all_cols(self, _: Any) -> None:
        self.editable_cols.clear()
        self._render_grid()
        self._set_status("All columns are now read-only.")

    # ── Core grid: full rebuild ────────────────────────────────────────────

    def _render_grid(self) -> None:
        """
        Rebuild the entire widget pool.  Expensive — only called on structural
        changes (config edits, data load, col/row mutations, page-size change).
        Pagination uses _update_page_data instead.
        """
        row_start   = self.page * self.page_size
        row_end     = min(row_start + self.page_size, self.sheet.rows)
        num_visible = row_end - row_start
        row_layout  = widgets.Layout(flex_flow="row nowrap")

        # ── Column-letter header row ───────────────────────────────────────
        corner = widgets.Label(value="", layout=self._fixed40())
        self.col_header_widgets.clear()
        header_cells: list[widgets.Widget] = [corner]
        for c in range(self.sheet.cols):
            cw = self._col_w(c)
            h = widgets.Label(
                value=col_letter(c),
                layout=widgets.Layout(width=cw, min_width=cw, flex="0 0 auto", border="1px solid #ccc"),
            )
            self.col_header_widgets[c] = h
            header_cells.append(h)
        header_row = widgets.HBox(header_cells, layout=row_layout)
        if self.freeze_top:
            header_row.add_class("mini-excel-freeze-header")

        # ── Frozen copy of data-row 0 (shown when freeze_top and page > 0) ─
        self._frozen_row_labels.clear()
        pin_cells: list[widgets.Widget] = [widgets.Label(value="1", layout=self._fixed40())]
        for c in range(self.sheet.cols):
            cw = self._col_w(c)
            lbl = widgets.Label(
                value=self.sheet.get_display(0, c),
                layout=widgets.Layout(
                    width=cw, min_width=cw, flex="0 0 auto",
                    border="1px solid #ccc", background_color="#fffde7",
                ),
            )
            pin_cells.append(lbl)
            self._frozen_row_labels[c] = lbl
        self._frozen_row_hbox = widgets.HBox(pin_cells, layout=row_layout)
        self._frozen_row_hbox.add_class("mini-excel-freeze-row1")
        # Shown/hidden by _update_page_data; start in the correct state
        self._frozen_row_hbox.layout.display = (
            "" if (self.freeze_top and self.page > 0 and self.sheet.rows > 0) else "none"
        )

        # ── Widget pool (page_size slots) ──────────────────────────────────
        self._pool_rows = []
        for pool_idx in range(self.page_size):
            r       = row_start + pool_idx
            visible = pool_idx < num_visible

            rn_label = widgets.Label(
                value=str(r + 1) if visible else "",
                layout=self._fixed40(),
            )
            cell_pool: list[widgets.Widget] = []
            row_cells: list[widgets.Widget] = [rn_label]

            for c in range(self.sheet.cols):
                current_val = self.sheet.get_display(r, c) if visible else ""
                layout      = self._cell_layout(c)

                if c in self.editable_cols:
                    if c in self.col_dropdowns:
                        opts = self.col_dropdowns[c]
                        dd_opts = [""] + opts
                        if current_val and current_val not in dd_opts:
                            dd_opts = ["", current_val] + opts
                        w: widgets.Widget = widgets.Dropdown(
                            options=dd_opts,
                            value=current_val if current_val in dd_opts else "",
                            layout=layout,
                        )
                    else:
                        w = widgets.Text(
                            value=current_val, layout=layout, continuous_update=False
                        )
                    w._coords = (r, c)  # type: ignore[attr-defined]
                    w.observe(self._on_cell_change, names="value")
                else:
                    w = widgets.Label(
                        value=current_val,
                        layout=widgets.Layout(
                            width=layout.width, min_width=layout.min_width,
                            flex="0 0 auto", border="1px solid #e0e0e0",
                            background_color="#fafafa",
                        ),
                    )
                    w._coords = (r, c)  # type: ignore[attr-defined]

                cell_pool.append(w)
                row_cells.append(w)

            data_row = widgets.HBox(row_cells, layout=row_layout)
            if not visible:
                data_row.layout.display = "none"
            # Pin data-row 0 only when freeze is on and we are on page 0
            if pool_idx == 0 and self.page == 0 and self.freeze_top:
                data_row.add_class("mini-excel-freeze-row1")

            self._pool_rows.append((data_row, rn_label, cell_pool))

        # ── Assemble grid ──────────────────────────────────────────────────
        all_rows: list[widgets.Widget] = [header_row, self._frozen_row_hbox]
        all_rows.extend(hbox for hbox, _, _ in self._pool_rows)
        self.grid_box.children = all_rows

        for col_idx in self.hidden_cols:
            self._apply_col_visibility(col_idx, True)

        self._update_page_indicator()

    # ── Core grid: fast value-only update (pagination) ────────────────────

    def _update_page_data(self) -> None:
        """
        Update widget values for the new page without creating any widgets.
        Only `.value`, `.options`, and `layout.display` are mutated.
        """
        row_start   = self.page * self.page_size
        row_end     = min(row_start + self.page_size, self.sheet.rows)
        num_visible = row_end - row_start

        self._programmatic_update = True
        try:
            # Show/hide frozen header copy
            if self._frozen_row_hbox is not None:
                show_frozen = self.freeze_top and self.page > 0 and self.sheet.rows > 0
                self._frozen_row_hbox.layout.display = "" if show_frozen else "none"
                if show_frozen:
                    for c, lbl in self._frozen_row_labels.items():
                        nd = self.sheet.get_display(0, c)
                        if lbl.value != nd:
                            lbl.value = nd

            # Manage freeze class on pool row 0
            if self._pool_rows:
                first_hbox = self._pool_rows[0][0]
                if self.page == 0 and self.freeze_top:
                    first_hbox.add_class("mini-excel-freeze-row1")
                else:
                    first_hbox.remove_class("mini-excel-freeze-row1")

            # Update each pool slot
            for pool_idx, (data_row, rn_label, cell_pool) in enumerate(self._pool_rows):
                r       = row_start + pool_idx
                visible = pool_idx < num_visible

                data_row.layout.display = "" if visible else "none"
                if not visible:
                    continue

                rn_label.value = str(r + 1)

                for c, w in enumerate(cell_pool):
                    w._coords = (r, c)  # type: ignore[attr-defined]
                    nd = self.sheet.get_display(r, c)

                    if isinstance(w, widgets.Label):
                        if w.value != nd:
                            w.value = nd
                    elif isinstance(w, widgets.Dropdown):
                        opts = list(w.options)
                        if nd not in opts:
                            opts.insert(1, nd)
                            w.options = opts
                        target = nd if nd in w.options else ""
                        if w.value != target:
                            w.value = target
                    else:  # Text
                        if w.value != nd:
                            w.value = nd
        finally:
            self._programmatic_update = False

        self._update_page_indicator()

    # ── Cell-change events ─────────────────────────────────────────────────

    def _on_cell_change(self, change: dict) -> None:
        if self._programmatic_update:
            return
        w: Any = change["owner"]
        r, c = w._coords
        self.selected = (r, c)
        self.cell_label.value = rc_to_a1(r, c)
        self.sheet.set_cell(r, c, change["new"])
        self._refresh_visible_cells()
        self._programmatic_update = True
        self.formula_bar.value = self.sheet.get_raw(r, c)
        self._programmatic_update = False
        self._set_status(
            f"Set {rc_to_a1(r, c)} = {self.sheet.get_raw(r, c)!r} → {self.sheet.get_display(r, c)}"
        )

    def _on_formula_bar_change(self, change: dict) -> None:
        if self._programmatic_update:
            return
        r, c = self.selected
        if c not in self.editable_cols:
            return
        self.sheet.set_cell(r, c, change["new"])
        self._refresh_visible_cells()

    def _refresh_visible_cells(self, *, skip_selected: bool = True) -> None:
        """Push fresh computed values into every visible pool widget."""
        row_start = self.page * self.page_size
        self._programmatic_update = True
        try:
            for c, lbl in self._frozen_row_labels.items():
                nd = self.sheet.get_display(0, c)
                if lbl.value != nd:
                    lbl.value = nd
            for pool_idx, (_, _, cell_pool) in enumerate(self._pool_rows):
                r = row_start + pool_idx
                if r >= self.sheet.rows:
                    break
                for c, w in enumerate(cell_pool):
                    nd = self.sheet.get_display(r, c)
                    if isinstance(w, widgets.Label):
                        if w.value != nd:
                            w.value = nd
                    elif isinstance(w, widgets.Dropdown):
                        opts = list(w.options)
                        if nd not in opts:
                            opts.insert(1, nd)
                            w.options = opts
                        target = nd if nd in w.options else ""
                        if w.value != target:
                            w.value = target
                    else:  # Text
                        if skip_selected and w._coords == self.selected:  # type: ignore[attr-defined]
                            continue
                        if w.value != nd:
                            w.value = nd
        finally:
            self._programmatic_update = False

    # ── Dropdown config ────────────────────────────────────────────────────

    def _on_toggle_dd_config(self, _: Any) -> None:
        cur = self.dd_config_row.layout.display
        self.dd_config_row.layout.display = "none" if cur != "none" else ""

    def _on_apply_dropdown(self, _: Any) -> None:
        c = self._parse_col_input(self.dd_col_input)
        if c is None: return
        items_str = self.dd_items_input.value.strip()
        if not items_str:
            self._set_status("Enter at least one item."); return
        options = [x.strip() for x in items_str.split(",") if x.strip()]
        if not options:
            self._set_status("No valid items found."); return
        self.col_dropdowns[c] = options
        self._render_grid()
        self._set_status(f"Dropdown applied to column {col_letter(c)}: {options}")

    def _on_remove_dropdown(self, _: Any) -> None:
        c = self._parse_col_input(self.dd_col_input)
        if c is None: return
        if c in self.col_dropdowns:
            del self.col_dropdowns[c]
            self._render_grid()
            self._set_status(f"Dropdown removed from column {col_letter(c)}.")
        else:
            self._set_status(f"Column {col_letter(c)} has no dropdown configured.")

    # ── Column width ───────────────────────────────────────────────────────

    def _on_toggle_cw_config(self, _: Any) -> None:
        cur = self.cw_config_row.layout.display
        self.cw_config_row.layout.display = "none" if cur != "none" else ""

    def _on_apply_col_width(self, _: Any) -> None:
        c = self._parse_col_input(self.cw_col_input)
        if c is None: return
        px = self.cw_width_input.value
        self.col_widths[c] = px
        self._apply_col_width(c, px)   # in-place, no rebuild
        self._set_status(f"Column {col_letter(c)} width set to {px}px.")

    def _on_reset_col_width(self, _: Any) -> None:
        c = self._parse_col_input(self.cw_col_input)
        if c is None: return
        if c in self.col_widths:
            del self.col_widths[c]
            self._apply_col_width(c, 90)
            self._set_status(f"Column {col_letter(c)} width reset to 90px.")
        else:
            self._set_status(f"Column {col_letter(c)} is already at default width.")

    # ── Hide / unhide column ───────────────────────────────────────────────

    def _on_toggle_hide_config(self, _: Any) -> None:
        cur = self.hc_config_row.layout.display
        self.hc_config_row.layout.display = "none" if cur != "none" else ""

    def _on_hide_col(self, _: Any) -> None:
        c = self._parse_col_input(self.hc_col_input)
        if c is None: return
        if c in self.hidden_cols:
            self._set_status(f"Column {col_letter(c)} is already hidden."); return
        self.hidden_cols.add(c)
        self._apply_col_visibility(c, True)
        self._set_status(f"Column {col_letter(c)} hidden. Use Unhide to restore it.")

    def _on_unhide_col(self, _: Any) -> None:
        c = self._parse_col_input(self.hc_col_input)
        if c is None: return
        if c not in self.hidden_cols:
            self._set_status(f"Column {col_letter(c)} is not hidden."); return
        self.hidden_cols.discard(c)
        self._apply_col_visibility(c, False)
        self._set_status(f"Column {col_letter(c)} restored.")

    # ── Toolbar button handlers ────────────────────────────────────────────

    def _on_load_data(self, _: Any) -> None:
        text = self.input_area.value.strip()
        if not text:
            self._set_status("Nothing to load — paste data into the Input data box first."); return
        delimiter = "\t" if "\t" in text else ","
        reader = csv.reader(io.StringIO(text), delimiter=delimiter)
        rows = [row for row in reader if any(cell.strip() for cell in row)]
        if not rows:
            self._set_status("Could not parse data."); return
        self.sheet.load_from_2d(rows)
        self.page = 0
        self.editable_cols.clear()
        self._render_grid()
        self._set_status(
            f"Loaded {len(rows)} row(s), {max(len(r) for r in rows)} col(s). "
            "Use Column Edit to enable editing."
        )

    def _on_undo(self, _: Any) -> None:
        prev_rows, prev_cols = self.sheet.rows, self.sheet.cols
        if self.sheet.undo():
            if self.sheet.rows != prev_rows or self.sheet.cols != prev_cols:
                self._render_grid()
            else:
                self._refresh_visible_cells(skip_selected=False)
            self._set_status("Undid last change.")
        else:
            self._set_status("Nothing to undo.")

    def _on_redo(self, _: Any) -> None:
        prev_rows, prev_cols = self.sheet.rows, self.sheet.cols
        if self.sheet.redo():
            if self.sheet.rows != prev_rows or self.sheet.cols != prev_cols:
                self._render_grid()
            else:
                self._refresh_visible_cells(skip_selected=False)
            self._set_status("Redid change.")
        else:
            self._set_status("Nothing to redo.")

    def _on_ins_row(self, _: Any) -> None:
        r, _ = self.selected
        self.sheet.insert_row(r)
        self._render_grid()
        self._set_status(f"Inserted row at {r + 1}.")

    def _on_del_row(self, _: Any) -> None:
        r, _ = self.selected
        self.sheet.delete_row(r)
        total = self._total_pages()
        if self.page >= total:
            self.page = max(0, total - 1)
        self._render_grid()
        self._set_status(f"Deleted row {r + 1}.")

    def _on_ins_col(self, _: Any) -> None:
        _, c = self.selected
        self.sheet.insert_col(c)
        self.editable_cols = {(x + 1 if x >= c else x) for x in self.editable_cols}
        self._render_grid()
        self._set_status(f"Inserted column at {col_letter(c)}.")

    def _on_del_col(self, _: Any) -> None:
        _, c = self.selected
        self.sheet.delete_col(c)
        self.editable_cols = {(x - 1 if x > c else x) for x in self.editable_cols if x != c}
        self._render_grid()
        self._set_status(f"Deleted column {col_letter(c)}.")

    def _on_save_csv(self, _: Any) -> None:
        path = (self.path_input.value or "miniexcel_output") + ".csv"
        self.sheet.save_csv(path)
        self._set_status(f"Saved → {path}")

    def _on_save_xlsx(self, _: Any) -> None:
        path = (self.path_input.value or "miniexcel_output") + ".xlsx"
        self.sheet.save_xlsx(path)
        self._set_status(f"Saved → {path}")

    def _on_load_csv(self, _: Any) -> None:
        raw = self.path_input.value.strip()
        path = raw if raw.lower().endswith(".csv") else (raw or "miniexcel_output") + ".csv"
        try:
            self.sheet.load_csv(path)
        except FileNotFoundError:
            self._set_status(f"File not found: {path}"); return
        except Exception as e:
            self._set_status(f"Error loading {path}: {e}"); return
        self.page = 0
        self.editable_cols.clear()
        self._render_grid()
        self._set_status(f"Loaded {self.sheet.rows} rows × {self.sheet.cols} cols from {path}. "
                         "Use Column Edit to enable editing.")

    def _on_load_xlsx(self, _: Any) -> None:
        raw = self.path_input.value.strip()
        path = raw if raw.lower().endswith(".xlsx") else (raw or "miniexcel_output") + ".xlsx"
        try:
            self.sheet.load_xlsx(path)
        except FileNotFoundError:
            self._set_status(f"File not found: {path}"); return
        except Exception as e:
            self._set_status(f"Error loading {path}: {e}"); return
        self.page = 0
        self.editable_cols.clear()
        self._render_grid()
        self._set_status(f"Loaded {self.sheet.rows} rows × {self.sheet.cols} cols from {path}. "
                         "Use Column Edit to enable editing.")

    def _on_clear(self, _: Any) -> None:
        self.sheet = Spreadsheet(self.sheet.rows, self.sheet.cols)
        self.page = 0
        self.editable_cols.clear()
        self._render_grid()
        self._set_status("Cleared.")

    # ── Public API ─────────────────────────────────────────────────────────

    def load_sample(self) -> None:
        """Populate the sheet with built-in demo data."""
        sample = [
            ["Product", "Q1",   "Q2",   "Q3",   "Q4",   "Total",          "Avg"],
            ["Widgets", "1200", "1450", "1600", "1800", "=SUM(B2:E2)",   "=AVERAGE(B2:E2)"],
            ["Gadgets", "800",  "950",  "1100", "1300", "=SUM(B3:E3)",   "=AVERAGE(B3:E3)"],
            ["Gizmos",  "500",  "600",  "750",  "900",  "=SUM(B4:E4)",   "=AVERAGE(B4:E4)"],
            ["Doodads", "300",  "350",  "400",  "450",  "=SUM(B5:E5)",   "=AVERAGE(B5:E5)"],
            ["Total",   "=SUM(B2:B5)", "=SUM(C2:C5)", "=SUM(D2:D5)", "=SUM(E2:E5)",
             "=SUM(F2:F5)", "=AVERAGE(G2:G5)"],
            [],
            ["Stats",  "Value"],
            ["Best quarter Q4 total",    "=E6"],
            ["Growth Q1→Q4 (Widgets)",   "=(E2-B2)/B2"],
            ["High performer?",          '=IF(F2>5000, "Yes", "No")'],
        ]
        self.sheet.load_from_2d(sample)
        self.page = 0
        self.editable_cols.clear()
        self._render_grid()
        self._set_status("Sample data loaded. Use Column Edit → Enable to make columns editable.")

    def show(self) -> None:
        """Render the spreadsheet UI in the notebook output cell."""
        display(self.container)

    # ── Internal helpers ───────────────────────────────────────────────────

    def _parse_col_input(self, text_widget: widgets.Text) -> int | None:
        """Parse a column letter from *text_widget*; emit a status message and return None on error."""
        col_str = text_widget.value.strip().upper()
        if not col_str:
            self._set_status("Enter a column letter (e.g. A)."); return None
        try:
            c = col_index(col_str)
        except Exception:
            self._set_status(f"Invalid column: {col_str!r}"); return None
        if c >= self.sheet.cols:
            self._set_status(f"Column {col_str} is out of range (sheet has {self.sheet.cols} columns).")
            return None
        return c

    def _set_status(self, msg: str) -> None:
        self.status.value = f'<span style="color:#333">{msg}</span>'
