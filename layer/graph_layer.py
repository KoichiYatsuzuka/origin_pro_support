"""
Graph layer classes for OriginExt wrappers.

This module contains wrapper classes for Origin graph layers including:
- DataPlot
- Axis
- GraphLayer
"""
from __future__ import annotations
import math
import re
import warnings
import OriginExt.OriginExt as oext_types
import OriginExt._OriginExt as oext
from enum import Enum
from typing import Optional, Union, TYPE_CHECKING, List
from collections.abc import Iterator

from ..base import OriginCommandResponceError, OriginNotFoundError, OriginObjectWrapper
from .enums import ColorMap, AxisType, XYPlotType, GroupMode, OriginColorIndex, ColorSpec, color_to_lt_str, LegendLayout, LegendAnchor, TickType, MarkerShape, LineStyle
from .worksheet import Worksheet

from ..base import APP
from ..lab_talk.lab_talk_commands import (
    axis_pg, axis_ps,
    layer_axis_get, layer_axis_set,
    layer_axis_get_from, layer_axis_set_from,
    layer_axis_get_to, layer_axis_set_to,
    layer_axis_get_ticks, layer_axis_set_ticks,
    layer_plot_get, layer_plot_set,
    layer_plot_get_name, layer_plot_count,
    plot_set_symbol_size, plot_set_symbol_kind,
    plot_set_line_style, plot_set_line_width,
    win_activate,
    legend_get_num, legend_set_num,
    legend_get_str, legend_set_str,
    legend_update, legend_reconstruct, legend_reset_position, legend_set_layout,
    active_layer_x_get_from, active_layer_x_get_to,
    active_layer_y_get_from, active_layer_y_get_to,
)

if TYPE_CHECKING:
    from ..pages import GraphPage


# Regex to extract the 1-based plot index from GetRange() strings like
# '[Graph1]1!Plot(3)' or '[Graph1]Layer1!Plot(3)'.
_RE_PLOT_ID = re.compile(r'Plot\((\d+)\)', re.IGNORECASE)


# ================== DataPlot Class ==================

class DataPlot:
    """
    Data plot in a graph layer.
    Wrapper class that wraps OriginExt.OriginExt.DataPlot.

    Holds a reference to the parent ``GraphLayer`` and the underlying
    ``oext_types.DataPlot`` object.  Index-based access is not supported;
    the 1-based LabTalk plot ID is resolved at runtime by comparing
    ``DatasetName`` values reported by LabTalk.

    Corresponds to: originpro.DataPlot, OriginExt.OriginExt.DataPlot
    """

    def __init__(self, plot: oext_types.DataPlot, graph_layer: 'GraphLayer',
                 plot_type: Optional[XYPlotType] = None):
        """
        Initialize DataPlot wrapper.

        Args:
            plot: ``OriginExt.OriginExt.DataPlot`` instance to wrap.
            graph_layer: Parent ``GraphLayer`` that owns this plot.
            plot_type: The ``XYPlotType`` used when creating this plot.
                       When provided, ``_has_line`` and ``_has_symbol`` use it
                       directly instead of calling ``GetTheme()``, which can
                       raise a server error after the plot list is rebuilt.
        """
        if plot is None:
            raise ValueError("DataPlot cannot be created with None plot object")
        if not isinstance(plot, oext_types.DataPlot):
            raise TypeError(
                f"plot must be an OriginExt.DataPlot instance, got {type(plot).__name__}"
            )
        self._plot: oext_types.DataPlot = plot
        self._graph_layer: 'GraphLayer' = graph_layer
        self._plot_type: Optional[XYPlotType] = plot_type

    # ── helpers ──────────────────────────────────────────────────────────

    @property
    def api_core(self) -> 'APP':
        """Get the API core reference from the parent layer."""
        return self._graph_layer.api_core

    def _page_context(self) -> tuple[str, int]:
        """Return ``(page_name, layer_id)`` for use in LabTalk command helpers.

        Raises:
            RuntimeError: If the parent GraphPage is unavailable.
        """
        parent_page = self._graph_layer._parent_page
        if parent_page is None:
            raise RuntimeError(
                "Cannot build LabTalk command: parent GraphPage is not set on the parent GraphLayer."
            )
        return parent_page.name, self._graph_layer._id

    def _resolve_plot_id(self) -> int:
        """Resolve the 1-based LabTalk plot ID for this DataPlot.

        Extracts the 1-based plot index from the range string returned by
        ``OriginExt.DataPlot.GetRange()``, which has the form
        ``"[PageName]LayerN!Plot(M)"``.

        Returns:
            int: 1-based plot index within the parent layer.

        Raises:
            OriginNotFoundError: If the range string cannot be parsed.
        """
        range_str: str = self._plot.GetRange()
        match = _RE_PLOT_ID.search(range_str)
        if match is None:
            raise OriginNotFoundError(
                f"Cannot extract plot ID from range string '{range_str}'."
            )
        return int(match.group(1))

    def _execute_with_active_page(self, command: str) -> None:
        """Execute a LabTalk command after activating this plot's page.

        Saves the currently active window name, activates the page that owns
        this plot, executes *command*, then restores the previously active
        window.  This ensures ``set`` commands (which require an active graph
        window) target the correct page without leaving side-effects on the
        user's active window.

        Args:
            command: LabTalk command string to execute.
        """
        page_name, _ = self._page_context()
        current_win = self.api_core.LT_get_str('%H')
        self.api_core.LT_execute(f"win -a {page_name}")
        self.api_core.LT_execute(command)
        if current_win:
            self.api_core.LT_execute(f"win -a {current_win}")

    # ── plot-type capability sets ─────────────────────────────────────────

    _LINE_TYPES: frozenset = frozenset({
        XYPlotType.LINE,
        XYPlotType.LINE_SYMBOL,
        XYPlotType.AREA,
    })
    _SYMBOL_TYPES: frozenset = frozenset({
        XYPlotType.SCATTER,
        XYPlotType.LINE_SYMBOL,
    })

    def _has_symbol(self) -> bool:
        """Return True if this plot type supports symbol (marker) display.

        Uses the stored ``_plot_type`` when available (fast, no Origin call).
        Falls back to inspecting the theme tree only when ``_plot_type`` is
        unknown, because ``GetTheme()`` can raise a server error after the
        plot list is rebuilt by a subsequent ``add_xy_plot`` call.
        """
        if self._plot_type is not None:
            return self._plot_type in self._SYMBOL_TYPES
        theme = self._plot.GetTheme()
        if theme is None:
            return False
        child = theme.firstChild
        while child is not None:
            if child.Name == 'Symbol':
                return True
            child = child.nextSibling
        return False

    def _assert_has_symbol(self) -> None:
        """Raise ValueError if this plot does not support symbol display."""
        if not self._has_symbol():
            raise ValueError(
                f"Plot '{self.name}' does not support symbol (marker) properties. "
                "Only scatter and line+symbol plots have markers."
            )

    def _has_line(self) -> bool:
        """Return True if this plot type supports line display.

        Uses the stored ``_plot_type`` when available (fast, no Origin call).
        Falls back to inspecting the theme tree only when ``_plot_type`` is
        unknown, because ``GetTheme()`` can raise a server error after the
        plot list is rebuilt by a subsequent ``add_xy_plot`` call.
        """
        if self._plot_type is not None:
            return self._plot_type in self._LINE_TYPES
        theme = self._plot.GetTheme()
        if theme is None:
            return False
        child = theme.firstChild
        while child is not None:
            if child.Name == 'Line':
                return True
            child = child.nextSibling
        return False

    def _assert_has_line(self) -> None:
        """Raise ValueError if this plot does not support line display."""
        if not self._has_line():
            raise ValueError(
                f"Plot '{self.name}' does not support line properties. "
                "Only line and line+symbol plots have lines."
            )

    # ── properties ───────────────────────────────────────────────────────

    @property
    def name(self) -> str:
        """Dataset name of the data plot."""
        return self._plot.GetDatasetName()

    @property
    def worksheet(self):
        """Worksheet containing the plotted data."""
        worksheet_obj = oext.DataPlot_GetWorksheet(self._plot)
        return Worksheet(worksheet_obj, self.api_core)

    @property
    def color(self) -> int:
        """Line/symbol color as an Origin color value.

        Reads via ``DataPlot.GetNumProp("color")``.
        For index colors (1-24) the returned integer matches
        ``OriginColorIndex``.  For custom RGB colors the returned value is
        the Windows COLORREF integer (R + G*256 + B*65536).

        Setting accepts an int, (R,G,B) tuple, or ``OriginColorIndex``.
        """
        raw = self._plot.GetNumProp("color")
        if math.isnan(raw):
            page_name, layer_id = self._page_context()
            plot_id = self._resolve_plot_id()
            raise OriginCommandResponceError(f"NaN value is received while getting plot color. (page={page_name}, layer={layer_id}, plot={plot_id})")
        return int(raw)

    @color.setter
    def color(self, value: ColorSpec) -> None:
        """Set the plot color.

        Uses ``DataPlot.SetNumProp("color", color_int)`` for index colors,
        or ``DataPlot.SetStrProp("color", color_str)`` for RGB colors.
        """
        if isinstance(value, tuple) and len(value) == 3:
            # RGB tuple: must use LabTalk layer_plot_set since SetNumProp can't encode RGB
            page_name, layer_id = self._page_context()
            plot_id = self._resolve_plot_id()
            color_str = color_to_lt_str(value)
            self.api_core.LT_execute(layer_plot_set(page_name, layer_id, plot_id, "color", color_str))
        else:
            if isinstance(value, OriginColorIndex):
                int_val = value.value
            else:
                int_val = int(value)
            self._plot.SetNumProp("color", int_val)

    @property
    def color_map(self) -> ColorMap:
        """Color map used by the data plot."""
        color_map_str = oext.DataPlot_GetColorMap(self._plot)
        return ColorMap(color_map_str)

    @color_map.setter
    def color_map(self, value: ColorMap) -> None:
        """Set color map for the data plot."""
        try:
            oext.DataPlot_SetColorMap(self._plot, value.value)
        except Exception as e:
            raise RuntimeError(f"Failed to set color map: {e}") from e

    @property
    def symbol_size(self) -> float:
        """Symbol (marker) size for this plot in **points** (1 pt = 1/72 inch).

        Reads via ``[Page]!layerN.plot(M).symbol.size`` LabTalk notation.
        Writes via ``set <dataset> -z <size>`` (unit: points).

        Raises:
            ValueError: If the plot type does not support symbols.
            OriginCommandResponceError: If LabTalk returns NaN.
        """
        self._assert_has_symbol()
        page_name, layer_id = self._page_context()
        plot_id = self._resolve_plot_id()
        self.api_core.LT_execute(
            layer_plot_get(page_name, layer_id, plot_id, "symbol.size", "_dp_symsize")
        )
        raw = self.api_core.LT_get_var("_dp_symsize")
        if math.isnan(raw):
            raise OriginCommandResponceError(
                f"NaN returned while getting symbol.size for plot '{self.name}'."
            )
        return float(raw)

    @symbol_size.setter
    def symbol_size(self, value: float) -> None:
        """Set the symbol (marker) size.

        Uses ``set <dataset> -z <size>`` LabTalk command.

        Args:
            value: Symbol size in **points** (1 pt = 1/72 inch), must be positive.

        Raises:
            ValueError: If the plot type does not support symbols, or value is not positive.
        """
        self._assert_has_symbol()
        if value <= 0:
            raise ValueError(f"symbol_size must be positive, got {value}")
        self._execute_with_active_page(plot_set_symbol_size(self._plot.GetDatasetName(), value))

    @property
    def symbol_kind(self) -> MarkerShape:
        """Symbol (marker) shape for this plot.

        Reads and writes ``[Page]!layerN.plot(M).symbol.kind`` via LabTalk.

        Raises:
            ValueError: If the plot type does not support symbols.
            OriginCommandResponceError: If LabTalk returns NaN.
        """
        self._assert_has_symbol()
        page_name, layer_id = self._page_context()
        plot_id = self._resolve_plot_id()
        self.api_core.LT_execute(
            layer_plot_get(page_name, layer_id, plot_id, "symbol.kind", "_dp_symkind")
        )
        raw = self.api_core.LT_get_var("_dp_symkind")
        if math.isnan(raw):
            raise OriginCommandResponceError(
                f"NaN returned while getting symbol.kind for plot '{self.name}'."
            )
        return MarkerShape(int(raw))

    @symbol_kind.setter
    def symbol_kind(self, value: MarkerShape) -> None:
        """Set the symbol (marker) shape.

        Uses ``set <dataset> -k <kind>`` LabTalk command.

        Args:
            value: A ``MarkerShape`` enum member.

        Raises:
            ValueError: If the plot type does not support symbols.
        """
        self._assert_has_symbol()
        self._execute_with_active_page(plot_set_symbol_kind(self._plot.GetDatasetName(), value.value))

    @property
    def line_style(self) -> LineStyle:
        """Line style for this plot.

        Reads via ``DataPlot.GetNumProp("line.type")`` which returns an integer
        matching the ``LineStyle`` enum values (1 = SOLID, 2 = DASH, 3 = DOT, …).
        Writes via ``DataPlot.SetNumProp("line.type", value)``.

        Raises:
            ValueError: If the plot type does not support lines.
            OriginCommandResponceError: If the returned value is NaN.
        """
        self._assert_has_line()
        raw = self._plot.GetNumProp("line.type")
        if math.isnan(raw):
            raise OriginCommandResponceError(
                f"NaN returned while getting line.type for plot '{self.name}'."
            )
        return LineStyle(int(raw))

    @line_style.setter
    def line_style(self, value: LineStyle) -> None:
        """Set the line style.

        Uses ``DataPlot.SetNumProp("line.type", value.value)``.
        The stored integer matches the ``LineStyle`` enum
        (1 = SOLID, 2 = DASH, 3 = DOT, 4 = DASH_DOT, 5 = DASH_DOT_DOT).

        Args:
            value: A ``LineStyle`` enum member.

        Raises:
            ValueError: If the plot type does not support lines.
        """
        self._assert_has_line()
        self._plot.SetNumProp("line.type", value.value)

    @property
    def line_width(self) -> float:
        """Line width for this plot in **points** (1 pt = 1/72 inch).

        Reads via ``DataPlot.GetNumProp("line.width")`` which returns the value
        directly in points.
        Note: the LabTalk ``set -w`` command uses units of pts × 500
        (e.g. 500 = 1 pt), but ``GetNumProp``/``SetNumProp`` use plain points.

        Raises:
            ValueError: If the plot type does not support lines.
            OriginCommandResponceError: If the returned value is NaN.
        """
        self._assert_has_line()
        raw = self._plot.GetNumProp("line.width")
        if math.isnan(raw):
            raise OriginCommandResponceError(
                f"NaN returned while getting line.width for plot '{self.name}'."
            )
        return float(raw)

    @line_width.setter
    def line_width(self, value: float) -> None:
        """Set the line width.

        Uses ``DataPlot.SetNumProp("line.width", value)``.

        Args:
            value: Line width in **points** (1 pt = 1/72 inch), must be positive.

        Raises:
            ValueError: If the plot type does not support lines, or value is not positive.
        """
        self._assert_has_line()
        if value <= 0:
            raise ValueError(f"line_width must be positive, got {value}")
        self._plot.SetNumProp("line.width", value)

    def change_data(self, data_obj, designation: str = 'Y', keep_modifiers: bool = False) -> bool:
        """
        Change data for the data plot.

        Args:
            data_obj: Data object to use
            designation: Data designation (X, Y, Z, etc.)
            keep_modifiers: Whether to keep existing modifiers

        Returns:
            bool: True if successful
        """
        return self._plot.ChangeData(data_obj, designation, keep_modifiers)


# ================== Legend Class ==================

class Legend:
    """
    Legend object for a graph layer.

    Wraps the LabTalk ``legend`` graphic object that Origin places on each
    graph layer.  All properties and methods delegate to LabTalk via
    ``api_core.LT_execute`` / ``LT_get_var`` / ``LT_get_str``.

    Activate the parent layer before every operation so that LabTalk's
    ``legend`` object always points to the correct layer.

    Ref: https://www.originlab.com/doc/LabTalk/ref/Graphic-objs
         https://www.originlab.com/doc/LabTalk/ref/Legend-cmd
    """

    def __init__(self, graph_layer: 'GraphLayer') -> None:
        """
        Initialize Legend wrapper.

        Args:
            graph_layer: Parent GraphLayer that owns this legend.
        """
        self._layer = graph_layer

    # ── helpers ──────────────────────────────────────────────────────────

    def _activate(self) -> None:
        """Activate the parent page so LabTalk's ``legend`` object is correct."""
        if self._layer._parent_page is not None:
            page_name = self._layer._parent_page.name
            self._layer.api_core.LT_execute(win_activate(page_name))

    # ── visibility ───────────────────────────────────────────────────────

    @property
    def visible(self) -> bool:
        """Whether the legend is visible.

        Corresponds to: ``legend.show`` in LabTalk.
        """
        self._activate()
        self._layer.api_core.LT_execute(legend_get_num("show", "_leg_show"))
        val = self._layer.api_core.LT_get_var("_leg_show")
        if math.isnan(val):
            return True
        return bool(int(val))

    @visible.setter
    def visible(self, value: bool) -> None:
        """Show or hide the legend.

        Args:
            value: True to show, False to hide.
        """
        self._activate()
        self._layer.api_core.LT_execute(legend_set_num("show", 1 if value else 0))

    # ── text ─────────────────────────────────────────────────────────────

    @property
    def text(self) -> str:
        """Raw text content of the legend object.

        Line breaks are stored internally as ``\\r\\n`` (CRLF) by Origin.
        The getter returns them as ``\\n`` (LF) for consistency.

        Corresponds to: ``legend.text$`` in LabTalk.
        """
        self._activate()
        self._layer.api_core.LT_execute(legend_get_str("text", "_leg_text"))
        raw = self._layer.api_core.LT_get_str("_leg_text")
        return raw.replace("\r\n", "\n").replace("\r", "\n")

    @text.setter
    def text(self, value: str) -> None:
        """Set the raw text of the legend.

        Line breaks must be Python ``\\n`` characters (or ``\\r\\n``).
        They are converted to ``\\r\\n`` (CRLF) before passing to Origin,
        which is the only newline form that Origin's legend object recognises.

        Do **not** use the literal string ``\\\\n`` — Origin does not
        interpret that as a line break in legend text.

        Example::

            legend.text = r"\\l(1) 0 min" + "\\n" + r"\\l(2) 50 min"

        Args:
            value: New legend text.  Use Python ``\\n`` for line breaks.
        """
        self._activate()
        normalised = value.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
        escaped = normalised.replace('"', '\\"')
        self._layer.api_core.LT_execute(legend_set_str("text", escaped))

    # ── position ─────────────────────────────────────────────────────────

    def get_position(self) -> tuple:
        """Get the legend position in axis-scale units.

        Corresponds to: ``legend.x``, ``legend.y`` in LabTalk.

        Returns:
            tuple: (x, y) centre coordinates of the legend in axis-scale units.
        """
        self._activate()
        self._layer.api_core.LT_execute(legend_get_num("x", "_leg_x"))
        self._layer.api_core.LT_execute(legend_get_num("y", "_leg_y"))
        x = self._layer.api_core.LT_get_var("_leg_x")
        y = self._layer.api_core.LT_get_var("_leg_y")
        x = 0.0 if math.isnan(x) else float(x)
        y = 0.0 if math.isnan(y) else float(y)
        return (x, y)

    def set_position(self, x: float, y: float) -> None:
        """Move the legend to the specified position.

        Coordinates are in axis-scale units (same coordinate system as the
        layer axes).

        Corresponds to: ``legend.x = x; legend.y = y`` in LabTalk.

        Args:
            x: Horizontal centre of the legend in axis-scale units.
            y: Vertical centre of the legend in axis-scale units.
        """
        self._activate()
        self._layer.api_core.LT_execute(legend_set_num("x", x))
        self._layer.api_core.LT_execute(legend_set_num("y", y))

    def get_position_pct(self) -> tuple:
        """Get the legend centre position as a percentage of the layer axis range.

        (0, 0) corresponds to the bottom-left corner of the axis frame;
        (100, 100) corresponds to the top-right corner.

        Reads ``layer.x.from/to`` and ``layer.y.from/to`` via LabTalk and
        converts the current ``legend.x``/``legend.y`` axis-scale values to
        percentages.

        Returns:
            tuple: (x_pct, y_pct) as percentages of the layer axis width/height.
        """
        self._activate()
        api = self._layer.api_core
        api.LT_execute(active_layer_x_get_from("_lxf"))
        api.LT_execute(active_layer_x_get_to("_lxt"))
        api.LT_execute(active_layer_y_get_from("_lyf"))
        api.LT_execute(active_layer_y_get_to("_lyt"))
        api.LT_execute(legend_get_num("x", "_leg_x"))
        api.LT_execute(legend_get_num("y", "_leg_y"))
        xf = api.LT_get_var("_lxf"); xt = api.LT_get_var("_lxt")
        yf = api.LT_get_var("_lyf"); yt = api.LT_get_var("_lyt")
        lx = api.LT_get_var("_leg_x"); ly = api.LT_get_var("_leg_y")
        xf = 0.0 if math.isnan(xf) else float(xf)
        xt = 1.0 if math.isnan(xt) else float(xt)
        yf = 0.0 if math.isnan(yf) else float(yf)
        yt = 1.0 if math.isnan(yt) else float(yt)
        lx = 0.0 if math.isnan(lx) else float(lx)
        ly = 0.0 if math.isnan(ly) else float(ly)
        x_pct = (lx - xf) / (xt - xf) * 100.0 if xt != xf else 0.0
        y_pct = (ly - yf) / (yt - yf) * 100.0 if yt != yf else 0.0
        return (x_pct, y_pct)

    def set_position_pct(
        self, x_pct: float, y_pct: float, anchor: LegendAnchor = LegendAnchor.CENTER
    ) -> None:
        """Move the legend to a position specified as a percentage of the layer axis range.

        (0, 0) corresponds to the bottom-left corner of the axis frame;
        (100, 100) corresponds to the top-right corner.

        Internally reads ``layer.x.from/to`` and ``layer.y.from/to`` via
        LabTalk, converts the percentage values to axis-scale units, and
        applies an offset determined by *anchor* so that the chosen corner of
        the legend lands on the target point.

        ``legend.x``/``legend.y`` in LabTalk refer to the **centre** of the
        legend object; ``legend.dx``/``legend.dy`` give its full width/height
        in axis-scale units, so the offset is ±dx/2 and ±dy/2.

        Args:
            x_pct:  Horizontal position as % of axis width  (0–100).
            y_pct:  Vertical   position as % of axis height (0–100).
            anchor: Which point of the legend is placed at (x_pct, y_pct).
                    Use ``LegendAnchor`` enum values:

                    - ``LegendAnchor.CENTER``       – centre of the legend *(default)*
                    - ``LegendAnchor.TOP_LEFT``     – top-left  corner
                    - ``LegendAnchor.TOP_RIGHT``    – top-right corner
                    - ``LegendAnchor.BOTTOM_LEFT``  – bottom-left  corner
                    - ``LegendAnchor.BOTTOM_RIGHT`` – bottom-right corner
        """
        self._activate()
        api = self._layer.api_core
        api.LT_execute(active_layer_x_get_from("_lxf"))
        api.LT_execute(active_layer_x_get_to("_lxt"))
        api.LT_execute(active_layer_y_get_from("_lyf"))
        api.LT_execute(active_layer_y_get_to("_lyt"))
        api.LT_execute(legend_get_num("dx", "_ldx"))
        api.LT_execute(legend_get_num("dy", "_ldy"))
        xf = api.LT_get_var("_lxf"); xt = api.LT_get_var("_lxt")
        yf = api.LT_get_var("_lyf"); yt = api.LT_get_var("_lyt")
        dx = api.LT_get_var("_ldx"); dy = api.LT_get_var("_ldy")
        if math.isnan(xf) or math.isnan(xt) or math.isnan(yf) or math.isnan(yt):
            raise RuntimeError("Could not read layer axis range from LabTalk.")
        dx = 0.0 if math.isnan(dx) else float(dx)
        dy = 0.0 if math.isnan(dy) else float(dy)
        sx, sy = anchor.value
        x = float(xf) + (float(xt) - float(xf)) * x_pct / 100.0 + sx * dx
        y = float(yf) + (float(yt) - float(yf)) * y_pct / 100.0 + sy * dy
        api.LT_execute(legend_set_num("x", x))
        api.LT_execute(legend_set_num("y", y))

    def reset_position(self) -> None:
        """Reset the legend to its default position.

        Corresponds to: ``legend -d`` in LabTalk.
        """
        self._activate()
        self._layer.api_core.LT_execute(legend_reset_position())

    # ── font size ─────────────────────────────────────────────────────────

    @property
    def font_size(self) -> int:
        """Font size of the legend text in points.

        Corresponds to: ``legend.fsize`` in LabTalk.
        """
        self._activate()
        self._layer.api_core.LT_execute(legend_get_num("fsize", "_leg_fsize"))
        val = self._layer.api_core.LT_get_var("_leg_fsize")
        if math.isnan(val):
            return 12
        return int(val)

    @font_size.setter
    def font_size(self, value: int) -> None:
        """Set the legend font size.

        Args:
            value: Font size in points (positive integer).
        """
        if value <= 0:
            raise ValueError(f"font_size must be a positive integer, got {value}")
        self._activate()
        self._layer.api_core.LT_execute(legend_set_num("fsize", value))

    # ── background box ────────────────────────────────────────────────────

    @property
    def background(self) -> int:
        """Background box style of the legend.

        0 = none, 1 = black border, 2 = shadow, 3 = white-out,
        4 = black border + white-out.

        Corresponds to: ``legend.background`` in LabTalk.
        """
        self._activate()
        self._layer.api_core.LT_execute(legend_get_num("background", "_leg_bg"))
        val = self._layer.api_core.LT_get_var("_leg_bg")
        if math.isnan(val):
            return 1
        return int(val)

    @background.setter
    def background(self, value: int) -> None:
        """Set the legend background box style.

        Args:
            value: 0=none, 1=black border, 2=shadow, 3=white-out,
                   4=black border + white-out.
        """
        self._activate()
        self._layer.api_core.LT_execute(legend_set_num("background", value))

    # ── layout ────────────────────────────────────────────────────────────

    def set_layout(self, layout: LegendLayout) -> None:
        """Rearrange legend entries horizontally or vertically.

        Args:
            layout: LegendLayout.VERTICAL or LegendLayout.HORIZONTAL.
        """
        self._activate()
        self._layer.api_core.LT_execute(legend_set_layout(layout.value))

    # ── reconstruct ───────────────────────────────────────────────────────

    def reconstruct(self) -> None:
        """Reconstruct the legend to best-fit the current plots.

        Equivalent to menu Graph > Legend > Reconstruct Legend.

        Corresponds to: ``legend -r`` in LabTalk.
        """
        self._activate()
        self._layer.api_core.LT_execute(legend_reconstruct())

    def update(self) -> None:
        """Create or update the legend on the active graph layer.

        Corresponds to: ``legend`` (no option) in LabTalk.
        """
        self._activate()
        self._layer.api_core.LT_execute(legend_update())


# ================== Axis Class ==================


class Axis:
    """
    Axis object for manipulating graph layer axes.
    Provides methods to get and modify axis ranges and properties.
    """

    def __init__(self, graph_layer: 'GraphLayer', axis_type: AxisType):
        """
        Initialize Axis wrapper.

        Args:
            graph_layer: Parent GraphLayer object
            axis_type: Type of axis (X, Y, Z, ERROR, X2, or Y2)
        """
        self._axis_type = axis_type
        self._graph_layer: 'GraphLayer' = graph_layer
        self._api_core: APP = graph_layer.api_core

    # ── parent references ────────────────────────────────────────────────

    @property
    def __graph_layer(self) -> 'GraphLayer':
        """Return the parent GraphLayer instance."""
        return self._graph_layer

    @property
    def __graph_page(self) -> 'GraphPage':
        """Return the parent GraphPage instance, or None if unavailable."""
        return self._graph_layer._parent_page

    # ── helpers ──────────────────────────────────────────────────────────

    def _page_context(self) -> tuple[str, int]:
        """Return ``(page_name, layer_id)`` for use in LabTalk command helpers.

        Raises:
            RuntimeError: If the parent GraphPage is unavailable.
        """
        parent_page = self._graph_layer._parent_page
        if parent_page is None:
            raise RuntimeError(
                "Cannot build LabTalk command: parent GraphPage is not set on this Axis."
            )
        return parent_page.name, self._graph_layer._id

    def _get_axis_letter(self) -> str:
        """Get the LabTalk axis object prefix for this axis type."""
        _map = {
            AxisType.X: "x",
            AxisType.Y: "y",
            AxisType.Z: "z",
            AxisType.ERROR: "e",
            AxisType.X2: "x2",
            AxisType.Y2: "y2",
        }
        return _map.get(self._axis_type, "x")

    # ── axis type ────────────────────────────────────────────────────────

    @property
    def axis_type(self) -> AxisType:
        """Get axis type"""
        return self._axis_type

    # ── range ────────────────────────────────────────────────────────────

    def get_range(self) -> tuple[float, float]:
        """
        Get the axis range (min, max).

        Returns:
            tuple: (min_value, max_value)
        """
        ax = self._get_axis_letter()
        page, lid = self._page_context()
        self._api_core.LT_execute(layer_axis_get_from(page, lid, ax, "_ax_from"))
        self._api_core.LT_execute(layer_axis_get_to(page, lid, ax, "_ax_to"))
        min_val = self._api_core.LT_get_var("_ax_from")
        max_val = self._api_core.LT_get_var("_ax_to")
        if math.isnan(min_val) or math.isnan(max_val):
            raise OriginCommandResponceError(f"NaN value is received while getting axis range. (page={page}, layer={lid})")
        return (float(min_val), float(max_val))

    def set_range(self, min_val: float, max_val: float) -> None:
        """
        Set the axis range.

        Args:
            min_val: Minimum value
            max_val: Maximum value
        """
        ax = self._get_axis_letter()
        page, lid = self._page_context()
        self._api_core.LT_execute(layer_axis_set(page, lid, ax, "rescale", 1))
        self._api_core.LT_execute(layer_axis_set_from(page, lid, ax, min_val))
        self._api_core.LT_execute(layer_axis_set_to(page, lid, ax, max_val))

    # ── reverse ──────────────────────────────────────────────────────────

    def get_reverse(self) -> bool:
        """
        Get whether the axis is reversed.

        Returns:
            bool: True if reversed
        """
        ax = self._get_axis_letter()
        page, lid = self._page_context()
        self._api_core.LT_execute(layer_axis_get(page, lid, ax, "reverse", "_ax_rev"))
        val = self._api_core.LT_get_var("_ax_rev")
        if math.isnan(val):
            raise OriginCommandResponceError(f"NaN value is received while getting axis reverse. (page={page}, layer={lid})")
        return bool(int(val))

    def set_reverse(self, reverse: bool) -> None:
        """
        Set whether the axis is reversed.

        Args:
            reverse: True to reverse the axis
        """
        ax = self._get_axis_letter()
        page, lid = self._page_context()
        self._api_core.LT_execute(layer_axis_set(page, lid, ax, "reverse", 1 if reverse else 0))

    # ── tick marks (TODO-2) ───────────────────────────────────────────────

    def _get_ticks_raw(self, ax: str) -> int:
        """Read the raw ``layer.axis.ticks`` bitmask value."""
        page, lid = self._page_context()
        self._api_core.LT_execute(layer_axis_get_ticks(page, lid, ax, "_ax_ticks"))
        val = self._api_core.LT_get_var("_ax_ticks")
        if math.isnan(val):
            raise OriginCommandResponceError(f"NaN value is received while getting axis ticks. (page={page}, layer={lid}, graph={self._graph_layer._name})")
        return int(val)

    def _set_ticks_raw(self, ax: str, raw: int) -> None:
        """Write the raw ``layer.axis.ticks`` bitmask value."""
        page, lid = self._page_context()
        self._api_core.LT_execute(layer_axis_set_ticks(page, lid, ax, raw))

    def get_major_tick(self) -> TickType:
        """
        Get the major tick style for this axis.

        Reads ``layer.axis.ticks`` and extracts bits 0-1 (major in/out).
        Ref: https://www.originlab.com/doc/LabTalk/ref/Layer-Axis-obj

        Returns:
            TickType: Current major tick style.
        """
        ax = self._get_axis_letter()
        raw = self._get_ticks_raw(ax)
        return TickType(raw & 0x03)

    def set_major_tick(self, tick_type: TickType) -> None:
        """
        Set the major tick style for this axis.

        Modifies bits 0-1 of ``layer.axis.ticks`` while preserving the
        minor tick bits (2-3).
        Ref: https://www.originlab.com/doc/LabTalk/ref/Layer-Axis-obj

        Args:
            tick_type: TickType enum value.
        """
        ax = self._get_axis_letter()
        raw = self._get_ticks_raw(ax)
        new_raw = (raw & ~0x03) | (tick_type.value & 0x03)
        self._set_ticks_raw(ax, new_raw)

    def get_minor_tick(self) -> TickType:
        """
        Get the minor tick style for this axis.

        Reads ``layer.axis.ticks`` and extracts bits 2-3 (minor in/out),
        shifting them down to match TickType values.
        Ref: https://www.originlab.com/doc/LabTalk/ref/Layer-Axis-obj

        Returns:
            TickType: Current minor tick style.
        """
        ax = self._get_axis_letter()
        raw = self._get_ticks_raw(ax)
        return TickType((raw >> 2) & 0x03)

    def set_minor_tick(self, tick_type: TickType) -> None:
        """
        Set the minor tick style for this axis.

        Modifies bits 2-3 of ``layer.axis.ticks`` while preserving the
        major tick bits (0-1).
        Ref: https://www.originlab.com/doc/LabTalk/ref/Layer-Axis-obj

        Args:
            tick_type: TickType enum value.
        """
        ax = self._get_axis_letter()
        raw = self._get_ticks_raw(ax)
        new_raw = (raw & ~0x0C) | ((tick_type.value & 0x03) << 2)
        self._set_ticks_raw(ax, new_raw)

    # ── tick spacing / count ─────────────────────────────────────────────

    def get_major_tick_spacing(self) -> float:
        """
        Get the major tick increment for this axis.

        Corresponds to: ``layer.axis.inc`` in LabTalk.
        Ref: https://www.originlab.com/doc/LabTalk/ref/Layer-Axis-obj

        Returns:
            float: Current major tick increment value.
        """
        ax = self._get_axis_letter()
        page, lid = self._page_context()
        self._api_core.LT_execute(layer_axis_get(page, lid, ax, "inc", "_ax_inc"))
        val = self._api_core.LT_get_var("_ax_inc")
        if math.isnan(val):
            raise OriginCommandResponceError(f"NaN value is received while getting major tick spacing. (page={page}, layer={lid}, graph={self._graph_layer._name})")
        return float(val)

    def set_major_tick_spacing(self, spacing: Union[int, float]) -> None:
        """
        Set the major tick increment for this axis.

        Corresponds to: ``layer.axis.inc = value`` in LabTalk.

        Args:
            spacing: Major tick increment (positive int or float).
        """
        if spacing <= 0:
            raise ValueError(f"spacing must be positive, got {spacing}")
        ax = self._get_axis_letter()
        page, lid = self._page_context()
        self._api_core.LT_execute(layer_axis_set(page, lid, ax, "inc", spacing))

    def get_minor_tick_count(self) -> int:
        """
        Get the number of minor ticks between major ticks for this axis.

        Uses ``axis -pg axis M varName`` (type ``M`` = number of Minor ticks).
        Requires an active graph window; activates the parent page automatically.

        Returns:
            int: Current number of minor ticks between major ticks.
        """
        pg_ax = self._get_axis_pg_letter()
        page, lid = self._page_context()
        self._api_core.LT_execute(f"win -a {page}")
        self._api_core.LT_execute(axis_pg(pg_ax, "M", "_ax_minor"))
        val = self._api_core.LT_get_var("_ax_minor")
        if math.isnan(val):
            raise OriginCommandResponceError(f"NaN value is received while getting minor tick count. (page={page}, layer={lid})")
        return int(val)

    def set_minor_tick_count(self, count: int) -> None:
        """
        Set the number of minor ticks between major ticks for this axis.

        Uses ``axis -ps axis M value`` (type ``M`` = number of Minor ticks).
        Requires an active graph window; activates the parent page automatically.

        Args:
            count: Number of minor ticks between major ticks (non-negative int).
        """
        if count < 0:
            raise ValueError(f"count must be non-negative, got {count}")
        pg_ax = self._get_axis_pg_letter()
        page, lid = self._page_context()
        self._api_core.LT_execute(f"win -a {page}")
        self._api_core.LT_execute(axis_ps(pg_ax, "M", count))

    # ── opposite axis visibility (TODO-3) ─────────────────────────────────

    def show_opposite_axis(self) -> None:
        """
        Show the opposite (top/right) axis line for this axis.

        For X-axis: shows the top axis (``x2.showAxes``).
        For Y-axis: shows the right axis (``y2.showAxes``).
        Only applicable for X and Y axes; raises ValueError otherwise.

        Corresponds to: ``x2.showAxes = 1`` / ``y2.showAxes = 1`` in LabTalk.
        Ref: https://www.originlab.com/doc/LabTalk/ref/Layer-Axis-obj
        """
        ax2 = self._get_opposite_axis_letter()
        page, lid = self._page_context()
        self._api_core.LT_execute(layer_axis_set(page, lid, ax2, "showAxes", 1))

    def hide_opposite_axis(self) -> None:
        """
        Hide the opposite (top/right) axis line for this axis.

        For X-axis: hides the top axis (``x2.showAxes``).
        For Y-axis: hides the right axis (``y2.showAxes``).
        Only applicable for X and Y axes; raises ValueError otherwise.

        Corresponds to: ``x2.showAxes = 0`` / ``y2.showAxes = 0`` in LabTalk.
        """
        ax2 = self._get_opposite_axis_letter()
        page, lid = self._page_context()
        self._api_core.LT_execute(layer_axis_set(page, lid, ax2, "showAxes", 0))

    def get_opposite_axis_visible(self) -> bool:
        """
        Get whether the opposite (top/right) axis is visible.

        Returns:
            bool: True if visible.
        """
        ax2 = self._get_opposite_axis_letter()
        page, lid = self._page_context()
        self._api_core.LT_execute(layer_axis_get(page, lid, ax2, "showAxes", "_ax_show"))
        val = self._api_core.LT_get_var("_ax_show")
        if math.isnan(val):
            raise OriginCommandResponceError(f"NaN value is received while getting opposite axis visibility. (page={page}, layer={lid})")
        return bool(int(val))

    def _get_axis_pg_letter(self) -> str:
        """Return the axis identifier for ``axis -pg/-ps`` commands (X or Y only).

        Raises:
            ValueError: If this axis type is not supported by axis -pg/-ps.
        """
        _map = {
            AxisType.X:  "X",
            AxisType.Y:  "Y",
            AxisType.X2: "X",
            AxisType.Y2: "Y",
        }
        letter = _map.get(self._axis_type)
        if letter is None:
            raise ValueError(
                f"Axis type {self._axis_type.name} is not supported by axis -pg/-ps. "
                "Only X, Y, X2, and Y2 axes are supported."
            )
        return letter

    def _get_opposite_axis_letter(self) -> str:
        """Return the LabTalk object name for the opposite axis (x2 or y2).

        Raises:
            ValueError: If this axis type has no opposite axis.
        """
        if self._axis_type == AxisType.X:
            return "x2"
        if self._axis_type == AxisType.Y:
            return "y2"
        raise ValueError(
            f"Axis type {self._axis_type.name} does not have an opposite axis. "
            "Only X and Y axes support show/hide of the opposite axis."
        )

    def _get_title_obj_name(self) -> str:
        """Return the LabTalk graphic object name for the axis title.

        Origin uses pre-defined named text objects for axis titles:
          xb = bottom X axis title
          xl = left Y axis title
          xt = top X2 axis title
          yr = right Y2 axis title

        Raises:
            ValueError: If this axis type has no dedicated title object.
        """
        _map = {
            AxisType.X:  "xb",
            AxisType.Y:  "yl",
            AxisType.X2: "xt",
            AxisType.Y2: "yr",
        }
        obj_name = _map.get(self._axis_type)
        if obj_name is None:
            raise ValueError(
                f"Axis type {self._axis_type.name} does not have a title object. "
                "Only X, Y, X2, and Y2 axes support label_text."
            )
        return obj_name

    # ── axis label text (TODO-4) ──────────────────────────────────────────

    @property
    def label_text(self) -> str:
        """
        Get the axis title text.

        Corresponds to: ``xb.text$`` / ``yl.text$`` etc. in LabTalk.
        Ref: https://www.originlab.com/doc/LabTalk/guide/Creating-and-Accessing-Graphical-objs

        Returns:
            str: Current axis title string.
        """
        obj = self._get_title_obj_name()
        page, lid = self._page_context()
        self._api_core.LT_execute(f"win -a {page}")
        self._api_core.LT_execute(f"string _ax_lbl$; _ax_lbl$ = {obj}.text$")
        return self._api_core.LT_get_str("_ax_lbl")

    @label_text.setter
    def label_text(self, value: str) -> None:
        """
        Set the axis title text.

        Corresponds to: ``xb.text$ = "..."`` / ``yl.text$ = "..."`` etc. in LabTalk.

        Args:
            value: New title string for the axis.
        """
        obj = self._get_title_obj_name()
        page, lid = self._page_context()
        escaped = value.replace('"', '\\"')
        self._api_core.LT_execute(f"win -a {page}")
        self._api_core.LT_execute(f'{obj}.text$ = "{escaped}"')

    # ── axis label visibility (TODO-5) ───────────────────────────────────

    def show_label(self) -> None:
        """
        Show the axis label (title).

        Corresponds to: ``xb.show = 1`` / ``yl.show = 1`` etc. in LabTalk.
        Ref: https://www.originlab.com/doc/LabTalk/ref/Graphic-objs
        """
        obj = self._get_title_obj_name()
        page, lid = self._page_context()
        self._api_core.LT_execute(f"win -a {page}")
        self._api_core.LT_execute(f"{obj}.show = 1")

    def hide_label(self) -> None:
        """
        Hide the axis label (title).

        Corresponds to: ``xb.show = 0`` / ``yl.show = 0`` etc. in LabTalk.
        """
        obj = self._get_title_obj_name()
        page, lid = self._page_context()
        self._api_core.LT_execute(f"win -a {page}")
        self._api_core.LT_execute(f"{obj}.show = 0")

    def get_label_visible(self) -> bool:
        """
        Get whether the axis label (title) is currently shown.

        Returns:
            bool: True if the label is visible.
        """
        obj = self._get_title_obj_name()
        page, lid = self._page_context()
        self._api_core.LT_execute(f"win -a {page}")
        self._api_core.LT_execute(f"_ax_sl = {obj}.show")
        val = self._api_core.LT_get_var("_ax_sl")
        if math.isnan(val):
            raise OriginCommandResponceError(f"NaN value is received while getting label visibility. (page={page}, layer={lid})")
        return bool(int(val))


# ================== GraphLayer Class ==================

class GraphLayer(OriginObjectWrapper[oext_types.GraphLayer]):
    """
    Graph layer for plotting data.
    Wrapper class that wraps OriginExt.OriginExt.GraphLayer.

    Corresponds to: originpro.GLayer, OriginExt.OriginExt.GraphLayer
    """

    def __init__(self, layer: oext_types.GraphLayer, api_core: APP, id: int, parent_page: Optional['GraphPage'] = None):
        """
        Initialize GraphLayer wrapper with hierarchical references.

        Args:
            layer: Original OriginExt.GraphLayer instance to wrap
            api_core: APP instance reference for LabTalk access
            id: Layer ID
            parent_page: Parent GraphPage reference (protected member)
        """
        super().__init__(layer, api_core)

        self._id = id
        self._parent_page = parent_page


    @property
    def id(self)->int:
        """Get the graph layer ID"""
        return self._id

    def _get_data_plot_list(self) -> list[DataPlot]:
        """Return all DataPlot objects in this layer.

        Iterates over the underlying OriginExt GraphLayer's DataPlots
        collection and wraps each entry as a ``DataPlot`` instance.

        Returns:
            list[DataPlot]: All plots in this layer.
        """
        plots: list[DataPlot] = []
        for raw_plot in self._obj.DataPlots:
            plots.append(DataPlot(raw_plot, self))
        return plots

    @property
    def data_plots(self) -> list[DataPlot]:
        """All data plots in this layer as a list."""
        return self._get_data_plot_list()

    @property
    def graph_objects(self):
        """Collection of graph objects in this layer"""
        return self._obj.GraphObjects

    def __iter__(self) -> Iterator[DataPlot]:
        """Iterate over data plots in this layer."""
        return iter(self._get_data_plot_list())

    def __getitem__(self, index: int) -> DataPlot:
        """Get data plot by 0-based index."""
        plots = self._get_data_plot_list()
        return plots[index]

    def get_axis(self, axis_type: AxisType) -> Axis:
        """
        Get an axis object for the specified axis type.

        Args:
            axis_type: Type of axis (X, Y, Z, ERROR, X2, or Y2)

        Returns:
            Axis: Axis wrapper object
        """
        return Axis(self, axis_type)

    # def rescale(self) -> None:
    #     """
    #     Rescale all axes in the layer.
    #     """
    #     cmd = "layer -a 100"  # Rescale all axes
    #     self.api_core.LT_execute(cmd)

    def rescale_axis(self, axis_type: AxisType) -> None:
        """
        Rescale a specific axis.

        Args:
            axis_type: Type of axis to rescale
        """
        axis_letter = axis_type.name[0]  # X, Y, Z, or E
        cmd = f"layer -a {axis_letter}"
        self.api_core.LT_execute(cmd)

    def add_xy_plot(self, worksheet: Worksheet, x_col: int, y_col: int,
                   plot_type: XYPlotType,
                   color: Optional[ColorSpec] = None) -> DataPlot:
        """
        Add an XY plot to this layer.

        Args:
            worksheet: Worksheet containing data
            x_col: X column index (0-based)
            y_col: Y column index (0-based) or -1 for all columns after x_col
            plot_type: XYPlotType enum (defaults to LINE_SYMBOL)
            color: Plot color. Accepts an int (Origin color index 1-24),
                   an (R, G, B) tuple, or an OriginColorIndex constant.
                   When omitted, Origin uses the template default.

        Returns:
            DataPlot: The created data plot
        """

        # Get fullname using worksheet's page name
        book_name = worksheet._obj.GetPage().GetName()
        worksheet_full_name = f"[{book_name}]{worksheet.name}"

        # Use simple layer name without page specification
        # Get parent page name from the parent page reference if available
        if self._parent_page is not None:
            parent_page_name = self._parent_page.name
        else:
            try:
                parent_page_name = self._obj.GetPage().GetName()
            except Exception as e:
                warnings.warn(
                    f"Could not retrieve parent page name from layer object: {e}. Falling back to 'Graph1'.",
                    RuntimeWarning,
                    stacklevel=2,
                )
                parent_page_name = "Graph1"

        layer_name = f"Layer{self._id + 1}"
        graph_full_name = f"[{parent_page_name}]{layer_name}"

        # Build the plotxy LabTalk command
        color_str = f" color:={color_to_lt_str(color)}" if color is not None else ""
        cmd = (
            f"plotxy iy:={worksheet_full_name}!({x_col+1},{y_col+1}) "
            f"plot:={plot_type.value.plot_id}"
            f"{color_str} "
            f"ogl:={graph_full_name}!"
        )

        self.api_core.LT_execute(cmd)

        raw_plots = list(self._obj.DataPlots)
        if not raw_plots:
            raise OriginNotFoundError("No DataPlot found in layer after executing plotxy.")
        return DataPlot(raw_plots[-1], self, plot_type)

    def group_plots(self, group_mode: Optional[GroupMode] = None) -> None:
        """
        Group plots in this layer.

        Args:
            group_mode: GroupMode enum for grouping type
        """
        if group_mode is None:
            group_mode = GroupMode.DEPENDENT

        if self._parent_page is not None:
            self.api_core.LT_execute(f"win -a {self._parent_page.name}")
        self.api_core.LT_execute(f"layer -g {group_mode.value}")

    def get_parent_graph(self) -> oext_types.GraphPage:
        """
        Get the parent graph page.

        Returns:
            GraphPage: Parent graph page object
        """
        return oext.GraphPage_GetPage(self._obj)

    def set_scale(self, x: float, y: float, width: float, height: float) -> None:
        """
        Set the drawing position and size of this layer on the page.

        All values are in **percentage of the page dimensions** (0-100 %).
        Origin's default layer occupies roughly x=15, y=15, width=70, height=70.

        Internally uses ``[page]!layerN.left/top/width/height`` LabTalk
        properties, which accept percentage values.

        Args:
            x: Left edge position as % of page width (0-100).
            y: Top edge position as % of page height (0-100).
            width: Layer width as % of page width (0-100).
            height: Layer height as % of page height (0-100).
        """
        if self._parent_page is not None:
            page_name = self._parent_page.name
            layer_no = self._id + 1  # 1-based layer number
            # Use explicit [page]!layerN.prop notation so the target is
            # unambiguous and independent of the currently active layer.
            self.api_core.LT_execute(f"[{page_name}]!layer{layer_no}.left = {x}")
            self.api_core.LT_execute(f"[{page_name}]!layer{layer_no}.top = {y}")
            self.api_core.LT_execute(f"[{page_name}]!layer{layer_no}.width = {width}")
            self.api_core.LT_execute(f"[{page_name}]!layer{layer_no}.height = {height}")
        else:
            # Fallback: activate and use the generic layer object
            self.api_core.LT_execute(f"layer.left = {x}")
            self.api_core.LT_execute(f"layer.top = {y}")
            self.api_core.LT_execute(f"layer.width = {width}")
            self.api_core.LT_execute(f"layer.height = {height}")

    def get_legend(self) -> Legend:
        """
        Get the legend object for this layer.

        Returns:
            Legend: Legend wrapper for this graph layer.
        """
        return Legend(self)

    def get_scale(self) -> tuple:
        """
        Get the current drawing position and size of this layer.

        All values are in **percentage of the page dimensions** (0-100 %).

        Reads ``[page]!layerN.left``, ``.top``, ``.width``, ``.height``
        via LabTalk.

        Returns:
            tuple: ``(x, y, width, height)`` in % of page dimensions
                   (left, top, layer-width, layer-height).
                   Returns ``(0.0, 0.0, 0.0, 0.0)`` if the parent page is unavailable.
        """
        if self._parent_page is None:
            return (0.0, 0.0, 0.0, 0.0)

        page_name = self._parent_page.name
        layer_no  = self._id + 1

        def _read(prop: str) -> float:
            self.api_core.LT_execute(
                f"_gs_tmp = [{page_name}]!layer{layer_no}.{prop}"
            )
            val = self.api_core.LT_get_var("_gs_tmp")
            return 0.0 if math.isnan(val) else float(val)

        x      = _read("left")
        y      = _read("top")
        width  = _read("width")
        height = _read("height")
        return (x, y, width, height)
