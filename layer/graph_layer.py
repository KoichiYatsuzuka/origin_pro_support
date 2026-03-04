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
from .enums import ColorMap, AxisType, XYPlotType, GroupMode, OriginColorIndex, ColorSpec, color_to_lt_str, LegendLayout, TickType, MarkerShape
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

    def __init__(self, plot: oext_types.DataPlot, graph_layer: 'GraphLayer'):
        """
        Initialize DataPlot wrapper.

        Args:
            plot: ``OriginExt.OriginExt.DataPlot`` instance to wrap.
            graph_layer: Parent ``GraphLayer`` that owns this plot.
        """
        if plot is None:
            raise ValueError("DataPlot cannot be created with None plot object")
        if not isinstance(plot, oext_types.DataPlot):
            raise TypeError(
                f"plot must be an OriginExt.DataPlot instance, got {type(plot).__name__}"
            )
        self._plot: oext_types.DataPlot = plot
        self._graph_layer: 'GraphLayer' = graph_layer

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

    def _has_symbol(self) -> bool:
        """Return True if this plot type supports symbol (marker) display.

        Checks for the presence of a top-level ``Symbol`` node in the plot's
        theme tree, which is present for scatter and line+symbol plots but
        absent for line-only and other non-symbol plot types.
        """
        theme = self._plot.GetTheme()
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

        Reads the plot color via ``[Page]!layerN.plot(M).color`` in LabTalk.
        For index colors (1-24) the returned integer matches
        ``OriginColorIndex``.  For custom RGB colors the returned value is
        the Windows COLORREF integer (R + G*256 + B*65536).

        Note: Colors applied at plot-creation time (via ``add_xy_plot``'s
        *color* argument) are stored in the group's style and may only be
        readable after ungrouping the plots first
        (``layer.group_plots(GroupMode.NONE)``).
        Setting accepts an int, (R,G,B) tuple, or ``OriginColorIndex``.
        """
        page_name, layer_id = self._page_context()
        plot_id = self._resolve_plot_id()
        self.api_core.LT_execute(layer_plot_get(page_name, layer_id, plot_id, "color", "_dp_color"))
        raw = self.api_core.LT_get_var("_dp_color")
        if math.isnan(raw):
            raise OriginCommandResponceError(f"NaN value is received while getting plot color. (page={page_name}, layer={layer_id}, plot={plot_id})")
        return int(raw)

    @color.setter
    def color(self, value: ColorSpec) -> None:
        """Set the plot color.

        Works reliably after ungrouping
        (``layer.group_plots(GroupMode.NONE)``).
        """
        page_name, layer_id = self._page_context()
        plot_id = self._resolve_plot_id()
        color_str = color_to_lt_str(value)
        self.api_core.LT_execute(layer_plot_set(page_name, layer_id, plot_id, "color", color_str))

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
        """Symbol (marker) size for this plot.

        Reads and writes ``[Page]!layerN.plot(M).symbol.size`` via LabTalk.

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
            value: Symbol size as a positive float.

        Raises:
            ValueError: If the plot type does not support symbols, or value is not positive.
        """
        self._assert_has_symbol()
        if value <= 0:
            raise ValueError(f"symbol_size must be positive, got {value}")
        self.api_core.LT_execute(plot_set_symbol_size(self._plot.GetDatasetName(), value))

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
        self.api_core.LT_execute(plot_set_symbol_kind(self._plot.GetDatasetName(), value.value))

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
        """Activate the parent layer so LabTalk's ``legend`` object is correct."""
        if self._layer._parent_page is not None:
            page_name = self._layer._parent_page.name
            layer_no = self._layer._id + 1
            self._layer.api_core.LT_execute(f"win -a {page_name}")
            self._layer.api_core.LT_execute(f"layer {layer_no}")

    # ── visibility ───────────────────────────────────────────────────────

    @property
    def visible(self) -> bool:
        """Whether the legend is visible.

        Corresponds to: ``legend.show`` in LabTalk.
        """
        self._activate()
        self._layer.api_core.LT_execute("_leg_show = legend.show")
        import math
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
        self._layer.api_core.LT_execute(f"legend.show = {1 if value else 0}")

    # ── text ─────────────────────────────────────────────────────────────

    @property
    def text(self) -> str:
        """Raw text content of the legend object.

        Corresponds to: ``legend.text$`` in LabTalk.
        """
        self._activate()
        self._layer.api_core.LT_execute("_leg_text$ = legend.text$")
        return self._layer.api_core.LT_get_str("_leg_text")

    @text.setter
    def text(self, value: str) -> None:
        """Set the raw text of the legend.

        Args:
            value: New legend text. Use ``\\(\\n)`` for line breaks.
        """
        self._activate()
        escaped = value.replace('"', '\\"')
        self._layer.api_core.LT_execute(f'legend.text$ = "{escaped}"')

    # ── position ─────────────────────────────────────────────────────────

    def get_position(self) -> tuple:
        """Get the legend position in axis-scale units.

        Corresponds to: ``legend.x``, ``legend.y`` in LabTalk.

        Returns:
            tuple: (x, y) centre coordinates of the legend in axis-scale units.
        """
        import math
        self._activate()
        self._layer.api_core.LT_execute("_leg_x = legend.x")
        self._layer.api_core.LT_execute("_leg_y = legend.y")
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
        self._layer.api_core.LT_execute(f"legend.x = {x}; legend.y = {y}")

    def reset_position(self) -> None:
        """Reset the legend to its default position.

        Corresponds to: ``legend -d`` in LabTalk.
        """
        self._activate()
        self._layer.api_core.LT_execute("legend -d")

    # ── font size ─────────────────────────────────────────────────────────

    @property
    def font_size(self) -> int:
        """Font size of the legend text in points.

        Corresponds to: ``legend.fsize`` in LabTalk.
        """
        import math
        self._activate()
        self._layer.api_core.LT_execute("_leg_fsize = legend.fsize")
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
        self._layer.api_core.LT_execute(f"legend.fsize = {value}")

    # ── background box ────────────────────────────────────────────────────

    @property
    def background(self) -> int:
        """Background box style of the legend.

        0 = none, 1 = black border, 2 = shadow, 3 = white-out,
        4 = black border + white-out.

        Corresponds to: ``legend.background`` in LabTalk.
        """
        import math
        self._activate()
        self._layer.api_core.LT_execute("_leg_bg = legend.background")
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
        self._layer.api_core.LT_execute(f"legend.background = {value}")

    # ── layout ────────────────────────────────────────────────────────────

    def set_layout(self, layout: LegendLayout) -> None:
        """Rearrange legend entries horizontally or vertically.

        Args:
            layout: LegendLayout.VERTICAL or LegendLayout.HORIZONTAL.
        """
        self._activate()
        self._layer.api_core.LT_execute(f"legend -{layout.value}")

    # ── reconstruct ───────────────────────────────────────────────────────

    def reconstruct(self) -> None:
        """Reconstruct the legend to best-fit the current plots.

        Equivalent to menu Graph > Legend > Reconstruct Legend.

        Corresponds to: ``legend -r`` in LabTalk.
        """
        self._activate()
        self._layer.api_core.LT_execute("legend -r")

    def update(self) -> None:
        """Create or update the legend on the active graph layer.

        Corresponds to: ``legend`` (no option) in LabTalk.
        """
        self._activate()
        self._layer.api_core.LT_execute("legend")


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
        Ref: https://www.originlab.com/doc/LabTalk/ref/Axis-cmd

        Returns:
            int: Current number of minor ticks between major ticks.
        """
        pg_ax = self._get_axis_pg_letter()
        page, lid = self._page_context()
        self._api_core.LT_execute(axis_pg(pg_ax, "M", "_ax_minor"))
        val = self._api_core.LT_get_var("_ax_minor")
        if math.isnan(val):
            raise OriginCommandResponceError(f"NaN value is received while getting minor tick count. (page={page}, layer={lid})")
        return int(val)

    def set_minor_tick_count(self, count: int) -> None:
        """
        Set the number of minor ticks between major ticks for this axis.

        Uses ``axis -ps axis M value`` (type ``M`` = number of Minor ticks).
        Ref: https://www.originlab.com/doc/LabTalk/ref/Axis-cmd

        Args:
            count: Number of minor ticks between major ticks (non-negative int).
        """
        if count < 0:
            raise ValueError(f"count must be non-negative, got {count}")
        pg_ax = self._get_axis_pg_letter()
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
        self._api_core.LT_execute(f"layer {lid + 1}")
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
        self._api_core.LT_execute(f"layer {lid + 1}")
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
        self._api_core.LT_execute(f"layer {lid + 1}")
        self._api_core.LT_execute(f"{obj}.show = 1")

    def hide_label(self) -> None:
        """
        Hide the axis label (title).

        Corresponds to: ``xb.show = 0`` / ``yl.show = 0`` etc. in LabTalk.
        """
        obj = self._get_title_obj_name()
        page, lid = self._page_context()
        self._api_core.LT_execute(f"win -a {page}")
        self._api_core.LT_execute(f"layer {lid + 1}")
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
        self._api_core.LT_execute(f"layer {lid + 1}")
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

    def rescale(self) -> None:
        """
        Rescale all axes in the layer.
        """
        cmd = "layer -a 100"  # Rescale all axes
        self.api_core.LT_execute(cmd)

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

        plots = self._get_data_plot_list()
        if not plots:
            raise OriginNotFoundError("No DataPlot found in layer after executing plotxy.")
        return plots[-1]

    def group_plots(self, group_mode: Optional[GroupMode] = None) -> None:
        """
        Group plots in this layer.

        Args:
            group_mode: GroupMode enum for grouping type
        """
        if group_mode is None:
            group_mode = GroupMode.DEPENDENT

        cmd = f"layer -g {group_mode.value}"
        self.api_core.LT_execute(cmd)

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

        All values are in percentage of the page dimensions (0-100).
        Origin's default layer occupies roughly x=15, y=15, width=70, height=70.

        Corresponds to: layer.left = x; layer.top = y; layer.width = w; layer.height = h

        Args:
            x: Left position as percentage of page width (0-100)
            y: Top position as percentage of page height (0-100)
            width: Layer width as percentage of page width (0-100)
            height: Layer height as percentage of page height (0-100)
        """
        if self._parent_page is not None:
            page_name = self._parent_page.name
            layer_no = self._id + 1  # 1-based layer number
            # Activate the target layer so LabTalk's 'layer' object points to it
            self.api_core.LT_execute(f"win -a {page_name}")
            self.api_core.LT_execute(f"layer {layer_no}")
        # Set all four geometry properties in one command
        self.api_core.LT_execute(
            f"layer.left = {x}; layer.top = {y}; "
            f"layer.width = {width}; layer.height = {height}"
        )

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

        All values are in percentage of the page dimensions (0-100).

        Corresponds to: GetNumProp("left/top/width/height") on the OriginExt layer object.

        Returns:
            tuple: (x, y, width, height) in percentage of page dimensions,
                   or (0.0, 0.0, 0.0, 0.0) if unavailable
        """
        if self._parent_page is None:
            return (0.0, 0.0, 0.0, 0.0)

        import math

        def _safe(v):
            return 0.0 if math.isnan(v) else float(v)

        # GetNumProp reads geometry directly from the OriginExt layer object.
        # Values are in percentage of page dimensions (0-100).
        x      = _safe(self._obj.GetNumProp("left"))
        y      = _safe(self._obj.GetNumProp("top"))
        width  = _safe(self._obj.GetNumProp("width"))
        height = _safe(self._obj.GetNumProp("height"))

        return (x, y, width, height)
