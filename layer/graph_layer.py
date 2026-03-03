"""
Graph layer classes for OriginExt wrappers.

This module contains wrapper classes for Origin graph layers including:
- DataPlot
- Axis
- GraphLayer
"""
from __future__ import annotations

import OriginExt.OriginExt as oext_types
import OriginExt._OriginExt as oext
from enum import Enum
from typing import Optional, Union, TYPE_CHECKING, List
from collections.abc import Iterator

from ..base import OriginObjectWrapper
from .enums import ColorMap, AxisType, XYPlotType, GroupMode, OriginColorIndex, ColorSpec, color_to_lt_str, LegendLayout
from .worksheet import Worksheet

from ..base import APP

if TYPE_CHECKING:
    from ..pages import GraphPage


# ================== DataPlot Class ==================

class DataPlot:
    """
    Data plot in a graph layer.
    Wrapper class that wraps OriginExt.OriginExt.DataPlot.

    Corresponds to: originpro.DataPlot, OriginExt.OriginExt.DataPlot
    """

    def __init__(self, plot, api_core: 'APP',
                 page_name: Optional[str] = None,
                 layer_index: Optional[int] = None):
        """
        Initialize DataPlot wrapper with hierarchical references.

        Args:
            plot: Original OriginExt.DataPlot instance to wrap or plot index
            api_core: APP instance reference for LabTalk access
            page_name: Short name of the parent graph page (enables layer activation)
            layer_index: 0-based index of the parent graph layer
        """
        if plot is None:
            raise ValueError("DataPlot cannot be created with None plot object")
        self._plot = plot
        self.__API_core = api_core
        self._is_index = isinstance(plot, int)
        self._page_name = page_name
        self._layer_index = layer_index


    @property
    def api_core(self) -> 'APP':
        """Get the API core reference"""
        return self.__API_core

    @property
    def name(self) -> str:
        """Short name of the data plot"""
        if self._is_index:
            # Use LabTalk to get the plot name
            self.api_core.LT_execute(f"plot_name = layer.plot({self._plot + 1}).name$")
            return self.api_core.LT_get_str("plot_name")
        return self._plot.Name

    # @property
    # def parent_layer(self) -> 'GraphLayer':
    #     """Parent graph layer"""
    #     return GraphLayer(self._plot.Parent, self.api_core)

    @property
    def worksheet(self):
        """Worksheet containing the plotted data"""
        if self._is_index:
            # Use LabTalk to get the worksheet name
            self.api_core.LT_execute(f"ws_name = layer.plot({self._plot + 1}).name$")
            ws_name = self.api_core.LT_get_str("ws_name")
            # Get the worksheet object from the active page
            self.api_core.LT_execute(f"ws = {ws_name}")
            # Return None for now since we can't easily get the worksheet object
            return None
        worksheet_obj = oext.DataPlot_GetWorksheet(self._plot)
        return Worksheet(worksheet_obj, self.api_core)

    def _plot_range(self) -> Optional[str]:
        """Return the LabTalk range string for this plot, e.g. ``[Graph1]1!(1)``.

        Returns None when the page/layer context is unavailable.
        """
        if self._page_name is not None and self._layer_index is not None:
            # Range notation: [PageName]LayerIndex(1-based)!(PlotIndex(1-based))
            return f"[{self._page_name}]{self._layer_index + 1}!({self._plot + 1})"
        return None

    @property
    def color(self) -> int:
        """Line/symbol color as an Origin color value.

        Reads the plot color via ``layer.plot(n).color`` in LabTalk after
        activating the parent graph layer.  For index colors (1-24) the
        returned integer matches ``OriginColorIndex``.  For custom RGB
        colors the returned value is the Windows COLORREF integer
        (R + G*256 + B*65536).

        Note: Colors applied at plot-creation time (via ``add_xy_plot``'s
        *color* argument) are stored in the group's style and may only be
        readable after ungrouping the plots first
        (``layer.group_plots(GroupMode.NONE)``).
        Setting accepts an int, (R,G,B) tuple, or ``OriginColorIndex``.
        """
        if self._page_name is None or self._layer_index is None:
            return 0
        self.api_core.LT_execute(f"win -a {self._page_name}")
        self.api_core.LT_execute(f"layer {self._layer_index + 1}")
        self.api_core.LT_execute(
            f"colorTmpVal = layer.plot({self._plot + 1}).color"
        )
        raw = self.api_core.LT_get_var("colorTmpVal")
        import math
        if math.isnan(raw):
            return 0
        return int(raw)

    @color.setter
    def color(self, value: ColorSpec) -> None:
        """Set the plot color.

        Activates the parent graph layer then assigns the color via
        ``layer.plot(n).color``.  Works reliably after ungrouping
        (``layer.group_plots(GroupMode.NONE)``).
        """
        color_str = color_to_lt_str(value)
        if self._page_name is not None and self._layer_index is not None:
            self.api_core.LT_execute(f"win -a {self._page_name}")
            self.api_core.LT_execute(f"layer {self._layer_index + 1}")
        self.api_core.LT_execute(
            f"layer.plot({self._plot + 1}).color = {color_str}"
        )

    @property
    def color_map(self) -> ColorMap:
        color_map_str = oext.DataPlot_GetColorMap(self._plot)
        return ColorMap(color_map_str)


    @color_map.setter
    def color_map(self, value: ColorMap) -> None:
        """Set color map for the data plot"""
        try:
            oext.DataPlot_SetColorMap(self._plot, value.value)
        except Exception as e:
            print(f"Failed to set color map: {e}")
            # Try alternative method
            try:
                self._plot.Execute(f"set {self._plot.Range} -cmap {value.value}")
            except Exception as e2:
                print(f"Alternative colormap setting also failed: {e2}")
                raise RuntimeError(f"Failed to set color map: {e2}") from e2

    @property
    def shape_list(self) -> list[int]:
        """Get shape list"""
        return oext.DataPlot_GetShapeList(self._plot)

    @shape_list.setter
    def shape_list(self, value: list[int]) -> None:
        """Set shape list for the data plot"""
        oext.DataPlot_SetShapeList(self._plot, value)

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
            axis_type: Type of axis (X, Y, Z, or ERROR)
        """
        self._axis_type = axis_type
        self._obj = self._get_originext_axis(graph_layer)

    def _get_originext_axis(self, graph_layer: 'GraphLayer'):
        """Get the OriginExt axis object from a GraphLayer wrapper"""
        obj = graph_layer._obj
        # Ensure it's the correct type by checking if it has axis methods
        return obj

    @property
    def axis_type(self) -> AxisType:
        """Get axis type"""
        return self._axis_type

    def get_range(self) -> tuple[float, float]:
        """
        Get the axis range (min, max).

        Returns:
            tuple: (min_value, max_value)
        """
        axis_letter = self._get_axis_letter()
        min_val = self._obj.LT_get_float(f"{axis_letter}.from")
        max_val = self._obj.LT_get_float(f"{axis_letter}.to")
        return (min_val, max_val)

    def set_range(self, min_val: float, max_val: float) -> None:
        """
        Set the axis range.

        Args:
            min_val: Minimum value
            max_val: Maximum value

        Returns:
            bool: True if successful
        """
        axis_letter = self._get_axis_letter()
        self._obj.LT_execute(f"{axis_letter}.from = {min_val}; {axis_letter}.to = {max_val}")
        return

    def get_reverse(self) -> bool:
        """
        Get whether the axis is reversed.

        Returns:
            bool: True if reversed
        """
        axis_letter = self._get_axis_letter()
        return bool(self._obj.LT_get_int(f"{axis_letter}.reverse"))

    def set_reverse(self, reverse: bool) -> None:
        """
        Set whether the axis is reversed.

        Args:
            reverse: True to reverse the axis

        Returns:
            bool: True if successful
        """
        axis_letter = self._get_axis_letter()
        self._obj.LT_execute(f"{axis_letter}.reverse = {1 if reverse else 0}")
        return 

    def _get_axis_letter(self) -> str:
        """Get the axis letter for LabTalk commands"""
        if self._axis_type == AxisType.X:
            return "X"
        elif self._axis_type == AxisType.Y:
            return "Y"
        elif self._axis_type == AxisType.Z:
            return "Z"
        elif self._axis_type == AxisType.ERROR:
            return "E"
        else:
            return "X"  # Default fallback


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

    @property
    def data_plots(self):
        """Collection of data plots in this layer"""
        # Ensure this layer is active
        self.api_core.LT_execute("layer -a")
        # Force layer refresh
        self.api_core.LT_execute("layer -u")
        # Use LabTalk to get plot count instead of direct access
        self.api_core.LT_execute("layer -c")
        # Return a list with index-based DataPlot objects
        plot_count = self.api_core.LT_get_var("count")
        print(f"[DEBUG] Plot count: {plot_count}")
        return [DataPlot(i, self.api_core) for i in range(int(plot_count))]

    @property
    def graph_objects(self):
        """Collection of graph objects in this layer"""
        return self._obj.GraphObjects

    def __iter__(self) -> Iterator[DataPlot]:
        """Iterate over data plots"""
        # Use LabTalk to get plot count
        self.api_core.LT_execute("layer -c")
        plot_count = self.api_core.LT_get_var("count")
        for i in range(int(plot_count)):
            yield DataPlot(i, self.api_core)

    def __getitem__(self, index: int) -> DataPlot:
        """Get data plot by index"""
        return DataPlot(index, self.api_core)

    def get_axis(self, axis_type: AxisType) -> Axis:
        """
        Get an axis object for the specified axis type.

        Args:
            axis_type: Type of axis (X, Y, Z, or ERROR)

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
            # Fallback: try to get from layer object
            try:
                parent_page_name = self._obj.GetPage().GetName()
            except:
                # Final fallback: use a generic name
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

        print(f"[DEBUG] Executing plot command: {cmd}")

        # Use API core for LabTalk execution
        self.api_core.LT_execute(cmd)

        print("[DEBUG] Plot command executed successfully")

        # Activate the target graph layer so that LabTalk's 'layer' object
        # points to it before querying the plot count.
        self.api_core.LT_execute(f"win -a {parent_page_name}")
        self.api_core.LT_execute(f"layer {self._id + 1}")
        self.api_core.LT_execute("layer -c")  # stores count in 'count' variable
        plot_count = self.api_core.LT_get_var("count")
        plot_index = int(plot_count) - 1
        print(f"[DEBUG] Plot count: {plot_count}, plot index: {plot_index}")

        return DataPlot(plot_index, self.api_core,
                        page_name=parent_page_name,
                        layer_index=self._id)

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
