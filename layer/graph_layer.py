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
from .enums import ColorMap, AxisType, XYPlotType, GroupMode
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

    def __init__(self, plot, api_core: 'APP'):
        """
        Initialize DataPlot wrapper with hierarchical references.

        Args:
            plot: Original OriginExt.DataPlot instance to wrap or plot index
            api_core: APP instance reference for LabTalk access
        """
        if plot is None:
            raise ValueError("DataPlot cannot be created with None plot object")
        self._plot = plot
        self.__API_core = api_core
        self._is_index = isinstance(plot, int)


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
                   plot_type: XYPlotType) -> DataPlot:
        """
        Add an XY plot to this layer.

        Args:
            worksheet: Worksheet containing data
            x_col: X column index (0-based)
            y_col: Y column index (0-based) or -1 for all columns after x_col
            plot_type: XYPlotType enum (defaults to LINE_SYMBOL)

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

        # Create the plot using LabTalk
            # Plot specific X and Y columns
        cmd = \
            f"plotxy iy:={worksheet_full_name}!({x_col+1},{y_col+1}) "+\
            f"plot:={plot_type.value.plot_id} "+\
            f"ogl:={graph_full_name}!"
        
        print(f"[DEBUG] Executing plot command: {cmd}")
        
        # Use API core for LabTalk execution
        self.api_core.LT_execute(cmd)
        
        print("[DEBUG] Plot command executed successfully")
        
        # Force layer refresh after plot creation
        self.api_core.LT_execute("layer -u")
        
        # Since we can't access the plot object directly through OriginExt,
        # we'll create a DataPlot with a reference to the layer and index
        # The actual plot object can be accessed through LabTalk when needed
        self.api_core.LT_execute("plot_index = layer.plot.count - 1")
        
        # Use LT_get_var to get the variable value
        plot_index = self.api_core.LT_get_var("plot_index")
        print(f"[DEBUG] Plot index: {plot_index}")
        
        # Create a DataPlot with layer reference and index
        return DataPlot(int(plot_index), self.api_core)

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
