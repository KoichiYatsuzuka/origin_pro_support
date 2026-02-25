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
from .enums import ColorMap, AxisType, PlotType, GroupMode

if TYPE_CHECKING:
    from ..origin_instance import OriginInstance


# ================== DataPlot Class ==================

class DataPlot:
    """
    Data plot in a graph layer.
    Wrapper class that wraps OriginExt.OriginExt.DataPlot.

    Corresponds to: originpro.DataPlot, OriginExt.OriginExt.DataPlot
    """

    def __init__(self, plot, parent: Optional['OriginObjectWrapper'] = None, 
                 origin_instance: Optional['OriginInstance'] = None):
        """
        Initialize DataPlot wrapper with hierarchical references.

        Args:
            plot: Original OriginExt.DataPlot instance to wrap
            parent: Parent wrapper object (for hierarchical navigation)
            origin_instance: Root OriginInstance reference (for LabTalk access)
        """
        self._plot = plot
        self._parent = parent
        self._origin_instance = origin_instance

    @property
    def parent(self) -> Optional['OriginObjectWrapper']:
        return self._parent

    @property
    def origin_instance(self) -> Optional['OriginInstance']:
        """Get the root OriginInstance reference"""
        if self._origin_instance:
            return self._origin_instance
        elif self._parent:
            return self._parent.origin_instance
        return None

    @property
    def name(self) -> str:
        """Short name of the data plot"""
        return self._plot.Name

    @property
    def parent_layer(self) -> 'GraphLayer':
        """Parent graph layer"""
        return GraphLayer(self._plot.Parent)

    @property
    def color_map(self) -> ColorMap:
        try:
            color_map_str = oext.DataPlot_GetColorMap(self._plot)
            return ColorMap(color_map_str)
        except ValueError:
            return ColorMap.CANDY  # Default fallback

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

    @property
    def shape_list(self) -> list[int]:
        """Get shape list"""
        try:
            return oext.DataPlot_GetShapeList(self._plot)
        except Exception as e:
            print(f"Failed to get shape list: {e}")
            return []

    @shape_list.setter
    def shape_list(self, value: list[int]) -> None:
        """Set shape list for the data plot"""
        try:
            oext.DataPlot_SetShapeList(self._plot, value)
        except Exception as e:
            print(f"Failed to set shape list: {e}")

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
        try:
            return self._plot.ChangeData(data_obj, designation, keep_modifiers)
        except Exception as e:
            print(f"Failed to change data: {e}")
            return False


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
        try:
            axis_letter = self._get_axis_letter()
            min_val = self._obj.LT_get_float(f"{axis_letter}.from")
            max_val = self._obj.LT_get_float(f"{axis_letter}.to")
            return (min_val, max_val)
        except Exception as e:
            print(f"Failed to get {self._axis_type.name} axis range: {e}")
            return (0.0, 1.0)

    def set_range(self, min_val: float, max_val: float) -> bool:
        """
        Set the axis range.

        Args:
            min_val: Minimum value
            max_val: Maximum value

        Returns:
            bool: True if successful
        """
        try:
            axis_letter = self._get_axis_letter()
            self._obj.LT_execute(f"{axis_letter}.from = {min_val}; {axis_letter}.to = {max_val}")
            return True
        except Exception as e:
            print(f"Failed to set {self._axis_type.name} axis range: {e}")
            return False

    def get_reverse(self) -> bool:
        """
        Get whether the axis is reversed.

        Returns:
            bool: True if reversed
        """
        try:
            axis_letter = self._get_axis_letter()
            return bool(self._obj.LT_get_int(f"{axis_letter}.reverse"))
        except Exception as e:
            print(f"Failed to get {self._axis_type.name} axis reverse status: {e}")
            return False

    def set_reverse(self, reverse: bool) -> bool:
        """
        Set whether the axis is reversed.

        Args:
            reverse: True to reverse the axis

        Returns:
            bool: True if successful
        """
        try:
            axis_letter = self._get_axis_letter()
            self._obj.LT_execute(f"{axis_letter}.reverse = {1 if reverse else 0}")
            return True
        except Exception as e:
            print(f"Failed to set {self._axis_type.name} axis reverse: {e}")
            return False

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

    def __init__(self, layer: oext_types.GraphLayer, parent: Optional['OriginObjectWrapper'] = None, 
                 origin_instance: Optional['OriginInstance'] = None):
        """
        Initialize GraphLayer wrapper with hierarchical references.

        Args:
            layer: Original OriginExt.GraphLayer instance to wrap
            parent: Parent wrapper object (for hierarchical navigation)
            origin_instance: Root OriginInstance reference (for LabTalk access)
        """
        super().__init__(layer, parent, origin_instance)

    @property
    def data_plots(self):
        """Collection of data plots in this layer"""
        return self._obj.DataPlots

    @property
    def graph_objects(self):
        """Collection of graph objects in this layer"""
        return self._obj.GraphObjects

    def __iter__(self) -> Iterator[DataPlot]:
        """Iterate over data plots"""
        for plot in self._obj:
            yield DataPlot(plot, self, self.origin_instance)

    def __getitem__(self, index: int) -> DataPlot:
        """Get data plot by index"""
        return DataPlot(self._obj[index], self, self.origin_instance)

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
        try:
            self._obj.Execute("layer -a 100")  # Rescale all axes
        except Exception as e:
            print(f"Failed to rescale layer: {e}")

    def rescale_axis(self, axis_type: AxisType) -> None:
        """
        Rescale a specific axis.

        Args:
            axis_type: Type of axis to rescale
        """
        try:
            axis_letter = axis_type.name[0]  # X, Y, Z, or E
            self._obj.Execute(f"layer -a {axis_letter}")
        except Exception as e:
            print(f"Failed to rescale {axis_type.name} axis: {e}")

    def add_xy_plot(self, worksheet, x_col: int, y_col: int = -1, 
                   plot_type: Optional[PlotType] = None) -> DataPlot:
        """
        Add an XY plot to this layer.

        Args:
            worksheet: Worksheet containing data
            x_col: X column index (0-based)
            y_col: Y column index (0-based) or -1 for all columns after x_col
            plot_type: PlotType enum (defaults to LINE_SYMBOL)

        Returns:
            DataPlot: The created data plot
        """
        if plot_type is None:
            plot_type = PlotType.LINE_SYMBOL

        try:
            # Get the OriginExt worksheet object
            worksheet_obj = worksheet._obj
            
            # Create the plot using LabTalk
            if y_col == -1:
                # Plot all columns after x_col as Y
                cmd = f"plotxy {worksheet_obj.Name}!col({x_col + 1}) plot:={plot_type.value}"
            else:
                # Plot specific X and Y columns
                cmd = f"plotxy {worksheet_obj.Name}!col({x_col + 1}) col({y_col + 1}) plot:={plot_type.value}"
            
            self._obj.Execute(cmd)
            
            # Get the newly created plot (last one in the collection)
            plots = list(self._obj.DataPlots)
            if plots:
                return DataPlot(plots[-1], self, self.origin_instance)
            
        except Exception as e:
            print(f"Failed to add XY plot: {e}")
            
        return None

    def group_plots(self, group_mode: Optional[GroupMode] = None) -> None:
        """
        Group plots in this layer.

        Args:
            group_mode: GroupMode enum for grouping type
        """
        if group_mode is None:
            group_mode = GroupMode.NONE
        
        try:
            group_value = group_mode.value
            self._obj.Execute(f"layer -g {group_value}")
        except Exception as e:
            print(f"Failed to group plots: {e}")

    def get_parent_graph(self) -> oext_types.GraphPage:
        """
        Get the parent graph page.

        Returns:
            GraphPage: Parent graph page object
        """
        return oext.GraphPage_GetPage(self._obj)
