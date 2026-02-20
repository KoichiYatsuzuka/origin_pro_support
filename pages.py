"""
Page classes for OriginExt wrappers.

This module contains wrapper classes for Origin pages including:
- Folder
- PageBase (base class)
- Page
- WorksheetPage
- GraphPage
- MatrixPage
- NotePage
"""
from __future__ import annotations

import OriginExt.OriginExt as oext_types

from typing import Iterator, TypeVar, TYPE_CHECKING, Optional
import numpy as np

from .base import OriginObjectWrapper

if TYPE_CHECKING:
    from .layers import Layer, Worksheet, GraphLayer, Matrixsheet, DataPlot, PlotType, ColorMap, GroupMode
    from . import OriginInstance


# ================== Folder Class ==================

class Folder(OriginObjectWrapper[oext_types.Folder]):
    """
    Folder in Origin project.
    Wrapper class that wraps OriginExt.OriginExt.Folder.

    Corresponds to: originpro.Folder, OriginExt.OriginExt.Folder
    """

    def __init__(self, folder: oext_types.Folder, parent: Optional['OriginObjectWrapper'] = None, 
                 origin_instance: Optional['OriginInstance'] = None):
        """
        Initialize Folder wrapper with hierarchical references.

        Args:
            folder: Original OriginExt.Folder instance to wrap
            parent: Parent wrapper object (for hierarchical navigation)
            origin_instance: Root OriginInstance reference (for LabTalk access)
        """
        super().__init__(folder, parent, origin_instance)

    @property
    def Name(self) -> str:
        """Folder name"""
        return self._obj.Name

    def get_pages(self) -> list[PageBase]:
        """
        Get all pages in this folder.

        Returns:
            list[PageBase]: List of pages in the folder
        """
        pages = []
        for page in self._obj.Pages:
            if page.Type == 1:  # Worksheet
                pages.append(WorksheetPage(page, self, self.origin_instance))
            elif page.Type == 2:  # Graph
                pages.append(GraphPage(page, self, self.origin_instance))
            elif page.Type == 3:  # Matrix
                pages.append(MatrixPage(page, self, self.origin_instance))
            elif page.Type == 4:  # Notes
                pages.append(NotePage(page, self, self.origin_instance))
        return pages


# ================== Type Variables ==================

TPageBase = TypeVar('TPageBase', bound=oext_types.PageBase)
TPage = TypeVar('TPage', bound=oext_types.Page)

# ================== Page Classes ==================

class PageBase(OriginObjectWrapper[TPageBase]):
    """
    Base class for all Origin page types.
    Wrapper class that wraps OriginExt.OriginExt.PageBase.
    Inherits common methods from OriginObjectWrapper.

    Corresponds to: originpro.PageBase, OriginExt.OriginExt.PageBase
    """

    def __init__(self, page: TPageBase, parent: Optional['OriginObjectWrapper'] = None, 
                 origin_instance: Optional['OriginInstance'] = None):
        """
        Initialize PageBase wrapper with hierarchical references.

        Args:
            page: Original OriginExt.PageBase instance to wrap
            parent: Parent wrapper object (for hierarchical navigation)
            origin_instance: Root OriginInstance reference (for LabTalk access)
        """
        super().__init__(page, parent, origin_instance)

    @property
    def Type(self) -> int:
        """Page type identifier"""
        return self._obj.Type

    @property
    def Parent(self) -> Folder:
        """Parent folder of this page"""
        return Folder(self._obj.Parent, None, self.origin_instance)

    def get_type(self) -> int:
        """
        Get the page type.

        Corresponds to: OriginExt.OriginExt.PageBase.GetType()

        Returns:
            int: Page type identifier
        """
        return self._obj.GetType()

    def get_parent(self) -> Folder:
        """
        Get the parent folder.

        Corresponds to: OriginExt.OriginExt.PageBase.GetParent()

        Returns:
            Folder: Parent folder
        """
        return Folder(self._obj.GetParent())


class Page(PageBase[TPage]):
    """
    Generic page class with layer management.
    Wrapper class that wraps OriginExt.OriginExt.Page.
    Extends PageBase with layer-specific methods.

    Corresponds to: OriginExt.OriginExt.Page
    """

    def __init__(self, page: TPage, parent: Optional['OriginObjectWrapper'] = None, 
                 origin_instance: Optional['OriginInstance'] = None):
        """
        Initialize Page wrapper with hierarchical references.

        Args:
            page: Original OriginExt.Page instance to wrap
            parent: Parent wrapper object (for hierarchical navigation)
            origin_instance: Root OriginInstance reference (for LabTalk access)
        """
        super().__init__(page, parent, origin_instance)

    @property
    def Layers(self):
        """Collection of layers in this page"""
        return self._obj.Layers

    def __iter__(self) -> Iterator[Layer]:
        """Iterate over layers"""
        return iter(self._obj)

    def __getitem__(self, index: int) -> Layer:
        """Get layer by index"""
        from .layers import Layer
        return Layer(self._obj[index], self, self.origin_instance)

    def __len__(self) -> int:
        """Get number of layers"""
        return len(self._obj)

    def add_layer(self, name: str = '') -> Layer:
        """
        Add a new layer to this page.

        Corresponds to: OriginExt.OriginExt.Page.AddLayer()

        Args:
            name: Optional name for the new layer

        Returns:
            Layer: The newly created layer
        """
        from .layers import Layer
        return Layer(self._obj.AddLayer(name), self, self.origin_instance)

    def get_layers(self) -> list[Layer]:
        """
        Get list of layers in this page.

        Corresponds to: OriginExt.OriginExt.Page.GetLayers()

        Returns:
            list[Layer]: List of layers
        """
        from .layers import Layer
        return [Layer(l, self, self.origin_instance) for l in self._obj.GetLayers()]

    def get_layer(self, index: int) -> Layer:
        """
        Get layer by index.

        Corresponds to: OriginExt.OriginExt.Page.GetLayer()

        Args:
            index: Layer index

        Returns:
            Layer: The layer at the specified index
        """
        from .layers import Layer
        return Layer(self._obj.GetLayer(index), self, self.origin_instance)

    def preview(self, fname: str) -> bool:
        """
        Generate a preview image.

        Corresponds to: OriginExt.OriginExt.Page.Preview()

        Args:
            fname: File name for the preview image

        Returns:
            bool: True if successful
        """
        return self._obj.Preview(fname)


class WorksheetPage(Page[oext_types.WorksheetPage]):
    """
    Workbook page containing worksheet layers.
    Wrapper class that wraps OriginExt.OriginExt.WorksheetPage.

    Corresponds to: originpro.WBook, OriginExt.OriginExt.WorksheetPage
    """

    def __init__(self, page: oext_types.WorksheetPage, parent: Optional['OriginObjectWrapper'] = None, 
                 origin_instance: Optional['OriginInstance'] = None):
        """
        Initialize WorksheetPage wrapper with hierarchical references.

        Args:
            page: Original OriginExt.WorksheetPage instance to wrap
            parent: Parent wrapper object (for hierarchical navigation)
            origin_instance: Root OriginInstance reference (for LabTalk access)
        """
        super().__init__(page, parent, origin_instance)

    def __iter__(self) -> Iterator[Worksheet]:
        """Iterate over worksheets"""
        from .layers import Worksheet
        for layer in self._obj:
            yield Worksheet(layer, self, self.origin_instance)

    def __getitem__(self, index: int) -> Worksheet:
        """Get worksheet by index"""
        from .layers import Worksheet
        return Worksheet(self._obj[index], self, self.origin_instance)

    def get_layers(self) -> list[Worksheet]:
        """
        Get list of worksheets in this workbook.

        Corresponds to: OriginExt.OriginExt.WorksheetPage.GetLayers()

        Returns:
            list[Worksheet]: List of worksheets
        """
        from .layers import Worksheet
        return [Worksheet(l, self, self.origin_instance) for l in self._obj.GetLayers()]

    def get_layer(self, index: int) -> Worksheet:
        """
        Get worksheet by index.

        Corresponds to: OriginExt.OriginExt.WorksheetPage.GetLayer()

        Args:
            index: Worksheet index

        Returns:
            Worksheet: The worksheet at the specified index
        """
        from .layers import Worksheet
        return Worksheet(self._obj.GetLayer(index), self, self.origin_instance)


class GraphPage(Page[oext_types.GraphPage]):
    """
    Graph page containing graph layers.
    Wrapper class that wraps OriginExt.OriginExt.GraphPage.

    Corresponds to: originpro.GPage, OriginExt.OriginExt.GraphPage
    """

    def __init__(self, page: oext_types.GraphPage, parent: Optional['OriginObjectWrapper'] = None, 
                 origin_instance: Optional['OriginInstance'] = None):
        """
        Initialize GraphPage wrapper with hierarchical references.

        Args:
            page: Original OriginExt.GraphPage instance to wrap
            parent: Parent wrapper object (for hierarchical navigation)
            origin_instance: Root OriginInstance reference (for LabTalk access)
        """
        super().__init__(page, parent, origin_instance)

    @property
    def BaseColor(self) -> int:
        """Base color of the graph page"""
        return self._obj.BaseColor

    @property
    def GradColor(self) -> int:
        """Gradient color of the graph page"""
        return self._obj.GradColor

    @property
    def Width(self) -> float:
        """Width of the graph page"""
        return self._obj.Width

    @property
    def Height(self) -> float:
        """Height of the graph page"""
        return self._obj.Height

    @property
    def Units(self) -> int:
        """Units for dimensions"""
        return self._obj.Units

    def __iter__(self) -> Iterator[GraphLayer]:
        """Iterate over graph layers"""
        from .layers import GraphLayer
        for layer in self._obj:
            yield GraphLayer(layer, self, self.origin_instance)

    def __getitem__(self, index: int) -> GraphLayer:
        """Get graph layer by index"""
        from .layers import GraphLayer
        return GraphLayer(self._obj[index], self, self.origin_instance)

    def get_layers(self) -> list[GraphLayer]:
        """
        Get list of graph layers in this page.

        Corresponds to: OriginExt.OriginExt.GraphPage.GetLayers()

        Returns:
            list[GraphLayer]: List of graph layers
        """
        from .layers import GraphLayer
        return [GraphLayer(l, self, self.origin_instance) for l in self._obj.GetLayers()]

    def get_layer(self, index: int) -> GraphLayer:
        """
        Get graph layer by index.

        Corresponds to: OriginExt.OriginExt.GraphPage.GetLayer()

        Args:
            index: Layer index

        Returns:
            GraphLayer: The graph layer at the specified index
        """
        from .layers import GraphLayer
        return GraphLayer(self._obj.GetLayer(index), self, self.origin_instance)

    def get_base_color(self) -> int:
        """
        Get base color.

        Corresponds to: OriginExt.OriginExt.GraphPage.GetBaseColor()

        Returns:
            int: Base color
        """
        return self._obj.GetBaseColor()

    def set_base_color(self, color: int) -> None:
        """
        Set base color.

        Corresponds to: OriginExt.OriginExt.GraphPage.SetBaseColor()

        Args:
            color: Base color
        """
        self._obj.SetBaseColor(color)

    def get_width(self) -> float:
        """
        Get page width.

        Corresponds to: OriginExt.OriginExt.GraphPage.GetWidth()

        Returns:
            float: Page width
        """
        return self._obj.GetWidth()

    def set_width(self, width: float) -> None:
        """
        Set page width.

        Corresponds to: OriginExt.OriginExt.GraphPage.SetWidth()

        Args:
            width: Page width
        """
        self._obj.SetWidth(width)

    def get_height(self) -> float:
        """
        Get page height.

        Corresponds to: OriginExt.OriginExt.GraphPage.GetHeight()

        Returns:
            float: Page height
        """
        return self._obj.GetHeight()

    def set_height(self, height: float) -> None:
        """
        Set page height.

        Corresponds to: OriginExt.OriginExt.GraphPage.SetHeight()

        Args:
            height: Page height
        """
        self._obj.SetHeight(height)


class MatrixPage(Page[oext_types.MatrixPage]):
    """
    Matrix book page containing matrix sheets.
    Wrapper class that wraps OriginExt.OriginExt.MatrixPage.

    Corresponds to: originpro.MBook, OriginExt.OriginExt.MatrixPage
    """

    def __init__(self, page: oext_types.MatrixPage):
        """
        Initialize MatrixPage wrapper.

        Args:
            page: Original OriginExt.MatrixPage instance to wrap
        """
        super().__init__(page)

    def __iter__(self) -> Iterator[Matrixsheet]:
        """Iterate over matrix sheets"""
        from .layers import Matrixsheet
        for layer in self._obj:
            yield Matrixsheet(layer)

    def __getitem__(self, index: int) -> Matrixsheet:
        """Get matrix sheet by index"""
        from .layers import Matrixsheet
        return Matrixsheet(self._obj[index])

    def get_layers(self) -> list[Matrixsheet]:
        """
        Get list of matrix sheets in this page.

        Corresponds to: OriginExt.OriginExt.MatrixPage.GetLayers()

        Returns:
            list[Matrixsheet]: List of matrix sheets
        """
        from .layers import Matrixsheet
        return [Matrixsheet(l) for l in self._obj.GetLayers()]

    def get_layer(self, index: int) -> Matrixsheet:
        """
        Get matrix sheet by index.

        Corresponds to: OriginExt.OriginExt.MatrixPage.GetLayer()

        Args:
            index: Matrix sheet index

        Returns:
            Matrixsheet: The matrix sheet at the specified index
        """
        from .layers import Matrixsheet
        return Matrixsheet(self._obj.GetLayer(index))


class NotePage(PageBase[oext_types.NotePage]):
    """
    Notes window page.
    Wrapper class that wraps OriginExt.OriginExt.NotePage.

    Corresponds to: originpro.Note, OriginExt.OriginExt.NotePage
    """

    def __init__(self, page: oext_types.NotePage):
        """
        Initialize NotePage wrapper.

        Args:
            page: Original OriginExt.NotePage instance to wrap
        """
        super().__init__(page)

    @property
    def Text(self) -> str:
        """Text content of the notes"""
        return self._obj.Text

    def get_text(self) -> str:
        """
        Get the text content.

        Corresponds to: OriginExt.OriginExt.NotePage.GetText()

        Returns:
            str: Text content
        """
        return self._obj.GetText()

    def set_text(self, text: str) -> None:
        """
        Set the text content.

        Corresponds to: OriginExt.OriginExt.NotePage.SetText()

        Args:
            text: Text content
        """
        self._obj.SetText(text)


class FigurePage(GraphPage):
    """
    Enhanced graph page with plotting functionality.
    Extends GraphPage with high-level plotting methods using DataPlot and GraphLayer.
    Provides object-oriented interface for creating and managing plots with enum-based controls.
    """

    def __init__(self, graph_page: GraphPage, template: str = 'Origin'):
        """
        Initialize FigurePage with hierarchical references.

        Args:
            graph_page: GraphPage object to wrap
            template: Template name for the graph
        """
        super().__init__(graph_page._obj, graph_page.parent, graph_page.origin_instance)
        self._template = template

    @classmethod
    def create_new(cls, name: str = '', template: str = 'scatter') -> FigurePage:
        """
        Create a new FigurePage with specified template.

        Args:
            name: Optional name for the new page
            template: Graph template name

        Returns:
            FigurePage: New FigurePage instance
        """
        # Import here to avoid circular imports
        import sys
        import os
        
        # Get the OriginInstance from the current context
        # This is a workaround - in practice, users should use origin.new_graph()
        # But for standalone creation, we need to access the global OriginExt
        
        # Try to get the current Origin instance
        try:
            # Check if we have an active Origin instance
            from . import OriginInstance
            if hasattr(OriginInstance, '_OriginInstance__instance_count') and OriginInstance._OriginInstance__instance_count > 0:
                # Use LabTalk through the existing instance
                import OriginExt.OriginExt as oext
                cmd = f'newpanel name:="{name}" template:="{template}"' if name else f'newpanel template:="{template}"'
                oext.LT_execute(cmd.strip())
                
                # Get the newly created graph page
                pages = list(oext.GetGraphPages())
                if pages:
                    return cls(pages[-1], template)
            
            raise RuntimeError("No active Origin instance found. Use OriginInstance.new_graph() instead.")
            
        except Exception as e:
            raise RuntimeError(f"Failed to create FigurePage: {e}. Please use OriginInstance.new_graph() method.")

    def get_active_layer(self) -> GraphLayer:
        """
        Get the active layer of the graph page.

        Returns:
            GraphLayer: The active graph layer
        """
        from .layers import GraphLayer
        import OriginExt._OriginExt as oext
        return GraphLayer(oext.GraphPage_GetLayer(self._obj, 0), self, self.origin_instance)

    def add_layer(self, name: str = '') -> GraphLayer:
        """
        Add a new graph layer to this page.

        Args:
            name: Optional name for the new layer

        Returns:
            GraphLayer: The newly created layer
        """
        from .layers import GraphLayer
        return GraphLayer(self._obj.AddLayer(name), self, self.origin_instance)

    def plot_xy_data(self, worksheet, x_col: int, y_col: int = -1,
                    plot_type = None, layer_index: int = 0,
                    color_map = None, shape_list: list[int] = None,
                    group_mode = None) -> DataPlot:
        """
        Plot XY data from worksheet with hierarchical structure.

        Args:
            worksheet: Worksheet containing data
            x_col: X column index (0-based)
            y_col: Y column index (0-based) or -1 for all columns after x_col
            plot_type: PlotType enum (defaults to LINE_SYMBOL)
            layer_index: Layer index to plot on (0 for first layer)
            color_map: Optional ColorMap enum
            shape_list: Optional list of shape indices
            group_mode: GroupMode enum for plot grouping

        Returns:
            DataPlot: The created data plot
        """
        from .layers import PlotType, GroupMode
        
        if plot_type is None:
            plot_type = PlotType.LINE_SYMBOL
        if group_mode is None:
            group_mode = GroupMode.DEPENDENT

        # Get or create the target layer
        if layer_index >= len(self):
            # Add new layers if needed
            while layer_index >= len(self):
                self.add_layer()
        
        layer = self[layer_index]

        # Add the plot
        plot = layer.add_xy_plot(worksheet, x_col, y_col, plot_type)

        # Apply optional styling
        if color_map:
            plot.color_map = color_map
        
        if shape_list:
            plot.shape_list = shape_list

        # Apply grouping if requested
        if group_mode != GroupMode.NONE:
            layer.group_plots(group_mode)

        # Rescale the layer
        layer.rescale()

        return plot

    def plot_multiple_series(self, worksheet, x_col: int, y_cols: list[int],
                           plot_type = None, layer_index: int = 0,
                           color_map = None, shape_list: list[int] = None,
                           group_mode = None) -> list[DataPlot]:
        """
        Plot multiple Y series against the same X column.

        Args:
            worksheet: Worksheet containing data
            x_col: X column index (0-based)
            y_cols: List of Y column indices (0-based)
            plot_type: PlotType enum (defaults to LINE_SYMBOL)
            layer_index: Layer index to plot on
            color_map: Optional ColorMap enum
            shape_list: Optional list of shape indices
            group_mode: GroupMode enum for plot grouping

        Returns:
            list[DataPlot]: List of created data plots
        """
        from .layers import GroupMode
        
        plots = []
        for y_col in y_cols:
            plot = self.plot_xy_data(worksheet, x_col, y_col, plot_type, 
                                   layer_index, color_map, shape_list, GroupMode.NONE)
            plots.append(plot)

        # Apply grouping after all plots are created
        if group_mode and group_mode != GroupMode.NONE:
            layer = self[layer_index]
            layer.group_plots(group_mode)
            layer.rescale()

        return plots

    def create_grouped_plot(self, worksheet, x_col: int, 
                          color_map = None,
                          shape_list: list[int] = None,
                          plot_type = None) -> DataPlot:
        """
        Create a grouped plot similar to the Origin example (Sample #5).

        Args:
            worksheet: Worksheet containing data
            x_col: X column index (0-based)
            color_map: ColorMap enum for coloring
            shape_list: List of shape indices for different series
            plot_type: PlotType enum

        Returns:
            DataPlot: The created grouped plot
        """
        from .layers import PlotType, ColorMap, GroupMode
        
        if plot_type is None:
            plot_type = PlotType.LINE_SYMBOL
        if color_map is None:
            color_map = ColorMap.CANDY

        # Plot whole sheet as XY plot (all columns after x_col as Y)
        plot = self.plot_xy_data(worksheet, x_col, -1, plot_type, 0, color_map, shape_list)

        # Get the layer and apply grouping
        layer = self.get_active_layer()
        layer.group_plots(GroupMode.DEPENDENT)

        return plot

    def set_page_size(self, width: float, height: float, units: int = 0) -> None:
        """
        Set the page size.

        Args:
            width: Page width
            height: Page height
            units: Units (0=inch, 1=cm, 2=mm, 3=pixel, 4=point)
        """
        self.set_width(width)
        self.set_height(height)
        self._obj.Units = units

    def set_colors(self, base_color: int = None, grad_color: int = None) -> None:
        """
        Set page colors.

        Args:
            base_color: Base color index
            grad_color: Gradient color index
        """
        if base_color is not None:
            self.set_base_color(base_color)
        if grad_color is not None:
            self._obj.SetGradColor(grad_color)

    def export_preview(self, filename: str) -> bool:
        """
        Export preview image of the page.

        Args:
            filename: Output filename

        Returns:
            bool: True if successful
        """
        return self.preview(filename)
