"""
Layer classes for OriginExt wrappers.

This module contains wrapper classes for Origin layers including:
- Layer (base class)
- Datasheet
- Worksheet
- GraphLayer
- Matrixsheet
- Column
- DataPlot
"""
from __future__ import annotations

import OriginExt.OriginExt as oext_types

from typing import Iterator, TypeVar, TYPE_CHECKING

from .base import OriginObjectWrapper

if TYPE_CHECKING:
    from .pages import Page, WorksheetPage, GraphPage, MatrixPage


# ================== Type Variables ==================

TLayer = TypeVar('TLayer', bound=oext_types.Layer)
TDatasheet = TypeVar('TDatasheet', bound=oext_types.Datasheet)


# ================== Layer Classes ==================

class Layer(OriginObjectWrapper[TLayer]):
    """
    Base class for layers (worksheets, graph layers, matrix sheets).
    Wrapper class that wraps OriginExt.OriginExt.Layer.
    Inherits common methods from OriginObjectWrapper.

    Corresponds to: originpro.DSheet / originpro.GLayer, OriginExt.OriginExt.Layer
    """

    def __init__(self, layer: TLayer):
        """
        Initialize Layer wrapper.

        Args:
            layer: Original OriginExt.Layer instance to wrap
        """
        super().__init__(layer)

    @property
    def Parent(self) -> Page:
        """Parent page of this layer"""
        from .pages import Page
        return Page(self._obj.GetPage())

    def get_page(self) -> Page:
        """
        Get the parent page of this layer.

        Corresponds to: OriginExt.OriginExt.Layer.GetPage()

        Returns:
            Page: Parent page
        """
        from .pages import Page
        return Page(self._obj.GetPage())

    def get_data_object_bases(self):
        """
        Get collection of data objects in this layer.

        Corresponds to: OriginExt.OriginExt.Layer.GetDataObjectBases()

        Returns:
            Collection of data objects
        """
        return self._obj.GetDataObjectBases()


class Datasheet(Layer[TDatasheet]):
    """
    Base class for data sheets (Worksheet, Matrixsheet).
    Wrapper class that wraps OriginExt.OriginExt.Datasheet.

    Corresponds to: OriginExt.OriginExt.Datasheet
    """

    def __init__(self, datasheet: TDatasheet):
        """
        Initialize Datasheet wrapper.

        Args:
            datasheet: Original OriginExt.Datasheet instance to wrap
        """
        super().__init__(datasheet)

    @property
    def Cols(self) -> int:
        """Number of columns"""
        return self._obj.Cols

    @property
    def Rows(self) -> int:
        """Number of rows"""
        return self._obj.Rows

    def get_cols(self) -> int:
        """
        Get number of columns.

        Corresponds to: OriginExt.OriginExt.Datasheet.GetCols()

        Returns:
            int: Number of columns
        """
        return self._obj.GetCols()

    def set_cols(self, cols: int) -> None:
        """
        Set number of columns.

        Corresponds to: OriginExt.OriginExt.Datasheet.SetCols()

        Args:
            cols: Number of columns
        """
        self._obj.SetCols(cols)

    def get_rows(self) -> int:
        """
        Get number of rows.

        Corresponds to: OriginExt.OriginExt.Datasheet.GetRows()

        Returns:
            int: Number of rows
        """
        return self._obj.GetRows()

    def set_rows(self, rows: int) -> None:
        """
        Set number of rows.

        Corresponds to: OriginExt.OriginExt.Datasheet.SetRows()

        Args:
            rows: Number of rows
        """
        self._obj.SetRows(rows)

    def clear_data(self, col_start: int = 0, col_end: int = -1) -> None:
        """
        Clear data in specified column range.

        Corresponds to: OriginExt.OriginExt.Datasheet.ClearData()

        Args:
            col_start: Starting column index
            col_end: Ending column index (-1 for all)
        """
        self._obj.ClearData(col_start, col_end)

    def find_col(self, label: str, begin: int = 0, case_sensitive: bool = False,
                 full_match: bool = True, end: int = -1, allow_short_name: bool = True) -> int:
        """
        Find column by label.

        Corresponds to: OriginExt.OriginExt.Datasheet.FindCol()

        Args:
            label: Label to search for
            begin: Starting column index
            case_sensitive: Case sensitive search
            full_match: Require full match
            end: Ending column index
            allow_short_name: Allow short name matching

        Returns:
            int: Column index or -1 if not found
        """
        return self._obj.FindCol(label, begin, case_sensitive, full_match, end, allow_short_name)


class Column:
    """
    Column in a worksheet.
    Wrapper class that wraps OriginExt.OriginExt.Column.

    Corresponds to: OriginExt.OriginExt.Column
    """

    _column: oext_types.Column

    def __init__(self, column: oext_types.Column):
        """
        Initialize Column wrapper.

        Args:
            column: Original OriginExt.Column instance to wrap
        """
        self._column = column

    @property
    def Name(self) -> str:
        """Short name of the column"""
        return self._column.Name

    @property
    def LongName(self) -> str:
        """Long name of the column"""
        return self._column.LongName

    @property
    def Type(self) -> int:
        """Column type"""
        return self._column.Type

    @property
    def Units(self) -> str:
        """Column units"""
        return self._column.Units

    @property
    def Comments(self) -> str:
        """Column comments"""
        return self._column.Comments

    @property
    def Parent(self) -> Worksheet:
        """Parent worksheet"""
        return Worksheet(self._column.Parent)

    def get_parent(self) -> Worksheet:
        """
        Get the parent worksheet.

        Corresponds to: OriginExt.OriginExt.Column.GetParent()

        Returns:
            Worksheet: Parent worksheet
        """
        return Worksheet(self._column.GetParent())

    def get_data(self, format: int, start: int = 0, end: int = -1, lowbound: int = 1):
        """
        Get column data.

        Corresponds to: OriginExt.OriginExt.Column.GetData()

        Args:
            format: Data format
            start: Starting row
            end: Ending row
            lowbound: Lower bound

        Returns:
            Column data
        """
        return self._column.GetData(format, start, end, lowbound)

    def set_data(self, *args):
        """
        Set column data.

        Corresponds to: OriginExt.OriginExt.Column.SetData()
        """
        return self._column.SetData(*args)

    def is_valid(self) -> bool:
        """
        Check if this column is valid.

        Corresponds to: OriginExt.OriginExt.Column.IsValid()

        Returns:
            bool: True if valid
        """
        return self._column.IsValid()


class Worksheet(Datasheet[oext_types.Worksheet]):
    """
    Worksheet layer for tabular data.
    Wrapper class that wraps OriginExt.OriginExt.Worksheet.

    Corresponds to: originpro.WSheet, OriginExt.OriginExt.Worksheet
    """

    def __init__(self, worksheet: oext_types.Worksheet):
        """
        Initialize Worksheet wrapper.

        Args:
            worksheet: Original OriginExt.Worksheet instance to wrap
        """
        super().__init__(worksheet)

    @property
    def Columns(self):
        """Collection of columns in this worksheet"""
        return self._obj.Columns

    def __iter__(self) -> Iterator[Column]:
        """Iterate over columns"""
        return iter(self._obj)

    def __getitem__(self, index: int) -> Column:
        """Get column by index"""
        return Column(self._obj[index])

    def get_cell(self, row: int, col: int):
        """
        Get cell value at specified row and column.

        Corresponds to: OriginExt.OriginExt.Worksheet.GetCell()

        Args:
            row: Row index
            col: Column index

        Returns:
            Cell value
        """
        return self._obj.GetCell(row, col)

    def get_columns(self) -> list[Column]:
        """
        Get list of columns in this worksheet.

        Corresponds to: OriginExt.OriginExt.Worksheet.GetColumns()

        Returns:
            list[Column]: List of columns
        """
        return [Column(c) for c in self._obj.GetColumns()]

    def get_page(self) -> WorksheetPage:
        """
        Get the parent workbook page.

        Corresponds to: OriginExt.OriginExt.Worksheet.GetPage()

        Returns:
            WorksheetPage: Parent workbook page
        """
        from .pages import WorksheetPage
        return WorksheetPage(self._obj.GetPage())

    def get_data(self, row_start: int = 0, col_start: int = 0, row_end: int = -1, col_end: int = -1, format: int = 0):
        """
        Get worksheet data.

        Corresponds to: OriginExt.OriginExt.Worksheet.GetData()

        Args:
            row_start: Starting row
            col_start: Starting column
            row_end: Ending row
            col_end: Ending column
            format: Data format

        Returns:
            Worksheet data
        """
        return self._obj.GetData(row_start, col_start, row_end, col_end, format)

    def set_data(self, *args):
        """
        Set worksheet data.

        Corresponds to: OriginExt.OriginExt.Worksheet.SetData()
        """
        return self._obj.SetData(*args)


class DataPlot:
    """
    Data plot in a graph layer.
    Wrapper class that wraps OriginExt.OriginExt.DataPlot.

    Corresponds to: OriginExt.OriginExt.DataPlot
    """

    _plot: oext_types.DataPlot

    def __init__(self, plot: oext_types.DataPlot):
        """
        Initialize DataPlot wrapper.

        Args:
            plot: Original OriginExt.DataPlot instance to wrap
        """
        self._plot = plot

    @property
    def Name(self) -> str:
        """Short name of the data plot"""
        return self._plot.Name

    @property
    def Parent(self) -> GraphLayer:
        """Parent graph layer"""
        return GraphLayer(self._plot.Parent)

    def get_parent(self) -> GraphLayer:
        """
        Get the parent graph layer.

        Corresponds to: OriginExt.OriginExt.DataPlot.GetParent()

        Returns:
            GraphLayer: Parent graph layer
        """
        return GraphLayer(self._plot.GetParent())

    def get_color_map(self):
        """
        Get color map.

        Corresponds to: OriginExt.OriginExt.DataPlot.GetColorMap()

        Returns:
            Color map
        """
        return self._plot.GetColorMap()

    def change_data(self, data_obj, designation: str, keep_modifiers: bool = False):
        """
        Change data source.

        Corresponds to: OriginExt.OriginExt.DataPlot.ChangeData()

        Args:
            data_obj: Data object
            designation: Designation string
            keep_modifiers: Keep modifiers
        """
        return self._plot.ChangeData(data_obj, designation, keep_modifiers)


class GraphLayer(Layer[oext_types.GraphLayer]):
    """
    Graph layer for plotting data.
    Wrapper class that wraps OriginExt.OriginExt.GraphLayer.

    Corresponds to: originpro.GLayer, OriginExt.OriginExt.GraphLayer
    """

    def __init__(self, layer: oext_types.GraphLayer):
        """
        Initialize GraphLayer wrapper.

        Args:
            layer: Original OriginExt.GraphLayer instance to wrap
        """
        super().__init__(layer)

    @property
    def DataPlots(self):
        """Collection of data plots in this layer"""
        return self._obj.DataPlots

    @property
    def GraphObjects(self):
        """Collection of graph objects in this layer"""
        return self._obj.GraphObjects

    def __iter__(self) -> Iterator[DataPlot]:
        """Iterate over data plots"""
        for plot in self._obj:
            yield DataPlot(plot)

    def __getitem__(self, index: int) -> DataPlot:
        """Get data plot by index"""
        return DataPlot(self._obj[index])

    def add_plot(self, data_range, plot_type: int, composite: bool = False) -> DataPlot:
        """
        Add a plot from a data range.

        Corresponds to: OriginExt.OriginExt.GraphLayer.AddPlot()

        Args:
            data_range: Data range object
            plot_type: Plot type identifier
            composite: Create composite plot

        Returns:
            DataPlot: The newly created data plot
        """
        return DataPlot(self._obj.AddPlot(data_range, plot_type, composite))

    def add_plot_from_string(self, range_str: str, plot_type: int) -> DataPlot:
        """
        Add a plot from a range string.

        Corresponds to: OriginExt.OriginExt.GraphLayer.AddPlotFromString()

        Args:
            range_str: Range string
            plot_type: Plot type identifier

        Returns:
            DataPlot: The newly created data plot
        """
        return DataPlot(self._obj.AddPlotFromString(range_str, plot_type))

    def get_data_plots(self) -> list[DataPlot]:
        """
        Get list of data plots in this layer.

        Corresponds to: OriginExt.OriginExt.GraphLayer.GetDataPlots()

        Returns:
            list[DataPlot]: List of data plots
        """
        return [DataPlot(p) for p in self._obj.GetDataPlots()]

    def get_graph_objects(self):
        """
        Get graph objects in this layer.

        Corresponds to: OriginExt.OriginExt.GraphLayer.GetGraphObjects()

        Returns:
            Graph objects collection
        """
        return self._obj.GetGraphObjects()

    def get_page(self) -> GraphPage:
        """
        Get the parent graph page.

        Corresponds to: OriginExt.OriginExt.GraphLayer.GetPage()

        Returns:
            GraphPage: Parent graph page
        """
        from .pages import GraphPage
        return GraphPage(self._obj.GetPage())


class Matrixsheet(Datasheet[oext_types.Matrixsheet]):
    """
    Matrix sheet layer for 2D array data.
    Wrapper class that wraps OriginExt.OriginExt.Matrixsheet.

    Corresponds to: originpro.MSheet, OriginExt.OriginExt.Matrixsheet
    """

    def __init__(self, matrixsheet: oext_types.Matrixsheet):
        """
        Initialize Matrixsheet wrapper.

        Args:
            matrixsheet: Original OriginExt.Matrixsheet instance to wrap
        """
        super().__init__(matrixsheet)

    @property
    def MatrixObjects(self):
        """Collection of matrix objects in this sheet"""
        return self._obj.MatrixObjects

    def get_matrix_objects(self):
        """
        Get matrix objects in this sheet.

        Corresponds to: OriginExt.OriginExt.Matrixsheet.GetMatrixObjects()

        Returns:
            Matrix objects collection
        """
        return self._obj.GetMatrixObjects()

    def get_page(self) -> MatrixPage:
        """
        Get the parent matrix book page.

        Corresponds to: OriginExt.OriginExt.Matrixsheet.GetPage()

        Returns:
            MatrixPage: Parent matrix book page
        """
        from .pages import MatrixPage
        return MatrixPage(self._obj.GetPage())

    def set_shape(self, rows: int, cols: int, keep_data: bool = False) -> None:
        """
        Set matrix shape.

        Corresponds to: OriginExt.OriginExt.Matrixsheet.SetShape()

        Args:
            rows: Number of rows
            cols: Number of columns
            keep_data: Keep existing data
        """
        self._obj.SetShape(rows, cols, keep_data)
