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
import OriginExt._OriginExt as oext
import pandas as pd
from enum import Enum

from typing import Iterator, TypeVar, TYPE_CHECKING, Union, Optional

from .base import OriginObjectWrapper

if TYPE_CHECKING:
    from .pages import Page, WorkbookPage, GraphPage, MatrixPage


# ================== Helper Functions ==================

def get_originext_obj(wrapper_obj):
    """
    Get the underlying OriginExt object from a wrapper.
    
    Args:
        wrapper_obj: A wrapper object (GraphLayer, Worksheet, etc.)
        
    Returns:
        The underlying OriginExt object
    """
    if hasattr(wrapper_obj, '_obj'):
        # If _obj is another wrapper, get its _obj
        if hasattr(wrapper_obj._obj, '_obj'):
            return wrapper_obj._obj._obj
        return wrapper_obj._obj
    return wrapper_obj

def get_originext_graphlayer(graphlayer_wrapper):
    """
    Get the OriginExt GraphLayer object from a GraphLayer wrapper.
    
    Args:
        graphlayer_wrapper: GraphLayer wrapper instance
        
    Returns:
        OriginExt GraphLayer object
    """
    obj = get_originext_obj(graphlayer_wrapper)
    # Ensure it's the correct type by checking if it has GraphLayer methods
    return obj

def get_originext_worksheet(worksheet_wrapper):
    """
    Get the OriginExt Worksheet object from a Worksheet wrapper.
    
    Args:
        worksheet_wrapper: Worksheet wrapper instance
        
    Returns:
        OriginExt Worksheet object
    """
    obj = get_originext_obj(worksheet_wrapper)
    return obj


# ================== Type Variables ==================

TLayer = TypeVar('TLayer', bound=oext_types.Layer)
TDatasheet = TypeVar('TDatasheet', bound=oext_types.Datasheet)


# ================== Enum Types ==================

class PlotType(Enum):
    """Enumeration for plot types to avoid literal values."""
    LINE = 200  # Line plot
    SCATTER = 201  # Scatter plot
    LINE_SYMBOL = 202  # Line + Symbol plot
    COLUMN = 203  # Column/bar plot
    BAR = 204  # Horizontal bar plot
    AREA = 205  # Area plot
    PIE = 206  # Pie chart
    SURFACE = 207  # 3D surface plot
    CONTOUR = 208  # Contour plot
    HISTOGRAM = 209  # Histogram


class XYTemplate(Enum):
    """Enumeration for XY plot templates (input data range = XY)."""
    LINE = "line"  # Line plot
    SCATTER = "scatter"  # Scatter plot
    LINE_SYMBOL = "linesymb"  # Line + Symbol plot
    COLUMN = "column"  # Column plot
    BAR = "bar"  # Bar plot
    AREA = "area"  # Area plot
    STEP = "step"  # Step plot
    SPLINE = "spline"  # Spline curve
    DROP_LINE = "dropline"  # Drop line plot
    FLOATING_BAR = "floatbar"  # Floating bar plot
    FLOATING_COLUMN = "floatcol"  # Floating column plot
    HIGH_LOW = "hilo"  # High-Low plot
    HIGH_LOW_CLOSE = "hiloclose"  # High-Low-Close plot
    OHLC = "ohlc"  # Open-High-Low-Close plot
    BOX_CHART = "box"  # Box chart
    HISTOGRAM = "histogram"  # Histogram
    HISTOGRAM_PLUS = "histogram+"  # Histogram + distribution
    BINS_2D = "2dbins"  # 2D Binning
    DENSITY = "density"  # Density plot
    VIOLIN = "violin"  # Violin plot
    PROFILE = "p_profile"  # Profile plot
    ZONES = "zones"  # Zones plot
    PIE = "pie"  # Pie chart
    DONUT = "donut"  # Donut chart
    PLOT_2D = "2d"  # 2D plot
    POLAR = "polar"  # Polar plot
    TERNARY = "ternary"  # Ternary plot
    SMITH = "smith"  # Smith chart
    VECTOR = "vector"  # Vector plot
    CONTOUR = "contour"  # Contour plot
    HEAT_MAP = "heatmap"  # Heat map
    IMAGE = "image"  # Image plot


class ColorMap(Enum):
    """Enumeration for color maps to avoid literal values."""
    CANDY = "Candy"
    RAINBOW = "Rainbow"
    HEATMAP = "Heatmap"
    GRAYSCALE = "Grayscale"
    OCEAN = "Ocean"
    TERRAIN = "Terrain"
    VIRIDIS = "Viridis"
    PLASMA = "Plasma"


class GroupMode(Enum):
    """Enumeration for grouping modes."""
    NONE = 0  # No grouping
    INDEPENDENT = 1  # Independent plots
    DEPENDENT = 2  # Dependent plots


class AxisType(Enum):
    """Enumeration for axis types."""
    X = 1
    Y = 2
    Z = 3
    ERROR = 4


# ================== Layer Classes ==================

class Layer(OriginObjectWrapper[TLayer]):
    """
    Base class for layers (worksheets, graph layers, matrix sheets).
    Wrapper class that wraps OriginExt.OriginExt.Layer.
    Inherits common methods from OriginObjectWrapper.

    Corresponds to: originpro.DSheet / originpro.GLayer, OriginExt.OriginExt.Layer
    """

    def __init__(self, layer: TLayer, parent: Optional['OriginObjectWrapper'] = None, 
                 origin_instance: Optional['OriginInstance'] = None):
        """
        Initialize Layer wrapper with hierarchical references.

        Args:
            layer: Original OriginExt.Layer instance to wrap
            parent: Parent wrapper object (for hierarchical navigation)
            origin_instance: Root OriginInstance reference (for LabTalk access)
        """
        super().__init__(layer, parent, origin_instance)

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

    def __init__(self, datasheet: TDatasheet, parent: Optional['OriginObjectWrapper'] = None, 
                 origin_instance: Optional['OriginInstance'] = None):
        """
        Initialize Datasheet wrapper with hierarchical references.

        Args:
            datasheet: Original OriginExt.Datasheet instance to wrap
            parent: Parent wrapper object (for hierarchical navigation)
            origin_instance: Root OriginInstance reference (for LabTalk access)
        """
        super().__init__(datasheet, parent, origin_instance)

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


class ColumnCollection:
    """
    Wrapper for Origin column collection that returns wrapped Column objects.
    """
    
    def __init__(self, columns_collection, parent: Optional['OriginObjectWrapper'] = None):
        self._columns = columns_collection
        self._parent = parent
    
    def __getitem__(self, index: int) -> Column:
        """Get column by index"""
        return Column(self._columns[index], self._parent, self._parent.origin_instance if self._parent else None)
    
    def __call__(self, index: int) -> Column:
        """Get column by index (Origin-style access)"""
        return Column(self._columns(index), self._parent, self._parent.origin_instance if self._parent else None)
    
    def __len__(self) -> int:
        """Get number of columns"""
        return len(self._columns)
    
    def __iter__(self):
        """Iterate over columns, yielding wrapped Column objects"""
        for col in self._columns:
            yield Column(col, self._parent, self._parent.origin_instance if self._parent else None)


class Column(OriginObjectWrapper[oext_types.Column]):
    """
    Column in a worksheet.
    Wrapper class that wraps OriginExt.OriginExt.Column.

    Corresponds to: originpro.Column, OriginExt.OriginExt.Column
    """

    def __init__(self, column: oext_types.Column, parent: Optional['OriginObjectWrapper'] = None, 
                 origin_instance: Optional['OriginInstance'] = None):
        """
        Initialize Column wrapper with hierarchical references.

        Args:
            column: Original OriginExt.Column instance to wrap
            parent: Parent wrapper object (for hierarchical navigation)
            origin_instance: Root OriginInstance reference (for LabTalk access)
        """
        super().__init__(column, parent, origin_instance)

    @property
    def name(self) -> str:
        """Short name of the column"""
        return self._obj.Name

    @name.setter
    def name(self, value: str) -> None:
        """Set short name"""
        self._obj.Name = value

    # Backward compatibility alias
    @property
    def Name(self) -> str:
        """Short name of the column (backward compatibility)"""
        return self.name

    @Name.setter
    def Name(self, value: str) -> None:
        """Set short name (backward compatibility)"""
        self.name = value

    @property
    def long_name(self) -> str:
        """Long name of the column"""
        return self._obj.LongName

    @long_name.setter
    def long_name(self, value: str) -> None:
        """Set long name"""
        self._obj.LongName = value

    # Backward compatibility alias
    @property
    def LongName(self) -> str:
        """Long name of the column (backward compatibility)"""
        return self.long_name

    @LongName.setter
    def LongName(self, value: str) -> None:
        """Set long name (backward compatibility)"""
        self.long_name = value

    @property
    def type(self) -> int:
        """Column type"""
        return self._obj.Type

    # Backward compatibility alias
    @property
    def Type(self) -> int:
        """Column type (backward compatibility)"""
        return self.type

    @Type.setter
    def Type(self, value: int) -> None:
        """Set column type (backward compatibility)"""
        self._obj.Type = value

    @property
    def units(self) -> str:
        """Column units"""
        return self._obj.Units

    @units.setter
    def units(self, value: str) -> None:
        """Set column units"""
        self._obj.Units = value

    # Backward compatibility alias
    @property
    def Units(self) -> str:
        """Column units (backward compatibility)"""
        return self.units

    @Units.setter
    def Units(self, value: str) -> None:
        """Set column units (backward compatibility)"""
        self.units = value

    @property
    def comments(self) -> str:
        """Column comments"""
        return self._obj.Comments

    @comments.setter
    def comments(self, value: str) -> None:
        """Set column comments"""
        self._obj.Comments = value

    # Backward compatibility alias
    @property
    def Comments(self) -> str:
        """Column comments (backward compatibility)"""
        return self.comments

    @Comments.setter
    def Comments(self, value: str) -> None:
        """Set column comments (backward compatibility)"""
        self.comments = value

    @property
    def parent(self) -> Worksheet:
        """Parent worksheet"""
        return Worksheet(self._obj.Parent, self._parent, self.origin_instance)

    # Backward compatibility alias
    @property
    def Parent(self) -> Worksheet:
        """Parent worksheet (backward compatibility)"""
        return self.parent

    def get_parent(self) -> Worksheet:
        """
        Get the parent worksheet.

        Corresponds to: OriginExt.OriginExt.Column.GetParent()

        Returns:
            Worksheet: Parent worksheet
        """
        return Worksheet(self._obj.GetParent(), self._parent, self.origin_instance)

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
        return self._obj.GetData(format, start, end, lowbound)

    def set_data(self, *args):
        """
        Set column data.

        Corresponds to: OriginExt.OriginExt.Column.SetData()
        """
        return self._obj.SetData(*args)

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

    def __init__(self, worksheet: oext_types.Worksheet, parent: Optional['OriginObjectWrapper'] = None, 
                 origin_instance: Optional['OriginInstance'] = None):
        """
        Initialize Worksheet wrapper with hierarchical references.

        Args:
            worksheet: Original OriginExt.Worksheet instance to wrap
            parent: Parent wrapper object (for hierarchical navigation)
            origin_instance: Root OriginInstance reference (for LabTalk access)
        """
        super().__init__(worksheet, parent, origin_instance)
        
        # Automatically set header rows to show: Long Name, Units, Sparklines, F(x), Comments
        self.header_rows('LUSCO')
        
        # Ensure sparklines are properly activated and generated
        self._ensure_sparklines()

    @property
    def Columns(self):
        """Collection of columns in this worksheet"""
        # Return a wrapper that provides access to wrapped Column objects
        return ColumnCollection(self._obj.Columns, self)

    def __iter__(self) -> Iterator[Column]:
        """Iterate over columns"""
        for col in self._obj:
            yield Column(col, self, self.origin_instance)

    def __getitem__(self, index: int) -> Column:
        """Get column by index"""
        return Column(self._obj[index], self, self.origin_instance)

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
        return [Column(c, self, self.origin_instance) for c in self._obj.GetColumns()]

    def get_page(self) -> WorkbookPage:
        """
        Get the parent workbook page.

        Returns:
            WorkbookPage: Parent workbook page
        """
        from .pages import WorkbookPage
        return WorkbookPage(oext.Worksheet_GetPage(self._obj))

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

    def from_df(self, df: pd.DataFrame) -> None:
        """
        Load pandas DataFrame into worksheet.
        Uses DataFrame column names as Origin column Long Name.
        Attempts to generate sparklines for all columns after loading data.

        Args:
            df: pandas DataFrame to load into worksheet
        """
        if not isinstance(df, pd.DataFrame):
            raise TypeError("Input must be a pandas DataFrame")
        
        # Ensure worksheet has enough columns
        num_cols = len(df.columns)
        if self.get_cols() < num_cols:
            self.set_cols(num_cols)
        
        # Ensure worksheet has enough rows
        num_rows = len(df)
        if self.get_rows() < num_rows:
            self.set_rows(num_rows)
        
        # Set column long names from DataFrame column names
        for i, col_name in enumerate(df.columns):
            col = self.Columns(i)
            col.LongName = str(col_name)
        
        # Load data column by column
        for i, col_name in enumerate(df.columns):
            col_data = df[col_name].tolist()
            col = self.Columns(i)
            col.set_data(col_data)
        
        # Attempt to generate sparklines for all columns after data is loaded
        try:
            self.generate_sparklines()
        except Exception as e:
            print(f"Warning: Could not automatically generate sparklines: {e}")

    def from_list(self, col_index: Union[int, str], data: list, lname: Optional[str] = None, 
                  units: Optional[str] = None, comments: Optional[str] = None, 
                  axis: Optional[str] = None) -> None:
        """
        Load list data into specified column with optional properties.
        
        Args:
            col_index: Column index (0-based) or column letter ('A', 'B', etc.)
            data: List of data to load
            lname: Optional long name for the column
            units: Optional units for the column
            comments: Optional comments for the column
            axis: Optional axis designation ('X', 'Y', 'Z', 'E', etc.)
        """
        # Convert column letter to index if necessary
        if isinstance(col_index, str):
            if len(col_index) == 1 and col_index.isalpha():
                col_index = ord(col_index.upper()) - ord('A')
            else:
                raise ValueError("Column letter must be a single alphabetic character (A, B, C, etc.)")
        
        # Ensure worksheet has enough columns
        if self.get_cols() <= col_index:
            self.set_cols(col_index + 1)
        
        # Ensure worksheet has enough rows
        if self.get_rows() < len(data):
            self.set_rows(len(data))
        
        # Get the column and set data
        col = self.Columns(col_index)
        col.set_data(data)
        
        # Set optional properties
        if lname is not None:
            col.LongName = lname
        if units is not None:
            col.Units = units
        if comments is not None:
            col.Comments = comments
        if axis is not None:
            # Set axis designation using LabTalk command
            # This corresponds to setting the column type/designation
            axis_map = {'X': 1, 'Y': 2, 'Z': 3, 'E': 4}  # Common axis types
            if axis.upper() in axis_map:
                col.Type = axis_map[axis.upper()]
            else:
                # Try to set using LabTalk command for other designations
                range_str = f"[{self.get_page().Name}]${self.Name}!{col_index + 1}"
                self._obj.Execute(f"set {range_str} -t {axis}")

    def header_rows(self, spec: str = '') -> None:
        """
        Controls which worksheet label rows to show, same as LabTalk wks.labels string.
        
        Args:
            spec: A combination of letters for column label rows to display.
                  Common options:
                  - 'L': Long Name
                  - 'U': Units  
                  - 'C': Comments
                  - 'G': Short Name
                  - 'S': Sparklines (attempted for all columns)
                  - 'O': F(x)= (Formula)
                  - 'LU': Long Name + Units
                  - 'LUC': Long Name + Units + Comments
                  - 'LUSCO': Long Name + Units + Sparklines + F(x) + Comments
                  - '': Hide all label rows, keep only heading
                  See: https://www.originlab.com/doc/LabTalk/ref/Column-Label-Row-Characters
                  
        Examples:
            worksheet.header_rows('L')     # Show only Long Name
            worksheet.header_rows('LU')    # Show Long Name and Units
            worksheet.header_rows('LUC')   # Show Long Name, Units, and Comments
            worksheet.header_rows('LUSCO') # Show Long Name, Units, Sparklines, F(x), Comments
            worksheet.header_rows()        # Hide all label rows
            
        Note: When 'S' (Sparklines) is included, sparklines generation will be attempted
        for ALL columns regardless of data type. Warning messages will be displayed for
        columns where sparklines cannot be generated, but the process will continue.
        """
        if len(spec) == 0:
            specs = '0'  # Hide all label rows
        else:
            specs = spec.upper()
        
        # Use the Labels method (OriginExt equivalent of ShowLabels)
        try:
            self._obj.Labels(specs)
        except Exception as e:
            # Fallback: try using LabTalk command
            try:
                self._obj.Execute(f"wks.labels {specs}")
            except Exception as e2:
                print(f"Failed to set header rows: {e}, LabTalk fallback also failed: {e2}")

    def _ensure_sparklines(self) -> None:
        """
        Ensure sparklines are properly activated for this worksheet.
        This method activates sparklines and attempts to generate them for all columns.
        """
        try:
            # Use LabTalk command to ensure sparklines are available
            self._obj.Execute("wks.labels *S")  # Add sparklines row if not present
            
            # Try to generate sparklines for all columns
            self.generate_sparklines()
            
        except Exception as e:
            print(f"Warning: Could not ensure sparklines: {e}")

    def generate_sparklines(self, start_col: int = 0, end_col: int = -1) -> None:
        """
        Generate sparklines for all columns in the worksheet.
        Attempts to create sparklines for every column, regardless of data type.
        If sparklines cannot be generated for a column, displays a warning and continues.
        
        Args:
            start_col: Starting column index (0-based)
            end_col: Ending column index (-1 for all columns)
        """
        try:
            # Get the actual number of columns
            num_cols = self.get_cols()
            if end_col == -1 or end_col >= num_cols:
                end_col = num_cols - 1
            
            #print(f"Attempting to generate sparklines for columns {start_col} to {end_col}...")
            
            # Generate sparklines using LabTalk sparklines X-function for ALL columns
            for col_idx in range(start_col, end_col + 1):
                try:
                    # Get column name/letter for LabTalk command
                    if col_idx < 26:
                        col_letter = chr(ord('A') + col_idx)
                    else:
                        col_letter = f"Col{col_idx + 1}"
                    
                    # Attempt to generate sparklines for this column
                    self._obj.Execute(f"sparklines sel:=0 c1:={col_letter} c2:={col_letter}")
                    #print(f"  [OK] Sparklines generated for column {col_idx} ({col_letter})")
                        
                except Exception as e:
                    # If sparklines generation fails for this column, show warning and continue
                    print(f"  [WARNING] Could not generate sparklines for column {col_idx} ({col_letter}): {e}")
                    continue
                    
        except Exception as e:
            print(f"[ERROR] Failed to generate sparklines: {e}")
        
        return

    def _has_numeric_data(self, col_idx: int) -> bool:
        """
        Check if a column contains numeric data that can be used for sparklines.
        
        Args:
            col_idx: Column index to check
            
        Returns:
            bool: True if column has numeric data
        """
        try:
            # Get a sample of data from the column
            col = self.Columns(col_idx)
            
            # Try to get first few values to check if they're numeric
            for row in range(min(5, self.get_rows())):
                try:
                    value = self.get_cell(row, col_idx)
                    if value is not None and value != '':
                        # Try to convert to float to check if numeric
                        float(value)
                        return True
                except (ValueError, TypeError):
                    continue
            
            return False
            
        except Exception:
            return False

    def refresh_sparklines(self) -> None:
        """
        Refresh sparklines for all columns in the worksheet.
        This is useful after data changes.
        """
        try:
            # Clear existing sparklines and regenerate
            self._obj.Execute("wks.labels -S")  # Remove sparklines row
            self._obj.Execute("wks.labels *S")  # Add sparklines row back
            self.generate_sparklines()
            
        except Exception as e:
            print(f"Warning: Could not refresh sparklines: {e}")


class DataPlot:
    """
    Data plot in a graph layer.
    Wrapper class that wraps OriginExt.OriginExt.DataPlot.

    Corresponds to: OriginExt.OriginExt.DataPlot
    """

    _plot: oext_types.DataPlot

    def __init__(self, plot: oext_types.DataPlot, parent: Optional['OriginObjectWrapper'] = None, 
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
    def Name(self) -> str:
        """Short name of the data plot"""
        return self._plot.Name

    @property
    def Parent(self) -> GraphLayer:
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
        try:
            # Use direct property access like in Origin sample: plot.colormap = 'Candy'
            self._plot.colormap = value.value
        except Exception as e:
            print(f"Direct colormap setting failed: {e}")
            # Try alternative approach using _OriginExt
            try:
                oext.OriginObject_DoMethod(self._plot, f"SetColorMap({value.value})")
            except Exception as e2:
                print(f"Alternative colormap setting also failed: {e2}")

    @property
    def shape_list(self) -> list[int]:
        """Get shape list"""
        try:
            return list(self._plot.shapelist)
        except:
            return []

    @shape_list.setter
    def shape_list(self, shapes: list[int]) -> None:
        """Set shape list like in Origin sample: plot.shapelist = [3, 2, 1]"""
        try:
            self._plot.shapelist = shapes
        except Exception as e:
            print(f"Shape list setting failed: {e}")

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

    def set_color_map(self, color_map: ColorMap) -> None:
        """
        Set color map using enum.

        Args:
            color_map: ColorMap enum value
        """
        self._plot.SetColorMap(color_map.value)

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
            yield DataPlot(plot, self, self.origin_instance)

    def __getitem__(self, index: int) -> DataPlot:
        """Get data plot by index"""
        return DataPlot(self._obj[index], self, self.origin_instance)

    def get_page(self) -> GraphPage:
        """
        Get the parent graph page.

        Returns:
            GraphPage: Parent graph page
        """
        from .pages import GraphPage
        return GraphPage(oext.GraphLayer_GetPage(self._obj))

    def add_plot(self, data_range, plot_type: PlotType, composite: bool = False) -> DataPlot:
        """
        Add a plot from a data range using PlotType enum.

        Args:
            data_range: Data range object
            plot_type: PlotType enum value
            composite: Create composite plot

        Returns:
            DataPlot: The newly created data plot
        """
        return DataPlot(oext.GraphLayer_AddPlot(self._obj, data_range, plot_type.value, composite), 
                       self, self.origin_instance)

    def add_plot_from_string(self, range_str: str, plot_type: PlotType) -> DataPlot:
        """
        Add a plot from a range string using PlotType enum.
        Based on Origin Sample #5: gl.add_plot(f'{wks.lt_range()}!(?,1:end)')

        Args:
            range_str: Range string
            plot_type: PlotType enum value

        Returns:
            DataPlot: The newly created data plot
        """
        print(f"Attempting to create plot with range: {range_str}")
        print(f"Plot type: {plot_type} (value: {plot_type.value})")
        
        # Debug: Check if we can access Origin instance through hierarchy
        origin_instance = self.get_origin_instance()
        print(f"Origin instance accessible through hierarchy: {origin_instance is not None}")
        
        try:
            # Use GraphLayer_AddPlotFromString which accepts string directly
            
            # Get the underlying OriginExt GraphLayer object
            gl_obj = get_originext_graphlayer(self)
            
            # Call AddPlotFromString with range string and plot type
            plot_obj = oext.GraphLayer_AddPlotFromString(gl_obj, range_str, plot_type.value)
            
            if plot_obj:
                return DataPlot(plot_obj, self, self.origin_instance)
            else:
                print("GraphLayer_AddPlotFromString returned None")
                
        except Exception as e:
            print(f"GraphLayer_AddPlotFromString failed: {e}")
            # Try alternative approach using the wrapper's method
            try:
                # Try to use string directly with the wrapper
                plot = self._obj.AddPlot(range_str, plot_type.value)
                if plot:
                    return DataPlot(plot, self, self.origin_instance)
            except Exception as e2:
                print(f"Alternative approach also failed: {e2}")
        
        raise RuntimeError(f"Failed to create plot with range: {range_str}")

    def group_plots(self, group_mode: GroupMode = GroupMode.DEPENDENT) -> None:
        """
        Group plots in this layer using GroupMode enum.
        Based on Origin Sample #5: gl.group()

        Args:
            group_mode: GroupMode enum value
        """
        if group_mode != GroupMode.NONE:
            try:
                # Use the wrapper's group method directly like in Origin sample
                self._obj.Group()
            except Exception as e:
                print(f"Direct Group method failed: {e}")
                # Try alternative approach using _OriginExt
                try:
                    gl_obj = get_originext_graphlayer(self)
                    oext.OriginObject_DoMethod(gl_obj, "Group")
                except Exception as e2:
                    print(f"Alternative group method also failed: {e2}")

    def rescale(self) -> None:
        """
        Rescale the layer to fit all data.
        Based on Origin Sample #5: gl.rescale()
        """
        try:
            # Use the wrapper's rescale method directly like in Origin sample
            self._obj.Rescale()
        except Exception as e:
            print(f"Direct Rescale method failed: {e}")
            # Try alternative approach using _OriginExt
            try:
                gl_obj = get_originext_graphlayer(self)
                oext.OriginObject_DoMethod(gl_obj, "Rescale")
            except Exception as e2:
                print(f"Alternative rescale method also failed: {e2}")

    def add_xy_plot(self, worksheet, x_col: int, y_col: int, 
                   plot_type: PlotType = PlotType.LINE_SYMBOL) -> DataPlot:
        """
        Add an XY plot from worksheet columns.
        Based on Origin Sample #4: gl.add_plot(wks, coly='B', colx='A', type=202)

        Args:
            worksheet: Worksheet object containing data
            x_col: X column index (0-based)
            y_col: Y column index (0-based) or -1 for all columns after x_col
            plot_type: PlotType enum value

        Returns:
            DataPlot: The newly created data plot
        """
        # Convert column indices to column letters (A, B, C, ...) like in Sample #4
        def col_index_to_letter(col_idx: int) -> str:
            """Convert 0-based column index to column letter (A, B, C, ...)"""
            result = ""
            col_idx += 1  # Convert to 1-based for calculation
            while col_idx > 0:
                col_idx -= 1
                result = chr(65 + (col_idx % 26)) + result
                col_idx //= 26
            return result
        
        x_col_letter = col_index_to_letter(x_col)
        
        if y_col == -1:
            # Plot all columns after x_col as Y - use the first Y column for now
            # This is a limitation - Sample #4 plots one Y at a time
            y_col_letter = col_index_to_letter(x_col + 1)
        else:
            y_col_letter = col_index_to_letter(y_col)
        
        # Use full worksheet reference like in Sample #4: [BookName]SheetName!col
        # Get worksheet name and parent book name
        wks_name = worksheet.Name
        page = worksheet.get_page()
        book_name = page.Name
        
        # Create range string with full worksheet reference
        range_str = f"[{book_name}]{wks_name}!({x_col_letter},{y_col_letter})"
        
        print(f"Using column range: {range_str}")
        return self.add_plot_from_string(range_str, plot_type)

    def get_data_plots(self) -> list[DataPlot]:
        """
        Get list of data plots in this layer.

        Returns:
            list[DataPlot]: List of data plots
        """
        return [DataPlot(p, self, self.origin_instance) for p in oext.GraphLayer_GetDataPlots(self._obj)]

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
