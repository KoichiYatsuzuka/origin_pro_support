"""
Worksheet classes for OriginExt wrappers.

This module contains wrapper classes for Origin worksheets including:
- Column
- ColumnCollection
- Worksheet
"""
from __future__ import annotations

import OriginExt.OriginExt as oext_types
import OriginExt._OriginExt as oext
import pandas as pd
import numpy as np
from typing import Optional, Tuple, Union, TypeVar, TYPE_CHECKING, overload, List
from collections.abc import Iterator

from ..base import OriginObjectWrapper

if TYPE_CHECKING:
    from ..base import APP


# ================== Type Variables ==================

TColumn = TypeVar('TColumn', bound=oext_types.Column)
TWorksheet = TypeVar('TWorksheet', bound=oext_types.Worksheet)
TDatasheet = TypeVar('TDatasheet', bound=oext_types.Datasheet)


# ================== Datasheet Class ==================

class Datasheet(OriginObjectWrapper[TDatasheet]):
    """
    Base class for data sheets (Worksheet, Matrixsheet).
    Wrapper class that wraps OriginExt.OriginExt.Datasheet.
    """

    def __init__(self, datasheet: TDatasheet, api_core: APP):
        """
        Initialize Datasheet wrapper with hierarchical references.
        """
        super().__init__(datasheet, api_core)

    @property
    def cols(self) -> int:
        """Number of columns"""
        return self._obj.Cols

    @property
    def rows(self) -> int:
        """Number of rows"""
        return self._obj.Rows

    def get_cols(self) -> int:
        """Get number of columns"""
        return self._obj.GetCols()

    def get_rows(self) -> int:
        """Get number of rows"""
        return self._obj.GetRows()

    def set_cols(self, num_cols: int) -> None:
        """Set number of columns"""
        self._obj.SetCols(num_cols)

    def set_rows(self, num_rows: int) -> None:
        """Set number of rows"""
        self._obj.SetRows(num_rows)


# ================== Column Collection ==================

class ColumnCollection:
    """
    Wrapper for Origin column collection that returns wrapped Column objects.
    """
    
    def __init__(self, columns_collection, api_core: 'APP'):
        self._columns = columns_collection
        self.__API_core = api_core
    
    def __getitem__(self, index: int) -> 'Column':
        """Get column by index"""
        return Column(self._columns[index], self.__API_core)
    
    def __len__(self) -> int:
        """Get number of columns"""
        return len(self._columns)
    
    def __iter__(self):
        """Iterate over columns, yielding wrapped Column objects"""
        for col in self._columns:
            yield Column(col, self.__API_core)


# ================== Column Class ==================

class Column(OriginObjectWrapper[TColumn]):
    """
    Column in a worksheet.
    Wrapper class that wraps OriginExt.OriginExt.Column.

    Corresponds to: originpro.Column, OriginExt.OriginExt.Column
    """

    def __init__(self, column: TColumn, api_core: 'APP'):
        """
        Initialize Column wrapper with hierarchical references.

        Args:
            column: Original OriginExt.Column instance to wrap
            api_core: APP instance reference for LabTalk access
        """
        super().__init__(column, api_core)

    @property
    def name(self) -> str:
        """Short name of the column"""
        return self._obj.Name

    @name.setter
    def name(self, value: str) -> None:
        """Set short name"""
        self._obj.Name = value

    @property
    def long_name(self) -> str:
        """Long name of the column"""
        return self._obj.LongName

    @long_name.setter
    def long_name(self, value: str) -> None:
        """Set long name"""
        self._obj.LongName = value

    @property
    def type(self) -> int:
        """Column type"""
        return self._obj.Type

    @type.setter
    def type(self, value: int) -> None:
        """Set column type"""
        self._obj.Type = value

    @property
    def units(self) -> str:
        """Column units"""
        return self._obj.Units

    @units.setter
    def units(self, value: str) -> None:
        """Set column units"""
        self._obj.Units = value

    @property
    def comments(self) -> str:
        """Column comments"""
        return self._obj.Comments

    @comments.setter
    def comments(self, value: str) -> None:
        """Set column comments"""
        self._obj.Comments = value

    @property
    def parent(self) -> 'Worksheet':
        """Parent worksheet"""
        return Worksheet(self._obj.Parent, self, self.api_core)

    def get_parent(self) -> 'Worksheet':
        """
        Get the parent worksheet.

        Corresponds to: OriginExt.OriginExt.Column.GetParent()

        Returns:
            Worksheet: Parent worksheet
        """
        return Worksheet(self._obj.GetParent(), self, self.api_core)

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

    def set_data(self, data, offset: int = 0):
        """
        Set column data.

        Corresponds to: OriginExt.OriginExt.Column.SetData()

        Args:
            data: Data array to set
            offset: Row offset (default: 0)

        Returns:
            bool: True if successful
        """
        return self._obj.SetData(data, offset)

    def is_valid(self) -> bool:
        """
        Check if this column is valid.

        Corresponds to: OriginExt.OriginExt.Column.IsValid()

        Returns:
            bool: True if valid
        """
        return self._column.IsValid()


# ================== Worksheet Class ==================

class Worksheet(Datasheet[TWorksheet]):
    """
    Worksheet layer for tabular data.
    Wrapper class that wraps OriginExt.OriginExt.Worksheet.

    Corresponds to: originpro.WSheet, OriginExt.OriginExt.Worksheet
    """

    def __init__(self, worksheet: TWorksheet, api_core: APP, 
                data: Optional[pd.DataFrame | np.ndarray | List[List[Any]]] = None):
        """
        Initialize Worksheet wrapper with hierarchical references.
        Can optionally load data during initialization.

        Args:
            worksheet: Original OriginExt.Worksheet instance to wrap
            api_core: APP instance reference for LabTalk access
            parent: Parent wrapper object (for hierarchical navigation)
            data: Optional data to load (pd.DataFrame, 2D np.ndarray, or 2D list)
        
        Raises:
            TypeError: If data is not None and not one of the supported types
            ValueError: If np.ndarray is not 2D or list is not 2D
        """
        super().__init__(worksheet, api_core)
        
        # Automatically set header rows to show: Long Name, Units, Sparklines, F(x), Comments
        self.header_rows('LUSCO')
         
        # Ensure sparklines are properly activated and generated
        self._ensure_sparklines()
        
        # Load data if provided
        if data is not None:
            self.add_column_from_data(data)
        else:
            # Create two empty columns if no data provided
            self.set_cols(2)
            self.set_rows(0)

    @property
    def columns(self):
        """Collection of columns in this worksheet"""
        # Return a wrapper that provides access to wrapped Column objects
        return ColumnCollection(self._obj.Columns, self)

    def __iter__(self) -> Iterator[Column]:
        """Iterate over columns"""
        for col in self._obj:
            yield Column(col, self.api_core)

    def __getitem__(self, index: int) -> Column:
        """Get column by index"""
        return Column(self._obj[index], self.api_core)

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

    def set_cell(self, row: int, col: int, value):
        """
        Set cell value at specified row and column.

        Args:
            row: Row index
            col: Column index
            value: Value to set
        """
        self._obj.SetCell(row, col, value)

    def get_columns(self) -> list[Column]:
        """
        Get list of columns in this worksheet.

        Corresponds to: OriginExt.OriginExt.Worksheet.GetColumns()

        Returns:
            list[Column]: List of columns
        """
        return [Column(c, self.api_core) for c in self._obj.GetColumns()]

    def header_rows(self, spec: str = 'LUSCO') -> None:
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
    @overload
    def add_column_from_data(self, data: List, lname: Optional[str] = None,
                           units: Optional[str] = None, comments: Optional[str] = None,
                           axis: Optional[str] = None) -> 'Column': ...

    @overload
    def add_column_from_data(self, data: pd.Series, lname: Optional[str] = None,
                           units: Optional[str] = None, comments: Optional[str] = None,
                           axis: Optional[str] = None) -> 'Column': ...

    @overload
    def add_column_from_data(self, data: np.ndarray, lname: Optional[List[str]] = None,
                           units: Optional[str] = None, comments: Optional[str] = None,
                           axis: Optional[str] = None) -> Union['Column', List['Column']]: ...

    @overload
    def add_column_from_data(self, data: pd.DataFrame, lname: Optional[List[str]] = None,
                           units: Optional[str] = None, comments: Optional[str] = None,
                           axis: Optional[str] = None) -> Union['Column', List['Column']]: ...

    def add_column_from_data(self, data, lname=None,
                           units: Optional[str] = None, comments: Optional[str] = None,
                           axis: Optional[str] = None):
        """
        Add a new column to the worksheet from various data types.

        lname behaviour by data type:

        * 1-D ``list`` / 1-D ``np.ndarray``:
            ``Optional[str]``.  When *None* a serial long name ``list_1``,
            ``list_2``, … is generated automatically (continuing from the
            highest existing ``list_N`` in the current worksheet).

        * 2-D ``list`` / 2-D ``np.ndarray``:
            ``Optional[list[str]]``.  When *None* each new column receives an
            auto-generated ``list_N`` name.  When a ``list[str]`` is supplied
            its length must equal the number of columns (axis-1 size);
            a ``ValueError`` is raised otherwise.

        * ``pd.Series``:
            *lname* is **ignored**; the series' ``.name`` attribute is used as
            the column long name (auto-generated if ``.name`` is *None*).
            A warning is printed when a non-None *lname* is supplied.

        * ``pd.DataFrame``:
            *lname* is **ignored**; each DataFrame column's name is used.
            A warning is printed when a non-None *lname* is supplied.

        Args:
            data: Input data (list, np.ndarray, pd.Series, or pd.DataFrame)
            lname: Long name(s) — see above for type and semantics per data type.
            units: Optional units for the column(s)
            comments: Optional comments for the column(s)
            axis: Optional axis designation ('X', 'Y', 'Z', 'E', etc.)

        Returns:
            Column or List[Column]: The newly created column(s)

        Raises:
            TypeError: If data is an unsupported type.
            ValueError: If data is 3-D+, or if lname length mismatches column count.
        """

        if not isinstance(data, (list, pd.DataFrame, pd.Series, np.ndarray)):
            raise TypeError(
                f"Unsupported data type: {type(data)}. "
                "Supported types: list, numpy.ndarray, pandas.Series, pandas.DataFrame"
            )

        # ── pd.Series ──────────────────────────────────────────────────────
        if isinstance(data, pd.Series):
            if lname is not None:
                print(
                    "Warning: lname is ignored for pd.Series; "
                    "the series .name attribute is used instead."
                )
            series_name = str(data.name) if data.name is not None else None
            return self._add_column_from_1d_data(data.tolist(), series_name, units, comments, axis)

        # ── pd.DataFrame ───────────────────────────────────────────────────
        if isinstance(data, pd.DataFrame):
            if lname is not None:
                print(
                    "Warning: lname is ignored for pd.DataFrame; "
                    "the DataFrame column names are used instead."
                )
            return self._add_column_from_dataframe(data, units, comments, axis)

        # ── 1-D np.ndarray ─────────────────────────────────────────────────
        if isinstance(data, np.ndarray):
            if data.ndim == 1:
                return self._add_column_from_1d_data(data.tolist(), lname, units, comments, axis)
            elif data.ndim == 2:
                return self._add_column_from_2d_array(data, lname, units, comments, axis)
            else:
                raise ValueError("numpy.ndarray must be 1-D or 2-D.")

        # ── list ───────────────────────────────────────────────────────────
        if isinstance(data, list):
            if not data or not isinstance(data[0], list):
                # 1-D list
                valid_types = (int, float, str)
                if any(not isinstance(item, valid_types) for item in data if item is not None):
                    raise ValueError("1-D list elements must be int, float, or str.")
                return self._add_column_from_1d_data(data, lname, units, comments, axis)
            else:
                # 2-D list
                if data[0] and isinstance(data[0][0], list):
                    raise ValueError(
                        "3-D or deeper nested lists are not supported. "
                        "Use a 2-D list, numpy.ndarray, pd.Series, or pd.DataFrame."
                    )
                return self._add_column_from_2d_list(data, lname, units, comments, axis)

        raise ValueError(
            f"Unsupported data format: {type(data)}. "
            "Expected 1-D/2-D list, numpy.ndarray, pd.Series, or pd.DataFrame."
        )
    def _get_next_list_number(self) -> int:
        """Return the next available serial number for auto-generated 'list_N' long names.

        Scans all existing columns for long names that match the pattern ``list_<integer>``,
        finds the maximum N, and returns N+1 (or 1 if none are found).
        """
        import re
        pattern = re.compile(r'^list_(\d+)$')
        max_n = 0
        for col in self._obj.GetColumns():
            m = pattern.match(col.LongName)
            if m:
                n = int(m.group(1))
                if n > max_n:
                    max_n = n
        return max_n + 1

    def _set_axis_type(self, col: 'Column', axis: Optional[str]) -> None:
        """Apply axis designation to a column (shared helper)."""
        if axis is None:
            return
        axis_map = {'X': 1, 'Y': 2, 'Z': 3, 'E': 4}
        if isinstance(axis, str) and axis.upper() in axis_map:
            col.type = axis_map[axis.upper()]
        elif isinstance(axis, int):
            col.type = axis

    def _add_column_from_1d_data(self, data_list: list, lname: Optional[str] = None,
                                units: Optional[str] = None, comments: Optional[str] = None,
                                axis: Optional[str] = None) -> 'Column':
        """Add a single column from 1-D list data.

        Args:
            data_list: 1-D list of values.
            lname: Long name.  When *None* an auto-generated ``list_N`` name is used.
            units: Optional units.
            comments: Optional comments.
            axis: Optional axis designation ('X', 'Y', 'Z', 'E', …).

        Returns:
            Column: The newly created column.
        """
        current_cols = self.get_cols()
        self.set_cols(current_cols + 1)
        new_col = self.columns[current_cols]

        effective_lname = lname if lname is not None else f"list_{self._get_next_list_number()}"
        new_col.long_name = effective_lname
        if units is not None:
            new_col.units = units
        if comments is not None:
            new_col.comments = comments
        self._set_axis_type(new_col, axis)

        new_col.set_data(data_list)
        return new_col

    def _add_column_from_dataframe(self, df: pd.DataFrame,
                                  units: Optional[str] = None, comments: Optional[str] = None,
                                  axis: Optional[str] = None) -> Union['Column', List['Column']]:
        """Add columns from a pandas DataFrame.

        The DataFrame's own column names are always used as long names.

        Args:
            df: pandas DataFrame to load.
            units: Optional units for all columns.
            comments: Optional comments for all columns.
            axis: Optional axis designation for all columns.

        Returns:
            Column or List[Column]: The newly created column(s).
        """
        num_cols = len(df.columns)
        current_cols = self.get_cols()
        self.set_cols(current_cols + num_cols)

        new_columns = []
        for i, col_name in enumerate(df.columns):
            new_col = self.columns[current_cols + i]
            new_col.long_name = str(col_name)
            if units is not None:
                new_col.units = units
            if comments is not None:
                new_col.comments = comments
            self._set_axis_type(new_col, axis)
            new_col.set_data(df[col_name].tolist())
            new_columns.append(new_col)

        return new_columns[0] if len(new_columns) == 1 else new_columns

    def _add_column_from_2d_array(self, arr: np.ndarray,
                                 lname: Optional[List[str]] = None,
                                 units: Optional[str] = None, comments: Optional[str] = None,
                                 axis: Optional[str] = None) -> Union['Column', List['Column']]:
        """Add columns from a 2-D numpy array (each column of the array → one Origin column).

        Args:
            arr: 2-D numpy array.  Shape is (rows, num_cols).
            lname: ``None`` for auto-generated ``list_N`` names, or a ``list[str]``
                   whose length must equal ``arr.shape[1]``.
            units: Optional units for all columns.
            comments: Optional comments for all columns.
            axis: Optional axis designation for all columns.

        Returns:
            Column or List[Column]: The newly created column(s).

        Raises:
            ValueError: If *lname* is a list whose length != number of columns.
        """
        num_cols = arr.shape[1]

        if lname is not None and len(lname) != num_cols:
            raise ValueError(
                f"lname list length ({len(lname)}) must match the number of columns ({num_cols})."
            )

        current_cols = self.get_cols()
        self.set_cols(current_cols + num_cols)
        next_n = self._get_next_list_number()

        new_columns = []
        for i in range(num_cols):
            new_col = self.columns[current_cols + i]
            new_col.long_name = lname[i] if lname is not None else f"list_{next_n + i}"
            if units is not None:
                new_col.units = units
            if comments is not None:
                new_col.comments = comments
            self._set_axis_type(new_col, axis)
            new_col.set_data(arr[:, i].tolist())
            new_columns.append(new_col)

        return new_columns[0] if len(new_columns) == 1 else new_columns

    def _add_column_from_2d_list(self, data: List[List],
                                lname: Optional[List[str]] = None,
                                units: Optional[str] = None, comments: Optional[str] = None,
                                axis: Optional[str] = None) -> Union['Column', List['Column']]:
        """Add columns from a 2-D list (list of rows).

        Args:
            data: 2-D list with shape [rows][cols].
            lname: ``None`` for auto-generated ``list_N`` names, or a ``list[str]``
                   whose length must equal the number of columns (``len(data[0])``).
            units: Optional units for all columns.
            comments: Optional comments for all columns.
            axis: Optional axis designation for all columns.

        Returns:
            Column or List[Column]: The newly created column(s).

        Raises:
            ValueError: If *lname* is a list whose length != number of columns.
        """
        num_cols = len(data[0]) if data else 0

        if lname is not None and len(lname) != num_cols:
            raise ValueError(
                f"lname list length ({len(lname)}) must match the number of columns ({num_cols})."
            )

        current_cols = self.get_cols()
        self.set_cols(current_cols + num_cols)
        next_n = self._get_next_list_number()

        new_columns = []
        for i in range(num_cols):
            new_col = self.columns[current_cols + i]
            new_col.long_name = lname[i] if lname is not None else f"list_{next_n + i}"
            if units is not None:
                new_col.units = units
            if comments is not None:
                new_col.comments = comments
            self._set_axis_type(new_col, axis)
            col_data = [row[i] for row in data]
            new_col.set_data(col_data)
            new_columns.append(new_col)

        return new_columns[0] if len(new_columns) == 1 else new_columns
