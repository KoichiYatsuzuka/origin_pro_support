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
    from ..origin_instance import OriginInstance


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

    def __init__(self, datasheet: TDatasheet, parent: Optional['OriginObjectWrapper'] = None, 
                 origin_instance: Optional['OriginInstance'] = None):
        """
        Initialize Datasheet wrapper with hierarchical references.
        """
        super().__init__(datasheet, parent, origin_instance)

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
    
    def __init__(self, columns_collection, parent: Optional['OriginObjectWrapper'] = None):
        self._columns = columns_collection
        self._parent = parent
    
    def __getitem__(self, index: int) -> 'Column':
        """Get column by index"""
        return Column(self._columns[index], self._parent, self._parent.origin_instance if self._parent else None)
    
    def __len__(self) -> int:
        """Get number of columns"""
        return len(self._columns)
    
    def __iter__(self):
        """Iterate over columns, yielding wrapped Column objects"""
        for col in self._columns:
            yield Column(col, self._parent, self._parent.origin_instance if self._parent else None)


# ================== Column Class ==================

class Column(OriginObjectWrapper[TColumn]):
    """
    Column in a worksheet.
    Wrapper class that wraps OriginExt.OriginExt.Column.

    Corresponds to: originpro.Column, OriginExt.OriginExt.Column
    """

    def __init__(self, column: TColumn, parent: Optional['OriginObjectWrapper'] = None, 
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
        return Worksheet(self._obj.Parent, self._parent, self.origin_instance)

    def get_parent(self) -> 'Worksheet':
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


# ================== Worksheet Class ==================

class Worksheet(Datasheet[TWorksheet]):
    """
    Worksheet layer for tabular data.
    Wrapper class that wraps OriginExt.OriginExt.Worksheet.

    Corresponds to: originpro.WSheet, OriginExt.OriginExt.Worksheet
    """

    @overload
    def __init__(self, worksheet: TWorksheet, origin_instance: 'OriginInstance', 
                 parent: 'OriginObjectWrapper') -> None: ...
    
    @overload
    def __init__(self, worksheet: TWorksheet, origin_instance: 'OriginInstance', 
                 parent: 'OriginObjectWrapper',
                 data: pd.DataFrame) -> None: ...
    
    @overload
    def __init__(self, worksheet: TWorksheet, origin_instance: 'OriginInstance', 
                 parent: 'OriginObjectWrapper', 
                 data: np.ndarray) -> None: ...
    
    @overload
    def __init__(self, worksheet: TWorksheet, origin_instance: 'OriginInstance', 
                 parent: 'OriginObjectWrapper', data: Optional[pd.DataFrame | np.ndarray | List[List[Any]]] = None):
        """
        Initialize Worksheet wrapper with hierarchical references.
        Can optionally load data during initialization.

        Args:
            worksheet: Original OriginExt.Worksheet instance to wrap
            origin_instance: Root OriginInstance reference (for LabTalk access)
            parent: Parent wrapper object (for hierarchical navigation)
            data: Optional data to load (pd.DataFrame, 2D np.ndarray, or 2D list)
        
        Raises:
            TypeError: If data is not None and not one of the supported types
            ValueError: If np.ndarray is not 2D or list is not 2D
        """
        super().__init__(worksheet, parent, origin_instance)
        
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
