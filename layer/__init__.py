"""
Layer classes for OriginExt wrappers.

This module contains wrapper classes for Origin layers including:
- Layer (base class)
- Datasheet
- Matrixsheet
- Enum types for plotting
"""
from __future__ import annotations

import OriginExt.OriginExt as oext_types
from enum import Enum
from typing import Optional, TypeVar, TYPE_CHECKING

from ..base import OriginObjectWrapper

# Import from submodules
from .worksheet import Column, Worksheet, Datasheet
from .graph_layer import DataPlot, Axis, GraphLayer, Legend
from .enums import (
    XYPlotType,
    ColorMap,
    OriginColorIndex,
    ColorSpec,
    color_to_lt_str,
    GroupMode,
    AxisType,
    ScaleType,
    TickType,
    LegendLayout,
    MarkerShape,
    LineStyle,
)
from base import APP


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

    def __init__(self, layer: TLayer, api_core: 'APP'):
        """
        Initialize Layer wrapper with hierarchical references.

        Args:
            layer: Original OriginExt layer instance to wrap
            api_core: APP instance reference for LabTalk access
        """
        super().__init__(layer, api_core)

    def get_data_object_bases(self):
        """
        Get collection of data objects in this layer.

        Corresponds to: OriginExt.OriginExt.Layer.GetDataObjectBases()

        Returns:
            Collection of data objects
        """
        return self._obj.GetDataObjectBases()


# ================== Matrixsheet Class ==================

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
    def matrix_objects(self):
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


# ================== Re-export ==================

__all__ = [
    # Base classes
    'Layer',
    'Datasheet',
    # Worksheet classes
    'Column',
    'Worksheet',
    # Graph layer classes
    'DataPlot',
    'Axis',
    'GraphLayer',
    'Legend',
    # Matrix classes
    'Matrixsheet',
    # Enum types
    'XYPlotType',
    'ColorMap',
    'OriginColorIndex',
    'ColorSpec',
    'color_to_lt_str',
    'GroupMode',
    'AxisType',
    'ScaleType',
    'TickType',
    'MarkerShape',
    'LineStyle',
    'LegendLayout',
]
