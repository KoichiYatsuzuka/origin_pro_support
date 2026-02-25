"""
Enumeration types for OriginExt wrappers.

This module contains all enum types used across the layer modules.
"""
from __future__ import annotations

from dataclasses import dataclass
from enum import Enum


# ================== Data Classes ==================

@dataclass(frozen=True)
class PlotTypeInfo:
    """Immutable data class for plot type information.
    See the docstring of XYPlotType for more details."""
    numeric_value: int
    template_value: str
    
    def __int__(self) -> int:
        """Cast to int, returns numeric value."""
        return self.numeric_value
    
    def __str__(self) -> str:
        """Cast to str, returns template value."""
        return self.template_value


# ================== Plot Enums ==================

class XYPlotType(Enum):
    """Enumeration for XY plot types to avoid literal values.
    The inteer numbers correspond to the values used for LAbTalk commands.
    Ref: https://www.originlab.com/doc/LabTalk/ref/Plot-Type-IDs \n
    The string values correspond to the names of template files.
    Ref: https://www.originlab.com/doc/X-Function/ref/newpanel
    Note that the roles of theese values are slighly different.

    User-defeined plot types may be created to use user-defiend template.  
    """
    LINE = PlotTypeInfo(200, "line")
    SCATTER = PlotTypeInfo(201, "scatter")
    LINE_SYMBOL = PlotTypeInfo(202, "linesymb")
    COLUMN = PlotTypeInfo(203, "column")
    BAR = PlotTypeInfo(204, "bar")
    AREA = PlotTypeInfo(205, "area")
    PIE = PlotTypeInfo(206, "pie")
    SURFACE = PlotTypeInfo(207, "surface")
    CONTOUR = PlotTypeInfo(208, "contour")
    HISTOGRAM = PlotTypeInfo(209, "histogram")
    

# ================== Color and Style Enums ==================

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


# ================== Axis Enums ==================

class AxisType(Enum):
    """Enumeration for axis types."""
    X = 1
    Y = 2
    Z = 3
    ERROR = 4


class ScaleType(Enum):
    """Enumeration for axis scale types.
    The integer numbers correspond to the values used for LAbTalk commands.
    Ref: https://www.originlab.com/doc/LabTalk/ref/Layer-Axis-obj"""
    LINEAR = 1
    LOG10 = 2
    PROBABILITY = 3
    PROBIT = 4
    RECIPROCAL = 5
    OFFSET_RECIPROCAL = 6
    LOGIT = 7
    LN = 8
    LOG2 = 9


class TickType(Enum):
    """Enumeration for tick types."""
    NONE = 0
    IN = 1
    OUT = 2
    IN_OUT = 3


# ================== Export ==================

__all__ = [
    # Data classes
    'PlotTypeInfo',
    # Plot enums
    'XYPlotType',
    # Color and style enums
    'ColorMap',
    'GroupMode',
    # Axis enums
    'AxisType',
    'ScaleType',
    'TickType',
]
