"""
Enumeration types for OriginExt wrappers.

This module contains all enum types used across the layer modules.
"""
from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from typing import Union


# ================== Data Classes ==================

@dataclass(frozen=True)
class PlotTypeInfo:
    """Immutable data class for plot type information.
    See the docstring of XYPlotType for more details."""
    plot_id: int
    template_name: str
    
    def __int__(self) -> int:
        """Cast to int, returns numeric value."""
        return self.plot_id
    
    def __str__(self) -> str:
        """Cast to str, returns template value."""
        return self.template_name


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


class OriginColorIndex(Enum):
    """Named color indices from Origin's standard color list (1-24).

    These correspond to Origin's built-in color palette.
    Ref: https://www.originlab.com/doc/LabTalk/guide/Specifying-Colors
    """
    BLACK = 1
    RED = 2
    GREEN = 3
    BLUE = 4
    CYAN = 5
    MAGENTA = 6
    YELLOW = 7
    DARK_YELLOW = 8
    NAVY = 9
    PURPLE = 10
    WINE = 11
    OLIVE = 12
    DARK_CYAN = 13
    ROYAL = 14
    ORANGE = 15
    VIOLET = 16
    PINK = 17
    WHITE = 18
    LIGHT_GRAY = 19
    GRAY = 20
    LT_YELLOW = 21
    LT_CYAN = 22
    LT_MAGENTA = 23
    DARK_GRAY = 24


# Type alias for color specifications accepted throughout this library.
# - int: Origin color index (1-24)
# - tuple[int, int, int]: RGB values (0-255 each)
# - OriginColorIndex: named color constant
ColorSpec = Union[int, tuple, 'OriginColorIndex']


def color_to_lt_str(color: ColorSpec) -> str:
    """Convert a ColorSpec to the LabTalk color string used in commands.

    Args:
        color: int index, (R, G, B) tuple, or OriginColorIndex.

    Returns:
        str: LabTalk representation, e.g. ``"2"`` or ``"color(240,208,0)"``.
    """
    if isinstance(color, OriginColorIndex):
        return str(color.value)
    if isinstance(color, tuple):
        r, g, b = color
        return f"color({r},{g},{b})"
    return str(int(color))


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
    'OriginColorIndex',
    'ColorSpec',
    'color_to_lt_str',
    'GroupMode',
    # Axis enums
    'AxisType',
    'ScaleType',
    'TickType',
]
