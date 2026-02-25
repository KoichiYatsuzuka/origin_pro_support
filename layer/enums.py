"""
Enumeration types for OriginExt wrappers.

This module contains all enum types used across the layer modules.
"""
from __future__ import annotations

from enum import Enum


# ================== Plot Enums ==================

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
    """Enumeration for axis scale types."""
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
    # Plot enums
    'PlotType',
    'XYTemplate',
    # Color and style enums
    'ColorMap',
    'GroupMode',
    # Axis enums
    'AxisType',
    'ScaleType',
    'TickType',
]
