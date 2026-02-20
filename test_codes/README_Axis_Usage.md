# Axis Manipulation Functions

This document describes the axis manipulation functions implemented for FigureLayer and GraphLayer classes.

## Overview

The axis manipulation functions allow you to:
- Get and set axis ranges (min/max values)
- Modify axis properties (scale, title, ticks)
- Rescale and reverse axes
- Access individual axis objects for detailed control

## Basic Usage

### Getting and Setting Axis Ranges

```python
from origin_pro_support import OriginInstance
from origin_pro_support.layers import AxisType

# Create Origin instance and graph
origin = OriginInstance()
graph_page = origin.new_graph()
figure_page = origin.FigurePage(graph_page)
layer = figure_page.get_active_layer()

# Get current axis ranges
x_range = layer.get_x_range()  # Returns (min_val, max_val)
y_range = layer.get_y_range()
print(f"X range: {x_range}")
print(f"Y range: {y_range}")

# Set new axis ranges
layer.set_x_range(0, 100)
layer.set_y_range(-10, 10)

# Generic axis methods
layer.set_axis_range(AxisType.X, 0, 100)
layer.set_axis_range(AxisType.Y, -10, 10)
x_range = layer.get_axis_range(AxisType.X)
```

### Working with Axis Objects

```python
# Get axis objects
x_axis = layer.get_x_axis()
y_axis = layer.get_y_axis()
z_axis = layer.get_z_axis()  # For 3D plots

# Or get by type
x_axis = layer.get_axis(AxisType.X)
y_axis = layer.get_axis(AxisType.Y)
```

## Axis Object Methods

### Range Operations

```python
# Get and set range
min_val, max_val = x_axis.get_range()
x_axis.set_range(0, 50)

# Rescale axis to fit all data
x_axis.rescale()
```

### Scale Types

```python
# Get current scale
scale = x_axis.get_scale()  # 'linear', 'log10', 'ln', etc.

# Set scale type
x_axis.set_scale('linear')
x_axis.set_scale('log10')
x_axis.set_scale('ln')
x_axis.set_scale('reciprocal')
x_axis.set_scale('probability')
x_axis.set_scale('probit')
x_axis.set_scale('logit')
```

### Axis Title

```python
# Get current title
title = x_axis.get_title()

# Set new title
x_axis.set_title("Time (seconds)")
y_axis.set_title("Temperature (°C)")
```

### Tick Settings

```python
# Major tick type
tick_type = x_axis.get_major_tick_type()  # 'none', 'in', 'out', 'in_out'
x_axis.set_major_tick_type('in_out')

# Minor ticks
num_minor = x_axis.get_minor_ticks()  # Number of minor ticks between major ticks
x_axis.set_minor_ticks(4)
```

### Axis Direction

```python
# Check if axis is reversed
is_reversed = x_axis.is_reversed()

# Reverse axis direction
x_axis.reverse(True)   # Reverse axis
x_axis.reverse(False)  # Normal direction
```

## Convenience Methods

### GraphLayer Shortcuts

```python
# Direct axis rescaling
layer.rescale_x_axis()
layer.rescale_y_axis()
layer.rescale_z_axis()

# Generic rescaling
layer.rescale_axis(AxisType.X)
```

## Complete Example

```python
from origin_pro_support import OriginInstance
from origin_pro_support.layers import AxisType
import numpy as np

def create_customized_plot():
    origin = OriginInstance()
    
    try:
        # Create test data
        x = np.linspace(0.1, 100, 200)
        y1 = x**2
        y2 = np.log10(x)
        
        # Create workbook
        workbook = origin.new_workbook()
        worksheet = workbook[0]
        worksheet.from_list(0, x.tolist(), "X", "", "X values")
        worksheet.from_list(1, y1.tolist(), "Quadratic", "", "y = x²")
        worksheet.from_list(2, y2.tolist(), "Logarithmic", "", "y = log₁₀(x)")
        
        # Create graph
        graph_page = origin.new_graph()
        figure_page = origin.FigurePage(graph_page)
        layer = figure_page.get_active_layer()
        
        # Plot data
        layer.add_xy_plot(worksheet, 0, 1)
        layer.add_xy_plot(worksheet, 0, 2)
        
        # Customize axes
        x_axis = layer.get_x_axis()
        y_axis = layer.get_y_axis()
        
        # Set axis ranges
        x_axis.set_range(0, 120)
        y_axis.set_range(-2, 10)
        
        # Set axis titles
        x_axis.set_title("X Values (log scale)")
        y_axis.set_title("Y Values")
        
        # Set X axis to log scale
        x_axis.set_scale('log10')
        
        # Customize ticks
        x_axis.set_major_tick_type('out')
        x_axis.set_minor_ticks(8)
        y_axis.set_major_tick_type('in_out')
        y_axis.set_minor_ticks(4)
        
        # Group and rescale
        layer.group_plots()
        
        print(f"Final X range: {layer.get_x_range()}")
        print(f"Final Y range: {layer.get_y_range()}")
        print(f"X scale: {x_axis.get_scale()}")
        print(f"X title: {x_axis.get_title()}")
        
        # Save
        origin.save("customized_axes.opju")
        
    finally:
        origin.close()

create_customized_plot()
```

## Error Handling

All axis methods include error handling with fallback approaches:

```python
try:
    x_range = layer.get_x_range()
    print(f"X range: {x_range}")
except Exception as e:
    print(f"Failed to get X range: {e}")
```

## Supported Axis Types

- `AxisType.X`: X-axis (horizontal)
- `AxisType.Y`: Y-axis (vertical)  
- `AxisType.Z`: Z-axis (for 3D plots)
- `AxisType.ERROR`: Error bars axis

## Supported Scale Types

- `'linear'`: Linear scale
- `'log10'`: Base-10 logarithmic scale
- `'ln'`: Natural logarithmic scale
- `'reciprocal'`: Reciprocal scale (1/x)
- `'offset reciprocal'`: Offset reciprocal scale
- `'probability'`: Probability scale
- `'probit'`: Probit scale
- `'logit'`: Logit scale

## Supported Tick Types

- `'none'`: No ticks
- `'in'`: Ticks pointing inward
- `'out'`: Ticks pointing outward
- `'in_out'`: Ticks on both sides

## Notes

- All methods work with both 2D and 3D graphs (where applicable)
- Z-axis operations will fail for 2D plots with appropriate error messages
- Functions use LabTalk commands internally for maximum compatibility
- Automatic fallback to alternative methods if primary approach fails
