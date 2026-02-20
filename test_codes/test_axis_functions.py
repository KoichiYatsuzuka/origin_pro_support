"""
Test script for axis manipulation functions in GraphLayer.

This script demonstrates how to use the new axis functions to:
- Get and set axis ranges
- Modify axis properties like scale, title, ticks
- Rescale and reverse axes
"""
import os
import sys
import numpy as np

# Add the parent directory to the path to import our module
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    import origin_pro_support as ops
    from origin_pro_support import (
        FigurePage, PlotType, ColorMap, GroupMode, AxisType,
        OriginInstance, OriginPath
    )
except ImportError as e:
    print(f"Import error: {e}")
    print("Make sure Origin is installed and the OriginExt module is available")
    sys.exit(1)


def test_axis_functions():
    """Test axis manipulation functions."""
    
    # Define test file path
    test_file = os.path.join(os.path.dirname(__file__), "axis_test.opju")
    
    # Clean up existing file
    if os.path.exists(test_file):
        os.remove(test_file)
    
    # Create Origin instance
    origin = OriginInstance(test_file)
    
    try:
        # Create test data
        x_data = np.linspace(0, 10, 100)
        y_data = np.sin(x_data) + 0.1 * np.random.randn(100)
        
        # Create workbook and add data
        workbook = origin.new_workbook("TestData")
        worksheet = workbook[0]
        worksheet.from_list(0, x_data.tolist(), "Time", "s", "X-axis data")
        worksheet.from_list(1, y_data.tolist(), "Signal", "V", "Y-axis data")
        
        # Create graph page
        graph_page = origin.new_graph("TestGraph")
        figure_page = FigurePage(graph_page)
        
        # Plot data
        layer = figure_page.get_active_layer()
        plot = layer.add_xy_plot(worksheet, 0, 1)
        
        print("=== Axis Functions Test ===\n")
        
        # Test basic axis range operations
        print("1. Testing axis range operations:")
        x_range = layer.get_x_range()
        y_range = layer.get_y_range()
        print(f"   Initial X range: {x_range}")
        print(f"   Initial Y range: {y_range}")
        
        # Set custom ranges
        layer.set_x_range(-1, 11)
        layer.set_y_range(-2, 2)
        
        x_range_new = layer.get_x_range()
        y_range_new = layer.get_y_range()
        print(f"   New X range: {x_range_new}")
        print(f"   New Y range: {y_range_new}")
        
        # Test axis object methods
        print("\n2. Testing axis object methods:")
        x_axis = layer.get_x_axis()
        y_axis = layer.get_y_axis()
        
        # Test axis titles
        print(f"   X axis title: '{x_axis.get_title()}'")
        print(f"   Y axis title: '{y_axis.get_title()}'")
        
        # Set axis titles
        x_axis.set_title("Time (seconds)")
        y_axis.set_title("Signal Amplitude (V)")
        
        print(f"   Updated X axis title: '{x_axis.get_title()}'")
        print(f"   Updated Y axis title: '{y_axis.get_title()}'")
        
        # Test axis scale
        print(f"   X axis scale: {x_axis.get_scale()}")
        print(f"   Y axis scale: {y_axis.get_scale()}")
        
        # Test tick settings
        print(f"   X major tick type: {x_axis.get_major_tick_type()}")
        print(f"   X minor ticks: {x_axis.get_minor_ticks()}")
        
        # Change tick settings
        x_axis.set_major_tick_type('in_out')
        x_axis.set_minor_ticks(4)
        
        print(f"   Updated X major tick type: {x_axis.get_major_tick_type()}")
        print(f"   Updated X minor ticks: {x_axis.get_minor_ticks()}")
        
        # Test axis reversal
        print(f"   X axis reversed: {x_axis.is_reversed()}")
        x_axis.reverse(True)
        print(f"   X axis reversed after setting: {x_axis.is_reversed()}")
        x_axis.reverse(False)  # Reset to normal
        print(f"   X axis reversed after reset: {x_axis.is_reversed()}")
        
        # Test axis rescaling
        print("\n3. Testing axis rescaling:")
        layer.set_x_range(0, 5)  # Narrow range
        layer.set_y_range(0, 1)
        print(f"   Before rescaling - X: {layer.get_x_range()}, Y: {layer.get_y_range()}")
        
        layer.rescale_x_axis()
        print(f"   After X rescaling - X: {layer.get_x_range()}, Y: {layer.get_y_range()}")
        
        layer.rescale_y_axis()
        print(f"   After Y rescaling - X: {layer.get_x_range()}, Y: {layer.get_y_range()}")
        
        # Test log scale
        print("\n4. Testing log scale:")
        # Create data suitable for log scale
        y_log_data = np.exp(x_data/2)
        worksheet.from_list(2, y_log_data.tolist(), "ExpSignal", "V", "Exponential data")
        
        plot2 = layer.add_xy_plot(worksheet, 0, 2)
        
        y_axis.set_scale('log10')
        print(f"   Y axis scale changed to: {y_axis.get_scale()}")
        
        # Reset to linear
        y_axis.set_scale('linear')
        print(f"   Y axis scale reset to: {y_axis.get_scale()}")
        
        # Test Z axis (for 3D plots - will show error for 2D)
        print("\n5. Testing Z axis (should fail for 2D plot):")
        try:
            z_axis = layer.get_z_axis()
            z_range = layer.get_z_range()
            print(f"   Z range: {z_range}")
        except Exception as e:
            print(f"   Expected error for Z axis in 2D plot: {e}")
        
        print("\n=== Test completed successfully! ===")
        
    except Exception as e:
        print(f"Error during test: {e}")
        import traceback
        traceback.print_exc()
        
    finally:
        # Clean up
        origin.close()

if __name__ == "__main__":
    test_axis_functions()
