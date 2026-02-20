"""
Debug script for axis title setting.
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


def debug_axis_title():
    """Debug axis title setting with various approaches."""
    
    # Define test file path
    test_file = os.path.join(os.path.dirname(__file__), "axis_title_debug.opju")
    
    # Clean up existing file
    if os.path.exists(test_file):
        os.remove(test_file)
    
    # Create Origin instance
    origin = OriginInstance(test_file)
    
    try:
        # Create simple test data
        x_data = [1, 2, 3, 4, 5]
        y_data = [1, 4, 9, 16, 25]
        
        # Create workbook and add data
        workbook = origin.new_workbook("DebugData")
        worksheet = workbook[0]
        worksheet.from_list(0, x_data, "X", "", "X values")
        worksheet.from_list(1, y_data, "Y", "", "Y values")
        
        # Create graph
        graph_page = origin.new_graph("DebugGraph")
        figure_page = FigurePage(graph_page)
        
        # Plot data
        layer = figure_page.get_active_layer()
        plot = layer.add_xy_plot(worksheet, 0, 1)
        
        print("=== Axis Title Debug Test ===\n")
        
        # Get the raw OriginExt object for direct LabTalk testing
        from origin_pro_support.layers import get_originext_graphlayer
        gl_obj = get_originext_graphlayer(layer)
        
        print("Testing various LabTalk commands for axis title:")
        
        # Test different LabTalk commands
        commands = [
            'X.label.text$ = "Test X Title"',
            'layer -xlabel "Test X Title 2"',
            'X.label$ = "Test X Title 3"',
            'win -t "Test X Title 4" xtitle',
            'label -xb "Test X Title 5"',  # Bottom X axis
        ]
        
        for i, cmd in enumerate(commands, 1):
            try:
                print(f"\n{i}. Trying: {cmd}")
                result = gl_obj.Execute(cmd)
                print(f"   Result: {result}")
                
                # Try to get the title back
                try:
                    title = gl_obj.GetStrProp("X.label.text")
                    print(f"   Current title: '{title}'")
                except Exception as e:
                    print(f"   Could not get title: {e}")
                    
            except Exception as e:
                print(f"   Failed: {e}")
        
        print("\n=== Testing with Axis class ===")
        
        # Test our Axis class
        x_axis = layer.get_x_axis()
        
        print(f"Current title: '{x_axis.get_title()}'")
        
        print("Setting title with Axis class...")
        x_axis.set_title("Axis Class Test Title")
        
        print(f"Title after setting: '{x_axis.get_title()}'")
        
        print("\n=== Debug completed ===")
        
    except Exception as e:
        print(f"Error during debug: {e}")
        import traceback
        traceback.print_exc()
        
    finally:
        # Clean up
        origin.close()


if __name__ == "__main__":
    debug_axis_title()
