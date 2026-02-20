"""
Test to verify if axis titles are actually saved in the Origin file.
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


def test_axis_title_save():
    """Test if axis titles are saved in Origin file."""
    
    # Define test file path
    test_file = os.path.join(os.path.dirname(__file__), "axis_title_save_test.opju")
    
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
        workbook = origin.new_workbook("TitleTest")
        worksheet = workbook[0]
        worksheet.from_list(0, x_data, "X", "", "X values")
        worksheet.from_list(1, y_data, "Y", "", "Y values")
        
        # Create graph
        graph_page = origin.new_graph("TitleTestGraph")
        figure_page = FigurePage(graph_page)
        
        # Plot data
        layer = figure_page.get_active_layer()
        plot = layer.add_xy_plot(worksheet, 0, 1)
        
        print("=== Axis Title Save Test ===\n")
        
        # Get axis objects
        x_axis = layer.get_x_axis()
        y_axis = layer.get_y_axis()
        
        # Set titles
        print("Setting axis titles...")
        x_title = "Test X Axis Title"
        y_title = "Test Y Axis Title"
        
        x_axis.set_title(x_title)
        y_axis.set_title(y_title)
        
        # Try to get titles immediately
        print(f"Immediate get - X: '{x_axis.get_title()}'")
        print(f"Immediate get - Y: '{y_axis.get_title()}'")
        
        # Save the file
        print(f"\nSaving file to: {test_file}")
        origin.save()
        
        # Wait a moment
        import time
        time.sleep(1)
        
        # Try to get titles after saving
        print(f"After save - X: '{x_axis.get_title()}'")
        print(f"After save - Y: '{y_axis.get_title()}'")
        
        # Close and reopen to test persistence
        print("\nClosing and reopening file...")
        origin.close()
        
        # Reopen the file
        origin2 = OriginInstance(test_file)
        
        try:
            # Get the graph page again
            graph_page2 = origin2.pages[0]  # First page should be our graph
            layer2 = graph_page2.get_active_layer()
            
            x_axis2 = layer2.get_x_axis()
            y_axis2 = layer2.get_y_axis()
            
            print(f"After reopen - X: '{x_axis2.get_title()}'")
            print(f"After reopen - Y: '{y_axis2.get_title()}'")
            
            # Try setting titles again
            print("\nSetting titles again...")
            x_axis2.set_title("Reopened X Title")
            y_axis2.set_title("Reopened Y Title")
            
            print(f"After reset - X: '{x_axis2.get_title()}'")
            print(f"After reset - Y: '{y_axis2.get_title()}'")
            
        finally:
            origin2.close()
        
        print(f"\nTest completed. Check the file: {test_file}")
        print("Open it in Origin to see if titles are visible in the GUI.")
        
    except Exception as e:
        print(f"Error during test: {e}")
        import traceback
        traceback.print_exc()
        
    finally:
        # Clean up
        try:
            origin.close()
        except:
            pass


if __name__ == "__main__":
    test_axis_title_save()
