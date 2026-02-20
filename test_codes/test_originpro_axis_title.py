"""
Test originpro.graph.Axis.title property to find correct way to get/set axis titles.
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
    import originpro
    from originpro.graph import Axis
except ImportError as e:
    print(f"Import error: {e}")
    print("Make sure Origin is installed and the OriginExt module is available")
    sys.exit(1)


def test_originpro_axis_title():
    """Test originpro.graph.Axis.title property."""
    
    # Define test file path
    test_file = os.path.join(os.path.dirname(__file__), "originpro_axis_title_test.opju")
    
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
        workbook = origin.new_workbook("OriginProTest")
        worksheet = workbook[0]
        worksheet.from_list(0, x_data, "X", "", "X values")
        worksheet.from_list(1, y_data, "Y", "", "Y values")
        
        # Create graph
        graph_page = origin.new_graph("OriginProTestGraph")
        figure_page = FigurePage(graph_page)
        
        # Plot data
        layer = figure_page.get_active_layer()
        plot = layer.add_xy_plot(worksheet, 0, 1)
        
        print("=== Testing originpro.graph.Axis.title ===\n")
        
        # Get the raw OriginExt GraphLayer object
        from origin_pro_support.layers import get_originext_graphlayer
        gl_obj = get_originext_graphlayer(layer)
        
        print(f"GraphLayer object: {gl_obj}")
        print(f"GraphLayer type: {type(gl_obj)}")
        
        # Try to access X axis using different methods
        print("\nTrying to access X axis:")
        
        # Method 1: Try XAxis property
        try:
            x_axis = gl_obj.XAxis
            print(f"Method 1 - XAxis: {x_axis}")
            print(f"Method 1 - Type: {type(x_axis)}")
            
            if hasattr(x_axis, 'title'):
                print("Method 1 - Has title property")
                try:
                    current_title = x_axis.title
                    print(f"Method 1 - Current title: {repr(current_title)}")
                except Exception as e:
                    print(f"Method 1 - Error getting title: {e}")
                
                try:
                    x_axis.title = "Test from Method 1"
                    print("Method 1 - Title set successfully")
                    new_title = x_axis.title
                    print(f"Method 1 - New title: {repr(new_title)}")
                except Exception as e:
                    print(f"Method 1 - Error setting title: {e}")
            else:
                print("Method 1 - No title property")
                
        except Exception as e:
            print(f"Method 1 failed: {e}")
        
        # Method 2: Try using originpro.graph.Axis directly
        try:
            print("\nMethod 2 - Using originpro.graph.Axis:")
            # Try to create Axis object from GraphLayer
            x_axis_originpro = originpro.graph.Axis(gl_obj.XAxis)
            print(f"Method 2 - Created Axis: {x_axis_originpro}")
            print(f"Method 2 - Type: {type(x_axis_originpro)}")
            
            if hasattr(x_axis_originpro, 'title'):
                print("Method 2 - Has title property")
                try:
                    current_title = x_axis_originpro.title
                    print(f"Method 2 - Current title: {repr(current_title)}")
                except Exception as e:
                    print(f"Method 2 - Error getting title: {e}")
                
                try:
                    x_axis_originpro.title = "Test from Method 2"
                    print("Method 2 - Title set successfully")
                    new_title = x_axis_originpro.title
                    print(f"Method 2 - New title: {repr(new_title)}")
                except Exception as e:
                    print(f"Method 2 - Error setting title: {e}")
            else:
                print("Method 2 - No title property")
                
        except Exception as e:
            print(f"Method 2 failed: {e}")
        
        # Method 3: Try using originpro.Axis directly
        try:
            print("\nMethod 3 - Using originpro.Axis:")
            x_axis_direct = originpro.Axis(gl_obj.XAxis)
            print(f"Method 3 - Created Axis: {x_axis_direct}")
            print(f"Method 3 - Type: {type(x_axis_direct)}")
            
            if hasattr(x_axis_direct, 'title'):
                print("Method 3 - Has title property")
                try:
                    current_title = x_axis_direct.title
                    print(f"Method 3 - Current title: {repr(current_title)}")
                except Exception as e:
                    print(f"Method 3 - Error getting title: {e}")
                
                try:
                    x_axis_direct.title = "Test from Method 3"
                    print("Method 3 - Title set successfully")
                    new_title = x_axis_direct.title
                    print(f"Method 3 - New title: {repr(new_title)}")
                except Exception as e:
                    print(f"Method 3 - Error setting title: {e}")
            else:
                print("Method 3 - No title property")
                
        except Exception as e:
            print(f"Method 3 failed: {e}")
        
        # Method 4: Check all available properties
        try:
            print("\nMethod 4 - Checking all GraphLayer properties:")
            
            print("Method 4 - All attributes (non-private):")
            for attr in dir(gl_obj):
                if not attr.startswith('_'):
                    try:
                        value = getattr(gl_obj, attr)
                        if 'Axis' in attr or 'axis' in attr.lower():
                            print(f"  AXIS-RELATED: {attr}: {type(value)} = {repr(value)}")
                        elif callable(value):
                            print(f"  METHOD: {attr}")
                        else:
                            print(f"  PROPERTY: {attr}: {type(value)}")
                    except Exception as e:
                        print(f"  {attr}: Error - {e}")
                        
        except Exception as e:
            print(f"Method 4 failed: {e}")
        
        # Method 5: Try to find axis through different approaches
        try:
            print("\nMethod 5 - Alternative axis access:")
            
            # Try to get axes through different methods
            methods_to_try = [
                'GetXAxis', 'get_xaxis', 'get_XAxis', 'xaxis', 'XAxis',
                'GetYAxis', 'get_yaxis', 'get_YAxis', 'yaxis', 'YAxis',
                'GetAxis', 'get_axis', 'Axes', 'axes'
            ]
            
            for method_name in methods_to_try:
                if hasattr(gl_obj, method_name):
                    try:
                        method = getattr(gl_obj, method_name)
                        if callable(method):
                            result = method()
                            print(f"  {method_name}(): {type(result)} = {repr(result)}")
                        else:
                            print(f"  {method_name}: {type(result)} = {repr(result)}")
                    except Exception as e:
                        print(f"  {method_name}: Error - {e}")
                        
        except Exception as e:
            print(f"Method 5 failed: {e}")
        
        # Method 6: Try using originpro directly with the graph page
        try:
            print("\nMethod 6 - Using originpro with graph page:")
            
            # Get the graph page object
            graph_page_obj = gl_obj.GetPage()
            print(f"Graph page: {graph_page_obj}")
            print(f"Graph page type: {type(graph_page_obj)}")
            
            # Try to get axis from graph page
            if hasattr(graph_page_obj, 'Layers'):
                layers = graph_page_obj.Layers
                print(f"Layers: {layers}")
                
                if layers and len(layers) > 0:
                    layer0 = layers(0)  # Try both indexing methods
                    print(f"Layer 0: {layer0}")
                    
                    # Try to get axis from layer
                    for attr in dir(layer0):
                        if 'Axis' in attr or 'axis' in attr.lower():
                            try:
                                value = getattr(layer0, attr)
                                print(f"  Layer axis attr {attr}: {type(value)} = {repr(value)}")
                            except Exception as e:
                                print(f"  Layer axis attr {attr}: Error - {e}")
                                
        except Exception as e:
            print(f"Method 6 failed: {e}")
        
        # Method 7: Try LabTalk to get axis object
        try:
            print("\nMethod 7 - LabTalk axis access:")
            
            # Try to get axis using LabTalk
            cmd = "XAxis"
            result = gl_obj.Execute(cmd)
            print(f"LabTalk XAxis result: {result} (type: {type(result)})")
            
            # Try different LabTalk commands
            commands = ["get XAxis", "layer.xaxis", "layer -xaxis"]
            for cmd in commands:
                try:
                    result = gl_obj.Execute(cmd)
                    print(f"LabTalk '{cmd}': {result} (type: {type(result)})")
                except Exception as e:
                    print(f"LabTalk '{cmd}': Error - {e}")
                    
        except Exception as e:
            print(f"Method 7 failed: {e}")
        
        # Method 8: Try using originpro.graph.GLayer
        try:
            print("\nMethod 8 - Using originpro.graph.GLayer:")
            
            # Create GLayer from our GraphLayer
            glayer = originpro.graph.GLayer(gl_obj)
            print(f"Created GLayer: {glayer}")
            print(f"GLayer type: {type(glayer)}")
            
            # Check GLayer attributes for axis
            print("GLayer axis-related attributes:")
            for attr in dir(glayer):
                if 'Axis' in attr or 'axis' in attr.lower():
                    try:
                        value = getattr(glayer, attr)
                        print(f"  {attr}: {type(value)} = {repr(value)}")
                    except Exception as e:
                        print(f"  {attr}: Error - {e}")
            
            # Try common axis property names
            axis_props = ['XAxis', 'xaxis', 'Xaxis', 'x_axis', 'X_axis']
            for prop in axis_props:
                if hasattr(glayer, prop):
                    try:
                        axis = getattr(glayer, prop)
                        print(f"  Found axis {prop}: {type(axis)} = {repr(axis)}")
                        
                        # Check if this axis has title property
                        if hasattr(axis, 'title'):
                            print(f"    Axis has title property")
                            try:
                                current_title = axis.title
                                print(f"    Current title: {repr(current_title)}")
                            except Exception as e:
                                print(f"    Error getting title: {e}")
                            
                            try:
                                axis.title = "Test from GLayer"
                                print(f"    Title set successfully")
                                new_title = axis.title
                                print(f"    New title: {repr(new_title)}")
                            except Exception as e:
                                print(f"    Error setting title: {e}")
                        else:
                            print(f"    Axis has no title property")
                    except Exception as e:
                        print(f"  Error accessing {prop}: {e}")
                        
        except Exception as e:
            print(f"Method 8 failed: {e}")
        
        # Method 9: Try creating Axis directly with GLayer
        try:
            print("\nMethod 9 - Creating Axis from GLayer:")
            
            glayer = originpro.graph.GLayer(gl_obj)
            
            # Try to create Axis object with lowercase specs
            axis_specs = ['x', 'y', 'z']
            for spec in axis_specs:
                try:
                    axis = originpro.graph.Axis(glayer, spec)
                    print(f"Created {spec} axis: {type(axis)}")
                    
                    # Check if this axis has title property
                    if hasattr(axis, 'title'):
                        print(f"  {spec} axis has title property")
                        try:
                            current_title = axis.title
                            print(f"  Current {spec} title: {repr(current_title)}")
                        except Exception as e:
                            print(f"  Error getting {spec} title: {e}")
                        
                        try:
                            test_title = f"Test {spec.upper()} Title"
                            axis.title = test_title
                            print(f"  {spec} title set successfully")
                            new_title = axis.title
                            print(f"  New {spec} title: {repr(new_title)}")
                            
                            if new_title == test_title:
                                print(f"  ✓ SUCCESS: {spec} axis title works perfectly!")
                                return  # Success! We can stop here
                            else:
                                print(f"  ✗ FAILED: {spec} title not set correctly")
                        except Exception as e:
                            print(f"  Error setting {spec} title: {e}")
                    else:
                        print(f"  {spec} axis has no title property")
                        
                except Exception as e:
                    print(f"Error creating {spec} axis: {e}")
                    
        except Exception as e:
            print(f"Method 9 failed: {e}")
        
        # Method 10: Try using GLayer.axis() method
        try:
            print("\nMethod 10 - Using GLayer.axis() method:")
            
            glayer = originpro.graph.GLayer(gl_obj)
            
            # Try the axis method
            if hasattr(glayer, 'axis'):
                print("GLayer has axis method")
                
                for spec in ['x', 'y', 'z']:
                    try:
                        axis = glayer.axis(spec)
                        print(f"Got {spec} axis via glayer.axis(): {type(axis)}")
                        
                        if hasattr(axis, 'title'):
                            print(f"  {spec} axis has title property")
                            try:
                                current_title = axis.title
                                print(f"  Current {spec} title: {repr(current_title)}")
                                
                                test_title = f"GLayer {spec.upper()} Title"
                                axis.title = test_title
                                new_title = axis.title
                                print(f"  New {spec} title: {repr(new_title)}")
                                
                                if new_title == test_title:
                                    print(f"  ✓ SUCCESS: GLayer.axis() {spec} title works!")
                                    return
                            except Exception as e:
                                print(f"  Error with {spec} title: {e}")
                        else:
                            print(f"  {spec} axis has no title property")
                            
                    except Exception as e:
                        print(f"Error getting {spec} axis via glayer.axis(): {e}")
            else:
                print("GLayer does not have axis method")
                
        except Exception as e:
            print(f"Method 10 failed: {e}")
        
        print(f"\nTest completed. File saved to: {test_file}")
        
    except Exception as e:
        print(f"Error during test: {e}")
        import traceback
        traceback.print_exc()
        
    finally:
        # Clean up
        origin.close()


if __name__ == "__main__":
    test_originpro_axis_title()
