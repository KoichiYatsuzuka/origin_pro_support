#!/usr/bin/env python3
"""
Debug test for graph creation.
"""

import os
import sys
from pathlib import Path

# Add the parent directory to the path to import the library
sys.path.insert(0, str(Path(__file__).parent.parent))

from origin_pro_support import OriginInstance, XYTemplate


def test_graph_creation():
    """Test graph creation with debugging."""
    print("Testing graph creation...")
    
    # Set up project path
    project_path = os.path.join(os.path.dirname(__file__), "test_graph_debug.opju")
    
    # Delete existing file if it exists
    if os.path.exists(project_path):
        print(f"Removing existing file: {project_path}")
        os.remove(project_path)
    
    origin = None
    try:
        print("Creating Origin instance...")
        origin = OriginInstance(project_path)
        origin.set_show(True)
        
        print("Testing different XYTemplate values...")
        
        # Test 1: Basic scatter plot
        print("\n1. Testing SCATTER template...")
        graph1 = origin.new_graph("ScatterPlot", XYTemplate.SCATTER)
        print(f"   Result: {graph1}")
        
        # Test 2: Line plot
        print("\n2. Testing LINE template...")
        graph2 = origin.new_graph("LinePlot", XYTemplate.LINE)
        print(f"   Result: {graph2}")
        
        # Test 3: Direct string template (for comparison)
        print("\n3. Testing direct string template...")
        try:
            # This should work if we pass the string directly
            root_folder = origin.get_root_dir()
            graph3 = root_folder.create_graph("DirectString", "scatter")
            print(f"   Result: {graph3}")
        except Exception as e:
            print(f"   Error: {e}")
        
        print("\nGraph creation test completed.")
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        
    finally:
        if origin:
            print("Closing Origin instance...")
            origin.close(False)


if __name__ == "__main__":
    test_graph_creation()
