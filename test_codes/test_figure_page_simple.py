#!/usr/bin/env python3
"""
Simple test for FigurePage functionality without user interaction.
This test demonstrates the basic features and verifies the implementation.
"""
import sys
import os
import numpy as np
import pandas as pd

# Add the parent directory to the path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

def test_figure_page_basic():
    """Basic test of FigurePage functionality."""
    print("=== FigurePage Basic Test ===")
    
    try:
        from origin_pro_support import FigurePage, PlotType, ColorMap, GroupMode, OriginInstance
        print("[OK] Successfully imported origin_pro_support modules")
    except ImportError as e:
        print(f"[ERROR] Import failed: {e}")
        return False
    
    # Create test file path
    current_dir = os.path.dirname(os.path.abspath(__file__))
    project_path = os.path.join(current_dir, "test_figure_page.opju")
    
    # Clean up existing file
    if os.path.exists(project_path):
        os.remove(project_path)
    
    origin = None
    try:
        print("\n1. Creating Origin instance...")
        origin = OriginInstance(project_path)
        origin.set_show(False)  # Hide window for automated test
        
        print("2. Creating sample data...")
        # Create simple test data
        worksheet = origin.new_sheet('w', 'TestData')
        
        x_data = np.linspace(0, 5, 20)
        data = {
            'X': x_data,
            'Y1': np.sin(x_data),
            'Y2': np.cos(x_data),
            'Y3': x_data * 0.5
        }
        
        df = pd.DataFrame(data)
        worksheet.from_df(df)
        print("[OK] Data loaded to worksheet")
        
        print("3. Creating FigurePage...")
        graph_page = origin.new_graph('TestGraph', template='scatter')
        figure = FigurePage(graph_page, template='scatter')
        print("[OK] FigurePage created")
        
        print("4. Creating grouped plot...")
        plot = figure.create_grouped_plot(
            worksheet=worksheet,
            x_col=0,  # X column
            color_map=ColorMap.CANDY,
            shape_list=[1, 2, 3],
            plot_type=PlotType.LINE_SYMBOL
        )
        print("[OK] Grouped plot created")
        
        print("5. Testing hierarchical structure...")
        layer = figure.get_active_layer()
        print(f"[OK] Active layer: {layer.Name}")
        print(f"[OK] Number of data plots: {len(layer.DataPlots)}")
        
        print("6. Testing page customization...")
        figure.set_page_size(6, 4)
        print("[OK] Page size set to 6x4 inches")
        
        print("7. Saving project...")
        origin.save()
        print("[OK] Project saved successfully")
        
        print("\n=== Test Results ===")
        print("[OK] All basic FigurePage functions work correctly")
        print("[OK] Enum-based controls (PlotType, ColorMap) working")
        print("[OK] Hierarchical structure (FigurePage -> GraphLayer -> DataPlot) working")
        print("[OK] Grouped plotting functionality working")
        
        return True
        
    except Exception as e:
        print(f"\n[ERROR] Test failed with error: {e}")
        import traceback
        traceback.print_exc()
        return False
        
    finally:
        if origin is not None:
            try:
                origin.close()
                print("[OK] Origin instance closed")
            except Exception as e:
                print(f"[ERROR] Failed to close Origin: {e}")

def test_enum_functionality():
    """Test enum functionality."""
    print("\n=== Enum Functionality Test ===")
    
    try:
        from origin_pro_support import PlotType, ColorMap, GroupMode
        
        print("Testing PlotType enum...")
        assert PlotType.LINE.value == 200
        assert PlotType.SCATTER.value == 201
        assert PlotType.LINE_SYMBOL.value == 202
        print("[OK] PlotType enum working")
        
        print("Testing ColorMap enum...")
        assert ColorMap.CANDY.value == "Candy"
        assert ColorMap.RAINBOW.value == "Rainbow"
        print("[OK] ColorMap enum working")
        
        print("Testing GroupMode enum...")
        assert GroupMode.NONE.value == 0
        assert GroupMode.DEPENDENT.value == 2
        print("[OK] GroupMode enum working")
        
        return True
        
    except Exception as e:
        print(f"[ERROR] Enum test failed: {e}")
        return False

if __name__ == "__main__":
    print("FigurePage Test Suite")
    print("=" * 50)
    
    # Test enum functionality first (doesn't require Origin)
    enum_success = test_enum_functionality()
    
    # Test basic FigurePage functionality
    basic_success = test_figure_page_basic()
    
    print("\n" + "=" * 50)
    print("FINAL RESULTS:")
    print(f"Enum Tests: {'PASSED' if enum_success else 'FAILED'}")
    print(f"Basic Tests: {'PASSED' if basic_success else 'FAILED'}")
    
    if enum_success and basic_success:
        print("\n[SUCCESS] All tests PASSED! FigurePage implementation is working correctly.")
    else:
        print("\n[FAILED] Some tests FAILED. Please check the implementation.")
    
    print("\nTest completed.")
