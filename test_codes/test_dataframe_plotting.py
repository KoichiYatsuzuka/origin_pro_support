#!/usr/bin/env python3
"""
Test script for creating worksheets from pandas DataFrames and plotting them.

This test demonstrates the proper hierarchical approach:
OriginInstance -> Folder -> Page -> Layer -> Plot
"""

import os
import sys
import numpy as np
import pandas as pd
from pathlib import Path

# Add the parent directory to the path to import the library
sys.path.insert(0, str(Path(__file__).parent.parent))

from origin_pro_support import (
    OriginInstance, 
    OriginNameConflictError,
    ColorMap, 
    PlotType, 
    GroupMode,
    FigurePage
)


def create_sample_data():
    """Create sample data for testing."""
    print("Creating sample data...")
    
    # Create time series data
    np.random.seed(42)  # For reproducible results
    time_points = np.linspace(0, 10, 50)
    
    # Create multiple datasets with different characteristics
    data = {
        'Time': time_points,
        'Sine_Wave': np.sin(time_points) + np.random.normal(0, 0.1, len(time_points)),
        'Cosine_Wave': np.cos(time_points) + np.random.normal(0, 0.1, len(time_points)),
        'Damped_Oscillation': np.exp(-time_points/5) * np.sin(2*time_points) + np.random.normal(0, 0.05, len(time_points)),
        'Linear_Trend': 0.5 * time_points + np.random.normal(0, 0.2, len(time_points))
    }
    
    df = pd.DataFrame(data)
    print(f"Created DataFrame with shape: {df.shape}")
    print(f"Columns: {list(df.columns)}")
    
    return df


def test_dataframe_to_worksheet():
    """Test creating a worksheet from pandas DataFrame."""
    print("\n" + "="*60)
    print("TEST: DataFrame to Worksheet")
    print("="*60)
    
    # Set up project path
    project_path = os.path.join(os.path.dirname(__file__), "test_dataframe_plotting.opju")
    
    # Delete existing file if it exists
    if os.path.exists(project_path):
        print(f"Removing existing file: {project_path}")
        os.remove(project_path)
    
    origin = None
    try:
        print("Creating Origin instance...")
        origin = OriginInstance(project_path)
        origin.set_show(True)
        
        # Create sample data
        df = create_sample_data()
        
        print("Creating workbook using proper hierarchy...")
        # Use proper hierarchy: OriginInstance -> new_workbook
        workbook = origin.new_workbook("DataWorkbook")
        worksheet = workbook[0]  # Get first worksheet
        
        print("Loading DataFrame into worksheet...")
        worksheet.from_df(df)
        print("[OK] DataFrame loaded to worksheet")
        
        # Verify data was loaded correctly
        print(f"Worksheet dimensions: {worksheet.get_rows()} rows x {worksheet.get_cols()} columns")
        
        # Check column names
        for i, col in enumerate(worksheet.Columns):
            print(f"  Column {i}: {col.LongName}")
        
        print("[SUCCESS] DataFrame to worksheet test completed")
        return worksheet
        
    except Exception as e:
        print(f"[ERROR] DataFrame to worksheet test failed: {e}")
        return None
        
    finally:
        if origin:
            origin.close()


def test_worksheet_plotting(worksheet):
    """Test plotting data from worksheet."""
    print("\n" + "="*60)
    print("TEST: Worksheet Plotting")
    print("="*60)
    
    # Set up project path
    project_path = os.path.join(os.path.dirname(__file__), "test_dataframe_plotting.opju")
    
    origin = None
    try:
        print("Opening existing Origin instance...")
        origin = OriginInstance(project_path)
        origin.set_show(True)
        
        print("Creating graph page...")
        # Create graph page using proper hierarchy
        graph_page = origin.new_graph("DataPlot", template="scatter")
        
        print("Creating FigurePage for enhanced plotting...")
        figure = FigurePage(graph_page, template="scatter")
        
        print("Plotting multiple Y series against Time...")
        # Plot all data series against Time (column 0)
        plots = figure.plot_multiple_series(
            worksheet=worksheet,
            x_col=0,  # Time column
            y_cols=[1, 2, 3, 4],  # All data columns
            plot_type=PlotType.LINE_SYMBOL,
            color_map=ColorMap.VIRIDIS,
            group_mode=GroupMode.INDEPENDENT
        )
        
        print(f"Created {len(plots)} data plots")
        
        print("Customizing plot appearance...")
        # Get the active layer for further customization
        layer = figure.get_active_layer()
        
        # Set page size and colors
        figure.set_page_size(8, 6, units=1)  # 8x6 cm
        figure.set_colors(base_color=1, grad_color=15)
        
        print("Exporting preview...")
        preview_path = os.path.join(os.path.dirname(__file__), "plot_preview.png")
        success = figure.export_preview(preview_path)
        if success:
            print(f"[OK] Preview exported to: {preview_path}")
        else:
            print("[WARNING] Preview export failed")
        
        print("[SUCCESS] Worksheet plotting test completed")
        
    except Exception as e:
        print(f"[ERROR] Worksheet plotting test failed: {e}")
        
    finally:
        if origin:
            origin.close()


def test_grouped_plotting(worksheet):
    """Test creating grouped plots."""
    print("\n" + "="*60)
    print("TEST: Grouped Plotting")
    print("="*60)
    
    # Set up project path
    project_path = os.path.join(os.path.dirname(__file__), "test_dataframe_plotting.opju")
    
    origin = None
    try:
        print("Opening existing Origin instance...")
        origin = OriginInstance(project_path)
        origin.set_show(True)
        
        print("Creating grouped plot graph...")
        # Create graph page using proper hierarchy
        graph_page = origin.new_graph("GroupedPlot", template="scatter")
        figure = FigurePage(graph_page, template="scatter")
        
        print("Creating grouped plot with automatic coloring...")
        # Create grouped plot - all Y columns against X with automatic grouping
        plot = figure.create_grouped_plot(
            worksheet=worksheet,
            x_col=0,  # Time column
            color_map=ColorMap.CANDY,
            plot_type=PlotType.LINE_SYMBOL
        )
        
        print("Customizing grouped plot...")
        # Customize the grouped plot
        plot.shape_list = [1, 2, 3, 4]  # Different shapes for each series
        
        print("Exporting grouped plot preview...")
        preview_path = os.path.join(os.path.dirname(__file__), "grouped_plot_preview.png")
        success = figure.export_preview(preview_path)
        if success:
            print(f"[OK] Grouped plot preview exported to: {preview_path}")
        else:
            print("[WARNING] Grouped plot preview export failed")
        
        print("[SUCCESS] Grouped plotting test completed")
        
    except Exception as e:
        print(f"[ERROR] Grouped plotting test failed: {e}")
        
    finally:
        if origin:
            origin.close()


def main():
    """Main test function."""
    print("Starting DataFrame plotting tests...")
    print(f"Working directory: {os.getcwd()}")
    print(f"Test directory: {os.path.dirname(__file__)}")
    
    try:
        # Test 1: Create worksheet from DataFrame
        worksheet = test_dataframe_to_worksheet()
        
        if worksheet is not None:
            # Test 2: Plot data from worksheet
            test_worksheet_plotting(worksheet)
            
            # Test 3: Create grouped plots
            test_grouped_plotting(worksheet)
        else:
            print("[ERROR] Cannot proceed with plotting tests - worksheet creation failed")
            
    except Exception as e:
        print(f"[FATAL ERROR] Test suite failed: {e}")
        return 1
    
    print("\n" + "="*60)
    print("ALL TESTS COMPLETED")
    print("="*60)
    return 0


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)
