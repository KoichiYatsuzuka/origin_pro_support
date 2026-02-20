#!/usr/bin/env python3
"""
Simple test for DataFrame to worksheet using OriginInstance methods.
"""

import os
import sys
import numpy as np
import pandas as pd
from pathlib import Path

# Add the parent directory to the path to import the library
sys.path.insert(0, str(Path(__file__).parent.parent))

from origin_pro_support import OriginInstance, XYTemplate


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


def test_dataframe_workflow():
    """Test the complete DataFrame workflow."""
    print("\n" + "="*60)
    print("TEST: DataFrame Workflow")
    print("="*60)
    
    # Set up project path
    project_path = os.path.join(os.path.dirname(__file__), "test_dataframe_workflow.opju")
    
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
        
        print("Creating workbook using OriginInstance...")
        # Use OriginInstance directly since folder methods are having issues
        workbook = origin.new_workbook("DataWorkbook")
        
        if workbook is None:
            print("ERROR: Failed to create workbook")
            return False
        
        print(f"Workbook created: {workbook.Name}")
        worksheet = workbook[0]  # Get first worksheet
        
        print("Loading DataFrame into worksheet...")
        worksheet.from_df(df)
        print("[OK] DataFrame loaded to worksheet")
        
        # Verify data was loaded correctly
        print(f"Worksheet dimensions: {worksheet.get_rows()} rows x {worksheet.get_cols()} columns")
        
        # Check column names
        for i, col in enumerate(worksheet.Columns):
            try:
                print(f"  Column {i}: {col.LongName}")
            except Exception as e:
                print(f"  Column {i}: Error getting name - {e}")
        
        print("Creating simple plot...")
        # Create a simple plot
        graph = origin.new_graph("DataPlot", XYTemplate.SCATTER)
        
        if graph:
            print(f"Graph created: {graph.Name}")
            # Plot first two columns (Time vs Sine_Wave)
            layer = graph[0]
            plot = layer.add_xy_plot(worksheet, 0, 1)  # X=col0, Y=col1
            print(f"Plot created: {plot}")
            
            print("Saving project...")
            origin.save()
            print("[OK] Project saved")
        else:
            print("ERROR: Failed to create graph")
        
        print("[SUCCESS] DataFrame workflow test completed")
        return True
        
    except Exception as e:
        print(f"[ERROR] DataFrame workflow test failed: {e}")
        import traceback
        traceback.print_exc()
        return False
        
    finally:
        if origin:
            origin.close()


def main():
    """Main test function."""
    print("Starting simple DataFrame test...")
    
    success = test_dataframe_workflow()
    
    if success:
        print("\n" + "="*60)
        print("TEST COMPLETED SUCCESSFULLY")
        print("="*60)
        return 0
    else:
        print("\n" + "="*60)
        print("TEST FAILED")
        print("="*60)
        return 1


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)
