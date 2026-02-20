#%%
"""
Simple example demonstrating FigurePage usage.

This example shows the key features of the FigurePage class
with enum-based controls and hierarchical structure.
"""
import sys
import os
import numpy as np
import pandas as pd

# Add the parent directory to the path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    import origin_pro_support as ops
    from origin_pro_support import FigurePage, PlotType, ColorMap, GroupMode, OriginInstance
except ImportError as e:
    print(f"Import error: {e}")
    print("Make sure Origin is installed and available")
    sys.exit(1)


def main():
    """Main example function."""
    print("FigurePage Example - Similar to Origin Sample #5")
    print("=" * 50)
    
    # Create Origin instance with full path
    import os
    current_dir = os.path.dirname(os.path.abspath(__file__))
    project_path = os.path.join(current_dir, "figure_page_test.opju")
    
    # Delete existing file if it exists
    if os.path.exists(project_path):
        print(f"Removing existing file: {project_path}")
        os.remove(project_path)
    
    origin = None
    try:
        print("Creating Origin instance...")
        origin = OriginInstance(project_path)
        origin.set_show(True)
        
        print("Creating sample data...")
        # Create sample data
        worksheet = origin.new_sheet('w', 'GroupData')
        
        # Create data similar to Origin's Group.dat
        x_data = np.linspace(0, 10, 30)
        data = {
            'Time': x_data,
            'Group1': np.sin(x_data) + np.random.normal(0, 0.1, len(x_data)),
            'Group2': np.cos(x_data) + np.random.normal(0, 0.1, len(x_data)),
            'Group3': np.sin(x_data * 0.5) + np.random.normal(0, 0.1, len(x_data))
        }
        
        df = pd.DataFrame(data)
        worksheet.from_df(df)
        print("[OK] Data loaded to worksheet")
        
        print("Creating FigurePage...")
        # Create FigurePage using OriginInstance (recommended approach)
        graph_page = origin.new_graph('GroupedPlot', template='scatter')
        figure = FigurePage(graph_page, template='scatter')
        
        print("Creating grouped plot...")
        # Create grouped plot similar to Origin Sample #5
        # Using enum types instead of literal values
        plot = figure.create_grouped_plot(
            worksheet=worksheet,
            x_col=0,  # Time column
            color_map=ColorMap.CANDY,  # Enum instead of 'Candy' string
            shape_list=[3, 2, 1],  # Different shapes for each group
            plot_type=PlotType.LINE_SYMBOL  # Enum instead of 202
        )
        
        print(f"[OK] Created grouped plot: {plot.Name}")
        print("Using enum types for all parameters")
        
        # Demonstrate hierarchical structure
        layer = figure.get_active_layer()
        print(f"Active layer: {layer.Name}")
        print(f"Number of data plots: {len(layer.get_data_plots())}")
        
        # Access individual plots and configure them
        for i, data_plot in enumerate(layer):
            print(f"Plot {i}: {data_plot.Name}")
            print(f"  Color map: {data_plot.color_map}")
            print(f"  Shape list: {data_plot.shape_list}")
        
        # Customize the page
        figure.set_page_size(8, 6)  # 8x6 inches
        print("Set page size to 8x6 inches")
        
        # Export preview
        if figure.export_preview('example_plot.png'):
            print("Exported preview to example_plot.png")
        
        print("\nExample completed successfully!")
        print("Key features demonstrated:")
        print("- FigurePage with enum-based controls")
        print("- Hierarchical structure (FigurePage -> GraphLayer -> DataPlot)")
        print("- Grouped plotting with ColorMap and PlotType enums")
        print("- Object-oriented approach vs functional")
        
        input("\nPress Enter to close Origin...")
        
    except Exception as e:
        print(f"[ERROR] Example failed with error: {e}")
        import traceback
        traceback.print_exc()
        
    finally:
        # Always save and close if origin instance exists
        if origin is not None:
            try:
                print("Saving Origin file...")
                origin.save()
                print("[OK] File saved successfully")
            except Exception as save_error:
                print(f"[ERROR] Failed to save: {save_error}")
            
            try:
                print("Closing Origin instance...")
                origin.close()
                print("[OK] Origin instance closed successfully")
            except Exception as close_error:
                print(f"[ERROR] Failed to close: {close_error}")

    print("Example execution finished.")


if __name__ == "__main__":
    main()
