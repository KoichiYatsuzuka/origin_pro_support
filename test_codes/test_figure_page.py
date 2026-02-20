"""
Test script demonstrating FigurePage plotting functionality.

This script shows how to use the new FigurePage class with enum-based controls
to create plots similar to the Origin Sample #5 example.
"""
import os
import sys
import numpy as np
import pandas as pd

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


def create_sample_data():
    """Create sample data similar to the Origin Group.dat example."""
    # Create sample data with multiple Y columns
    np.random.seed(42)
    x_data = np.linspace(0, 10, 50)
    
    # Create multiple Y datasets with different characteristics
    y1 = np.sin(x_data) + np.random.normal(0, 0.1, len(x_data))
    y2 = np.cos(x_data) + np.random.normal(0, 0.1, len(x_data))
    y3 = np.sin(x_data * 0.5) + np.random.normal(0, 0.1, len(x_data))
    
    # Create DataFrame
    df = pd.DataFrame({
        'Time': x_data,
        'Signal_A': y1,
        'Signal_B': y2,
        'Signal_C': y3
    })
    
    return df


def test_basic_plotting():
    """Test basic plotting functionality with FigurePage."""
    print("Testing basic FigurePage plotting...")
    
    try:
        # Create Origin instance
        origin = OriginInstance("test_figure_page.opju", create_new_if_not_exist=True)
        origin.set_show(True)
        
        # Create a new worksheet with sample data
        worksheet = origin.new_sheet('w', 'SampleData')
        df = create_sample_data()
        worksheet.from_df(df)
        
        # Create a new FigurePage
        figure = FigurePage.create_new('MyPlot', template='scatter')
        
        # Test 1: Basic XY plot with enum types
        print("Creating basic XY plot...")
        plot = figure.plot_xy_data(
            worksheet=worksheet,
            x_col=0,  # Time column
            y_col=1,  # Signal_A column
            plot_type=PlotType.LINE_SYMBOL,
            color_map=ColorMap.CANDY
        )
        print(f"Created plot: {plot.Name}")
        
        # Test 2: Multiple series plot
        print("Creating multiple series plot...")
        plots = figure.plot_multiple_series(
            worksheet=worksheet,
            x_col=0,  # Time column
            y_cols=[1, 2, 3],  # All signal columns
            plot_type=PlotType.SCATTER,
            color_map=ColorMap.RAINBOW,
            group_mode=GroupMode.DEPENDENT
        )
        print(f"Created {len(plots)} plots")
        
        # Test 3: Grouped plot (similar to Origin Sample #5)
        print("Creating grouped plot...")
        grouped_plot = figure.create_grouped_plot(
            worksheet=worksheet,
            x_col=0,
            color_map=ColorMap.VIRIDIS,
            shape_list=[3, 2, 1],  # Different shapes for each series
            plot_type=PlotType.LINE_SYMBOL
        )
        print(f"Created grouped plot: {grouped_plot.Name}")
        
        # Test 4: Page customization
        print("Customizing page...")
        figure.set_page_size(8, 6, units=0)  # 8x6 inches
        figure.set_colors(base_color=1, grad_color=2)
        
        # Export preview
        preview_path = "test_plot_preview.png"
        if figure.export_preview(preview_path):
            print(f"Exported preview to: {preview_path}")
        
        print("Basic plotting test completed successfully!")
        
        # Keep Origin open for inspection
        input("Press Enter to close Origin and continue...")
        
        origin.close()
        
    except Exception as e:
        print(f"Error in basic plotting test: {e}")
        if 'origin' in locals():
            origin.close(False)
        raise


def test_advanced_features():
    """Test advanced features like multi-layer plots."""
    print("Testing advanced FigurePage features...")
    
    try:
        # Create Origin instance
        origin = OriginInstance("test_figure_page_advanced.opju", create_new_if_not_exist=True)
        origin.set_show(True)
        
        # Create sample data
        worksheet = origin.new_sheet('w', 'AdvancedData')
        df = create_sample_data()
        worksheet.from_df(df)
        
        # Create FigurePage
        figure = FigurePage.create_new('AdvancedPlot', template='scatter')
        
        # Test multi-layer plotting
        print("Creating multi-layer plot...")
        
        # Layer 0: Original data
        plot1 = figure.plot_xy_data(
            worksheet=worksheet,
            x_col=0,
            y_col=1,
            plot_type=PlotType.LINE,
            layer_index=0,
            color_map=ColorMap.OCEAN
        )
        
        # Layer 1: Processed data (add a new layer)
        layer1 = figure.add_layer("Processed Data")
        plot2 = layer1.add_xy_plot(
            worksheet=worksheet,
            x_col=0,
            y_col=2,
            plot_type=PlotType.COLUMN
        )
        plot2.color_map = ColorMap.TERRAIN
        
        print(f"Created plots on layers: {plot1.Name} (layer 0), {plot2.Name} (layer 1)")
        
        # Test different plot types
        print("Testing different plot types...")
        
        # Create a new worksheet for histogram
        hist_worksheet = origin.new_sheet('w', 'HistogramData')
        hist_data = np.random.normal(0, 1, 100)
        hist_worksheet.from_list(0, hist_data, 'Random Data', 'units', 'Sample data')
        
        # Create histogram
        hist_figure = FigurePage.create_new('Histogram', template='histogram')
        hist_plot = hist_figure.plot_xy_data(
            worksheet=hist_worksheet,
            x_col=0,
            y_col=0,
            plot_type=PlotType.HISTOGRAM
        )
        
        print("Advanced features test completed successfully!")
        
        # Keep Origin open for inspection
        input("Press Enter to close Origin and continue...")
        
        origin.close()
        
    except Exception as e:
        print(f"Error in advanced features test: {e}")
        if 'origin' in locals():
            origin.close(False)
        raise


def demonstrate_object_oriented_approach():
    """Demonstrate the object-oriented approach vs functional approach."""
    print("Demonstrating object-oriented plotting approach...")
    
    try:
        # Create Origin instance
        origin = OriginInstance("test_oo_approach.opju", create_new_if_not_exist=True)
        origin.set_show(True)
        
        # Create data
        worksheet = origin.new_sheet('w', 'OOData')
        df = create_sample_data()
        worksheet.from_df(df)
        
        # Object-oriented approach with hierarchical structure
        print("Using object-oriented approach with enums...")
        
        # Create figure
        figure = FigurePage.create_new('OOPlot', 'scatter')
        
        # Access layers through the page
        layer = figure.get_active_layer()
        
        # Create plots using enum types
        plots = []
        for i, col in enumerate([1, 2, 3]):
            plot = layer.add_xy_plot(
                worksheet=worksheet,
                x_col=0,
                y_col=col,
                plot_type=PlotType.LINE_SYMBOL
            )
            
            # Configure plot using enum properties
            plot.color_map = list(ColorMap)[i]  # Different color for each plot
            plot.shape_list = [3 - i]  # Different shapes
            
            plots.append(plot)
        
        # Group plots using enum
        layer.group_plots(GroupMode.DEPENDENT)
        layer.rescale()
        
        # Configure page using object methods
        figure.set_page_size(10, 8)
        
        print(f"Created {len(plots)} plots using OO approach")
        print("Each plot configured with enum-based properties")
        print("Plots grouped and rescaled using layer methods")
        
        # Keep Origin open for inspection
        input("Press Enter to close Origin and continue...")
        
        origin.close()
        
    except Exception as e:
        print(f"Error in OO approach demo: {e}")
        if 'origin' in locals():
            origin.close(False)
        raise


if __name__ == "__main__":
    print("FigurePage Plotting Test Suite")
    print("=" * 50)
    
    try:
        # Run tests
        test_basic_plotting()
        print()
        test_advanced_features()
        print()
        demonstrate_object_oriented_approach()
        
        print()
        print("All tests completed successfully!")
        
    except Exception as e:
        print(f"Test failed: {e}")
        sys.exit(1)
