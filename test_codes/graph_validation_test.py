"""
Graph creation and plot validation test
"""
import os
import sys
import numpy as np
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from origin_pro_support import OriginInstance
from origin_pro_support.layer.enums import XYPlotType

def test_graph_validation():
    """Graph creation and plot validation test"""
    print('=== Graph creation and plot validation test start ===')
    
    # Test Origin project file path
    import time
    timestamp = int(time.time())
    test_file = os.path.join(os.path.dirname(__file__), f'graph_validation_test_{timestamp}.opju')
    
    try:
        # Create Origin instance
        print('1. Creating Origin instance...')
        origin = OriginInstance(test_file)
        
        # Create workbook and add test data
        print('2. Creating workbook...')
        workbook = origin.new_workbook("TestData")
        
        if workbook:
            print(f'   [OK] Workbook creation success: {workbook.name}')
            
            # Create test data
            x_data = np.linspace(0, 10, 20)  # Smaller data count
            y_data = np.sin(x_data) + np.random.normal(0, 0.1, 20)
            
            # Add worksheet and data
            worksheet = workbook.add_worksheet("Data1")
            
            # Manually create columns and set data
            worksheet.set_cols(2)
            worksheet.set_rows(len(x_data))
            
            # Set X data
            x_col = worksheet[0]
            x_col.name = "X"
            x_col.long_name = "X data"
            x_col.units = "units"
            
            # Set Y data  
            y_col = worksheet[1]
            y_col.name = "Y"
            y_col.long_name = "Y data"
            y_col.units = "units"
            
            # Set data in batch
            try:
                x_result = x_col.set_data(x_data.tolist())
                y_result = y_col.set_data(y_data.tolist())
                print(f'   [OK] Data setting success: X={x_result}, Y={y_result}')
                
                # Verify data setting by getting values
                try:
                    # Specify get_data format parameter (usually 1 for numeric data)
                    retrieved_x = x_col.get_data(1, 0, len(x_data) - 1, 1)
                    retrieved_y = y_col.get_data(1, 0, len(y_data) - 1, 1)
                    
                    # Compare first few points for verification
                    if len(retrieved_x) >= 3 and len(retrieved_y) >= 3:
                        x_match = (abs(retrieved_x[0] - x_data[0]) < 1e-10 and 
                                  abs(retrieved_x[1] - x_data[1]) < 1e-10 and 
                                  abs(retrieved_x[2] - x_data[2]) < 1e-10)
                        y_match = (abs(retrieved_y[0] - y_data[0]) < 1e-10 and 
                                  abs(retrieved_y[1] - y_data[1]) < 1e-10 and 
                                  abs(retrieved_y[2] - y_data[2]) < 1e-10)
                        
                        if x_match and y_match:
                            print(f'   [OK] Data verification success: First 3 points match')
                        else:
                            print(f'   [ERROR] Data verification failed: Data does not match')
                            print(f'      X: Set={x_data[0]:.3f}, Get={retrieved_x[0]:.3f}')
                            print(f'      Y: Set={y_data[0]:.3f}, Get={retrieved_y[0]:.3f}')
                            return False
                    else:
                        print(f'   [ERROR] Data verification failed: Insufficient retrieved data')
                        print(f'      X: Retrieved={len(retrieved_x)} points, Expected={len(x_data)} points')
                        print(f'      Y: Retrieved={len(retrieved_y)} points, Expected={len(y_data)} points')
                        return False
                        
                except Exception as e:
                    print(f'   [ERROR] Data verification failed: {e}')
                    return False
                    
            except Exception as e:
                print(f'   [ERROR] Data setting failed: {e}')
                return False
            
            print(f'   [OK] Worksheet creation success: {worksheet.name}')
            print(f'   [OK] Data setting complete: X={len(x_data)} points, Y={len(y_data)} points')
            
            # Create new graph
            print('3. Creating graph...')
            graph_name = f"TestGraph_{timestamp}"
            print(f'   [DEBUG] XYPlotType.SCATTER.value: {XYPlotType.SCATTER.value}')
            print(f'   [DEBUG] str(XYPlotType.SCATTER.value): {str(XYPlotType.SCATTER.value)}')
            
            try:
                graph_page = origin.new_graph(graph_name, XYPlotType.SCATTER)
                print(f'   [DEBUG] new_graph return value: {graph_page}')
            except Exception as e:
                print(f'   [ERROR] Exception during graph creation: {e}')
                import traceback
                traceback.print_exc()
                return False
            
            if graph_page:
                print(f'   [OK] Graph creation success: {graph_page.name}')
                print(f'   [OK] Graph type: {graph_page.type}')
                print(f'   [OK] Layer count: {len(graph_page)}')
                
                # Create FigurePage
                print('4. Creating FigurePage...')
                
                # Get active layer
                active_layer = graph_page.get_layer(0)
                print(f'   [OK] Active layer acquisition success: {type(active_layer).__name__}')
                
                # Plot data
                print('5. Plotting data...')
                data_plot = graph_page.plot_xy_data(worksheet, 0, 1, XYPlotType.SCATTER)
                
                if data_plot is not None:
                    print(f'   [OK] Data plot creation success: {type(data_plot).__name__}')
                    
                    # Check plot count immediately after plot creation
                    print('5.5. Checking plot count immediately after plot creation...')
                    plot_count_after = len(active_layer.data_plots)
                    print(f'   [INFO] Plot count after plot creation: {plot_count_after}')
                    
                    # Verify plot content
                    print('6. Verifying plot content...')
                    
                    # Get data plot information
                    if plot_count_after > 0:
                        # Get first plot
                        plot = active_layer[0]
                        print(f'   [INFO] Plot name: {plot.name}')
                        
                        # Get plotted data worksheet
                        plot_worksheet = plot.worksheet
                        if plot_worksheet:
                            print(f'   [INFO] Plot source worksheet: {plot_worksheet.name}')
                            
                            # Compare and verify data
                            print('7. Data comparison verification...')
                            original_x = worksheet[0].get_data()
                            original_y = worksheet[1].get_data()
                            
                            plotted_x = plot_worksheet[0].get_data()
                            plotted_y = plot_worksheet[1].get_data()
                            
                            print(f'   [INFO] Original X length: {len(original_x)}, Plot X length: {len(plotted_x)}')
                            print(f'   [INFO] Original Y length: {len(original_y)}, Plot Y length: {len(plotted_y)}')
                            
                            # Data comparison (show first 5 points)
                            print('   [INFO] Comparing first 5 points:')
                            for i in range(min(5, len(original_x))):
                                print(f'      Point{i}: Original=({original_x[i]:.3f}, {original_y[i]:.3f}), Plot=({plotted_x[i]:.3f}, {plotted_y[i]:.3f})')
                            
                            # Data comparison
                            x_match = len(original_x) == len(plotted_x) and all(abs(a - b) < 1e-10 for a, b in zip(original_x, plotted_x))
                            y_match = len(original_y) == len(plotted_y) and all(abs(a - b) < 1e-10 for a, b in zip(original_y, plotted_y))
                            
                            if x_match and y_match:
                                print('   [SUCCESS] Plot data matches original data!')
                                print('   [SUCCESS] Graph creation and plot validation test success!')
                            else:
                                print('   [ERROR] Plot data does not match original data')
                        else:
                            print('   [ERROR] Cannot get plot source worksheet')
                            return False
                    else:
                        print('   [ERROR] Plot not found')
                        return False
                else:
                    print('   [ERROR] Data plot creation failed')
                    return False
            else:
                print('   [ERROR] Graph creation failed')
                return False
        else:
            print('   [ERROR] Workbook creation failed')
            return False
            
        # Save and exit
        origin.close()
        print('[SUCCESS] Test complete')
        return True
        
    except Exception as e:
        print(f'[ERROR] Error occurred: {e}')
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_graph_validation()
    if success:
        print("Test passed!")
        sys.exit(0)
    else:
        print("Test failed!")
        sys.exit(1)
