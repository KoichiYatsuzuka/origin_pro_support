"""
包括的なグラフ作成機能のテスト
"""
import os
import sys
import numpy as np
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from origin_pro_support import OriginInstance
from origin_pro_support.layer.enums import XYPlotType
from origin_pro_support.pages import FigurePage

def test_comprehensive_graph():
    """包括的なグラフ作成のテスト"""
    print('=== 包括的グラフ作成テスト開始 ===')
    
    # テスト用のOriginプロジェクトファイルパス
    test_file = os.path.join(os.path.dirname(__file__), 'comprehensive_graph_test.opju')
    
    try:
        # Originインスタンスを作成
        print('Originインスタンスを作成中...')
        origin = OriginInstance(test_file)
        
        # ワークブックを作成してテストデータを追加
        print('ワークブックを作成中...')
        workbook = origin.new_workbook("TestData")
        
        if workbook:
            print(f'ワークブック作成成功: {workbook.name}')
            
            # テストデータを作成
            x_data = np.linspace(0, 10, 50)
            y_data = np.sin(x_data) + np.random.normal(0, 0.1, 50)
            
            # ワークシートにデータを追加
            worksheet = workbook.add_worksheet("Data1")
            
            # 手動で列を作成してデータを設定
            worksheet.set_cols(2)
            worksheet.set_rows(len(x_data))
            
            # Xデータを設定
            x_col = worksheet[0]
            x_col.name = "X"
            x_col.long_name = "X data"
            x_col.units = "units"
            x_col.comments = "X data"
            
            # Yデータを設定  
            y_col = worksheet[1]
            y_col.name = "Y"
            y_col.long_name = "Y data"
            y_col.units = "units"
            y_col.comments = "Y data"
            
            # データを設定（列のデータを直接設定）
            x_data_list = x_data.tolist()
            y_data_list = y_data.tolist()
            
            # データを一括で設定
            x_col.set_data(x_data_list)
            y_col.set_data(y_data_list)
            
            print(f'ワークシート作成成功: {worksheet.name}')
            print(f'列数: {len(worksheet.columns)}')
            
            # 新しいグラフを作成
            print('散布図を作成中...')
            graph_page = origin.new_graph("ScatterPlot", XYPlotType.SCATTER)
            
            if graph_page:
                print(f'グラフ作成成功: {graph_page.name}')
                print(f'グラフタイプ: {graph_page.type}')
                print(f'レイヤー数: {len(graph_page)}')
                
                # FigurePageを作成
                print('FigurePageを作成中...')
                figure_page = FigurePage(graph_page, XYPlotType.SCATTER)
                
                # アクティブレイヤーを取得
                active_layer = figure_page.get_active_layer()
                print(f'アクティブレイヤー取得成功: {type(active_layer).__name__}')
                
                # データをプロット
                print('データをプロット中...')
                data_plot = figure_page.plot_xy_data(worksheet, 0, 1, XYPlotType.SCATTER)
                
                if data_plot:
                    print(f'データプロット作成成功: {data_plot.name}')
                    print(f'プロットタイプ: {type(data_plot).__name__}')
                    
                    # プロット内容を検証
                    print('プロット内容を検証中...')
                    active_layer = figure_page.get_active_layer()
                    
                    # データプロットの情報を取得
                    plot_count = len(active_layer.data_plots)
                    print(f'アクティブレイヤーのプロット数: {plot_count}')
                    
                    if plot_count > 0:
                        # 最初のプロットを取得
                        plot = active_layer[0]
                        print(f'プロット名: {plot.name}')
                        
                        # プロットされたデータのワークシートを取得
                        plot_worksheet = plot.worksheet
                        print(f'プロット元ワークシート: {plot_worksheet.name}')
                        
                        # データを比較検証
                        original_x = worksheet[0].get_data()
                        original_y = worksheet[1].get_data()
                        
                        plotted_x = plot_worksheet[0].get_data()
                        plotted_y = plot_worksheet[1].get_data()
                        
                        # データ比較
                        x_match = len(original_x) == len(plotted_x) and all(abs(a - b) < 1e-10 for a, b in zip(original_x, plotted_x))
                        y_match = len(original_y) == len(plotted_y) and all(abs(a - b) < 1e-10 for a, b in zip(original_y, plotted_y))
                        
                        if x_match and y_match:
                            print('✅ プロットデータが元データと一致しました！')
                        else:
                            print('❌ プロットデータが元データと一致しません')
                            print(f'元データX長さ: {len(original_x)}, プロットX長さ: {len(plotted_x)}')
                            print(f'元データY長さ: {len(original_y)}, プロットY長さ: {len(plotted_y)}')
                    else:
                        print('❌ プロットが見つかりません')
                else:
                    print('❌ データプロット作成失敗')
                
                # 線グラフも作成
                print('線グラフを作成中...')
                line_graph = origin.new_graph("LinePlot", XYPlotType.LINE)
                
                if line_graph:
                    line_figure = FigurePage(line_graph, XYPlotType.LINE)
                    line_plot = line_figure.plot_xy_data(worksheet, 0, 1, XYPlotType.LINE)
                    
                    if line_plot:
                        print(f'線グラフプロット作成成功: {line_plot.name}')
                
        else:
            print('ワークブック作成失敗')
            
        # 保存して終了
        origin.close()
        print('包括的テスト完了')
        
    except Exception as e:
        print(f'エラーが発生しました: {e}')
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    test_comprehensive_graph()
