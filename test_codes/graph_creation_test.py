"""
グラフ作成機能のテスト
"""
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from origin_pro_support import OriginInstance
from origin_pro_support.layer.enums import XYPlotType
from origin_pro_support.pages import FigurePage

def test_graph_creation():
    """グラフ作成のテスト"""
    print('=== グラフ作成テスト開始 ===')
    
    # テスト用のOriginプロジェクトファイルパス
    test_file = os.path.join(os.path.dirname(__file__), 'graph_test.opju')
    
    try:
        # Originインスタンスを作成
        print('Originインスタンスを作成中...')
        origin = OriginInstance(test_file)
        
        # 新しいグラフを作成
        print('新しいグラフを作成中...')
        graph_page = origin.new_graph("TestGraph", XYPlotType.SCATTER)
        
        if graph_page:
            print(f'グラフ作成成功: {graph_page.name}')
            print(f'グラフタイプ: {graph_page.type}')
            print(f'レイヤー数: {len(graph_page)}')
            
            # FigurePageを作成してテスト
            print('FigurePageを作成中...')
            figure_page = FigurePage(graph_page, XYPlotType.LINE)
            print(f'FigurePage作成成功: {figure_page.name}')
            
            # アクティブレイヤーを取得
            active_layer = figure_page.get_active_layer()
            print(f'アクティブレイヤー取得成功: {type(active_layer).__name__}')
            
        else:
            print('グラフ作成失敗')
            
        # 保存して終了
        origin.close()
        print('テスト完了')
        
    except Exception as e:
        print(f'エラーが発生しました: {e}')
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    test_graph_creation()
