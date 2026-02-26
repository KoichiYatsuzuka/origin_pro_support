from origin_pro_support.base import OriginObjectWrapper, APP
from origin_pro_support.pages import PageBase, WorkbookPage, GraphPage
from origin_pro_support.layer import Layer, Column, GraphLayer

print('=== __API_core変更テスト ===')

# OriginObjectWrapperのテスト
print('OriginObjectWrapper.__API_core属性:', hasattr(OriginObjectWrapper, '__API_core'))
print('OriginObjectWrapper._origin_instance属性:', hasattr(OriginObjectWrapper, '_origin_instance'))

# PageBaseのテスト
print('PageBase.__API_core属性:', hasattr(PageBase, '__API_core'))
print('PageBase._origin_instance属性:', hasattr(PageBase, '_origin_instance'))

# Layerのテスト
print('Layer.__API_core属性:', hasattr(Layer, '__API_core'))
print('Layer._origin_instance属性:', hasattr(Layer, '_origin_instance'))

# Columnのテスト
print('Column.__API_core属性:', hasattr(Column, '__API_core'))
print('Column._origin_instance属性:', hasattr(Column, '_origin_instance'))

# GraphLayerのテスト
print('GraphLayer.__API_core属性:', hasattr(GraphLayer, '__API_core'))
print('GraphLayer._origin_instance属性:', hasattr(GraphLayer, '_origin_instance'))

# APPクラスのテスト
print('APPクラスの存在:', hasattr(APP, '__init__'))

print('__API_core変更完了')
