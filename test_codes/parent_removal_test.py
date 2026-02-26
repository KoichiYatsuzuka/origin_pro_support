from origin_pro_support.base import OriginObjectWrapper
from origin_pro_support.pages import PageBase, WorkbookPage, GraphPage
from origin_pro_support.layer import Layer, Column, GraphLayer

print('=== _parentメンバ削除テスト ===')

# OriginObjectWrapperのテスト
print('OriginObjectWrapper._parent属性:', hasattr(OriginObjectWrapper, '_parent'))

# PageBaseのテスト
print('PageBase._parent属性:', hasattr(PageBase, '_parent'))

# Layerのテスト
print('Layer._parent属性:', hasattr(Layer, '_parent'))

# Columnのテスト
print('Column._parent属性:', hasattr(Column, '_parent'))

# GraphLayerのテスト
print('GraphLayer._parent属性:', hasattr(GraphLayer, '_parent'))

print('_parentメンバ削除完了')
