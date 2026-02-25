from origin_pro_support.layer.enums import XYPlotType, PlotTypeInfo

print('=== XYPlotType名前変更テスト ===')

# テスト
print('XYPlotType.LINE:', XYPlotType.LINE)
print('XYPlotType.SCATTER:', XYPlotType.SCATTER)
print('XYPlotType.LINE_SYMBOL:', XYPlotType.LINE_SYMBOL)

# キャストテスト
info = XYPlotType.LINE.value
print('int(info):', int(info))
print('str(info):', str(info))

# LabTalkコマンドテスト
cmd1 = f'plot:={int(XYPlotType.LINE.value)}'
cmd2 = f'template:="{str(XYPlotType.LINE.value)}"'

print('Numeric command:', cmd1)
print('Template command:', cmd2)

print('XYPlotType名前変更完了')
