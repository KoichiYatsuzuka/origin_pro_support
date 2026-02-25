from origin_pro_support.layer.enums import PlotType, PlotTypeInfo

# テスト
print('=== PlotTypeInfoとPlotTypeのテスト ===')

# PlotTypeのテスト
print('PlotType.LINE:', PlotType.LINE)
print('numeric value:', PlotType.LINE.numeric)
print('template value:', PlotType.LINE.template)
print('description:', PlotType.LINE.description)
print()

# PlotTypeInfoが不変であることを確認
info = PlotType.LINE.value
print('PlotTypeInfo type:', type(info))
print('PlotTypeInfo attributes:', info)
print()

# 不変性テスト（エラーになるはず）
try:
    info.numeric_value = 999
    print('ERROR: PlotTypeInfo should be immutable!')
except Exception as e:
    print('✅ PlotTypeInfo is immutable:', type(e).__name__)

# LabTalkコマンドの例
cmd1 = f'plot:={PlotType.LINE.numeric}'
cmd2 = f'template:="{PlotType.LINE.template}"'

print('LabTalk numeric command:', cmd1)
print('LabTalk template command:', cmd2)

# XYTemplateから移行したプロットタイプのテスト
print()
print('=== XYTemplateから移行したプロットタイプ ===')
print('PlotType.STEP:', PlotType.STEP)
print('numeric:', PlotType.STEP.numeric)
print('template:', PlotType.STEP.template)
print('description:', PlotType.STEP.description)
