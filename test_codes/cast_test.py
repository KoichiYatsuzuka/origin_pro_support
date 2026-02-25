from origin_pro_support.layer.enums import PlotType, PlotTypeInfo

# テスト
print('=== PlotTypeInfoのキャスト関数テスト ===')

# PlotTypeのテスト
plot_type = PlotType.LINE
info = plot_type.value

print('PlotType.LINE:', plot_type)
print('PlotTypeInfo:', info)
print()

# キャストのテスト
print('=== キャストテスト ===')
print('int(info):', int(info))
print('str(info):', str(info))
print()

# 直接アクセスとの比較
print('=== 直接アクセスとの比較 ===')
print('info.numeric_value:', info.numeric_value)
print('info.template_value:', info.template_value)
print('int(info) == info.numeric_value:', int(info) == info.numeric_value)
print('str(info) == info.template_value:', str(info) == info.template_value)
print()

# LabTalkコマンドでの使用例
print('=== LabTalkコマンドでの使用例 ===')
cmd1 = f'plot:={int(plot_type.value)}'  # int()キャスト
cmd2 = f'template:="{str(plot_type.value)}"'  # str()キャスト

print('Numeric command:', cmd1)
print('Template command:', cmd2)

# プロパティアクセスとの比較
cmd3 = f'plot:={plot_type.numeric}'  # プロパティ
cmd4 = f'template:="{plot_type.template}"'  # プロパティ

print('Property numeric command:', cmd3)
print('Property template command:', cmd4)
print()

print('=== 両方の方法が同じ結果を返す ===')
print('int() vs numeric:', int(plot_type.value) == plot_type.numeric)
print('str() vs template:', str(plot_type.value) == plot_type.template)
