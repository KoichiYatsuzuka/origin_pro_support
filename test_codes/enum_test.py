import enum

class PlotInfo:
    def __init__(self, numeric_value: int, template_value: str, description: str):
        self.numeric_value = numeric_value
        self.template_value = template_value
        self.description = description
    
    def __repr__(self):
        return f'PlotInfo({self.numeric_value}, "{self.template_value}", "{self.description}")'

class PlotType(enum.Enum):
    LINE = PlotInfo(200, 'line', 'Line plot')
    SCATTER = PlotInfo(201, 'scatter', 'Scatter plot')
    LINE_SYMBOL = PlotInfo(202, 'linesymb', 'Line + Symbol plot')
    COLUMN = PlotInfo(203, 'column', 'Column/bar plot')
    
    @property
    def numeric(self) -> int:
        return self.value.numeric_value
    
    @property
    def template(self) -> str:
        return self.value.template_value
    
    @property
    def description(self) -> str:
        return self.value.description

# テスト
print('PlotType.LINE:', PlotType.LINE)
print('numeric value:', PlotType.LINE.numeric)
print('template value:', PlotType.LINE.template)
print('description:', PlotType.LINE.description)
print()

# LabTalkコマンドの例
cmd1 = f'plot:={PlotType.LINE.numeric}'
cmd2 = f'template:="{PlotType.LINE.template}"'

print('LabTalk numeric command:', cmd1)
print('LabTalk template command:', cmd2)
