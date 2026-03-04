

def axis_pg(
    axis: str,
    type: str,
    varName: str,
)->str:
    """
    https://www.originlab.com/doc/ja/LabTalk/ref/Axis-cmd
    Syntax: axis -pg axis type varName

    Get parameters for the specified axis (X or Y, Z is not supported) and place the value into varName. Where type is one of:

    S	Scale type: 0 = Linear, 1 = Log10, etc.
    G	Grids displayed: 0 = None, 1 = Major, 2 = Minor, 3 = Both
    A	Axis display: 0 = None, 1 = First, 2 = Second, 3 = Both
    L	axis tick Label display: 0 = None, 1 = First, 2 = Second, 3 = Both
    R	Rescale margin: in % of axis scale
    M	number of Minor ticks
    I	major tick Increment
    """
    return f'axis -pg {axis} {type} {varName}'


def axis_ps(
    axis: str,
    type: str,
    value,
) -> str:
    """
    https://www.originlab.com/doc/ja/LabTalk/ref/Axis-cmd
    Syntax: axis -ps axis type value

    Set parameters for the specified axis (X or Y). Valid types are as described for axis_pg, plus:

    S	Scale type: 0 = Linear, 1 = Log10, etc.
    G	Grids displayed: 0 = None, 1 = Major, 2 = Minor, 3 = Both
    A	Axis display: 0 = None, 1 = First, 2 = Second, 3 = Both
    L	axis tick Label display: 0 = None, 1 = First, 2 = Second, 3 = Both
    R	Rescale margin: in % of axis scale
    M	number of Minor ticks
    I	major tick Increment
    """
    return f'axis -ps {axis} {type} {value}'
