from typing import Union


def axis_pg(
    axis: str,
    type: str,
    varName: str,
) -> str:
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


# ── Layer.Axis object property helpers ────────────────────────────────────────
# Ref: https://www.originlab.com/doc/LabTalk/ref/Layer-Axis-obj
#
# All functions below return a LabTalk command string using the
# ``[PageName]!layerN.axis.prop`` notation so the target page and layer are
# explicit and do NOT rely on the currently active window.


def layer_axis_get(page_name: str, layer_id: int, axis: str, prop: str, var: str) -> str:
    """Get a Layer.Axis property and store the result in a LabTalk variable.

    ### page_name
        Short name of the graph page (e.g. ``"Graph1"``).

    ### layer_id
        0-based layer index.

    ### axis
        Axis identifier: ``"x"``, ``"y"``, ``"z"``, ``"x2"``, or ``"y2"``.

    ### prop
        Property name on the Layer.Axis object, e.g. ``"from"``, ``"to"``,
        ``"inc"``, ``"reverse"``, ``"showAxes"``, ``"ticks"``.

    ### var
        Name of the LabTalk variable to receive the value.

    Generates: ``var = [PageName]!layerN.axis.prop``

    Ref: https://www.originlab.com/doc/LabTalk/ref/Layer-Axis-obj
    """
    return f"{var} = [{page_name}]!layer{layer_id + 1}.{axis}.{prop}"


def layer_axis_set(page_name: str, layer_id: int, axis: str, prop: str, value: Union[int, float, str]) -> str:
    """Set a Layer.Axis property.

    ### page_name
        Short name of the graph page (e.g. ``"Graph1"``).

    ### layer_id
        0-based layer index.

    ### axis
        Axis identifier: ``"x"``, ``"y"``, ``"z"``, ``"x2"``, or ``"y2"``.

    ### prop
        Property name on the Layer.Axis object, e.g. ``"from"``, ``"to"``,
        ``"inc"``, ``"reverse"``, ``"showAxes"``, ``"ticks"``.

    ### value
        Value to assign.

    Generates: ``[PageName]!layerN.axis.prop = value``

    Ref: https://www.originlab.com/doc/LabTalk/ref/Layer-Axis-obj
    """
    return f"[{page_name}]!layer{layer_id + 1}.{axis}.{prop} = {value}"


def layer_axis_get_from(page_name: str, layer_id: int, axis: str, var: str) -> str:
    """Get the ``from`` (minimum) scale value of an axis.

    ### page_name
        Short name of the graph page (e.g. ``"Graph1"``).

    ### layer_id
        0-based layer index.

    ### axis
        Axis identifier: ``"x"``, ``"y"``, ``"z"``, ``"x2"``, or ``"y2"``.

    ### var
        LabTalk variable name to receive the value.

    Generates: ``var = [PageName]!layerN.axis.from``

    Ref: https://www.originlab.com/doc/LabTalk/ref/Layer-Axis-obj
    """
    return f"{var} = [{page_name}]!layer{layer_id + 1}.{axis}.from"


def layer_axis_set_from(page_name: str, layer_id: int, axis: str, value: Union[int, float]) -> str:
    """Set the ``from`` (minimum) scale value of an axis.

    ### page_name
        Short name of the graph page (e.g. ``"Graph1"``).

    ### layer_id
        0-based layer index.

    ### axis
        Axis identifier: ``"x"``, ``"y"``, ``"z"``, ``"x2"``, or ``"y2"``.

    ### value
        From (minimum) scale value.

    Generates: ``[PageName]!layerN.axis.from = value``

    Ref: https://www.originlab.com/doc/LabTalk/ref/Layer-Axis-obj
    """
    return f"[{page_name}]!layer{layer_id + 1}.{axis}.from = {value}"


def layer_axis_get_to(page_name: str, layer_id: int, axis: str, var: str) -> str:
    """Get the ``to`` (maximum) scale value of an axis.

    ### page_name
        Short name of the graph page (e.g. ``"Graph1"``).

    ### layer_id
        0-based layer index.

    ### axis
        Axis identifier: ``"x"``, ``"y"``, ``"z"``, ``"x2"``, or ``"y2"``.

    ### var
        LabTalk variable name to receive the value.

    Generates: ``var = [PageName]!layerN.axis.to``

    Ref: https://www.originlab.com/doc/LabTalk/ref/Layer-Axis-obj
    """
    return f"{var} = [{page_name}]!layer{layer_id + 1}.{axis}.to"


def layer_axis_set_to(page_name: str, layer_id: int, axis: str, value: Union[int, float]) -> str:
    """Set the ``to`` (maximum) scale value of an axis.

    ### page_name
        Short name of the graph page (e.g. ``"Graph1"``).

    ### layer_id
        0-based layer index.

    ### axis
        Axis identifier: ``"x"``, ``"y"``, ``"z"``, ``"x2"``, or ``"y2"``.

    ### value
        To (maximum) scale value.

    Generates: ``[PageName]!layerN.axis.to = value``

    Ref: https://www.originlab.com/doc/LabTalk/ref/Layer-Axis-obj
    """
    return f"[{page_name}]!layer{layer_id + 1}.{axis}.to = {value}"


def layer_axis_get_ticks(page_name: str, layer_id: int, axis: str, var: str) -> str:
    """Get the ``ticks`` bitmask controlling major/minor tick display.

    ### page_name
        Short name of the graph page (e.g. ``"Graph1"``).

    ### layer_id
        0-based layer index.

    ### axis
        Axis identifier: ``"x"``, ``"y"``, ``"z"``, ``"x2"``, or ``"y2"``.

    ### var
        LabTalk variable name to receive the bitmask value.

    Bitmask encoding:
        - 0 = no ticks
        - 1 = major in
        - 2 = major out
        - 4 = minor in
        - 8 = minor out
        - Values combine: e.g. 1+2+8 = 11 (major in+out, minor out)

    Generates: ``var = [PageName]!layerN.axis.ticks``

    Ref: https://www.originlab.com/doc/LabTalk/ref/Layer-Axis-obj
    """
    return f"{var} = [{page_name}]!layer{layer_id + 1}.{axis}.ticks"


def layer_axis_set_ticks(page_name: str, layer_id: int, axis: str, value: int) -> str:
    """Set the ``ticks`` bitmask controlling major/minor tick display.

    ### page_name
        Short name of the graph page (e.g. ``"Graph1"``).

    ### layer_id
        0-based layer index.

    ### axis
        Axis identifier: ``"x"``, ``"y"``, ``"z"``, ``"x2"``, or ``"y2"``.

    ### value
        Bitmask value (see ``layer_axis_get_ticks`` for encoding).

    Generates: ``[PageName]!layerN.axis.ticks = value``

    Ref: https://www.originlab.com/doc/LabTalk/ref/Layer-Axis-obj
    """
    return f"[{page_name}]!layer{layer_id + 1}.{axis}.ticks = {value}"


def axis_ps(
    axis: str,
    type: str,
    value: Union[int, float, str],
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
