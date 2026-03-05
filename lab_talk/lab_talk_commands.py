from typing import Union


# ── Window helpers ─────────────────────────────────────────────────────────────

def win_activate(page_name: str) -> str:
    """Activate the specified window.

    ### page_name
        Short name of the page to activate.

    Generates: ``win -a <page_name>``

    Ref: https://www.originlab.com/doc/LabTalk/ref/Win-cmd
    """
    return f"win -a {page_name}"


# ── Legend graphic object helpers ──────────────────────────────────────────────
# Ref: https://www.originlab.com/doc/LabTalk/ref/Graphic-objs
#      https://www.originlab.com/doc/LabTalk/ref/Legend-cmd
#
# These functions assume the parent page has been activated first so that
# LabTalk's ``legend`` object resolves to the correct layer.


def legend_get_num(prop: str, var: str) -> str:
    """Get a numeric property of the active layer's legend object.

    ### prop
        Property name, e.g. ``"show"``, ``"fsize"``, ``"background"``,
        ``"x"``, ``"y"``, ``"dx"``, ``"dy"``.

    ### var
        LabTalk variable name to receive the value.

    Generates: ``var = legend.prop``
    """
    return f"{var} = legend.{prop}"


def legend_set_num(prop: str, value: Union[int, float]) -> str:
    """Set a numeric property of the active layer's legend object.

    ### prop
        Property name, e.g. ``"show"``, ``"fsize"``, ``"background"``,
        ``"x"``, ``"y"``.

    ### value
        Numeric value to assign.

    Generates: ``legend.prop = value``
    """
    return f"legend.{prop} = {value}"


def legend_get_str(prop: str, var: str) -> str:
    """Get a string property of the active layer's legend object.

    ### prop
        String property name without ``$``, e.g. ``"text"``.

    ### var
        LabTalk string variable name to receive the value (without ``$``).

    Generates: ``var$ = legend.prop$``
    """
    return f"{var}$ = legend.{prop}$"


def legend_set_str(prop: str, value: str) -> str:
    """Set a string property of the active layer's legend object.

    *value* must already be escaped (double-quotes inside the string should
    be backslash-escaped before calling this function).

    ### prop
        String property name without ``$``, e.g. ``"text"``.

    ### value
        Escaped string value (without surrounding double-quotes).

    Generates: ``legend.prop$ = "value"``
    """
    return f'legend.{prop}$ = "{value}"'


def legend_update() -> str:
    """Create or update the legend on the active layer.

    Generates: ``legend``

    Ref: https://www.originlab.com/doc/LabTalk/ref/Legend-cmd
    """
    return "legend"


def legend_reconstruct() -> str:
    """Reconstruct the legend to best-fit the current plots.

    Equivalent to menu Graph > Legend > Reconstruct Legend.

    Generates: ``legend -r``

    Ref: https://www.originlab.com/doc/LabTalk/ref/Legend-cmd
    """
    return "legend -r"


def legend_reset_position() -> str:
    """Reset the legend to its default position.

    Generates: ``legend -d``

    Ref: https://www.originlab.com/doc/LabTalk/ref/Legend-cmd
    """
    return "legend -d"


def legend_set_layout(option: str) -> str:
    """Rearrange legend entries.

    ### option
        Layout option flag without leading ``-``, e.g. ``"ah"`` for horizontal
        or ``"av"`` for vertical.

    Generates: ``legend -<option>``

    Ref: https://www.originlab.com/doc/LabTalk/ref/Legend-cmd
    """
    return f"legend -{option}"


# ── Active-layer axis range helpers ───────────────────────────────────────────
# These use the unqualified ``layer.axis.prop`` form and therefore require
# the parent page to have been activated beforehand.


def active_layer_x_get_from(var: str) -> str:
    """Get the x-axis ``from`` (minimum) value of the active layer.

    Generates: ``var = layer.x.from``
    """
    return f"{var} = layer.x.from"


def active_layer_x_get_to(var: str) -> str:
    """Get the x-axis ``to`` (maximum) value of the active layer.

    Generates: ``var = layer.x.to``
    """
    return f"{var} = layer.x.to"


def active_layer_y_get_from(var: str) -> str:
    """Get the y-axis ``from`` (minimum) value of the active layer.

    Generates: ``var = layer.y.from``
    """
    return f"{var} = layer.y.from"


def active_layer_y_get_to(var: str) -> str:
    """Get the y-axis ``to`` (maximum) value of the active layer.

    Generates: ``var = layer.y.to``
    """
    return f"{var} = layer.y.to"


def active_layer_x_set_from(value: Union[int, float]) -> str:
    """Set the x-axis ``from`` (minimum) value of the active layer.

    Generates: ``layer.x.from = value``
    """
    return f"layer.x.from = {value}"


def active_layer_x_set_to(value: Union[int, float]) -> str:
    """Set the x-axis ``to`` (maximum) value of the active layer.

    Generates: ``layer.x.to = value``
    """
    return f"layer.x.to = {value}"


def active_layer_y_set_from(value: Union[int, float]) -> str:
    """Set the y-axis ``from`` (minimum) value of the active layer.

    Generates: ``layer.y.from = value``
    """
    return f"layer.y.from = {value}"


def active_layer_y_set_to(value: Union[int, float]) -> str:
    """Set the y-axis ``to`` (maximum) value of the active layer.

    Generates: ``layer.y.to = value``
    """
    return f"layer.y.to = {value}"


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


# ── Layer.Plot object property helpers ────────────────────────────────────────
# Ref: https://www.originlab.com/doc/LabTalk/ref/Layer-plot-obj
#
# All functions below use the ``[PageName]!layerN.plot(M).prop`` notation so
# the target page and layer are explicit and do NOT rely on the active window.


def layer_plot_get(page_name: str, layer_id: int, plot_id: int, prop: str, var: str) -> str:
    """Get a Layer.plot property and store the result in a LabTalk variable.

    ### page_name
        Short name of the graph page (e.g. ``"Graph1"``).

    ### layer_id
        0-based layer index.

    ### plot_id
        1-based plot index within the layer.

    ### prop
        Property path on the Layer.plot object, e.g. ``"color"``,
        ``"symbol.size"``, ``"symbol.kind"``.

    ### var
        Name of the LabTalk variable to receive the value.

    Generates: ``var = [PageName]!layerN.plot(M).prop``

    Ref: https://www.originlab.com/doc/LabTalk/ref/Layer-plot-obj
    """
    return f"{var} = [{page_name}]!layer{layer_id + 1}.plot({plot_id}).{prop}"


def layer_plot_set(page_name: str, layer_id: int, plot_id: int, prop: str, value: Union[int, float, str]) -> str:
    """Set a Layer.plot property.

    ### page_name
        Short name of the graph page (e.g. ``"Graph1"``).

    ### layer_id
        0-based layer index.

    ### plot_id
        1-based plot index within the layer.

    ### prop
        Property path on the Layer.plot object, e.g. ``"color"``,
        ``"symbol.size"``, ``"symbol.kind"``.

    ### value
        Value to assign.

    Generates: ``[PageName]!layerN.plot(M).prop = value``

    Ref: https://www.originlab.com/doc/LabTalk/ref/Layer-plot-obj
    """
    return f"[{page_name}]!layer{layer_id + 1}.plot({plot_id}).{prop} = {value}"


def layer_plot_get_name(page_name: str, layer_id: int, plot_id: int, var: str) -> str:
    """Get the dataset name of a plot and store it in a LabTalk string variable.

    ### page_name
        Short name of the graph page (e.g. ``"Graph1"``).

    ### layer_id
        0-based layer index.

    ### plot_id
        1-based plot index within the layer.

    ### var
        Name of the LabTalk *string* variable to receive the value (without ``$``).

    Generates: ``var$ = [PageName]!layerN.plot(M).name$``

    Ref: https://www.originlab.com/doc/LabTalk/ref/Layer-plot-obj
    """
    return f"{var}$ = [{page_name}]!layer{layer_id + 1}.plot({plot_id}).name$"


def layer_plot_count(page_name: str, layer_id: int, var: str) -> str:
    """Get the number of plots in a layer and store it in a LabTalk variable.

    ### page_name
        Short name of the graph page (e.g. ``"Graph1"``).

    ### layer_id
        0-based layer index.

    ### var
        Name of the LabTalk variable to receive the count.

    Generates: ``var = [PageName]!layerN.numplots``

    Ref: https://www.originlab.com/doc/LabTalk/ref/Layer-obj
    """
    return f"{var} = [{page_name}]!layer{layer_id + 1}.numplots"


def plot_set_symbol_size(dataset_name: str, size: float) -> str:
    """Set the symbol size of a plot via the ``set`` LabTalk command.

    ### dataset_name
        The dataset name as returned by ``DataObjectBase.GetDatasetName()``,
        e.g. ``"Book1_B"``.

    ### size
        Symbol size as a positive float.

    Generates: ``set <dataset_name> -z <size>``

    Ref: https://www.originlab.com/doc/LabTalk/ref/Set-cmd
    """
    return f"set {dataset_name} -z {size}"


def plot_set_symbol_kind(dataset_name: str, kind: int) -> str:
    """Set the symbol shape (kind) of a plot via the ``set`` LabTalk command.

    ### dataset_name
        The dataset name as returned by ``DataObjectBase.GetDatasetName()``,
        e.g. ``"Book1_B"``.

    ### kind
        Symbol kind integer corresponding to ``MarkerShape`` enum values.

    Generates: ``set <dataset_name> -k <kind>``

    Ref: https://www.originlab.com/doc/LabTalk/ref/Set-cmd
    """
    return f"set {dataset_name} -k {kind}"


def plot_set_line_style(dataset_name: str, style: int) -> str:
    """Set the line style of a plot via the ``set`` LabTalk command.

    ### dataset_name
        The dataset name as returned by ``DataObjectBase.GetDatasetName()``,
        e.g. ``"Book1_B"``.

    ### style
        Line style integer corresponding to ``LineStyle`` enum values.
        1=Solid, 2=Dash, 3=Dot, 4=DashDot, 5=DashDotDot.

    Generates: ``set <dataset_name> -d <style>``

    Ref: https://www.originlab.com/doc/LabTalk/ref/Set-cmd
    """
    return f"set {dataset_name} -d {style}"


def plot_set_line_width(dataset_name: str, width: float) -> str:
    """Set the line width of a plot via the ``set`` LabTalk command.

    ### dataset_name
        The dataset name as returned by ``DataObjectBase.GetDatasetName()``,
        e.g. ``"Book1_B"``.

    ### width
        Line width as a positive float (in points).

    Generates: ``set <dataset_name> -w <width>``

    Ref: https://www.originlab.com/doc/LabTalk/ref/Set-cmd
    """
    return f"set {dataset_name} -w {width}"


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
