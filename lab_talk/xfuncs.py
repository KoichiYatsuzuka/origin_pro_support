from typing import Optional

def xyplot(
    iy:str,
    plot: int,
    color: int,
    ogl: str,
    size: Optional[int] = None,
    rescale: int = 1,
    legend: int = 1,
    hide: int = 0,
)->str:
    """
    Create plot from XY data by specifying plot type and properties.
    
    ### iy
    Specifies the input data, including X column, Y column, Y error column.
        example:
        - (?, 5)  // if col(5) happens to be X column call fails
        - (1, 2, 3)  // plot col(2) vs col(1), with col(3) as Y error
        - (1,2:4) // plot col(2) to col(4) vs col(1)
        - [Book1]2!(1,3) // plot col(3) vs col(1) from 2nd sheet of Book1
        - [Book2]"Sheet1"!(Col(1),Col(5)[60]:Col(8)[80]) // plot data from column 5, row 60 to column 8, row 80 from Sheet1 of Book2
        
    ### plot
        Specifies the plot type.
        The most common are 200, line; 201, scatter; 202, line + symbol.
        For the list of all plot types, see Plot Type IDs.
        Note: only the plot types for XY range can be used.

    ### color
        Specifies the plot color.
        Works for line, scatter, and line+symbol plots.
        Color follows Origin's color list: 1 (black), 2 (red), up to 24 (dark gray), or up to 40 if 16 custom colors are defined.
        Can also use color function for RGB colors: color:=color(240,208,0)

    ### size
        Specifies the symbol size in points.
        Default value is -1, which means the size will follow the symbol size specified in the template automatically.

    ### rescale
        Specifies whether to rescale the graph (1) or not (0).

    ### legend
        Specifies whether to create a graph legend.
        legend:=0 can be used when the template doesn't have a legend object saved.
        legend:=1 creates a legend regardless of whether the template has one.

    ### hide
        Specifies whether to hide the created graph (1) or show it (0).

    ### ogl
        Specifies the graph layer to add plots.
        Use range syntax.
        example:
        - [Graph1]1! //add plot to 1st graph layer of Graph1
        - [<new template:=doubley name:=SamplePlot>] // plotting to a new graph layer with doubley template and name SamplePlot
        - 2 // add plot to layer 2
        - <new template:=PolarXrYTheta> // create new polar graph with r(X) theta(Y)
        - <new template:=Polar> // create new polar graph with theta(X) r(Y)

    
    
    Ref: https://www.originlab.com/doc/ja/X-Function/ref/plotxy
    """
    if size is not None:
        seze_str = ""
    else:
        seze_str = f" size:={size}"

    return f"xyplot iy:={iy}"+\
        f" plot:={plot} color:={color} ogl:={ogl}"+\
        f"{seze_str} rescale:={rescale} legend:={legend} hide:={hide}"


def pe_cd(path: str) -> str:
    """
    Change Project Explorer directory or go to root folder.
    
    ### path
        Specifies the path to change to.
        Both absolute and relative paths are supported.
        example:
        - ".." // Moves active folder up one level
        - "/" // Moves active folder to root folder
        - "../subfolder1" // Moves to another folder at the same level
        - "abc" // Moves to a subfolder of the currently active folder
        - "/abc/def" // Absolute path to a folder
    
    Ref: https://www.originlab.com/doc/en/X-Function/ref/pe_cd
    """
    return f"pe_cd path:=\"{path}\""


def pe_dir(
    name: str = "*",
    page: Optional[str] = None,
    oname: Optional[str] = None,
    recursive: int = 0,
    display: int = 0,
    sensitive: int = 0,
    sep: str = "",
) -> str:
    """
    List Project Explorer subfolders and windows.
    
    ### name
        Name of the desired page; "*" will match all the pages.
        Supports wildcard patterns (e.g., "B*").
    
    ### page
        Type of page to list.
        - W = Workbook
        - P / G = Graph
        - M = Matrix
        - L = Layout
        - N = Notes
        - I = Image (Origin 2025b)
        - <optional> = Folder
    
    ### oname
        Output variable name to store the names of the pages found.
    
    ### recursive
        Specify whether to search the subfolders for files.
        - 0 = Do not search subfolders
        - 1 = Search subfolders
    
    ### display
        Specify the list display form. Only available when recursive=1.
        - 0 = Simple list
        - 1 = With description
        - 2 = Full details
    
    ### sensitive
        Specify whether the page name display is case sensitive or not.
        - 0 = Case insensitive
        - 1 = Case sensitive
    
    ### sep
        Specify the separators for the description.
        Only available when display=1 to show the description for the folders and windows.
        Note: If use the left bracket (e.g., (, [, {, <) as the separators,
        it will be converted to right bracket for the right side separators.
    
    Ref: https://www.originlab.com/doc/X-Function/ref/pe_dir
    """
    cmd = "pe_dir"
    
    if name != "*":
        cmd += f" name:=\"{name}\""
    
    if page is not None:
        cmd += f" page:={page}"
    
    if oname is not None:
        cmd += f" oname:={oname}"
    
    if recursive != 0:
        cmd += f" recursive:={recursive}"
    
    if display != 0:
        cmd += f" display:={display}"
    
    if sensitive != 0:
        cmd += f" sensitive:={sensitive}"
    
    if sep:
        cmd += f" sep:={sep}"
    
    return cmd


def pe_load(fname: str, path: Optional[str] = None) -> str:
    """
    Load an Origin project into an existing folder in the current project.
    
    ### fname
        Name of the Origin project file to load (.opj).
    
    ### path
        Path to the folder where the project should be loaded.
        If not specified, loads into the current active folder.
    
    Ref: https://www.originlab.com/doc/en/X-Function/ref/pe_load
    """
    cmd = f"pe_load fname:={fname}"
    
    if path is not None:
        cmd += f" path:={path}"
    
    return cmd


def pe_mkdir(
    folder: str,
    chk: int = 0,
    cd: int = 0,
    path: Optional[str] = None,
) -> str:
    """
    Create a new folder in Project Explorer.
    
    ### folder
        The name of the new folder to be created.
        Depending on the value of chk, if the name already exists,
        it will either add a number to end of new folder or it will not create the folder.
    
    ### chk
        If set to 0 and folder already exists, it will create a new folder with number added to the end of the name.
        If set to 1 and folder already exists, it does not create a new folder.
    
    ### cd
        If set to 1, make the newly created folder the active one in Project Explorer.
    
    ### path
        The name of string variable (e.g. strPath$) to hold the full path to the newly-created folder
        (or existing folder) when the X-Function returns.
    
    Ref: https://www.originlab.com/doc/X-Function/ref/pe_mkdir
    """
    cmd = f"pe_mkdir folder:=\"{folder}\""
    
    if chk != 0:
        cmd += f" chk:={chk}"
    
    if cd != 0:
        cmd += f" cd:={cd}"
    
    if path is not None:
        cmd += f" path:={path}"
    
    return cmd


def pe_move(move: str, path: str) -> str:
    """
    Move page or folder to specified folder in Project Explorer.
    
    ### move
        Name of the window or folder to move.
    
    ### path
        Path to Project Explorer folder to move the window or folder to.
    
    Ref: https://www.originlab.com/doc/X-Function/ref/pe_move
    """
    return f"pe_move move:=\"{move}\" path:=\"{path}\""


def pe_path(
    page: str,
    path: Optional[str] = None,
    type: int = 0,
    active: int = 0,
) -> str:
    """
    Get current Project Explorer path or path of specified window.
    
    ### page
        Name of the page or folder to get the path for.
    
    ### path
        Output variable name to store the full path.
    
    ### type
        Type of search:
        - 0 (window) = Search for a window (Page Short Name)
        - 1 (folder) = Search for a subfolder (Subfolder Name)
    
    ### active
        If set to 1, changes active folder in Project Explorer to the specified folder.
    
    Ref: https://www.originlab.com/doc/X-Function/ref/pe_path
    """
    cmd = f"pe_path page:={page}"
    
    if path is not None:
        cmd += f" path:={path}"
    
    if type != 0:
        cmd += f" type:={type}"
    
    if active != 0:
        cmd += f" active:={active}"
    
    return cmd


def pe_rename(old: str, newname: str) -> str:
    """
    Rename Project Explorer folder or window.
    
    ### old
        Name of the existing page or subfolder under current folder.
    
    ### newname
        New name of the page or subfolder.
    
    Ref: https://www.originlab.com/doc/en/X-Function/ref/pe_rename
    """
    return f"pe_rename old:=\"{old}\" newname:={newname}"


def pe_rmdir(
    folder: str,
    folpromt: int = 0,
    pgpromt: int = 0,
) -> str:
    """
    Delete subfolder in Project Explorer.
    
    ### folder
        Specify the name and the path of the folder to delete.
        You can use both absolute or relative way to specify path,
        like "/FolderA/FolderB" or "../FolderB".
    
    ### folpromt
        Specify whether to prompt for folder deletion.
        If TRUE(1), Origin will prompt the user to verify before deleting the folder.
    
    ### pgpromt
        Specify whether to prompt for deletion of pages in subfolder.
        If TRUE(1), Origin will prompt the user to verify before deleting the pages.
    
    Ref: https://www.originlab.com/doc/X-Function/ref/pe_rmdir
    """
    cmd = f"pe_rmdir folder:=\"{folder}\""
    
    if folpromt != 0:
        cmd += f" folpromt:={folpromt}"
    
    if pgpromt != 0:
        cmd += f" pgpromt:={pgpromt}"
    
    return cmd


def pe_save(
    fname: str,
    path: Optional[str] = None,
) -> str:
    """
    Save contents of Project Explorer subfolder as an Origin project file.
    
    ### fname
        Name of the Origin project file to save (.opj).
    
    ### path
        Path to the folder to save.
        If not specified, saves the current active folder.
    
    Ref: https://www.originlab.com/doc/en/X-Function/ref/pe_save
    """
    cmd = f"pe_save fname:=\"{fname}\""

    if path is not None:
        cmd += f" path:={path}"

    return cmd


# ============================================================
# XFunction Wrappers
# ============================================================

# --- add_graph_to_graph ---
def add_graph_to_graph(igname: str, ogname: Optional[str] = None) -> str:
    """
    Paste a selected graph as an EMF object onto a layout window.

    ### igname
        Short name of the graph to be pasted.

    ### ogname
        Short name of the layout window to add onto.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/add_graph_to_graph
    """
    cmd = f"add_graph_to_graph igname:={igname}"
    if ogname is not None:
        cmd += f" ogname:={ogname}"
    return cmd


# --- add_table_to_graph ---
def add_table_to_graph(
    igl: str = "<active>",
    cols: int = 5,
    rows: int = 8,
    title: Optional[str] = None,
    col: int = 0,
) -> str:
    """
    Add a linked table to a graph or layout.

    ### igl
        Graph layer to add the table to.

    ### cols
        Number of columns in the table.

    ### rows
        Number of rows in the table.

    ### title
        Title of the table.

    ### col
        Whether to show the column label row (1=show, 0=hide).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/add_table_to_graph
    """
    cmd = f"add_table_to_graph igl:={igl} cols:={cols} rows:={rows} col:={col}"
    if title is not None:
        cmd += f" title:=\"{title}\""
    return cmd


# --- add_wks_to_graph ---
def add_wks_to_graph(iwname: str, owname: Optional[str] = None) -> str:
    """
    Paste a worksheet as an EMF object onto a layout window.

    ### iwname
        Name of the worksheet to be pasted.

    ### owname
        Name of the layout window to add onto.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/add_wks_to_graph
    """
    cmd = f"add_wks_to_graph iwname:={iwname}"
    if owname is not None:
        cmd += f" owname:={owname}"
    return cmd


# --- add_xyscale_obj ---
def add_xyscale_obj(
    igl: str = "<active>",
    xscale: float = 10,
    xUnits: str = "X Scale",
    yscale: float = 10,
    yUnits: str = "Y Scale",
    left: Optional[float] = None,
    top: Optional[float] = None,
) -> str:
    """
    Add a new XY scale object to a graph layer.

    ### igl
        Graph layer to add the scale object to.

    ### xscale
        Percentage of X scale.

    ### xUnits
        Units label for X scale.

    ### yscale
        Percentage of Y scale.

    ### yUnits
        Units label for Y scale.

    ### left
        Horizontal position (% of layer) from the left edge.

    ### top
        Vertical position (% of layer) from the top edge.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/add_xyscale_obj
    """
    cmd = (
        f"add_xyscale_obj igl:={igl}"
        f" xscale:={xscale} xUnits:=\"{xUnits}\""
        f" yscale:={yscale} yUnits:=\"{yUnits}\""
    )
    if left is not None:
        cmd += f" left:={left}"
    if top is not None:
        cmd += f" top:={top}"
    return cmd


# --- addline ---
def addline(
    type: int = 0,
    value: Optional[float] = None,
    color: int = 0,
    style: int = 0,
    location: int = 0,
    select: int = 1,
    move: int = 1,
    name: Optional[str] = None,
) -> str:
    """
    Draw a vertical or horizontal line on a graph layer.

    ### type
        0=Vertical, 1=Horizontal.

    ### value
        Position of the line in axis units.

    ### color
        Color of the line (Origin color index).

    ### style
        Line style (0=solid, 1=dash, etc.).

    ### location
        Label location (0=top, 1=bottom, 2=left, 3=right).

    ### select
        Allow the line to be selectable (1=yes, 0=no).

    ### move
        Allow the line to be moved on the graph (1=yes, 0=no).

    ### name
        Name for the line object.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/addline
    """
    cmd = f"addline type:={type} color:={color} style:={style} location:={location} select:={select} move:={move}"
    if value is not None:
        cmd += f" value:={value}"
    if name is not None:
        cmd += f" name:=\"{name}\""
    return cmd


# --- addsheet ---
def addsheet(
    book: str = "<active>",
    fname: Optional[str] = None,
    index: int = 0,
    active: int = 1,
) -> str:
    """
    Add a worksheet or matrix sheet to a workbook or matrix book.

    ### book
        The workbook or matrix book to add the sheet to.

    ### fname
        Template file name (.OTW, .OTM, .OGW, or .OGM).
        If 0, all sheets from the template are added.

    ### index
        Index of the sheet in the template to add (0=all sheets).

    ### active
        Whether to activate the newly added sheet (1=yes, 0=no).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/addsheet
    """
    cmd = f"addsheet book:={book} index:={index} active:={active}"
    if fname is not None:
        cmd += f" fname:=\"{fname}\""
    return cmd


# --- axis_scrollbar ---
def axis_scrollbar(
    axis: str = "bottom",
    begin: Optional[float] = None,
    end: Optional[float] = None,
    layer: Optional[int] = None,
    rescale: int = 0,
) -> str:
    """
    Add a scrollbar object to a graph for easy zoom and pan.

    ### axis
        Position of the axis scrollbar: bottom, top, right, left.

    ### begin
        Initial scale value for the scrollbar range.

    ### end
        Final scale value for the scrollbar range.

    ### layer
        Index of the layer to add the scrollbar to.

    ### rescale
        Whether to rescale the layer when adding the scrollbar (1=yes, 0=no).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/axis_scrollbar
    """
    cmd = f"axis_scrollbar axis:={axis} rescale:={rescale}"
    if begin is not None:
        cmd += f" begin:={begin}"
    if end is not None:
        cmd += f" end:={end}"
    if layer is not None:
        cmd += f" layer:={layer}"
    return cmd


# --- axis_scroller ---
def axis_scroller(
    igl: str = "<active>",
    top: int = 0,
    xsize: float = 1,
    ysize: float = 1,
) -> str:
    """
    Add a pair of inverted triangles on the X axis for scale change.

    ### igl
        Graph layer to add the scroller to.

    ### top
        0=bottom X axis (default), 1=top X axis.

    ### xsize
        Horizontal size of the triangles.

    ### ysize
        Vertical size of the triangles.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/axis_scroller
    """
    return f"axis_scroller igl:={igl} top:={top} xsize:={xsize} ysize:={ysize}"


# --- colcopy ---
def colcopy(
    irng: str = "<active>",
    orng: str = "<new>!<new>",
    data: int = 1,
    interval: int = 0,
    hidden: int = 1,
    mask: int = 0,
    format: int = 1,
    lname: int = 1,
    units: int = 1,
    comments: int = 1,
) -> str:
    """
    Copy data, labels, and format between worksheet columns.

    ### irng
        Source column(s) to copy.

    ### orng
        Destination column(s) for copied data.

    ### data
        Copy data (1=yes, 0=no).

    ### interval
        Copy sampling intervals (1=yes, 0=no).

    ### hidden
        Ignore hidden rows (1=yes, 0=no).

    ### mask
        Copy mask status (1=yes, 0=no).

    ### format
        Copy column data format (1=yes, 0=no).

    ### lname
        Copy Long Name (1=yes, 0=no).

    ### units
        Copy Units (1=yes, 0=no).

    ### comments
        Copy Comments (1=yes, 0=no).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/colcopy
    """
    return (
        f"colcopy irng:={irng} orng:={orng}"
        f" data:={data} interval:={interval} hidden:={hidden}"
        f" mask:={mask} format:={format}"
        f" lname:={lname} units:={units} comments:={comments}"
    )


# --- coldesig ---
def coldesig(rng: str = "<active>", desig: Optional[str] = None) -> str:
    """
    Set the plot designation (X/Y/Z/E/L/G etc.) of worksheet columns.

    ### rng
        The columns to set plot designation for.

    ### desig
        Plot designation: X, Y, Z, M (X error), E (Y error),
        L (Label), G (Group), S (Subject), N (Disregard).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/coldesig
    """
    cmd = f"coldesig rng:={rng}"
    if desig is not None:
        cmd += f" desig:={desig}"
    return cmd


# --- colhide ---
def colhide(irng: str = "<active>", operation: int = 0) -> str:
    """
    Hide or show specified worksheet columns.

    ### irng
        Column range to hide or show.

    ### operation
        0=hide, 1=unhide, 2=show all columns in the worksheet.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/colhide
    """
    return f"colhide irng:={irng} operation:={operation}"


# --- colint ---
def colint(
    rng: str = "<active>",
    x0: float = 1,
    inc: float = 1,
    units: Optional[str] = None,
    lname: Optional[str] = None,
) -> str:
    """
    Set sampling interval for selected Y columns.

    ### rng
        The XYRange of data to set the sampling interval for.

    ### x0
        Initial X value.

    ### inc
        X increment.

    ### units
        Units label.

    ### lname
        Long name.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/colint
    """
    cmd = f"colint rng:={rng} x0:={x0} inc:={inc}"
    if units is not None:
        cmd += f" units:=\"{units}\""
    if lname is not None:
        cmd += f" lname:=\"{lname}\""
    return cmd


# --- coljoinbydesig ---
def coljoinbydesig(rng: str = "<active>", rd: str = "<new>") -> str:
    """
    Join columns with the same plot designation into one column group.

    ### rng
        Input data range.

    ### rd
        Output data range.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/coljoinbydesig
    """
    return f"coljoinbydesig rng:={rng} rd:={rd}"


# --- colmask ---
def colmask(
    irng: str = "<active>",
    cond: int = 0,
    nSD: float = 1,
    val: Optional[float] = None,
    abs: int = 0,
    keep: int = 0,
) -> str:
    """
    Mask cells based on a condition.

    ### irng
        Range to be masked.

    ### cond
        Masking condition: 0=sd, 1=gt, 2=ge, 3=lt, 4=le, 5=eq, 6=iqr.

    ### nSD
        Multiple of SD/IQR (available when cond is sd or iqr).

    ### val
        Comparison threshold (available when cond is not sd or iqr).

    ### abs
        Use absolute values to test the condition (1=yes, 0=no).

    ### keep
        Preserve existing masks (1=yes, 0=replace).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/colmask
    """
    cmd = f"colmask irng:={irng} cond:={cond} nSD:={nSD} abs:={abs} keep:={keep}"
    if val is not None:
        cmd += f" val:={val}"
    return cmd


# --- colmove ---
def colmove(
    rng: str = "<active>",
    operation: str = "first",
    col: Optional[str] = None,
) -> str:
    """
    Move selected column(s) to a specified position.

    ### rng
        Column(s) to be moved.

    ### operation
        Movement action: first, last, left, right, or pos (before specific column).

    ### col
        Target column (only used when operation=pos).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/colmove
    """
    cmd = f"colmove rng:={rng} operation:={operation}"
    if col is not None:
        cmd += f" col:={col}"
    return cmd


# --- colreverse ---
def colreverse(rng: str = "<active>") -> str:
    """
    Reverse the order of rows within a selected column range.

    ### rng
        Range to reverse.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/colreverse
    """
    return f"colreverse rng:={rng}"


# --- colshowx ---
def colshowx(rng: str = "<active>", clear: int = 0) -> str:
    """
    Show the sampling interval of Y columns as explicit X columns.

    ### rng
        Y columns to show the sampling intervals for.

    ### clear
        Clear the sampling interval information after generating X columns (1=yes, 0=no).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/colshowx
    """
    return f"colshowx rng:={rng} clear:={clear}"


# --- colsplit ---
def colsplit(
    irng: str = "<active>",
    method: int = 0,
    nrows: int = 2,
    ref: Optional[str] = None,
    rd: str = "[<input>]<new>",
) -> str:
    """
    Split one column into multiple columns by row index grouping.

    ### irng
        Input data column(s).

    ### method
        0=By Every Nth Row, 1=By Sequential N Rows, 2=By Reference Column.

    ### nrows
        N value used in the method.

    ### ref
        Reference column (only when method=2).

    ### rd
        Output result location.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/colsplit
    """
    cmd = f"colsplit irng:={irng} method:={method} nrows:={nrows} rd:={rd}"
    if ref is not None:
        cmd += f" ref:={ref}"
    return cmd


# --- colswap ---
def colswap(rng: str = "<active>") -> str:
    """
    Swap the positions of two selected columns.

    ### rng
        Columns to be swapped. If more than two columns are selected,
        the leftmost and rightmost are swapped.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/colswap
    """
    return f"colswap rng:={rng}"


# --- dlgChkList ---
def dlgChkList(
    inames: str = "Case 1|Case 2|Case 3",
    title: str = "Checkbox List",
    desc: str = "Please select the needed cases",
    osel: Optional[str] = None,
    olist: Optional[str] = None,
) -> str:
    """
    Display a checklist dialog and output selected indices.

    ### inames
        Checkbox labels separated by '|'.

    ### title
        Dialog title bar text.

    ### desc
        Description text displayed above the checkboxes.

    ### osel
        Output vector variable name for selected indices.

    ### olist
        Output string variable name for comma-separated selected indices.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/dlgChkList
    """
    cmd = f"dlgChkList inames:=\"{inames}\" title:=\"{title}\" desc:=\"{desc}\""
    if osel is not None:
        cmd += f" osel:={osel}"
    if olist is not None:
        cmd += f" olist:={olist}"
    return cmd


# --- filltext ---
def filltext(group: int = 26, rng: str = "<active>") -> str:
    """
    Fill worksheet cells with random letters from a specified alphabet group.

    ### group
        Number of letters available for filling (e.g., 4 = A, B, C, D).

    ### rng
        Destination range to fill with random text.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/filltext
    """
    return f"filltext group:={group} rng:={rng}"


# --- findrelated ---
def findrelated(
    id: Optional[str] = None,
    txt: Optional[str] = None,
    title: Optional[str] = None,
    desc: Optional[str] = None,
    delim: int = 0,
) -> str:
    """
    Find graphs, analysis results, and worksheets related to selected data.

    ### id
        Column to use as the Record ID.

    ### txt
        Column to use as the keywords.

    ### title
        Column to use as the title.

    ### desc
        Column to use as the description.

    ### delim
        Delimiter: 0=comma, 1=semicolon, 2=pipe, 3=whitespace.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/findrelated
    """
    cmd = f"findrelated delim:={delim}"
    if id is not None:
        cmd += f" id:={id}"
    if txt is not None:
        cmd += f" txt:={txt}"
    if title is not None:
        cmd += f" title:={title}"
    if desc is not None:
        cmd += f" desc:={desc}"
    return cmd


# --- findreplace ---
def findreplace(option: int = 0, str: Optional[str] = None) -> str:
    """
    Open the Find and Replace dialog.

    ### option
        0=open on Find tab, 1=open on Replace tab.

    ### str
        Initial search string.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/findreplace
    """
    cmd = f"findreplace option:={option}"
    if str is not None:
        cmd += f" str:=\"{str}\""
    return cmd


# --- g2layout ---
def g2layout(
    option: Optional[int] = None,
    graphs: Optional[str] = None,
    row: Optional[int] = None,
    col: Optional[int] = None,
    aspectratio: int = 0,
    xgap: int = 5,
    ygap: int = 5,
    leftmg: int = 15,
    rightmg: int = 10,
    topmg: int = 10,
    bottommg: int = 15,
    portrait: int = 0,
) -> str:
    """
    Add selected graphs to a new layout page.

    ### option
        Which graphs to add: active page, folder, open graphs, etc.

    ### graphs
        Names of graphs to add (when option=specified).

    ### row
        Number of rows in the layout grid.

    ### col
        Number of columns in the layout grid.

    ### aspectratio
        Keep graph aspect ratio (1=yes, 0=no).

    ### xgap
        Horizontal gap between graphs.

    ### ygap
        Vertical gap between graphs.

    ### leftmg / rightmg / topmg / bottommg
        Margins of the layout page.

    ### portrait
        Page orientation: 0=landscape, 1=portrait.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/g2layout
    """
    cmd = (
        f"g2layout aspectratio:={aspectratio}"
        f" xgap:={xgap} ygap:={ygap}"
        f" leftmg:={leftmg} rightmg:={rightmg}"
        f" topmg:={topmg} bottommg:={bottommg} portrait:={portrait}"
    )
    if option is not None:
        cmd += f" option:={option}"
    if graphs is not None:
        cmd += f" graphs:=\"{graphs}\""
    if row is not None:
        cmd += f" row:={row}"
    if col is not None:
        cmd += f" col:={col}"
    return cmd


# --- get_wks_sel ---
def get_wks_sel(
    str: str = "sel",
    sep: str = "|",
    wks: str = "<active>",
) -> str:
    """
    Get the selection range string from a worksheet.

    ### str
        Output string variable name to hold the selection.

    ### sep
        Separator character for non-contiguous selections.

    ### wks
        Worksheet to get the selection from.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/get_wks_sel
    """
    return f"get_wks_sel str:={str} sep:={sep} wks:={wks}"


# --- gfitp ---
def gfitp(igp: str = "<active>", margin: float = 0, aspect: int = 0) -> str:
    """
    Auto-adjust graph layers to tight-fit within the page.

    ### igp
        Graph page to manipulate.

    ### margin
        Margin as % of page size.

    ### aspect
        Maintain layer aspect ratios (1=yes, 0=no).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/gfitp
    """
    return f"gfitp igp:={igp} margin:={margin} aspect:={aspect}"


# --- gxy2w ---
def gxy2w(
    x: float = 0,
    igp: str = "<active>",
    bname: str = "Points",
    n1: int = 1,
    n2: int = 0,
    xname: Optional[str] = None,
    xunits: Optional[str] = None,
) -> str:
    """
    Extract Y values corresponding to a specified X value from graph data to worksheet.

    ### x
        X value to pick up corresponding Y values for.

    ### igp
        Graph page to extract data from.

    ### bname
        Name of the output workbook.

    ### n1
        First layer to extract from (1-based index).

    ### n2
        Last layer to extract from (0=last layer).

    ### xname
        Long name of the X column in output.

    ### xunits
        Units of the X column in output.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/gxy2w
    """
    cmd = f"gxy2w x:={x} igp:={igp} bname:=\"{bname}\" n1:={n1} n2:={n2}"
    if xname is not None:
        cmd += f" xname:=\"{xname}\""
    if xunits is not None:
        cmd += f" xunits:=\"{xunits}\""
    return cmd


# --- insert_graph_to_layer ---
def insert_graph_to_layer(igname: str) -> str:
    """
    Insert a graph window into the active graph or layout window as a BMP object.

    ### igname
        Short name of the graph window to insert.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/insert_graph_to_layer
    """
    return f"insert_graph_to_layer igname:={igname}"


# --- insert_wks_to_layer ---
def insert_wks_to_layer(name: str, table: int = 0) -> str:
    """
    Insert a worksheet into the active graph layer or layout page.

    ### name
        Name of the worksheet to insert.

    ### table
        0=insert as bitmap, 1=insert as table object.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/insert_wks_to_layer
    """
    return f"insert_wks_to_layer name:={name} table:={table}"


# --- label_layer ---
def label_layer(
    type: int = 0,
    custom: Optional[str] = None,
    pos: int = 0,
) -> str:
    """
    Add labels (letter/number) to graph layers.

    ### type
        0=A (uppercase), 1=a (lowercase), 2=custom.

    ### custom
        Custom label pattern. Use a$/A$ for letters, r$/R$ for Roman numerals,
        n$/# for index numbers.

    ### pos
        Label position: 0=top-left outside, 1=top-left inside,
        2=top-right outside, 3=top-right inside.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/label_layer
    """
    cmd = f"label_layer type:={type} pos:={pos}"
    if custom is not None:
        cmd += f" custom:=\"{custom}\""
    return cmd


# --- layadd ---
def layadd(
    igp: str = "<active>",
    type: str = "normal",
    activate: int = 1,
    linkto: Optional[int] = None,
    xaxis: Optional[int] = None,
    yaxis: Optional[int] = None,
) -> str:
    """
    Add a new layer to a graph.

    ### igp
        Graph page to manipulate.

    ### type
        Layer type: normal, topx, righty, lefty, txry, bxry, inset, insetdata, noxy.

    ### activate
        Activate the newly added layer (1=yes, 0=no).

    ### linkto
        Parent layer index to link to.

    ### xaxis
        X-axis linking method: none, straight, or custom.

    ### yaxis
        Y-axis linking method: none, straight, or custom.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/layadd
    """
    cmd = f"layadd igp:={igp} type:={type} activate:={activate}"
    if linkto is not None:
        cmd += f" linkto:={linkto}"
    if xaxis is not None:
        cmd += f" xaxis:={xaxis}"
    if yaxis is not None:
        cmd += f" yaxis:={yaxis}"
    return cmd


# --- layalign ---
def layalign(
    igp: str = "<active>",
    igl: str = "<active>",
    destlayer: str = "<active>",
    direction: str = "left",
) -> str:
    """
    Align graph layers by their edges.

    ### igp
        Graph page to manipulate.

    ### igl
        Source layer.

    ### destlayer
        Destination layer to align.

    ### direction
        Alignment direction: left, top, right, bottom.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/layalign
    """
    return f"layalign igp:={igp} igl:={igl} destlayer:={destlayer} direction:={direction}"


# --- layarrange ---
def layarrange(
    igp: str = "<active>",
    row: Optional[int] = None,
    col: Optional[int] = None,
    xgap: int = 5,
    ygap: int = 5,
    left: int = 15,
    right: int = 15,
    top: int = 15,
    bottom: int = 15,
) -> str:
    """
    Automatically arrange multiple graph layers.

    ### igp
        Graph page to manipulate.

    ### row
        Number of rows.

    ### col
        Number of columns.

    ### xgap
        Horizontal gap between adjacent layers.

    ### ygap
        Vertical gap between adjacent layers.

    ### left / right / top / bottom
        Margins of the graph.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/layarrange
    """
    cmd = (
        f"layarrange igp:={igp}"
        f" xgap:={xgap} ygap:={ygap}"
        f" left:={left} right:={right} top:={top} bottom:={bottom}"
    )
    if row is not None:
        cmd += f" row:={row}"
    if col is not None:
        cmd += f" col:={col}"
    return cmd


# --- laycolor ---
def laycolor(
    igp: str = "<active>",
    layer: str = "1:0",
    color: Optional[int] = None,
    fill: Optional[int] = None,
    border: Optional[int] = None,
) -> str:
    """
    Set the background color of graph layers.

    ### igp
        Graph page to manipulate.

    ### layer
        Layer(s) to modify (e.g., "1:3" for layers 1 to 3).

    ### color
        Background fill color (Origin color index).

    ### fill
        Color of the space between background and borderline.

    ### border
        Color of the layer borderline.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/laycolor
    """
    cmd = f"laycolor igp:={igp} layer:={layer}"
    if color is not None:
        cmd += f" color:={color}"
    if fill is not None:
        cmd += f" fill:={fill}"
    if border is not None:
        cmd += f" border:={border}"
    return cmd


# --- laycopyscale ---
def laycopyscale(
    igp: str = "<active>",
    igl: str = "Layer1",
    dest: Optional[str] = None,
    axis: int = 0,
) -> str:
    """
    Copy axis scales from a source layer to a destination layer.

    ### igp
        Graph page to manipulate.

    ### igl
        Source layer.

    ### dest
        Destination layer(s).

    ### axis
        Axis scale to copy: 0=All, 1=X, 2=Y, 3=Z.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/laycopyscale
    """
    cmd = f"laycopyscale igp:={igp} igl:={igl} axis:={axis}"
    if dest is not None:
        cmd += f" dest:={dest}"
    return cmd


# --- layextract ---
def layextract(
    igp: str = "<active>",
    layer: str = "1:0",
    keep: int = 1,
    fullpage: int = 0,
) -> str:
    """
    Extract layers from the active graph into new graph windows.

    ### igp
        Graph page to manipulate.

    ### layer
        Layers to extract (e.g., "1:3" or "1:3,5:7").

    ### keep
        Keep the source graph (1=yes, 0=no).

    ### fullpage
        Resize extracted layers to full page (1=yes, 0=no).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/layextract
    """
    return f"layextract igp:={igp} layer:={layer} keep:={keep} fullpage:={fullpage}"


# --- laylink ---
def laylink(
    igp: str = "<active>",
    igl: str = "layer1",
    destlayers: str = "<active>",
    XAxis: int = -1,
    YAxis: int = -1,
    unit: str = "link",
) -> str:
    """
    Link multiple graph layers to a source layer.

    ### igp
        Graph page to manipulate.

    ### igl
        Source layer to link to.

    ### destlayers
        Destination layers to link (use colon notation for ranges).

    ### XAxis
        X axis link: -1=keep current, 0=no link, 1=link.

    ### YAxis
        Y axis link: -1=keep current, 0=no link, 1=link.

    ### unit
        Measurement unit: page, inch, cm, mm, pixel, point, link.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/laylink
    """
    return (
        f"laylink igp:={igp} igl:={igl} destlayers:={destlayers}"
        f" XAxis:={XAxis} YAxis:={YAxis} unit:={unit}"
    )


# --- laymanage ---
def laymanage() -> str:
    """
    Open the Layer Management dialog.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/laymanage
    """
    return "laymanage"


# --- laymplot ---
def laymplot(
    plot: str = "1",
    target: int = 0,
    axis: int = 0,
    show: int = 0,
    rescale: int = 0,
) -> str:
    """
    Move plot(s) from the active layer to another layer.

    ### plot
        Index (or space-separated indices) of the data plot(s) to move.

    ### target
        Index of the destination layer (0=create new right-Y layer).

    ### axis
        Axis linking mode when target=0:
        0=Independent Y Linked X, 1=Independent X Linked Y, 2=Independent XY.

    ### show
        Show linked axis on both sides (1=yes, 0=no).

    ### rescale
        Which layer(s) to rescale: 0=new layer only, 1=source, 2=target, 3=both.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/laymplot
    """
    return f"laymplot plot:={plot} target:={target} axis:={axis} show:={show} rescale:={rescale}"


# --- laymxn ---
def laymxn(
    row: Optional[int] = None,
    col: Optional[int] = None,
    dir: int = 0,
    alternate: int = 0,
    duplayer: int = 0,
    autofit: int = 1,
    xgap: Optional[float] = None,
    ygap: Optional[float] = None,
    igp: str = "<active>",
) -> str:
    """
    Arrange graph layers in an m×n grid.

    ### row
        Number of rows.

    ### col
        Number of columns.

    ### dir
        Arrangement order: 0=horizontal first, 1=vertical first.

    ### alternate
        Show axis ticks/labels alternately (stack graphs only).

    ### duplayer
        Duplicate layers when grid exceeds current layer count.

    ### autofit
        Keep layer size and auto-fit page size (1=yes, 0=no).

    ### xgap
        Horizontal gap between adjacent layers.

    ### ygap
        Vertical gap between adjacent layers.

    ### igp
        Graph page to manipulate.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/laymxn
    """
    cmd = (
        f"laymxn igp:={igp} dir:={dir} alternate:={alternate}"
        f" duplayer:={duplayer} autofit:={autofit}"
    )
    if row is not None:
        cmd += f" row:={row}"
    if col is not None:
        cmd += f" col:={col}"
    if xgap is not None:
        cmd += f" xgap:={xgap}"
    if ygap is not None:
        cmd += f" ygap:={ygap}"
    return cmd


# --- laysetfont ---
def laysetfont(option: int = 1) -> str:
    """
    Set font scaling for graph layer(s).

    ### option
        1=active layer only, 2=all layers in active graph.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/laysetfont
    """
    return f"laysetfont option:={option}"


# --- laysetpos ---
def laysetpos(
    igp: str = "<active>",
    layer: str = "<active>",
    left: Optional[float] = None,
    top: Optional[float] = None,
    unit: int = 0,
) -> str:
    """
    Set the position of graph layer(s).

    ### igp
        Graph page to manipulate.

    ### layer
        Layer(s) to position (use colon notation for ranges).

    ### left
        Position from the page's left edge.

    ### top
        Position from the page's top edge.

    ### unit
        Measurement unit: 0=% of page, 1=inch, 2=cm, 3=mm, 4=pixel, 5=point.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/laysetpos
    """
    cmd = f"laysetpos igp:={igp} layer:={layer} unit:={unit}"
    if left is not None:
        cmd += f" left:={left}"
    if top is not None:
        cmd += f" top:={top}"
    return cmd


# --- laysetratio ---
def laysetratio(
    igp: str = "<active>",
    destlayer: str = "<active>",
    ratio: Optional[float] = None,
) -> str:
    """
    Set the aspect ratio (width/height) of graph layer(s).

    ### igp
        Graph page to manipulate.

    ### destlayer
        Layer(s) to modify (use colon notation for ranges).

    ### ratio
        New width-to-height ratio for each destination layer.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/laysetratio
    """
    cmd = f"laysetratio igp:={igp} destlayer:={destlayer}"
    if ratio is not None:
        cmd += f" ratio:={ratio}"
    return cmd


# --- laysetscale ---
def laysetscale(
    igp: str = "<active>",
    layer: str = "<active>",
    axis: int = 0,
    from_: Optional[float] = None,
    to: Optional[float] = None,
    inc: Optional[float] = None,
    type: Optional[int] = None,
) -> str:
    """
    Set the axis scale of graph layer(s).

    ### igp
        Graph page to manipulate.

    ### layer
        Layer(s) to modify.

    ### axis
        Axis to set: 0=x, 1=y, 2=z.

    ### from_
        First scale value (From).

    ### to
        Last scale value (To).

    ### inc
        Scale increment.

    ### type
        Scale type: 0=linear, 1=log10, 2=probability, etc.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/laysetscale
    """
    cmd = f"laysetscale igp:={igp} layer:={layer} axis:={axis}"
    if from_ is not None:
        cmd += f" from:={from_}"
    if to is not None:
        cmd += f" to:={to}"
    if inc is not None:
        cmd += f" inc:={inc}"
    if type is not None:
        cmd += f" type:={type}"
    return cmd


# --- laysetunit ---
def laysetunit(
    igp: str = "<active>",
    layer: str = "<active>",
    unit: int = 0,
) -> str:
    """
    Set the size unit of graph layer(s).

    ### igp
        Graph page to manipulate.

    ### layer
        Layer(s) to modify (use colon notation for ranges).

    ### unit
        Unit: 0=% of page, 1=inch, 2=cm, 3=mm, 4=pixel, 5=point.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/laysetunit
    """
    return f"laysetunit igp:={igp} layer:={layer} unit:={unit}"


# --- layswap ---
def layswap(igl1: str, igl2: str, reorder: int = 0) -> str:
    """
    Swap the plot positions and order of two graph layers.

    ### igl1
        First graph layer.

    ### igl2
        Second graph layer.

    ### reorder
        Also swap the order of layers (1=yes, 0=no).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/layswap
    """
    return f"layswap igl1:={igl1} igl2:={igl2} reorder:={reorder}"


# --- laytoggle ---
def laytoggle(
    igp: str = "<active>",
    destlayer: str = "<active>",
) -> str:
    """
    Toggle the visibility of left and bottom axes for graph layer(s).

    ### igp
        Graph page to manipulate.

    ### destlayer
        Layer(s) to toggle (use colon notation for ranges).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/laytoggle
    """
    return f"laytoggle igp:={igp} destlayer:={destlayer}"


# --- layzoom ---
def layzoom(igp: str = "<active>", layer: Optional[str] = None) -> str:
    """
    Center zoom on a specified graph layer.

    ### igp
        Graph page to manipulate.

    ### layer
        Layer to zoom.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/layzoom
    """
    cmd = f"layzoom igp:={igp}"
    if layer is not None:
        cmd += f" layer:={layer}"
    return cmd


# --- m2w ---
def m2w(
    im: str = "<active>",
    method: str = "direct",
    xy: str = "xcol",
    ycol: int = 0,
    xcol: int = 0,
    trim: int = 0,
    ow: str = "[<new>]<new>",
) -> str:
    """
    Convert matrix data to a worksheet.

    ### im
        Source matrix object.

    ### method
        Conversion approach: direct or xyz.

    ### xy
        X/Y arrangement when using direct conversion: xcol, ycol, or noxy.

    ### ycol
        Output matrix Y values to first worksheet column (1=yes, 0=no).

    ### xcol
        Output matrix X values to first worksheet column (1=yes, 0=no).

    ### trim
        Trim missing values (1=yes, 0=no).

    ### ow
        Output worksheet.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/m2w
    """
    return f"m2w im:={im} method:={method} xy:={xy} ycol:={ycol} xcol:={xcol} trim:={trim} ow:={ow}"


# --- merge_book ---
def merge_book(
    fld: str = "folder",
    keep: int = 1,
    single: int = 1,
    match: str = "none",
    key: Optional[str] = None,
    rename: str = "lname",
    owp: str = "<new>",
) -> str:
    """
    Merge multiple workbooks into one.

    ### fld
        Which books to merge: folder, recursive, open, or project.

    ### keep
        Keep source workbooks after merging (1=yes, 0=no).

    ### single
        Only merge single-worksheet workbooks (1=yes, 0=no).

    ### match
        Workbook property to match: none, lname, sname, or comments.

    ### key
        Search pattern for matching (supports wildcards).

    ### rename
        How to name merged worksheets: lname, sname, shname, or reset.

    ### owp
        Output workbook page.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/merge_book
    """
    cmd = f"merge_book fld:={fld} keep:={keep} single:={single} match:={match} rename:={rename} owp:={owp}"
    if key is not None:
        cmd += f" key:=\"{key}\""
    return cmd


# --- merge_graph ---
def merge_graph(
    option: Optional[int] = None,
    graphs: Optional[str] = None,
    keep: int = 1,
    arrange: int = 1,
    row: Optional[int] = None,
    col: Optional[int] = None,
    dir: int = 0,
    xgap: int = 2,
    ygap: int = 2,
    leftmg: int = 5,
    rightmg: int = 5,
    topmg: int = 5,
    bottommg: int = 5,
    portrait: int = 2,
    ogp: str = "[<new>]",
) -> str:
    """
    Merge multiple graphs into one graph window.

    ### option
        Graph source: page, folder, recursive, open, embed, project, or specified.

    ### graphs
        Graph names to merge (use /n as separator, for option=specified).

    ### keep
        Keep source graphs (1=yes, 0=no).

    ### arrange
        Rearrange layers from multi-layer graphs (1=yes, 0=no).

    ### row / col
        Grid rows/columns for the merged layout.

    ### dir
        Direction: 0=horizontal first, 1=vertical first.

    ### xgap / ygap
        Horizontal/vertical gap between layers.

    ### leftmg / rightmg / topmg / bottommg
        Page margins.

    ### portrait
        Page orientation: 0=landscape, 1=portrait, 2=auto.

    ### ogp
        Output merged graph destination.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/merge_graph
    """
    cmd = (
        f"merge_graph keep:={keep} arrange:={arrange} dir:={dir}"
        f" xgap:={xgap} ygap:={ygap}"
        f" leftmg:={leftmg} rightmg:={rightmg} topmg:={topmg} bottommg:={bottommg}"
        f" portrait:={portrait} ogp:={ogp}"
    )
    if option is not None:
        cmd += f" option:={option}"
    if graphs is not None:
        cmd += f" graphs:=\"{graphs}\""
    if row is not None:
        cmd += f" row:={row}"
    if col is not None:
        cmd += f" col:={col}"
    return cmd


# --- mg2layout ---
def mg2layout(
    option: Optional[int] = None,
    graphs: Optional[str] = None,
    col: Optional[int] = None,
    row: Optional[int] = None,
    dir: int = 0,
    xgap: int = 0,
    ygap: int = 0,
    xmg: int = 0,
    ymg: int = 0,
) -> str:
    """
    Merge multiple graphs into a new layout page in a grid arrangement.

    ### option
        Which graphs to include: active page, active folder, open graphs, etc.

    ### graphs
        Names of graphs (when option=specified).

    ### col
        Number of columns.

    ### row
        Number of rows.

    ### dir
        0=horizontal first, 1=vertical first.

    ### xgap / ygap
        Horizontal/vertical gap between graphs.

    ### xmg / ymg
        Left/right and top/bottom margins.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/mg2layout
    """
    cmd = f"mg2layout dir:={dir} xgap:={xgap} ygap:={ygap} xmg:={xmg} ymg:={ymg}"
    if option is not None:
        cmd += f" option:={option}"
    if graphs is not None:
        cmd += f" graphs:=\"{graphs}\""
    if col is not None:
        cmd += f" col:={col}"
    if row is not None:
        cmd += f" row:={row}"
    return cmd


# --- mlabelbywcol ---
def mlabelbywcol(
    ms: str = "<active>",
    ref: Optional[str] = None,
    pos: int = 0,
) -> str:
    """
    Label matrix objects using corresponding worksheet column names.

    ### ms
        Matrix layer to label.

    ### ref
        Reference column for matrix labels.

    ### pos
        Label position: 0=Long Name, 1=Comments.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/mlabelbywcol
    """
    cmd = f"mlabelbywcol ms:={ms} pos:={pos}"
    if ref is not None:
        cmd += f" ref:={ref}"
    return cmd


# --- mo2s ---
def mo2s(ims: str = "<active>", omp: str = "<new>") -> str:
    """
    Convert a matrix layer with multiple objects to a matrix page with multiple layers.

    ### ims
        Source matrix layer with multiple matrix objects.

    ### omp
        Output matrix page with multiple matrix layers.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/mo2s
    """
    return f"mo2s ims:={ims} omp:={omp}"


# --- newbook ---
def newbook(
    name: Optional[str] = None,
    sheet: int = -1,
    template: str = "Origin",
    hidden: int = 0,
    mat: int = 0,
    keep: int = 0,
    chkname: int = 1,
) -> str:
    """
    Create a new workbook or matrix book.

    ### name
        Desired name for the workbook.

    ### sheet
        Number of sheets (-1=use template default).

    ### template
        Template name to use.

    ### hidden
        Create hidden workbook (1=yes, 0=no).

    ### mat
        Create matrix book (1=yes, 0=workbook).

    ### keep
        Preserve template sheet names (1=yes, 0=no).

    ### chkname
        Prevent creation if name exists (1=yes, 0=enumerate).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/newbook
    """
    cmd = f"newbook sheet:={sheet} template:=\"{template}\" hidden:={hidden} mat:={mat} keep:={keep} chkname:={chkname}"
    if name is not None:
        cmd += f" name:=\"{name}\""
    return cmd


# --- newpanel ---
def newpanel(
    col: int = 1,
    row: int = 1,
    name: str = "Graph1",
    vg: int = 5,
    hg: int = 5,
    template: str = "origin.otp",
) -> str:
    """
    Create a new graph window with a specified m×n layer arrangement.

    ### col
        Number of columns.

    ### row
        Number of rows.

    ### name
        Desired graph page name.

    ### vg
        Vertical gap between rows (% of page).

    ### hg
        Horizontal gap between columns (% of page).

    ### template
        Graph template for the new page.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/newpanel
    """
    return f"newpanel col:={col} row:={row} name:=\"{name}\" vg:={vg} hg:={hg} template:=\"{template}\""


# --- newsheet ---
def newsheet(
    book: str = "<active>",
    name: str = "Sheet1",
    cols: int = -1,
    rows: int = -1,
    labels: Optional[str] = None,
    xy: Optional[str] = None,
    template: str = "origin",
    use: int = 0,
    mat: int = 0,
    active: int = 1,
) -> str:
    """
    Create a new worksheet.

    ### book
        Workbook to add the sheet to.

    ### name
        Desired sheet name.

    ### cols
        Number of columns (-1=use template default).

    ### rows
        Number of rows (-1=use template default).

    ### labels
        Long names of columns separated by '|'.

    ### xy
        Column designations (X, Y, Z, E, L, M, N).

    ### template
        Template name.

    ### use
        Reuse existing sheet with same name (1=yes, 0=create new).

    ### mat
        Create matrix sheet (1=yes, 0=worksheet).

    ### active
        Activate the new sheet (1=yes, 0=no).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/newsheet
    """
    cmd = (
        f"newsheet book:={book} name:=\"{name}\""
        f" cols:={cols} rows:={rows}"
        f" template:=\"{template}\" use:={use} mat:={mat} active:={active}"
    )
    if labels is not None:
        cmd += f" labels:=\"{labels}\""
    if xy is not None:
        cmd += f" xy:={xy}"
    return cmd


# --- palApply ---
def palApply(
    igl: str = "<active>",
    fname: str = "grayscale",
    stretch: int = 1,
    flip: int = 0,
    scope: int = 1,
) -> str:
    """
    Apply a color palette to a graph layer.

    ### igl
        Graph layer to apply the palette to.

    ### fname
        Path and file name of the palette file.

    ### stretch
        Stretch palette to full color map range (1=yes, 0=no).

    ### flip
        Flip the palette color order (1=yes, 0=no).

    ### scope
        0=active plot only, 1=all plots on the layer.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/palApply
    """
    return f"palApply igl:={igl} fname:=\"{fname}\" stretch:={stretch} flip:={flip} scope:={scope}"


# --- patternD ---
def patternD(
    irng: str = "<active>",
    format: int = 0,
    display: int = 0,
    from_: Optional[float] = None,
    to: Optional[float] = None,
    inc: float = 1,
    unit: int = 2,
    mode: int = 0,
    size: int = 10,
) -> str:
    """
    Detect and convert date/time patterns in worksheet data.

    ### irng
        Range of data columns to fill values to.

    ### format
        0=date, 1=time.

    ### display
        Display format index of the date/time data.

    ### from_
        Start value (Julian date for date; 0-1 for time).

    ### to
        End value.

    ### inc
        Increment within the generated data sequence.

    ### unit
        Increment unit: second, minute, hour, day, week, month, year, workday, quarter.

    ### mode
        0=repeat, 1=random.

    ### size
        Total number of values (random mode only).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/patternD
    """
    cmd = f"patternD irng:={irng} format:={format} display:={display} inc:={inc} unit:={unit} mode:={mode} size:={size}"
    if from_ is not None:
        cmd += f" from:={from_}"
    if to is not None:
        cmd += f" to:={to}"
    return cmd


# --- patternN ---
def patternN(
    irng: str = "<active>",
    from_: float = 1,
    to: float = 10,
    inc: float = 1,
    mode: int = 0,
    onerepeat: int = 1,
    seqrepeat: int = 1,
    size: int = 10,
) -> str:
    """
    Detect and convert numeric patterns in worksheet data.

    ### irng
        Range of data columns to fill.

    ### from_
        Starting value of the generated sequence.

    ### to
        Ending value of the generated sequence.

    ### inc
        Increment within the generated sequence.

    ### mode
        0=repeat, 1=random.

    ### onerepeat
        Repetitions per value (repeat mode only).

    ### seqrepeat
        Repetitions of whole sequence (repeat mode only).

    ### size
        Total dataset size (random mode only).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/patternN
    """
    return (
        f"patternN irng:={irng} from:={from_} to:={to} inc:={inc}"
        f" mode:={mode} onerepeat:={onerepeat} seqrepeat:={seqrepeat} size:={size}"
    )


# --- patternT ---
def patternT(
    irng: str = "<active>",
    text: Optional[str] = None,
    mode: int = 0,
    onerepeat: int = 1,
    seqrepeat: int = 1,
    size: int = 10,
) -> str:
    """
    Detect and convert text patterns in worksheet data.

    ### irng
        Range of data columns to fill.

    ### text
        Value series to fill (text or numbers, values separated by spaces).

    ### mode
        0=repeat (by order), 1=random.

    ### onerepeat
        Repetitions per value (repeat mode only).

    ### seqrepeat
        Repetitions of whole sequence (repeat mode only).

    ### size
        Total dataset size (random mode only).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/patternT
    """
    cmd = f"patternT irng:={irng} mode:={mode} onerepeat:={onerepeat} seqrepeat:={seqrepeat} size:={size}"
    if text is not None:
        cmd += f" text:=\"{text}\""
    return cmd


# --- plot_2x ---
def plot_2x(
    iy: str = "<active>",
    type: int = 2,
    doublex: int = 0,
    template: Optional[str] = None,
    rd: Optional[str] = None,
) -> str:
    """
    Create a double X-axis graph.

    ### iy
        Input data columns to plot.

    ### type
        Plot type: 0=Line, 1=Scatter, 2=Line+Symbol, 3=Column.

    ### doublex
        Double X mode: 0=Independent XY, 1=Custom Formula, 2=Arrhenius.

    ### template
        Graph template file.

    ### rd
        Output plot data worksheet.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/plot_2x
    """
    cmd = f"plot_2x iy:={iy} type:={type} doublex:={doublex}"
    if template is not None:
        cmd += f" template:=\"{template}\""
    if rd is not None:
        cmd += f" rd:={rd}"
    return cmd


# --- plot_BA ---
def plot_BA(
    irng: str = "<active>",
    xaxis: str = "Mean",
    yaxis: int = 1,
    sd: float = 1.96,
    conf: float = 95,
    rd: str = "[<input>]<new>",
) -> str:
    """
    Create a Bland-Altman (method comparison) plot.

    ### irng
        Input data ranges of Method1, Method2, and Subject (in order).

    ### xaxis
        X-axis values: Mean, m1 (Method 1), m2 (Method 2), geo (Geometric Mean).

    ### yaxis
        Y-axis values: 1=diffp (difference percent), 0=diff (difference), 2=ratio.

    ### sd
        Multiplier factor of Standard Deviation for Limits of Agreement.

    ### conf
        Confidence level (%).

    ### rd
        Output results location.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/plot_BA
    """
    return f"plot_BA irng:={irng} xaxis:={xaxis} yaxis:={yaxis} sd:={sd} conf:={conf} rd:={rd}"


# --- plot_bygroup ---
def plot_bygroup(
    iy: str = "<active>",
    type: int = 0,
    template: Optional[str] = None,
    hgroup: Optional[str] = None,
    vgroup: Optional[str] = None,
    color: Optional[str] = None,
    arrange: int = 0,
    link: int = 1,
    rd: str = "[<input>]<new>",
) -> str:
    """
    Create a multi-panel graph based on group information.

    ### iy
        Input data range (multiple XY ranges supported).

    ### type
        Plot type: 0=line, 1=scatter, 2=line+symbol, 3=column, 4=bar.

    ### template
        Graph template for the group plot.

    ### hgroup
        Group column(s) for horizontal panel arrangement.

    ### vgroup
        Group column(s) for vertical panel arrangement.

    ### color
        Column used for color-mapping data points.

    ### arrange
        Layer arrangement: 0=grid, 1=overlap.

    ### link
        Link layers to Layer1 (1=yes, 0=no).

    ### rd
        Output worksheet for unstacked data.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/plot_bygroup
    """
    cmd = f"plot_bygroup iy:={iy} type:={type} arrange:={arrange} link:={link} rd:={rd}"
    if template is not None:
        cmd += f" template:=\"{template}\""
    if hgroup is not None:
        cmd += f" hgroup:={hgroup}"
    if vgroup is not None:
        cmd += f" vgroup:={vgroup}"
    if color is not None:
        cmd += f" color:={color}"
    return cmd


# --- plot_chord ---
def plot_chord(
    irng: str = "<active>",
    format: int = 0,
    rowpos: Optional[int] = None,
    colpos: int = 0,
) -> str:
    """
    Create a chord (string) diagram.

    ### irng
        Input Z-values or existing virtual matrix.

    ### format
        Data layout: 0=link row to column, 1=link column to row.

    ### rowpos
        Row position for Y/X values (0=none, 1=1st row in selection, etc.).

    ### colpos
        Column position for X/Y values (0=none, 1=1st column, etc.).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/plot_chord
    """
    cmd = f"plot_chord irng:={irng} format:={format} colpos:={colpos}"
    if rowpos is not None:
        cmd += f" rowpos:={rowpos}"
    return cmd


# --- plot_durov ---
def plot_durov(
    irng: str = "<active>",
    id: Optional[str] = None,
    tds: Optional[str] = None,
    pH: Optional[str] = None,
) -> str:
    """
    Create a Durov trilinear diagram.

    ### irng
        Input data range (6 columns in order: Ca, Mg, Na+K, Cl, SO4, CO3+HCO3).

    ### id
        Sample ID column (controls symbol color and shape).

    ### tds
        Total Dissolved Solids column.

    ### pH
        pH values column.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/plot_durov
    """
    cmd = f"plot_durov irng:={irng}"
    if id is not None:
        cmd += f" id:={id}"
    if tds is not None:
        cmd += f" tds:={tds}"
    if pH is not None:
        cmd += f" pH:={pH}"
    return cmd


# --- plot_gboxindexed ---
def plot_gboxindexed(
    irng: str = "<active>",
    group: Optional[str] = None,
    template: Optional[str] = None,
    theme: Optional[str] = None,
    rd: str = "[<input>]<new>",
) -> str:
    """
    Create a grouped box chart from indexed data.

    ### irng
        Input data range.

    ### group
        Grouping range.

    ### template
        Graph template to apply.

    ### theme
        Built-in graph theme (e.g., Box_I-shaped, Box_Violin).

    ### rd
        Output location for unstacked data.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/plot_gboxindexed
    """
    cmd = f"plot_gboxindexed irng:={irng} rd:={rd}"
    if group is not None:
        cmd += f" group:={group}"
    if template is not None:
        cmd += f" template:={template}"
    if theme is not None:
        cmd += f" theme:=\"{theme}\""
    return cmd


# --- plot_gdot ---
def plot_gdot(
    irng: str = "<active>",
    group: Optional[str] = None,
    order: int = 0,
    stack: int = 4,
    template: Optional[str] = None,
) -> str:
    """
    Create a grouped dot plot (with optional stacking).

    ### irng
        Source Y data column(s).

    ### group
        Grouping column(s).

    ### order
        Plot order: 0=Input variables first, 1=Group info first.

    ### stack
        Stack control: 0=All, 1=Input variables, 2=Last group, 3=All groups, 4=None.

    ### template
        Graph template file.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/plot_gdot
    """
    cmd = f"plot_gdot irng:={irng} order:={order} stack:={stack}"
    if group is not None:
        cmd += f" group:={group}"
    if template is not None:
        cmd += f" template:=\"{template}\""
    return cmd


# --- plot_gfloatbar ---
def plot_gfloatbar(
    iy: str = "<active>",
    subgroup: int = 0,
    size: int = 2,
    label: str = "L",
    type: int = 2,
    rd: str = "[<input>]<new>",
) -> str:
    """
    Create a grouped floating bar chart.

    ### iy
        Input XYRange data.

    ### subgroup
        Subgroup method: 0=by size, 1=by column label.

    ### size
        Subgroup size (when subgroup=0).

    ### label
        Column label row for subgroup (when subgroup=1): L=Long Name, U=Units, C=Comments.

    ### type
        Plot type: 0=column, 1=bar.

    ### rd
        Output location for calculated data.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/plot_gfloatbar
    """
    return f"plot_gfloatbar iy:={iy} subgroup:={subgroup} size:={size} label:={label} type:={type} rd:={rd}"


# --- plot_gindexed ---
def plot_gindexed(
    iy: str = "<active>",
    group: Optional[str] = None,
    plottype: int = 0,
    template: Optional[str] = None,
    rd: str = "[<input>]<new>",
) -> str:
    """
    Create a grouped column plot from indexed data.

    ### iy
        Input data range (one or more Y columns, optionally with Y error).

    ### group
        Grouped column data range.

    ### plottype
        0=column, 1=bar, 2=scatter.

    ### template
        Graph template to use.

    ### rd
        Output location for calculated data.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/plot_gindexed
    """
    cmd = f"plot_gindexed iy:={iy} plottype:={plottype} rd:={rd}"
    if group is not None:
        cmd += f" group:={group}"
    if template is not None:
        cmd += f" template:=\"{template}\""
    return cmd


# --- plot_group ---
def plot_group(
    iy: str = "<active>",
    type: int = 0,
    horz: Optional[str] = None,
    vert: Optional[str] = None,
    color: Optional[str] = None,
    template: Optional[str] = None,
    ogp: str = "[<new template:=grouped>]",
) -> str:
    """
    Create a trellis plot organized by group variables.

    ### iy
        Input range.

    ### type
        Plot type: 0=Line, 1=Scatter, 2=Line+Symbol, 3=Column, 4=Bar, etc.

    ### horz
        Group column(s) for horizontal panel arrangement.

    ### vert
        Group column(s) for vertical panel arrangement.

    ### color
        Column used for color-mapping.

    ### template
        Graph template file.

    ### ogp
        Output graph page.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/plot_group
    """
    cmd = f"plot_group iy:={iy} type:={type} ogp:={ogp}"
    if horz is not None:
        cmd += f" horz:={horz}"
    if vert is not None:
        cmd += f" vert:={vert}"
    if color is not None:
        cmd += f" color:={color}"
    if template is not None:
        cmd += f" template:=\"{template}\""
    return cmd


# --- plot_heatmapxy ---
def plot_heatmapxy(
    iy: str = "<active>",
    stats: int = 5,
    base: int = 1,
    rd: str = "[<input>]<new>",
    template: str = "Heat_Map",
) -> str:
    """
    Create a heatmap from XY data.

    ### iy
        Input data range.

    ### stats
        Quantity to compute: 0=min, 1=max, 2=mean, 3=median, 4=sum, 5=count, 6=percent.

    ### base
        Column to compute quantity for: 1=X, 2=Y.

    ### rd
        Output worksheet for binned results.

    ### template
        Graph template to use.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/plot_heatmapxy
    """
    return f"plot_heatmapxy iy:={iy} stats:={stats} base:={base} rd:={rd} template:=\"{template}\""


# --- plot_heatmapxyz ---
def plot_heatmapxyz(
    iz: str = "<active>",
    stats: int = 2,
    base: int = 2,
    rd: str = "[<input>]<new>",
    template: str = "Heat_Map",
) -> str:
    """
    Create a heatmap from XYZ data.

    ### iz
        Input XYZ data range.

    ### stats
        Quantity to compute: 0=min, 1=max, 2=mean, 3=median, 4=sum, 5=count, 6=percent.

    ### base
        Column to compute quantity for: 1=X, 2=Y, 3=Z.

    ### rd
        Output worksheet for binned results.

    ### template
        Graph template to use.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/plot_heatmapxyz
    """
    return f"plot_heatmapxyz iz:={iz} stats:={stats} base:={base} rd:={rd} template:=\"{template}\""


# --- plot_lineser ---
def plot_lineser(irng: str = "<active>", rd: str = "[<new>]<new>") -> str:
    """
    Create a graph connecting 2-3 series of Y values with lines.

    ### irng
        Input data range.

    ### rd
        Output location for calculated data.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/plot_lineser
    """
    return f"plot_lineser irng:={irng} rd:={rd}"


# --- plot_marginal ---
def plot_marginal(
    iy: str = "<active>",
    type: int = 0,
    gap: float = 0,
    size: float = 33.3,
    rugs: int = 0,
) -> str:
    """
    Create a main graph with marginal distribution plots attached.

    ### iy
        Input data.

    ### type
        Main layer plot type: 0=scatter, 1=scatter+linear regression, 2=kernel density.

    ### gap
        Gap between the top/right layer and the main layer.

    ### size
        Size of the top and right layer as percentage.

    ### rugs
        Show axis rugs for the main layer (1=yes, 0=no).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/plot_marginal
    """
    return f"plot_marginal iy:={iy} type:={type} gap:={gap} size:={size} rugs:={rugs}"


# --- plot_matrix ---
def plot_matrix(
    irng: str = "<active>",
    group: Optional[str] = None,
    ellipse: int = 0,
    fit: int = 0,
    missing: int = 0,
    rd: str = "[<input>]<new>",
) -> str:
    """
    Create a scatter matrix (Scatter Matrix) for multiple columns.

    ### irng
        Input data range (at least two columns).

    ### group
        Grouping range.

    ### ellipse
        Add confidence ellipse (1=yes, 0=no).

    ### fit
        Perform linear fit to each pair (1=yes, 0=no).

    ### missing
        Exclude rows with missing values listwise (1=yes, 0=no).

    ### rd
        Output location for calculated data.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/plot_matrix
    """
    cmd = f"plot_matrix irng:={irng} ellipse:={ellipse} fit:={fit} missing:={missing} rd:={rd}"
    if group is not None:
        cmd += f" group:={group}"
    return cmd


# --- plot_mosaic ---
def plot_mosaic(
    row: str = "<active>",
    col: str = "<active>",
    template: Optional[str] = None,
    rdplot: str = "[<input>]<new>",
) -> str:
    """
    Create a mosaic plot for categorical data.

    ### row
        Column range for the result sheet row (shown on X axis).

    ### col
        Column range for the result sheet column (shown on Y axis).

    ### template
        Graph template for the plot.

    ### rdplot
        Output location for calculated data.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/plot_mosaic
    """
    cmd = f"plot_mosaic row:={row} col:={col} rdplot:={rdplot}"
    if template is not None:
        cmd += f" template:=\"{template}\""
    return cmd


# --- plot_multivari ---
def plot_multivari(
    irng: str = "<active>",
    factor: Optional[str] = None,
    mean1: int = 1,
    mean2: int = 1,
    mean3: int = 1,
    points: int = 0,
    grand: int = 0,
    template: str = "multivari",
    rd: str = "[<input>]<new>",
) -> str:
    """
    Create a multi-vari chart.

    ### irng
        Input data column.

    ### factor
        Factor column(s) (up to 3 columns).

    ### mean1 / mean2 / mean3
        Connect mean points for each factor (1=yes, 0=no).

    ### points
        Show individual data points (1=yes, 0=no).

    ### grand
        Show grand mean reference line (1=yes, 0=no).

    ### template
        Graph template file.

    ### rd
        Output location for result table.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/plot_multivari
    """
    cmd = (
        f"plot_multivari irng:={irng}"
        f" mean1:={mean1} mean2:={mean2} mean3:={mean3}"
        f" points:={points} grand:={grand}"
        f" template:=\"{template}\" rd:={rd}"
    )
    if factor is not None:
        cmd += f" factor:={factor}"
    return cmd


# --- plot_network ---
def plot_network(
    irng: str = "<active>",
    dtype: int = 0,
    directed: int = 0,
    weighted: int = 0,
    method: int = 3,
    template: str = "network",
) -> str:
    """
    Create a network diagram (nodes and edges).

    ### irng
        Input data.

    ### dtype
        Data type: 0=adjacency matrix, 1=incidence matrix, 2=edge list.

    ### directed
        Use arrows to show link direction (1=yes, 0=no).

    ### weighted
        Add weighting on the links (1=yes, 0=no).

    ### method
        Node positioning method: 0=Fruchterman-Reingold, 1=Kamada-Kawai,
        2=MDS, 3=Pivot MDS, etc.

    ### template
        Graph template file.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/plot_network
    """
    return f"plot_network irng:={irng} dtype:={dtype} directed:={directed} weighted:={weighted} method:={method} template:=\"{template}\""


# --- plot_paretobin ---
def plot_paretobin(
    datarng: str = "<active>",
    countrng: str = "<active>",
    group: Optional[str] = None,
    cum: int = 1,
    symr: int = 1,
    combine: int = 0,
    rd: str = "[<input>]<new>",
) -> str:
    """
    Create a Pareto chart from binned data.

    ### datarng
        Input range of data categories.

    ### countrng
        Input range of counts.

    ### group
        Grouping column.

    ### cum
        Show cumulative percent plot (1=yes, 0=no).

    ### symr
        Show symbol at right side of bar (1=yes, 0=center).

    ### combine
        Combine smaller values into "Other" (1=yes, 0=no).

    ### rd
        Output location for calculated plot data.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/plot_paretobin
    """
    cmd = f"plot_paretobin datarng:={datarng} countrng:={countrng} cum:={cum} symr:={symr} combine:={combine} rd:={rd}"
    if group is not None:
        cmd += f" group:={group}"
    return cmd


# --- plot_paretoraw ---
def plot_paretoraw(
    irng: str = "<active>",
    group: Optional[str] = None,
    cum: int = 1,
    symr: int = 1,
    combine: int = 0,
    rd: str = "[<input>]<new>",
) -> str:
    """
    Create a Pareto chart from raw data.

    ### irng
        Input data range.

    ### group
        Grouping column.

    ### cum
        Show cumulative percent plot (1=yes, 0=no).

    ### symr
        Show symbol at right side of bar (1=yes, 0=center).

    ### combine
        Combine smaller values into "Other" (1=yes, 0=no).

    ### rd
        Output location for calculated plot data.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/plot_paretoraw
    """
    cmd = f"plot_paretoraw irng:={irng} cum:={cum} symr:={symr} combine:={combine} rd:={rd}"
    if group is not None:
        cmd += f" group:={group}"
    return cmd


# --- plot_prob ---
def plot_prob(
    irng: str = "<active>",
    group: Optional[str] = None,
    overlay: int = 0,
    distr: int = 0,
    estimate: int = 1,
    method: int = 0,
    conf: int = 1,
    level: float = 95,
    xmin: float = 1,
    xmax: float = 99.5,
    rd: str = "[<input>]<new>",
) -> str:
    """
    Create a probability plot or Q-Q plot for a specified distribution.

    ### irng
        Input variable(s) to plot.

    ### group
        Group column to separate input variables.

    ### overlay
        Arrangement: 0=overlay all, 1=overlay groups/different layers for variables.

    ### distr
        Distribution: 0=Normal, 1=Lognormal, 2=Exponential, 3=Weibull, 4=Gamma.

    ### estimate
        Estimate distribution parameters from data (1=yes, 0=no).

    ### method
        Score method: 0=Blom, 1=Benard, 2=Hazen, 3=Van der Waerden, 4=Kaplan-Meier.

    ### conf
        Output confidence band (1=yes, 0=no).

    ### level
        Confidence level (%).

    ### xmin / xmax
        X-axis minimum/maximum values.

    ### rd
        Output location for calculated data.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/plot_prob
    """
    cmd = (
        f"plot_prob irng:={irng} overlay:={overlay} distr:={distr}"
        f" estimate:={estimate} method:={method} conf:={conf} level:={level}"
        f" xmin:={xmin} xmax:={xmax} rd:={rd}"
    )
    if group is not None:
        cmd += f" group:={group}"
    return cmd


# --- plot_rowwise ---
def plot_rowwise(
    iy: str = "<active>",
    type: int = 0,
    template: Optional[str] = None,
    rd: str = "[<input>]<new>",
) -> str:
    """
    Create a plot from row-wise data.

    ### iy
        Input data range (one or more Y columns).

    ### type
        Plot type: 0=line, 1=scatter, 2=line+symbol, 3=column, 4=bar.

    ### template
        Graph template for the plot.

    ### rd
        Output location for calculated data.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/plot_rowwise
    """
    cmd = f"plot_rowwise iy:={iy} type:={type} rd:={rd}"
    if template is not None:
        cmd += f" template:={template}"
    return cmd


# --- plot_transpose_donut ---
def plot_transpose_donut(irng: str = "<active>", rd: str = "<new>") -> str:
    """
    Create a donut chart supporting transposed data.

    ### irng
        Input X and Y value range.

    ### rd
        Output location for calculated data.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/Plot_transpose_donut
    """
    return f"Plot_transpose_donut irng:={irng} rd:={rd}"


# --- plot_vari ---
def plot_vari(
    irng: str = "<active>",
    factor: Optional[str] = None,
    cell: int = 1,
    gmean: int = 1,
    gmedian: int = 0,
    sd: int = 0,
    points: int = 1,
    box: int = 0,
    rline: int = 1,
    template: str = "variability",
    rd: str = "[<input>]<new>",
) -> str:
    """
    Create a multi-vari (variability) chart.

    ### irng
        Input data column.

    ### factor
        Factor column(s).

    ### cell
        Connect mean values of the final factor (1=yes, 0=no).

    ### gmean
        Show grand mean line (1=yes, 0=no).

    ### gmedian
        Show grand median line (1=yes, 0=no).

    ### sd
        Show Standard Deviation chart (1=yes, 0=no).

    ### points
        Show individual data points (1=yes, 0=no).

    ### box
        Show box plot per group (1=yes, 0=no).

    ### rline
        Show whisker line (1=yes, 0=no).

    ### template
        Graph template file.

    ### rd
        Output table location.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/plot_vari
    """
    cmd = (
        f"plot_vari irng:={irng} cell:={cell} gmean:={gmean} gmedian:={gmedian}"
        f" sd:={sd} points:={points} box:={box} rline:={rline}"
        f" template:=\"{template}\" rd:={rd}"
    )
    if factor is not None:
        cmd += f" factor:={factor}"
    return cmd


# --- plot_windrose ---
def plot_windrose(
    iy: str = "<active>",
    xinc: float = 30,
    xintervals: int = 16,
    labels: int = 2,
    spoke: int = 0,
    stats: int = 0,
    rd: str = "<new>",
) -> str:
    """
    Create a wind rose (wind direction/speed) chart.

    ### iy
        Input data range (X=direction, Y=speed).

    ### xinc
        Angle increment for direction bins.

    ### xintervals
        Number of direction sectors.

    ### labels
        Direction label style: 0=N-E-S-W, 1=N-NE-E..., 2=N-NNE-NE...

    ### spoke
        Visual style: 0=Sector, 1=Paddle.

    ### stats
        Quantity to compute: 0=Count, 1=Percent Frequency.

    ### rd
        Output worksheet.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/plot_windrose
    """
    return f"plot_windrose iy:={iy} xinc:={xinc} xintervals:={xintervals} labels:={labels} spoke:={spoke} stats:={stats} rd:={rd}"


# --- plot_xyz3dstack ---
def plot_xyz3dstack(
    iz: str = "<active>",
    split: Optional[str] = None,
    type: int = 0,
    oms: str = "[<new>]",
) -> str:
    """
    Create a 3D stacked graph from XYZ data.

    ### iz
        Input XYZ data (XYZ or XYZZZ for multiple Z).

    ### split
        Column to split XYZ data (for single Z input).

    ### type
        Plot type: 0=heatmap, 1=surface.

    ### oms
        Output matrix with the dataset.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/plot_xyz3dstack
    """
    cmd = f"plot_xyz3dstack iz:={iz} type:={type} oms:={oms}"
    if split is not None:
        cmd += f" split:={split}"
    return cmd


# --- plot_xyzquiver ---
def plot_xyzquiver(
    iz: str = "<active>",
    type: int = 0,
    skip: int = 1,
    xpts: int = 10,
    ypts: int = 10,
    template: str = "multivari",
) -> str:
    """
    Overlay gradient vectors or streamlines on a contour plot.

    ### iz
        Input data column.

    ### type
        0=Contour Line + Gradient Vector, 1=Contour + Streamline.

    ### skip
        Keep one point and skip N-1 points (gradient mode only).

    ### xpts
        Number of points in X direction (streamline mode only).

    ### ypts
        Number of points in Y direction (streamline mode only).

    ### template
        Graph template file.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/plot_xyzquiver
    """
    return f"plot_xyzquiver iz:={iz} type:={type} skip:={skip} xpts:={xpts} ypts:={ypts} template:=\"{template}\""


# --- plotbylabel ---
def plotbylabel(
    iy: str = "<active>",
    group: Optional[str] = None,
    plottype: int = 0,
    template: Optional[str] = None,
    rows: Optional[int] = None,
    cols: Optional[int] = None,
    hide: int = 0,
) -> str:
    """
    Create a multi-layer graph grouped by column labels.

    ### iy
        Input data ranges.

    ### group
        Grouping method using column labels.

    ### plottype
        0=line, 1=scatter, 2=line+symbol, 3=column.

    ### template
        Graph template to apply.

    ### rows
        Number of rows in the layout grid.

    ### cols
        Number of columns in the layout grid.

    ### hide
        Hide the newly created graph (1=yes, 0=no).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/plotbylabel
    """
    cmd = f"plotbylabel iy:={iy} plottype:={plottype} hide:={hide}"
    if group is not None:
        cmd += f" group:={group}"
    if template is not None:
        cmd += f" template:={template}"
    if rows is not None:
        cmd += f" rows:={rows}"
    if cols is not None:
        cmd += f" cols:={cols}"
    return cmd


# --- plotcpack ---
def plotcpack(
    irng: str = "<active>",
    type: int = 0,
    unit: int = 0,
    root: int = 0,
) -> str:
    """
    Create a circular packing graph to visualize hierarchical structure.

    ### irng
        Input data range.

    ### type
        Data structure: 0=edgelist (4 columns), 1=multlevels (multiple level columns).

    ### unit
        Size based on: 0=area, 1=radius.

    ### root
        Add a root node encircling all circles (1=yes, 0=no).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/plotcpack
    """
    return f"plotcpack irng:={irng} type:={type} unit:={unit} root:={root}"


# --- plotgboxraw ---
def plotgboxraw(
    irng: str = "<active>",
    num: int = 1,
    g1: Optional[str] = None,
    g2: Optional[str] = None,
    g3: Optional[str] = None,
    g4: Optional[str] = None,
    g5: Optional[str] = None,
    sort: int = 1,
    theme: Optional[str] = None,
) -> str:
    """
    Create a grouped box plot from raw data.

    ### irng
        Input data range.

    ### num
        Number of grouping variables (0-5).

    ### g1 - g5
        Column label rows for grouping variables 1-5.

    ### sort
        Sort plots by ascending group labels (1=yes, 0=no).

    ### theme
        Built-in graph theme for box chart visualization.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/plotgboxraw
    """
    cmd = f"plotgboxraw irng:={irng} num:={num} sort:={sort}"
    for i, g in enumerate([g1, g2, g3, g4, g5], 1):
        if g is not None:
            cmd += f" g{i}:=\"{g}\""
    if theme is not None:
        cmd += f" theme:=\"{theme}\""
    return cmd


# --- plotgroup ---
def plotgroup(
    iy: str = "<active>",
    pgrp: Optional[str] = None,
    lgrp: Optional[str] = None,
    dgrp: Optional[str] = None,
    arrange: int = 0,
    template: str = "origin",
    type: int = 1,
    sort: int = 0,
    hide: int = 0,
) -> str:
    """
    Organize plots by page, layer, and data group variables.

    ### iy
        Input XYRange data.

    ### pgrp
        Column for page group (creates new graph window when data changes).

    ### lgrp
        Column for layer group (creates new graph layer when data changes).

    ### dgrp
        Column for data group (creates new data group when data changes).

    ### arrange
        Layer arrangement: 0=vertical, 1=horizontal.

    ### template
        Graph template for plot layers.

    ### type
        Plot type: 0=line, 1=scatter, 2=linesymb, 3=column.

    ### sort
        Sort worksheet by groups: 0=none, 1=ascending, 2=descending.

    ### hide
        Hide created graphs (1=yes, 0=no).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/plotgroup
    """
    cmd = f"plotgroup iy:={iy} arrange:={arrange} template:=\"{template}\" type:={type} sort:={sort} hide:={hide}"
    if pgrp is not None:
        cmd += f" pgrp:={pgrp}"
    if lgrp is not None:
        cmd += f" lgrp:={lgrp}"
    if dgrp is not None:
        cmd += f" dgrp:={dgrp}"
    return cmd


# --- plotm ---
def plotm(
    im: str = "<active>",
    plot: int = 103,
    rescale: int = 1,
    ogl: str = "[<new template:=GLMesh>]<new>",
    hide: int = 0,
) -> str:
    """
    Create a 3D graph from a matrix object.

    ### im
        Matrix object as Z matrix for the 3D plot.

    ### plot
        Plot type ID (103=3D color map surface, etc.).

    ### rescale
        Rescale the graph (1=yes, 0=no).

    ### ogl
        Target graph layer or template.

    ### hide
        Hide the newly created graph (1=yes, 0=no).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/plotm
    """
    return f"plotm im:={im} plot:={plot} rescale:={rescale} ogl:={ogl} hide:={hide}"


# --- plotmatrix ---
def plotmatrix(
    irng: str = "<active>",
    ellipse: int = 1,
    conflevel: float = 95,
    fit: int = 0,
    missing: int = 0,
    rd: str = "[<input>]<new>",
) -> str:
    """
    Create a scatter matrix for all selected datasets.

    ### irng
        Input data range.

    ### ellipse
        Add confidence ellipse for each graph (1=yes, 0=no).

    ### conflevel
        Confidence level (%) for the ellipses.

    ### fit
        Perform linear fit to each variable pair (1=yes, 0=no).

    ### missing
        Exclude missing values listwise (1=yes, 0=no).

    ### rd
        Output location for calculated data.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/plotmatrix
    """
    return f"plotmatrix irng:={irng} ellipse:={ellipse} conflevel:={conflevel} fit:={fit} missing:={missing} rd:={rd}"


# --- plotms ---
def plotms(ml: str = "<active>", type: int = 0, hide: int = 0) -> str:
    """
    Create a color fill surface graph from all objects in a matrix sheet.

    ### ml
        Matrix sheet containing all objects to plot.

    ### type
        Plot type: 0=color fill surface, 1=colormap surface,
        2=stacked surface, 3=stacked heatmap.

    ### hide
        Hide the newly created graph (1=yes, 0=no).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/plotms
    """
    return f"plotms ml:={ml} type:={type} hide:={hide}"


# --- plotmultivm ---
def plotmultivm(
    irng: str = "<active>",
    datatype: int = 1,
    format: int = 0,
    type: int = 105,
    ogl: str = "[<new template:=Heat_Map_Multi_var>]",
) -> str:
    """
    Create a 3D or contour graph from multiple virtual matrices.

    ### irng
        Input Z-values to plot.

    ### datatype
        Input type: 1=multiple sheets as virtual matrix.

    ### format
        Data layout: 0=Y across columns, 1=X across columns.

    ### type
        Plot type for 3D or contour graph.

    ### ogl
        Output graph layer specification.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/plotmultivm
    """
    return f"plotmultivm irng:={irng} datatype:={datatype} format:={format} type:={type} ogl:={ogl}"


# --- plotmyaxes ---
def plotmyaxes(
    iy: str = "<active>",
    plottype: int = 0,
    ly: Optional[int] = None,
    ry: Optional[int] = None,
    my: int = 0,
    number: int = 1,
    axiscolor: int = 1,
    ytitle: int = 0,
    topx: int = 0,
    hide: int = 0,
) -> str:
    """
    Customize the axis arrangement of a multi-axis plot.

    ### iy
        Input data range.

    ### plottype
        Plot type: 0=line, 1=scatter, 2=linesymb, 3=custom, 4=column.

    ### ly
        Number of left Y axes.

    ### ry
        Number of right Y axes.

    ### my
        Add middle Y axis (1=yes, 0=no).

    ### number
        Number of plots in each layer.

    ### axiscolor
        Link axis color to corresponding plot (1=yes, 0=no).

    ### ytitle
        Show Y axis title (1=yes, 0=no).

    ### topx
        Show top X axis (1=yes, 0=no).

    ### hide
        Hide the newly created graph (1=yes, 0=no).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/plotmyaxes
    """
    cmd = (
        f"plotmyaxes iy:={iy} plottype:={plottype} my:={my}"
        f" number:={number} axiscolor:={axiscolor}"
        f" ytitle:={ytitle} topx:={topx} hide:={hide}"
    )
    if ly is not None:
        cmd += f" ly:={ly}"
    if ry is not None:
        cmd += f" ry:={ry}"
    return cmd


# --- plotpClamp ---
def plotpClamp(
    iw: str = "<active>",
    tag: int = 0,
    stimulur: int = 0,
    mode: str = "sweeps",
    scrollbar: int = 1,
    ctrl: int = 1,
) -> str:
    """
    Create a graph from pClamp imported data.

    ### iw
        Worksheet with imported pClamp data.

    ### tag
        Show tags when available in the imported file (1=yes, 0=no).

    ### stimulur
        Plot stimulus waveform if present (1=yes, 0=no).

    ### mode
        Display mode: sweeps, Continuous, or Concatenated.

    ### scrollbar
        Include horizontal axis scrollbar (1=yes, 0=no).

    ### ctrl
        Display control panel in signal graph (1=yes, 0=no).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/plotpClamp
    """
    return f"plotpClamp iw:={iw} tag:={tag} stimulur:={stimulur} mode:={mode} scrollbar:={scrollbar} ctrl:={ctrl}"


# --- plotpiper ---
def plotpiper(
    irng: str = "<active>",
    id: Optional[str] = None,
    tds: Optional[str] = None,
    ogl: str = "[<new template:=piper>]<new>",
) -> str:
    """
    Create or extend a Piper plot (water quality trilinear diagram).

    ### irng
        Input XYZXYZ data range for the piper plot.

    ### id
        Sample ID column (controls symbol color/shape and legend labels).

    ### tds
        Total Dissolved Solids column (controls symbol size in Layer 1).

    ### ogl
        Output graph layer.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/plotpiper
    """
    cmd = f"plotpiper irng:={irng} ogl:={ogl}"
    if id is not None:
        cmd += f" id:={id}"
    if tds is not None:
        cmd += f" tds:={tds}"
    return cmd


# --- plotstack ---
def plotstack(
    iy: str = "<active>",
    plottype: int = 0,
    template: str = "stack",
    portrait: int = 1,
    dir: int = 0,
    layer: Optional[int] = None,
    number: str = "1",
    alternate: int = 0,
    link: int = 1,
    hide: int = 0,
) -> str:
    """
    Create a stacked multi-panel graph.

    ### iy
        Input data range.

    ### plottype
        Plot type: 0=line, 1=scatter, 2=linesymb, 3=column, 4=histogram, etc.

    ### template
        Graph template for the page.

    ### portrait
        Orientation: 0=landscape, 1=portrait.

    ### dir
        Stack direction: 0=vertical, 1=horizontal (axes exchanged), 2=horizontal.

    ### layer
        Number of layers.

    ### number
        Number of plots in each layer.

    ### alternate
        Show axis ticks/labels alternately (1=yes, 0=no).

    ### link
        Link the layers (1=yes, 0=no).

    ### hide
        Hide the newly created graph (1=yes, 0=no).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/plotstack
    """
    cmd = (
        f"plotstack iy:={iy} plottype:={plottype} template:=\"{template}\""
        f" portrait:={portrait} dir:={dir}"
        f" number:={number} alternate:={alternate} link:={link} hide:={hide}"
    )
    if layer is not None:
        cmd += f" layer:={layer}"
    return cmd


# --- plotvm ---
def plotvm(
    irng: str = "<active>",
    format: int = 0,
    xtitle: str = "X Title",
    ytitle: str = "Y Title",
    ztitle: str = "Z Title",
    vmname: str = "VM1",
    type: int = 226,
    hide: int = 0,
    ogl: str = "[<new template:=contour>]",
) -> str:
    """
    Plot a worksheet cell range as a virtual matrix in a 3D or contour graph.

    ### irng
        Z-values or an existing virtual matrix to plot.

    ### format
        Data layout: 0=Y across columns, 1=X across columns.

    ### xtitle / ytitle / ztitle
        Axis labels.

    ### vmname
        Virtual matrix name.

    ### type
        Plot type for 3D or contour graphs (226=contour).

    ### hide
        Hide the newly created graph (1=yes, 0=no).

    ### ogl
        Output graph layer specification.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/plotvm
    """
    return (
        f"plotvm irng:={irng} format:={format}"
        f" xtitle:=\"{xtitle}\" ytitle:=\"{ytitle}\" ztitle:=\"{ztitle}\""
        f" vmname:={vmname} type:={type} hide:={hide} ogl:={ogl}"
    )


# --- plotxyz ---
def plotxyz(
    iz: str = "<active>",
    plot: int = 103,
    rescale: int = 1,
    hide: int = 0,
    ogl: str = "<new template:=GLMesh>",
) -> str:
    """
    Create a 3D plot from XYZ data.

    ### iz
        Input XYZ data (X column, Y column, Z column).

    ### plot
        Plot type ID. See Plot Type IDs for XYZ-compatible options.

    ### rescale
        Rescale the newly created plot (1=yes, 0=no).

    ### hide
        Hide the newly created graph (1=yes, 0=no).

    ### ogl
        Target graph layer or template.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/plotxyz
    """
    return f"plotxyz iz:={iz} plot:={plot} rescale:={rescale} hide:={hide} ogl:={ogl}"


# --- setdatafmt ---
def setdatafmt(rng: str = "<active>", type: int = 0) -> str:
    """
    Set the data format (numeric, text, date, etc.) of worksheet columns.

    ### rng
        Target column range.

    ### type
        Data format type (0=numeric, 1=text, 2=time, 3=date,
        4=month, 5=day of week, 9=column bracket).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/setdatafmt
    """
    return f"setdatafmt rng:={rng} type:={type}"


# --- sparklines ---
def sparklines(
    iw: str = "<active>",
    sel: int = 1,
    plottype: int = 0,
    template: str = "sparkline_label",
    row_height: int = 200,
    endpts: int = 1,
    xy: int = 0,
    logy: int = 0,
) -> str:
    """
    Add or update sparkline thumbnail graphs in worksheet column headers.

    ### iw
        Worksheet to manipulate.

    ### sel
        Use selected columns as input (1=yes, 0=use c1/c2 range).

    ### plottype
        Chart style: 0=line, 1=histogram, 2=box.

    ### template
        Graph template for the sparkline.

    ### row_height
        Sparkline row height (% of normal cell height).

    ### endpts
        Mark first and last points (1=yes, 0=no).

    ### xy
        Plot Y against X dataset (1=yes, 0=against row numbers).

    ### logy
        Apply log10 scaling to Y axis (1=yes, 0=no).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/sparklines
    """
    return (
        f"sparklines iw:={iw} sel:={sel} plottype:={plottype}"
        f" template:=\"{template}\" row_height:={row_height}"
        f" endpts:={endpts} xy:={xy} logy:={logy}"
    )


# --- speedmode ---
def speedmode(
    index: str = "page",
    sm: str = "off",
    max: int = 3000,
) -> str:
    """
    Set the speed mode (fast draw) properties of graph layers.

    ### index
        Target scope: page, folder, recursive, open, embed, project, or layer.

    ### sm
        Speed mode: off, low (5000pts), medium (3000pts), high (800pts), or custom.

    ### max
        Maximum points per curve (custom mode only).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/speedmode
    """
    return f"speedmode index:={index} sm:={sm} max:={max}"


# --- splitcol ---
def splitcol(
    irng: str = "<active>",
    sep: str = ",",
    fixed: int = 0,
    width: Optional[int] = None,
) -> str:
    """
    Split worksheet columns by delimiter or fixed width into multiple columns.

    ### irng
        Input column range to split.

    ### sep
        Delimiter character (when fixed=0).

    ### fixed
        Use fixed-width splitting (1=yes, 0=delimiter-based).

    ### width
        Column width for fixed-width splitting.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/splitcol
    """
    cmd = f"splitcol irng:={irng} sep:=\"{sep}\" fixed:={fixed}"
    if width is not None:
        cmd += f" width:={width}"
    return cmd


# --- updateEmbedGraphs ---
def updateEmbedGraphs(
    rng: str = "<active>",
    keepaspect: int = 1,
    axes: int = 0,
    legends: int = 0,
    texts: int = 0,
) -> str:
    """
    Update embedded graph objects in a worksheet.

    ### rng
        Range containing embedded graphs to update.

    ### keepaspect
        Preserve the original aspect ratio (1=yes, 0=no).

    ### axes
        Hide axes (1=yes, 0=no).

    ### legends
        Hide legends (1=yes, 0=no).

    ### texts
        Hide text elements (1=yes, 0=no).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/updateEmbedGraphs
    """
    return f"updateEmbedGraphs rng:={rng} keepaspect:={keepaspect} axes:={axes} legends:={legends} texts:={texts}"


# --- updateSparklines ---
def updateSparklines(
    orng: str = "<active>",
    plottype: int = 0,
    template: str = "sparkline_label",
    keepAspect: int = 0,
    label: int = 1,
    endpts: int = 1,
) -> str:
    """
    Update or add sparklines to worksheet column headers.

    ### orng
        Column(s) receiving the sparkline updates.

    ### plottype
        Sparkline plot type: 0=line, 1=histogram, 2=box.

    ### template
        Template used for the sparkline.

    ### keepAspect
        Scale X and Y axes proportionally (1=yes, 0=no).

    ### label
        Hide annotations in the sparkline (1=yes, 0=no).

    ### endpts
        Mark first and last data points (1=yes, 0=no).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/updateSparklines
    """
    return (
        f"updateSparklines orng:={orng} plottype:={plottype}"
        f" template:=\"{template}\" keepAspect:={keepAspect}"
        f" label:={label} endpts:={endpts}"
    )


# --- v2m ---
def v2m(
    ix: str = "<active>",
    method: str = "v2col",
    index: int = 1,
    direction: str = "row",
    om: str = "<new>",
) -> str:
    """
    Convert a vector (column data) to a matrix.

    ### ix
        Input vector.

    ### method
        Conversion method: v2col (copy to column), v2row (copy to row),
        v2m (fill whole matrix).

    ### index
        Target column/row index (for v2col and v2row methods).

    ### direction
        Fill order for v2m method: row or col.

    ### om
        Output matrix object.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/v2m
    """
    return f"v2m ix:={ix} method:={method} index:={index} direction:={direction} om:={om}"


# --- w2m ---
def w2m(
    iw: str = "<active>",
    xy: str = "noxy",
    ycol: int = 0,
    xcol: int = 0,
    trim: int = 1,
    om: str = "<new>",
) -> str:
    """
    Convert worksheet data directly to a matrix.

    ### iw
        Worksheet to convert.

    ### xy
        X/Y coordinate specification: xcol, ycol, or noxy.

    ### ycol
        Use first column values as Y coordinates (1=yes, 0=no).

    ### xcol
        Use first column values as X coordinates (1=yes, 0=no).

    ### trim
        Trim missing rows or columns (1=yes, 0=no).

    ### om
        Output matrix.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/w2m
    """
    return f"w2m iw:={iw} xy:={xy} ycol:={ycol} xcol:={xcol} trim:={trim} om:={om}"


# --- w2xyz ---
def w2xyz(
    iw: str = "<active>",
    format: int = 2,
    xlabel: int = 0,
    ylabel: int = 0,
    trim: int = 0,
    oz: str = "[<input>]<new>!(<new>,<new>,<new>)",
) -> str:
    """
    Convert worksheet data to XYZ column format.

    ### iw
        Input worksheet.

    ### format
        Data format: 0=X across columns, 1=Y across columns, 2=no X/Y indices.

    ### xlabel
        X values in: 0=none, 1=first data row, 2=column label, 3=first row in selection.

    ### ylabel
        Y values in: same options as xlabel.

    ### trim
        Exclude data points with missing coordinates (1=yes, 0=no).

    ### oz
        Output three-column XYZ range.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/w2xyz
    """
    return f"w2xyz iw:={iw} format:={format} xlabel:={xlabel} ylabel:={ylabel} trim:={trim} oz:={oz}"


# --- wautofill ---
def wautofill(
    irng: str = "<active>",
    orng: Optional[str] = None,
    action: str = "repeat",
    mode: int = 0,
) -> str:
    """
    Auto-fill worksheet cells based on input range values.

    ### irng
        Input range to use as the fill pattern.

    ### orng
        Destination range to fill.

    ### action
        Fill mode: repeat, enumerate, randomize, insert, single_enumerate.

    ### mode
        Method: 0=Drag & Drop, 1=Double Click.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/wautofill
    """
    cmd = f"wautofill irng:={irng} action:={action} mode:={mode}"
    if orng is not None:
        cmd += f" orng:={orng}"
    return cmd


# --- wautosize ---
def wautosize(
    iw: str = "<active>",
    cw1: float = 1.5,
    cw2: float = 25,
    hw: str = "w",
) -> str:
    """
    Auto-size worksheet column widths to fit the maximum string length.

    ### iw
        Worksheet to resize.

    ### cw1
        Minimum column width (in characters).

    ### cw2
        Maximum column width (in characters).

    ### hw
        Resize target: w=width only, h=height only, b=both.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/wautosize
    """
    return f"wautosize iw:={iw} cw1:={cw1} cw2:={cw2} hw:={hw}"


# --- wcellcolor ---
def wcellcolor(
    irng: str = "<active>",
    color: int = 0,
    type: str = "fill",
) -> str:
    """
    Set the fill or font color of worksheet cells.

    ### irng
        Cell range to color.

    ### color
        Color index (0=Black, follow Origin's color list).

    ### type
        Color type: fill (background) or font (text color).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/wcellcolor
    """
    return f"wcellcolor irng:={irng} color:={color} type:={type}"


# --- wcellformat ---
def wcellformat() -> str:
    """
    Open the Format Cells dialog for the selected worksheet cells.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/wcellformat
    """
    return "wcellformat"


# --- wcellmask ---
def wcellmask(irng: str = "<active>", mask: str = "mask") -> str:
    """
    Mask or unmask a range of worksheet cells.

    ### irng
        Range to be masked/unmasked.

    ### mask
        mask=mask the range, mask=unmask to unmask the range.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/wcellmask
    """
    return f"wcellmask irng:={irng} mask:={mask}"


# --- wcellsel ---
def wcellsel(
    rng: str = "<active>",
    condition: str = "none",
    val: float = 0.05,
    tol: float = 1e-6,
) -> str:
    """
    Select worksheet cells based on a condition.

    ### rng
        Range to search for cells matching the condition.

    ### condition
        Selection criteria: eq, lt, le, gt, ge, ne, or none (clear selection).

    ### val
        Cutoff value for the condition.

    ### tol
        Tolerance used in comparing cell values to the cutoff.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/wcellsel
    """
    return f"wcellsel rng:={rng} condition:={condition} val:={val} tol:={tol}"


# --- wclear ---
def wclear(
    w: str = "<active>",
    reduce: int = 1,
    msg: int = 1,
    r1: int = 1,
    hidden: int = 1,
) -> str:
    """
    Clear worksheet data.

    ### w
        Worksheet to clear.

    ### reduce
        Reduce dataset size to 0 and minimal rows (1=yes, 0=no).

    ### msg
        Show warning message (1=yes, 0=no).

    ### r1
        Start row number for clearing (data from r1 to end is cleared).

    ### hidden
        Reset hidden rows (1=reset to reveal all, 0=keep existing).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/wclear
    """
    return f"wclear w:={w} reduce:={reduce} msg:={msg} r1:={r1} hidden:={hidden}"


# --- wcolwidth ---
def wcolwidth(irng: str = "<active>", width: float = 7) -> str:
    """
    Set the width of worksheet columns.

    ### irng
        Column(s) to resize.

    ### width
        Column width (ratio to default font width; 7 = 7 characters wide).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/wcolwidth
    """
    return f"wcolwidth irng:={irng} width:={width}"


# --- wcopy ---
def wcopy(
    iw: str = "<active>",
    ow: str = "<new>",
    copydata: int = 1,
    script: int = 1,
    hidden: int = 0,
    activate: int = 1,
) -> str:
    """
    Create a copy of a worksheet.

    ### iw
        Source worksheet to copy.

    ### ow
        Destination worksheet.

    ### copydata
        Copy the data (1=yes, 0=no).

    ### script
        Copy the worksheet script (1=yes, 0=no).

    ### hidden
        Create a hidden copy (1=yes, 0=no).

    ### activate
        Activate the output sheet if not new (1=yes, 0=no).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/wcopy
    """
    return f"wcopy iw:={iw} ow:={ow} copydata:={copydata} script:={script} hidden:={hidden} activate:={activate}"


# --- wdeldup ---
def wdeldup(
    irng: str = "<active>",
    keep1st: int = 1,
    sensitive: int = 0,
    tol: float = 1e-8,
    ow: str = "<input>",
) -> str:
    """
    Delete or merge duplicate rows based on a reference column.

    ### irng
        Column(s) to check for duplicates.

    ### keep1st
        How to handle duplicates: 0=removeAll, 1=keep1st, 2=keepLast,
        3=average, 4=min, 5=max, 6=sum, 7=sd.

    ### sensitive
        Case-sensitive string comparison (1=yes, 0=no).

    ### tol
        Tolerance for treating close values as duplicates.

    ### ow
        Output worksheet.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/wdeldup
    """
    return f"wdeldup irng:={irng} keep1st:={keep1st} sensitive:={sensitive} tol:={tol} ow:={ow}"


# --- wdelrows ---
def wdelrows(
    irng: str = "<active>",
    method: int = 0,
    del_: int = 1,
    skip: int = 1,
    start: int = 1,
) -> str:
    """
    Delete rows from a worksheet.

    ### irng
        Source range.

    ### method
        Deletion method: 0=delete N rows then skip M rows,
        1=remove rows with missing values, 2=remove masked rows.

    ### del_
        Number of rows to delete.

    ### skip
        Number of rows to skip.

    ### start
        Starting row for deletion.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/wdelrows
    """
    return f"wdelrows irng:={irng} method:={method} del:={del_} skip:={skip} start:={start}"


# --- wexpand2m ---
def wexpand2m(
    iw: str = "<active>",
    expand: int = 1,
    order: str = "row",
    om: str = "<new>",
) -> str:
    """
    Expand worksheet rows or columns into a matrix.

    ### iw
        Worksheet to convert.

    ### expand
        Number of rows/columns comprising one matrix row/column.

    ### order
        Expanding direction: row (expand in row) or col (expand in column).

    ### om
        Output matrix.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/wexpand2m
    """
    return f"wexpand2m iw:={iw} expand:={expand} order:={order} om:={om}"


# --- wkeepdup ---
def wkeepdup(
    irng: str = "<active>",
    num: int = 2,
    sensitive: int = 0,
    ow: str = "<input>",
) -> str:
    """
    Keep only duplicate rows (meeting minimum duplication threshold).

    ### irng
        Column to search for duplicates.

    ### num
        Minimum number of duplicate rows to keep.

    ### sensitive
        Case-sensitive duplication check (1=yes, 0=no).

    ### ow
        Output worksheet for duplicate rows.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/wkeepdup
    """
    return f"wkeepdup irng:={irng} num:={num} sensitive:={sensitive} ow:={ow}"


# --- wmergexy ---
def wmergexy(
    iy: Optional[str] = None,
    oy: Optional[str] = None,
    label: Optional[str] = None,
) -> str:
    """
    Merge XY datasets by matching X values (insert blank rows if X doesn't match).

    ### iy
        XY range to copy.

    ### oy
        Destination XY range.

    ### label
        Column labels to copy (e.g., "L"=Long Name, "LC"=Long Name+Comments).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/wmergexy
    """
    cmd = "wmergexy"
    if iy is not None:
        cmd += f" iy:={iy}"
    if oy is not None:
        cmd += f" oy:={oy}"
    if label is not None:
        cmd += f" label:={label}"
    return cmd


# --- wmove_sheet ---
def wmove_sheet(
    source: str,
    target: str = "<new>",
    index: int = 0,
) -> str:
    """
    Move a specified worksheet to another workbook.

    ### source
        Source worksheet to be moved.

    ### target
        Destination workbook.

    ### index
        Index of the moved worksheet in the destination workbook.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/wmove_sheet
    """
    return f"wmove_sheet source:={source} target:={target} index:={index}"


# --- wmvsn ---
def wmvsn(
    w: str = "<active>",
    prefix: Optional[str] = None,
    label: int = 0,
) -> str:
    """
    Reset short names of all columns in a worksheet.

    ### w
        Worksheet to reset.

    ### prefix
        Prefix for the new short names.

    ### label
        Column label row to move original short names to
        (0=None, 1=Long Name, 2=Units, 3=Comments, etc.).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/wmvsn
    """
    cmd = f"wmvsn w:={w} label:={label}"
    if prefix is not None:
        cmd += f" prefix:=\"{prefix}\""
    return cmd


# --- wproperties ---
def wproperties(
    iw: str = "<active>",
    execute: str = "get",
) -> str:
    """
    Get or set worksheet properties in tree form.

    ### iw
        Input worksheet.

    ### execute
        Mode: get (Get Properties) or set (Set Properties).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/wproperties
    """
    return f"wproperties iw:={iw} execute:={execute}"


# --- wrcopy ---
def wrcopy(
    iw: str = "<active>",
    ow: str = "<new>",
    c1: int = 1,
    c2: int = 0,
    r1: int = 1,
    r2: int = 0,
    dc1: int = 1,
    dr1: int = 1,
    transpose: int = 0,
    label: str = "0",
    format: int = 0,
    clear: int = 0,
) -> str:
    """
    Copy a cell range from one worksheet to another.

    ### iw
        Source worksheet.

    ### ow
        Destination worksheet.

    ### c1 / c2
        Source column begin/end.

    ### r1 / r2
        Source row begin/end (0=last row).

    ### dc1 / dr1
        Destination column/row begin.

    ### transpose
        Transpose the copied range (1=yes, 0=no).

    ### label
        Copy column labels: 0=none, 1=all, or specific label chars.

    ### format
        Copy column formatting (1=yes, 0=no).

    ### clear
        Clear destination columns before copying (1=yes, 0=no).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/wrcopy
    """
    return (
        f"wrcopy iw:={iw} ow:={ow}"
        f" c1:={c1} c2:={c2} r1:={r1} r2:={r2}"
        f" dc1:={dc1} dr1:={dr1}"
        f" transpose:={transpose} label:={label} format:={format} clear:={clear}"
    )


# --- wreplace ---
def wreplace(
    rng: str = "<active>",
    type: int = 0,
    find_value: Optional[float] = None,
    find_str: Optional[str] = None,
    cond_value: int = 0,
    replace_value: Optional[float] = None,
    replace_str: Optional[str] = None,
    tolerance: float = 1e-8,
    label: int = 0,
) -> str:
    """
    Search and replace cell values in a worksheet.

    ### rng
        Range to perform the replacing on.

    ### type
        Data type: 0=numeric, 1=string.

    ### find_value
        Numeric value to find.

    ### find_str
        String to find.

    ### cond_value
        Condition operator for numeric: 0=eq, 1=lt, 2=le, 3=gt, 4=ge, 5=ne.

    ### replace_value
        New numeric value for matching cells.

    ### replace_str
        New string for matching cells.

    ### tolerance
        Tolerance for finding numeric values.

    ### label
        Include label rows in search (1=yes, 0=no).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/wreplace
    """
    cmd = f"wreplace rng:={rng} type:={type} cond_value:={cond_value} tolerance:={tolerance} label:={label}"
    if find_value is not None:
        cmd += f" find_value:={find_value}"
    if find_str is not None:
        cmd += f" find_str:=\"{find_str}\""
    if replace_value is not None:
        cmd += f" replace_value:={replace_value}"
    if replace_str is not None:
        cmd += f" replace_str:=\"{replace_str}\""
    return cmd


# --- wrow2label ---
def wrow2label(
    iw: str = "<active>",
    longname: int = 0,
    unit: int = 0,
    comment: int = 0,
    removerow: int = 1,
    showlabel: int = 1,
) -> str:
    """
    Promote worksheet data rows to column label rows.

    ### iw
        Input worksheet.

    ### longname
        Row index to convert to Long Name.

    ### unit
        Row index to convert to Units.

    ### comment
        Row index to convert to Comments.

    ### removerow
        Remove the converted rows (1=yes, 0=no).

    ### showlabel
        Show the label rows after conversion (1=yes, 0=no).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/wrow2label
    """
    return (
        f"wrow2label iw:={iw} longname:={longname} unit:={unit}"
        f" comment:={comment} removerow:={removerow} showlabel:={showlabel}"
    )


# --- wrowheight ---
def wrowheight(irng: str = "<active>", height: float = 1) -> str:
    """
    Set the height of worksheet rows.

    ### irng
        Row(s) to resize.

    ### height
        Row height (ratio to default font height; 1 = 1 character tall).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/wrowheight
    """
    return f"wrowheight irng:={irng} height:={height}"


# --- wsort ---
def wsort(
    w: str = "<active>",
    bycol: int = 1,
    descending: int = 0,
    c1: int = 1,
    c2: int = 0,
    r1: int = 1,
    r2: int = 0,
) -> str:
    """
    Sort a worksheet by one or more columns.

    ### w
        Worksheet to sort.

    ### bycol
        Column index to use as sort criteria.

    ### descending
        Sort order: 0=ascending, 1=descending,
        2=categorical ascending, 3=categorical descending.

    ### c1 / c2
        Start/end column index for the sort range.

    ### r1 / r2
        Start/end row index for the sort range.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/wsort
    """
    return f"wsort w:={w} bycol:={bycol} descending:={descending} c1:={c1} c2:={c2} r1:={r1} r2:={r2}"


# --- wsplit_book ---
def wsplit_book(
    fld: str = "active",
    mode: str = "duplicate",
    keep: int = 1,
    rename: int = 1,
    match: str = "none",
    key: Optional[str] = None,
) -> str:
    """
    Split a workbook into multiple single-sheet workbooks.

    ### fld
        Which workbooks to split: active, folder, recursive, open, or project.

    ### mode
        Split mode: duplicate (copy sheets) or drag (drag sheets out).

    ### keep
        Keep the source workbook (1=yes, 0=no).

    ### rename
        Rename new workbooks using their source sheet names (1=yes, 0=no).

    ### match
        Where to search the key string: none, lname, sname, or comments.

    ### key
        Search pattern for filtering workbooks (supports wildcards).

    Ref: https://www.originlab.com/doc/en/X-Function/ref/wsplit_book
    """
    cmd = f"wsplit_book fld:={fld} mode:={mode} keep:={keep} rename:={rename} match:={match}"
    if key is not None:
        cmd += f" key:=\"{key}\""
    return cmd


# --- wtranspose ---
def wtranspose(
    iw: str = "<active>",
    select: int = 0,
    exchange: int = 1,
    ow: str = "<input>",
) -> str:
    """
    Transpose the active worksheet.

    ### iw
        Worksheet to transpose.

    ### select
        Transpose only the selected region (1=yes, 0=entire sheet).

    ### exchange
        Swap column label rows with data columns (1=yes, 0=no).

    ### ow
        Destination worksheet.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/wtranspose
    """
    return f"wtranspose iw:={iw} select:={select} exchange:={exchange} ow:={ow}"


# --- wunstackcol ---
def wunstackcol(
    irng1: str = "<active>",
    irng2: Optional[str] = None,
    nonstack: int = 0,
    sort: int = 0,
    ow: str = "<new>",
) -> str:
    """
    Unstack data columns using a group column.

    ### irng1
        Data to be unstacked.

    ### irng2
        Group column(s) for unstacking.

    ### nonstack
        Include non-unstack columns from the original sheet (1=yes, 0=no).

    ### sort
        Sort output columns: 0=by group (alphanumeric), 1=by data order.

    ### ow
        Output worksheet.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/wunstackcol
    """
    cmd = f"wunstackcol irng1:={irng1} nonstack:={nonstack} sort:={sort} ow:={ow}"
    if irng2 is not None:
        cmd += f" irng2:={irng2}"
    return cmd


# --- wxt ---
def wxt(
    test: str,
    iw: str = "<active>",
    c1: int = 1,
    c2: int = -1,
    r1: int = 1,
    r2: int = -1,
    sel: int = 0,
    ow: Optional[str] = None,
    pass_: Optional[str] = None,
    num: Optional[str] = None,
) -> str:
    """
    Extract worksheet data using a condition.

    ### test
        Test condition string (e.g., 'col(A) > 5').

    ### iw
        Input worksheet.

    ### c1 / c2
        Begin/end column index of the range to extract.

    ### r1 / r2
        Begin/end row index of the range to extract.

    ### sel
        Mark passed cells: 0=None, 1=Select Cells, 2=Mask Cells.

    ### ow
        Output worksheet for extracted data.

    ### pass_
        Output vector variable name for row indices that pass.

    ### num
        Output variable name for the number of rows that pass.

    Ref: https://www.originlab.com/doc/en/X-Function/ref/wxt
    """
    cmd = f"wxt test:=\"{test}\" iw:={iw} c1:={c1} c2:={c2} r1:={r1} r2:={r2} sel:={sel}"
    if ow is not None:
        cmd += f" ow:={ow}"
    if pass_ is not None:
        cmd += f" pass:={pass_}"
    if num is not None:
        cmd += f" num:={num}"
    return cmd