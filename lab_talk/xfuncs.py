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