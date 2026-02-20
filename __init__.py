"""
必須要件: Origin 2019以降がインストールされている（外部pythonからOriginのAPIが叩ける）こと
それ以前のOriginだとエラーはいて止まる。

originproパッケージが微妙に使いにくいので、改良したAPIっぽいものを提供する。
たとえば、シートを選択するときにはシートの有無を先に検査する必要があるが、シートが無かった場合には自動で作成する関数を提供したりする。

Originのインスタンスは一個しか管理できない。
これはoriginproライブラリの仕様に基づく。

外部PythonからOriginを起動したとき、PythonがOiriginをバインドしたままになるので、手動でウィンドウを閉じれなくなることに注意。

"""
from __future__ import annotations

#%%
import OriginExt as oext

import sys
import os

from enum import Enum
from typing import Optional, Union

# ================== Re-export from submodules ==================

from .base import (
    OriginCollection,
    OriginObjectWrapper,
)

from .pages import (
    PageBase,
    Page,
    WorkbookPage,
    GraphPage,
    MatrixPage,
    NotePage,
    FigurePage,
)

from .layers import (
    Layer,
    Datasheet,
    Column,
    Worksheet,
    DataPlot,
    GraphLayer,
    Matrixsheet,
    PlotType,
    XYTemplate,
    ColorMap,
    GroupMode,
    AxisType,
)

from .folder import (
    Folder,
)

# ================== Application API ==================

# originproよりコピー
# originproは読み込むだけでoriginを起動する
class APP:
    'OriginExt.Application() wrapper'
    def __init__(self):
        self._app = None
        self._first = True
    def __getattr__(self, name):
        try:
            return getattr(oext, name)
        except AttributeError:
            pass
        if self._app is None:
            self._app = oext.Application()
            self._app.LT_execute('sec -poc') # wait until OC ready
        return getattr(self._app, name)
    def __bool__(self):
        return self._app is not None
    def Exit(self, releaseonly=False):
        'Exit if Application exists'
        if self._app is not None:
            self._app.Exit(releaseonly)
            self._app = None
    def Attach(self):
        'Attach to exising Origin instance'
        releaseonly = True
        if self._first:
            releaseonly = False
            self._first = False
        self.Exit(releaseonly)
        self._app = oext.ApplicationSI()
    def Detach(self):
        'Detach from Origin instance'
        self.Exit(True)

class OriginPath(Enum):
    USER_FILES_DIR = "u",
    ORIGIN_EXE_DIR = "e",
    PROJECT_DIR = "p",
    ATTACHED_FILE_DIR = "a",
    LEANING_CENTER = "l"


class OriginNotFoundError(BaseException):
    pass

class OriginInstanceGenerationError(BaseException):
    pass

class OriginTooManyInstancesError(BaseException):
    pass

class OriginNameConflictError(BaseException):
    """Exception raised when trying to create an object with a name that already exists."""
    pass


if "ORIGIN_INSTANCE_LIMIT" not in globals():
    ORIGIN_INSTANCE_LIMIT = 5



class OriginInstance:
    """
    ## Abstract
    Originのインスタンスを管理するクラス。
    ここからファイル構造を辿ったり、シートを取得したりする。

    ## Member variables
    __core (op.APP): originproのAPPクラスのインスタンス。

    path (str): このインスタンスが管理しているOriginプロジェクトファイルのパス。
    パスはグローバル変数で管理しており、同じパスのものは複数生成できない。


    ## note
    デストラクタが呼ばれた際、自動的に保存するので、保存したくないときには明示的にclose(False)を呼び出さなければならない。



    余談だが、originproはその設計上、インスタンスを一度だけ、一個だけ呼べるようになっており、そのインスタンスそのものを操作することがサポートされていない。
    オリジナルのoriginproではop.poの実態がモジュールではなくAPPクラスになっている。
    実質的にこれがOriginのインスタンスとなるのだが、このクラスのインスタンス生成がライブラリ呼び出し時の一回だけである。
    インテリジェンスの予測もop.poをモジュールとして認識してしまっているため、わかりにくくなっている。
    """
    __core: APP
    @property
    def get_API(self) -> APP:
        return self.__core

    path: str
    @property
    def get_path(self) -> str:
        return self.path

    # 疑似static変数
    __instance_count: int = 0
    __instance_path_list: set[str] = set()

    def __init__(
            self,
            path_to_origin_file: str,
            create_new_if_not_exist: bool = True
            ):
        """
        Originのインスタンスを生成する。
        Args:
            path_to_origin_file (str): 開きたいOriginプロジェクトファイルのパス(**フルパス、絶対参照**)
            create_new_if_not_exist (bool, optional): ファイルが無かった場合に新規作成するかどうか. Defaults to True.

        Raises:
            FileNotFoundError: path_to_origin_fileで指定したパスにファイルが存在しない場合に発生
            OriginTooManyInstancesError: Originのインスタンス数が多すぎる場合に発生
        Returns:
            このクラスのインスタンス
        """

        # //があるとフォルダだと認識されないため、置換
        self.path = path_to_origin_file.replace("//", "\\")

        dir = os.path.dirname(self.path)
        if not os.path.exists(dir):
            raise OriginNotFoundError(
                "The directory was not found:\n\
                {}".format(dir)
            )

        # 同パスへのインスタンスが既にあればエラー
        if path_to_origin_file in OriginInstance.__instance_path_list:
            raise OriginInstanceGenerationError(
                f"An Origin instance with the path {path_to_origin_file} already active."
                )
        OriginInstance.__instance_path_list.add(path_to_origin_file)


        # インスタンス数が多すぎたらエラー
        if OriginInstance.__instance_count > ORIGIN_INSTANCE_LIMIT:
            raise OriginTooManyInstancesError(
                "Too many Origin instances are being created.\n"+ \
                "Run close() function to free instances\n"+ \
                "To increase the limit of the number of instances,"+ \
                "define ORIGIN_INSTANCE_LIMIT variable before importing this module."
                )

        # インスタンス生成開始
        print("Generating Origin instance...")
        self.__core = APP()

        # すでにあるファイルへのパスならばそれを読み込み
        if os.path.isfile(self.path):
            print("Opening the file...")
            r = self.__core.Load(self.path, False)
            # 書き込み可で開く

            if r == False:
                self.close()
                raise OriginInstanceGenerationError(
                    "Failed to load.\n \
                    Please check the extension, the version of Origin, and so on.\n"+\
                    "The path: {}".format(self.path)
                    )

        # 新規作成がFalseでパスが見つからない場合
        elif not create_new_if_not_exist:
            raise OriginNotFoundError(
                "The file was not found\n \
                The path: {}".format(path_to_origin_file)
            )
        # 一旦新規作成して保存
        else:
            print("Generating a new file...")
            self.__core.Save(self.path)


        OriginInstance.__instance_count += 1

        self.__core.LT_set_var("@VIS", 100)

        print("Origin booted")

        # エラーが投げられたら即座に終了
        def origin_shutdown_exception_hook(exctype, value, traceback):
            '''Ensures Origin gets shut down if an uncaught exception'''
            self.close()
            sys.__excepthook__(exctype, value, traceback)
        if oext:
            sys.excepthook = origin_shutdown_exception_hook

    def __del__(self):
        self.close()

    def close(self, save_flag: bool = True) -> None:
        '''Originのインスタンスを終了する'''
        if save_flag:
            self.__core.Save(self.path)
        self.__core.Exit()

        OriginInstance.__instance_count = OriginInstance.__instance_count - 1

        OriginInstance.__instance_path_list.remove(self.path)

    def get_root_dir(self) -> Folder:
        '''Originのルートディレクトリを取得する'''
        return Folder(self.__core.GetRootFolder())

    def lt_exec_cmnd(self, command: str)->None:
        """
        ## 概要
        OriginにLab talkのコマンドを送って実行する。
        一部のコマンドには例外処理を挟み、バグを減らしているが、網羅しているわけではないので利用は計画的に。

        ## 例外処理
        exit: 代わりにclose()処理を実行する
        """

        match (command):
            case "exit":
                self.close()
            case _:
                self.__core.lt_exec(command)

    def is_valid(self) -> bool:
        return bool(self.__core)

    # ================== Visibility / Display ==================

    def set_show(self, show: bool = True) -> None:
        """
        Set Origin window visibility.

        Corresponds to: originpro.set_show()

        Args:
            show: True to show the window, False to hide it
        """
        self.__core.LT_set_var("@VIS", 100 if show else 0)

    def set_display(self, mode: int) -> None:
        """
        Set display mode.

        Corresponds to: originpro.set_display()

        Args:
            mode: 0=hidden, 1=minimized, 2=normal, 3=maximized
        """
        self.__core.LT_set_var("@VIS", mode)

    # ================== Path Functions ==================

    def get_origin_path(self, path_type: str | OriginPath) -> str:
        """
        Get Origin system paths.

        Corresponds to: originpro.path()

        Args:
            path_type: One of the following:
                'u' or OriginPath.USER_FILES_DIR: User Files folder
                'e' or OriginPath.ORIGIN_EXE_DIR: Origin exe folder
                'p' or OriginPath.PROJECT_DIR: Current project folder
                'a' or OriginPath.ATTACHED_FILE_DIR: Attached files folder
                'l' or OriginPath.LEANING_CENTER: Learning Center folder
        Returns:
            str: The requested path
        """
        if isinstance(path_type, OriginPath):
            path_type = path_type.value[0]

        path_map = {
            'u': 'system.path.program$',
            'e': 'system.path.program$',
            'p': 'system.path.project$',
            'a': 'system.path.appdata$',
            'l': 'system.path.learningcenter$',
        }

        if path_type == 'u':
            return self.__core.LT_get_str('%Y')
        elif path_type == 'e':
            return self.__core.LT_get_str('%X')
        elif path_type == 'p':
            return os.path.dirname(self.path)
        else:
            var_name = path_map.get(path_type, 'system.path.program$')
            return self.__core.LT_get_str(var_name)

    # ================== Project Operations ==================

    def save(self, path: str | None = None) -> bool:
        """
        Save the current project.

        Corresponds to: originpro.save()

        Args:
            path: Optional path to save to. If None, saves to current path.
        Returns:
            bool: True if save succeeded
        """
        if path is not None:
            self.path = path.replace("//", "\\")
        return self.__core.Save(self.path)

    # def new_project(self) -> None:
    #     """
    #     Create a new empty project, closing the current one without saving.

    #     Corresponds to: originpro.new()
    #     """
    #     self.__core.LT_execute("doc -n")

    # ================== Folder Operations ==================

    def get_folder(self, path: str | None = None) -> Folder:
        """
        Get a folder object.

        Corresponds to: originpro.root_folder() (when path is None),
                        originpro.get_folder()

        Args:
            path: Folder path in Origin project. If None, returns root folder.
        Returns:
            Folder object
        """
        if path is None:
            return Folder(self.__core.GetRootFolder())
        return Folder(self.__core.GetFolder(path))

    def root_folder(self) -> Folder:
        """
        Get the root folder of the project.

        Corresponds to: originpro.root_folder()

        Returns:
            Root folder object
        """
        return Folder(self.__core.GetRootFolder())

    def make_folder(self, name: str, parent_path: str | None = None) -> Folder:
        """
        Create a new folder in the project.

        Corresponds to: originpro.Folder.create_folder()

        Args:
            name: Name of the new folder
            parent_path: Path to parent folder. If None, creates in root.
        Returns:
            The created folder object
        """
        if parent_path is None:
            parent = self.__core.GetRootFolder()
        else:
            parent = self.__core.GetFolder(parent_path)
        return Folder(parent.AddFolder(name))

    # ================== Workbook / Worksheet Operations ==================

    def get_workbook_pages(self) -> list[WorkbookPage]:
        """
        Get all workbook pages in the project.

        Corresponds to: originpro.pages('w')

        Returns:
            list: List of workbook page objects
        """
        return [WorkbookPage(p) for p in self.__core.GetWorksheetPages()]

    def get_worksheet_pages(self) -> list[WorkbookPage]:
        """
        Get all worksheet pages in the project (alias for get_workbook_pages).

        Corresponds to: originpro.pages('w')

        Returns:
            list: List of workbook page objects
        """
        return [WorkbookPage(p) for p in self.__core.GetWorksheetPages()]

    # def find_sheet(self, name: str) -> Optional[Worksheet]:
    #     """
    #     Find a worksheet by name (short name or long name).

    #     Corresponds to: originpro.find_sheet('w', name)

    #     Args:
    #         name: Name of the worksheet to find
    #     Returns:
    #         Worksheet object or None if not found
    #     """
    #     for page in self.__core.GetWorksheetPages():
    #         if page.Name == name or page.LongName == name:
    #             # Get the first sheet in the workbook
    #             if page.Layers.Count > 0:
    #                 return Worksheet(page.Layers(0))
    #     return None

    # NOTE: To find worksheets following proper hierarchy, use:
    # workbook = origin.find_book('WorkbookName')
    # if workbook:
    #     worksheet = workbook[0]  # Get first worksheet

    def find_book(self, name: str) -> Optional[WorkbookPage]:
        """
        Find a workbook by name (short name or long name).

        Corresponds to: originpro.find_book('w', name)

        Args:
            name: Name of the workbook to find
        Returns:
            Workbook page object or None if not found
        """
        for page in self.__core.GetWorksheetPages():
            if page.Name == name or page.LongName == name:
                return WorkbookPage(page)
        return None

    def new_workbook(self, name: str, template: str = '') -> Optional[WorkbookPage]:
        """
        Create a new workbook page in the root folder.
        Uses direct LabTalk execution for now.

        Args:
            name: Name for the workbook (required)
            template: Optional template name
        Returns:
            The created workbook page object, or None if creation failed
        
        Raises:
            OriginNameConflictError: If a page with the same name already exists
        """
        # Check for name conflicts first
        root_folder = self.get_root_dir()
        if root_folder.has_page(name):
            raise OriginNameConflictError(f"A page with name '{name}' already exists in root folder")
        
        # Use direct LabTalk execution
        cmd = f'newbook name:="{name}" template:="{template}"'
        self.__core.LT_execute(cmd.strip())
        
        # Find the newly created workbook
        for page in self.__core.GetWorksheetPages():
            if page.Name == name or page.LongName == name:
                return WorkbookPage(page)
        return None

    def new_matrixbook(self, name: str, template: str = '') -> Optional[MatrixPage]:
        """
        Create a new matrix book page in the root folder.
        Delegates to the root folder to maintain proper hierarchy.

        Args:
            name: Name for the matrix book (required)
            template: Optional template name
        Returns:
            The created matrix book page object, or None if creation failed
        
        Raises:
            OriginNameConflictError: If a page with the same name already exists
        """
        root_folder = self.get_root_dir()
        return root_folder.create_matrix(name, template)

    # def new_sheet(self, type_: str = 'w', long_name: str = '', template: str = '') -> Optional[Union[Worksheet, Matrixsheet]]:
    #     """
    #     Create a new workbook with a single sheet or matrix.

    #     Corresponds to: originpro.new_sheet()

    #     Args:
    #         type_: 'w' for worksheet, 'm' for matrix
    #         long_name: Optional long name for the sheet
    #         template: Optional template name
    #     Returns:
    #         The created sheet object
    #     """
    #     book = self.new_book(type_, long_name, template)
    #     if book and book.Layers.Count > 0:
    #         layer = book.Layers(0)
    #         if type_ == 'w':
    #             return Worksheet(layer)
    #         else:
    #             return Matrixsheet(layer)
    #     return None

    # NOTE: To create worksheets following proper hierarchy, use:
    # workbook = origin.new_book('w', 'WorkbookName')
    # worksheet = workbook[0]  # Get first worksheet

    # ================== Graph Operations ==================

    def get_graph_pages(self) -> list[GraphPage]:
        """
        Get all graph pages in the project.

        Corresponds to: originpro.pages('g')

        Returns:
            list: List of graph page objects
        """
        return [GraphPage(p) for p in self.__core.GetGraphPages()]

    def find_graph(self, name: str) -> Optional[GraphPage]:
        """
        Find a graph page by name (short name or long name).

        Corresponds to: originpro.find_graph(name)

        Args:
            name: Name of the graph to find
        Returns:
            Graph page object or None if not found
        """
        for page in self.__core.GetGraphPages():
            if page.Name == name or page.LongName == name:
                return GraphPage(page)
        return None

    def new_graph(self, name: str, template = None) -> Optional[GraphPage]:
        """
        Create a new graph page in the root folder.
        Uses direct LabTalk execution for reliability.

        Args:
            name: Name for the graph (required)
            template: XY template enum (default: XYTemplate.LINE)
        Returns:
            The created graph page object
        
        Raises:
            OriginNameConflictError: If a page with the same name already exists
        """
        # Import XYTemplate at runtime to avoid circular dependency
        if template is None:
            from .layers import XYTemplate
            template = XYTemplate.LINE
        
        # Check for name conflicts first
        root_folder = self.get_root_dir()
        if root_folder.has_page(name):
            raise OriginNameConflictError(f"A page with name '{name}' already exists in root folder")
        
        # Use direct LabTalk execution
        cmd = f'newpanel name:="{name}" template:="{template.value}"'
        self.__core.LT_execute(cmd.strip())
        
        # Find the newly created graph
        for page in self.__core.GetGraphPages():
            if page.Name == name or page.LongName == name:
                return GraphPage(page)
        
        return None

    # ================== Matrix Operations ==================

    def get_matrix_pages(self) -> list[MatrixPage]:
        """
        Get all matrix pages in the project.

        Corresponds to: originpro.pages('m')

        Returns:
            list: List of matrix page objects
        """
        return [MatrixPage(p) for p in self.__core.GetMatrixPages()]

    def find_matrix(self, name: str) -> Optional[MatrixPage]:
        """
        Find a matrix book by name (short name or long name).

        Corresponds to: originpro.find_sheet('m', name) / originpro.find_book('m', name)

        Args:
            name: Name of the matrix to find
        Returns:
            Matrix page object or None if not found
        """
        for page in self.__core.GetMatrixPages():
            if page.Name == name or page.LongName == name:
                return MatrixPage(page)
        return None

    # ================== Notes ==================

    def get_notes_pages(self) -> list[NotePage]:
        """
        Get all notes pages in the project.

        Corresponds to: originpro.pages('n')

        Returns:
            list: List of notes page objects
        """
        return [NotePage(p) for p in self.__core.GetNotesPages()]

    def new_notes(self, name: str) -> Optional[NotePage]:
        """
        Create a new notes page in the root folder.
        Delegates to the root folder to maintain proper hierarchy.

        Args:
            name: Name for the notes window (required)
        Returns:
            The created notes page object
        
        Raises:
            OriginNameConflictError: If a page with the same name already exists
        """
        root_folder = self.get_root_dir()
        return root_folder.create_notes(name)

    # ================== LabTalk Variables ==================

    def lt_get_var(self, name: str) -> float:
        """
        Get a LabTalk variable value.

        Corresponds to: originpro.lt_int() / originpro.lt_float()

        Args:
            name: Variable name
        Returns:
            The variable value (numeric)
        """
        return self.__core.LT_get_var(name)

    def lt_set_var(self, name: str, value: float) -> None:
        """
        Set a LabTalk variable value.

        Corresponds to: originpro.lt_set_var()

        Args:
            name: Variable name
            value: Numeric value to set
        """
        self.__core.LT_set_var(name, value)

    def lt_get_str(self, name: str) -> str:
        """
        Get a LabTalk string variable value.

        Corresponds to: originpro.lt_string()

        Args:
            name: Variable name (use $ suffix for string vars)
        Returns:
            The string value
        """
        return self.__core.LT_get_str(name)

    def lt_set_str(self, name: str, value: str) -> None:
        """
        Set a LabTalk string variable.

        Corresponds to: originpro.lt_set_str()

        Args:
            name: Variable name (use $ suffix for string vars)
            value: String value to set
        """
        self.__core.LT_set_str(name, value)

    # ================== Page Iteration Utilities ==================

    def pages(self, type_: str = '') -> list[PageBase]:
        """
        Get all pages of specified type.

        Corresponds to: originpro.pages()

        Args:
            type_: 'w' for workbooks, 'g' for graphs, 'm' for matrices,
                   'n' for notes, '' for all pages
        Returns:
            list: List of page objects
        """
        if type_ == 'w':
            return [WorkbookPage(p) for p in self.__core.GetWorksheetPages()]
        elif type_ == 'g':
            return [GraphPage(p) for p in self.__core.GetGraphPages()]
        elif type_ == 'm':
            return [MatrixPage(p) for p in self.__core.GetMatrixPages()]
        elif type_ == 'n':
            return [NotePage(p) for p in self.__core.GetNotesPages()]
        else:
            # Get workbook pages
            result = [WorkbookPage(p) for p in self.__core.GetWorksheetPages()]
            result.extend([GraphPage(p) for p in self.__core.GetGraphPages()])
            result.extend([MatrixPage(p) for p in self.__core.GetMatrixPages()])
            result.extend([NotePage(p) for p in self.__core.GetNotesPages()])
            return result

    # ================== Wait/Flush Operations ==================

    def wait(self, type_: str = '') -> None:
        """
        Wait for Origin operations to complete.

        Corresponds to: originpro.wait()

        Args:
            type_: 'r' for recalculate, '' for general wait
        """
        if type_ == 'r':
            self.__core.LT_execute('run -p au')
        else:
            self.__core.LT_execute('sec -poc')

    def flush(self) -> None:
        """
        Flush pending operations.

        Corresponds to: originpro.doc_flush()
        """
        self.__core.LT_execute('doc -uw')


# ================== Public API ==================

__all__ = [
    # Base classes
    'OriginCollection',
    'OriginObjectWrapper',
    # Page classes
    'Folder',
    'PageBase',
    'Page',
    'WorkbookPage',
    'GraphPage',
    'MatrixPage',
    'NotePage',
    'FigurePage',
    # Layer classes
    'Layer',
    'Datasheet',
    'Column',
    'Worksheet',
    'DataPlot',
    'GraphLayer',
    'Matrixsheet',
    # Enum types
    'PlotType',
    'XYTemplate',
    'ColorMap',
    'GroupMode',
    'AxisType',
    # Application API
    'APP',
    'OriginPath',
    'OriginNotFoundError',
    'OriginInstanceGenerationError',
    'OriginTooManyInstancesError',
    'OriginNameConflictError',
    'OriginInstance',
    'ORIGIN_INSTANCE_LIMIT',
]
