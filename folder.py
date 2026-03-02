"""
Folder class for OriginExt wrappers.

This module contains the Folder wrapper class for organizing pages in Origin projects.
"""
from __future__ import annotations

import OriginExt.OriginExt as oext_types

from typing import Iterator, Optional, TYPE_CHECKING

from .base import APP, OriginObjectWrapper, OriginNameConflictError, OriginPageGenerationError
from .pages import PageBase, WorkbookPage, GraphPage, MatrixPage, NotePage

if TYPE_CHECKING:
    from .layer.enums import XYPlotType

# ================== Folder Class ==================

class Folder:
    """
    Project folder for organizing pages.
    Wrapper class that wraps OriginExt.OriginExt.Folder with additional features.

    Corresponds to: originpro.Folder, OriginExt.OriginExt.Folder
    """

    _folder: oext_types.Folder

    def __init__(self, folder: oext_types.Folder, core: APP):
        """
        Initialize Folder wrapper.

        Args:
            folder: Original OriginExt.Folder instance to wrap
            core: APP instance reference for LabTalk access
        """
        self._folder = folder
        self._core = core

    @property
    def path(self) -> str:
        """Full path of the folder in the project"""
        return self._folder.Path

    @property
    def parent(self) -> Folder:
        """Parent folder"""
        return Folder(self._folder.Parent, self._core)

    @property
    def folders(self) -> list[Folder]:
        """Collection of subfolders"""
        return [Folder(f, self._core) for f in self._folder.Folders]

    def __iter__(self) -> Iterator[PageBase]:
        """Iterate over pages in this folder"""
        return iter(self._folder)

    @property
    def subfolders(self) -> list[Folder]:
        """
        Get list of subfolders in this folder.

        Corresponds to: OriginExt.OriginExt.Folder.GetFolders()

        Returns:
            list[Folder]: List of subfolders
        """
        return [Folder(f, self._core) for f in self._folder.GetFolders()]

    def get_path(self) -> str:
        """
        Get the full path of this folder.

        Corresponds to: OriginExt.OriginExt.Folder.GetPath()

        Returns:
            str: Full path of the folder in the project
        """
        return self._folder.GetPath()

    @property
    def name(self) -> str:
        """
        Get the name of the folder.

        Returns:
            str: Name of the folder
        """
        path = self._folder.Path
        # Extract the folder name from the path (last part after '/')
        return path.split('/')[-1] if path and '/' in path else path

    @property
    def index(self) -> int:
        """
        Get the index of this folder.

        Corresponds to: OriginExt.OriginExt.Folder.GetIndex()

        Returns:
            int: Index of the folder
        """
        return self._folder.GetIndex()

    def get_parent(self) -> Optional[Folder]:
        """
        Get the parent folder.

        Corresponds to: OriginExt.OriginExt.Folder.GetParent()

        Returns:
            Folder: Parent folder, or None if this is the root folder
        """

        parent = self._folder.GetParent()
        if parent is oext_types.ApplicationBase:
            return None
        return Folder(parent, self._core)

    def result_text(self, recursive: bool = False) -> str:
        """
        Get result text from this folder.

        Corresponds to: OriginExt.OriginExt.Folder.GetResultText()

        Args:
            recursive: If True, include results from subfolders

        Returns:
            str: Result text
        """
        return self._folder.GetResultText(recursive)

    @property
    def pages(self) -> list[PageBase]:
        """
        Get list of pages in this folder.

        Corresponds to: OriginExt.OriginExt.Folder.PageBases()

        Returns:
            list[PageBase]: List of pages in this folder
        """
        return list(self._folder.PageBases())

    def has_page(self, name: str) -> bool:
        """
        Check if a page with the specified name exists in this folder.

        Args:
            name: Name of the page to check for (short name or long name)

        Returns:
            bool: True if a page with the name exists, False otherwise
        """
        for page in self.pages:
            if page.Name == name or page.LongName == name:
                return True
        return False

    def find_page(self, name: str) -> Optional[PageBase]:
        """
        Find a page with the specified name in this folder.

        Args:
            name: Name of the page to find (short name or long name)

        Returns:
            PageBase: The found page object, or None if not found
        """
        for page in self.pages:
            if page.Name == name or page.LongName == name:
                return page
        return None

    def create_folder(self, name: str) -> Folder:
        """
        Add a new subfolder to this folder.

        Args:
            name: Name of the new subfolder

        Returns:
            Folder: The newly created subfolder
        """
        new_folder = self._folder.Folders.Add(name)
        return Folder(new_folder, self._core)

    def new_workbook(self, name: str, template: str = '') -> Optional['WorkbookPage']:
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
        # # Check for name conflicts first
        # root_folder = self.get_root_dir()
        # if root_folder.has_page(name):
        #     raise OriginNameConflictError(f"A page with name '{name}' already exists in root folder")
        
        # Check if this is root folder or subfolder
        full_path = self.get_path()
        
        if full_path.count('/') == 2:
            # Root folder - no cd needed
            cmd = f'newbook name:="{name}" template:="{template}"'
            print(f'Root folder - executing: {cmd}')
            self._core.LT_execute(cmd.strip())
        else:
            # Subfolder - need to change directory
            cd_cmd = f'pe_cd "{full_path}"'
            wb_cmd = f'newbook name:="{name}" template:="{template}"'
            combined_cmd = f'{cd_cmd}; {wb_cmd}; cd /'
            print(f'Subfolder - executing: {combined_cmd}')
            self._core.LT_execute(combined_cmd.strip())
        
        # Find the newly created workbook
        app = self._core
        for page in app.GetWorksheetPages():
            if page.Name == name or page.LongName == name:
                return WorkbookPage(page, self._core)
        
        # If not found, raise error instead of returning None
        raise OriginPageGenerationError(f"Failed to create workbook '{name}'. Command executed but workbook not found.")
    def create_graph(self, name: str, template: str) -> GraphPage:
        """
        Create a new graph page in this folder.

        Args:
            name: Name for the graph (required)
            template: XY template string (e.g., "scatter", "line")

        Returns:
            GraphPage: The newly created graph page
        
        Raises:
            OriginNameConflictError: If a page with the same name already exists in this folder
        """
        # Check if a page with the same name already exists
        if self.has_page(name):
            raise OriginNameConflictError(f"A page with name '{name}' already exists in folder '{self.name}'")
        
        # Use LabTalk command to create a new graph page
        cmd = f'newpanel name:="{name}" template:="{template}"'
        self._core.LT_execute(cmd.strip())

        # Search all graph pages for the newly created one
        for page in self._core.GetGraphPages():
            if page.Name == name or page.LongName == name:
                return GraphPage(page, self._core)

        raise OriginPageGenerationError(
            f"Failed to create graph '{name}'. Command executed but graph not found."
        )

    def create_matrix(self, name: str, template: str = '') -> MatrixPage:
        """
        Create a new matrix book page in this folder.

        Args:
            name: Name for the matrix book (required)
            template: Optional template name

        Returns:
            MatrixPage: The newly created matrix page
        
        Raises:
            OriginNameConflictError: If a page with the same name already exists in this folder
        """
        # Check if a page with the same name already exists
        if self.has_page(name):
            raise OriginNameConflictError(f"A page with name '{name}' already exists in folder '{self.get_name()}'")
        
        # Use LabTalk command to create matrix book in this folder
        cmd = f'cd "{self.get_path()}"; newmatrix name:="{name}" template:="{template}"'
        self._folder.Execute(cmd.strip())
        
        # Get the newly created matrix book by finding it by name
        for page in self._folder.PageBases():
            if page.Name == name or page.LongName == name:
                return MatrixPage(page, self._core)
        return None

    def create_notes(self, name: str) -> NotePage:
        """
        Create a new notes page in this folder.

        Args:
            name: Name for the notes window (required)

        Returns:
            NotePage: The newly created notes page
        
        Raises:
            OriginNameConflictError: If a page with the same name already exists in this folder
        """
        # Check if a page with the same name already exists
        if self.has_page(name):
            raise OriginNameConflictError(f"A page with name '{name}' already exists in folder '{self.get_name()}'")
        
        # Use LabTalk command to create notes in this folder
        cmd = f'cd "{self.get_path()}"; win -n n "{name}"'
        self._folder.Execute(cmd)
        
        # Get the newly created notes page by finding it by name
        for page in self._folder.PageBases():
            if page.Name == name or page.LongName == name:
                return NotePage(page, self._core)
        return None

    def get_pages_by_type(self, type_: str = '') -> list:
        """
        Get all pages of specified type in this folder.

        Args:
            type_: 'w' for workbooks, 'g' for graphs, 'm' for matrices,
                   'n' for notes, '' for all pages
        Returns:
            list: List of page objects
        """
     
        if type_ == 'w':
            return [WorkbookPage(p, self._core) for p in self._folder.GetWorksheetPages()]
        elif type_ == 'g':
            return [GraphPage(p, self._core) for p in self._folder.GetGraphPages()]
        elif type_ == 'm':
            return [MatrixPage(p, self._core) for p in self._folder.GetMatrixPages()]
        elif type_ == 'n':
            return [NotePage(p, self._core) for p in self._folder.GetNotesPages()]
        else:
            # Get all pages
            result = [WorkbookPage(p, self._core) for p in self._folder.GetWorksheetPages()]
            result.extend([GraphPage(p, self._core) for p in self._folder.GetGraphPages()])
            result.extend([MatrixPage(p, self._core) for p in self._folder.GetMatrixPages()])
            result.extend([NotePage(p, self._core) for p in self._folder.GetNotesPages()])
            return result

    def find_workbook(self, name: str):
        """
        Find a workbook by name in this folder.

        Args:
            name: Name of the workbook to find
        Returns:
            WorkbookPage object
        """

        for page in self._core.GetWorksheetPages():
            if page.Name == name or page.LongName == name:
                return WorkbookPage(page, self._core)
        return None

    def find_graph(self, name: str):
        """
        Find a graph page by name in this folder.

        Args:
            name: Name of the graph to find
        Returns:
            GraphPage object or None if not found
        """
        for page in self._core.GetGraphPages():
            if page.Name == name or page.LongName == name:
                return GraphPage(page, self._core)
        return None

    def find_matrix(self, name: str):
        """
        Find a matrix book by name in this folder.

        Args:
            name: Name of the matrix to find
        Returns:
            MatrixPage object or None if not found
        """
        for page in self._core.GetMatrixPages():
            if page.Name == name or page.LongName == name:
                return MatrixPage(page, self._core)
        return None

    def __repr__(self) -> str:
        """
        String representation of the Folder.
        
        Returns:
            str: Formatted string showing folder name, subfolders, sheets, and parent
        """
        folder_name = self.Path
        subfolders = [subfolder.Path for subfolder in self.subfolders]
        sheets = [page.Name for page in self.pages]
        parent_folder = self.Parent.Path if self.Parent else "None"
        
        return (f"Folder(name='{folder_name}', "
                f"subfolders={subfolders}, "
                f"sheets={sheets}, "
                f"parent='{parent_folder}')")


