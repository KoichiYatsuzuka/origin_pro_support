"""
Folder class for OriginExt wrappers.

This module contains the Folder wrapper class for organizing pages in Origin projects.
"""
from __future__ import annotations

import OriginExt.OriginExt as oext_types

from typing import Iterator, Optional

from .base import OriginObjectWrapper

from .pages import PageBase

# ================== Folder Class ==================

class Folder:
    """
    Project folder for organizing pages.
    Wrapper class that wraps OriginExt.OriginExt.Folder with additional features.

    Corresponds to: originpro.Folder, OriginExt.OriginExt.Folder
    """

    _folder: oext_types.Folder

    def __init__(self, folder: oext_types.Folder):
        """
        Initialize Folder wrapper.

        Args:
            folder: Original OriginExt.Folder instance to wrap
        """
        self._folder = folder

    @property
    def Path(self) -> str:
        """Full path of the folder in the project"""
        return self._folder.Path

    @property
    def Parent(self) -> Folder:
        """Parent folder"""
        return Folder(self._folder.Parent)

    @property
    def Folders(self):
        """Collection of subfolders"""
        return self._folder.Folders

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
        return [Folder(f) for f in self._folder.GetFolders()]

    def get_path(self) -> str:
        """
        Get the full path of this folder.

        Corresponds to: OriginExt.OriginExt.Folder.GetPath()

        Returns:
            str: Full path of the folder in the project
        """
        return self._folder.GetPath()

    def get_name(self) -> str:
        """
        Get the name of this folder (last component of the path).

        Returns:
            str: Name of the folder
        """
        path = self.Path
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
        return Folder(parent)

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

    def create_folder(self, name: str) -> Folder:
        """
        Add a new subfolder to this folder.

        Args:
            name: Name of the new subfolder

        Returns:
            Folder: The newly created subfolder
        """
        new_folder = self._folder.Folders.Add(name)
        return Folder(new_folder)

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


