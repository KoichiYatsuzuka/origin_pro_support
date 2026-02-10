"""
Page classes for OriginExt wrappers.

This module contains wrapper classes for Origin pages including:
- Folder
- PageBase (base class)
- Page
- WorksheetPage
- GraphPage
- MatrixPage
- NotePage
"""
from __future__ import annotations

import OriginExt.OriginExt as oext_types

from typing import Iterator, TypeVar, TYPE_CHECKING

from .base import OriginObjectWrapper

if TYPE_CHECKING:
    from .layers import Layer, Worksheet, GraphLayer, Matrixsheet


# ================== Type Variables ==================

TPageBase = TypeVar('TPageBase', bound=oext_types.PageBase)
TPage = TypeVar('TPage', bound=oext_types.Page)


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

    def get_sub_folders(self) -> list[Folder]:
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

    def get_index(self) -> int:
        """
        Get the index of this folder.

        Corresponds to: OriginExt.OriginExt.Folder.GetIndex()

        Returns:
            int: Index of the folder
        """
        return self._folder.GetIndex()

    def get_parent(self) -> Folder:
        """
        Get the parent folder.

        Corresponds to: OriginExt.OriginExt.Folder.GetParent()

        Returns:
            Folder: Parent folder
        """
        return Folder(self._folder.GetParent())

    def get_result_text(self, recursive: bool = False) -> str:
        """
        Get result text from this folder.

        Corresponds to: OriginExt.OriginExt.Folder.GetResultText()

        Args:
            recursive: If True, include results from subfolders

        Returns:
            str: Result text
        """
        return self._folder.GetResultText(recursive)

    def get_pages(self) -> list[PageBase]:
        """
        Get list of pages in this folder.

        Corresponds to: OriginExt.OriginExt.Folder.PageBases()

        Returns:
            list[PageBase]: List of pages in this folder
        """
        return list(self._folder.PageBases())

    def add_folder(self, name: str) -> Folder:
        """
        Add a new subfolder to this folder.

        Args:
            name: Name of the new subfolder

        Returns:
            Folder: The newly created subfolder
        """
        new_folder = self._folder.Folders.Add(name)
        return Folder(new_folder)


# ================== Page Classes ==================

class PageBase(OriginObjectWrapper[TPageBase]):
    """
    Base class for all Origin page types.
    Wrapper class that wraps OriginExt.OriginExt.PageBase.
    Inherits common methods from OriginObjectWrapper.

    Corresponds to: originpro.PageBase, OriginExt.OriginExt.PageBase
    """

    def __init__(self, page: TPageBase):
        """
        Initialize PageBase wrapper.

        Args:
            page: Original OriginExt.PageBase instance to wrap
        """
        super().__init__(page)

    @property
    def Type(self) -> int:
        """Page type identifier"""
        return self._obj.Type

    @property
    def Parent(self) -> Folder:
        """Parent folder of this page"""
        return Folder(self._obj.Parent)

    def get_type(self) -> int:
        """
        Get the page type.

        Corresponds to: OriginExt.OriginExt.PageBase.GetType()

        Returns:
            int: Page type identifier
        """
        return self._obj.GetType()

    def get_parent(self) -> Folder:
        """
        Get the parent folder.

        Corresponds to: OriginExt.OriginExt.PageBase.GetParent()

        Returns:
            Folder: Parent folder
        """
        return Folder(self._obj.GetParent())


class Page(PageBase[TPage]):
    """
    Page with layers (base for WorksheetPage, GraphPage, MatrixPage).
    Wrapper class that wraps OriginExt.OriginExt.Page.

    Corresponds to: OriginExt.OriginExt.Page
    """

    def __init__(self, page: TPage):
        """
        Initialize Page wrapper.

        Args:
            page: Original OriginExt.Page instance to wrap
        """
        super().__init__(page)

    @property
    def Layers(self):
        """Collection of layers in this page"""
        return self._obj.Layers

    def __iter__(self) -> Iterator[Layer]:
        """Iterate over layers"""
        return iter(self._obj)

    def __getitem__(self, index: int) -> Layer:
        """Get layer by index"""
        from .layers import Layer
        return Layer(self._obj[index])

    def __len__(self) -> int:
        """Get number of layers"""
        return len(self._obj)

    def add_layer(self, name: str = '') -> Layer:
        """
        Add a new layer to this page.

        Corresponds to: OriginExt.OriginExt.Page.AddLayer()

        Args:
            name: Optional name for the new layer

        Returns:
            Layer: The newly created layer
        """
        from .layers import Layer
        return Layer(self._obj.AddLayer(name))

    def get_layers(self) -> list[Layer]:
        """
        Get list of layers in this page.

        Corresponds to: OriginExt.OriginExt.Page.GetLayers()

        Returns:
            list[Layer]: List of layers
        """
        from .layers import Layer
        return [Layer(l) for l in self._obj.GetLayers()]

    def get_layer(self, index: int) -> Layer:
        """
        Get layer by index.

        Corresponds to: OriginExt.OriginExt.Page.GetLayer()

        Args:
            index: Layer index

        Returns:
            Layer: The layer at the specified index
        """
        from .layers import Layer
        return Layer(self._obj.GetLayer(index))

    def preview(self, fname: str) -> bool:
        """
        Generate a preview image.

        Corresponds to: OriginExt.OriginExt.Page.Preview()

        Args:
            fname: File name for the preview image

        Returns:
            bool: True if successful
        """
        return self._obj.Preview(fname)


class WorksheetPage(Page[oext_types.WorksheetPage]):
    """
    Workbook page containing worksheet layers.
    Wrapper class that wraps OriginExt.OriginExt.WorksheetPage.

    Corresponds to: originpro.WBook, OriginExt.OriginExt.WorksheetPage
    """

    def __init__(self, page: oext_types.WorksheetPage):
        """
        Initialize WorksheetPage wrapper.

        Args:
            page: Original OriginExt.WorksheetPage instance to wrap
        """
        super().__init__(page)

    def __iter__(self) -> Iterator[Worksheet]:
        """Iterate over worksheets"""
        from .layers import Worksheet
        for layer in self._obj:
            yield Worksheet(layer)

    def __getitem__(self, index: int) -> Worksheet:
        """Get worksheet by index"""
        from .layers import Worksheet
        return Worksheet(self._obj[index])

    def get_layers(self) -> list[Worksheet]:
        """
        Get list of worksheets in this workbook.

        Corresponds to: OriginExt.OriginExt.WorksheetPage.GetLayers()

        Returns:
            list[Worksheet]: List of worksheets
        """
        from .layers import Worksheet
        return [Worksheet(l) for l in self._obj.GetLayers()]

    def get_layer(self, index: int) -> Worksheet:
        """
        Get worksheet by index.

        Corresponds to: OriginExt.OriginExt.WorksheetPage.GetLayer()

        Args:
            index: Worksheet index

        Returns:
            Worksheet: The worksheet at the specified index
        """
        from .layers import Worksheet
        return Worksheet(self._obj.GetLayer(index))


class GraphPage(Page[oext_types.GraphPage]):
    """
    Graph page containing graph layers.
    Wrapper class that wraps OriginExt.OriginExt.GraphPage.

    Corresponds to: originpro.GPage, OriginExt.OriginExt.GraphPage
    """

    def __init__(self, page: oext_types.GraphPage):
        """
        Initialize GraphPage wrapper.

        Args:
            page: Original OriginExt.GraphPage instance to wrap
        """
        super().__init__(page)

    @property
    def BaseColor(self) -> int:
        """Base color of the graph page"""
        return self._obj.BaseColor

    @property
    def GradColor(self) -> int:
        """Gradient color of the graph page"""
        return self._obj.GradColor

    @property
    def Width(self) -> float:
        """Width of the graph page"""
        return self._obj.Width

    @property
    def Height(self) -> float:
        """Height of the graph page"""
        return self._obj.Height

    @property
    def Units(self) -> int:
        """Units for dimensions"""
        return self._obj.Units

    def __iter__(self) -> Iterator[GraphLayer]:
        """Iterate over graph layers"""
        from .layers import GraphLayer
        for layer in self._obj:
            yield GraphLayer(layer)

    def __getitem__(self, index: int) -> GraphLayer:
        """Get graph layer by index"""
        from .layers import GraphLayer
        return GraphLayer(self._obj[index])

    def get_layers(self) -> list[GraphLayer]:
        """
        Get list of graph layers in this page.

        Corresponds to: OriginExt.OriginExt.GraphPage.GetLayers()

        Returns:
            list[GraphLayer]: List of graph layers
        """
        from .layers import GraphLayer
        return [GraphLayer(l) for l in self._obj.GetLayers()]

    def get_layer(self, index: int) -> GraphLayer:
        """
        Get graph layer by index.

        Corresponds to: OriginExt.OriginExt.GraphPage.GetLayer()

        Args:
            index: Layer index

        Returns:
            GraphLayer: The graph layer at the specified index
        """
        from .layers import GraphLayer
        return GraphLayer(self._obj.GetLayer(index))

    def get_base_color(self) -> int:
        """
        Get base color.

        Corresponds to: OriginExt.OriginExt.GraphPage.GetBaseColor()

        Returns:
            int: Base color
        """
        return self._obj.GetBaseColor()

    def set_base_color(self, color: int) -> None:
        """
        Set base color.

        Corresponds to: OriginExt.OriginExt.GraphPage.SetBaseColor()

        Args:
            color: Base color
        """
        self._obj.SetBaseColor(color)

    def get_width(self) -> float:
        """
        Get page width.

        Corresponds to: OriginExt.OriginExt.GraphPage.GetWidth()

        Returns:
            float: Page width
        """
        return self._obj.GetWidth()

    def set_width(self, width: float) -> None:
        """
        Set page width.

        Corresponds to: OriginExt.OriginExt.GraphPage.SetWidth()

        Args:
            width: Page width
        """
        self._obj.SetWidth(width)

    def get_height(self) -> float:
        """
        Get page height.

        Corresponds to: OriginExt.OriginExt.GraphPage.GetHeight()

        Returns:
            float: Page height
        """
        return self._obj.GetHeight()

    def set_height(self, height: float) -> None:
        """
        Set page height.

        Corresponds to: OriginExt.OriginExt.GraphPage.SetHeight()

        Args:
            height: Page height
        """
        self._obj.SetHeight(height)


class MatrixPage(Page[oext_types.MatrixPage]):
    """
    Matrix book page containing matrix sheets.
    Wrapper class that wraps OriginExt.OriginExt.MatrixPage.

    Corresponds to: originpro.MBook, OriginExt.OriginExt.MatrixPage
    """

    def __init__(self, page: oext_types.MatrixPage):
        """
        Initialize MatrixPage wrapper.

        Args:
            page: Original OriginExt.MatrixPage instance to wrap
        """
        super().__init__(page)

    def __iter__(self) -> Iterator[Matrixsheet]:
        """Iterate over matrix sheets"""
        from .layers import Matrixsheet
        for layer in self._obj:
            yield Matrixsheet(layer)

    def __getitem__(self, index: int) -> Matrixsheet:
        """Get matrix sheet by index"""
        from .layers import Matrixsheet
        return Matrixsheet(self._obj[index])

    def get_layers(self) -> list[Matrixsheet]:
        """
        Get list of matrix sheets in this page.

        Corresponds to: OriginExt.OriginExt.MatrixPage.GetLayers()

        Returns:
            list[Matrixsheet]: List of matrix sheets
        """
        from .layers import Matrixsheet
        return [Matrixsheet(l) for l in self._obj.GetLayers()]

    def get_layer(self, index: int) -> Matrixsheet:
        """
        Get matrix sheet by index.

        Corresponds to: OriginExt.OriginExt.MatrixPage.GetLayer()

        Args:
            index: Matrix sheet index

        Returns:
            Matrixsheet: The matrix sheet at the specified index
        """
        from .layers import Matrixsheet
        return Matrixsheet(self._obj.GetLayer(index))


class NotePage(PageBase[oext_types.NotePage]):
    """
    Notes window page.
    Wrapper class that wraps OriginExt.OriginExt.NotePage.

    Corresponds to: originpro.Note, OriginExt.OriginExt.NotePage
    """

    def __init__(self, page: oext_types.NotePage):
        """
        Initialize NotePage wrapper.

        Args:
            page: Original OriginExt.NotePage instance to wrap
        """
        super().__init__(page)

    @property
    def Text(self) -> str:
        """Text content of the notes"""
        return self._obj.Text

    def get_text(self) -> str:
        """
        Get the text content.

        Corresponds to: OriginExt.OriginExt.NotePage.GetText()

        Returns:
            str: Text content
        """
        return self._obj.GetText()

    def set_text(self, text: str) -> None:
        """
        Set the text content.

        Corresponds to: OriginExt.OriginExt.NotePage.SetText()

        Args:
            text: Text content
        """
        self._obj.SetText(text)
