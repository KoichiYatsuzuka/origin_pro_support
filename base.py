"""
Base classes for OriginExt wrappers.

This module contains the abstract base classes that provide common functionality
for all OriginExt wrapper classes.
"""
from __future__ import annotations

import OriginExt.OriginExt as oext_types

from typing import Iterator, Generic, TypeVar


# ================== Type Variables ==================

T = TypeVar('T')
TOriginObject = TypeVar('TOriginObject', bound=oext_types.OriginObject)


# ================== Collection Types ==================

class OriginCollection(Generic[T]):
    """Generic base class for Origin collections."""
    Count: int
    """Number of items in the collection"""

    def __len__(self) -> int:
        ...

    def __iter__(self) -> Iterator[T]:
        ...

    def __getitem__(self, index: int) -> T:
        ...

    def __call__(self, index: int) -> T:
        """Get item by index (Origin-style access)"""
        ...


# ================== Abstract Base Wrapper Class ==================

class OriginObjectWrapper(Generic[TOriginObject]):
    """
    Abstract base class for wrapping OriginExt.OriginObject.
    Provides common properties and methods inherited from OriginBase and OriginObject.

    Corresponds to: OriginExt.OriginExt.OriginObject, OriginExt.OriginExt.OriginBase
    """

    _obj: TOriginObject

    def __init__(self, obj: TOriginObject):
        """
        Initialize OriginObjectWrapper.

        Args:
            obj: Original OriginExt.OriginObject instance to wrap
        """
        self._obj = obj

    # ================== Properties from OriginObject ==================

    @property
    def Name(self) -> str:
        """Short name of the object"""
        return self._obj.Name

    @Name.setter
    def Name(self, value: str) -> None:
        """Set short name"""
        self._obj.Name = value

    @property
    def LongName(self) -> str:
        """Long name of the object"""
        return self._obj.LongName

    @LongName.setter
    def LongName(self, value: str) -> None:
        """Set long name"""
        self._obj.LongName = value

    @property
    def Show(self) -> bool:
        """Visibility state of the object"""
        return self._obj.Show

    @Show.setter
    def Show(self, value: bool) -> None:
        """Set visibility"""
        self._obj.Show = value

    @property
    def Index(self) -> int:
        """Index of the object in its collection"""
        return self._obj.Index

    @property
    def Range(self) -> str:
        """Range string representation"""
        return self._obj.Range

    @property
    def TypeName(self) -> str:
        """Type name of the object"""
        return self._obj.TypeName

    @property
    def Theme(self):
        """Theme of the object"""
        return self._obj.Theme

    @Theme.setter
    def Theme(self, value) -> None:
        """Set theme"""
        self._obj.Theme = value

    # ================== Magic Methods ==================

    def __str__(self) -> str:
        """String representation"""
        return str(self._obj)

    def __len__(self) -> int:
        """Length/count of the object"""
        return len(self._obj)

    def __bool__(self) -> bool:
        """Boolean validity check"""
        return bool(self._obj)

    # ================== Methods from OriginObject (snake_case) ==================

    def is_valid(self) -> bool:
        """
        Check if this object is valid.

        Corresponds to: OriginExt.OriginExt.OriginObject.IsValid()

        Returns:
            bool: True if valid
        """
        return self._obj.IsValid()

    def get_name(self) -> str:
        """
        Get short name.

        Corresponds to: OriginExt.OriginExt.OriginObject.GetName()

        Returns:
            str: Short name
        """
        return self._obj.GetName()

    def set_name(self, name: str) -> None:
        """
        Set short name.

        Corresponds to: OriginExt.OriginExt.OriginObject.SetName()

        Args:
            name: New short name
        """
        self._obj.SetName(name)

    def get_long_name(self) -> str:
        """
        Get long name.

        Corresponds to: OriginExt.OriginExt.OriginObject.GetLongName()

        Returns:
            str: Long name
        """
        return self._obj.GetLongName()

    def set_long_name(self, name: str) -> None:
        """
        Set long name.

        Corresponds to: OriginExt.OriginExt.OriginObject.SetLongName()

        Args:
            name: New long name
        """
        self._obj.SetLongName(name)

    def destroy(self) -> None:
        """
        Destroy this object.

        Corresponds to: OriginExt.OriginExt.OriginObject.Destroy()
        """
        self._obj.Destroy()

    def activate(self) -> None:
        """
        Activate this object.

        Corresponds to: OriginExt.OriginExt.OriginObject.Activate()
        """
        self._obj.Activate()

    def execute(self, labtalk_str: str) -> bool:
        """
        Execute a LabTalk command on this object.

        Corresponds to: OriginExt.OriginExt.OriginObject.Execute()

        Args:
            labtalk_str: LabTalk command string

        Returns:
            bool: Execution result
        """
        return self._obj.Execute(labtalk_str)

    def get_num_prop(self, prop_name: str) -> float:
        """
        Get a numeric property.

        Corresponds to: OriginExt.OriginExt.OriginObject.GetNumProp()

        Args:
            prop_name: Property name

        Returns:
            float: Property value
        """
        return self._obj.GetNumProp(prop_name)

    def set_num_prop(self, prop_name: str, value: float) -> None:
        """
        Set a numeric property.

        Corresponds to: OriginExt.OriginExt.OriginObject.SetNumProp()

        Args:
            prop_name: Property name
            value: Property value
        """
        self._obj.SetNumProp(prop_name, value)

    def get_str_prop(self, prop_name: str) -> str:
        """
        Get a string property.

        Corresponds to: OriginExt.OriginExt.OriginObject.GetStrProp()

        Args:
            prop_name: Property name

        Returns:
            str: Property value
        """
        return self._obj.GetStrProp(prop_name)

    def set_str_prop(self, prop_name: str, value: str) -> None:
        """
        Set a string property.

        Corresponds to: OriginExt.OriginExt.OriginObject.SetStrProp()

        Args:
            prop_name: Property name
            value: Property value
        """
        self._obj.SetStrProp(prop_name, value)

    def do_method(self, cmd: str, arg: str | None = None) -> float:
        """
        Execute a method command.

        Corresponds to: OriginExt.OriginExt.OriginObject.DoMethod()

        Args:
            cmd: Command string
            arg: Optional argument

        Returns:
            float: Result
        """
        return self._obj.DoMethod(cmd, arg)

    def do_str_method(self, cmd: str, arg: str | None = None) -> str:
        """
        Execute a method command that returns a string.

        Corresponds to: OriginExt.OriginExt.OriginObject.DoStrMethod()

        Args:
            cmd: Command string
            arg: Optional argument

        Returns:
            str: Result
        """
        return self._obj.DoStrMethod(cmd, arg)

    def get_meta_data(self, name: str, visible_to_user: bool) -> str:
        """
        Get metadata.

        Corresponds to: OriginExt.OriginExt.OriginObject.GetMetaData()

        Args:
            name: Metadata name
            visible_to_user: Visibility flag

        Returns:
            str: Metadata XML
        """
        return self._obj.GetMetaData(name, visible_to_user)

    def set_meta_data(self, xml: str, name: str, visible_to_user: bool) -> None:
        """
        Set metadata.

        Corresponds to: OriginExt.OriginExt.OriginObject.SetMetaData()

        Args:
            xml: Metadata XML
            name: Metadata name
            visible_to_user: Visibility flag
        """
        self._obj.SetMetaData(xml, name, visible_to_user)
