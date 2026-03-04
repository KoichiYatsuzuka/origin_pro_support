"""
Base classes for OriginExt wrappers.

This module contains the abstract base classes that provide common functionality
for all OriginExt wrapper classes, and custom exception classes.
Also contains the APP class for Origin application management.
"""
from __future__ import annotations

import OriginExt as oext
import OriginExt.OriginExt as oext_types

from typing import Iterator, Generic, TypeVar, Optional


# ================== Custom Exception Classes ==================

class OriginNotFoundError(BaseException):
    """Raised when Origin directory or file is not found."""
    pass

class OriginInstanceGenerationError(BaseException):
    """Raised when Origin instance generation fails."""
    pass

class OriginTooManyInstancesError(BaseException):
    """Raised when too many Origin instances are created."""
    pass

class OriginNameConflictError(BaseException):
    """Exception raised when trying to create an object with a name that already exists."""
    pass

class OriginPageGenerationError(BaseException):
    """Exception raised when page generation fails."""
    pass

class OriginCommandResponceError(BaseException):
    """Exception raised when a command response is NaN, null, failed to cast, and so on."""
    pass
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


# ================== Application Management ==================

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


# ================== Abstract Base Wrapper Class ==================

class OriginObjectWrapper(Generic[TOriginObject]):
    """
    Abstract base class for wrapping OriginExt.OriginObject.
    Provides common properties and methods inherited from OriginBase and OriginObject.
    Includes parent and origin instance references for hierarchical access.

    Corresponds to: OriginExt.OriginExt.OriginObject, OriginExt.OriginExt.OriginBase
    """

    def __init__(self, obj: TOriginObject, api_core: 'APP'):
        """
        Initialize the wrapper with OriginExt object and API core reference.

        Args:
            obj: Original OriginExt object to wrap
            api_core: APP instance reference for LabTalk access
        """
        self._obj = obj
        self.__API_core = api_core

    @property
    def api_core(self) -> 'APP':
        """Get the API core reference"""
        return self.__API_core

    def get_origin_instance(self) -> 'APP':
        """
        Get the API core reference.
        
        Returns:
            APP: The API core
        """
        return self.__API_core

    # ================== Properties from OriginObject ==================

    @property
    def name(self) -> str:
        """Short name of the object"""
        return self._obj.Name

    @name.setter
    def name(self, value: str) -> None:
        """Set short name"""
        self._obj.Name = value

    @property
    def long_name(self) -> str:
        """Long name of the object"""
        return self._obj.LongName

    @long_name.setter
    def long_name(self, value: str) -> None:
        """Set long name"""
        self._obj.LongName = value

    @property
    def show(self) -> bool:
        """Visibility state of the object"""
        return self._obj.Show

    @show.setter
    def show(self, value: bool) -> None:
        """Set visibility"""
        self._obj.Show = value

    @property
    def index(self) -> int:
        """Index of the object in its collection"""
        return self._obj.Index

    @property
    def range(self) -> str:
        """Range string representation"""
        return self._obj.Range

    @property
    def type_name(self) -> str:
        """Type name of the object"""
        return self._obj.TypeName

    @property
    def theme(self):
        """Theme of the object"""
        return self._obj.Theme

    @theme.setter
    def theme(self, value) -> None:
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

    def __repr__(self) -> str:
        """String representation"""
        return (
            "type: {}\n".format(type(self)) +
            "short name: {}\n".format(self.Name) +
            "long name: {}".format(self.LongName)
        )

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
