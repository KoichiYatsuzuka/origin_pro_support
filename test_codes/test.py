#%%
import origin_pro_support as ops
import os


a = ops.OriginInstance(os.path.join(os.getcwd(), "test.opju"))

root = a.get_root_dir()
# root.create_folder("test_1")
# root.create_folder("test_2")

sub = root.subfolders[0]
print(root)

#%%
import enum

class MyClass:
    name: str
    number: int

class TestEnum(enum.Enum):
    a = MyClass("a", 1)
    b = MyClass("b", 2)
    c = MyClass("c", 3)

print(TestEnum.a.name)

#%%
import originpro as op