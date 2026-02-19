#%%
import origin_pro_support as ops
import os

a = ops.OriginInstance(os.path.join(os.getcwd(), "test.opju"))

root = a.get_root_dir()
# root.create_folder("test_1")
# root.create_folder("test_2")

sub = root.subfolders[0]
print(root)


