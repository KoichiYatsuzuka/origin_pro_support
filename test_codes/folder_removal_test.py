from origin_pro_support.pages import PageBase, GraphPage, WorkbookPage
from origin_pro_support.folder import Folder

print('=== Folderクラス削除テスト ===')
print('pages.pyからFolderをインポート:', 'Folder' in dir())
print('folder.pyからFolderをインポート:', 'Folder' in dir(Folder.__module__))
print('Folderクラスのモジュール:', Folder.__module__)
print('Folderクラス削除完了')
