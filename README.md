# OriginProSupport

Originの外部Pythonライブラリをより扱いやすくするためのラッパークラス群を提供するライブラリ。内部でOriginExtや_OriginExtを使用し、外部から本モジュールを読み込んだ際にはoriginproをインポートする必要がないようにしつつ、OriginExtや_OriginExtを隠蔽し、より扱いやすいAPIを提供する。

## 必須要件

- OriginPro 2019以降がインストールされていること（外部PythonからOriginのAPIが叩けること）。それ以前のOriginだとエラーが発生して停止します。
- Origin**Pro**であること。（外部Pythonから操作できるのはPro版限定の機能です）
- Python 3.10以上（使用している言語機能上の制約）

## 特徴
- オブジェクトを階層構造にして、より直観的に操作できるようにしました。
- 複数のウィンドウ同時呼び出しや、一度閉じたウィンドウをもう一度呼び出すことができるので、Jupyter NotebookなどでPythonインスタンスを終了していない場合でも、再度ウィンドウを呼び出せます。
- PandasやNumpyからのWorkseet作成をサポート

## 注意事項

外部PythonからOriginを起動したとき、PythonがOriginをバインドしたままになるため、手動でウィンドウを閉じれなくなることに注意してください。
OriginInstance.close()を呼び出すことで、Originを閉じることができます。


## クラス構成

### クラス依存関係ツリー

```
OriginInstance (アプリケーション管理)
├── Folder (フォルダ管理)
│   ├── WorkbookPage (ワークブックページ)
│   │   └── Worksheet (ワークシート)
│   │       ├── Column (列)
│   │       └── ...
│   ├── GraphPage (グラフページ)
│   │   └── GraphLayer (グラフレイヤ)
│   │       └── DataPlot (データプロット)
│   ├── MatrixPage (行列ページ)
│   │   └── Matrixsheet (行列シート)
│   ├── NotePage (ノートページ)
│   └── FigurePage (図ページ)
└── ...

基底クラス:
OriginObjectWrapper (すべてのラッパーの基底)
├── PageBase (ページクラスの基底)
│   ├── WorkbookPage
│   ├── GraphPage
│   ├── MatrixPage
│   ├── NotePage
│   └── FigurePage
├── Layer (レイヤクラスの基底)
│   ├── Datasheet
│   │   └── Worksheet
│   └── GraphLayer
└── ...
```

### コアクラス
