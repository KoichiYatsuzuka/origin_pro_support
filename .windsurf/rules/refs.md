---
trigger: always_on
---
# 本ライブラリの目的
OriginProSupportは、OriginLabのPython APIであるOriginProをより扱いやすくするためのラッパークラス群を提供するライブラリ。
内部でOriginExtや_OriginExtを使用する。
外部から本モジュールを読み込んだ際にはoriginproをインポートする必要がないようにしつつ、OriginExtや_OriginExtを隠蔽し、より扱いやすいAPIを提供する。

# 参照する公式ドキュメント
- originproのコードサンプル
  -これらのコード自体は使えないが、呼び出し元をたどっていくことで、参照すべき関数が分かることがある
  - https://www.originlab.com/doc/ja/ExternalPython/External-Python-Code-Samples

- LabTalk ドキュメント
  - OriginExtのメソッドのいくつかはLabTalkで実装されているため、OriginExtと_OriginExtからでは解決できない内容はLabTalkで実装したほうがよい。
  - 以下のURLの下層ページを参照する
  - https://www.originlab.com/doc/ja/LabTalk/ref

- originproのドキュメント
  - クラスメソッドなどの情報が分かる。必要であれば下層のURLを参照。
  - https://docs.originlab.com/originpro/index.html

