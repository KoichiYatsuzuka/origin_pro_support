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

# コーディングルール
- Noneは基本的に返さない。例外は明示的にraiseする。
- テストコード以外で例外は必ずさらにraiseするか、警告を表示する。
- 文中importは絶対に行わない。必要性があるのであれば、クラスの構成自体を変える必要があるため、警告したうえでコーディングを停止する。
- try, if文のネストは可能な限り浅くする。
- - 一つの関数でtry節は2つまでにする。それ以上のネストは、大きなtry節の中で個別の例外をcatchするようにする。
- - if文のネストが深くなりそうな場合、短い複数のif文を逐次評価することで避けられないか検討する。それも難しい場合はandで条件をつなげたり、例外をraiseしたりしてネストを避ける。
- - これらの回避のためにラムダ式は使わない。