# Work Notification by Email

## 目的
- 在宅勤務における就業開始/終了の連絡が面倒くさいのでツール化したい

## 方法
- Python で Outlook を操作してメールの下書きまで作成
- メール送信は自分で行う (誤送信防止)
- 日にちの入力は自動化
- 時刻や業務内容の入力はツール上で手動入力 (したい)
- 入力された時刻や業務内容は config.ini に保存されて次回はその内容を表示
- 最終的には Python の無い環境でも実行したいので exe 化

## ツール
- work_mail.py

## 設定ファイル
- config.ini

## 参考
- [【python】outlookでメール送信](https://qiita.com/watyanabe164/items/e1c21af0127497b74b2a)
- [pythonプログラムにおける設定ファイル管理モジュール～configparserの使い方と注意点～](https://qiita.com/mimitaro/items/3506a444f325c6f980b2)
- [configparser --- 設定ファイルのパーサー](https://docs.python.org/ja/3/library/configparser.html)
- [Python tkinterでボタンを押したときにラジオボタンで選択されている内容を取得したい](https://teratail.com/questions/239658)
- [PythonのTkinterを使ってみる](https://qiita.com/nnahito/items/ad1428a30738b3d93762)
- [Tkinter、Textウィジェットの使い方](https://blog.narito.ninja/detail/100/)
- [Pythonのファイルをexe化する方法【初心者向け】](https://techacademy.jp/magazine/18963)
- [ythonでコンソールウィンドウを非表示にする方法は？](https://www.it-swarm.dev/ja/python/python%E3%81%A7%E3%82%B3%E3%83%B3%E3%82%BD%E3%83%BC%E3%83%AB%E3%82%A6%E3%82%A3%E3%83%B3%E3%83%89%E3%82%A6%E3%82%92%E9%9D%9E%E8%A1%A8%E7%A4%BA%E3%81%AB%E3%81%99%E3%82%8B%E6%96%B9%E6%B3%95%E3%81%AF%EF%BC%9F/957922839/)