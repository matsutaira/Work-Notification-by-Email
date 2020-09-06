# Work Notification by Email

## 目的
- 在宅勤務における就業開始/終了の連絡が面倒くさいのでツール化したい

## 方法
- Python で Outlook を操作してメールの下書きまで作成
- メール送信は自分で行う (誤送信防止)
- 日にちの入力は自動化
- 時刻の入力はツール上で手動入力 (したい)
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