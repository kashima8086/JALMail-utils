# JALMail-utils

<br>
<p align="left">
  <div align="center"><img src="images/jal_work.svg" width="380"></div>
</p>

## 概要
　JAL修行で使ったツール類を集めたリポジトリです。
 
- JAL搭乗予定オートメーション.shortcut<br>
  JAL予約メール → 自動カレンダ登録オートメーション
  iOSで予約通知メールからカレンダーに自動で予定登録をするショートカット

- jal_mailto_route_map.py<br>
  JAL予約メール → 予約経路チェック図 HTML生成ツール
  IMAPメールサーバー上の特定フォルダに保存した予約メールを集計して日本地図上に経路を表示する htmlを生成する Pythonプログラム
  
- JALMail2CSV.py<br>
  JAL予約メール → CSV変換ツール
  IMAPメールサーバー上の特定フォルダに保存した予約メールを集計して CSVファイルに出力する Pythonプログラム
  
- Yahoo!トラベル予約オートメーション.shortcut、APAホテル予約オートメーション.shortcut<br>
  ホテル予約メール → 自動カレンダ登録オートメーション
  Yahoo!トラベル、APAホテルの予約メールからカレンダーに自動で予定登録をするショートカット  
  ※iOS端末への設定は JAL搭乗予定オートメーションの説明を参考にして、予約通知メールの差出人と題名を設定してください。
