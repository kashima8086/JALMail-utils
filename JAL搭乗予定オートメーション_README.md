# JAL搭乗予定オートメーション

<br>
<p align="left">
  <img src="images/jal-scheduled-example.png" width="760">
</p>

## 概要
「JAL搭乗予定オートメーション」は、JALからの予約・購入メールをトリガーとして、メール内容から自動的にカレンダーへ搭乗予定を登録する iOSショートカットです。

---

## 動作内容

- JALからのメール受信をトリガーに自動起動
- メール本文をショートカットに渡す
- フライト情報を解析
- カレンダーへ予定を自動登録

---

## 必要環境

- iOS（ショートカットAppが利用可能なバージョン）
- ショートカットApp
- 標準メールApp

---

# セットアップ手順

## 1. ショートカットのインストール

### 1.1 GitHubから `.shortcut` ファイルをダウンロード

GitHub 上で `.shortcut` ファイルを開き、ダウンロードボタンをタップします。

<p align="center">
  <img src="images/shortcut-download.png" width="500">
</p>

---

### 1.2 Safari のダウンロード一覧から開く

Safari のダウンロード一覧からショートカットを開きます。

<p align="center">
  <img src="images/safari-download.png" width="500">
</p>

---

### 1.3 ショートカットを追加

「ショートカットを追加」をタップします。

<p align="center">
  <img src="images/add-shortcut.png" width="400">
</p>

---

## 2. ショートカットの準備

ショートカット一覧に  
「JAL登場予定オートメーション」が追加されていることを確認します。

<p align="center">
  <img src="images/select-shortcut.png" width="400">
</p>

---

## 3. オートメーションの作成

※ 本ショートカットは、ショートカット入力が無い場合、クリップボードにコピーされているメール本文を処理します。  
そのため、メール受信トリガーを作成しなくても単体で動作確認が可能です。  
動作確認方法は「動作確認」の項を参照してください。

### 3.1 トリガーの選択

「ショートカット」→「オートメーション」→「新規作成」

<p align="center">
  <img src="images/automation-select-template.png" width="400">
</p>

- 「メール」を選択

---

### 3.2 メール条件の設定

<p align="center">
  <img src="images/jal-mail-trigger-when.png" width="400">
</p>

以下を設定します。

- 差出人  
  `no_reply-dom@booking.jal.com`

- 件名に含む  
  `【JAL国内線】購入内容のお知らせ`

---

### 3.3 詳細条件の設定

<p align="center">
  <img src="images/jal-mail-trigger-from.png" width="400">
</p>

必要に応じて設定してください。

- アカウント：任意
- 宛先：任意

---

### 3.4 実行設定

<p align="center">
  <img src="images/jal-mail-trigger-overview.png" width="400">
</p>

- 「すぐに実行」を選択

---

# 4. アクションの設定

### 4.1 「ショートカットを実行」アクションを追加

オートメーションの「行う」に、ショートカット実行アクションを追加します。

まず、アクション一覧から「スクリプティング」を選択します。

<p align="center">
  <img src="images/do-scripting.png" width="400">
</p>

次に、「ショートカットを実行」を選択します。

<p align="center">
  <img src="images/do-exec-shortcut.png" width="400">
</p>

ショートカット選択画面で「JAL搭乗予定オートメーションv1.0」を選択します。

<p align="center">
  <img src="images/do-exec-shortcut-select.png" width="400">
</p>

---

### 4.2 メールを入力として受け取る

<p align="center">
  <img src="images/jal-mail-trigger-do.png" width="400">
</p>

- 「メールを入力として受け取る」が設定されていることを確認

---

### 4.3 入力の設定

<p align="center">
  <img src="images/select-shortcut-input.png" width="400">
</p>

- 入力：ショートカットの入力（メール内容）

---

# 動作確認

## オートメーション動作確認

1. 条件に一致するメールを受信
2. 自動でショートカットが実行される
3. カレンダーへ予定が追加される

- 過去に受信した予約(購入)メールの内容をコピーし、オートメーションのメール受信条件を調整し(題名を「JALメールテスト」とするなど)、自分宛てにメールを送信すると動作確認ができます。

---

## クリップボード経由での追加

メールアプリで JALの予約メール本文をコピーした状態で「JAL登場予定オートメーション」ショートカットを直接実行すると、クリップボード内にある 1 件分のメール内容を解析、カレンダーへ予定を追加できます。

---

# 注意点

- 予約メールのフォーマットが変更された場合、正常に動作しない可能性があります
- 初回実行時にカレンダーアクセス許可が必要です
- 「実行の前に尋ねる」が ON の場合、自動実行されません

---

# カスタマイズ

- 件名条件を変更すれば予約内容のお知らせメールにも対応可能
- ショートカット編集でタイトルやメモ内容を変更可能
- カレンダー登録前に内容を確認したい場合は、  
  ショートカットアプリで「JAL搭乗予定オートメーション」を編集モードで開き、  
  スクリプトの最後の方にある「予定を追加」アクションの「作成シートを表示」を ON にしてください

<p align="center">
  <img src="images/shortcut-show-compose-sheet.png" width="500">
</p>

- ON にすると、予定登録前に内容確認・編集画面が表示されます
- 自動登録したい場合は OFF のままにしてください

---

## ライセンス

本ソフトウェアは MIT License に準拠します。

詳細は `LICENSE` ファイルを参照してください。

