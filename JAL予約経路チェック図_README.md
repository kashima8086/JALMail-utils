# JAL予約経路チェック図 生成ツール

<br>
<p align="left">
  <img src="images/route_viewer_example.png" width="760">
</p>

## 概要

JALの予約メールを解析し、日本地図上に移動経路を表示するHTMLを生成するツールです。

- IMAPからメールを直接取得可能
- 保存済みメール本文テキストでもテスト可能
- 時系列UIで移動経路を順番に確認可能

---

## ファイル構成

| ファイル | 説明 |
|---|---|
| `jal_mail_to_route_map.py` | メインプログラム |
| `JALMailEngine.py` | メール解析エンジン |
| `jal_route_map.tmpl` | SVG地図テンプレート |

---

## 基本実行（IMAP）

```bash
python jal_mail_to_route_map.py
```

---

## 保存済みメールでテスト

```bash
python jal_mail_to_route_map.py --text mail1.txt mail2.txt
```

> ⚠ 保存するメール本文テキストの文字コードは UTF-8 にしてください。  
> Shift_JIS など他の文字コードの場合、正しく解析できません。

---

## オプション

| オプション | 説明 |
|---|---|
| `--output` | 出力HTMLファイル |
| `--template` | テンプレートファイル |
| `--text` | 保存済みメール本文テキスト |
| `--debug-json` | 抽出データJSON出力 |
| `--print-flights` | 抽出フライト一覧表示 |

---

## IMAP設定について

`jal_mail_to_route_map.py` 内で以下を設定してください。

| 変数名 | 説明 |
|---|---|
| `IMAP_HOST` | IMAPサーバー |
| `IMAP_USER` | ログインユーザー名（メールアドレスまたはユーザーID） |
| `IMAP_PASSWORD` | パスワード（またはアプリパスワード） |
| `IMAP_FOLDER` | メールフォルダ |
| `MAIL_SUBJECT_REGEX` | 件名フィルタ（正規表現） |

---

## 設定例

```python
IMAP_HOST = "imap.gmail.com"
IMAP_USER = "your_id"
IMAP_PASSWORD = "your_password"
IMAP_FOLDER = "INBOX"
MAIL_SUBJECT_REGEX = r"JAL.*予約"
```

---

## Gmailを使用する場合の注意

- IMAPを有効にする必要があります
- 2段階認証を有効にしている場合は「アプリパスワード」を使用してください
- 通常のパスワードでは接続できない場合があります

---

## 出力ファイル

```text
jal_route_timeline.html
```

ブラウザで開くことで、予約経路を地図上で確認できます。<br>

- 出力例<br>
　[route_viewer_example.html](https://kashima8086.github.io/JALMail-utils/route_viewer_example.html)

---

## スマートフォン表示モード

スマートフォンで表示した場合は、縦画面向けの専用UIへ自動切り替えされます。

※ スマートフォン専用UIは縦画面表示時のみ有効です。  
横画面表示時はPC向けレイアウトで表示されます。

<br>
<p align="left">
  <img src="images/route_viewer_example_mobile.jpg" width="320">
</p>

### 特徴

- 初期状態では日本全体を縮小表示
- 「進む」「戻る」で現在区間へ自動ズーム
- 現在表示中の区間はオレンジ色で強調表示
- 到着空港もオレンジ色で表示
- 出発済み区間は緑色で保持表示
- 下部コントロールは常に画面下へ固定表示

---

### 下部ステータス表示

画面下部には現在表示中の区間情報を表示します。

- 表示中区間数
- 出発地 → 到着地
- 便名
- 発着時刻

表示例:

```text
表示中: 8 / 8 区間 / 福岡 → 羽田
JAL308 10:05-11:45
```

---

### 操作ボタン

| ボタン | 機能 |
|---|---|
| 戻る | 1区間前へ戻る |
| 進む | 次の区間を表示 |
| 全て | 全区間を表示 |
| リセット | 初期状態へ戻す |

---

### ズームコントロール

右側の丸いボタンでズーム操作が可能です。

| ボタン | 機能 |
|---|---|
| ↔ | 出発地と到着地が収まる範囲へ自動ズーム |
| ＋ | 拡大 |
| － | 縮小 |

---

### 自動ズーム動作

「進む」「戻る」で区間を切り替えた場合は、現在区間の出発地・到着地が画面内に収まるよう自動でズームします。

また、最大拡大時でも到着空港が画面外へ出ないよう補正されています。

---

## デバッグ機能

テンプレート内の設定で区間線の強制表示が可能です。

```javascript
/* ===== デバッグオプション =====
   通常運用では両方とも false / 空文字のままにしてください。

   1) 全区間線を表示したい場合:
      DEBUG_SHOW_ALL_ROUTES = true

   2) 特定空港に関係する区間線だけ表示したい場合:
      DEBUG_ROUTE_AIRPORT = "羽田" のように空港名を指定
      例: "千歳", "羽田", "那覇", "福岡"

   3) デバッグ表示時に発着両方向の矢印も表示する場合:
      DEBUG_SHOW_ROUTE_ARROWS = true

   安全設計:
      本番用の #routes は移動しません。
      デバッグ対象の path だけを #routesDebugOverlay に複製し、
      SVG末尾へ置くことで最前面に表示します。
      そのため、デバッグ無効時のタイムライン機能・通常表示には影響しません。
*/
```

---

## ライセンス

本ソフトウェアは MIT License に準拠します。

詳細は `LICENSE` ファイルを参照してください。
