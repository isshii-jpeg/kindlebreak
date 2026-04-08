# PageSnap

macOS 上のデスクトップウィンドウを定期的にスクリーンショットし、キーイベントによるページ送りを自動化して、**画像 → PDF → OCR テキスト**のパイプラインを実行する Python ツール。

pyautogui + AppleScript によるウィンドウ操作の自動化と、Gemini API を使った OCR テキスト抽出を組み合わせた汎用的なデスクトップオートメーションの実装例です。

## 機能

| モード | コマンド | 説明 |
|--------|----------|------|
| **キャプチャ** | `python3 pagesnap.py capture` | 対象ウィンドウをページ送りしながら自動キャプチャ → PDF → OCR |
| **バッチ処理** | `python3 pagesnap.py batch` | グリッド一覧画面から複数ドキュメントを一括処理 |
| **OCR のみ** | `python3 pagesnap.py ocr [PDF]` | 既存の PDF からテキスト抽出だけ実行 |

## 技術スタック

- **pyautogui** — クロスプラットフォームの GUI 自動化（クリック、キー入力、マウス移動）
- **AppleScript / Quartz** — macOS ネイティブのウィンドウ検出・アクティベーション
- **screencapture** — macOS 標準のウィンドウ単位スクリーンショット
- **NumPy** — ピクセル単位の画像類似度比較（最終ページ自動検出）
- **Pillow** — 画像処理・PNG→PDF 変換
- **PyMuPDF (fitz)** — PDF → 画像変換
- **Gemini API** — マルチモーダル LLM による高精度 OCR

## 動作の流れ

```
対象アプリのウィンドウを検出
  ↓
スクリーンキャプチャ → PNG 保存
  ↓
キーイベントでページ送り
  ↓
前ページとの類似度比較 → 同一なら最終ページと判定して停止
  ↓
PNG 一括 → PDF 結合
  ↓
Gemini API で各ページを OCR → テキストファイル出力
```

## 必要なもの

- **macOS**（screencapture / AppleScript を使用）
- **Python 3.10+**
- **Gemini API キー**（[Google AI Studio](https://aistudio.google.com/apikey) で無料取得）

## セットアップ

```bash
git clone https://github.com/isshii-jpeg/pagesnap.git
cd pagesnap

python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

cp .env.example .env
# .env を編集して Gemini API キーを設定
```

### macOS 権限の設定

ウィンドウ操作の自動化に以下の権限が必要です：

- **システム設定 → プライバシーとセキュリティ → アクセシビリティ** にターミナルアプリを追加
- **システム設定 → プライバシーとセキュリティ → 画面収録** にも追加

## 使い方

### 基本: ウィンドウキャプチャ

対象アプリでドキュメントを開いた状態で：

```bash
python3 pagesnap.py capture
```

途中から再開：

```bash
python3 pagesnap.py capture --start-page 50
```

別のアプリを対象にする場合：

```bash
python3 pagesnap.py capture --app "Preview"
```

### バッチ処理

グリッド一覧（サムネイルが並ぶ画面）を表示した状態で：

```bash
python3 pagesnap.py batch
```

各ドキュメントごとにフォルダが作られます：

```
output/
  ドキュメント1/
    document.pdf
    output.txt
    page_0001.png ...
  ドキュメント2/
    ...
```

### OCR のみ

```bash
python3 pagesnap.py ocr path/to/document.pdf
```

## 設定のカスタマイズ

`pagesnap.py` の `CONFIG` クラスで動作を調整できます：

```python
@dataclass
class CONFIG:
    APP_NAME: str = "Kindle"           # 対象アプリ名（任意のアプリに変更可）
    APP_BUNDLE_ID: str = "com.amazon.Lassen"
    PAGE_WAIT: float = 1.5             # ページ送り後の待機時間（秒）
    NEXT_PAGE_KEY_CODE: int = 123      # ページ送りキー（123=左矢印, 124=右矢印）
    SIMILARITY_THRESHOLD: float = 0.999  # 最終ページ検出の閾値
    MAX_PAGES: int = 5000
    GRID_COLS: int = 5                 # バッチモードのグリッド列数
    GRID_ROWS: int = 4                 # バッチモードのグリッド行数
```

### 別のアプリに対応させるには

1. `APP_NAME` をアプリのプロセス名に変更
2. `APP_BUNDLE_ID` をアプリのバンドルIDに変更（`mdls -name kMDItemCFBundleIdentifier /Applications/AppName.app` で確認）
3. `NEXT_PAGE_KEY_CODE` をそのアプリのページ送りキーに変更

### macOS キーコード一覧（よく使うもの）

| キー | コード |
|------|--------|
| ← | 123 |
| → | 124 |
| ↑ | 126 |
| ↓ | 125 |
| Space | 49 |
| Return | 36 |

## 仕組みの解説

### ウィンドウ検出 (Quartz API)

macOS の `CGWindowListCopyWindowInfo` を使って、アプリ名でウィンドウを検索し、ウィンドウIDと座標を取得しています。これにより、ウィンドウが他のウィンドウに隠れていても正確にキャプチャできます。

### ページ送りの自動化 (AppleScript + System Events)

`osascript` 経由で AppleScript を実行し、System Events の `key code` でキーイベントを送信しています。pyautogui のキー入力はアプリによっては効かない場合がありますが、AppleScript 経由なら確実に動作します。

### 最終ページの自動検出 (NumPy)

連続する2枚のスクリーンショットをピクセル単位で比較し、類似度が閾値（デフォルト 99.9%）を超えたら「ページが変わっていない＝最終ページ」と判定します。

### OCR (Gemini API)

Gemini のマルチモーダル入力に画像を直接渡して、テキストを抽出しています。従来の Tesseract ベースの OCR と比較して、日本語の縦書きやレイアウトが複雑なドキュメントでも高い精度が出ます。

## トラブルシューティング

| 症状 | 対処 |
|------|------|
| 「ウィンドウが見つかりません」 | 対象アプリを起動して `APP_NAME` が正しいか確認 |
| 「GEMINI_API_KEY が設定されていません」 | `.env` ファイルに API キーを設定 |
| キャプチャが真っ黒 | システム設定で画面収録の権限を許可 |
| ページ送りが動かない | システム設定でアクセシビリティの権限を許可 |
| 最終ページなのに止まらない | `SIMILARITY_THRESHOLD` を少し下げる |
| 途中で止まる | `SIMILARITY_THRESHOLD` を上げるか `PAGE_WAIT` を長くする |

## ライセンス

MIT License
