# KindleBreak

macOS の Kindle アプリから本を自動キャプチャし、Gemini API で OCR テキスト化するツール。

## できること

| モード | コマンド | 説明 |
|--------|----------|------|
| **単体キャプチャ** | `python3 kindlebreak.py run` | 今開いている本を1冊まるごとキャプチャ → PDF → テキスト化 |
| **コレクション一括** | `python3 kindlebreak.py collection` | Kindle のコレクション画面から全書籍を順番に処理 |
| **OCR のみ** | `python3 kindlebreak.py ocr [PDF]` | 既存の PDF や画像から OCR だけ実行 |

## 動作の流れ

```
Kindle アプリ → スクリーンキャプチャ → ページめくり自動化 → PNG → PDF → Gemini OCR → テキスト
```

- ページが変わらなくなったら自動で最終ページを検出して停止
- マウスを画面左上隅に移動すると緊急停止（pyautogui のフェイルセーフ）

## 必要なもの

- **macOS**（screencapture / AppleScript を使うため、Windows / Linux では動きません）
- **Python 3.10+**
- **Kindle デスクトップアプリ**（Amazon 公式の Mac 版）
- **Gemini API キー**（[Google AI Studio](https://aistudio.google.com/apikey) で無料取得）

## セットアップ

### 1. リポジトリをクローン

```bash
git clone https://github.com/isshii-jpeg/kindlebreak.git
cd kindlebreak
```

### 2. Python 仮想環境を作成＆依存パッケージをインストール

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### 3. API キーを設定

```bash
cp .env.example .env
```

`.env` を開いて自分の Gemini API キーを設定：

```
GEMINI_API_KEY=your-api-key-here
```

### 4. macOS のアクセシビリティ権限を許可

このツールは画面キャプチャとキー入力の自動化を行うため、以下の権限が必要です：

1. **システム設定 → プライバシーとセキュリティ → アクセシビリティ** にターミナル（または使用しているターミナルアプリ）を追加
2. **システム設定 → プライバシーとセキュリティ → 画面収録** にも同様に追加

## 使い方

### 1冊キャプチャ（`run`）

1. Kindle アプリで本を開き、**最初のページ**を表示しておく
2. 以下を実行：

```bash
python3 kindlebreak.py run
```

途中のページから再開する場合：

```bash
python3 kindlebreak.py run --start-page 50
```

出力：
- `pages/` — 各ページの PNG 画像
- `book.pdf` — 全ページ結合 PDF
- `output.txt` — OCR テキスト

### コレクション一括（`collection`）

1. Kindle アプリで**コレクション一覧画面**（本の表紙がグリッド表示されている画面）を開く
2. 以下を実行：

```bash
python3 kindlebreak.py collection
```

出力は `books/` ディレクトリ以下に本ごとのフォルダが作られます：

```
books/
  本のタイトル1/
    book.pdf
    output.txt
    page_0001.png
    ...
  本のタイトル2/
    ...
```

### OCR のみ（`ocr`）

既存の PDF からテキストを抽出：

```bash
python3 kindlebreak.py ocr path/to/book.pdf
```

## 設定のカスタマイズ

`kindlebreak.py` の `CONFIG` クラスで動作を調整できます：

```python
@dataclass
class CONFIG:
    PAGE_WAIT: float = 1.5          # ページめくり後の待機時間（秒）
    INITIAL_WAIT: float = 3.0       # 開始前の待機時間
    SIMILARITY_THRESHOLD: float = 0.999  # 最終ページ検出の閾値（下げると早めに停止）
    MAX_PAGES: int = 5000           # 最大ページ数の上限
    GRID_COLS: int = 5              # コレクションのグリッド列数
    GRID_ROWS: int = 4              # グリッド行数
```

**よくある調整：**

- ページめくりが速すぎてキャプチャが乱れる → `PAGE_WAIT` を大きくする（例: `2.0`）
- 最終ページ検出が甘い / 途中で止まる → `SIMILARITY_THRESHOLD` を調整
- コレクションのグリッドレイアウトが合わない → `GRID_COLS` / `GRID_ROWS` を変更

## 改造ガイド

### OCR プロンプトを変える

`ocr_with_gemini()` 関数内の `prompt` 変数を編集してください。例えば英語の本なら：

```python
prompt = (
    "Transcribe the text on this page exactly as written.\n"
    "Rules:\n"
    "- Output only the body text\n"
    "- Exclude headers, footers, page numbers\n"
    "- Break text by paragraph, not by line\n"
)
```

### OCR モデルを変える

`gemini-2.5-flash` の部分を変更すれば別のモデルを使えます：

```python
model = genai.GenerativeModel("gemini-2.5-pro")  # より高精度
```

### ページめくり方向を変える

`kindle_next_page()` 関数の `key code 123`（←キー）を `key code 124`（→キー）に変えると逆方向にめくれます。

### 出力先を変える

`CONFIG` クラスの `OUTPUT_BASE`、`OUTPUT_DIR`、`OUTPUT_PDF`、`OUTPUT_TEXT` を変更してください。

## トラブルシューティング

| 症状 | 対処 |
|------|------|
| 「Kindle ウィンドウが見つかりません」 | Kindle アプリを起動してください |
| 「GEMINI_API_KEY 環境変数が設定されていません」 | `.env` ファイルに API キーを設定してください |
| キャプチャが真っ黒 / 画面が映らない | システム設定で画面収録の権限を許可してください |
| ページめくりが動かない | システム設定でアクセシビリティの権限を許可してください |
| 最終ページなのに止まらない | `SIMILARITY_THRESHOLD` を少し下げてください |
| 途中で勝手に止まる | `SIMILARITY_THRESHOLD` を上げるか、`PAGE_WAIT` を長くしてください |

## ライセンス

個人利用を想定しています。Kindle の利用規約を遵守してご利用ください。
