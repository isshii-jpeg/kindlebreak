"""
Kindle自動ページめくり＆OCRテキスト化システム

使い方:
  1. python3 kindlebreak.py collection  → コレクション一覧から全書籍をキャプチャ→PDF→OCR
  2. python3 kindlebreak.py run         → 開いている本を1冊キャプチャ→PDF→OCR
  3. python3 kindlebreak.py ocr [pdf]   → 既存PDF/画像からOCRだけ実行
"""

import sys
import os
import time
import subprocess
import argparse
import tempfile
import json
import re
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv
load_dotenv()

import pyautogui
import numpy as np
from PIL import Image


# ============================================================
# CONFIG
# ============================================================
@dataclass
class CONFIG:
    # --- タイミング ---
    PAGE_WAIT: float = 1.5
    INITIAL_WAIT: float = 3.0
    BOOK_OPEN_WAIT: float = 2.0   # 本を開いた後の待機

    # --- 出力 ---
    OUTPUT_BASE: str = "books"    # コレクションモードの出力ベースディレクトリ
    OUTPUT_DIR: str = "pages"     # 単体モードの画像保存先
    OUTPUT_PDF: str = "book.pdf"
    OUTPUT_TEXT: str = "output.txt"

    # --- 終了検出 ---
    SIMILARITY_THRESHOLD: float = 0.999
    MAX_PAGES: int = 5000

    # --- コレクション ---
    GRID_COLS: int = 5            # コレクションのグリッド列数
    GRID_ROWS: int = 4            # 画面内に見えるグリッド行数

    # --- 安全装置 ---
    FAILSAFE: bool = True


cfg = CONFIG()

pyautogui.FAILSAFE = cfg.FAILSAFE
pyautogui.PAUSE = 0.3


# ============================================================
# Kindleウィンドウ操作 (macOS)
# ============================================================
def get_kindle_window_id():
    """KindleウィンドウのIDと位置を取得する。"""
    result = subprocess.run(
        ["python3", "-c", """
import json
try:
    import Quartz
    windows = Quartz.CGWindowListCopyWindowInfo(Quartz.kCGWindowListOptionAll, Quartz.kCGNullWindowID)
    best = None
    for w in windows:
        name = w.get('kCGWindowOwnerName', '')
        if 'Kindle' in name and w.get('kCGWindowLayer', -1) == 0:
            bounds = dict(w.get('kCGWindowBounds', {}))
            area = bounds.get('Width', 0) * bounds.get('Height', 0)
            if best is None or area > best[1]:
                best = ({'id': int(w['kCGWindowNumber']), 'bounds': bounds}, area)
    if best:
        print(json.dumps(best[0]))
    else:
        print('NOT_FOUND')
except Exception as e:
    print(f'ERROR:{e}')
"""],
        capture_output=True, text=True, timeout=10
    )

    output = result.stdout.strip()
    if not output or output == "NOT_FOUND":
        print("エラー: Kindleウィンドウが見つかりません。")
        print("  Kindleアプリを開いてください。")
        sys.exit(1)
    if output.startswith("ERROR:"):
        print(f"エラー: {output}")
        sys.exit(1)

    info = json.loads(output)
    return info["id"], info["bounds"]


def activate_kindle():
    """Kindleを最前面にする。"""
    subprocess.run(
        ["osascript", "-e", 'tell application id "com.amazon.Lassen" to activate'],
        capture_output=True, timeout=5
    )
    time.sleep(0.5)


def kindle_next_page():
    """左矢印キーでページめくり。"""
    subprocess.run(
        ["osascript", "-e", '''
        tell application id "com.amazon.Lassen" to activate
        delay 0.3
        tell application "System Events"
            tell process "Kindle"
                key code 123
            end tell
        end tell
        '''],
        capture_output=True, timeout=10
    )


def kindle_click(x: int, y: int):
    """Kindleをアクティブにして指定座標をクリック。"""
    activate_kindle()
    pyautogui.click(x, y)


def get_book_title_from_image(img: Image.Image) -> str:
    """本を開いた状態のスクリーンショットからタイトルをGeminiで読み取る。"""
    import google.generativeai as genai

    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        return "untitled"

    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-2.5-flash")

    # 上部だけ切り出す（タイトルバー付近）
    w, h = img.size
    header = img.crop((0, 0, w, min(80, h)))

    try:
        response = model.generate_content([
            header,
            "この画像はKindleアプリのヘッダー部分です。本のタイトルだけを返してください。余計な文字は不要です。"
        ])
        title = response.text.strip()
        return title if title else "untitled"
    except Exception:
        return "untitled"


def capture_kindle_window(window_id: int) -> Image.Image:
    """ウィンドウIDを指定してキャプチャ。"""
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
        tmp_path = tmp.name
    subprocess.run(
        ["screencapture", "-l", str(window_id), "-x", tmp_path],
        check=True, timeout=10
    )
    img = Image.open(tmp_path)
    Path(tmp_path).unlink(missing_ok=True)
    return img


def sanitize_filename(name: str) -> str:
    """ファイル名に使えない文字を除去する。"""
    name = re.sub(r'[\\/:*?"<>|]', '', name)
    name = name.strip('. ')
    return name[:100] if name else "untitled"


# ============================================================
# 画像比較
# ============================================================
def images_are_similar(img1: Image.Image, img2: Image.Image) -> bool:
    """2枚の画像がほぼ同一かどうかを判定する。"""
    arr1 = np.array(img1, dtype=np.float32)
    arr2 = np.array(img2, dtype=np.float32)
    if arr1.shape != arr2.shape:
        return False
    diff = np.abs(arr1 - arr2) / 255.0
    similarity = 1.0 - diff.mean()
    return similarity >= cfg.SIMILARITY_THRESHOLD


# ============================================================
# 1冊分のキャプチャ → PDF保存
# ============================================================
def capture_book(window_id: int, output_dir: Path) -> int:
    """現在開いている本の全ページをキャプチャし、PNGとPDFを保存する。
    戻り値: キャプチャしたページ数
    """
    output_dir.mkdir(parents=True, exist_ok=True)

    prev_screenshot = None
    page = 1

    while page <= cfg.MAX_PAGES:
        try:
            time.sleep(cfg.PAGE_WAIT)
            screenshot = capture_kindle_window(window_id)

            if prev_screenshot is not None and images_are_similar(screenshot, prev_screenshot):
                print(f"    最終ページ検出（{page - 1}ページ）")
                break

            save_path = output_dir / f"page_{page:04d}.png"
            screenshot.save(str(save_path))
            print(f"    [ページ {page}]", end=" ", flush=True)

            prev_screenshot = screenshot
            kindle_next_page()
            page += 1

        except pyautogui.FailSafeException:
            print(f"\n    緊急停止！ 最終処理ページ: {page}")
            break
        except Exception as e:
            print(f"\n    エラー (ページ {page}): {e}")
            break

    # PNG → PDF
    total = page - 1
    if total > 0:
        png_files = sorted(output_dir.glob("page_*.png"))
        images = [Image.open(f).convert("RGB") for f in png_files]
        pdf_path = output_dir / "book.pdf"
        images[0].save(str(pdf_path), save_all=True, append_images=images[1:], resolution=150)
        print(f"    → PDF保存: {pdf_path.name} ({total}ページ)")

    return total


# ============================================================
# コレクション操作
# ============================================================
def click_back_button(bounds: dict):
    """本のビューからコレクションに戻る（<ボタンをクリック）。"""
    activate_kindle()
    # まずウィンドウ上部にマウスを移動してヘッダーを表示させる
    top_center_x = int(bounds["X"]) + int(bounds["Width"]) // 2
    top_y = int(bounds["Y"]) + 40
    pyautogui.moveTo(top_center_x, top_y)
    time.sleep(1.0)
    # <ボタンをクリック
    back_x = int(bounds["X"]) + 25
    back_y = int(bounds["Y"]) + 45
    pyautogui.click(back_x, back_y)
    time.sleep(2.0)


def get_book_grid_positions(bounds: dict) -> list[tuple[int, int]]:
    """コレクションのグリッド内の各本のクリック座標を返す。"""
    win_x = int(bounds["X"])
    win_y = int(bounds["Y"])
    win_w = int(bounds["Width"])
    win_h = int(bounds["Height"])

    # グリッドの開始位置と範囲（ツールバー分を除外）
    grid_top = win_y + 80       # ツールバー下
    grid_left = win_x + 20
    grid_width = win_w - 40
    grid_height = win_h - 100

    cell_w = grid_width // cfg.GRID_COLS
    cell_h = grid_height // cfg.GRID_ROWS

    positions = []
    for row in range(cfg.GRID_ROWS):
        for col in range(cfg.GRID_COLS):
            cx = grid_left + col * cell_w + cell_w // 2
            cy = grid_top + row * cell_h + cell_h // 2
            positions.append((cx, cy))

    return positions


def scroll_collection(bounds: dict):
    """コレクションを1画面分下にスクロール。"""
    center_x = int(bounds["X"]) + int(bounds["Width"]) // 2
    center_y = int(bounds["Y"]) + int(bounds["Height"]) // 2
    activate_kindle()
    pyautogui.moveTo(center_x, center_y)
    pyautogui.scroll(-5)  # 下スクロール
    time.sleep(1.5)


def run_collection():
    """コレクション一覧から全書籍を順番にキャプチャ → PDF → OCR。"""

    print("Kindleウィンドウを検出中...")
    window_id, bounds = get_kindle_window_id()
    activate_kindle()

    base_dir = Path(cfg.OUTPUT_BASE)
    base_dir.mkdir(exist_ok=True)

    print("=" * 50)
    print("コレクション一括処理")
    print(f"  出力先: {base_dir.resolve()}/")
    print("=" * 50)

    # ── 第1段階: 全書籍のキャプチャ & PDF化 ──
    print("\n【第1段階】全書籍をキャプチャ & PDF化")
    print("-" * 50)

    captured_books = []  # (title, book_dir) のリスト
    processed_titles = set()

    while True:
        # 現在の画面でのグリッド座標を取得
        collection_before = capture_kindle_window(window_id)
        positions = get_book_grid_positions(bounds)

        found_new_book = False

        for pos_idx, (cx, cy) in enumerate(positions):
            # コレクション画面をキャプチャ（クリック前）
            before = capture_kindle_window(window_id)

            # 本をクリック
            kindle_click(cx, cy)
            time.sleep(cfg.BOOK_OPEN_WAIT)

            # クリック後のスクリーンショット
            after = capture_kindle_window(window_id)

            # 画面が変わっていなければ空セル → スキップ
            if images_are_similar(before, after):
                continue

            # 本が開いた → タイトルを画像から取得
            title = get_book_title_from_image(after)
            safe_title = sanitize_filename(title)

            # 既に処理済みならスキップ
            if safe_title in processed_titles:
                click_back_button(bounds)
                time.sleep(1)
                continue

            found_new_book = True
            print(f"\n📖 [{len(captured_books) + 1}] {title}")

            # キャプチャ
            book_dir = base_dir / safe_title
            page_count = capture_book(window_id, book_dir)

            if page_count > 0:
                captured_books.append((title, book_dir))
                processed_titles.add(safe_title)

            # コレクションに戻る
            click_back_button(bounds)
            time.sleep(1)

        # スクロールして次の行を探す
        scroll_before = capture_kindle_window(window_id)
        scroll_collection(bounds)
        scroll_after = capture_kindle_window(window_id)

        # スクロールしても変わらない → 全書籍処理済み
        if images_are_similar(scroll_before, scroll_after):
            if not found_new_book:
                break

    print(f"\n{'=' * 50}")
    print(f"キャプチャ完了！ {len(captured_books)} 冊")
    for i, (title, _) in enumerate(captured_books, 1):
        print(f"  {i}. {title}")

    if not captured_books:
        print("キャプチャされた本がありません。")
        return

    # ── 第2段階: 全書籍をGemini OCRでテキスト化 ──
    print(f"\n{'=' * 50}")
    print("【第2段階】Gemini OCRでテキスト化")
    print("-" * 50)

    for i, (title, book_dir) in enumerate(captured_books, 1):
        print(f"\n📖 [{i}/{len(captured_books)}] {title}")
        pdf_path = book_dir / "book.pdf"
        text_path = book_dir / "output.txt"

        if not pdf_path.exists():
            print("  PDF が見つかりません。スキップ。")
            continue

        images = pdf_to_images(pdf_path)
        ocr_with_gemini(images, text_path)

    print(f"\n{'=' * 50}")
    print(f"全処理完了！ {len(captured_books)} 冊")
    print(f"出力先: {base_dir.resolve()}/")


# ============================================================
# PNG → PDF
# ============================================================
def pngs_to_pdf(image_dir: Path, output_pdf: Path):
    """ページ画像をまとめて1つのPDFにする。"""
    png_files = sorted(image_dir.glob("page_*.png"))
    if not png_files:
        print("エラー: キャプチャ画像が見つかりません。")
        sys.exit(1)

    print(f"\nPDF作成中... ({len(png_files)} ページ)")
    images = []
    for f in png_files:
        img = Image.open(f).convert("RGB")
        images.append(img)

    images[0].save(
        str(output_pdf),
        save_all=True,
        append_images=images[1:],
        resolution=150,
    )
    print(f"PDF保存: {output_pdf.resolve()}")
    return png_files


# ============================================================
# Gemini OCR
# ============================================================
def pdf_to_images(pdf_path: Path) -> list[Image.Image]:
    """PDFファイルを画像のリストに変換する。"""
    import fitz
    doc = fitz.open(str(pdf_path))
    images = []
    for page in doc:
        pix = page.get_pixmap(dpi=150)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        images.append(img)
    doc.close()
    return images


def ocr_with_gemini(images: list[Image.Image], output_text: Path):
    """画像リストをGemini APIに送ってOCRする。"""
    import google.generativeai as genai

    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        print("エラー: GEMINI_API_KEY 環境変数が設定されていません。")
        print("  .env ファイルに GEMINI_API_KEY=your-key を設定してください。")
        sys.exit(1)

    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-2.5-flash")

    total = len(images)
    all_text = []

    print(f"  Gemini OCR: {total} ページ")

    prompt = (
        "この画像は本の1ページです。書かれているテキストをそのまま正確に書き起こしてください。\n"
        "ルール:\n"
        "- 本文のテキストだけを出力すること\n"
        "- ヘッダー、フッター、ページ番号、シリーズ名、アプリのUI要素は除外すること\n"
        "- 原文の改行位置を再現する必要はない。段落ごとに改行すること\n"
        "- 挿絵や図表のみのページは空で返すこと\n"
        "- 説明や注釈は一切付けないこと"
    )

    for i, img in enumerate(images):
        page_num = i + 1
        print(f"    [{page_num}/{total}] ...", end="", flush=True)

        try:
            response = model.generate_content([img, prompt])
            text = response.text.strip()
            all_text.append(text)
            print(f" OK ({len(text)}文字)")
        except Exception as e:
            print(f" エラー: {e}")
            all_text.append(f"[ページ {page_num}: OCRエラー]")

        if i < total - 1:
            time.sleep(1)

    filtered = [t for t in all_text if t and not t.startswith("[ページ")]
    combined = "\n\n".join(filtered)
    output_text.write_text(combined, encoding="utf-8")
    print(f"  → テキスト保存: {output_text} ({len(combined)}文字)")


# ============================================================
# 単体キャプチャ（1冊）
# ============================================================
def run(start_page: int = 1):
    """開いている本1冊をキャプチャ → PDF → OCR。"""
    print("Kindleウィンドウを検出中...")
    window_id, bounds = get_kindle_window_id()

    output_dir = Path(cfg.OUTPUT_DIR)
    output_dir.mkdir(exist_ok=True)

    print("=" * 50)
    print("Kindle 1冊キャプチャ → PDF → OCR")
    print(f"  ウィンドウサイズ: {bounds.get('Width', '?')}x{bounds.get('Height', '?')}")
    print(f"  開始ページ:       {start_page}")
    print("=" * 50)

    activate_kindle()
    time.sleep(cfg.INITIAL_WAIT)

    if start_page > 1:
        print(f"ページ {start_page} までスキップ中...")
        for i in range(1, start_page):
            kindle_next_page()
            time.sleep(0.3)
        print("スキップ完了。")
        time.sleep(cfg.PAGE_WAIT)

    prev_screenshot = None
    page = start_page

    while page <= cfg.MAX_PAGES:
        try:
            time.sleep(cfg.PAGE_WAIT)
            screenshot = capture_kindle_window(window_id)

            if prev_screenshot is not None and images_are_similar(screenshot, prev_screenshot):
                print(f"\n最終ページを検出しました（ページ {page - 1}）。")
                break

            save_path = output_dir / f"page_{page:04d}.png"
            screenshot.save(str(save_path))
            print(f"[ページ {page}] 保存: {save_path.name}")

            prev_screenshot = screenshot
            kindle_next_page()
            page += 1

        except pyautogui.FailSafeException:
            print(f"\n緊急停止！ 最終処理ページ: {page}")
            return
        except Exception as e:
            print(f"\nエラー発生 (ページ {page}): {e}")
            return

    total = page - start_page
    print(f"\nキャプチャ完了！ {total} ページ")

    pdf_path = Path(cfg.OUTPUT_PDF)
    png_files = pngs_to_pdf(output_dir, pdf_path)

    images = [Image.open(f) for f in png_files]
    text_path = Path(cfg.OUTPUT_TEXT)
    ocr_with_gemini(images, text_path)


# ============================================================
# OCRのみ
# ============================================================
def ocr_only(pdf_file: str = None):
    """PDFまたは既存キャプチャ画像からOCRを実行する。"""
    if pdf_file:
        pdf_path = Path(pdf_file)
        if not pdf_path.exists():
            print(f"エラー: ファイルが見つかりません: {pdf_path}")
            sys.exit(1)
        print(f"PDF読み込み中: {pdf_path}")
        images = pdf_to_images(pdf_path)
        print(f"  {len(images)} ページ検出")
    else:
        output_dir = Path(cfg.OUTPUT_DIR)
        pdf_path = Path(cfg.OUTPUT_PDF)
        png_files = pngs_to_pdf(output_dir, pdf_path)
        images = [Image.open(f) for f in png_files]

    text_path = Path(cfg.OUTPUT_TEXT)
    ocr_with_gemini(images, text_path)


# ============================================================
# エントリポイント
# ============================================================
def main():
    parser = argparse.ArgumentParser(description="Kindle自動ページめくり＆OCRシステム")
    subparsers = parser.add_subparsers(dest="command")

    subparsers.add_parser("collection", help="コレクション一覧から全書籍を一括処理")

    run_parser = subparsers.add_parser("run", help="開いている本1冊をキャプチャ→PDF→OCR")
    run_parser.add_argument(
        "--start-page", type=int, default=1,
        help="開始ページ番号（レジュームに使用）"
    )

    ocr_parser = subparsers.add_parser("ocr", help="PDFまたはキャプチャ画像からOCR実行")
    ocr_parser.add_argument(
        "pdf", nargs="?", default=None,
        help="OCR対象のPDFファイルパス"
    )

    args = parser.parse_args()

    if args.command == "collection":
        run_collection()
    elif args.command == "run":
        run(start_page=args.start_page)
    elif args.command == "ocr":
        ocr_only(pdf_file=args.pdf)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
