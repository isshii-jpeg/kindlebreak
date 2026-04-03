"""
Kindle自動ページめくり＆OCRテキスト化システム

使い方:
  1. Kindleで本を開く
  2. python3 kindlebreak.py run  → キャプチャ → PDF化 → Gemini OCR → テキスト出力
  3. python3 kindlebreak.py run --start-page 42  → 42ページ目から再開
  4. python3 kindlebreak.py ocr  → 既存のキャプチャ画像からPDF化 & OCRだけ実行
"""

import sys
import os
import time
import subprocess
import argparse
import tempfile
import json
from dataclasses import dataclass
from pathlib import Path

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

    # --- 出力 ---
    OUTPUT_DIR: str = "pages"
    OUTPUT_PDF: str = "book.pdf"
    OUTPUT_TEXT: str = "output.txt"

    # --- 終了検出 ---
    SIMILARITY_THRESHOLD: float = 0.999
    MAX_PAGES: int = 5000

    # --- Gemini OCR ---
    GEMINI_BATCH_SIZE: int = 10  # 一度にGeminiに送るページ数

    # --- 安全装置 ---
    FAILSAFE: bool = True


cfg = CONFIG()

pyautogui.FAILSAFE = cfg.FAILSAFE
pyautogui.PAUSE = 0.3


# ============================================================
# Kindleウィンドウ検出 & キャプチャ (macOS)
# ============================================================
def get_kindle_window_id():
    """KindleウィンドウのIDを取得する。"""
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
        print("  Kindleアプリを開いて本を表示してください。")
        sys.exit(1)
    if output.startswith("ERROR:"):
        print(f"エラー: {output}")
        sys.exit(1)

    info = json.loads(output)
    return info["id"], info["bounds"]


def activate_kindle():
    """Kindleウィンドウを最前面にしてフォーカスを当てる。"""
    subprocess.run(
        ["osascript", "-e", 'tell application id "com.amazon.Lassen" to activate'],
        capture_output=True, timeout=5
    )
    time.sleep(0.5)


def kindle_next_page():
    """AppleScriptでKindleに左矢印キーを送信してページをめくる。"""
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
def ocr_with_gemini(png_files: list[Path], output_text: Path):
    """キャプチャ画像をGemini APIに送ってOCRする。"""
    import google.generativeai as genai

    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        print("エラー: GEMINI_API_KEY 環境変数が設定されていません。")
        print("  export GEMINI_API_KEY='your-api-key' を実行してください。")
        sys.exit(1)

    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-2.5-flash")

    total = len(png_files)
    batch_size = cfg.GEMINI_BATCH_SIZE
    all_text = []

    print(f"\nGemini OCR開始... ({total} ページ, {batch_size}ページずつ処理)")

    for batch_start in range(0, total, batch_size):
        batch_end = min(batch_start + batch_size, total)
        batch_files = png_files[batch_start:batch_end]
        page_from = batch_start + 1
        page_to = batch_end

        print(f"  処理中: ページ {page_from}-{page_to} / {total} ...", end="", flush=True)

        # 画像をアップロード
        images = []
        for f in batch_files:
            img = Image.open(f)
            images.append(img)

        prompt_parts = []
        for img in images:
            prompt_parts.append(img)

        prompt_parts.append(
            "これらは本のページの画像です。各ページのテキストを正確に書き起こしてください。"
            "ページ間は空行で区切ってください。"
            "画像内のテキストだけを出力し、それ以外の説明は不要です。"
            "ヘッダー、フッター、ページ番号などのUI要素は無視してください。"
        )

        try:
            response = model.generate_content(prompt_parts)
            text = response.text
            all_text.append(text)
            print(f" OK ({len(text)}文字)")
        except Exception as e:
            print(f" エラー: {e}")
            all_text.append(f"[ページ {page_from}-{page_to}: OCRエラー: {e}]\n")

        # レート制限対策
        if batch_end < total:
            time.sleep(2)

    # テキスト保存
    combined = "\n\n".join(all_text)
    output_text.write_text(combined, encoding="utf-8")
    print(f"\nOCR完了！")
    print(f"テキスト保存: {output_text.resolve()}")
    print(f"合計文字数: {len(combined)}")


# ============================================================
# キャプチャ → PDF → OCR パイプライン
# ============================================================
def run(start_page: int = 1):
    """自動キャプチャ → PDF化 → Gemini OCRのフルパイプライン。"""

    # Kindleウィンドウを自動検出
    print("Kindleウィンドウを検出中...")
    window_id, bounds = get_kindle_window_id()

    output_dir = Path(cfg.OUTPUT_DIR)
    output_dir.mkdir(exist_ok=True)

    print("=" * 50)
    print("Kindle 自動キャプチャ → PDF → OCR")
    print(f"  ウィンドウID:     {window_id}")
    print(f"  ウィンドウサイズ: {bounds.get('Width', '?')}x{bounds.get('Height', '?')}")
    print(f"  開始ページ:       {start_page}")
    print(f"  画像保存先:       {output_dir.resolve()}/")
    print("=" * 50)

    # Kindleをアクティブにする
    print("\nKindleウィンドウをアクティブにしています...")
    activate_kindle()
    time.sleep(cfg.INITIAL_WAIT)

    # start_page > 1 の場合、そこまでスキップ
    if start_page > 1:
        print(f"ページ {start_page} までスキップ中...")
        for i in range(1, start_page):
            kindle_next_page()
            time.sleep(0.3)
        print("スキップ完了。")
        time.sleep(cfg.PAGE_WAIT)

    prev_screenshot = None
    page = start_page

    # ── ステップ1: キャプチャ ──
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
            print(f"再開コマンド: python3 kindlebreak.py run --start-page {page + 1}")
            return
        except Exception as e:
            print(f"\nエラー発生 (ページ {page}): {e}")
            print(f"再開コマンド: python3 kindlebreak.py run --start-page {page}")
            return

    total = page - start_page
    print(f"\nキャプチャ完了！ {total} ページ")

    # ── ステップ2: PDF化 ──
    pdf_path = Path(cfg.OUTPUT_PDF)
    png_files = pngs_to_pdf(output_dir, pdf_path)

    # ── ステップ3: Gemini OCR ──
    text_path = Path(cfg.OUTPUT_TEXT)
    ocr_with_gemini(png_files, text_path)


def ocr_only():
    """既存のキャプチャ画像からPDF化 & OCRだけ実行する。"""
    output_dir = Path(cfg.OUTPUT_DIR)

    # PDF化
    pdf_path = Path(cfg.OUTPUT_PDF)
    png_files = pngs_to_pdf(output_dir, pdf_path)

    # Gemini OCR
    text_path = Path(cfg.OUTPUT_TEXT)
    ocr_with_gemini(png_files, text_path)


# ============================================================
# エントリポイント
# ============================================================
def main():
    parser = argparse.ArgumentParser(description="Kindle自動ページめくり＆OCRシステム")
    subparsers = parser.add_subparsers(dest="command")

    run_parser = subparsers.add_parser("run", help="キャプチャ → PDF → OCR フルパイプライン")
    run_parser.add_argument(
        "--start-page", type=int, default=1,
        help="開始ページ番号（レジュームに使用）"
    )

    subparsers.add_parser("ocr", help="既存のキャプチャ画像からPDF化 & OCRだけ実行")

    args = parser.parse_args()

    if args.command == "run":
        run(start_page=args.start_page)
    elif args.command == "ocr":
        ocr_only()
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
