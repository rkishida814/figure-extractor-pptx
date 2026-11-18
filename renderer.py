"""pptx → PDF の変換と、PDF からの切り出し。

変換は次の順に試す（環境変数 FX_RENDERER = auto / powerpoint / libreoffice で固定できる）。
- Mac: Microsoft PowerPoint（AppleScript）
- Windows: Microsoft PowerPoint（COM、pywin32 が必要）
- LibreOffice（soffice）
"""
from __future__ import annotations

import os
import platform
import shutil
import subprocess
import tempfile
import threading
import uuid

import fitz  # PyMuPDF
import numpy as np
from PIL import Image

_lock = threading.Lock()
SYSTEM = platform.system()

MAC_PPT_APP = "/Applications/Microsoft PowerPoint.app"
MAC_PPT_CONTAINER = os.path.expanduser("~/Library/Containers/com.microsoft.Powerpoint/Data/tmp/figure-extractor")

_MAC_SCRIPT = r'''
on run argv
    set inPath to item 1 of argv
    set outPath to item 2 of argv
    set docName to item 3 of argv
    tell application "Microsoft PowerPoint"
        open (POSIX file inPath)
        set thePres to missing value
        repeat 50 times
            try
                set thePres to presentation docName
                exit repeat
            end try
            delay 0.2
        end repeat
        if thePres is missing value then set thePres to active presentation
        save thePres in (POSIX file outPath) as save as PDF
        close thePres saving no
    end tell
end run
'''


def _find_soffice() -> str | None:
    for c in ("soffice", "libreoffice"):
        p = shutil.which(c)
        if p:
            return p
    for p in ("/Applications/LibreOffice.app/Contents/MacOS/soffice",
              r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
        if os.path.exists(p):
            return p
    return None


def _has_win_powerpoint() -> bool:
    if SYSTEM != "Windows":
        return False
    try:
        import win32com.client  # noqa: F401
        return True
    except Exception:
        return False


def available_renderer() -> str | None:
    want = os.environ.get("FX_RENDERER", "auto").lower()
    has_ppt = (SYSTEM == "Darwin" and os.path.exists(MAC_PPT_APP)) or _has_win_powerpoint()
    soffice = _find_soffice()
    if want == "powerpoint":
        return "powerpoint" if has_ppt else None
    if want == "libreoffice":
        return "libreoffice" if soffice else None
    if has_ppt:
        return "powerpoint"
    if soffice:
        return "libreoffice"
    return None


def _mac_powerpoint(pptx: str, pdf: str) -> None:
    os.makedirs(MAC_PPT_CONTAINER, exist_ok=True)
    name = f"fx_{uuid.uuid4().hex[:10]}"
    in_path = os.path.join(MAC_PPT_CONTAINER, name + ".pptx")
    out_path = os.path.join(MAC_PPT_CONTAINER, name + ".pdf")
    shutil.copyfile(pptx, in_path)
    try:
        r = subprocess.run(["osascript", "-e", _MAC_SCRIPT, in_path, out_path, name + ".pptx"],
                           capture_output=True, text=True, timeout=600)
        if r.returncode != 0 or not os.path.exists(out_path):
            raise RuntimeError("PowerPoint での PDF 変換に失敗しました: " + (r.stderr or r.stdout).strip())
        shutil.move(out_path, pdf)
    finally:
        for p in (in_path, out_path):
            try:
                os.remove(p)
            except OSError:
                pass


def _win_powerpoint(pptx: str, pdf: str) -> None:
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
        app = win32com.client.DispatchEx("PowerPoint.Application")
        pres = app.Presentations.Open(os.path.abspath(pptx), True, False, False)  # ReadOnly, Untitled, WithWindow
        try:
            pres.SaveAs(os.path.abspath(pdf), 32)  # ppSaveAsPDF
        finally:
            pres.Close()
    finally:
        pythoncom.CoUninitialize()


def _libreoffice(pptx: str, pdf: str) -> None:
    soffice = _find_soffice()
    with tempfile.TemporaryDirectory() as outdir, tempfile.TemporaryDirectory() as profile:
        r = subprocess.run([soffice, f"-env:UserInstallation=file://{profile}", "--headless",
                            "--convert-to", "pdf", "--outdir", outdir, pptx],
                           capture_output=True, text=True, timeout=600)
        out = os.path.join(outdir, os.path.splitext(os.path.basename(pptx))[0] + ".pdf")
        if not os.path.exists(out):
            raise RuntimeError("LibreOffice での PDF 変換に失敗しました: " + (r.stderr or r.stdout).strip())
        shutil.move(out, pdf)


def pptx_to_pdf(pptx: str, pdf: str) -> str:
    kind = available_renderer()
    if kind is None:
        raise RuntimeError("PowerPoint も LibreOffice も見つかりません。どちらかを入れてください。")
    with _lock:
        if kind == "powerpoint":
            (_mac_powerpoint if SYSTEM == "Darwin" else _win_powerpoint)(pptx, pdf)
        else:
            _libreoffice(pptx, pdf)
    return kind


# ---------------------------------------------------------------- PDF → 画像

def render_pages(pdf: str, out_dir: str, width_px: int = 1600) -> list[str]:
    doc = fitz.open(pdf)
    paths = []
    for i, page in enumerate(doc):
        zoom = width_px / page.rect.width
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
        p = os.path.join(out_dir, f"slide{i + 1}.png")
        pix.save(p)
        paths.append(p)
    return paths


def _render(page, zoom: float, bg: int, clip=None) -> np.ndarray:
    """ページを描き、何も描かれていない部分を bg（0 か 255）で埋める。
    ページの端で背景の矩形がわずかに足りないとき、白ページと黒ページで縁の色が食い違わないようにする。"""
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=clip, alpha=True)
    arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n).astype(np.float32)
    a = arr[:, :, 3:4] / 255.0
    rgb = arr[:, :, :3] + bg * (1.0 - a)  # samples は乗算済み（premultiplied）
    return np.clip(rgb.round(), 0, 255).astype(np.uint8)


def _alpha(white: np.ndarray, black: np.ndarray) -> np.ndarray:
    """白背景と黒背景の2枚から不透明度（0..1）を求める。"""
    diff = white.astype(np.float32) - black.astype(np.float32)
    return np.clip(1.0 - diff.mean(axis=2) / 255.0, 0.0, 1.0)


def _bbox(alpha: np.ndarray, thr: float = 0.02):
    ys, xs = np.nonzero(alpha > thr)
    if len(xs) == 0:
        return None
    return int(xs.min()), int(ys.min()), int(xs.max()) + 1, int(ys.max()) + 1


def _safe_rect(*pages):
    """PowerPoint は、ページとわずかに大きさの違う背景の四角を重ねて描く（白の上に背景色など）。
    はみ出した端の細い帯を図と誤認しないよう、ページ全体を覆う四角すべてが重なる内側だけを調べる。"""
    r = fitz.Rect(pages[0].rect)
    for page in pages:
        area = page.rect.width * page.rect.height
        for dr in page.get_drawings():
            rect = dr.get("rect")
            if dr.get("fill") is not None and rect is not None and rect.width * rect.height >= 0.98 * area:
                r &= rect
    return r


def crop_object(page_white, page_black, dpi: int) -> tuple[Image.Image, Image.Image]:
    """白背景・黒背景のページから、オブジェクトを余白なしで切り出す。
    戻り値は (白背景の RGB 画像, 背景透明の RGBA 画像)。"""
    # 低解像度で範囲を見つけ、その周りだけを目的の解像度で描く
    lo_zoom = 100 / 72
    safe = _safe_rect(page_white, page_black)
    safe = fitz.Rect(safe.x0 + 0.5, safe.y0 + 0.5, safe.x1 - 0.5, safe.y1 - 0.5)
    bbox = _bbox(_alpha(_render(page_white, lo_zoom, 255, safe), _render(page_black, lo_zoom, 0, safe)))
    if bbox is None:
        raise RuntimeError("オブジェクトが描かれていません（スライドの枠外・非表示・透明の可能性）")
    pad = 3  # px（低解像度）
    x0, y0, x1, y1 = bbox
    clip = fitz.Rect(safe.x0 + (x0 - pad) / lo_zoom, safe.y0 + (y0 - pad) / lo_zoom,
                     safe.x0 + (x1 + pad) / lo_zoom, safe.y0 + (y1 + pad) / lo_zoom) & safe
    zoom = dpi / 72
    w = _render(page_white, zoom, 255, clip)
    b = _render(page_black, zoom, 0, clip)
    h = min(w.shape[0], b.shape[0])
    wd = min(w.shape[1], b.shape[1])
    w, b = w[:h, :wd], b[:h, :wd]
    a = _alpha(w, b)
    bb = _bbox(a)
    if bb is not None:
        x0, y0, x1, y1 = bb
        w, b, a = w[y0:y1, x0:x1], b[y0:y1, x0:x1], a[y0:y1, x0:x1]
    safe = np.where(a > 1e-3, a, 1.0)[:, :, None]
    color = np.clip(b.astype(np.float32) / safe, 0, 255)
    rgba = np.dstack([color, a * 255.0]).round().astype(np.uint8)
    return Image.fromarray(np.ascontiguousarray(w)), Image.fromarray(rgba, "RGBA")


def save_image(on_white: Image.Image, rgba: Image.Image, path: str, ext: str, dpi: int,
               transparent: bool) -> None:
    ext = ext.lower()
    out = rgba if (transparent and ext in ("png", "tif")) else on_white
    if ext == "png":
        out.save(path, "PNG", dpi=(dpi, dpi))
    elif ext == "tif":
        out.save(path, "TIFF", dpi=(dpi, dpi), compression="tiff_lzw")
    elif ext == "jpg":
        out.save(path, "JPEG", dpi=(dpi, dpi), quality=95, subsampling=0)
    elif ext == "bmp":
        out.save(path, "BMP", dpi=(dpi, dpi))
    else:
        raise ValueError(ext)
