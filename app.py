"""Figure Extractor for PPTX — ローカルで動く Web アプリ。

    python app.py            # ブラウザが http://127.0.0.1:8790 を開く（使用中なら次の空き）
    python app.py --port 9000 --no-browser
"""
from __future__ import annotations

import argparse
import os
import platform
import re
import shutil
import subprocess
import tempfile
import threading
import time
import uuid
import webbrowser

import fitz
from flask import Flask, abort, jsonify, request, send_from_directory

import pptx_tools
import renderer

BASE = os.path.dirname(os.path.abspath(__file__))
WORK = os.path.join(tempfile.gettempdir(), "figure-extractor-pptx")
EXTS = ("png", "tif", "bmp", "jpg")
BAD_CHARS = re.compile(r'[\\/:*?"<>|\x00-\x1f]')

app = Flask(__name__, static_folder=os.path.join(BASE, "static"), static_url_path="/static")
app.config["MAX_CONTENT_LENGTH"] = 1024 * 1024 * 1024  # 1 GB


def _session_dir(sid: str) -> str:
    if not re.fullmatch(r"[0-9a-f]{32}", sid or ""):
        abort(400, "不正なセッションです")
    d = os.path.join(WORK, sid)
    if not os.path.isdir(d):
        abort(404, "セッションが見つかりません。pptx を読み込み直してください")
    return d


def _cleanup_old(hours: float = 24) -> None:
    if not os.path.isdir(WORK):
        return
    now = time.time()
    for n in os.listdir(WORK):
        p = os.path.join(WORK, n)
        try:
            if now - os.path.getmtime(p) > hours * 3600:
                shutil.rmtree(p, ignore_errors=True)
        except OSError:
            pass


@app.get("/")
def index():
    return send_from_directory(app.static_folder, "index.html")


@app.get("/api/status")
def status():
    return jsonify({"renderer": renderer.available_renderer(), "home": os.path.expanduser("~")})


@app.post("/api/upload")
def upload():
    f = request.files.get("file")
    if f is None or not f.filename.lower().endswith((".pptx", ".pptm", ".ppsx", ".potx")):
        return jsonify({"error": "pptx ファイルを選んでください"}), 400
    _cleanup_old()
    sid = uuid.uuid4().hex
    d = os.path.join(WORK, sid)
    os.makedirs(d)
    src = os.path.join(d, "source.pptx")
    f.save(src)
    try:
        info = pptx_tools.list_shapes(src)
        prev = os.path.join(d, "preview.pptx")
        n = pptx_tools.build_preview_pptx(src, prev)
        pdf = os.path.join(d, "preview.pdf")
        kind = renderer.pptx_to_pdf(prev, pdf)
        pages = renderer.render_pages(pdf, d)
    except Exception as e:  # noqa: BLE001
        shutil.rmtree(d, ignore_errors=True)
        return jsonify({"error": f"読み込みに失敗しました: {e}"}), 500
    if len(pages) < n:
        shutil.rmtree(d, ignore_errors=True)
        return jsonify({"error": f"PDF のページ数（{len(pages)}）がスライド数（{n}）より少なく、対応が取れません"}), 500
    for s in info["slides"]:
        s["image"] = f"/api/slide/{sid}/{s['index'] + 1}"
    info.update({"session": sid, "filename": f.filename, "renderer": kind})
    return jsonify(info)


@app.get("/api/slide/<sid>/<int:n>")
def slide_image(sid, n):
    d = _session_dir(sid)
    return send_from_directory(d, f"slide{n}.png", max_age=3600)


def _validate_items(items):
    errors, seen = [], {}
    for i, it in enumerate(items, start=1):
        name = (it.get("name") or "").strip()
        ext = (it.get("ext") or "").lower().lstrip(".")
        if ext == "tiff":
            ext = "tif"
        if ext == "jpeg":
            ext = "jpg"
        if not name:
            errors.append(f"{i}番目: 名前が空です")
        elif BAD_CHARS.search(name) or name in (".", ".."):
            errors.append(f"{i}番目「{name}」: 使えない文字（\\ / : * ? \" < > |）が入っています")
        shapes = it.get("shapes")
        if not isinstance(shapes, list) or not shapes:
            errors.append(f"{i}番目: オブジェクトがありません")
        if ext not in EXTS:
            errors.append(f"{i}番目: 拡張子 {ext} には対応していません")
        key = f"{name}.{ext}".lower()
        if key in seen:
            errors.append(f"{seen[key]}番目と{i}番目の名前が同じです（{name}.{ext}）")
        seen.setdefault(key, i)
        it["name"], it["ext"] = name, ext
    return errors


@app.post("/api/export")
def export():
    body = request.get_json(force=True)
    d = _session_dir(body.get("session"))
    items = body.get("items") or []
    if not items:
        return jsonify({"error": "オブジェクトが選ばれていません"}), 400
    folder = os.path.abspath(os.path.expanduser((body.get("folder") or "").strip()))
    if not (body.get("folder") or "").strip():
        return jsonify({"error": "保存先のフォルダを指定してください"}), 400
    dpi = max(50, min(1200, int(body.get("dpi") or 300)))
    transparent = bool(body.get("transparent"))
    overwrite = bool(body.get("overwrite"))

    errors = _validate_items(items)
    if errors:
        return jsonify({"error": "\n".join(errors)}), 400

    targets = [os.path.join(folder, f"{it['name']}.{it['ext']}") for it in items]
    if os.path.exists(folder) and not os.path.isdir(folder):
        return jsonify({"error": f"{folder} はフォルダではありません"}), 400
    existing = [os.path.basename(t) for t in targets if os.path.exists(t)]
    if existing and not overwrite:
        return jsonify({"conflict": existing}), 409

    src = os.path.join(d, "source.pptx")
    job_dir = tempfile.mkdtemp(dir=d)
    try:
        exp = os.path.join(job_dir, "export.pptx")
        pptx_tools.build_export_pptx(src, exp, [(int(it["slide"]), [int(x) for x in it["shapes"]]) for it in items])
        pdf = os.path.join(job_dir, "export.pdf")
        renderer.pptx_to_pdf(exp, pdf)
        doc = fitz.open(pdf)
        n = len(items)
        if doc.page_count < 2 * n:
            return jsonify({"error": f"PDF のページ数（{doc.page_count}）が足りません"}), 500
        os.makedirs(folder, exist_ok=True)
        results = []
        for i, (it, path) in enumerate(zip(items, targets)):
            try:
                on_white, rgba = renderer.crop_object(doc[i], doc[n + i], dpi)
                renderer.save_image(on_white, rgba, path, it["ext"], dpi, transparent)
                results.append({"file": os.path.basename(path), "ok": True,
                                "size": [on_white.width, on_white.height]})
            except Exception as e:  # noqa: BLE001
                results.append({"file": os.path.basename(path), "ok": False, "error": str(e)})
        doc.close()
    except Exception as e:  # noqa: BLE001
        return jsonify({"error": f"書き出しに失敗しました: {e}"}), 500
    finally:
        shutil.rmtree(job_dir, ignore_errors=True)
    return jsonify({"folder": folder, "results": results})


@app.post("/api/choose-folder")
def choose_folder():
    start = (request.get_json(silent=True) or {}).get("start") or ""
    start = os.path.expanduser(start)
    as_start = start.replace("\\", "\\\\").replace('"', '\\"')
    ps_start = start.replace("'", "''")
    system = platform.system()
    try:
        if system == "Darwin":
            script = ['tell application "System Events"', "activate"]
            if os.path.isdir(start):
                script.append(f'set f to choose folder with prompt "保存先のフォルダ" default location (POSIX file "{as_start}")')
            else:
                script.append('set f to choose folder with prompt "保存先のフォルダ"')
            script += ["end tell", "POSIX path of f"]
            args = ["osascript"]
            for line in script:
                args += ["-e", line]
            r = subprocess.run(args, capture_output=True, text=True)
            path = r.stdout.strip()
        elif system == "Windows":
            ps = ("Add-Type -AssemblyName System.Windows.Forms;"
                  "$d=New-Object System.Windows.Forms.FolderBrowserDialog;"
                  "$d.Description='保存先のフォルダ';"
                  + (f"$d.SelectedPath='{ps_start}';" if os.path.isdir(start) else "")
                  + "if($d.ShowDialog() -eq 'OK'){[Console]::OutputEncoding=[Text.Encoding]::UTF8;$d.SelectedPath}")
            r = subprocess.run(["powershell", "-NoProfile", "-STA", "-Command", ps], capture_output=True, text=True,
                               encoding="utf-8")
            path = r.stdout.strip()
        else:
            import tkinter
            from tkinter import filedialog
            root = tkinter.Tk()
            root.withdraw()
            path = filedialog.askdirectory(initialdir=start or None) or ""
            root.destroy()
    except Exception as e:  # noqa: BLE001
        return jsonify({"error": str(e)}), 500
    return jsonify({"folder": path.rstrip("/") if len(path) > 1 else path})


@app.post("/api/open-folder")
def open_folder():
    folder = os.path.expanduser((request.get_json(silent=True) or {}).get("folder") or "")
    if not os.path.isdir(folder):
        return jsonify({"error": "フォルダがありません"}), 400
    system = platform.system()
    if system == "Darwin":
        subprocess.Popen(["open", folder])
    elif system == "Windows":
        os.startfile(folder)  # type: ignore[attr-defined]
    else:
        subprocess.Popen(["xdg-open", folder])
    return jsonify({"ok": True})


def _free_port(start: int, tries: int = 50) -> int:
    """start から順に、空いているポートを探す（ほかのアプリが使っているポートは避ける）。"""
    import socket
    for port in range(start, start + tries):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    raise SystemExit(f"{start}〜{start + tries - 1} のポートがすべて使われています。--port で指定してください")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--port", type=int, default=8790, help="使うポート（使用中なら次の空きを探す）")
    ap.add_argument("--no-browser", action="store_true")
    a = ap.parse_args()
    os.makedirs(WORK, exist_ok=True)
    port = _free_port(a.port)
    if port != a.port:
        print(f"ポート {a.port} は別のアプリが使っているため、{port} を使います")
    url = f"http://127.0.0.1:{port}/"
    print(f"Figure Extractor: {url}  （描画: {renderer.available_renderer() or '見つかりません'}）")
    if not a.no_browser:
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    app.run(host="127.0.0.1", port=port, threaded=True)


if __name__ == "__main__":
    main()
