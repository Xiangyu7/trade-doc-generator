# -*- coding: utf-8 -*-
"""
外贸单据 RPA —— 本地网页后端
启动：  python3 app.py
然后浏览器会自动打开 http://127.0.0.1:5000
"""
import webbrowser
import threading
import uuid
from pathlib import Path

from flask import Flask, request, jsonify, send_from_directory, render_template

import core

BASE = Path(__file__).resolve().parent
UPLOADS = BASE / "uploads"
OUTDIR = BASE / "输出单据"
UPLOADS.mkdir(exist_ok=True)
OUTDIR.mkdir(exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100MB


def _save_upload(file_storage) -> Path:
    """保存上传文件到 uploads/，保留原名、加唯一前缀防冲突，支持中文名。"""
    name = Path(file_storage.filename).name.replace("/", "_").replace("\\", "_")
    dest = UPLOADS / f"{uuid.uuid4().hex[:8]}__{name}"
    file_storage.save(dest)
    return dest


@app.route("/")
def index():
    return render_template("index.html",
                           fields=core.ORDER_FIELDS,
                           chem_fields=core.CHEM_DEFAULTS,
                           chem_pkg=core.CHEM_PKG_DEFAULTS,
                           pkg_default="200kg")


@app.route("/api/extract", methods=["POST"])
def api_extract():
    """上传一个 PO，返回尽量提取到的字段（草稿）。"""
    f = request.files.get("file")
    if not f:
        return jsonify({"ok": False, "msg": "没收到文件"}), 400
    path = _save_upload(f)
    data = core.extract_po(path)
    scanned = data.pop("_scanned", False)
    ocr = data.pop("_ocr", False)
    ai = data.pop("_ai", False)
    text = data.pop("_text", "")
    fields = {k: v for k, v in data.items() if not k.startswith("_")}
    return jsonify({
        "ok": True,
        "scanned": scanned,
        "ocr": ocr,
        "ai": ai,
        "fields": fields,
        "preview": text[:1500],
        "filename": f.filename,
    })


@app.route("/api/generate", methods=["POST"])
def api_generate():
    """根据表单字段，按勾选分别生成 PI / CI / PL（每个一份独立文件）。"""
    order = request.get_json(force=True) or {}
    if not (order.get("发票号") or "").strip():
        return jsonify({"ok": False, "msg": "请先填写发票号"}), 400
    which = order.pop("which", None)  # ["PI","CI","PL"] 的子集；缺省=全部
    if which is not None:
        which = [k for k in which if k in core.INVOICE_DOCS]
        if not which:
            return jsonify({"ok": False, "msg": "请至少勾选一种单据（PI/CI/PL）"}), 400
    try:
        results = core.generate_invoice_docs(order, which, OUTDIR)
    except Exception as e:
        return jsonify({"ok": False, "msg": f"生成失败：{e}"}), 500
    return jsonify({"ok": True, "results": results})


@app.route("/api/merge", methods=["POST"])
def api_merge():
    """合并多个上传文件成一个清关 PDF。前端用 order[] 传顺序。"""
    files = request.files.getlist("files")
    if not files:
        return jsonify({"ok": False, "msg": "没收到文件"}), 400
    saved = [_save_upload(f) for f in files]
    out_name = (request.form.get("name") or "清关资料").strip()
    out_name = core.safe_name(out_name) + ".pdf"
    out_path = OUTDIR / out_name
    _, skipped = core.merge_pdfs(saved, out_path)
    return jsonify({"ok": True, "pdf": out_path.name, "skipped": skipped})


@app.route("/api/chem_extract", methods=["POST"])
def api_chem_extract():
    """上传 性能单/鉴定报告 等源文件，尽量提取危化品字段（草稿）。"""
    files = request.files.getlist("files")
    if not files:
        return jsonify({"ok": False, "msg": "没收到文件"}), 400
    saved = [_save_upload(f) for f in files]
    data = core.extract_chem(saved)
    ocr = data.pop("_ocr", False)
    texts = data.pop("_texts", [])
    fields = {k: v for k, v in data.items() if not k.startswith("_")}
    preview = "\n\n".join(f"【{n}】\n{(t or '')[:500]}" for n, t in texts)
    return jsonify({"ok": True, "ocr": ocr, "fields": fields, "preview": preview})


@app.route("/api/chem", methods=["POST"])
def api_chem():
    """根据危化品字段 + 包装类型，生成 符合性声明/责任声明/对应厂检单。"""
    data = request.get_json(force=True) or {}
    packaging = data.pop("包装类型", None) or "200kg"
    if packaging not in core.CHEM_FACTORY:
        return jsonify({"ok": False, "msg": f"未知包装类型：{packaging}"}), 400
    try:
        results = core.generate_chem_docs(data, packaging, OUTDIR)
    except Exception as e:
        return jsonify({"ok": False, "msg": f"生成失败：{e}"}), 500
    return jsonify({"ok": True, "results": results})


@app.route("/download/<path:fname>")
def download(fname):
    return send_from_directory(OUTDIR, fname, as_attachment=True)


@app.route("/preview/<path:fname>")
def preview(fname):
    return send_from_directory(OUTDIR, fname, as_attachment=False)


import os
PORT = int(os.environ.get("FLASK_PORT", "5000"))


def _open_browser():
    webbrowser.open(f"http://127.0.0.1:{PORT}")


if __name__ == "__main__":
    if not os.environ.get("NO_BROWSER"):
        threading.Timer(1.2, _open_browser).start()
    print(f"启动中… 浏览器会自动打开 http://127.0.0.1:{PORT}  （按 Ctrl+C 停止）")
    app.run(host="127.0.0.1", port=PORT, debug=False)
