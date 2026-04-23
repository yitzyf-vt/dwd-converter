"""
DWD → DOCX converter  –  Flask web server
Run:  python server.py
Then visit  http://localhost:5000

Nikud and trup are always included (on by default, no user toggle).
"""
import os, uuid, threading, time
from pathlib import Path
from flask import Flask, request, send_file, jsonify

import dwd_to_docx

app = Flask(__name__)

UPLOAD_DIR = Path("/tmp/dwd_uploads")
OUTPUT_DIR = Path("/tmp/dwd_outputs")
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

MAX_FILE_MB = 20


# ── Background cleanup (delete files older than 30 min) ──────────────────────
def _cleanup():
    while True:
        time.sleep(300)
        cutoff = time.time() - 1800
        for d in (UPLOAD_DIR, OUTPUT_DIR):
            for f in d.iterdir():
                try:
                    if f.stat().st_mtime < cutoff:
                        f.unlink()
                except Exception:
                    pass

threading.Thread(target=_cleanup, daemon=True).start()


# ── Routes ────────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return open(Path(__file__).parent / "index.html").read()


@app.route("/convert", methods=["POST"])
def convert():
    f = request.files.get("file")
    if not f:
        return jsonify(error="No file uploaded"), 400
    if not f.filename.lower().endswith(".dwd"):
        return jsonify(error="Only .dwd files are supported"), 400

    data = f.read()
    if len(data) > MAX_FILE_MB * 1024 * 1024:
        return jsonify(error=f"File too large (max {MAX_FILE_MB} MB)"), 400

    job_id  = uuid.uuid4().hex
    in_path = UPLOAD_DIR / f"{job_id}.dwd"
    out_stem = Path(f.filename).stem
    out_name = f"{out_stem}.docx"
    out_path = OUTPUT_DIR / f"{job_id}.docx"

    in_path.write_bytes(data)
    try:
        # Nikud and trup always on; converter auto-detects images → returns .zip if any
        result = dwd_to_docx.convert(str(in_path), str(out_path),
                                     with_nikud=True, with_trup=True)
    except Exception as e:
        return jsonify(error=f"Conversion failed: {e}"), 500
    finally:
        in_path.unlink(missing_ok=True)

    result_path = Path(result)
    if not result_path.exists():
        return jsonify(error="Conversion produced no output"), 500

    # Rename output so download handler finds it by job_id
    actual_ext = result_path.suffix   # .docx or .zip
    final_path = OUTPUT_DIR / f"{job_id}{actual_ext}"
    result_path.rename(final_path)

    out_name = Path(f.filename).stem + actual_ext
    return jsonify(job_id=job_id, filename=out_name, has_images=(actual_ext == '.zip'))


@app.route("/download/<job_id>")
def download(job_id):
    if not job_id.replace("-", "").isalnum():
        return "Invalid ID", 400
    # Support both .docx and .zip output
    filename = request.args.get("name", "converted.docx")
    ext = ".zip" if filename.endswith(".zip") else ".docx"
    out_path = OUTPUT_DIR / f"{job_id}{ext}"
    if not out_path.exists():
        return "File not found or expired", 404
    mime = ("application/zip" if ext == ".zip"
            else "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    return send_file(out_path, as_attachment=True, download_name=filename, mimetype=mime)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
