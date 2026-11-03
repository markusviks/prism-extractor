"""
Prism Extractor — a small local web app around the `prism-ole-handler` package.

Drop PowerPoint (or Word / Excel) files on the page and get back the GraphPad
Prism graphs embedded in them, ready to open and edit in Prism.

Opens in its own window (pywebview). The page is served from 127.0.0.1 only.
    python app.py              # app window
    python app.py --browser    # serve on port 8765 and open in the web browser instead
"""

import atexit
import contextlib
import io
import os
import re
import shutil
import struct
import subprocess
import sys
import tempfile
import threading
import uuid
import webbrowser
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

from flask import Flask, abort, jsonify, request, send_file

from prism_ole_handler import PrismExtractor

HOST = "127.0.0.1"
PORT = int(os.environ.get("PRISM_EXTRACTOR_PORT", "8765"))

SUPPORTED = {
    ".pptx": "ppt", ".pptm": "ppt", ".ppsx": "ppt", ".potx": "ppt",
    ".docx": "word", ".docm": "word",
    ".xlsx": "xl", ".xlsm": "xl",
}
OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}
SAVE_ROOT = Path.home() / "Downloads" / "Prism Extracts"
# Bundled files live in sys._MEIPASS when frozen by PyInstaller.
RESOURCES = Path(getattr(sys, "_MEIPASS", Path(__file__).parent))
WINDOW_MODE = False  # set at startup; hides the "Quit app" link, since closing the window quits

WORK_ROOT = Path(tempfile.mkdtemp(prefix="prism-extractor-"))
atexit.register(shutil.rmtree, WORK_ROOT, ignore_errors=True)

# PrismExtractor prints progress to stdout; we capture it per job, and
# redirect_stdout is process-wide, so jobs run one at a time.
EXTRACT_LOCK = threading.Lock()
JOBS = {}  # job_id -> {"dir": Path, "files": {name: meta}}

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 1024 * 1024 * 1024  # 1 GB of uploads per request


# --------------------------------------------------------------------------- #
# Extraction
# --------------------------------------------------------------------------- #

def unwrap_length_prefix(data):
    """Prism 4-9 objects store a 4-byte little-endian length, then a binary .pzf file.
    Unlike Prism 10's zip payloads, that length counts the 4 prefix bytes themselves
    (checked on 21 Prism 8 objects: only this reading ends every file on its 0x0140
    trailer). The OLE stream is often longer still, with leftovers from earlier saves."""
    if data[4:8] == b"PCFF":
        n = struct.unpack("<I", data[:4])[0]
        if 8 < n <= len(data):
            return data[4:n]
    return data


def classify_payload(data, extractor):
    """Return 'prism' (Prism 10+ zip), 'pzfx' (Prism XML), 'pzf'/'pzm' (binary) or None."""
    if data[:8] == b"PCFFGRA4":  # Prism 4 and later (GraphPad KB #1846)
        return "pzf"
    if data[:8] == b"PCFFGRAF":  # Prism 3 and earlier
        return "pzm"
    if data[:2] == b"PK":
        try:
            names = zipfile.ZipFile(io.BytesIO(data)).namelist()
        except zipfile.BadZipFile:
            return None
        if "document.json" in names or any(n.startswith(("data/sets/", "graphs/")) for n in names):
            return "prism"
        return None
    if extractor.is_prism_xml(data):
        return "pzfx"
    return None


def pptx_ole_details(unzipped):
    """Map 'oleObjectN.bin' -> {slide, progId, order} by reading slide XML."""
    details = {}
    slides_dir = unzipped / "ppt" / "slides"
    for slide_xml in slides_dir.glob("slide*.xml"):
        m = re.fullmatch(r"slide(\d+)\.xml", slide_xml.name)
        if not m:
            continue
        slide_num = int(m.group(1))
        rels_file = slides_dir / "_rels" / f"{slide_xml.name}.rels"
        if not rels_file.exists():
            continue
        rid_to_target = {
            r.get("Id"): Path(r.get("Target", "")).name
            for r in ET.parse(rels_file).getroot().findall("rel:Relationship", NS)
        }
        seen = set()
        for order, ole in enumerate(ET.parse(slide_xml).getroot().iter(f"{{{NS['p']}}}oleObj")):
            rid = ole.get(f"{{{NS['r']}}}id")
            # Each object usually appears twice (mc:Choice + mc:Fallback).
            if rid in seen or rid not in rid_to_target:
                continue
            seen.add(rid)
            details[rid_to_target[rid]] = {
                "slide": slide_num,
                "progId": ole.get("progId", ""),
                "order": order,
            }
    return details


def extract_document(src_path, kind, job_dir):
    """Extract Prism objects from one Office file. Returns a result dict."""
    stem = re.sub(r"[^\w\- .]+", "_", src_path.stem).strip() or "document"
    extractor = PrismExtractor(src_path)
    extractor.temp_dir = job_dir / f"unzipped-{uuid.uuid4().hex[:8]}"
    log = io.StringIO()
    objects, skipped = [], []

    try:
        with contextlib.redirect_stdout(log):
            extractor.extract_pptx()  # plain unzip; works for any OOXML file
    except zipfile.BadZipFile:
        return {"name": src_path.name, "error": "Not a valid Office Open XML file "
                "(old binary .ppt/.doc/.xls files aren't supported — re-save as .pptx/.docx/.xlsx).",
                "objects": [], "skipped": []}

    try:
        unzipped = extractor.temp_dir
        if kind == "ppt":
            with contextlib.redirect_stdout(log):
                embedded, _ = extractor.find_embedded_objects()
            info = pptx_ole_details(unzipped)
            for obj in embedded:
                d = info.get(obj["path"].name, {})
                obj.setdefault("progId", d.get("progId", ""))
                obj["slide"] = obj["slide"] or d.get("slide")
                obj["order"] = d.get("order", 0)
        else:
            emb_dir = unzipped / kind / "embeddings"
            embedded = [
                {"path": p, "slide": None, "progId": "", "order": 0}
                for p in sorted(emb_dir.glob("*")) if p.is_file()
            ] if emb_dir.exists() else []

        embedded.sort(key=lambda o: (o["slide"] or 0, o["order"], o["path"].name))
        per_slide = {}

        for obj in embedded:
            data = obj["path"].read_bytes()
            candidates = []
            if data[:8] == OLE_MAGIC:
                with contextlib.redirect_stdout(log):
                    candidates = [c for _, c in extractor.extract_from_ole(data, "x")]
            else:
                candidates = [data]

            payload = fmt = None
            for c in candidates:
                c = unwrap_length_prefix(c)
                fmt = classify_payload(c, extractor)
                if fmt:
                    payload = c
                    break

            if payload is None:
                skipped.append({
                    "source": obj["path"].name,
                    "slide": obj["slide"],
                    "progId": obj.get("progId", ""),
                    "reason": "not a Prism object" if "prism" not in obj.get("progId", "").lower()
                    else "Prism object, but its data format wasn't recognised (older Prism version?)",
                })
                continue

            key = obj["slide"] or 0
            per_slide[key] = per_slide.get(key, 0) + 1
            if obj["slide"]:
                out_name = f"{stem}_slide{obj['slide']:02d}_graph{per_slide[key]}.{fmt}"
            else:
                out_name = f"{stem}_graph{per_slide[key]:02d}.{fmt}"
            n = 2
            while (job_dir / out_name).exists():  # two uploads with the same name
                out_name = re.sub(r"( \(\d+\))?\.(\w+)$", rf" ({n}).\2", out_name)
                n += 1
            (job_dir / out_name).write_bytes(payload)
            objects.append({
                "file": out_name,
                "slide": obj["slide"],
                "progId": obj.get("progId", ""),
                "format": fmt,
                "size": len(payload),
                "source": obj["path"].name,
            })
    finally:
        shutil.rmtree(extractor.temp_dir, ignore_errors=True)

    return {"name": src_path.name, "objects": objects, "skipped": skipped, "log": log.getvalue()}


# --------------------------------------------------------------------------- #
# Routes
# --------------------------------------------------------------------------- #

def get_job(job_id):
    job = JOBS.get(job_id)
    if not job:
        abort(404)
    return job


def get_job_file(job_id, name):
    job = get_job(job_id)
    if name not in job["files"]:
        abort(404)
    return job["dir"] / name


@app.get("/")
def index():
    return INDEX_HTML.replace("<body>", '<body class="windowed">') if WINDOW_MODE else INDEX_HTML


@app.post("/api/extract")
def api_extract():
    uploads = request.files.getlist("files")
    if not uploads:
        return jsonify(error="No files uploaded."), 400

    job_id = uuid.uuid4().hex
    job_dir = WORK_ROOT / job_id
    (job_dir / "_in").mkdir(parents=True)
    results = []

    with EXTRACT_LOCK:
        for up in uploads:
            name = Path(up.filename or "upload").name
            kind = SUPPORTED.get(Path(name).suffix.lower())
            if not kind:
                results.append({"name": name, "objects": [], "skipped": [],
                                "error": "Unsupported file type. Use .pptx, .docx or .xlsx."})
                continue
            src = job_dir / "_in" / uuid.uuid4().hex[:8] / name
            src.parent.mkdir()
            up.save(src)
            try:
                results.append(extract_document(src, kind, job_dir))
            except Exception as e:  # keep going with the other files
                results.append({"name": name, "objects": [], "skipped": [], "error": f"Extraction failed: {e}"})

    shutil.rmtree(job_dir / "_in", ignore_errors=True)
    files = {o["file"]: o for r in results for o in r["objects"]}
    JOBS[job_id] = {"dir": job_dir, "files": files}
    return jsonify(job=job_id, results=results, total=len(files))


@app.get("/api/jobs/<job_id>/files/<path:name>")
def api_download(job_id, name):
    return send_file(get_job_file(job_id, name), as_attachment=True, download_name=name)


@app.get("/api/jobs/<job_id>/all.zip")
def api_download_all(job_id):
    job = get_job(job_id)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_STORED) as zf:
        for name in job["files"]:
            zf.write(job["dir"] / name, name)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="prism_graphs.zip", mimetype="application/zip")


def save_to_disk(job_id, names):
    job = get_job(job_id)
    SAVE_ROOT.mkdir(parents=True, exist_ok=True)
    saved = []
    for name in names:
        src = get_job_file(job_id, name)
        dest = SAVE_ROOT / name
        i = 1
        # Reuse an identical earlier copy; never overwrite one that was edited in Prism.
        while dest.exists() and dest.read_bytes() != src.read_bytes():
            dest = SAVE_ROOT / f"{src.stem} ({i}){src.suffix}"
            i += 1
        if not dest.exists():
            shutil.copy2(src, dest)
        saved.append(dest)
    return saved


@app.post("/api/jobs/<job_id>/open/<path:name>")
def api_open(job_id, name):
    """Save the graph to ~/Downloads/Prism Extracts and open it in its default app (Prism)."""
    dest = save_to_disk(job_id, [name])[0]
    if sys.platform == "darwin":
        subprocess.Popen(["open", str(dest)])
    return jsonify(path=str(dest))


@app.post("/api/jobs/<job_id>/save-all")
def api_save_all(job_id):
    saved = save_to_disk(job_id, list(get_job(job_id)["files"]))
    if sys.platform == "darwin" and saved:
        subprocess.Popen(["open", "-R", str(saved[0])])
    return jsonify(folder=str(SAVE_ROOT), count=len(saved))


@app.post("/api/quit")
def api_quit():
    threading.Timer(0.3, lambda: os._exit(0)).start()
    return jsonify(ok=True)


@app.get("/favicon.png")
def favicon():
    icon = RESOURCES / "icon" / "AppIcon.png"
    if not icon.exists():
        abort(404)
    return send_file(icon, mimetype="image/png")


@app.get("/api/ping")
def api_ping():
    return jsonify(app="prism-extractor")


# --------------------------------------------------------------------------- #
# Page
# --------------------------------------------------------------------------- #

INDEX_HTML = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Prism Extractor</title>
<link rel="icon" type="image/png" href="/favicon.png">
<style>
  :root {
    --bg: #f6f7f9; --panel: #ffffff; --text: #1d2330; --muted: #677085;
    --border: #e1e4ea; --accent: #2f6fde; --accent-soft: #e8f0fd;
    --ok: #1f8a4c; --warn: #a3620a; --err: #c0392b;
  }
  @media (prefers-color-scheme: dark) {
    :root {
      --bg: #15181e; --panel: #1e222a; --text: #e6e9ef; --muted: #98a0b0;
      --border: #2f3541; --accent: #6b9cff; --accent-soft: #23304a;
      --ok: #4cc283; --warn: #e0a54a; --err: #ff7a6b;
    }
  }
  * { box-sizing: border-box; }
  body { margin: 0; background: var(--bg); color: var(--text);
         font: 15px/1.5 -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; }
  main { max-width: 820px; margin: 0 auto; padding: 32px 16px 64px; }
  header { display: flex; align-items: baseline; justify-content: space-between; gap: 12px; }
  h1 { font-size: 22px; margin: 0; }
  .sub { color: var(--muted); margin: 4px 0 24px; }
  button, .btn { font: inherit; border: 1px solid var(--border); background: var(--panel); color: var(--text);
         border-radius: 8px; padding: 6px 12px; cursor: pointer; text-decoration: none; display: inline-block; }
  button:hover, .btn:hover { border-color: var(--accent); }
  .primary { background: var(--accent); border-color: var(--accent); color: #fff; }
  .link { border: none; background: none; color: var(--muted); padding: 0; font-size: 13px; }
  .windowed #quit { display: none; }
  .windowed main { padding-top: 20px; }
  #drop { border: 2px dashed var(--border); border-radius: 14px; background: var(--panel);
          padding: 48px 16px; text-align: center; cursor: pointer; transition: .15s; }
  #drop.over { border-color: var(--accent); background: var(--accent-soft); }
  #drop strong { font-size: 17px; }
  #drop p { color: var(--muted); margin: 6px 0 0; }
  #status { margin: 20px 0 8px; color: var(--muted); min-height: 1.5em; }
  .toolbar { display: flex; gap: 8px; flex-wrap: wrap; margin-bottom: 16px; }
  .doc { background: var(--panel); border: 1px solid var(--border); border-radius: 12px; margin-bottom: 16px; overflow: hidden; }
  .doc h2 { font-size: 15px; margin: 0; padding: 12px 16px; border-bottom: 1px solid var(--border);
            display: flex; justify-content: space-between; gap: 8px; }
  .doc h2 span { color: var(--muted); font-weight: normal; white-space: nowrap; }
  .row { display: flex; align-items: center; gap: 12px; padding: 10px 16px; border-bottom: 1px solid var(--border); }
  .row:last-child { border-bottom: none; }
  .slide { flex: 0 0 64px; font-weight: 600; }
  .meta { flex: 1; min-width: 0; }
  .meta .fn { overflow-wrap: anywhere; }
  .meta small { color: var(--muted); }
  .badge { font-size: 12px; border-radius: 6px; padding: 1px 6px; background: var(--accent-soft); color: var(--accent); margin-left: 6px; }
  .actions { display: flex; gap: 6px; flex-shrink: 0; }
  .note { padding: 10px 16px; color: var(--muted); font-size: 13px; }
  .err { color: var(--err); }
  .warn { color: var(--warn); }
  details { padding: 0 16px 10px; font-size: 13px; color: var(--muted); }
  pre { white-space: pre-wrap; font-size: 12px; }
  .tip { margin-top: 28px; font-size: 13px; color: var(--muted); }
  @media (max-width: 560px) { .row { flex-wrap: wrap; } .actions { width: 100%; } }
</style>
</head>
<body>
<main>
  <header>
    <h1>Prism Extractor</h1>
    <button class="link" id="quit" title="Stop the local server">Quit app</button>
  </header>
  <p class="sub">Pull embedded GraphPad Prism graphs out of PowerPoint decks so you can edit them in Prism on your Mac.</p>

  <div id="drop" tabindex="0">
    <strong>Drop .pptx files here</strong>
    <p>or click to choose — several files at once is fine. Word and Excel files work too (experimental).</p>
    <input id="picker" type="file" multiple accept=".pptx,.pptm,.ppsx,.potx,.docx,.docm,.xlsx,.xlsm" hidden>
  </div>

  <div id="status"></div>
  <div id="results"></div>

  <p class="tip">Files stay on this computer — the app runs locally and nothing is uploaded anywhere.
  “Open in Prism” saves a copy to <code>~/Downloads/Prism Extracts</code> and opens it.</p>
</main>

<script>
const $ = s => document.querySelector(s);
const drop = $("#drop"), picker = $("#picker"), statusEl = $("#status"), resultsEl = $("#results");
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const kb = n => n > 1048576 ? (n/1048576).toFixed(1) + " MB" : Math.max(1, Math.round(n/1024)) + " KB";

drop.addEventListener("click", () => picker.click());
drop.addEventListener("keydown", e => { if (e.key === "Enter" || e.key === " ") picker.click(); });
picker.addEventListener("change", () => { if (picker.files.length) upload(picker.files); picker.value = ""; });
["dragenter","dragover"].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.add("over"); }));
["dragleave","drop"].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.remove("over"); }));
drop.addEventListener("drop", e => { if (e.dataTransfer.files.length) upload(e.dataTransfer.files); });
// Don't let a drop outside the zone navigate the page away.
window.addEventListener("dragover", e => e.preventDefault());
window.addEventListener("drop", e => e.preventDefault());

async function upload(files) {
  const fd = new FormData();
  for (const f of files) fd.append("files", f);
  statusEl.textContent = `Extracting from ${files.length} file${files.length > 1 ? "s" : ""}…`;
  resultsEl.innerHTML = "";
  try {
    const res = await fetch("/api/extract", { method: "POST", body: fd });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || res.statusText);
    render(data);
  } catch (err) {
    statusEl.innerHTML = `<span class="err">Something went wrong: ${esc(err.message)}</span>`;
  }
}

function render({ job, results, total }) {
  const base = `/api/jobs/${job}`;
  statusEl.innerHTML = total
    ? `Found <strong>${total}</strong> Prism graph${total > 1 ? "s" : ""}.`
    : `<span class="warn">No Prism graphs found.</span>`;

  let html = "";
  if (total > 1) {
    html += `<div class="toolbar">
      <a class="btn primary" href="${base}/all.zip">Download all (.zip)</a>
      <button data-save-all>Save all to Downloads</button></div>`;
  } else if (total === 1) {
    html += `<div class="toolbar"><button data-save-all>Save to Downloads</button></div>`;
  }

  for (const r of results) {
    html += `<section class="doc"><h2>${esc(r.name)}<span>${r.error ? "" : r.objects.length + " graph" + (r.objects.length === 1 ? "" : "s")}</span></h2>`;
    if (r.error) html += `<div class="note err">${esc(r.error)}</div>`;
    for (const o of r.objects) {
      html += `<div class="row">
        <div class="slide">${o.slide ? "Slide " + o.slide : "—"}</div>
        <div class="meta"><div class="fn">${esc(o.file)}<span class="badge">.${o.format}</span></div>
          <small>${esc(o.progId || "Prism")} · ${kb(o.size)}</small></div>
        <div class="actions">
          <button data-open="${esc(o.file)}">Open in Prism</button>
          <a class="btn" href="${base}/files/${encodeURIComponent(o.file)}">Download</a>
        </div></div>`;
    }
    if (r.skipped && r.skipped.length) {
      html += `<div class="note">Skipped ${r.skipped.length} other embedded object${r.skipped.length > 1 ? "s" : ""}: ` +
        r.skipped.map(s => `${s.slide ? "slide " + s.slide + " " : ""}${esc(s.progId || s.source)} (${esc(s.reason)})`).join("; ") + `</div>`;
    }
    if (!r.error && !r.objects.length && !(r.skipped && r.skipped.length)) {
      html += `<div class="note">No embedded objects in this file.</div>`;
    }
    if (r.log) html += `<details><summary>Extractor log</summary><pre>${esc(r.log)}</pre></details>`;
    html += `</section>`;
  }
  resultsEl.innerHTML = html;

  resultsEl.querySelectorAll("[data-open]").forEach(b => b.addEventListener("click", async () => {
    b.disabled = true; b.textContent = "Opening…";
    const res = await fetch(`${base}/open/${encodeURIComponent(b.dataset.open)}`, { method: "POST" });
    const data = await res.json();
    b.disabled = false; b.textContent = "Open in Prism";
    statusEl.textContent = res.ok ? `Saved and opened ${data.path}` : "Couldn't open the file.";
  }));
  resultsEl.querySelectorAll("[data-save-all]").forEach(b => b.addEventListener("click", async () => {
    const res = await fetch(`${base}/save-all`, { method: "POST" });
    const data = await res.json();
    statusEl.textContent = res.ok ? `Saved ${data.count} file${data.count > 1 ? "s" : ""} to ${data.folder}` : "Couldn't save the files.";
  }));
}

$("#quit").addEventListener("click", async () => {
  await fetch("/api/quit", { method: "POST" }).catch(() => {});
  document.body.innerHTML = '<main><h1>Prism Extractor has quit.</h1><p class="sub">You can close this tab.</p></main>';
});
</script>
</body>
</html>
"""


def run_browser():
    url = f"http://{HOST}:{PORT}/"
    if "--no-browser" not in sys.argv:
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    print(f"Prism Extractor running at {url}  (Ctrl+C to stop)")
    app.run(host=HOST, port=PORT, threaded=True)


def run_window():
    import webview
    from werkzeug.serving import make_server

    global WINDOW_MODE
    WINDOW_MODE = True
    server = make_server(HOST, 0, app, threaded=True)  # any free port
    threading.Thread(target=server.serve_forever, daemon=True).start()

    webview.settings["ALLOW_DOWNLOADS"] = True
    webview.create_window(
        "Prism Extractor", f"http://{HOST}:{server.server_port}/",
        width=900, height=760, min_size=(520, 480),
    )
    webview.start()  # returns when the window closes
    server.shutdown()


if __name__ == "__main__":
    if "--browser" in sys.argv or "--no-browser" in sys.argv:
        run_browser()
    else:
        run_window()
