#!/usr/bin/env python3
"""
JOOR -> Heartland converter: drag-and-drop app.  (companion to joor_to_heartland.py)

Run:  python3 joor_converter_app.py     (or double-click "JOOR Converter.command")
A browser page opens. Drag the JOOR export in, click Convert, and the
Heartland import spreadsheet downloads to your Downloads folder.

All conversion logic lives in joor_to_heartland.py (must be in the same
folder); this file is only the user interface around it.
"""
import base64, io, json, socket, sys, tempfile, threading, traceback, webbrowser, zipfile
from contextlib import redirect_stdout
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
try:
    from joor_to_heartland import read_joor, build_grids, build_po
except ImportError:
    sys.exit("ERROR: joor_to_heartland.py must be in the same folder as this app.")

PAGE = """<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>JOOR &rarr; Heartland Converter</title>
<style>
  :root { --ink:#2b2b2b; --soft:#8a8a8a; --line:#e3ded6; --bg:#faf8f4; --accent:#5b4636; }
  * { box-sizing:border-box; }
  body { font-family:-apple-system, "Helvetica Neue", Arial, sans-serif; background:var(--bg);
         color:var(--ink); max-width:680px; margin:40px auto; padding:0 20px; }
  h1 { font-weight:600; font-size:22px; letter-spacing:.3px; }
  h1 small { color:var(--soft); font-weight:400; font-size:13px; margin-left:8px; }
  .tabs { display:flex; gap:8px; margin:18px 0 14px; }
  .tab { padding:8px 18px; border:1px solid var(--line); border-radius:20px; cursor:pointer;
         background:#fff; font-size:14px; }
  .tab.active { background:var(--accent); color:#fff; border-color:var(--accent); }
  .hint { color:var(--soft); font-size:13px; margin:0 0 14px; line-height:1.5; }
  .drop { border:2px dashed var(--line); border-radius:12px; background:#fff; padding:34px 20px;
          text-align:center; color:var(--soft); font-size:15px; cursor:pointer; transition:.15s; }
  .drop.small { padding:20px; margin-top:12px; }
  .drop.over { border-color:var(--accent); color:var(--accent); background:#f5efe7; }
  .drop.loaded { border-style:solid; border-color:#9db89d; color:var(--ink); }
  .drop b { color:var(--accent); }
  button { margin-top:16px; width:100%; padding:13px; font-size:16px; border:none; border-radius:10px;
           background:var(--accent); color:#fff; cursor:pointer; }
  button:disabled { background:#cfc8bf; cursor:default; }
  #log { display:none; white-space:pre-wrap; background:#1f1d1a; color:#e8e4dc; border-radius:10px;
         padding:14px 16px; font:12.5px/1.55 ui-monospace, Menlo, monospace; margin-top:18px; }
  #log .warn { color:#f2b866; font-weight:600; }
  #log .ok { color:#9dd49d; }
  #files { margin-top:10px; }
  .chip { display:inline-block; background:#fff; border:1px solid var(--line); border-radius:8px;
          padding:6px 12px; margin:4px 6px 0 0; font-size:13px; }
</style></head><body>
<h1>JOOR &rarr; Heartland <small>Romi Boutique converter</small></h1>
<div class="tabs">
  <div class="tab active" id="tab-grids" onclick="setPhase('grids')">Step 1 &middot; Grids</div>
  <div class="tab" id="tab-po" onclick="setPhase('po')">Step 2 &middot; Purchase Order</div>
</div>
<p class="hint" id="hint-grids">Drop the JOOR order export below (either JOOR format works &mdash;
Basic Export or Order-to-Size). You&rsquo;ll get the <b>grid import file</b> for
Heartland &rarr; Inventory &rarr; Item Grids &rarr; Import. Don&rsquo;t open the JOOR file in Excel first.</p>
<p class="hint" id="hint-po" style="display:none">Drop the <b>same JOOR export</b> again.
Item #s are looked up live from Heartland, so no items export is needed. You&rsquo;ll get
the PO import file for Purchasing &rarr; Purchase Orders &rarr; Import. Remember: it lands
as Pending &mdash; open it manually.</p>

<div class="drop" id="drop-joor">Drag the <b>JOOR export</b> here<br>(or click to browse)</div>
<input type="file" id="file-joor" hidden accept=".xlsx,.csv,.txt,.tsv">
<button id="go" disabled onclick="convert()">Convert</button>
<div id="files"></div>
<div id="log"></div>

<script>
let phase = 'grids', joorFile = null;

function setPhase(p) {
  phase = p;
  document.getElementById('tab-grids').classList.toggle('active', p==='grids');
  document.getElementById('tab-po').classList.toggle('active', p==='po');
  document.getElementById('hint-grids').style.display = p==='grids' ? '' : 'none';
  document.getElementById('hint-po').style.display = p==='po' ? '' : 'none';
  refresh();
}
function wire(dropId, inputId, setter) {
  const d = document.getElementById(dropId), inp = document.getElementById(inputId);
  d.onclick = () => inp.click();
  inp.onchange = () => { if (inp.files[0]) setter(inp.files[0], d); };
  d.ondragover = e => { e.preventDefault(); d.classList.add('over'); };
  d.ondragleave = () => d.classList.remove('over');
  d.ondrop = e => { e.preventDefault(); d.classList.remove('over');
                    if (e.dataTransfer.files[0]) setter(e.dataTransfer.files[0], d); };
}
function setJoor(f, d)  { joorFile = f;  d.classList.add('loaded'); d.innerHTML = 'JOOR file: <b>'+f.name+'</b>'; refresh(); }
wire('drop-joor', 'file-joor', setJoor);
function refresh() {
  document.getElementById('go').disabled = !joorFile;
}
function b64(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result.split(',')[1]);
    r.onerror = rej; r.readAsDataURL(file);
  });
}
async function convert() {
  const btn = document.getElementById('go'); btn.disabled = true; btn.textContent = 'Converting\u2026';
  const body = { phase, joor: { name: joorFile.name, b64: await b64(joorFile) } };
  let resp;
  try {
    resp = await (await fetch('/convert', { method:'POST', body: JSON.stringify(body) })).json();
  } catch (e) {
    resp = { error: 'Could not reach the converter \u2014 is the app still running? ' + e };
  }
  const log = document.getElementById('log'); log.style.display = 'block';
  log.innerHTML = (resp.log || (resp.error||''))
      .replace(/&/g,'&amp;').replace(/</g,'&lt;')
      .replace(/^(\\s*WARNING.*)$/gm, '<span class="warn">$1</span>')
      .replace(/^(\\s*NOTE.*)$/gm, '<span class="warn">$1</span>')
      .replace(/^(wrote .*)$/gm, '<span class="ok">$1</span>');
  const chips = document.getElementById('files'); chips.innerHTML = '';
  for (const f of (resp.files || [])) {
    const a = document.createElement('a');
    a.href = 'data:application/octet-stream;base64,' + f.b64;
    a.download = f.name; document.body.appendChild(a); a.click(); a.remove();
    const c = document.createElement('span'); c.className = 'chip';
    c.textContent = '\u2b07 ' + f.name; chips.appendChild(c);
  }
  btn.textContent = 'Convert'; refresh();
}
</script></body></html>"""

def run_conversion(payload):
    """Returns dict {log, files:[{name,b64}]}; raises nothing (errors go in log)."""
    log = io.StringIO()
    out_files = []
    with tempfile.TemporaryDirectory() as td:
        tdir = Path(td)
        joor_path = tdir / payload["joor"]["name"]
        joor_path.write_bytes(base64.b64decode(payload["joor"]["b64"]))
        stem = joor_path.stem[:40]
        try:
            with redirect_stdout(log):
                rows, images = read_joor(joor_path)
                print(f"read {len(rows)} JOOR line(s) from {joor_path.name}")
                if payload["phase"] == "grids":
                    out = tdir / f"{stem}_GRIDS_IMPORT.xlsx"
                    build_grids(rows, out)
                    out_files.append(out)
                else:
                    out = tdir / f"{stem}_PO_IMPORT.xlsx"
                    build_po(rows, out)
                    out_files.append(out)
        except SystemExit as e:                    # converter's sys.exit(...) messages
            print(e, file=log)
        except Exception:
            print("ERROR:\n" + traceback.format_exc(), file=log)
        files = [{"name": f.name,
                  "b64": base64.b64encode(f.read_bytes()).decode()} for f in out_files]
    return {"log": log.getvalue(), "files": files}

class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):                     # keep the terminal quiet
        pass
    def _send(self, code, body, ctype):
        data = body.encode()
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)
    def do_GET(self):
        if self.path in ("/", "/index.html"):
            self._send(200, PAGE, "text/html; charset=utf-8")
        else:
            self._send(404, "not found", "text/plain")
    def do_POST(self):
        if self.path != "/convert":
            return self._send(404, "not found", "text/plain")
        try:
            payload = json.loads(self.rfile.read(int(self.headers["Content-Length"])))
            result = run_conversion(payload)
        except Exception:
            result = {"log": "ERROR:\n" + traceback.format_exc(), "files": []}
        self._send(200, json.dumps(result), "application/json")

def main():
    with socket.socket() as s:                     # find a free port
        s.bind(("127.0.0.1", 0))
        port = s.getsockname()[1]
    srv = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    url = f"http://127.0.0.1:{port}/"
    print(f"JOOR -> Heartland converter running at {url}")
    print("Leave this window open while you work; close it to quit.")
    threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass

if __name__ == "__main__":
    main()
