"""
EIS Viewer — Electrochemical Impedance Spectroscopy file viewer.
Supports: BioLogic .mpt, Gamry .dta, Zahner .ism, Neware text/xlsx,
          generic CSV / TSV, Excel (.xlsx).
Run:  python eis_viewer.py
"""

import io
import json
import math
import os
import re
import struct
import threading
import webbrowser
from pathlib import Path

from flask import Flask, jsonify, render_template_string, request

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

app = Flask(__name__)
PORT = 5558

# ---------------------------------------------------------------------------
# Parsers
# ---------------------------------------------------------------------------

def _rows_to_eis(rows, freq_col, zre_col, zim_col):
    """Convert numeric rows into EIS dicts, skipping bad rows."""
    out = []
    for r in rows:
        try:
            freq = float(r[freq_col])
            zre  = float(r[zre_col])
            zim  = float(r[zim_col])
            if freq > 0:
                out.append({"freq": freq, "zre": zre, "zim": zim})
        except (ValueError, IndexError):
            pass
    return out


def _detect_delimiter(line):
    tab = line.count("\t")
    comma = line.count(",")
    semi = line.count(";")
    return "\t" if tab >= comma and tab >= semi else ("," if comma >= semi else ";")


def parse_biologic_mpt(text):
    """BioLogic EC-Lab .mpt — tab-delimited after 'Nb header lines' preamble."""
    lines = text.splitlines()
    skip = 0
    for ln in lines[:30]:
        m = re.match(r"Nb header lines\s*:\s*(\d+)", ln)
        if m:
            skip = int(m.group(1))
            break
    data_lines = lines[skip:]
    if not data_lines:
        return []
    hdr = data_lines[0].split("\t")
    hdr_lower = [h.lower() for h in hdr]
    def col(candidates):
        for c in candidates:
            for i, h in enumerate(hdr_lower):
                if c in h:
                    return i
        return None
    fc = col(["freq"])
    rc = col(["re(z)", "z_re", "zre", "re"])
    ic = col(["-im(z)", "im(z)", "z_im", "zim", "im"])
    if fc is None or rc is None or ic is None:
        return []
    rows = [ln.split("\t") for ln in data_lines[1:] if ln.strip()]
    pts = _rows_to_eis(rows, fc, rc, ic)
    # BioLogic stores -Im(Z) so imaginary is already negated
    return pts


def parse_gamry_dta(text):
    """Gamry .dta — find ZCURVE or ZDATA table."""
    lines = text.splitlines()
    start = None
    for i, ln in enumerate(lines):
        if re.match(r"(ZCURVE|ZDATA)\b", ln):
            start = i + 2  # skip ZCURVE header + column header
            break
    if start is None:
        return []
    rows = []
    for ln in lines[start:]:
        if not ln.strip() or ln.startswith("ENDD"):
            break
        rows.append(ln.split("\t"))
    # Gamry columns: Pt, Time, Freq, Zreal, Zimag, Zsig, Zmod, Zphz, IERange
    if not rows:
        return []
    return _rows_to_eis(rows, 2, 3, 4)


def parse_zahner_ism(text):
    """Zahner .ism — data after EOH marker."""
    lines = text.splitlines()
    start = None
    for i, ln in enumerate(lines):
        if ln.strip() == "EOH":
            start = i + 1
            break
    if start is None:
        return []
    rows = [ln.split() for ln in lines[start:] if ln.strip()]
    # columns: freq, |Z|, phase
    out = []
    for r in rows:
        try:
            freq = float(r[0])
            zmag = float(r[1])
            phase_deg = float(r[2])
            phase_rad = math.radians(phase_deg)
            zre = zmag * math.cos(phase_rad)
            zim = zmag * math.sin(phase_rad)
            if freq > 0:
                out.append({"freq": freq, "zre": zre, "zim": zim})
        except (ValueError, IndexError):
            pass
    return out


def parse_neware_text(text):
    """Neware / Hioki / generic instrument text export — tab or comma delimited."""
    lines = text.splitlines()

    FREQ_KEYS  = ["freq", "frequency"]
    REAL_KEYS  = ["zreal", "z_re", "re(z)", "real", "r(ohm", "resistance"]
    IMAG_KEYS  = ["zimag", "z_im", "-im(z)", "im(z)", "imag", "x(ohm", "reactance"]

    def col(hdr, candidates):
        for c in candidates:
            for i, h in enumerate(hdr):
                if c in h:
                    return i
        return None

    # Scan every line looking for one that has freq + real + imag columns
    for i, ln in enumerate(lines[:80]):
        ll = ln.lower()
        # Must contain a freq keyword
        if not any(k in ll for k in FREQ_KEYS):
            continue
        delim = _detect_delimiter(ln)
        hdr = [h.strip().lower() for h in ln.split(delim)]
        if len(hdr) < 3:
            continue
        fc = col(hdr, FREQ_KEYS)
        rc = col(hdr, REAL_KEYS)
        ic = col(hdr, IMAG_KEYS)
        if fc is not None and rc is not None and ic is not None:
            rows = [row.split(delim) for row in lines[i+1:] if row.strip()]
            pts = _rows_to_eis(rows, fc, rc, ic)
            if pts:
                return pts

    return _parse_generic(text)


def _parse_generic(text):
    """Auto-detect delimited text: extract numeric columns, take first three as freq/Zre/Zim."""
    lines = [ln for ln in text.splitlines() if ln.strip()]
    # Collect rows that contain at least 3 numeric fields (ignoring non-numeric columns)
    data_lines = []
    numeric_cols = None  # column indices that are consistently numeric
    for ln in lines:
        delim = _detect_delimiter(ln)
        parts = [p.strip() for p in ln.split(delim)]
        num_idx = []
        for i, p in enumerate(parts):
            try:
                float(p)
                num_idx.append(i)
            except ValueError:
                pass
        if len(num_idx) < 3:
            if data_lines:
                break
            continue
        if numeric_cols is None:
            numeric_cols = num_idx
        data_lines.append(parts)
    if len(data_lines) < 2 or numeric_cols is None or len(numeric_cols) < 3:
        return []
    fc, rc, ic = numeric_cols[0], numeric_cols[1], numeric_cols[2]
    return _rows_to_eis(data_lines, fc, rc, ic)


def parse_excel(file_bytes):
    """Excel .xlsx — search for freq/Zre/Zim columns."""
    if not HAS_OPENPYXL:
        return []
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        # Find header row
        for hi, row in enumerate(rows[:20]):
            cells = [str(c).lower() if c is not None else "" for c in row]
            if any("freq" in c or "hz" in c for c in cells):
                hdr = cells
                def col(candidates):
                    for cand in candidates:
                        for i, h in enumerate(hdr):
                            if cand in h:
                                return i
                    return None
                fc = col(["freq"])
                rc = col(["zreal", "z_re", "re(z)", "real"])
                ic = col(["zimag", "z_im", "-im(z)", "im(z)", "imag"])
                if fc is None:
                    continue
                if rc is None:
                    rc = 1
                if ic is None:
                    ic = 2
                out = []
                for r in rows[hi+1:]:
                    try:
                        freq = float(r[fc])
                        zre  = float(r[rc])
                        zim  = float(r[ic])
                        if freq > 0:
                            out.append({"freq": freq, "zre": zre, "zim": zim})
                    except (TypeError, ValueError, IndexError):
                        pass
                if out:
                    return out
    return []


def parse_eis_file(filename, file_bytes):
    """Dispatch to the right parser based on extension / content sniffing."""
    name_lower = filename.lower()
    # Try to decode as text first
    text = None
    for enc in ("utf-8", "latin-1", "cp1252"):
        try:
            text = file_bytes.decode(enc)
            break
        except UnicodeDecodeError:
            pass

    if name_lower.endswith(".mpt") and text:
        pts = parse_biologic_mpt(text)
        if pts:
            return pts

    if name_lower.endswith(".dta") and text:
        pts = parse_gamry_dta(text)
        if pts:
            return pts

    if name_lower.endswith(".ism") and text:
        pts = parse_zahner_ism(text)
        if pts:
            return pts

    if name_lower.endswith((".xlsx", ".xls")):
        pts = parse_excel(file_bytes)
        if pts:
            return pts

    if text:
        # Try BioLogic
        if "Nb header lines" in text:
            pts = parse_biologic_mpt(text)
            if pts:
                return pts
        # Try Gamry
        if "ZCURVE" in text or "ZDATA" in text:
            pts = parse_gamry_dta(text)
            if pts:
                return pts
        # Try Zahner
        if "EOH" in text:
            pts = parse_zahner_ism(text)
            if pts:
                return pts
        # Neware / generic
        pts = parse_neware_text(text)
        if pts:
            return pts

    return []


def enrich(pts):
    """Add |Z|, phase_deg to each point."""
    for p in pts:
        zre = p["zre"]
        zim = p["zim"]
        p["zmag"]      = math.sqrt(zre**2 + zim**2)
        p["phase_deg"] = math.degrees(math.atan2(zim, zre))
    return pts


# ---------------------------------------------------------------------------
# HTML template
# ---------------------------------------------------------------------------

HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>EIS Viewer</title>
<script src="https://cdn.plot.ly/plotly-2.32.0.min.js"></script>
<style>
  :root{--navy:#0c2143;--blue:#2563eb;--bg:#f0f4f8;--card:#fff;--border:#d1d5db;--muted:#64748b}
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:system-ui,sans-serif;background:var(--bg);color:#1e293b;min-height:100vh}
  header{background:var(--navy);color:#fff;padding:14px 28px;display:flex;align-items:center;gap:16px}
  header h1{font-size:1.25rem;font-weight:700;letter-spacing:.02em}
  header span{font-size:.85rem;opacity:.65}
  .main{max-width:1200px;margin:0 auto;padding:24px 20px}

  /* Drop zone */
  #dropzone{border:2px dashed var(--blue);border-radius:12px;background:#eff6ff;
    padding:40px 20px;text-align:center;cursor:pointer;transition:.2s}
  #dropzone.over{background:#dbeafe;border-color:#1d4ed8}
  #dropzone p{color:var(--blue);font-weight:600;font-size:1.05rem;margin-bottom:6px}
  #dropzone small{color:var(--muted)}
  #file-input{display:none}

  /* File list */
  #file-list{margin-top:18px;display:flex;flex-wrap:wrap;gap:10px}
  .file-chip{display:flex;align-items:center;gap:8px;background:var(--card);
    border:1px solid var(--border);border-radius:6px;padding:6px 12px;font-size:.85rem}
  .file-chip .dot{width:10px;height:10px;border-radius:50%;flex-shrink:0}
  .file-chip button{background:none;border:none;cursor:pointer;color:var(--muted);
    font-size:1rem;line-height:1;padding:0 2px}
  .file-chip button:hover{color:#dc2626}

  /* Controls */
  .controls{margin-top:16px;display:flex;gap:10px;flex-wrap:wrap;align-items:center}
  .btn{padding:8px 18px;border-radius:6px;border:none;cursor:pointer;font-size:.9rem;font-weight:600;transition:.15s}
  .btn-primary{background:var(--blue);color:#fff}
  .btn-primary:hover{background:#1d4ed8}
  .btn-secondary{background:var(--card);color:var(--navy);border:1px solid var(--border)}
  .btn-secondary:hover{background:#f1f5f9}
  #status{font-size:.85rem;color:var(--muted);align-self:center}

  /* Plot grid */
  .plot-grid{display:grid;grid-template-columns:1fr 1fr;gap:18px;margin-top:24px}
  @media(max-width:760px){.plot-grid{grid-template-columns:1fr}}
  .plot-card{background:var(--card);border:1px solid var(--border);border-radius:10px;
    padding:14px;box-shadow:0 1px 4px rgba(0,0,0,.06)}
  .plot-card h3{font-size:.9rem;color:var(--navy);margin-bottom:8px;font-weight:600}
  .plot-wrap{width:100%;height:340px}

  /* Error */
  .err{background:#fef2f2;border:1px solid #fca5a5;border-radius:6px;
    padding:10px 14px;color:#dc2626;font-size:.85rem;margin-top:12px}
</style>
</head>
<body>

<header>
  <div>
    <h1>EIS Viewer</h1>
    <span>Nyquist &amp; Bode plots for electrochemical impedance data</span>
  </div>
</header>

<div class="main">
  <div id="dropzone">
    <p>Drop EIS files here or click to browse</p>
    <small>Supports: BioLogic .mpt &nbsp;|&nbsp; Gamry .dta &nbsp;|&nbsp; Zahner .ism &nbsp;|&nbsp; Neware text &nbsp;|&nbsp; CSV / TSV &nbsp;|&nbsp; Excel .xlsx</small>
    <input type="file" id="file-input" multiple>
  </div>

  <div id="file-list"></div>

  <div class="controls">
    <button class="btn btn-primary" id="btn-plot">Plot</button>
    <button class="btn btn-secondary" id="btn-clear">Clear all</button>
    <span id="status"></span>
  </div>
  <div id="error-box"></div>

  <div class="plot-grid" id="plot-grid" style="display:none">
    <div class="plot-card" style="grid-column:1/-1">
      <h3>Nyquist Plot — Z′ vs −Z″</h3>
      <div class="plot-wrap" id="plt-nyquist"></div>
    </div>
    <div class="plot-card">
      <h3>Bode — |Z| vs Frequency</h3>
      <div class="plot-wrap" id="plt-bode-mag"></div>
    </div>
    <div class="plot-card">
      <h3>Bode — Phase vs Frequency</h3>
      <div class="plot-wrap" id="plt-bode-phase"></div>
    </div>
    <div class="plot-card">
      <h3>Z′ &amp; Z″ vs Frequency</h3>
      <div class="plot-wrap" id="plt-components"></div>
    </div>
  </div>
</div>

<script>
const PALETTE = [
  "#2563eb","#dc2626","#16a34a","#d97706","#7c3aed",
  "#0891b2","#db2777","#65a30d","#ea580c","#0f766e"
];

let fileQueue = [];  // {file, name, color, id}
let idCounter = 0;

const dropzone   = document.getElementById("dropzone");
const fileInput  = document.getElementById("file-input");
const fileList   = document.getElementById("file-list");
const btnPlot    = document.getElementById("btn-plot");
const btnClear   = document.getElementById("btn-clear");
const status     = document.getElementById("status");
const errorBox   = document.getElementById("error-box");
const plotGrid   = document.getElementById("plot-grid");

dropzone.addEventListener("click", () => fileInput.click());
dropzone.addEventListener("dragover", e => { e.preventDefault(); dropzone.classList.add("over"); });
dropzone.addEventListener("dragleave", () => dropzone.classList.remove("over"));
dropzone.addEventListener("drop", e => {
  e.preventDefault();
  dropzone.classList.remove("over");
  addFiles(e.dataTransfer.files);
});
fileInput.addEventListener("change", () => addFiles(fileInput.files));

function addFiles(files) {
  for (const f of files) {
    const id    = ++idCounter;
    const color = PALETTE[(fileQueue.length) % PALETTE.length];
    fileQueue.push({ file: f, name: f.name, color, id });
  }
  renderFileList();
  fileInput.value = "";
}

function removeFile(id) {
  fileQueue = fileQueue.filter(f => f.id !== id);
  renderFileList();
}

function renderFileList() {
  fileList.innerHTML = "";
  fileQueue.forEach(f => {
    const chip = document.createElement("div");
    chip.className = "file-chip";
    chip.innerHTML = `<span class="dot" style="background:${f.color}"></span>
      <span>${f.name}</span>
      <button title="Remove" onclick="removeFile(${f.id})">×</button>`;
    fileList.appendChild(chip);
  });
}

btnClear.addEventListener("click", () => {
  fileQueue = [];
  renderFileList();
  plotGrid.style.display = "none";
  errorBox.innerHTML = "";
  status.textContent = "";
});

btnPlot.addEventListener("click", async () => {
  if (!fileQueue.length) { status.textContent = "No files selected."; return; }
  status.textContent = "Parsing…";
  errorBox.innerHTML = "";
  btnPlot.disabled = true;

  const datasets = [];
  const errors   = [];

  for (const entry of fileQueue) {
    const fd = new FormData();
    fd.append("file", entry.file, entry.name);
    try {
      const res = await fetch("/parse", { method: "POST", body: fd });
      const json = await res.json();
      if (json.error)  { errors.push(`${entry.name}: ${json.error}`); continue; }
      if (!json.points.length) { errors.push(`${entry.name}: no data found`); continue; }
      datasets.push({ name: entry.name, color: entry.color, points: json.points });
    } catch(e) {
      errors.push(`${entry.name}: network error`);
    }
  }

  btnPlot.disabled = false;
  if (errors.length) {
    errorBox.innerHTML = `<div class="err">${errors.join("<br>")}</div>`;
  }
  if (!datasets.length) { status.textContent = "No data plotted."; return; }

  status.textContent = `Plotted ${datasets.length} dataset(s).`;
  buildPlots(datasets);
});

function sortByFreq(pts) {
  return [...pts].sort((a,b) => b.freq - a.freq);
}

const LAYOUT_BASE = {
  margin: {l:56,r:20,t:20,b:50},
  paper_bgcolor:"#fff", plot_bgcolor:"#f8fafc",
  font: {family:"system-ui,sans-serif", size:12},
  legend:{orientation:"h", y:-0.18},
  hovermode:"closest"
};

function buildPlots(datasets) {
  plotGrid.style.display = "";

  // --- Nyquist ---
  const nyqTraces = datasets.map(ds => {
    const s = sortByFreq(ds.points);
    return {
      x: s.map(p => p.zre),
      y: s.map(p => -p.zim),
      text: s.map(p => `${p.freq.toFixed(3)} Hz`),
      mode:"lines+markers",
      name: ds.name,
      line:{color: ds.color, width:1.5},
      marker:{size:4, color: ds.color},
      hovertemplate:"Z′=%{x:.3f} Ω<br>−Z″=%{y:.3f} Ω<br>%{text}<extra>%{fullData.name}</extra>"
    };
  });
  Plotly.newPlot("plt-nyquist", nyqTraces, {
    ...LAYOUT_BASE,
    xaxis:{title:"Z′ (Ω)", zeroline:true, zerolinecolor:"#94a3b8"},
    yaxis:{title:"−Z″ (Ω)", zeroline:true, zerolinecolor:"#94a3b8", scaleanchor:"x", scaleratio:1}
  }, {responsive:true, displayModeBar:true});

  // --- Bode magnitude ---
  const magTraces = datasets.map(ds => {
    const s = sortByFreq(ds.points);
    return {
      x: s.map(p => p.freq),
      y: s.map(p => p.zmag),
      mode:"lines+markers",
      name: ds.name,
      line:{color: ds.color, width:1.5},
      marker:{size:4},
      hovertemplate:"f=%{x:.3g} Hz<br>|Z|=%{y:.4f} Ω<extra>%{fullData.name}</extra>"
    };
  });
  Plotly.newPlot("plt-bode-mag", magTraces, {
    ...LAYOUT_BASE,
    xaxis:{title:"Frequency (Hz)", type:"log"},
    yaxis:{title:"|Z| (Ω)", type:"log"}
  }, {responsive:true, displayModeBar:true});

  // --- Bode phase ---
  const phTraces = datasets.map(ds => {
    const s = sortByFreq(ds.points);
    return {
      x: s.map(p => p.freq),
      y: s.map(p => p.phase_deg),
      mode:"lines+markers",
      name: ds.name,
      line:{color: ds.color, width:1.5},
      marker:{size:4},
      hovertemplate:"f=%{x:.3g} Hz<br>Phase=%{y:.2f}°<extra>%{fullData.name}</extra>"
    };
  });
  Plotly.newPlot("plt-bode-phase", phTraces, {
    ...LAYOUT_BASE,
    xaxis:{title:"Frequency (Hz)", type:"log"},
    yaxis:{title:"Phase (°)"}
  }, {responsive:true, displayModeBar:true});

  // --- Z components vs freq ---
  const compTraces = [];
  datasets.forEach(ds => {
    const s = sortByFreq(ds.points);
    compTraces.push({
      x: s.map(p => p.freq), y: s.map(p => p.zre),
      mode:"lines+markers", name:`Z′ — ${ds.name}`,
      line:{color: ds.color, width:1.5, dash:"solid"},
      marker:{size:4},
      hovertemplate:"f=%{x:.3g} Hz<br>Z′=%{y:.4f} Ω<extra>%{fullData.name}</extra>"
    });
    compTraces.push({
      x: s.map(p => p.freq), y: s.map(p => -p.zim),
      mode:"lines+markers", name:`−Z″ — ${ds.name}`,
      line:{color: ds.color, width:1.5, dash:"dot"},
      marker:{size:4, symbol:"triangle-up"},
      hovertemplate:"f=%{x:.3g} Hz<br>−Z″=%{y:.4f} Ω<extra>%{fullData.name}</extra>"
    });
  });
  Plotly.newPlot("plt-components", compTraces, {
    ...LAYOUT_BASE,
    xaxis:{title:"Frequency (Hz)", type:"log"},
    yaxis:{title:"Impedance (Ω)"}
  }, {responsive:true, displayModeBar:true});
}
</script>
</body>
</html>
"""


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template_string(HTML)


@app.route("/parse", methods=["POST"])
def parse_route():
    if "file" not in request.files:
        return jsonify({"error": "no file"}), 400
    f = request.files["file"]
    raw = f.read()
    try:
        pts = parse_eis_file(f.filename, raw)
        pts = enrich(pts)
    except Exception as e:
        return jsonify({"error": str(e)}), 200
    if not pts:
        return jsonify({"error": "Could not parse — unrecognised format or no numeric data found."}), 200
    return jsonify({"points": pts})


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    url = f"http://localhost:{PORT}"
    threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    print(f"EIS Viewer running at {url}")
    app.run(port=PORT, debug=False)
