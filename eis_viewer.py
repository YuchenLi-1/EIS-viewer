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
import threading
import webbrowser

import numpy as np
from flask import Flask, jsonify, render_template_string, request
from scipy.optimize import least_squares

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
    tab   = line.count("\t")
    comma = line.count(",")
    semi  = line.count(";")
    return "\t" if tab >= comma and tab >= semi else ("," if comma >= semi else ";")


def parse_biologic_mpt(text):
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
    return _rows_to_eis(rows, fc, rc, ic)


def parse_gamry_dta(text):
    lines = text.splitlines()
    start = None
    for i, ln in enumerate(lines):
        if re.match(r"(ZCURVE|ZDATA)\b", ln):
            start = i + 2
            break
    if start is None:
        return []
    rows = []
    for ln in lines[start:]:
        if not ln.strip() or ln.startswith("ENDD"):
            break
        rows.append(ln.split("\t"))
    if not rows:
        return []
    return _rows_to_eis(rows, 2, 3, 4)


def parse_zahner_ism(text):
    lines = text.splitlines()
    start = None
    for i, ln in enumerate(lines):
        if ln.strip() == "EOH":
            start = i + 1
            break
    if start is None:
        return []
    rows = [ln.split() for ln in lines[start:] if ln.strip()]
    out = []
    for r in rows:
        try:
            freq      = float(r[0])
            zmag      = float(r[1])
            phase_rad = math.radians(float(r[2]))
            if freq > 0:
                out.append({"freq": freq,
                            "zre":  zmag * math.cos(phase_rad),
                            "zim":  zmag * math.sin(phase_rad)})
        except (ValueError, IndexError):
            pass
    return out


def parse_neware_text(text):
    lines = text.splitlines()
    FREQ_KEYS = ["freq", "frequency"]
    REAL_KEYS = ["zreal", "z_re", "re(z)", "real", "r(ohm", "resistance"]
    IMAG_KEYS = ["zimag", "z_im", "-im(z)", "im(z)", "imag", "x(ohm", "reactance"]

    def col(hdr, candidates):
        for c in candidates:
            for i, h in enumerate(hdr):
                if c in h:
                    return i
        return None

    for i, ln in enumerate(lines[:80]):
        if not any(k in ln.lower() for k in FREQ_KEYS):
            continue
        delim = _detect_delimiter(ln)
        hdr   = [h.strip().lower() for h in ln.split(delim)]
        if len(hdr) < 3:
            continue
        fc = col(hdr, FREQ_KEYS)
        rc = col(hdr, REAL_KEYS)
        ic = col(hdr, IMAG_KEYS)
        if fc is not None and rc is not None and ic is not None:
            rows = [row.split(delim) for row in lines[i+1:] if row.strip()]
            pts  = _rows_to_eis(rows, fc, rc, ic)
            if pts:
                return pts
    return _parse_generic(text)


def _parse_generic(text):
    lines        = [ln for ln in text.splitlines() if ln.strip()]
    data_lines   = []
    numeric_cols = None
    for ln in lines:
        delim   = _detect_delimiter(ln)
        parts   = [p.strip() for p in ln.split(delim)]
        num_idx = [i for i, p in enumerate(parts) if _is_float(p)]
        if len(num_idx) < 3:
            if data_lines:
                break
            continue
        if numeric_cols is None:
            numeric_cols = num_idx
        data_lines.append(parts)
    if len(data_lines) < 2 or not numeric_cols or len(numeric_cols) < 3:
        return []
    return _rows_to_eis(data_lines, numeric_cols[0], numeric_cols[1], numeric_cols[2])


def _is_float(s):
    try:
        float(s)
        return True
    except ValueError:
        return False


def parse_excel(file_bytes):
    if not HAS_OPENPYXL:
        return []
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        for hi, row in enumerate(rows[:20]):
            cells = [str(c).lower() if c is not None else "" for c in row]
            if not any("freq" in c or "hz" in c for c in cells):
                continue
            hdr = cells
            def col(candidates):
                for cand in candidates:
                    for i, h in enumerate(hdr):
                        if cand in h:
                            return i
                return None
            fc = col(["freq"])
            rc = col(["zreal","z_re","re(z)","real"]) or 1
            ic = col(["zimag","z_im","-im(z)","im(z)","imag"]) or 2
            if fc is None:
                continue
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
    name_lower = filename.lower()
    text = None
    for enc in ("utf-8", "latin-1", "cp1252"):
        try:
            text = file_bytes.decode(enc)
            break
        except UnicodeDecodeError:
            pass

    if name_lower.endswith(".mpt") and text:
        pts = parse_biologic_mpt(text)
        if pts: return pts
    if name_lower.endswith(".dta") and text:
        pts = parse_gamry_dta(text)
        if pts: return pts
    if name_lower.endswith(".ism") and text:
        pts = parse_zahner_ism(text)
        if pts: return pts
    if name_lower.endswith((".xlsx", ".xls")):
        pts = parse_excel(file_bytes)
        if pts: return pts
    if text:
        if "Nb header lines" in text:
            pts = parse_biologic_mpt(text)
            if pts: return pts
        if "ZCURVE" in text or "ZDATA" in text:
            pts = parse_gamry_dta(text)
            if pts: return pts
        if "EOH" in text:
            pts = parse_zahner_ism(text)
            if pts: return pts
        pts = parse_neware_text(text)
        if pts: return pts
    return []


def enrich(pts):
    for p in pts:
        p["zmag"]      = math.sqrt(p["zre"]**2 + p["zim"]**2)
        p["phase_deg"] = math.degrees(math.atan2(p["zim"], p["zre"]))
    return pts


# ---------------------------------------------------------------------------
# Equivalent circuit fitting  —  L + Rs + (Rct + W) || CPE
# ---------------------------------------------------------------------------
# Parameters: [L, Rs, Rct, Q, n, sigma]
#   L     — inductance (H)
#   Rs    — ohmic/series resistance (Ω)
#   Rct   — charge-transfer resistance (Ω)
#   Q     — CPE coefficient (F·s^(n-1))
#   n     — CPE exponent  (0 < n ≤ 1)
#   sigma — Warburg coefficient (Ω·s^-0.5)

def _randles_Z(params, omega):
    L, Rs, Rct, Q, n, sigma = params
    jw    = 1j * omega
    Z_W   = sigma * (1.0 - 1j) / np.sqrt(omega)        # Warburg
    Z_CPE = 1.0 / (Q * (jw ** n))                      # CPE
    Z_rct = Rct + Z_W                                   # Rct + Warburg series
    Z_par = Z_rct * Z_CPE / (Z_rct + Z_CPE)            # parallel arm
    return 1j * omega * L + Rs + Z_par


def _residuals(params, omega, Z_meas):
    Z_calc = _randles_Z(params, omega)
    return np.concatenate([Z_calc.real - Z_meas.real,
                           Z_calc.imag - Z_meas.imag])


def fit_randles(points):
    """
    Fit Randles circuit to EIS points.
    Returns dict with fitted params, curve, and quality metric.
    """
    pts   = sorted(points, key=lambda p: p["freq"], reverse=True)
    freqs = np.array([p["freq"] for p in pts])
    Z     = np.array([p["zre"] + 1j * p["zim"] for p in pts])
    omega = 2 * np.pi * freqs

    # --- initial estimates ---
    # Rs: real part where -Zim is closest to zero (high-freq side)
    hf_idx = np.argmin(np.abs(Z.imag[:len(pts)//3 + 1]))
    Rs0    = max(Z[hf_idx].real, 1e-9)

    # L: from the inductive (positive Zim) region at highest freq
    inductive = Z.imag > 0
    if inductive.any():
        idx_ind = np.where(inductive)[0]
        # L ≈ Zim / omega at the highest inductive frequency
        L0 = float(np.median(Z[idx_ind].imag / omega[idx_ind]))
        L0 = max(L0, 1e-12)
    else:
        L0 = 1e-9

    # Rct: arc diameter ≈ difference between low-freq and Rs real parts
    Rct0 = max(Z[-1].real - Rs0, Rs0 * 0.1, 1e-9)

    # peak of -Zim gives characteristic frequency → C estimate
    neg_zim = -Z.imag
    peak_i  = int(np.argmax(neg_zim))
    omega_p = omega[peak_i] if neg_zim[peak_i] > 0 else float(np.median(omega))
    C0      = 1.0 / (omega_p * Rct0)
    Q0, n0  = max(C0, 1e-9), 0.85

    # sigma: small Warburg start
    sigma0 = abs(Z[-1].imag) * math.sqrt(omega[-1]) * 0.5 if omega[-1] > 0 else 1e-4

    x0     = [L0, Rs0, Rct0, Q0, n0, sigma0]
    bounds = ([0,    0,    0,    1e-12, 0.3, 0   ],
              [1e-3, 1e3,  1e6,  1e3,   1.0, 1e6 ])

    try:
        result = least_squares(_residuals, x0, args=(omega, Z),
                               bounds=bounds, max_nfev=5000,
                               ftol=1e-10, xtol=1e-10)
        p = result.x
    except Exception:
        p = x0   # return initial estimates on failure

    # Quality: normalised RMS error (%)
    Z_fit  = _randles_Z(p, omega)
    rms    = float(np.sqrt(np.mean(np.abs(Z_fit - Z)**2)))
    z_rms  = float(np.sqrt(np.mean(np.abs(Z)**2)))
    quality = round(rms / z_rms * 100, 2) if z_rms > 0 else 999.0

    # Generate smooth fitted curve (100 pts log-spaced)
    f_fit  = np.logspace(np.log10(freqs[-1]), np.log10(freqs[0]), 200)
    w_fit  = 2 * np.pi * f_fit
    Z_curve = _randles_Z(p, w_fit)

    L, Rs, Rct, Q, n, sigma = p
    return {
        "L":      round(float(L),     12),
        "Rs":     round(float(Rs),    9),
        "Rct":    round(float(Rct),   9),
        "Q":      round(float(Q),     12),
        "n":      round(float(n),     4),
        "sigma":  round(float(sigma), 6),
        "quality_pct": quality,
        "curve": [
            {"freq": float(f), "zre": float(z.real), "zim": float(z.imag)}
            for f, z in zip(f_fit, Z_curve)
        ],
    }


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
  :root{--navy:#0c2143;--blue:#2563eb;--green:#16a34a;--bg:#f0f4f8;--card:#fff;--border:#d1d5db;--muted:#64748b}
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:system-ui,sans-serif;background:var(--bg);color:#1e293b;min-height:100vh}
  header{background:var(--navy);color:#fff;padding:14px 28px;display:flex;align-items:center;gap:16px}
  header h1{font-size:1.25rem;font-weight:700;letter-spacing:.02em}
  header span{font-size:.85rem;opacity:.65}
  .main{max-width:1200px;margin:0 auto;padding:24px 20px}

  #dropzone{border:2px dashed var(--blue);border-radius:12px;background:#eff6ff;
    padding:40px 20px;text-align:center;cursor:pointer;transition:.2s}
  #dropzone.over{background:#dbeafe;border-color:#1d4ed8}
  #dropzone p{color:var(--blue);font-weight:600;font-size:1.05rem;margin-bottom:6px}
  #dropzone small{color:var(--muted)}
  #file-input{display:none}

  #file-list{margin-top:18px;display:flex;flex-wrap:wrap;gap:10px}
  .file-chip{display:flex;align-items:center;gap:8px;background:var(--card);
    border:1px solid var(--border);border-radius:6px;padding:6px 12px;font-size:.85rem}
  .file-chip .dot{width:10px;height:10px;border-radius:50%;flex-shrink:0}
  .file-chip button{background:none;border:none;cursor:pointer;color:var(--muted);font-size:1rem;line-height:1;padding:0 2px}
  .file-chip button:hover{color:#dc2626}

  .controls{margin-top:16px;display:flex;gap:10px;flex-wrap:wrap;align-items:center}
  .btn{padding:8px 18px;border-radius:6px;border:none;cursor:pointer;font-size:.9rem;font-weight:600;transition:.15s}
  .btn-primary{background:var(--blue);color:#fff}
  .btn-primary:hover{background:#1d4ed8}
  .btn-green{background:var(--green);color:#fff}
  .btn-green:hover{background:#15803d}
  .btn-secondary{background:var(--card);color:var(--navy);border:1px solid var(--border)}
  .btn-secondary:hover{background:#f1f5f9}
  #status{font-size:.85rem;color:var(--muted);align-self:center}

  .plot-grid{display:grid;grid-template-columns:1fr 1fr;gap:18px;margin-top:24px}
  @media(max-width:760px){.plot-grid{grid-template-columns:1fr}}
  .plot-card{background:var(--card);border:1px solid var(--border);border-radius:10px;
    padding:14px;box-shadow:0 1px 4px rgba(0,0,0,.06)}
  .plot-card h3{font-size:.9rem;color:var(--navy);margin-bottom:8px;font-weight:600}
  .plot-wrap{width:100%;height:340px}

  /* Summary table */
  .summary-card{background:var(--card);border:1px solid var(--border);border-radius:10px;
    padding:20px;margin-top:24px;box-shadow:0 1px 4px rgba(0,0,0,.06)}
  .summary-card h3{font-size:1rem;font-weight:700;color:var(--navy);margin-bottom:14px;
    display:flex;justify-content:space-between;align-items:center}
  .tbl-wrap{overflow-x:auto}
  table{width:100%;border-collapse:collapse;font-size:.83rem;white-space:nowrap}
  thead th{background:var(--navy);color:#fff;padding:8px 12px;text-align:left;font-weight:600}
  tbody td{padding:7px 12px;border-bottom:1px solid #f1f5f9}
  tbody tr:hover td{background:#f8fafc}
  tbody tr:last-child td{border-bottom:none}
  .qual-good{color:var(--green);font-weight:600}
  .qual-ok{color:#d97706;font-weight:600}
  .qual-bad{color:#dc2626;font-weight:600}

  /* Circuit diagram legend */
  .circuit-label{font-size:.75rem;color:var(--muted);margin-top:8px}

  .err{background:#fef2f2;border:1px solid #fca5a5;border-radius:6px;
    padding:10px 14px;color:#dc2626;font-size:.85rem;margin-top:12px}
</style>
</head>
<body>

<header>
  <div>
    <h1>EIS Viewer</h1>
    <span>Nyquist &amp; Bode · Equivalent circuit fitting · Batch summary</span>
  </div>
</header>

<div class="main">
  <div id="dropzone">
    <p>Drop EIS files here or click to browse</p>
    <small>BioLogic .mpt &nbsp;|&nbsp; Gamry .dta &nbsp;|&nbsp; Zahner .ism &nbsp;|&nbsp; Hioki / Neware CSV &nbsp;|&nbsp; Excel .xlsx</small>
    <input type="file" id="file-input" multiple>
  </div>

  <div id="file-list"></div>

  <div class="controls">
    <button class="btn btn-primary"   id="btn-plot">Plot</button>
    <button class="btn btn-green"     id="btn-fit" style="display:none">Fit Circuits</button>
    <button class="btn btn-secondary" id="btn-clear">Clear all</button>
    <span id="status"></span>
  </div>
  <div id="error-box"></div>

  <div class="plot-grid" id="plot-grid" style="display:none">
    <div class="plot-card" style="grid-column:1/-1">
      <h3>Nyquist Plot — Z′ vs −Z″</h3>
      <div class="circuit-label" id="circuit-label" style="display:none">
        Fitted model: <b>L + R<sub>s</sub> + (R<sub>ct</sub> + W) ∥ CPE</b> &nbsp;(dashed lines)
      </div>
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

  <!-- Batch summary table -->
  <div class="summary-card" id="summary-card" style="display:none">
    <h3>
      Batch Summary
      <button class="btn btn-secondary" id="btn-csv" style="font-size:.8rem;padding:5px 12px">
        Download CSV
      </button>
    </h3>
    <div class="tbl-wrap">
      <table id="summary-table">
        <thead>
          <tr>
            <th>#</th><th>File</th><th>Points</th>
            <th>f max (Hz)</th><th>f min (Hz)</th>
            <th>|Z| @ f_max (Ω)</th><th>|Z| @ f_min (Ω)</th>
            <th>R<sub>s</sub> est. (Ω)</th>
            <th>R<sub>ct</sub> est. (Ω)</th>
            <th>R<sub>s</sub> fit (Ω)</th>
            <th>R<sub>ct</sub> fit (Ω)</th>
            <th>L fit (nH)</th>
            <th>CPE-Q</th><th>CPE-n</th>
            <th>Warburg σ</th>
            <th>Fit error (%)</th>
          </tr>
        </thead>
        <tbody id="summary-body"></tbody>
      </table>
    </div>
  </div>
</div>

<script>
const PALETTE = [
  "#2563eb","#dc2626","#16a34a","#d97706","#7c3aed",
  "#0891b2","#db2777","#65a30d","#ea580c","#0f766e"
];

let fileQueue = [];
let idCounter = 0;
let currentDatasets = [];
let currentFits = {};

const dropzone    = document.getElementById("dropzone");
const fileInput   = document.getElementById("file-input");
const fileList    = document.getElementById("file-list");
const btnPlot     = document.getElementById("btn-plot");
const btnFit      = document.getElementById("btn-fit");
const btnClear    = document.getElementById("btn-clear");
const btnCsv      = document.getElementById("btn-csv");
const status      = document.getElementById("status");
const errorBox    = document.getElementById("error-box");
const plotGrid    = document.getElementById("plot-grid");
const summaryCard = document.getElementById("summary-card");
const summaryBody = document.getElementById("summary-body");
const circuitLabel= document.getElementById("circuit-label");

dropzone.addEventListener("click", () => fileInput.click());
dropzone.addEventListener("dragover",  e => { e.preventDefault(); dropzone.classList.add("over"); });
dropzone.addEventListener("dragleave", ()  => dropzone.classList.remove("over"));
dropzone.addEventListener("drop", e => {
  e.preventDefault(); dropzone.classList.remove("over"); addFiles(e.dataTransfer.files);
});
fileInput.addEventListener("change", () => addFiles(fileInput.files));

function addFiles(files) {
  for (const f of files) {
    const id    = ++idCounter;
    const color = PALETTE[fileQueue.length % PALETTE.length];
    fileQueue.push({ file: f, name: f.name, color, id });
  }
  renderFileList(); fileInput.value = "";
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
  fileQueue = []; currentDatasets = []; currentFits = {};
  renderFileList();
  plotGrid.style.display = summaryCard.style.display = "none";
  btnFit.style.display = "none";
  circuitLabel.style.display = "none";
  errorBox.innerHTML = ""; status.textContent = "";
});

// ── PLOT ───────────────────────────────────────────────────────────────────
btnPlot.addEventListener("click", async () => {
  if (!fileQueue.length) { status.textContent = "No files selected."; return; }
  status.textContent = "Parsing…";
  errorBox.innerHTML = "";
  btnPlot.disabled   = true;

  const datasets = [], errors = [];
  for (const entry of fileQueue) {
    const fd = new FormData();
    fd.append("file", entry.file, entry.name);
    try {
      const res  = await fetch("/parse", { method: "POST", body: fd });
      const json = await res.json();
      if (json.error) { errors.push(`${entry.name}: ${json.error}`); continue; }
      if (!json.points.length) { errors.push(`${entry.name}: no data found`); continue; }
      datasets.push({ name: entry.name, color: entry.color, points: json.points });
    } catch(e) { errors.push(`${entry.name}: network error`); }
  }

  btnPlot.disabled = false;
  if (errors.length) errorBox.innerHTML = `<div class="err">${errors.join("<br>")}</div>`;
  if (!datasets.length) { status.textContent = "No data plotted."; return; }

  currentDatasets = datasets;
  currentFits     = {};
  circuitLabel.style.display = "none";
  status.textContent = `Plotted ${datasets.length} dataset(s).`;
  btnFit.style.display = "";
  buildPlots(datasets, {});
  buildSummaryTable(datasets, {});
});

// ── FIT CIRCUITS ───────────────────────────────────────────────────────────
btnFit.addEventListener("click", async () => {
  if (!currentDatasets.length) return;
  status.textContent = "Fitting circuits…";
  btnFit.disabled    = true;

  for (const ds of currentDatasets) {
    try {
      const res  = await fetch("/fit", {
        method:  "POST",
        headers: { "Content-Type": "application/json" },
        body:    JSON.stringify({ points: ds.points })
      });
      const json = await res.json();
      if (!json.error) currentFits[ds.name] = json;
    } catch(e) {}
  }

  btnFit.disabled = false;
  circuitLabel.style.display = "";
  status.textContent = `Fitted ${Object.keys(currentFits).length} dataset(s).`;
  buildPlots(currentDatasets, currentFits);
  buildSummaryTable(currentDatasets, currentFits);
});

// ── PLOTS ──────────────────────────────────────────────────────────────────
function sortByFreq(pts) { return [...pts].sort((a,b) => b.freq - a.freq); }

const LAYOUT_BASE = {
  margin: {l:56,r:20,t:20,b:50},
  paper_bgcolor:"#fff", plot_bgcolor:"#f8fafc",
  font: {family:"system-ui,sans-serif", size:12},
  legend:{orientation:"h", y:-0.18},
  hovermode:"closest"
};

function hexToRgba(hex, a) {
  const r = parseInt(hex.slice(1,3),16);
  const g = parseInt(hex.slice(3,5),16);
  const b = parseInt(hex.slice(5,7),16);
  return `rgba(${r},${g},${b},${a})`;
}

function buildPlots(datasets, fits) {
  plotGrid.style.display = "";

  // Nyquist
  const nyqTraces = [];
  datasets.forEach(ds => {
    const s = sortByFreq(ds.points);
    nyqTraces.push({
      x: s.map(p => p.zre), y: s.map(p => -p.zim),
      text: s.map(p => `${p.freq.toFixed(3)} Hz`),
      mode:"lines+markers", name: ds.name,
      line:{color:ds.color,width:1.5}, marker:{size:4,color:ds.color},
      hovertemplate:"Z′=%{x:.4g} Ω<br>−Z″=%{y:.4g} Ω<br>%{text}<extra>%{fullData.name}</extra>"
    });
    if (fits[ds.name]) {
      const c = fits[ds.name].curve;
      nyqTraces.push({
        x: c.map(p => p.zre), y: c.map(p => -p.zim),
        mode:"lines", name:`${ds.name} (fit)`, showlegend:false,
        line:{color:ds.color, width:2, dash:"dash"},
        hoverinfo:"skip"
      });
    }
  });
  Plotly.newPlot("plt-nyquist", nyqTraces, {
    ...LAYOUT_BASE,
    xaxis:{title:"Z′ (Ω)", zeroline:true, zerolinecolor:"#94a3b8"},
    yaxis:{title:"−Z″ (Ω)", zeroline:true, zerolinecolor:"#94a3b8", scaleanchor:"x", scaleratio:1}
  }, {responsive:true, displayModeBar:true});

  // Bode magnitude
  const magTraces = datasets.map(ds => {
    const s = sortByFreq(ds.points);
    return {
      x:s.map(p=>p.freq), y:s.map(p=>p.zmag), mode:"lines+markers", name:ds.name,
      line:{color:ds.color,width:1.5}, marker:{size:4},
      hovertemplate:"f=%{x:.3g} Hz<br>|Z|=%{y:.4g} Ω<extra>%{fullData.name}</extra>"
    };
  });
  Plotly.newPlot("plt-bode-mag", magTraces, {
    ...LAYOUT_BASE,
    xaxis:{title:"Frequency (Hz)",type:"log"},
    yaxis:{title:"|Z| (Ω)",type:"log"}
  }, {responsive:true, displayModeBar:true});

  // Bode phase
  const phTraces = datasets.map(ds => {
    const s = sortByFreq(ds.points);
    return {
      x:s.map(p=>p.freq), y:s.map(p=>p.phase_deg), mode:"lines+markers", name:ds.name,
      line:{color:ds.color,width:1.5}, marker:{size:4},
      hovertemplate:"f=%{x:.3g} Hz<br>Phase=%{y:.2f}°<extra>%{fullData.name}</extra>"
    };
  });
  Plotly.newPlot("plt-bode-phase", phTraces, {
    ...LAYOUT_BASE,
    xaxis:{title:"Frequency (Hz)",type:"log"},
    yaxis:{title:"Phase (°)"}
  }, {responsive:true, displayModeBar:true});

  // Z components
  const compTraces = [];
  datasets.forEach(ds => {
    const s = sortByFreq(ds.points);
    compTraces.push({
      x:s.map(p=>p.freq), y:s.map(p=>p.zre),
      mode:"lines+markers", name:`Z′ — ${ds.name}`,
      line:{color:ds.color,width:1.5}, marker:{size:4},
      hovertemplate:"f=%{x:.3g} Hz<br>Z′=%{y:.4g} Ω<extra>%{fullData.name}</extra>"
    });
    compTraces.push({
      x:s.map(p=>p.freq), y:s.map(p=>-p.zim),
      mode:"lines+markers", name:`−Z″ — ${ds.name}`,
      line:{color:ds.color,width:1.5,dash:"dot"}, marker:{size:4,symbol:"triangle-up"},
      hovertemplate:"f=%{x:.3g} Hz<br>−Z″=%{y:.4g} Ω<extra>%{fullData.name}</extra>"
    });
  });
  Plotly.newPlot("plt-components", compTraces, {
    ...LAYOUT_BASE,
    xaxis:{title:"Frequency (Hz)",type:"log"},
    yaxis:{title:"Impedance (Ω)"}
  }, {responsive:true, displayModeBar:true});
}

// ── SUMMARY TABLE ──────────────────────────────────────────────────────────
function buildSummaryTable(datasets, fits) {
  summaryCard.style.display = "";
  summaryBody.innerHTML = "";

  datasets.forEach((ds, i) => {
    const pts  = sortByFreq(ds.points);
    const fmax = pts[0].freq,   fmin = pts[pts.length-1].freq;
    const zmax = pts[0].zmag,   zmin = pts[pts.length-1].zmag;

    // Estimated Rs: Zre at highest frequency
    const Rs_est = pts[0].zre;
    // Estimated Rct: Zre at low-freq zero crossing minus Rs
    const neg_zim = pts.map(p => -p.zim);
    const peak_i  = neg_zim.indexOf(Math.max(...neg_zim));
    const Rct_est = (pts[pts.length-1].zre - Rs_est);

    const fit = fits[ds.name];
    const fmt = (v, d=4) => v == null ? "—" : v.toPrecision(d);
    const qualClass = fit
      ? (fit.quality_pct < 5 ? "qual-good" : fit.quality_pct < 15 ? "qual-ok" : "qual-bad")
      : "";

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${i+1}</td>
      <td><span style="display:inline-block;width:10px;height:10px;border-radius:50%;
          background:${ds.color};margin-right:6px;vertical-align:middle"></span>${ds.name}</td>
      <td>${pts.length}</td>
      <td>${fmax.toFixed(2)}</td>
      <td>${fmin.toFixed(3)}</td>
      <td>${fmt(zmax)}</td>
      <td>${fmt(zmin)}</td>
      <td>${fmt(Rs_est)}</td>
      <td>${fmt(Rct_est > 0 ? Rct_est : null)}</td>
      <td>${fit ? fmt(fit.Rs) : "—"}</td>
      <td>${fit ? fmt(fit.Rct) : "—"}</td>
      <td>${fit ? fmt(fit.L * 1e9, 3) : "—"}</td>
      <td>${fit ? fit.Q.toExponential(2) : "—"}</td>
      <td>${fit ? fit.n.toFixed(3) : "—"}</td>
      <td>${fit ? fit.sigma.toExponential(2) : "—"}</td>
      <td class="${qualClass}">${fit ? fit.quality_pct.toFixed(1)+"%" : "—"}</td>`;
    summaryBody.appendChild(tr);
  });
}

// ── CSV DOWNLOAD ───────────────────────────────────────────────────────────
btnCsv.addEventListener("click", () => {
  if (!currentDatasets.length) return;
  const headers = [
    "File","Points","f_max_Hz","f_min_Hz","|Z|_at_fmax_Ohm","|Z|_at_fmin_Ohm",
    "Rs_est_Ohm","Rct_est_Ohm",
    "Rs_fit_Ohm","Rct_fit_Ohm","L_fit_nH","CPE_Q","CPE_n","Warburg_sigma","Fit_error_pct"
  ];
  const rows = [headers.join(",")];

  currentDatasets.forEach(ds => {
    const pts  = sortByFreq(ds.points);
    const fmax = pts[0].freq, fmin = pts[pts.length-1].freq;
    const zmax = pts[0].zmag, zmin = pts[pts.length-1].zmag;
    const Rs_est  = pts[0].zre;
    const Rct_est = pts[pts.length-1].zre - Rs_est;
    const fit = currentFits[ds.name];

    rows.push([
      `"${ds.name}"`, pts.length,
      fmax.toFixed(3), fmin.toFixed(4),
      zmax.toExponential(4), zmin.toExponential(4),
      Rs_est.toExponential(4), (Rct_est > 0 ? Rct_est : "").toExponential ? Rct_est.toExponential(4) : "",
      fit ? fit.Rs.toExponential(4) : "",
      fit ? fit.Rct.toExponential(4) : "",
      fit ? (fit.L*1e9).toFixed(3) : "",
      fit ? fit.Q.toExponential(4) : "",
      fit ? fit.n.toFixed(4) : "",
      fit ? fit.sigma.toExponential(4) : "",
      fit ? fit.quality_pct.toFixed(2) : ""
    ].join(","));
  });

  const blob = new Blob([rows.join("\n")], {type:"text/csv"});
  const a    = document.createElement("a");
  a.href     = URL.createObjectURL(blob);
  a.download = "EIS_summary.csv";
  a.click();
});
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
    f   = request.files["file"]
    raw = f.read()
    try:
        pts = parse_eis_file(f.filename, raw)
        pts = enrich(pts)
    except Exception as e:
        return jsonify({"error": str(e)}), 200
    if not pts:
        return jsonify({"error": "Could not parse — unrecognised format or no numeric data found."}), 200
    return jsonify({"points": pts})


@app.route("/fit", methods=["POST"])
def fit_route():
    data   = request.get_json(force=True)
    points = data.get("points", [])
    if len(points) < 5:
        return jsonify({"error": "need at least 5 points to fit"}), 200
    try:
        result = fit_randles(points)
    except Exception as e:
        return jsonify({"error": str(e)}), 200
    return jsonify(result)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    url = f"http://localhost:{PORT}"
    threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    print(f"EIS Viewer running at {url}")
    app.run(port=PORT, debug=False)
