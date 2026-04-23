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
# Circuit models
# ---------------------------------------------------------------------------

def _prep_eis(points):
    pts   = sorted(points, key=lambda p: p["freq"], reverse=True)
    freqs = np.array([p["freq"] for p in pts])
    Z     = np.array([p["zre"] + 1j * p["zim"] for p in pts])
    omega = 2 * np.pi * freqs
    return freqs, Z, omega

def _quality(Z_fit, Z_meas):
    rms   = float(np.sqrt(np.mean(np.abs(Z_fit - Z_meas)**2)))
    z_rms = float(np.sqrt(np.mean(np.abs(Z_meas)**2)))
    return round(rms / z_rms * 100, 2) if z_rms > 0 else 999.0

def _smooth_curve(model_fn, params, freqs, n=200):
    f  = np.logspace(np.log10(freqs[-1]), np.log10(freqs[0]), n)
    Zc = model_fn(params, 2 * np.pi * f)
    return [{"freq": float(fi), "zre": float(z.real), "zim": float(z.imag)}
            for fi, z in zip(f, Zc)]

def _rs_est(Z, n):
    hf = np.argmin(np.abs(Z.imag[:n//3 + 1]))
    return max(float(Z[hf].real), 1e-9)

def _L_est(Z, omega):
    ind = Z.imag > 0
    if ind.any():
        return max(float(np.median(Z[ind].imag / omega[ind])), 1e-12)
    return 1e-9

def _peak_omega(Z, omega):
    neg = -Z.imag
    i   = int(np.argmax(neg))
    return omega[i] if neg[i] > 0 else float(np.median(omega))

# ── Model: Simple  Rs + Rct∥CPE ─────────────────────────────────────────────

def _simple_Z(params, omega):
    Rs, Rct, Q, n = params
    Z_CPE = 1.0 / (Q * ((1j * omega) ** n))
    Z_par = Rct * Z_CPE / (Rct + Z_CPE)
    return Rs + Z_par

def fit_simple(points):
    freqs, Z, omega = _prep_eis(points)
    Rs0  = _rs_est(Z, len(Z))
    Rct0 = max(float(Z[-1].real) - Rs0, Rs0 * 0.1, 1e-9)
    Q0   = max(1.0 / (_peak_omega(Z, omega) * Rct0), 1e-9)

    def _res(p, w, Zm):
        Zc = _simple_Z(p, w)
        return np.concatenate([Zc.real - Zm.real, Zc.imag - Zm.imag])

    x0     = [Rs0, Rct0, Q0, 0.85]
    bounds = ([0,    0,    1e-12, 0.3],
              [1e3,  1e6,  1e3,   1.0])
    try:
        p = least_squares(_res, x0, args=(omega, Z), bounds=bounds,
                          max_nfev=5000, ftol=1e-10, xtol=1e-10).x
    except Exception:
        p = x0
    Rs, Rct, Q, n = p
    return {"model": "simple",
            "Rs": round(float(Rs), 9), "Rct": round(float(Rct), 9),
            "Q":  round(float(Q), 12), "n":   round(float(n), 4),
            "quality_pct": _quality(_simple_Z(p, omega), Z),
            "curve": _smooth_curve(_simple_Z, p, freqs)}

# ── Model: Randles  L + Rs + (Rct+W)∥CPE ───────────────────────────────────

def _randles_Z(params, omega):
    L, Rs, Rct, Q, n, sigma = params
    jw    = 1j * omega
    Z_W   = sigma * (1.0 - 1j) / np.sqrt(omega)
    Z_CPE = 1.0 / (Q * (jw ** n))
    Z_rct = Rct + Z_W
    Z_par = Z_rct * Z_CPE / (Z_rct + Z_CPE)
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
        "model":  "randles",
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


# ── Model: Two-arc  L + Rs + R1∥CPE1 + (R2+W)∥CPE2 ────────────────────────

def _two_arc_Z(params, omega):
    L, Rs, R1, Q1, n1, R2, Q2, n2, sigma = params
    jw     = 1j * omega
    Z_CPE1 = 1.0 / (Q1 * (jw ** n1))
    Z_par1 = R1 * Z_CPE1 / (R1 + Z_CPE1)
    Z_W    = sigma * (1.0 - 1j) / np.sqrt(omega)
    Z_CPE2 = 1.0 / (Q2 * (jw ** n2))
    Z_rct2 = R2 + Z_W
    Z_par2 = Z_rct2 * Z_CPE2 / (Z_rct2 + Z_CPE2)
    return 1j * omega * L + Rs + Z_par1 + Z_par2

def fit_two_arc(points):
    freqs, Z, omega = _prep_eis(points)
    Rs0    = _rs_est(Z, len(Z))
    L0     = _L_est(Z, omega)
    total  = max(float(Z[-1].real) - Rs0, 1e-9)
    R10    = max(total * 0.3, 1e-9)
    R20    = max(total * 0.7, 1e-9)
    op     = _peak_omega(Z, omega)
    Q10    = max(1.0 / (op * R10 * 10), 1e-9)
    Q20    = max(1.0 / (op * R20), 1e-9)
    sigma0 = max(abs(float(Z[-1].imag)) * math.sqrt(float(omega[-1])) * 0.5, 1e-4)

    def _res(p, w, Zm):
        Zc = _two_arc_Z(p, w)
        return np.concatenate([Zc.real - Zm.real, Zc.imag - Zm.imag])

    x0     = [L0, Rs0, R10, Q10, 0.85, R20, Q20, 0.85, sigma0]
    bounds = ([0,    0,    0,    1e-12, 0.3, 0,    1e-12, 0.3, 0   ],
              [1e-3, 1e3,  1e6,  1e3,   1.0, 1e6,  1e3,   1.0, 1e6 ])
    try:
        p = least_squares(_res, x0, args=(omega, Z), bounds=bounds,
                          max_nfev=8000, ftol=1e-10, xtol=1e-10).x
    except Exception:
        p = x0
    L, Rs, R1, Q1, n1, R2, Q2, n2, sigma = p
    return {
        "model": "two_arc",
        "L":     round(float(L),     12),
        "Rs":    round(float(Rs),    9),
        "R1":    round(float(R1),    9),
        "Q1":    round(float(Q1),    12),
        "n1":    round(float(n1),    4),
        "R2":    round(float(R2),    9),
        "Q2":    round(float(Q2),    12),
        "n2":    round(float(n2),    4),
        "sigma": round(float(sigma), 6),
        "quality_pct": _quality(_two_arc_Z(p, omega), Z),
        "curve": _smooth_curve(_two_arc_Z, p, freqs),
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
/* ── Reset & tokens ─────────────────────────────────────────────────────── */
*{box-sizing:border-box;margin:0;padding:0}
:root{
  --navy:#0d1b2e; --navy2:#162032; --accent:#3b82f6; --accent2:#2563eb;
  --green:#10b981; --amber:#f59e0b; --red:#ef4444;
  --bg:#f1f5f9; --surface:#fff; --surface2:#f8fafc;
  --border:#e2e8f0; --border2:#cbd5e1;
  --text:#0f172a; --muted:#64748b; --muted2:#94a3b8;
  --radius:10px; --shadow:0 1px 3px rgba(0,0,0,.08),0 4px 12px rgba(0,0,0,.04);
  --shadow-lg:0 4px 16px rgba(0,0,0,.12);
}
body{font-family:"Inter",system-ui,sans-serif;background:var(--bg);color:var(--text);
  min-height:100vh;display:flex;flex-direction:column}
button{cursor:pointer}

/* ── App shell ──────────────────────────────────────────────────────────── */
.app{display:flex;flex:1;height:100vh;overflow:hidden}

/* ── Sidebar ────────────────────────────────────────────────────────────── */
.sidebar{width:280px;min-width:280px;background:var(--navy);color:#e2e8f0;
  display:flex;flex-direction:column;height:100vh;overflow:hidden}
.sidebar-header{padding:20px 20px 16px;border-bottom:1px solid rgba(255,255,255,.07)}
.sidebar-logo{display:flex;align-items:center;gap:10px;margin-bottom:4px}
.sidebar-logo svg{flex-shrink:0}
.sidebar-logo h1{font-size:1.1rem;font-weight:700;color:#f8fafc;letter-spacing:.01em}
.sidebar-sub{font-size:.72rem;color:var(--muted2);padding-left:34px}

.sidebar-section{padding:14px 16px 8px;font-size:.68rem;font-weight:700;
  letter-spacing:.1em;text-transform:uppercase;color:var(--muted2)}

/* Drop zone */
.dropzone{margin:0 12px 12px;border:1.5px dashed rgba(59,130,246,.5);
  border-radius:var(--radius);background:rgba(59,130,246,.06);
  padding:20px 12px;text-align:center;cursor:pointer;transition:.2s}
.dropzone:hover,.dropzone.over{border-color:var(--accent);background:rgba(59,130,246,.12)}
.dropzone-icon{font-size:1.6rem;margin-bottom:6px;opacity:.7}
.dropzone p{font-size:.8rem;color:#93c5fd;font-weight:600;margin-bottom:3px}
.dropzone small{font-size:.68rem;color:var(--muted2);line-height:1.4;display:block}
#file-input{display:none}

/* File list */
.file-list{flex:1;overflow-y:auto;padding:0 12px 12px}
.file-list::-webkit-scrollbar{width:4px}
.file-list::-webkit-scrollbar-thumb{background:rgba(255,255,255,.1);border-radius:4px}
.file-item{display:flex;align-items:center;gap:8px;padding:8px 10px;
  border-radius:8px;margin-bottom:4px;background:rgba(255,255,255,.04);
  border:1px solid rgba(255,255,255,.06);transition:.15s;font-size:.8rem}
.file-item:hover{background:rgba(255,255,255,.08)}
.file-dot{width:9px;height:9px;border-radius:50%;flex-shrink:0}
.file-name{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:#cbd5e1}
.file-pts{font-size:.68rem;color:var(--muted2);flex-shrink:0;margin-right:2px}
.file-remove{background:none;border:none;color:var(--muted2);font-size:.9rem;
  padding:2px 4px;border-radius:4px;line-height:1;transition:.15s}
.file-remove:hover{color:#f87171;background:rgba(239,68,68,.15)}
.no-files{text-align:center;color:var(--muted2);font-size:.78rem;padding:20px 0}

/* Sidebar buttons */
.sidebar-actions{padding:12px 16px;border-top:1px solid rgba(255,255,255,.07);display:flex;flex-direction:column;gap:8px}
.btn-sidebar{width:100%;padding:9px 14px;border-radius:8px;border:none;
  font-size:.85rem;font-weight:600;display:flex;align-items:center;
  justify-content:center;gap:8px;transition:.15s}
.btn-plot{background:var(--accent);color:#fff}
.btn-plot:hover:not(:disabled){background:var(--accent2)}
.btn-plot:disabled{opacity:.5;cursor:not-allowed}
.btn-fit{background:rgba(16,185,129,.15);color:#34d399;border:1px solid rgba(16,185,129,.3)}
.btn-fit:hover:not(:disabled){background:rgba(16,185,129,.25)}
.btn-fit:disabled{opacity:.4;cursor:not-allowed}
.btn-clear-s{background:rgba(255,255,255,.06);color:#94a3b8;border:1px solid rgba(255,255,255,.08)}
.btn-clear-s:hover{background:rgba(255,255,255,.1);color:#cbd5e1}
.model-select-wrap{padding:0 16px}
.model-label{display:block;font-size:.72rem;color:#64748b;letter-spacing:.04em;text-transform:uppercase;margin-bottom:5px}
.model-select{width:100%;background:#0a1525;color:#cbd5e1;border:1px solid rgba(255,255,255,.12);
  border-radius:6px;padding:7px 10px;font-size:.8rem;outline:none;cursor:pointer}
.model-select:focus{border-color:var(--accent)}

/* ── Main area ──────────────────────────────────────────────────────────── */
.main{flex:1;display:flex;flex-direction:column;overflow:hidden}

/* Top bar with tabs */
.topbar{background:var(--surface);border-bottom:1px solid var(--border);
  padding:0 24px;display:flex;align-items:center;gap:0;flex-shrink:0}
.tab{padding:14px 18px;font-size:.85rem;font-weight:600;color:var(--muted);
  border:none;background:none;border-bottom:2px solid transparent;
  transition:.15s;white-space:nowrap}
.tab:hover{color:var(--text)}
.tab.active{color:var(--accent2);border-bottom-color:var(--accent2)}
.topbar-right{margin-left:auto;display:flex;align-items:center;gap:10px}
#status-badge{font-size:.78rem;padding:4px 10px;border-radius:20px;
  background:var(--surface2);border:1px solid var(--border);color:var(--muted)}
#status-badge.ok{background:#f0fdf4;border-color:#bbf7d0;color:#15803d}
#status-badge.busy{background:#eff6ff;border-color:#bfdbfe;color:#1d4ed8}
#status-badge.err{background:#fef2f2;border-color:#fecaca;color:#dc2626}
.btn-csv-top{padding:6px 14px;border-radius:6px;font-size:.8rem;font-weight:600;
  background:var(--surface);border:1px solid var(--border);color:var(--text);
  display:none;transition:.15s}
.btn-csv-top:hover{background:var(--surface2)}

/* Content panels */
.content{flex:1;overflow-y:auto;padding:20px 24px}
.content::-webkit-scrollbar{width:6px}
.content::-webkit-scrollbar-thumb{background:var(--border2);border-radius:4px}

.panel{display:none}
.panel.active{display:block}

/* Empty state */
.empty-state{display:flex;flex-direction:column;align-items:center;justify-content:center;
  height:60vh;color:var(--muted2);text-align:center}
.empty-state svg{opacity:.25;margin-bottom:16px}
.empty-state h2{font-size:1.1rem;color:var(--muted);margin-bottom:6px}
.empty-state p{font-size:.85rem;max-width:320px}

/* Plot grid */
.plot-grid{display:grid;grid-template-columns:1fr 1fr;gap:16px}
.plot-card{background:var(--surface);border:1px solid var(--border);
  border-radius:var(--radius);padding:16px 14px 10px;
  box-shadow:var(--shadow)}
.plot-card.wide{grid-column:1/-1}
.plot-title{font-size:.82rem;font-weight:700;color:var(--navy);margin-bottom:4px;
  display:flex;align-items:center;gap:8px}
.plot-title .badge{font-size:.67rem;padding:2px 7px;border-radius:10px;
  background:#eff6ff;color:var(--accent2);font-weight:600}
.plot-wrap{width:100%;height:320px}

/* Fit badge on Nyquist */
.fit-note{font-size:.72rem;color:var(--muted2);margin-bottom:6px}
.fit-note b{color:var(--accent2)}

/* ── Fit results panel ──────────────────────────────────────────────────── */
.fit-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(320px,1fr));gap:16px}
.fit-card{background:var(--surface);border:1px solid var(--border);
  border-radius:var(--radius);padding:18px;box-shadow:var(--shadow)}
.fit-card-header{display:flex;align-items:center;gap:10px;margin-bottom:14px}
.fit-dot{width:12px;height:12px;border-radius:50%;flex-shrink:0}
.fit-card-header h3{font-size:.92rem;font-weight:700;color:var(--text);flex:1;
  overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.fit-quality{font-size:.75rem;font-weight:700;padding:3px 9px;border-radius:10px}
.fq-good{background:#dcfce7;color:#15803d}
.fq-ok{background:#fef3c7;color:#92400e}
.fq-bad{background:#fee2e2;color:#b91c1c}

/* Circuit SVG area */
.circuit-svg-wrap{background:var(--surface2);border:1px solid var(--border);
  border-radius:8px;padding:12px;margin-bottom:14px;text-align:center}
.circuit-svg-wrap svg{max-width:100%}

/* Parameter grid */
.param-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.param-item{background:var(--surface2);border:1px solid var(--border);
  border-radius:7px;padding:9px 11px}
.param-label{font-size:.68rem;color:var(--muted);font-weight:600;
  text-transform:uppercase;letter-spacing:.04em;margin-bottom:3px}
.param-value{font-size:.9rem;font-weight:700;color:var(--text);font-variant-numeric:tabular-nums}
.param-unit{font-size:.72rem;color:var(--muted);font-weight:400;margin-left:3px}
.no-fit-msg{text-align:center;color:var(--muted2);font-size:.85rem;padding:30px}

/* ── Summary panel ──────────────────────────────────────────────────────── */
.summary-wrap{background:var(--surface);border:1px solid var(--border);
  border-radius:var(--radius);box-shadow:var(--shadow);overflow:hidden}
.summary-header{padding:14px 18px;border-bottom:1px solid var(--border);
  display:flex;align-items:center;justify-content:space-between}
.summary-header h3{font-size:.92rem;font-weight:700;color:var(--text)}
.tbl-scroll{overflow-x:auto}
table{width:100%;border-collapse:collapse;font-size:.8rem;white-space:nowrap}
thead th{background:var(--navy);color:#e2e8f0;padding:9px 14px;
  text-align:left;font-weight:600;font-size:.75rem;letter-spacing:.02em}
thead th:first-child{border-radius:0}
tbody td{padding:8px 14px;border-bottom:1px solid var(--border)}
tbody tr:nth-child(even) td{background:var(--surface2)}
tbody tr:hover td{background:#eff6ff}
tbody tr:last-child td{border-bottom:none}
.q-good{color:var(--green);font-weight:700}
.q-ok{color:var(--amber);font-weight:700}
.q-bad{color:var(--red);font-weight:700}
.num{font-variant-numeric:tabular-nums}

/* ── Toast ──────────────────────────────────────────────────────────────── */
#toast-container{position:fixed;bottom:24px;right:24px;display:flex;
  flex-direction:column-reverse;gap:8px;z-index:9999;pointer-events:none}
.toast{padding:10px 16px;border-radius:8px;font-size:.82rem;font-weight:600;
  box-shadow:var(--shadow-lg);pointer-events:auto;
  animation:slideIn .25s ease;max-width:340px;display:flex;align-items:center;gap:8px}
.toast-info{background:var(--navy);color:#e2e8f0}
.toast-ok{background:#052e16;color:#86efac;border:1px solid #166534}
.toast-err{background:#450a0a;color:#fca5a5;border:1px solid #7f1d1d}
@keyframes slideIn{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:none}}

/* ── Spinner ────────────────────────────────────────────────────────────── */
.spinner{width:14px;height:14px;border:2px solid rgba(255,255,255,.3);
  border-top-color:#fff;border-radius:50%;animation:spin .7s linear infinite;flex-shrink:0}
@keyframes spin{to{transform:rotate(360deg)}}

/* ── Responsive ─────────────────────────────────────────────────────────── */
@media(max-width:820px){
  .sidebar{width:240px;min-width:240px}
  .plot-grid{grid-template-columns:1fr}
  .plot-card.wide{grid-column:1}
}
</style>
</head>
<body>
<div class="app">

<!-- ═══ SIDEBAR ══════════════════════════════════════════════════════════ -->
<aside class="sidebar">
  <div class="sidebar-header">
    <div class="sidebar-logo">
      <svg width="24" height="24" viewBox="0 0 24 24" fill="none">
        <path d="M3 12 Q6 4 9 12 Q12 20 15 12 Q18 4 21 12" stroke="#3b82f6" stroke-width="2" stroke-linecap="round" fill="none"/>
        <circle cx="3" cy="12" r="1.5" fill="#3b82f6"/>
        <circle cx="21" cy="12" r="1.5" fill="#3b82f6"/>
      </svg>
      <h1>EIS Viewer</h1>
    </div>
    <div class="sidebar-sub">Impedance Analysis Tool</div>
  </div>

  <div class="sidebar-section">Files</div>

  <div class="dropzone" id="dropzone">
    <div class="dropzone-icon">⊕</div>
    <p>Drop files or click to browse</p>
    <small>.mpt .dta .ism .csv .xlsx</small>
    <input type="file" id="file-input" multiple>
  </div>

  <div class="file-list" id="file-list">
    <div class="no-files" id="no-files-msg">No files loaded</div>
  </div>

  <div class="sidebar-actions">
    <button class="btn-sidebar btn-plot" id="btn-plot">
      <svg width="14" height="14" viewBox="0 0 14 14" fill="none">
        <polygon points="3,2 12,7 3,12" fill="currentColor"/>
      </svg>
      Plot
    </button>
    <div class="model-select-wrap">
      <label class="model-label">Circuit model</label>
      <select id="model-select" class="model-select">
        <option value="simple">Simple — Rs + Rct∥CPE</option>
        <option value="randles" selected>Randles — L+Rs+(Rct+W)∥CPE</option>
        <option value="two_arc">Two-arc — SEI + Rct (9 params)</option>
      </select>
    </div>
    <button class="btn-sidebar btn-fit" id="btn-fit" disabled>
      <svg width="14" height="14" viewBox="0 0 14 14" fill="none">
        <circle cx="7" cy="7" r="5" stroke="currentColor" stroke-width="1.5"/>
        <path d="M4 9 Q5.5 4 7 7 Q8.5 10 10 5" stroke="currentColor" stroke-width="1.5" fill="none" stroke-linecap="round"/>
      </svg>
      Fit Circuits
    </button>
    <button class="btn-sidebar btn-clear-s" id="btn-clear">Clear All</button>
  </div>
</aside>

<!-- ═══ MAIN ══════════════════════════════════════════════════════════════ -->
<div class="main">
  <!-- Top bar / tabs -->
  <div class="topbar">
    <button class="tab active" data-panel="plots">Plots</button>
    <button class="tab" data-panel="fit-results">Fit Results</button>
    <button class="tab" data-panel="summary">Summary</button>
    <div class="topbar-right">
      <span id="status-badge">Ready</span>
      <button class="btn-csv-top" id="btn-csv">↓ Export CSV</button>
    </div>
  </div>

  <!-- Content -->
  <div class="content">

    <!-- ── Plots panel ── -->
    <div class="panel active" id="panel-plots">
      <div id="empty-plots" class="empty-state">
        <svg width="64" height="64" viewBox="0 0 64 64" fill="none">
          <rect x="8" y="40" width="10" height="16" rx="2" fill="#3b82f6"/>
          <rect x="22" y="28" width="10" height="28" rx="2" fill="#3b82f6"/>
          <rect x="36" y="18" width="10" height="38" rx="2" fill="#3b82f6"/>
          <rect x="50" y="32" width="10" height="24" rx="2" fill="#3b82f6"/>
        </svg>
        <h2>No data loaded</h2>
        <p>Drop EIS files in the sidebar and click Plot to visualise impedance spectra.</p>
      </div>
      <div class="plot-grid" id="plot-grid" style="display:none">
        <div class="plot-card wide">
          <div class="plot-title">
            Nyquist Plot
            <span class="badge">Z′ vs −Z″</span>
          </div>
          <div class="fit-note" id="fit-note" style="display:none">
            Dashed lines — fitted model: <b id="fit-note-label"></b>
          </div>
          <div class="plot-wrap" id="plt-nyquist"></div>
        </div>
        <div class="plot-card">
          <div class="plot-title">Bode Magnitude<span class="badge">|Z| vs f</span></div>
          <div class="plot-wrap" id="plt-bode-mag"></div>
        </div>
        <div class="plot-card">
          <div class="plot-title">Bode Phase<span class="badge">θ vs f</span></div>
          <div class="plot-wrap" id="plt-bode-phase"></div>
        </div>
        <div class="plot-card">
          <div class="plot-title">Z Components<span class="badge">Z′ &amp; Z″ vs f</span></div>
          <div class="plot-wrap" id="plt-components"></div>
        </div>
      </div>
    </div>

    <!-- ── Fit results panel ── -->
    <div class="panel" id="panel-fit-results">
      <div id="empty-fit" class="empty-state">
        <svg width="56" height="56" viewBox="0 0 56 56" fill="none">
          <circle cx="28" cy="28" r="20" stroke="#3b82f6" stroke-width="2"/>
          <path d="M14 36 Q18 20 22 28 Q26 36 30 24 Q34 12 42 20" stroke="#3b82f6" stroke-width="2" fill="none" stroke-linecap="round"/>
        </svg>
        <h2>No fit results yet</h2>
        <p>Load files, click Plot, then click Fit Circuits to run equivalent circuit fitting.</p>
      </div>
      <div class="fit-grid" id="fit-grid"></div>
    </div>

    <!-- ── Summary panel ── -->
    <div class="panel" id="panel-summary">
      <div id="empty-summary" class="empty-state">
        <svg width="56" height="56" viewBox="0 0 56 56" fill="none">
          <rect x="8" y="8" width="40" height="40" rx="4" stroke="#3b82f6" stroke-width="2"/>
          <line x1="8" y1="20" x2="48" y2="20" stroke="#3b82f6" stroke-width="1.5"/>
          <line x1="8" y1="32" x2="48" y2="32" stroke="#3b82f6" stroke-width="1"/>
          <line x1="24" y1="8" x2="24" y2="48" stroke="#3b82f6" stroke-width="1"/>
        </svg>
        <h2>No data yet</h2>
        <p>Load and plot files to see the batch summary table.</p>
      </div>
      <div class="summary-wrap" id="summary-wrap" style="display:none">
        <div class="summary-header">
          <h3>Batch Summary</h3>
        </div>
        <div class="tbl-scroll">
          <table>
            <thead>
              <tr>
                <th>#</th><th>File</th><th>Model</th><th>Points</th>
                <th>f max (Hz)</th><th>f min (Hz)</th>
                <th>|Z| @ f_max (Ω)</th><th>|Z| @ f_min (Ω)</th>
                <th>R<sub>s</sub> est.</th><th>R<sub>ct</sub> est.</th>
                <th>R<sub>s</sub> fit</th><th>R1 fit (SEI)</th><th>R<sub>ct</sub> fit</th>
                <th>L fit (nH)</th><th>CPE-Q</th><th>CPE-n</th>
                <th>Warburg σ</th><th>Fit error</th>
              </tr>
            </thead>
            <tbody id="summary-body"></tbody>
          </table>
        </div>
      </div>
    </div>

  </div><!-- /content -->
</div><!-- /main -->
</div><!-- /app -->

<div id="toast-container"></div>

<script>
const PALETTE = [
  "#3b82f6","#ef4444","#10b981","#f59e0b","#8b5cf6",
  "#06b6d4","#ec4899","#84cc16","#f97316","#14b8a6"
];

let fileQueue = [], idCounter = 0;
let currentDatasets = [], currentFits = {};

// ── DOM refs ──────────────────────────────────────────────────────────────
const $ = id => document.getElementById(id);
const dropzone    = $("dropzone");
const fileInput   = $("file-input");
const fileList    = $("file-list");
const noFilesMsg  = $("no-files-msg");
const btnPlot     = $("btn-plot");
const btnFit      = $("btn-fit");
const btnClear    = $("btn-clear");
const btnCsv      = $("btn-csv");
const statusBadge = $("status-badge");
const plotGrid    = $("plot-grid");
const emptyPlots  = $("empty-plots");
const fitGrid     = $("fit-grid");
const emptyFit    = $("empty-fit");
const summaryWrap = $("summary-wrap");
const emptySummary= $("empty-summary");
const summaryBody = $("summary-body");
const fitNote     = $("fit-note");

// ── Tabs ──────────────────────────────────────────────────────────────────
document.querySelectorAll(".tab").forEach(tab => {
  tab.addEventListener("click", () => {
    document.querySelectorAll(".tab").forEach(t => t.classList.remove("active"));
    document.querySelectorAll(".panel").forEach(p => p.classList.remove("active"));
    tab.classList.add("active");
    $("panel-" + tab.dataset.panel).classList.add("active");
  });
});

// ── Toast ─────────────────────────────────────────────────────────────────
function toast(msg, type="info", duration=3000) {
  const tc = $("toast-container");
  const t  = document.createElement("div");
  t.className = `toast toast-${type}`;
  const icon = type==="ok" ? "✓" : type==="err" ? "✕" : "ℹ";
  t.innerHTML = `<span>${icon}</span><span>${msg}</span>`;
  tc.appendChild(t);
  setTimeout(() => { t.style.opacity="0"; t.style.transition="opacity .3s";
    setTimeout(()=>t.remove(),300); }, duration);
}

function setStatus(msg, type="") {
  statusBadge.textContent = msg;
  statusBadge.className = "ok busy err".includes(type) ? type : "";
  statusBadge.className = type ? type : "";
}

// ── Drop & file management ────────────────────────────────────────────────
dropzone.addEventListener("click", () => fileInput.click());
dropzone.addEventListener("dragover",  e => { e.preventDefault(); dropzone.classList.add("over"); });
dropzone.addEventListener("dragleave", ()  => dropzone.classList.remove("over"));
dropzone.addEventListener("drop", e => {
  e.preventDefault(); dropzone.classList.remove("over");
  addFiles(e.dataTransfer.files);
});
fileInput.addEventListener("change", () => { addFiles(fileInput.files); fileInput.value=""; });

function addFiles(files) {
  for (const f of files) {
    if (fileQueue.find(q => q.name === f.name)) continue; // dedupe
    fileQueue.push({ file:f, name:f.name, color:PALETTE[fileQueue.length % PALETTE.length], id:++idCounter });
  }
  renderFileList();
}

function removeFile(id) {
  fileQueue = fileQueue.filter(f => f.id !== id);
  renderFileList();
}

function renderFileList() {
  // Remove all chips
  fileList.querySelectorAll(".file-item").forEach(e => e.remove());
  noFilesMsg.style.display = fileQueue.length ? "none" : "";
  fileQueue.forEach(f => {
    const el = document.createElement("div");
    el.className = "file-item";
    el.innerHTML = `<span class="file-dot" style="background:${f.color}"></span>
      <span class="file-name" title="${f.name}">${f.name}</span>
      <span class="file-pts" id="pts-${f.id}"></span>
      <button class="file-remove" onclick="removeFile(${f.id})" title="Remove">✕</button>`;
    fileList.appendChild(el);
  });
}

// ── PLOT ──────────────────────────────────────────────────────────────────
btnPlot.addEventListener("click", async () => {
  if (!fileQueue.length) { toast("No files selected", "err"); return; }

  btnPlot.innerHTML = '<span class="spinner"></span> Parsing…';
  btnPlot.disabled  = true;
  setStatus("Parsing…", "busy");

  const datasets = [], errors = [];
  for (const entry of fileQueue) {
    const fd = new FormData();
    fd.append("file", entry.file, entry.name);
    try {
      const res  = await fetch("/parse", { method:"POST", body:fd });
      const json = await res.json();
      if (json.error) { errors.push(`${entry.name}: ${json.error}`); continue; }
      datasets.push({ name:entry.name, color:entry.color, points:json.points });
      const el = $("pts-" + entry.id);
      if (el) el.textContent = json.points.length + " pts";
    } catch(e) { errors.push(`${entry.name}: network error`); }
  }

  btnPlot.innerHTML = `<svg width="14" height="14" viewBox="0 0 14 14"><polygon points="3,2 12,7 3,12" fill="currentColor"/></svg> Plot`;
  btnPlot.disabled  = false;

  if (errors.length) errors.forEach(e => toast(e, "err", 5000));
  if (!datasets.length) { setStatus("No data", "err"); return; }

  currentDatasets = datasets;
  currentFits     = {};
  fitNote.style.display = "none";
  btnFit.disabled = false;
  btnCsv.style.display  = "";
  setStatus(`${datasets.length} file${datasets.length>1?"s":""} plotted`, "ok");
  toast(`Plotted ${datasets.length} dataset${datasets.length>1?"s":""}`, "ok");

  buildPlots(datasets, {});
  buildSummaryTable(datasets, {});

  // Switch to plots tab
  document.querySelector('[data-panel="plots"]').click();
});

// ── FIT ───────────────────────────────────────────────────────────────────
const modelSelect = $("model-select");

btnFit.addEventListener("click", async () => {
  if (!currentDatasets.length) return;
  const model = modelSelect.value;
  btnFit.innerHTML = '<span class="spinner"></span> Fitting…';
  btnFit.disabled  = true;
  setStatus("Fitting circuits…", "busy");

  for (const ds of currentDatasets) {
    try {
      const res  = await fetch("/fit", {
        method:"POST", headers:{"Content-Type":"application/json"},
        body: JSON.stringify({ points: ds.points, model })
      });
      const json = await res.json();
      if (!json.error) currentFits[ds.name] = json;
    } catch(e) {}
  }

  const n = Object.keys(currentFits).length;
  btnFit.innerHTML = `<svg width="14" height="14" viewBox="0 0 14 14"><circle cx="7" cy="7" r="5" stroke="currentColor" stroke-width="1.5"/><path d="M4 9 Q5.5 4 7 7 Q8.5 10 10 5" stroke="currentColor" stroke-width="1.5" fill="none" stroke-linecap="round"/></svg> Fit Circuits`;
  btnFit.disabled  = false;
  fitNote.style.display = "";
  setStatus(`${n} circuit${n>1?"s":""} fitted`, "ok");
  toast(`Fitted ${n} equivalent circuit${n>1?"s":""}`, "ok");

  buildPlots(currentDatasets, currentFits);
  buildFitCards(currentDatasets, currentFits);
  buildSummaryTable(currentDatasets, currentFits);

  document.querySelector('[data-panel="fit-results"]').click();
});

// ── Clear ─────────────────────────────────────────────────────────────────
btnClear.addEventListener("click", () => {
  fileQueue = []; currentDatasets = []; currentFits = {};
  renderFileList();
  plotGrid.style.display = "none"; emptyPlots.style.display = "";
  fitGrid.innerHTML = ""; emptyFit.style.display = "";
  summaryWrap.style.display = "none"; emptySummary.style.display = "";
  fitNote.style.display = "none";
  btnFit.disabled = true; btnCsv.style.display = "none";
  setStatus("Ready");
  document.querySelector('[data-panel="plots"]').click();
});

// ── Plot building ─────────────────────────────────────────────────────────
function sortByFreq(pts) { return [...pts].sort((a,b) => b.freq - a.freq); }

const LAY = {
  margin:{l:52,r:16,t:12,b:46},
  paper_bgcolor:"#fff", plot_bgcolor:"#fafbff",
  font:{family:"Inter,system-ui,sans-serif",size:11.5},
  legend:{orientation:"h",y:-0.22,font:{size:10.5}},
  hovermode:"closest",
  xaxis:{gridcolor:"#f1f5f9",zeroline:false,linecolor:"#e2e8f0",tickfont:{size:10.5}},
  yaxis:{gridcolor:"#f1f5f9",zeroline:false,linecolor:"#e2e8f0",tickfont:{size:10.5}}
};

function buildPlots(datasets, fits) {
  emptyPlots.style.display = "none";
  plotGrid.style.display   = "";

  // Nyquist
  const nyqT = [];
  datasets.forEach(ds => {
    const s = sortByFreq(ds.points);
    nyqT.push({ x:s.map(p=>p.zre), y:s.map(p=>-p.zim),
      text:s.map(p=>`${p.freq.toFixed(3)} Hz`),
      mode:"lines+markers", name:ds.name,
      line:{color:ds.color,width:2}, marker:{size:4.5,color:ds.color},
      hovertemplate:"Z′ = %{x:.4g} Ω<br>−Z″ = %{y:.4g} Ω<br>%{text}<extra>%{fullData.name}</extra>"
    });
    if (fits[ds.name]) {
      const c = fits[ds.name].curve;
      nyqT.push({ x:c.map(p=>p.zre), y:c.map(p=>-p.zim), mode:"lines",
        name:`${ds.name} fit`, showlegend:false,
        line:{color:ds.color,width:2,dash:"dash"}, hoverinfo:"skip" });
    }
  });
  Plotly.newPlot("plt-nyquist", nyqT, {
    ...LAY,
    xaxis:{...LAY.xaxis,title:"Z′ (Ω)"},
    yaxis:{...LAY.yaxis,title:"−Z″ (Ω)",scaleanchor:"x",scaleratio:1}
  }, {responsive:true,displayModeBar:true,displaylogo:false});

  // Bode |Z|
  Plotly.newPlot("plt-bode-mag", datasets.map(ds => {
    const s = sortByFreq(ds.points);
    return { x:s.map(p=>p.freq), y:s.map(p=>p.zmag), mode:"lines+markers", name:ds.name,
      line:{color:ds.color,width:2}, marker:{size:4.5},
      hovertemplate:"f = %{x:.3g} Hz<br>|Z| = %{y:.4g} Ω<extra>%{fullData.name}</extra>" };
  }), {...LAY, xaxis:{...LAY.xaxis,title:"Frequency (Hz)",type:"log"},
    yaxis:{...LAY.yaxis,title:"|Z| (Ω)",type:"log"}
  }, {responsive:true,displayModeBar:true,displaylogo:false});

  // Bode phase
  Plotly.newPlot("plt-bode-phase", datasets.map(ds => {
    const s = sortByFreq(ds.points);
    return { x:s.map(p=>p.freq), y:s.map(p=>p.phase_deg), mode:"lines+markers", name:ds.name,
      line:{color:ds.color,width:2}, marker:{size:4.5},
      hovertemplate:"f = %{x:.3g} Hz<br>θ = %{y:.2f}°<extra>%{fullData.name}</extra>" };
  }), {...LAY, xaxis:{...LAY.xaxis,title:"Frequency (Hz)",type:"log"},
    yaxis:{...LAY.yaxis,title:"Phase (°)"}
  }, {responsive:true,displayModeBar:true,displaylogo:false});

  // Z components
  const compT = [];
  datasets.forEach(ds => {
    const s = sortByFreq(ds.points);
    compT.push({ x:s.map(p=>p.freq), y:s.map(p=>p.zre), mode:"lines+markers",
      name:`Z′ ${ds.name}`, line:{color:ds.color,width:2}, marker:{size:4.5},
      hovertemplate:"f = %{x:.3g} Hz<br>Z′ = %{y:.4g} Ω<extra>%{fullData.name}</extra>" });
    compT.push({ x:s.map(p=>p.freq), y:s.map(p=>-p.zim), mode:"lines+markers",
      name:`−Z″ ${ds.name}`, line:{color:ds.color,width:2,dash:"dot"},
      marker:{size:4.5,symbol:"triangle-up"},
      hovertemplate:"f = %{x:.3g} Hz<br>−Z″ = %{y:.4g} Ω<extra>%{fullData.name}</extra>" });
  });
  Plotly.newPlot("plt-components", compT, {
    ...LAY, xaxis:{...LAY.xaxis,title:"Frequency (Hz)",type:"log"},
    yaxis:{...LAY.yaxis,title:"Impedance (Ω)"}
  }, {responsive:true,displayModeBar:true,displaylogo:false});
}

// ── Fit result cards ──────────────────────────────────────────────────────
function fmtSI(v, unit) {
  if (v == null) return "—";
  const abs = Math.abs(v);
  if (abs >= 1)    return v.toFixed(4) + `<span class="param-unit">${unit}</span>`;
  if (abs >= 1e-3) return (v*1e3).toFixed(4) + `<span class="param-unit">m${unit}</span>`;
  if (abs >= 1e-6) return (v*1e6).toFixed(4) + `<span class="param-unit">µ${unit}</span>`;
  return v.toExponential(3) + `<span class="param-unit">${unit}</span>`;
}

const MODEL_LABELS = {
  simple:  "Rs + Rct∥CPE",
  randles: "L + Rs + (Rct+W)∥CPE",
  two_arc: "L + Rs + R₁∥CPE₁ + (R₂+W)∥CPE₂",
};

const SVG_CIRCUITS = {
  simple: `<svg viewBox="0 0 260 80" width="260" height="75" font-family="Inter,system-ui" font-size="11">
    <line x1="0" y1="40" x2="30" y2="40" stroke="#475569" stroke-width="1.5"/>
    <rect x="30" y="32" width="36" height="16" rx="3" fill="#f8fafc" stroke="#10b981" stroke-width="1.8"/>
    <text x="48" y="44" text-anchor="middle" fill="#10b981" font-weight="600">Rs</text>
    <line x1="66" y1="40" x2="82" y2="40" stroke="#475569" stroke-width="1.5"/>
    <circle cx="82" cy="40" r="3" fill="#475569"/>
    <line x1="82" y1="40" x2="82" y2="15" stroke="#475569" stroke-width="1.5"/>
    <line x1="82" y1="15" x2="110" y2="15" stroke="#475569" stroke-width="1.5"/>
    <rect x="110" y="7" width="36" height="16" rx="3" fill="#f8fafc" stroke="#ef4444" stroke-width="1.8"/>
    <text x="128" y="19" text-anchor="middle" fill="#ef4444" font-weight="600">Rct</text>
    <line x1="146" y1="15" x2="178" y2="15" stroke="#475569" stroke-width="1.5"/>
    <line x1="178" y1="15" x2="178" y2="40" stroke="#475569" stroke-width="1.5"/>
    <line x1="82" y1="40" x2="82" y2="65" stroke="#475569" stroke-width="1.5"/>
    <line x1="82" y1="65" x2="124" y2="65" stroke="#475569" stroke-width="1.5"/>
    <line x1="124" y1="57" x2="124" y2="73" stroke="#8b5cf6" stroke-width="2.5"/>
    <line x1="129" y1="57" x2="129" y2="73" stroke="#8b5cf6" stroke-width="2.5"/>
    <text x="148" y="69" fill="#8b5cf6" font-weight="600">CPE</text>
    <line x1="168" y1="65" x2="178" y2="65" stroke="#475569" stroke-width="1.5"/>
    <line x1="178" y1="65" x2="178" y2="40" stroke="#475569" stroke-width="1.5"/>
    <line x1="178" y1="40" x2="260" y2="40" stroke="#475569" stroke-width="1.5"/>
  </svg>`,

  randles: `<svg viewBox="0 0 340 80" width="300" height="75" font-family="Inter,system-ui" font-size="11">
    <line x1="0" y1="40" x2="30" y2="40" stroke="#475569" stroke-width="1.5"/>
    <path d="M30,40 q5,-9 10,0 q5,-9 10,0 q5,-9 10,0" fill="none" stroke="#3b82f6" stroke-width="1.8"/>
    <text x="42" y="26" text-anchor="middle" fill="#3b82f6" font-weight="600">L</text>
    <line x1="60" y1="40" x2="70" y2="40" stroke="#475569" stroke-width="1.5"/>
    <rect x="70" y="32" width="36" height="16" rx="3" fill="#f8fafc" stroke="#10b981" stroke-width="1.8"/>
    <text x="88" y="44" text-anchor="middle" fill="#10b981" font-weight="600">Rs</text>
    <line x1="106" y1="40" x2="128" y2="40" stroke="#475569" stroke-width="1.5"/>
    <circle cx="128" cy="40" r="3" fill="#475569"/>
    <line x1="128" y1="40" x2="128" y2="15" stroke="#475569" stroke-width="1.5"/>
    <line x1="128" y1="15" x2="155" y2="15" stroke="#475569" stroke-width="1.5"/>
    <rect x="155" y="7" width="36" height="16" rx="3" fill="#f8fafc" stroke="#ef4444" stroke-width="1.8"/>
    <text x="173" y="19" text-anchor="middle" fill="#ef4444" font-weight="600">Rct</text>
    <line x1="191" y1="15" x2="210" y2="15" stroke="#475569" stroke-width="1.5"/>
    <path d="M210,15 l5,8 l5,-8 l5,8 l5,-8" fill="none" stroke="#ef4444" stroke-width="1.8" stroke-linecap="round"/>
    <text x="223" y="7" text-anchor="middle" fill="#ef4444" font-weight="600" font-size="9">W</text>
    <line x1="230" y1="15" x2="262" y2="15" stroke="#475569" stroke-width="1.5"/>
    <line x1="262" y1="15" x2="262" y2="40" stroke="#475569" stroke-width="1.5"/>
    <line x1="128" y1="40" x2="128" y2="65" stroke="#475569" stroke-width="1.5"/>
    <line x1="128" y1="65" x2="178" y2="65" stroke="#475569" stroke-width="1.5"/>
    <line x1="178" y1="57" x2="178" y2="73" stroke="#8b5cf6" stroke-width="2.5"/>
    <line x1="183" y1="57" x2="183" y2="73" stroke="#8b5cf6" stroke-width="2.5"/>
    <text x="197" y="69" fill="#8b5cf6" font-weight="600">CPE</text>
    <line x1="216" y1="65" x2="262" y2="65" stroke="#475569" stroke-width="1.5"/>
    <line x1="262" y1="65" x2="262" y2="40" stroke="#475569" stroke-width="1.5"/>
    <line x1="262" y1="40" x2="310" y2="40" stroke="#475569" stroke-width="1.5"/>
  </svg>`,

  two_arc: `<svg viewBox="0 0 480 90" width="420" height="80" font-family="Inter,system-ui" font-size="10">
    <line x1="0" y1="44" x2="22" y2="44" stroke="#475569" stroke-width="1.5"/>
    <path d="M22,44 q4,-8 8,0 q4,-8 8,0 q4,-8 8,0" fill="none" stroke="#3b82f6" stroke-width="1.8"/>
    <text x="32" y="30" text-anchor="middle" fill="#3b82f6" font-weight="600">L</text>
    <line x1="46" y1="44" x2="54" y2="44" stroke="#475569" stroke-width="1.5"/>
    <rect x="54" y="36" width="30" height="14" rx="3" fill="#f8fafc" stroke="#10b981" stroke-width="1.8"/>
    <text x="69" y="47" text-anchor="middle" fill="#10b981" font-weight="600">Rs</text>
    <line x1="84" y1="44" x2="100" y2="44" stroke="#475569" stroke-width="1.5"/>
    <circle cx="100" cy="44" r="2.5" fill="#475569"/>
    <!-- Arc 1: R1∥CPE1 (SEI) -->
    <line x1="100" y1="44" x2="100" y2="20" stroke="#475569" stroke-width="1.5"/>
    <line x1="100" y1="20" x2="118" y2="20" stroke="#475569" stroke-width="1.5"/>
    <rect x="118" y="13" width="28" height="13" rx="2" fill="#f8fafc" stroke="#f59e0b" stroke-width="1.8"/>
    <text x="132" y="23" text-anchor="middle" fill="#f59e0b" font-weight="600">R1</text>
    <line x1="146" y1="20" x2="168" y2="20" stroke="#475569" stroke-width="1.5"/>
    <line x1="168" y1="20" x2="168" y2="44" stroke="#475569" stroke-width="1.5"/>
    <line x1="100" y1="44" x2="100" y2="68" stroke="#475569" stroke-width="1.5"/>
    <line x1="100" y1="68" x2="128" y2="68" stroke="#475569" stroke-width="1.5"/>
    <line x1="128" y1="62" x2="128" y2="74" stroke="#8b5cf6" stroke-width="2.2"/>
    <line x1="132" y1="62" x2="132" y2="74" stroke="#8b5cf6" stroke-width="2.2"/>
    <text x="144" y="72" fill="#8b5cf6" font-weight="600">CPE1</text>
    <line x1="162" y1="68" x2="168" y2="68" stroke="#475569" stroke-width="1.5"/>
    <line x1="168" y1="68" x2="168" y2="44" stroke="#475569" stroke-width="1.5"/>
    <line x1="168" y1="44" x2="190" y2="44" stroke="#475569" stroke-width="1.5"/>
    <circle cx="190" cy="44" r="2.5" fill="#475569"/>
    <!-- Arc 2: (R2+W)∥CPE2 (Rct) -->
    <line x1="190" y1="44" x2="190" y2="20" stroke="#475569" stroke-width="1.5"/>
    <line x1="190" y1="20" x2="208" y2="20" stroke="#475569" stroke-width="1.5"/>
    <rect x="208" y="13" width="28" height="13" rx="2" fill="#f8fafc" stroke="#ef4444" stroke-width="1.8"/>
    <text x="222" y="23" text-anchor="middle" fill="#ef4444" font-weight="600">R2</text>
    <line x1="236" y1="20" x2="252" y2="20" stroke="#475569" stroke-width="1.5"/>
    <path d="M252,20 l4,6 l4,-6 l4,6 l4,-6" fill="none" stroke="#ef4444" stroke-width="1.6" stroke-linecap="round"/>
    <text x="264" y="12" text-anchor="middle" fill="#ef4444" font-weight="600" font-size="8">W</text>
    <line x1="268" y1="20" x2="290" y2="20" stroke="#475569" stroke-width="1.5"/>
    <line x1="290" y1="20" x2="290" y2="44" stroke="#475569" stroke-width="1.5"/>
    <line x1="190" y1="44" x2="190" y2="68" stroke="#475569" stroke-width="1.5"/>
    <line x1="190" y1="68" x2="218" y2="68" stroke="#475569" stroke-width="1.5"/>
    <line x1="218" y1="62" x2="218" y2="74" stroke="#8b5cf6" stroke-width="2.2"/>
    <line x1="222" y1="62" x2="222" y2="74" stroke="#8b5cf6" stroke-width="2.2"/>
    <text x="236" y="72" fill="#8b5cf6" font-weight="600">CPE2</text>
    <line x1="254" y1="68" x2="290" y2="68" stroke="#475569" stroke-width="1.5"/>
    <line x1="290" y1="68" x2="290" y2="44" stroke="#475569" stroke-width="1.5"/>
    <line x1="290" y1="44" x2="480" y2="44" stroke="#475569" stroke-width="1.5"/>
  </svg>`,
};

function paramsHtml(fit) {
  if (!fit) return `<div class="no-fit-msg">Fit failed for this dataset.</div>`;
  const m = fit.model;
  if (m === "simple") return `
    <div class="param-grid">
      <div class="param-item"><div class="param-label">R<sub>s</sub> — Series</div><div class="param-value">${fmtSI(fit.Rs,"Ω")}</div></div>
      <div class="param-item"><div class="param-label">R<sub>ct</sub> — Charge Transfer</div><div class="param-value">${fmtSI(fit.Rct,"Ω")}</div></div>
      <div class="param-item"><div class="param-label">CPE — Q</div><div class="param-value">${fit.Q.toExponential(3)}</div></div>
      <div class="param-item"><div class="param-label">CPE — n</div><div class="param-value">${fit.n.toFixed(4)}<span class="param-unit">(0–1)</span></div></div>
    </div>`;
  if (m === "randles") return `
    <div class="param-grid">
      <div class="param-item"><div class="param-label">R<sub>s</sub> — Series</div><div class="param-value">${fmtSI(fit.Rs,"Ω")}</div></div>
      <div class="param-item"><div class="param-label">R<sub>ct</sub> — Charge Transfer</div><div class="param-value">${fmtSI(fit.Rct,"Ω")}</div></div>
      <div class="param-item"><div class="param-label">L — Inductance</div><div class="param-value">${fmtSI(fit.L,"H")}</div></div>
      <div class="param-item"><div class="param-label">Warburg σ</div><div class="param-value">${fit.sigma.toExponential(3)}<span class="param-unit">Ω·s⁻⁰·⁵</span></div></div>
      <div class="param-item"><div class="param-label">CPE — Q</div><div class="param-value">${fit.Q.toExponential(3)}</div></div>
      <div class="param-item"><div class="param-label">CPE — n</div><div class="param-value">${fit.n.toFixed(4)}<span class="param-unit">(0–1)</span></div></div>
    </div>`;
  if (m === "two_arc") return `
    <div class="param-grid">
      <div class="param-item"><div class="param-label">R<sub>s</sub> — Series</div><div class="param-value">${fmtSI(fit.Rs,"Ω")}</div></div>
      <div class="param-item"><div class="param-label">L — Inductance</div><div class="param-value">${fmtSI(fit.L,"H")}</div></div>
      <div class="param-item"><div class="param-label">R<sub>1</sub> — SEI</div><div class="param-value">${fmtSI(fit.R1,"Ω")}</div></div>
      <div class="param-item"><div class="param-label">CPE1 — Q</div><div class="param-value">${fit.Q1.toExponential(3)}</div></div>
      <div class="param-item"><div class="param-label">CPE1 — n</div><div class="param-value">${fit.n1.toFixed(4)}<span class="param-unit">(0–1)</span></div></div>
      <div class="param-item"><div class="param-label">R<sub>2</sub> — Charge Transfer</div><div class="param-value">${fmtSI(fit.R2,"Ω")}</div></div>
      <div class="param-item"><div class="param-label">CPE2 — Q</div><div class="param-value">${fit.Q2.toExponential(3)}</div></div>
      <div class="param-item"><div class="param-label">CPE2 — n</div><div class="param-value">${fit.n2.toFixed(4)}<span class="param-unit">(0–1)</span></div></div>
      <div class="param-item"><div class="param-label">Warburg σ</div><div class="param-value">${fit.sigma.toExponential(3)}<span class="param-unit">Ω·s⁻⁰·⁵</span></div></div>
    </div>`;
  return "";
}

function buildFitCards(datasets, fits) {
  fitGrid.innerHTML = "";
  const hasFits = Object.keys(fits).length > 0;
  emptyFit.style.display = hasFits ? "none" : "";
  if (!hasFits) return;

  // Update fit note label to reflect the model used
  const anyFit = Object.values(fits)[0];
  if (anyFit) {
    const lbl = $("fit-note-label");
    if (lbl) lbl.textContent = MODEL_LABELS[anyFit.model] || anyFit.model;
  }

  datasets.forEach(ds => {
    const fit = fits[ds.name];
    const card = document.createElement("div");
    card.className = "fit-card";

    const qc  = fit ? (fit.quality_pct < 5 ? "fq-good" : fit.quality_pct < 15 ? "fq-ok" : "fq-bad") : "";
    const qTxt= fit ? fit.quality_pct.toFixed(1)+"%" : "—";

    card.innerHTML = `
      <div class="fit-card-header">
        <span class="fit-dot" style="background:${ds.color}"></span>
        <h3 title="${ds.name}">${ds.name}</h3>
        ${fit ? `<span class="fit-quality ${qc}">${qTxt} error</span>` : ""}
      </div>
      ${fit ? `<div class="circuit-svg-wrap">${SVG_CIRCUITS[fit.model] || ""}</div>${paramsHtml(fit)}`
             : `<div class="no-fit-msg">Fit failed for this dataset.</div>`}`;
    fitGrid.appendChild(card);
  });
}

// ── Summary table ─────────────────────────────────────────────────────────
function buildSummaryTable(datasets, fits) {
  emptySummary.style.display = "none";
  summaryWrap.style.display  = "";
  summaryBody.innerHTML = "";

  datasets.forEach((ds, i) => {
    const pts  = sortByFreq(ds.points);
    const fmax = pts[0].freq, fmin = pts[pts.length-1].freq;
    const Rs_est  = pts[0].zre;
    const Rct_est = pts[pts.length-1].zre - Rs_est;
    const fit = fits[ds.name];
    const fmt = (v, d=4) => v==null ? "—" : Number(v).toPrecision(d);
    const qc  = fit ? (fit.quality_pct<5?"q-good":fit.quality_pct<15?"q-ok":"q-bad") : "";

    // Model-aware fit value extraction
    const modelName = fit ? ({simple:"Simple",randles:"Randles",two_arc:"Two-arc"}[fit.model]||fit.model) : "—";
    const f_Rs    = fit ? fmt(fit.Rs) : "—";
    const f_R1    = fit ? (fit.model==="two_arc" ? fmt(fit.R1) : "—") : "—";
    const f_Rct   = fit ? (fit.model==="two_arc" ? fmt(fit.R2) : fmt(fit.Rct)) : "—";
    const f_L     = fit ? (fit.L!=null ? fmt(fit.L*1e9,3) : "—") : "—";
    const f_Q     = fit ? (fit.model==="two_arc" ? fit.Q2.toExponential(2) : fit.Q.toExponential(2)) : "—";
    const f_n     = fit ? (fit.model==="two_arc" ? fit.n2.toFixed(3) : fit.n.toFixed(3)) : "—";
    const f_sigma = fit ? (fit.sigma!=null ? fit.sigma.toExponential(2) : "—") : "—";

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td class="num">${i+1}</td>
      <td><span style="display:inline-block;width:9px;height:9px;border-radius:50%;
          background:${ds.color};margin-right:7px;vertical-align:middle"></span>${ds.name}</td>
      <td>${modelName}</td>
      <td class="num">${pts.length}</td>
      <td class="num">${fmax.toFixed(2)}</td>
      <td class="num">${fmin.toFixed(3)}</td>
      <td class="num">${pts[0].zmag.toExponential(3)}</td>
      <td class="num">${pts[pts.length-1].zmag.toExponential(3)}</td>
      <td class="num">${Rs_est.toExponential(3)}</td>
      <td class="num">${Rct_est>0?Rct_est.toExponential(3):"—"}</td>
      <td class="num">${f_Rs}</td>
      <td class="num">${f_R1}</td>
      <td class="num">${f_Rct}</td>
      <td class="num">${f_L}</td>
      <td class="num">${f_Q}</td>
      <td class="num">${f_n}</td>
      <td class="num">${f_sigma}</td>
      <td class="num ${qc}">${fit?fit.quality_pct.toFixed(1)+"%":"—"}</td>`;
    summaryBody.appendChild(tr);
  });
}

// ── CSV export ────────────────────────────────────────────────────────────
btnCsv.addEventListener("click", () => {
  if (!currentDatasets.length) return;
  const hdr = ["File","Model","Points","f_max_Hz","f_min_Hz","|Z|_fmax_Ohm","|Z|_fmin_Ohm",
    "Rs_est_Ohm","Rct_est_Ohm","Rs_fit_Ohm","R1_SEI_fit_Ohm","Rct_fit_Ohm","L_fit_nH",
    "CPE_Q","CPE_n","Warburg_sigma","Fit_error_pct"];
  const rows = [hdr.join(",")];
  currentDatasets.forEach(ds => {
    const pts = sortByFreq(ds.points);
    const Rs  = pts[0].zre, Rct = pts[pts.length-1].zre - Rs;
    const fit = currentFits[ds.name];
    const m   = fit ? fit.model : "";
    rows.push([
      `"${ds.name}"`, m,
      pts.length,
      pts[0].freq.toFixed(3), pts[pts.length-1].freq.toFixed(4),
      pts[0].zmag.toExponential(4), pts[pts.length-1].zmag.toExponential(4),
      Rs.toExponential(4), Rct>0?Rct.toExponential(4):"",
      fit?fit.Rs.toExponential(4):"",
      fit&&m==="two_arc"?fit.R1.toExponential(4):"",
      fit?(m==="two_arc"?fit.R2:fit.Rct).toExponential(4):"",
      fit&&fit.L!=null?(fit.L*1e9).toFixed(3):"",
      fit?(m==="two_arc"?fit.Q2:fit.Q).toExponential(4):"",
      fit?(m==="two_arc"?fit.n2:fit.n).toFixed(4):"",
      fit&&fit.sigma!=null?fit.sigma.toExponential(4):"",
      fit?fit.quality_pct.toFixed(2):""
    ].join(","));
  });
  const a = Object.assign(document.createElement("a"),{
    href: URL.createObjectURL(new Blob([rows.join("\n")],{type:"text/csv"})),
    download: "EIS_summary.csv"
  });
  a.click();
  toast("CSV exported", "ok");
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
    model  = data.get("model", "randles")
    if len(points) < 5:
        return jsonify({"error": "need at least 5 points to fit"}), 200
    try:
        fn = {"simple": fit_simple, "randles": fit_randles, "two_arc": fit_two_arc}.get(model, fit_randles)
        result = fn(points)
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
