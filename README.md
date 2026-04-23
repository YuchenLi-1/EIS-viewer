# EIS Viewer

A local web app for viewing and plotting Electrochemical Impedance Spectroscopy (EIS) data. Select one or more files and instantly generate interactive Nyquist and Bode plots.

## Features

- Drag-and-drop or browse to load EIS files
- Overlay multiple datasets with automatic color coding
- Four interactive plots (Plotly.js):
  - **Nyquist** — Z′ vs −Z″ (equal aspect)
  - **Bode magnitude** — |Z| vs frequency (log-log)
  - **Bode phase** — Phase° vs frequency
  - **Z components** — Z′ and −Z″ vs frequency

## Supported file formats

| Format | Extension |
|--------|-----------|
| Hioki BT4560 | `.csv` |
| BioLogic EC-Lab | `.mpt` |
| Gamry | `.dta` |
| Zahner | `.ism` |
| Neware text export | `.csv` / `.txt` |
| Generic CSV / TSV | any |
| Excel | `.xlsx` |

## Requirements

```
pip install flask openpyxl
```

## Usage

```bash
python eis_viewer.py
```

Opens automatically at http://localhost:5558.
