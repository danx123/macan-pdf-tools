# Macan PDF Tools

A lightweight, standalone PDF utility built with **PySide6** and **Pillow** — designed specifically for **low-spec machines and older CPUs without AVX/AVX2 support** (e.g. early Core i3/i5, Celeron, Atom, and some older AMD chips).

Extracted from the PDF Tools page of [Macan Converter](https://github.com/danx123) and turned into its own focused, dependency-light application.

![Python](https://img.shields.io/badge/python-3.9%2B-blue)
![License](https://img.shields.io/badge/license-MIT-green)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20Linux%20%7C%20macOS-lightgrey)

---

## Why this exists

Most PDF/image toolkits pull in **numpy** and **opencv-python**, whose official wheels are commonly compiled with AVX2 instructions. On older or budget CPUs that lack AVX support, this causes a hard crash (`Illegal instruction`) the moment those libraries are imported.

Macan PDF Tools avoids that entirely:

- **No numpy, no opencv** — all image processing goes through **Pillow** only.
- PDF page rendering uses **pypdfium2** (a PDFium binding with no AVX requirement).
- PDF structural editing/compression uses **pikepdf** (optional, with an automatic fallback to plain pypdfium2 merging if it isn't installed).

The result: a tool that runs reliably on hardware that other PDF apps simply won't start on.

---

## Features

Four focused tools, each with its own page and options sidebar:

| Tool | Description |
|---|---|
| **Image to PDF** | Combine images into a single PDF or one PDF per image, with adjustable quality and target file size |
| **PDF to Image** | Export PDF pages to PNG / JPG / WEBP with selectable DPI and page ranges (e.g. `1,3,5-8`) |
| **PDF Merger** | Merge multiple PDFs with optional image recompression and target size shrinking |
| **PDF Document Conversion** | Convert PDF text content to TXT, DOCX, or XLSX |

Additional UI features:

- 🌐 **Bilingual interface** — Indonesian and English, switchable at runtime
- 🖼️ **Real thumbnails** — actual image previews and rendered first-page PDF previews (not generic placeholder icons), generated asynchronously so the UI never freezes
- 💾 **Persistent settings** — language, window size, and window position are remembered between sessions (via `QSettings`)
- 🖱️ **Drag-and-drop** batch file input
- 🎨 Dark UI with transparent label backgrounds

---

## Screenshots

<img width="1365" height="767" alt="Screenshot 2026-06-30 182333" src="https://github.com/user-attachments/assets/739ef1a3-6eba-45a4-a241-eeb12db133c2" />


---

## Installation

### Requirements

- Python 3.9+
- pip

### Install dependencies

```bash
pip install PySide6 Pillow pypdfium2 pikepdf python-docx openpyxl --break-system-packages
```

> `pikepdf`, `python-docx`, and `openpyxl` are **optional**. If any of them are missing, the related feature (PDF compression in the merger, DOCX export, or XLSX export) is automatically disabled with a clear message — the rest of the app keeps working normally.

### Run

```bash
git clone https://github.com/danx123/macan-pdf-tools.git
cd macan-pdf-tools
python macan_pdf_tools_standalone.py
```

---

## Optional: app logo

Drop a `logo.png` file in the same folder as the script (or the compiled `.exe`) to have it shown in the footer panel. If the file isn't present, the app simply hides that area — no error.

---

## Building a standalone executable

This app is designed to be compiled with [Nuitka](https://nuitka.net/) for native performance, matching the rest of the Macan suite:

```bash
nuitka --standalone --onefile --enable-plugin=pyside6 ^
    --include-data-files=logo.png=logo.png ^
    macan_pdf_tools_standalone.py
```

(Adjust flags for your platform; the `^` line continuation is for Windows `cmd` — use `\` on Linux/macOS.)

---

## Tech stack

- [PySide6](https://doc.qt.io/qtforpython/) — UI framework
- [Pillow](https://python-pillow.org/) — image processing
- [pypdfium2](https://github.com/pypdfium2-team/pypdfium2) — PDF rendering
- [pikepdf](https://github.com/pikepdf/pikepdf) — PDF structural editing (optional)
- [python-docx](https://python-docx.readthedocs.io/) — DOCX export (optional)
- [openpyxl](https://openpyxl.readthedocs.io/) — XLSX export (optional)

---

## License

MIT — see [LICENSE](LICENSE) for details.

---

## Author

Maintained by [danx123](https://github.com/danx123), part of the **Macan / MacanAngkasa** suite of desktop applications.
