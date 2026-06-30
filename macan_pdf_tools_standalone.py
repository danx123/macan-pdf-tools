#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Macan PDF Tools — Standalone Edition
=====================================
Diekstrak dari macan_converter (PDF Tools page) menjadi aplikasi mandiri.

Didesain khusus untuk PC/laptop spek rendah & CPU lama yang TIDAK mendukung
instruksi AVX/AVX2 (mis. prosesor generasi awal Core i3/i5, Celeron, Atom,
beberapa AMD lawas). Karena itu:

  - TIDAK menggunakan numpy / opencv (wheel resmi opencv-python & numpy versi
    baru sering dikompilasi dengan AVX2 dan akan crash "Illegal instruction"
    di CPU non-AVX).
  - Semua pemrosesan gambar memakai Pillow (PIL) murni.
  - Rendering halaman PDF memakai pypdfium2 (binding ke PDFium, tidak butuh AVX).
  - Manipulasi struktur PDF (merge & kompresi stream) memakai pikepdf (opsional,
    fallback otomatis ke pypdfium2 murni kalau pikepdf tidak terpasang).

Fitur (4 sub-tools, sama seperti page "PDF Tools" di Macan Converter):
  1. Image to PDF       — gabungkan gambar jadi satu/banyak PDF
  2. PDF to Image       — ekspor halaman PDF ke PNG/JPG/WEBP
  3. PDF Merger         — gabungkan banyak PDF + kompresi opsional
  4. PDF Document Conversion — PDF -> TXT / PDF -> DOCX / PDF -> XLSX

Mendukung 2 bahasa: Bahasa Indonesia & English (pilih dari pojok kanan atas).

Dependencies (semua ringan, tanpa AVX requirement):
    pip install PySide6 Pillow pypdfium2 pikepdf python-docx openpyxl --break-system-packages

Catatan: pikepdf, python-docx, openpyxl bersifat opsional — fitur terkait
akan otomatis nonaktif (dengan pesan jelas) jika library tidak terpasang,
aplikasi tetap berjalan normal untuk fitur lainnya.
"""

import sys
import os
import io

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QPushButton, QListWidget, QListWidgetItem, QAbstractItemView,
    QFileDialog, QLineEdit, QComboBox, QSpinBox, QFrame, QProgressBar,
    QMessageBox, QStackedWidget, QSplitter
)
from PySide6.QtCore import Qt, QSize, QThread, QObject, Signal, Slot
from PySide6.QtGui import QIcon, QPixmap, QDragEnterEvent, QDragMoveEvent, QDropEvent

from PIL import Image

try:
    import pypdfium2 as pdfium
    HAS_PDFIUM = True
except ImportError:
    HAS_PDFIUM = False

try:
    import pikepdf
    HAS_PIKEPDF = True
except ImportError:
    HAS_PIKEPDF = False

try:
    import docx as _docx_check  # python-docx
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

try:
    import openpyxl as _openpyxl_check
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False


APP_VERSION = "1.1.0"
IMAGE_EXT = ['.png', '.jpg', '.jpeg', '.bmp', '.webp', '.gif']
PDF_EXT = ['.pdf']


# ──────────────────────────────────────────────────────────────────────────
#  Bahasa / Language strings
# ──────────────────────────────────────────────────────────────────────────
LANGUAGES = {
    "id": {
        "window_title": "Macan PDF Tools — Standalone (v{version})",
        "nav_img2pdf": "Image to PDF",
        "nav_pdf2img": "PDF to Image",
        "nav_merger": "PDF Merger",
        "nav_docconv": "Konversi Dokumen PDF",
        "lang_label": "Bahasa:",
        "add_files_btn": "Tambah File",
        "clear_btn": "Bersihkan",
        "output_label": "Output:",
        "output_placeholder": "Pilih folder output...",
        "choose_folder_btn": "Pilih Folder",
        "choose_folder_title": "Pilih Folder Output",
        "status_ready": "Siap.",
        "status_stopped": "Dihentikan oleh pengguna.",
        "start_btn": "Mulai",
        "stop_btn": "Stop",
        "options_label": "Opsi",
        "drop_placeholder": "Seret file ke sini atau klik 'Tambah File'",
        "error_title": "Error",
        "done_title": "Selesai",
        "error_no_output": "Pilih folder output terlebih dahulu.",
        "dep_missing_title": "Dependency hilang",
        "dep_missing_body": "Beberapa fitur tidak akan berfungsi:\n\n{items}\n\nInstall dengan:\npip install {pkgs} --break-system-packages",

        # Image to PDF
        "img2pdf_title": "Image to PDF",
        "img2pdf_start_btn": "Gabungkan -> PDF",
        "img2pdf_no_files": "Tambahkan file gambar terlebih dahulu.",
        "quality_label": "Kualitas:",
        "qualities": ["Maksimum (100)", "Bagus (95)", "Baik (85)", "Sedang (75)", "Rendah (50)"],
        "out_mode_label": "Output:",
        "out_modes": ["Satu PDF untuk semua gambar", "Satu PDF per gambar"],
        "target_size_label": "Target Ukuran:",
        "target_size_hint": "0 = otomatis",

        # PDF to Image
        "pdf2img_title": "PDF to Image",
        "pdf2img_start_btn": "Konversi -> Gambar",
        "pdf2img_no_files": "Tambahkan file PDF terlebih dahulu.",
        "format_label": "Format:",
        "dpi_label": "Kualitas (DPI):",
        "dpis": ["Rendah (72 DPI)", "Sedang (150 DPI)", "Baik (200 DPI)", "Tinggi (300 DPI)", "Maksimum (600 DPI)"],
        "pages_label": "Halaman:",
        "pages_hint": "Cth: 1,3,5-8 (kosong = semua)",

        # Merger
        "merger_title": "PDF Merger",
        "merger_start_btn": "Gabungkan PDF",
        "merger_no_files": "Tambahkan minimal 2 file PDF.",
        "merger_qualities": ["Pertahankan asli", "Optimal (85)", "Rendah (60)"],
        "merger_pikepdf_note": "Catatan: 'pikepdf' tidak terpasang — kompresi tidak\ntersedia, hanya merge struktural (via pypdfium2).",

        # Doc conversion
        "docconv_start_btn": "Konversi",
        "docconv_no_files": "Tambahkan file PDF terlebih dahulu.",
        "convert_to_label": "Konversi ke:",
        "missing_libs_label": "Belum terpasang: {items}",
    },
    "en": {
        "window_title": "Macan PDF Tools — Standalone (v{version})",
        "nav_img2pdf": "Image to PDF",
        "nav_pdf2img": "PDF to Image",
        "nav_merger": "PDF Merger",
        "nav_docconv": "PDF Document Conversion",
        "lang_label": "Language:",
        "add_files_btn": "Add Files",
        "clear_btn": "Clear",
        "output_label": "Output:",
        "output_placeholder": "Select output folder...",
        "choose_folder_btn": "Choose Folder",
        "choose_folder_title": "Select Output Folder",
        "status_ready": "Ready.",
        "status_stopped": "Stopped by user.",
        "start_btn": "Start",
        "stop_btn": "Stop",
        "options_label": "Options",
        "drop_placeholder": "Drag files here or click 'Add Files'",
        "error_title": "Error",
        "done_title": "Done",
        "error_no_output": "Please select an output folder first.",
        "dep_missing_title": "Missing dependency",
        "dep_missing_body": "Some features will not work:\n\n{items}\n\nInstall with:\npip install {pkgs} --break-system-packages",

        # Image to PDF
        "img2pdf_title": "Image to PDF",
        "img2pdf_start_btn": "Combine -> PDF",
        "img2pdf_no_files": "Please add image files first.",
        "quality_label": "Quality:",
        "qualities": ["Maximum (100)", "Good (95)", "Good (85)", "Medium (75)", "Low (50)"],
        "out_mode_label": "Output:",
        "out_modes": ["Single PDF for all images", "One PDF per image"],
        "target_size_label": "Target Size:",
        "target_size_hint": "0 = automatic",

        # PDF to Image
        "pdf2img_title": "PDF to Image",
        "pdf2img_start_btn": "Convert -> Images",
        "pdf2img_no_files": "Please add PDF files first.",
        "format_label": "Format:",
        "dpi_label": "Quality (DPI):",
        "dpis": ["Low (72 DPI)", "Medium (150 DPI)", "Good (200 DPI)", "High (300 DPI)", "Maximum (600 DPI)"],
        "pages_label": "Pages:",
        "pages_hint": "E.g.: 1,3,5-8 (blank = all)",

        # Merger
        "merger_title": "PDF Merger",
        "merger_start_btn": "Merge PDFs",
        "merger_no_files": "Please add at least 2 PDF files.",
        "merger_qualities": ["Keep original", "Optimized (85)", "Low (60)"],
        "merger_pikepdf_note": "Note: 'pikepdf' is not installed — compression is\nunavailable, only structural merge (via pypdfium2).",

        # Doc conversion
        "docconv_start_btn": "Convert",
        "docconv_no_files": "Please add PDF files first.",
        "convert_to_label": "Convert to:",
        "missing_libs_label": "Not installed: {items}",
    },
}


# ──────────────────────────────────────────────────────────────────────────
#  Simple icon helpers (no numpy/opencv — generic colored glyph via Pillow)
# ──────────────────────────────────────────────────────────────────────────
def make_generic_icon(ext, size=96):
    """Bikin icon kotak sederhana berdasarkan ekstensi file, full PIL (ringan)."""
    ext = ext.lower().lstrip('.')
    colors = {
        'pdf': (200, 70, 60), 'png': (90, 150, 200), 'jpg': (90, 150, 200),
        'jpeg': (90, 150, 200), 'bmp': (90, 150, 200), 'webp': (90, 150, 200),
        'gif': (90, 150, 200),
    }
    color = colors.get(ext, (120, 120, 120))
    img = Image.new('RGB', (size, size), (45, 45, 45))
    pad = 10
    box = Image.new('RGB', (size - pad * 2, size - pad * 2), color)
    img.paste(box, (pad, pad))
    buf = io.BytesIO()
    img.save(buf, format='PNG')
    pix = QPixmap()
    pix.loadFromData(buf.getvalue(), 'PNG')
    return pix


# ──────────────────────────────────────────────────────────────────────────
#  Drag & drop file list (no thumbnail worker thread pool needed — lightweight)
# ──────────────────────────────────────────────────────────────────────────
class FileDropArea(QListWidget):
    files_changed = Signal()

    def __init__(self, accept_types=None, lang=None, parent=None):
        super().__init__(parent)
        self.accept_types = accept_types or ['image', 'pdf']
        self.file_paths = []
        self.lang = lang

        self.setAcceptDrops(True)
        self.setDragDropMode(QAbstractItemView.DragDropMode.NoDragDrop)
        self.setViewMode(QListWidget.ViewMode.IconMode)
        self.setIconSize(QSize(96, 96))
        self.setResizeMode(QListWidget.ResizeMode.Adjust)
        self.setWordWrap(True)
        self.setSpacing(10)
        self._icon_cache = {}
        self._placeholder_item = None
        self._set_placeholder()

    def _set_placeholder(self):
        text = self.lang["drop_placeholder"] if self.lang else "Drag files here or click 'Add Files'"
        self._placeholder_item = QListWidgetItem(text)
        self._placeholder_item.setFlags(Qt.ItemFlag.NoItemFlags)
        self._placeholder_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.addItem(self._placeholder_item)

    def retranslate(self, lang):
        self.lang = lang
        if self.count() == 1 and self.item(0) == self._placeholder_item:
            self._placeholder_item.setText(lang["drop_placeholder"])

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event: QDragMoveEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls():
            paths = [url.toLocalFile() for url in event.mimeData().urls()]
            self.add_files(paths)
            event.acceptProposedAction()
        else:
            event.ignore()

    def _allowed_ext(self):
        allowed = []
        if 'image' in self.accept_types:
            allowed.extend(IMAGE_EXT)
        if 'pdf' in self.accept_types:
            allowed.extend(PDF_EXT)
        return allowed

    def add_files(self, paths):
        if self.count() == 1 and self.item(0) == self._placeholder_item:
            self.clear()
            self.file_paths.clear()

        allowed = self._allowed_ext()
        valid = [p for p in paths if os.path.splitext(p)[1].lower() in allowed and os.path.isfile(p)]

        for path in valid:
            if path in self.file_paths:
                continue
            self.file_paths.append(path)
            ext = os.path.splitext(path)[1].lower().lstrip('.')
            if ext not in self._icon_cache:
                self._icon_cache[ext] = make_generic_icon(ext)
            item = QListWidgetItem(QIcon(self._icon_cache[ext]), os.path.basename(path))
            item.setData(Qt.ItemDataRole.UserRole, path)
            item.setSizeHint(QSize(120, 130))
            item.setTextAlignment(Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignHCenter)
            self.addItem(item)

        if valid:
            self.files_changed.emit()

    def clear_files(self):
        self.clear()
        self.file_paths.clear()
        self._set_placeholder()
        self.files_changed.emit()

    def get_all_file_paths(self):
        paths = []
        for i in range(self.count()):
            item = self.item(i)
            if item != self._placeholder_item:
                paths.append(item.data(Qt.ItemDataRole.UserRole))
        return paths


# ──────────────────────────────────────────────────────────────────────────
#  Worker — semua logika konversi (pure PIL / pypdfium2 / pikepdf)
# ──────────────────────────────────────────────────────────────────────────
class PdfToolsWorker(QObject):
    progress_updated = Signal(int, str)
    finished = Signal(str)
    error = Signal(str)

    def __init__(self, mode, input_paths, output_path, **kwargs):
        super().__init__()
        self.mode = mode
        self.input_paths = input_paths
        self.output_path = output_path
        self.kwargs = kwargs
        self.is_running = True

    def stop(self):
        self.is_running = False

    @staticmethod
    def _parse_page_selection(sel_str, total_pages):
        if not sel_str.strip():
            return list(range(total_pages))
        pages = set()
        for part in sel_str.split(','):
            part = part.strip()
            if '-' in part:
                a, b = part.split('-', 1)
                try:
                    pages.update(range(int(a) - 1, int(b)))
                except ValueError:
                    pass
            else:
                try:
                    pages.add(int(part) - 1)
                except ValueError:
                    pass
        valid = sorted(p for p in pages if 0 <= p < total_pages)
        return valid if valid else list(range(total_pages))

    @staticmethod
    def _quality_to_jpeg(quality_idx):
        return [100, 95, 85, 75, 50][min(quality_idx, 4)]

    @staticmethod
    def _dpi_from_index(dpi_idx):
        return [72, 150, 200, 300, 600][min(dpi_idx, 4)]

    def run(self):
        try:
            if self.mode == 'img2pdf_single':
                self._images_to_single_pdf()
            elif self.mode == 'img2pdf_multi':
                self._images_to_multi_pdf()
            elif self.mode == 'pdf2img':
                self._pdf_to_images()
            elif self.mode == 'merge':
                self._merge_pdfs()
            elif self.mode == 'pdf2txt':
                self._pdf_to_txt()
            elif self.mode == 'pdf2docx':
                self._pdf_to_docx()
            elif self.mode == 'pdf2xlsx':
                self._pdf_to_xlsx()
        except Exception as e:
            if self.is_running:
                self.error.emit(f"Error: {e}")

    # ---- Image -> PDF ----
    def _images_to_single_pdf(self):
        quality = self._quality_to_jpeg(self.kwargs.get('quality_idx', 2))
        target_mb = self.kwargs.get('target_mb', 0.0)
        output_name = self.kwargs.get('output_name', 'output.pdf')
        output_file = os.path.join(self.output_path, output_name)

        total = len(self.input_paths)
        pil_images = []
        for i, path in enumerate(self.input_paths):
            if not self.is_running:
                return
            self.progress_updated.emit(int((i / total) * 85), f"Processing image {i+1}/{total}...")
            pil_images.append(Image.open(path).convert('RGB'))

        if not pil_images:
            self.error.emit("No images.")
            return

        self.progress_updated.emit(90, "Saving PDF...")
        first, rest = pil_images[0], pil_images[1:]
        save_kwargs = {'format': 'PDF', 'save_all': True, 'append_images': rest,
                        'quality': max(20, quality - 10) if target_mb > 0 else quality}
        first.save(output_file, **save_kwargs)

        if target_mb > 0:
            actual_mb = os.path.getsize(output_file) / (1024 * 1024)
            if actual_mb > target_mb:
                q2 = max(20, int(quality * (target_mb / actual_mb) * 0.9))
                compressed = []
                for img in pil_images:
                    buf = io.BytesIO()
                    img.save(buf, format='JPEG', quality=q2)
                    buf.seek(0)
                    compressed.append(Image.open(buf).convert('RGB'))
                compressed[0].save(output_file, format='PDF', save_all=True,
                                    append_images=compressed[1:], quality=q2)

        self.progress_updated.emit(100, "Done!")
        self.finished.emit(f"OK|PDF saved: {os.path.basename(output_file)}")

    def _images_to_multi_pdf(self):
        quality = self._quality_to_jpeg(self.kwargs.get('quality_idx', 2))
        target_mb = self.kwargs.get('target_mb', 0.0)
        total = len(self.input_paths)
        for i, path in enumerate(self.input_paths):
            if not self.is_running:
                return
            self.progress_updated.emit(int((i / total) * 100), f"Processing image {i+1}/{total}...")
            base = os.path.splitext(os.path.basename(path))[0]
            out_file = os.path.join(self.output_path, f"{base}.pdf")
            img = Image.open(path).convert('RGB')
            q = quality
            if target_mb > 0:
                buf = io.BytesIO()
                img.save(buf, format='PDF', quality=quality)
                actual_mb = buf.tell() / (1024 * 1024)
                if actual_mb > target_mb:
                    q = max(20, int(quality * (target_mb / actual_mb) * 0.85))
            img.save(out_file, format='PDF', quality=q)

        self.progress_updated.emit(100, "Done!")
        self.finished.emit(f"OK|{total} PDF files created.")

    # ---- PDF -> Image ----
    def _pdf_to_images(self):
        if not HAS_PDFIUM:
            self.error.emit("'pypdfium2' is not installed.")
            return
        fmt = self.kwargs.get('fmt', 'PNG').lower()
        dpi = self._dpi_from_index(self.kwargs.get('dpi_idx', 2))
        page_sel_str = self.kwargs.get('page_sel', '')
        scale = dpi / 72.0

        total_exported = 0
        for pdf_path in self.input_paths:
            if not self.is_running:
                return
            pdf = pdfium.PdfDocument(pdf_path)
            total_pages = len(pdf)
            pages = self._parse_page_selection(page_sel_str, total_pages)
            base = os.path.splitext(os.path.basename(pdf_path))[0]
            for j, page_idx in enumerate(pages):
                if not self.is_running:
                    return
                self.progress_updated.emit(
                    int((j / max(len(pages), 1)) * 100),
                    f"Converting page {j+1}/{len(pages)}...")
                pil_img = pdf[page_idx].render(scale=scale).to_pil()
                out_file = os.path.join(self.output_path, f"{base}_p{page_idx+1}.{fmt}")
                if fmt in ('jpg', 'jpeg', 'webp'):
                    pil_img.convert('RGB').save(out_file)
                else:
                    pil_img.save(out_file)
                total_exported += 1

        self.progress_updated.emit(100, "Done!")
        self.finished.emit(f"OK|{total_exported} images exported.")

    # ---- PDF Merger ----
    @staticmethod
    def _recompress_images_in_pdf(pdf_pikepdf, jpeg_quality):
        count = 0
        for page in pdf_pikepdf.pages:
            resources = page.get("/Resources")
            if resources is None:
                continue
            xobjects = resources.get("/XObject")
            if xobjects is None:
                continue
            for name in list(xobjects.keys()):
                xobj = xobjects[name]
                try:
                    if xobj.get("/Subtype") != "/Image":
                        continue
                    w = int(xobj["/Width"]); h = int(xobj["/Height"])
                    if w * h < 4096:
                        continue
                    cs = xobj.get("/ColorSpace")
                    if cs in ("/DeviceGray", "/DeviceCMYK"):
                        continue
                    raw = bytes(xobj.read_raw_bytes())
                    img = Image.open(io.BytesIO(raw))
                    if img.mode not in ("RGB", "L"):
                        img = img.convert("RGB")
                    buf = io.BytesIO()
                    img.save(buf, "JPEG", quality=jpeg_quality, optimize=True)
                    buf.seek(0)
                    xobj.write(buf.read(), filter=pikepdf.Name("/DCTDecode"), decode_parms=None)
                    xobj["/BitsPerComponent"] = 8
                    count += 1
                except Exception:
                    pass
        return count

    def _merge_pdfs(self):
        quality_idx = self.kwargs.get('quality_idx', 0)
        target_mb = self.kwargs.get('target_mb', 0.0)
        output_name = self.kwargs.get('output_name', 'merged.pdf')
        output_file = os.path.join(self.output_path, output_name)

        if not HAS_PIKEPDF:
            self._merge_pdfs_fallback(output_file)
            return

        jpeg_quality_map = {0: None, 1: 72, 2: 45}
        jpeg_quality = jpeg_quality_map.get(quality_idx, None)

        total_files = len(self.input_paths)
        self.progress_updated.emit(2, "Opening files...")
        try:
            merged = pikepdf.Pdf.new()
            for i, path in enumerate(self.input_paths):
                if not self.is_running:
                    return
                src = pikepdf.Pdf.open(path)
                merged.pages.extend(src.pages)
                self.progress_updated.emit(int(((i + 1) / total_files) * 50), f"Merging {i+1}/{total_files}...")
        except Exception as e:
            self.error.emit(f"Merge error: {e}")
            return

        if not self.is_running:
            return

        self.progress_updated.emit(55, "Optimizing...")
        try:
            merged.remove_unreferenced_resources()
        except Exception:
            pass

        if jpeg_quality is not None:
            self.progress_updated.emit(60, "Recompressing images...")
            try:
                self._recompress_images_in_pdf(merged, jpeg_quality)
            except Exception:
                pass

        self.progress_updated.emit(85, "Saving...")
        try:
            save_opts = dict(compress_streams=True,
                              object_stream_mode=pikepdf.ObjectStreamMode.generate,
                              recompress_flate=True)
            merged.save(output_file, **save_opts)

            if target_mb > 0:
                actual_mb = os.path.getsize(output_file) / (1024 * 1024)
                iteration = 0
                q = jpeg_quality if jpeg_quality else 65
                while actual_mb > target_mb and iteration < 3 and self.is_running:
                    iteration += 1
                    q = max(10, int(q * (target_mb / actual_mb) * 0.88))
                    self.progress_updated.emit(85 + iteration * 4, f"Shrinking (q={q})...")
                    try:
                        self._recompress_images_in_pdf(merged, q)
                        merged.save(output_file, **save_opts)
                        actual_mb = os.path.getsize(output_file) / (1024 * 1024)
                    except Exception:
                        break
        except Exception as e:
            self.error.emit(f"Save error: {e}")
            return

        self.progress_updated.emit(100, "Done!")
        self.finished.emit(f"OK|Merged PDF saved: {os.path.basename(output_file)}")

    def _merge_pdfs_fallback(self, output_file):
        if not HAS_PDFIUM:
            self.error.emit("'pypdfium2' or 'pikepdf' is not installed.")
            return
        merged = pdfium.PdfDocument.new()
        docs, total_pages = [], 0
        for path in self.input_paths:
            doc = pdfium.PdfDocument(path)
            docs.append(doc)
            total_pages += len(doc)
        done = 0
        for doc in docs:
            if not self.is_running:
                return
            merged.import_pages(doc, list(range(len(doc))))
            done += len(doc)
            self.progress_updated.emit(int((done / max(total_pages, 1)) * 95), "Merging...")
        merged.save(output_file)
        self.progress_updated.emit(100, "Done!")
        self.finished.emit(f"OK|Merged PDF saved: {os.path.basename(output_file)}")

    # ---- PDF Document Conversion ----
    def _pdf_to_txt(self):
        if not HAS_PDFIUM:
            self.error.emit("'pypdfium2' is not installed.")
            return
        total = len(self.input_paths)
        done_files = []
        for i, pdf_path in enumerate(self.input_paths):
            if not self.is_running:
                return
            self.progress_updated.emit(int((i / total) * 95), f"Extracting text {i+1}/{total}...")
            pdf = pdfium.PdfDocument(pdf_path)
            text_parts = []
            for page in pdf:
                textpage = page.get_textpage()
                text_parts.append(textpage.get_text_range())
            base = os.path.splitext(os.path.basename(pdf_path))[0]
            out_file = os.path.join(self.output_path, f"{base}.txt")
            with open(out_file, 'w', encoding='utf-8') as f:
                f.write("\n\n".join(text_parts))
            done_files.append(out_file)
        self.progress_updated.emit(100, "Done!")
        self.finished.emit(f"OK|{len(done_files)} TXT files created.")

    def _pdf_to_docx(self):
        if not HAS_PDFIUM:
            self.error.emit("'pypdfium2' is not installed.")
            return
        if not HAS_DOCX:
            self.error.emit("'python-docx' is not installed. Install: pip install python-docx --break-system-packages")
            return
        import docx
        total = len(self.input_paths)
        done_files = []
        for i, pdf_path in enumerate(self.input_paths):
            if not self.is_running:
                return
            self.progress_updated.emit(int((i / total) * 95), f"Converting {i+1}/{total}...")
            pdf = pdfium.PdfDocument(pdf_path)
            document = docx.Document()
            for p_idx, page in enumerate(pdf):
                textpage = page.get_textpage()
                text = textpage.get_text_range()
                for line in text.split('\n'):
                    if line.strip():
                        document.add_paragraph(line)
                if p_idx < len(pdf) - 1:
                    document.add_page_break()
            base = os.path.splitext(os.path.basename(pdf_path))[0]
            out_file = os.path.join(self.output_path, f"{base}.docx")
            document.save(out_file)
            done_files.append(out_file)
        self.progress_updated.emit(100, "Done!")
        self.finished.emit(f"OK|{len(done_files)} DOCX files created.")

    def _pdf_to_xlsx(self):
        if not HAS_PDFIUM:
            self.error.emit("'pypdfium2' is not installed.")
            return
        if not HAS_OPENPYXL:
            self.error.emit("'openpyxl' is not installed. Install: pip install openpyxl --break-system-packages")
            return
        import openpyxl
        total = len(self.input_paths)
        done_files = []
        for i, pdf_path in enumerate(self.input_paths):
            if not self.is_running:
                return
            self.progress_updated.emit(int((i / total) * 95), f"Converting {i+1}/{total}...")
            pdf = pdfium.PdfDocument(pdf_path)
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "PDF Text"
            row = 1
            for page in pdf:
                textpage = page.get_textpage()
                text = textpage.get_text_range()
                for line in text.split('\n'):
                    if line.strip():
                        ws.cell(row=row, column=1, value=line)
                        row += 1
            base = os.path.splitext(os.path.basename(pdf_path))[0]
            out_file = os.path.join(self.output_path, f"{base}.xlsx")
            wb.save(out_file)
            done_files.append(out_file)
        self.progress_updated.emit(100, "Done!")
        self.finished.emit(f"OK|{len(done_files)} XLSX files created.")


# ──────────────────────────────────────────────────────────────────────────
#  Output folder picker (reusable row widget)
# ──────────────────────────────────────────────────────────────────────────
class OutputFolderRow(QWidget):
    def __init__(self, lang, parent=None):
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        self.output_label = QLabel(lang["output_label"])
        self.path_edit = QLineEdit()
        self.path_edit.setReadOnly(True)
        self.path_edit.setPlaceholderText(lang["output_placeholder"])
        self.browse_btn = QPushButton(lang["choose_folder_btn"])
        self.browse_btn.clicked.connect(self._browse)
        layout.addWidget(self.output_label)
        layout.addWidget(self.path_edit, 1)
        layout.addWidget(self.browse_btn)
        self.lang = lang

    def retranslate(self, lang):
        self.lang = lang
        self.output_label.setText(lang["output_label"])
        self.path_edit.setPlaceholderText(lang["output_placeholder"])
        self.browse_btn.setText(lang["choose_folder_btn"])

    def _browse(self):
        folder = QFileDialog.getExistingDirectory(self, self.lang["choose_folder_title"])
        if folder:
            self.path_edit.setText(folder)

    def get_path(self):
        return self.path_edit.text()


# ──────────────────────────────────────────────────────────────────────────
#  Base page: drop area + output + progress + start/stop (shared scaffolding)
# ──────────────────────────────────────────────────────────────────────────
class BaseToolPage(QWidget):
    def __init__(self, title_key, accept_types, lang, parent=None):
        super().__init__(parent)
        self.thread = None
        self.worker = None
        self.lang = lang
        self.title_key = title_key

        root = QHBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)
        root.addWidget(splitter)

        # ── main area ──
        main_area = QWidget()
        main_layout = QVBoxLayout(main_area)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(8)

        self.title_label = QLabel(f"<b>{lang[title_key]}</b>")
        self.title_label.setStyleSheet("font-size: 13pt; background: transparent;")
        main_layout.addWidget(self.title_label)

        btn_row = QHBoxLayout()
        self.add_btn = QPushButton(lang["add_files_btn"])
        self.add_btn.clicked.connect(self._browse_files)
        self.clear_btn = QPushButton(lang["clear_btn"])
        self.clear_btn.clicked.connect(self._clear_files)
        btn_row.addWidget(self.add_btn)
        btn_row.addWidget(self.clear_btn)
        btn_row.addStretch()
        main_layout.addLayout(btn_row)

        self.accept_types = accept_types
        self.drop_area = FileDropArea(accept_types=accept_types, lang=lang)
        main_layout.addWidget(self.drop_area, 1)

        self.output_row = OutputFolderRow(lang)
        main_layout.addWidget(self.output_row)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        main_layout.addWidget(self.progress_bar)

        self.status_label = QLabel(lang["status_ready"])
        self.status_label.setStyleSheet("background: transparent;")
        main_layout.addWidget(self.status_label)

        action_row = QHBoxLayout()
        self.start_btn = QPushButton(lang["start_btn"])
        self.start_btn.setObjectName("startButton")
        self.start_btn.clicked.connect(self._on_start_clicked)
        self.stop_btn = QPushButton(lang["stop_btn"])
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self._on_stop_clicked)
        action_row.addWidget(self.start_btn)
        action_row.addWidget(self.stop_btn)
        main_layout.addLayout(action_row)

        splitter.addWidget(main_area)

        # ── options sidebar (filled by subclass) ──
        self.sidebar = QFrame()
        self.sidebar.setObjectName("optionsSidebar")
        self.sidebar.setMinimumWidth(260)
        self.sidebar_layout = QVBoxLayout(self.sidebar)
        self.sidebar_layout.setContentsMargins(12, 12, 12, 12)
        self.sidebar_layout.setSpacing(8)
        splitter.addWidget(self.sidebar)
        splitter.setSizes([600, 280])

        self.options_label = QLabel(f"<b>{lang['options_label']}</b>")
        self.options_label.setStyleSheet("background: transparent;")
        self.sidebar_layout.addWidget(self.options_label)

        self.build_options(self.sidebar_layout, lang)
        self.sidebar_layout.addStretch()

    # subclasses override this to add their option widgets
    def build_options(self, layout, lang):
        pass

    def retranslate_options(self, lang):
        pass

    def retranslate(self, lang):
        self.lang = lang
        self.title_label.setText(f"<b>{lang[self.title_key]}</b>")
        self.add_btn.setText(lang["add_files_btn"])
        self.clear_btn.setText(lang["clear_btn"])
        self.drop_area.retranslate(lang)
        self.output_row.retranslate(lang)
        self.status_label.setText(lang["status_ready"])
        self.stop_btn.setText(lang["stop_btn"])
        self.options_label.setText(f"<b>{lang['options_label']}</b>")
        self.retranslate_options(lang)

    def _browse_files(self):
        filters = []
        if 'image' in self.accept_types:
            filters.append("*.png *.jpg *.jpeg *.bmp *.webp *.gif")
        if 'pdf' in self.accept_types:
            filters.append("*.pdf")
        filter_str = "Files (" + " ".join(filters) + ")"
        files, _ = QFileDialog.getOpenFileNames(self, self.lang["add_files_btn"], "", filter_str)
        if files:
            self.drop_area.add_files(files)

    def _clear_files(self):
        self.drop_area.clear_files()

    def _filtered_paths(self, exts):
        return [p for p in self.drop_area.get_all_file_paths()
                if os.path.splitext(p)[1].lower() in exts]

    def _validate_output(self):
        out = self.output_row.get_path()
        if not out or not os.path.isdir(out):
            QMessageBox.warning(self, self.lang["error_title"], self.lang["error_no_output"])
            return None
        return out

    def _on_start_clicked(self):
        raise NotImplementedError

    def _run_worker(self, worker):
        self.thread = QThread()
        self.worker = worker
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.progress_updated.connect(self._on_progress)
        self.worker.finished.connect(self._on_finished)
        self.worker.error.connect(self._on_error)
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.thread.start()

    def _on_stop_clicked(self):
        if self.worker:
            self.worker.stop()
        self.status_label.setText(self.lang["status_stopped"])
        self._cleanup_thread()

    @Slot(int, str)
    def _on_progress(self, value, message):
        self.progress_bar.setValue(value)
        self.status_label.setText(message)

    @Slot(str)
    def _on_finished(self, message):
        # message format "OK|detail text"
        detail = message.split("|", 1)[1] if "|" in message else message
        self.progress_bar.setValue(100)
        self.status_label.setText(detail)
        QMessageBox.information(self, self.lang["done_title"], detail)
        self._cleanup_thread()

    @Slot(str)
    def _on_error(self, message):
        self.status_label.setText(message)
        QMessageBox.critical(self, self.lang["error_title"], message)
        self._cleanup_thread()

    def _cleanup_thread(self):
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        if self.thread:
            self.thread.quit()
            self.thread.wait()
            self.thread = None
        self.worker = None


# ──────────────────────────────────────────────────────────────────────────
#  4 sub-tool pages
# ──────────────────────────────────────────────────────────────────────────
class ImageToPdfPage(BaseToolPage):
    def __init__(self, lang, parent=None):
        super().__init__("img2pdf_title", ['image'], lang, parent)
        self.start_btn.setText(lang["img2pdf_start_btn"])

    def build_options(self, layout, lang):
        grid = QGridLayout(); grid.setSpacing(6)
        grid.setColumnMinimumWidth(0, 110); grid.setColumnStretch(1, 1)

        self.quality_label = QLabel(lang["quality_label"])
        grid.addWidget(self.quality_label, 0, 0)
        self.quality_combo = QComboBox()
        self.quality_combo.addItems(lang["qualities"])
        self.quality_combo.setCurrentIndex(2)
        grid.addWidget(self.quality_combo, 0, 1)

        self.out_mode_label = QLabel(lang["out_mode_label"])
        grid.addWidget(self.out_mode_label, 1, 0)
        self.output_combo = QComboBox()
        self.output_combo.addItems(lang["out_modes"])
        grid.addWidget(self.output_combo, 1, 1)

        self.target_size_label = QLabel(lang["target_size_label"])
        grid.addWidget(self.target_size_label, 2, 0)
        self.target_spin = QSpinBox()
        self.target_spin.setRange(0, 9999)
        self.target_spin.setSuffix(" MB")
        self.target_spin.setToolTip(lang["target_size_hint"])
        grid.addWidget(self.target_spin, 2, 1)

        layout.addLayout(grid)

    def retranslate_options(self, lang):
        self.start_btn.setText(lang["img2pdf_start_btn"])
        self.quality_label.setText(lang["quality_label"])
        self.out_mode_label.setText(lang["out_mode_label"])
        self.target_size_label.setText(lang["target_size_label"])
        self.target_spin.setToolTip(lang["target_size_hint"])
        idx_q, idx_o = self.quality_combo.currentIndex(), self.output_combo.currentIndex()
        self.quality_combo.clear(); self.quality_combo.addItems(lang["qualities"]); self.quality_combo.setCurrentIndex(idx_q)
        self.output_combo.clear(); self.output_combo.addItems(lang["out_modes"]); self.output_combo.setCurrentIndex(idx_o)

    def _on_start_clicked(self):
        img_paths = self._filtered_paths(IMAGE_EXT)
        if not img_paths:
            QMessageBox.warning(self, self.lang["error_title"], self.lang["img2pdf_no_files"])
            return
        out = self._validate_output()
        if not out:
            return
        single = self.output_combo.currentIndex() == 0
        mode = 'img2pdf_single' if single else 'img2pdf_multi'
        worker = PdfToolsWorker(
            mode=mode, input_paths=img_paths, output_path=out,
            quality_idx=self.quality_combo.currentIndex(),
            target_mb=float(self.target_spin.value()),
            output_name="images_combined.pdf" if single else "")
        self._run_worker(worker)


class PdfToImagePage(BaseToolPage):
    def __init__(self, lang, parent=None):
        super().__init__("pdf2img_title", ['pdf'], lang, parent)
        self.start_btn.setText(lang["pdf2img_start_btn"])

    def build_options(self, layout, lang):
        grid = QGridLayout(); grid.setSpacing(6)
        grid.setColumnMinimumWidth(0, 110); grid.setColumnStretch(1, 1)

        self.format_label = QLabel(lang["format_label"])
        grid.addWidget(self.format_label, 0, 0)
        self.format_combo = QComboBox()
        self.format_combo.addItems(["PNG", "JPG", "WEBP"])
        grid.addWidget(self.format_combo, 0, 1)

        self.dpi_label = QLabel(lang["dpi_label"])
        grid.addWidget(self.dpi_label, 1, 0)
        self.dpi_combo = QComboBox()
        self.dpi_combo.addItems(lang["dpis"])
        self.dpi_combo.setCurrentIndex(2)
        grid.addWidget(self.dpi_combo, 1, 1)

        self.pages_label = QLabel(lang["pages_label"])
        grid.addWidget(self.pages_label, 2, 0)
        self.pages_edit = QLineEdit()
        self.pages_edit.setPlaceholderText(lang["pages_hint"])
        grid.addWidget(self.pages_edit, 2, 1)

        layout.addLayout(grid)

    def retranslate_options(self, lang):
        self.start_btn.setText(lang["pdf2img_start_btn"])
        self.format_label.setText(lang["format_label"])
        self.dpi_label.setText(lang["dpi_label"])
        self.pages_label.setText(lang["pages_label"])
        self.pages_edit.setPlaceholderText(lang["pages_hint"])
        idx_d = self.dpi_combo.currentIndex()
        self.dpi_combo.clear(); self.dpi_combo.addItems(lang["dpis"]); self.dpi_combo.setCurrentIndex(idx_d)

    def _on_start_clicked(self):
        pdf_paths = self._filtered_paths(PDF_EXT)
        if not pdf_paths:
            QMessageBox.warning(self, self.lang["error_title"], self.lang["pdf2img_no_files"])
            return
        out = self._validate_output()
        if not out:
            return
        worker = PdfToolsWorker(
            mode='pdf2img', input_paths=pdf_paths, output_path=out,
            fmt=self.format_combo.currentText(),
            dpi_idx=self.dpi_combo.currentIndex(),
            page_sel=self.pages_edit.text())
        self._run_worker(worker)


class PdfMergerPage(BaseToolPage):
    def __init__(self, lang, parent=None):
        super().__init__("merger_title", ['pdf'], lang, parent)
        self.start_btn.setText(lang["merger_start_btn"])

    def build_options(self, layout, lang):
        grid = QGridLayout(); grid.setSpacing(6)
        grid.setColumnMinimumWidth(0, 110); grid.setColumnStretch(1, 1)

        self.quality_label = QLabel(lang["quality_label"])
        grid.addWidget(self.quality_label, 0, 0)
        self.quality_combo = QComboBox()
        self.quality_combo.addItems(lang["merger_qualities"])
        grid.addWidget(self.quality_combo, 0, 1)

        self.target_size_label = QLabel(lang["target_size_label"])
        grid.addWidget(self.target_size_label, 1, 0)
        self.target_spin = QSpinBox()
        self.target_spin.setRange(0, 9999)
        self.target_spin.setSuffix(" MB")
        self.target_spin.setToolTip(lang["target_size_hint"])
        grid.addWidget(self.target_spin, 1, 1)

        layout.addLayout(grid)

        self.pikepdf_note = None
        if not HAS_PIKEPDF:
            self.pikepdf_note = QLabel(lang["merger_pikepdf_note"])
            self.pikepdf_note.setStyleSheet("color: #d08770; font-size: 8pt; background: transparent;")
            layout.addWidget(self.pikepdf_note)

    def retranslate_options(self, lang):
        self.start_btn.setText(lang["merger_start_btn"])
        self.quality_label.setText(lang["quality_label"])
        self.target_size_label.setText(lang["target_size_label"])
        self.target_spin.setToolTip(lang["target_size_hint"])
        idx_q = self.quality_combo.currentIndex()
        self.quality_combo.clear(); self.quality_combo.addItems(lang["merger_qualities"]); self.quality_combo.setCurrentIndex(idx_q)
        if self.pikepdf_note:
            self.pikepdf_note.setText(lang["merger_pikepdf_note"])

    def _on_start_clicked(self):
        pdf_paths = self._filtered_paths(PDF_EXT)
        if len(pdf_paths) < 2:
            QMessageBox.warning(self, self.lang["error_title"], self.lang["merger_no_files"])
            return
        out = self._validate_output()
        if not out:
            return
        worker = PdfToolsWorker(
            mode='merge', input_paths=pdf_paths, output_path=out,
            quality_idx=self.quality_combo.currentIndex(),
            target_mb=float(self.target_spin.value()),
            output_name="merged.pdf")
        self._run_worker(worker)


class PdfDocConversionPage(BaseToolPage):
    def __init__(self, lang, parent=None):
        super().__init__("nav_docconv", ['pdf'], lang, parent)
        self.start_btn.setText(lang["docconv_start_btn"])

    def build_options(self, layout, lang):
        grid = QGridLayout(); grid.setSpacing(6)
        grid.setColumnMinimumWidth(0, 110); grid.setColumnStretch(1, 1)

        self.convert_to_label = QLabel(lang["convert_to_label"])
        grid.addWidget(self.convert_to_label, 0, 0)
        self.target_combo = QComboBox()
        items = ["TXT"]
        if HAS_DOCX:
            items.append("DOCX")
        if HAS_OPENPYXL:
            items.append("XLSX")
        self.target_combo.addItems(items)
        grid.addWidget(self.target_combo, 0, 1)

        layout.addLayout(grid)

        self.missing_note = None
        missing = []
        if not HAS_DOCX:
            missing.append("python-docx (DOCX)")
        if not HAS_OPENPYXL:
            missing.append("openpyxl (XLSX)")
        if missing:
            self.missing_note = QLabel(lang["missing_libs_label"].format(items=", ".join(missing)))
            self.missing_note.setStyleSheet("color: #d08770; font-size: 8pt; background: transparent;")
            self.missing_note.setWordWrap(True)
            layout.addWidget(self.missing_note)
        self._missing_items = missing

    def retranslate_options(self, lang):
        self.start_btn.setText(lang["docconv_start_btn"])
        self.convert_to_label.setText(lang["convert_to_label"])
        if self.missing_note:
            self.missing_note.setText(lang["missing_libs_label"].format(items=", ".join(self._missing_items)))

    def _on_start_clicked(self):
        pdf_paths = self._filtered_paths(PDF_EXT)
        if not pdf_paths:
            QMessageBox.warning(self, self.lang["error_title"], self.lang["docconv_no_files"])
            return
        out = self._validate_output()
        if not out:
            return
        target = self.target_combo.currentText()
        mode_map = {"TXT": "pdf2txt", "DOCX": "pdf2docx", "XLSX": "pdf2xlsx"}
        worker = PdfToolsWorker(mode=mode_map[target], input_paths=pdf_paths, output_path=out)
        self._run_worker(worker)


# ──────────────────────────────────────────────────────────────────────────
#  Main window
# ──────────────────────────────────────────────────────────────────────────
class MacanPdfToolsApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.current_lang_code = "id"
        self.lang = LANGUAGES[self.current_lang_code]
        self.resize(1100, 760)

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # ── top bar: language switcher ──
        top_bar = QWidget()
        top_bar.setObjectName("topBar")
        top_layout = QHBoxLayout(top_bar)
        top_layout.setContentsMargins(12, 6, 12, 6)
        self.lang_label = QLabel(self.lang["lang_label"])
        self.lang_label.setStyleSheet("background: transparent;")
        top_layout.addWidget(self.lang_label)
        self.lang_combo = QComboBox()
        self.lang_combo.addItems(["Bahasa Indonesia", "English"])
        self.lang_combo.setFixedWidth(160)
        self.lang_combo.currentIndexChanged.connect(self._on_lang_changed)
        top_layout.addWidget(self.lang_combo)
        top_layout.addStretch()
        root.addWidget(top_bar)

        body = QWidget()
        body_layout = QHBoxLayout(body)
        body_layout.setContentsMargins(0, 0, 0, 0)
        body_layout.setSpacing(0)
        root.addWidget(body, 1)

        # left nav
        self.nav_list = QListWidget()
        self.nav_list.setObjectName("navList")
        self.nav_list.setFixedWidth(210)
        self._populate_nav()
        self.nav_list.currentRowChanged.connect(self._on_nav_changed)
        body_layout.addWidget(self.nav_list)

        # stacked pages
        self.stack = QStackedWidget()
        self.pages = [
            ImageToPdfPage(self.lang),
            PdfToImagePage(self.lang),
            PdfMergerPage(self.lang),
            PdfDocConversionPage(self.lang),
        ]
        for p in self.pages:
            self.stack.addWidget(p)
        body_layout.addWidget(self.stack, 1)

        self.nav_list.setCurrentRow(0)

        self._apply_stylesheet()
        self._set_window_title()
        self._check_dependencies()

    def _populate_nav(self):
        self.nav_list.clear()
        self.nav_list.addItems([
            self.lang["nav_img2pdf"],
            self.lang["nav_pdf2img"],
            self.lang["nav_merger"],
            self.lang["nav_docconv"],
        ])

    def _set_window_title(self):
        self.setWindowTitle(self.lang["window_title"].format(version=APP_VERSION))

    def _on_nav_changed(self, index):
        if index >= 0:
            self.stack.setCurrentIndex(index)

    def _on_lang_changed(self, index):
        self.current_lang_code = "id" if index == 0 else "en"
        self.lang = LANGUAGES[self.current_lang_code]

        current_row = self.nav_list.currentRow()
        self.lang_label.setText(self.lang["lang_label"])
        self._populate_nav()
        self.nav_list.setCurrentRow(current_row)
        self._set_window_title()

        for p in self.pages:
            p.retranslate(self.lang)

    def _check_dependencies(self):
        missing = []
        if not HAS_PDFIUM:
            missing.append("pypdfium2")
        if missing:
            QMessageBox.warning(
                self, self.lang["dep_missing_title"],
                self.lang["dep_missing_body"].format(items="\n".join(missing), pkgs=" ".join(missing)))

    def _apply_stylesheet(self):
        self.setStyleSheet("""
            QMainWindow, QWidget {
                background-color: #252525;
                color: #E0E0E0;
                font-family: Segoe UI, sans-serif;
            }
            QLabel {
                font-size: 10pt;
                background: transparent;
            }
            #topBar {
                background-color: #2c2c2c;
                border-bottom: 1px solid #444444;
            }
            #navList {
                background-color: #333333;
                border: none;
                border-right: 1px solid #444444;
            }
            #navList::item {
                padding: 12px 10px;
                border-bottom: 1px solid #3a3a3a;
                background: transparent;
            }
            #navList::item:selected {
                background-color: #5A5A5A;
                color: #FFFFFF;
            }
            #navList::item:hover:!selected {
                background-color: #4A4A4A;
            }
            FileDropArea {
                background-color: #333333;
                border: 2px dashed #555555;
                border-radius: 5px;
                color: #AAAAAA;
            }
            FileDropArea::item {
                background-color: #404040;
                border: 1px solid #555555;
                border-radius: 4px;
                padding: 5px;
            }
            FileDropArea::item:selected {
                background-color: #5A5A5A;
                border-color: #A3BE8C;
            }
            #optionsSidebar {
                background-color: #2c2c2c;
                border-left: 1px solid #444444;
            }
            QLineEdit, QComboBox, QSpinBox {
                background-color: #3A3A3A;
                border: 1px solid #555555;
                border-radius: 4px;
                padding: 4px;
                color: #E0E0E0;
            }
            QPushButton {
                background-color: #4A4A4A;
                border: 1px solid #5A5A5A;
                border-radius: 4px;
                padding: 6px 12px;
            }
            QPushButton:hover { background-color: #5A5A5A; }
            QPushButton:disabled { color: #777777; background-color: #3a3a3a; }
            #startButton {
                background-color: #4C7A4C;
                font-weight: bold;
            }
            #startButton:hover { background-color: #5C8A5C; }
            QProgressBar {
                border: 1px solid #555555;
                border-radius: 4px;
                text-align: center;
                background-color: #333333;
            }
            QProgressBar::chunk { background-color: #5A8A5A; }
        """)


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = MacanPdfToolsApp()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
