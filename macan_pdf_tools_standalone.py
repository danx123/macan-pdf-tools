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
    QMessageBox, QStackedWidget, QSplitter, QTextEdit
)
from PySide6.QtCore import Qt, QSize, QThread, QObject, Signal, Slot
from PySide6.QtGui import QIcon, QPixmap, QPainter, QColor, QFont, QDragEnterEvent, QDragMoveEvent, QDropEvent

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


APP_VERSION = "1.0.0"
IMAGE_EXT = ['.png', '.jpg', '.jpeg', '.bmp', '.webp', '.gif']
PDF_EXT = ['.pdf']


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


def pil_to_qpixmap(pil_img):
    buf = io.BytesIO()
    pil_img.save(buf, format='PNG')
    pix = QPixmap()
    pix.loadFromData(buf.getvalue(), 'PNG')
    return pix


# ──────────────────────────────────────────────────────────────────────────
#  Drag & drop file list (no thumbnail worker thread pool needed — lightweight)
# ──────────────────────────────────────────────────────────────────────────
class FileDropArea(QListWidget):
    files_changed = Signal()

    def __init__(self, accept_types=None, parent=None):
        super().__init__(parent)
        self.accept_types = accept_types or ['image', 'pdf']
        self.file_paths = []

        self.setAcceptDrops(True)
        self.setDragDropMode(QAbstractItemView.DragDropMode.NoDragDrop)
        self.setViewMode(QListWidget.ViewMode.IconMode)
        self.setIconSize(QSize(96, 96))
        self.setResizeMode(QListWidget.ResizeMode.Adjust)
        self.setWordWrap(True)
        self.setSpacing(10)
        self._icon_cache = {}
        self._set_placeholder()

    def _set_placeholder(self):
        self._placeholder_item = QListWidgetItem(
            "Seret file ke sini atau klik 'Tambah File'\n(Drag files here or click 'Add Files')")
        self._placeholder_item.setFlags(Qt.ItemFlag.NoItemFlags)
        self._placeholder_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.addItem(self._placeholder_item)

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
            if item != getattr(self, '_placeholder_item', None):
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
            self.progress_updated.emit(int((i / total) * 85), f"Memproses gambar {i+1}/{total}...")
            pil_images.append(Image.open(path).convert('RGB'))

        if not pil_images:
            self.error.emit("Tidak ada gambar.")
            return

        self.progress_updated.emit(90, "Menyimpan PDF...")
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

        self.progress_updated.emit(100, "Selesai!")
        self.finished.emit(f"Berhasil! PDF disimpan: {os.path.basename(output_file)}")

    def _images_to_multi_pdf(self):
        quality = self._quality_to_jpeg(self.kwargs.get('quality_idx', 2))
        target_mb = self.kwargs.get('target_mb', 0.0)
        total = len(self.input_paths)
        for i, path in enumerate(self.input_paths):
            if not self.is_running:
                return
            self.progress_updated.emit(int((i / total) * 100), f"Memproses gambar {i+1}/{total}...")
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

        self.progress_updated.emit(100, "Selesai!")
        self.finished.emit(f"Berhasil! {total} file PDF telah dibuat.")

    # ---- PDF -> Image ----
    def _pdf_to_images(self):
        if not HAS_PDFIUM:
            self.error.emit("Library 'pypdfium2' belum terpasang.")
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
                    f"Mengonversi halaman {j+1}/{len(pages)}...")
                pil_img = pdf[page_idx].render(scale=scale).to_pil()
                out_file = os.path.join(self.output_path, f"{base}_p{page_idx+1}.{fmt}")
                if fmt in ('jpg', 'jpeg', 'webp'):
                    pil_img.convert('RGB').save(out_file)
                else:
                    pil_img.save(out_file)
                total_exported += 1

        self.progress_updated.emit(100, "Selesai!")
        self.finished.emit(f"Berhasil! {total_exported} gambar diekspor.")

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
        self.progress_updated.emit(2, "Membuka file...")
        try:
            merged = pikepdf.Pdf.new()
            for i, path in enumerate(self.input_paths):
                if not self.is_running:
                    return
                src = pikepdf.Pdf.open(path)
                merged.pages.extend(src.pages)
                self.progress_updated.emit(int(((i + 1) / total_files) * 50), f"Menggabungkan {i+1}/{total_files}...")
        except Exception as e:
            self.error.emit(f"Merge error: {e}")
            return

        if not self.is_running:
            return

        self.progress_updated.emit(55, "Mengoptimalkan...")
        try:
            merged.remove_unreferenced_resources()
        except Exception:
            pass

        if jpeg_quality is not None:
            self.progress_updated.emit(60, "Mengompresi gambar...")
            try:
                self._recompress_images_in_pdf(merged, jpeg_quality)
            except Exception:
                pass

        self.progress_updated.emit(85, "Menyimpan...")
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
                    self.progress_updated.emit(85 + iteration * 4, f"Memperkecil ukuran (q={q})...")
                    try:
                        self._recompress_images_in_pdf(merged, q)
                        merged.save(output_file, **save_opts)
                        actual_mb = os.path.getsize(output_file) / (1024 * 1024)
                    except Exception:
                        break
        except Exception as e:
            self.error.emit(f"Save error: {e}")
            return

        self.progress_updated.emit(100, "Selesai!")
        self.finished.emit(f"Berhasil! PDF gabungan disimpan: {os.path.basename(output_file)}")

    def _merge_pdfs_fallback(self, output_file):
        if not HAS_PDFIUM:
            self.error.emit("Library 'pypdfium2' atau 'pikepdf' belum terpasang.")
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
            self.progress_updated.emit(int((done / max(total_pages, 1)) * 95), "Menggabungkan...")
        merged.save(output_file)
        self.progress_updated.emit(100, "Selesai!")
        self.finished.emit(f"Berhasil! PDF gabungan disimpan: {os.path.basename(output_file)}")

    # ---- PDF Document Conversion ----
    def _pdf_to_txt(self):
        if not HAS_PDFIUM:
            self.error.emit("Library 'pypdfium2' belum terpasang.")
            return
        total = len(self.input_paths)
        done_files = []
        for i, pdf_path in enumerate(self.input_paths):
            if not self.is_running:
                return
            self.progress_updated.emit(int((i / total) * 95), f"Mengekstrak teks {i+1}/{total}...")
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
        self.progress_updated.emit(100, "Selesai!")
        self.finished.emit(f"Berhasil! {len(done_files)} file TXT dibuat.")

    def _pdf_to_docx(self):
        if not HAS_PDFIUM:
            self.error.emit("Library 'pypdfium2' belum terpasang.")
            return
        if not HAS_DOCX:
            self.error.emit("Library 'python-docx' belum terpasang. Install: pip install python-docx --break-system-packages")
            return
        import docx
        total = len(self.input_paths)
        done_files = []
        for i, pdf_path in enumerate(self.input_paths):
            if not self.is_running:
                return
            self.progress_updated.emit(int((i / total) * 95), f"Mengonversi {i+1}/{total}...")
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
        self.progress_updated.emit(100, "Selesai!")
        self.finished.emit(f"Berhasil! {len(done_files)} file DOCX dibuat.")

    def _pdf_to_xlsx(self):
        if not HAS_PDFIUM:
            self.error.emit("Library 'pypdfium2' belum terpasang.")
            return
        if not HAS_OPENPYXL:
            self.error.emit("Library 'openpyxl' belum terpasang. Install: pip install openpyxl --break-system-packages")
            return
        import openpyxl
        total = len(self.input_paths)
        done_files = []
        for i, pdf_path in enumerate(self.input_paths):
            if not self.is_running:
                return
            self.progress_updated.emit(int((i / total) * 95), f"Mengonversi {i+1}/{total}...")
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
        self.progress_updated.emit(100, "Selesai!")
        self.finished.emit(f"Berhasil! {len(done_files)} file XLSX dibuat.")


# ──────────────────────────────────────────────────────────────────────────
#  Output folder picker (reusable row widget)
# ──────────────────────────────────────────────────────────────────────────
class OutputFolderRow(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        self.path_edit = QLineEdit()
        self.path_edit.setReadOnly(True)
        self.path_edit.setPlaceholderText("Pilih folder output...")
        browse_btn = QPushButton("Pilih Folder")
        browse_btn.clicked.connect(self._browse)
        layout.addWidget(QLabel("Output:"))
        layout.addWidget(self.path_edit, 1)
        layout.addWidget(browse_btn)

    def _browse(self):
        folder = QFileDialog.getExistingDirectory(self, "Pilih Folder Output")
        if folder:
            self.path_edit.setText(folder)

    def get_path(self):
        return self.path_edit.text()


# ──────────────────────────────────────────────────────────────────────────
#  Base page: drop area + output + progress + start/stop (shared scaffolding)
# ──────────────────────────────────────────────────────────────────────────
class BaseToolPage(QWidget):
    def __init__(self, title, accept_types, parent=None):
        super().__init__(parent)
        self.thread = None
        self.worker = None

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

        title_label = QLabel(f"<b>{title}</b>")
        title_label.setStyleSheet("font-size: 13pt;")
        main_layout.addWidget(title_label)

        btn_row = QHBoxLayout()
        add_btn = QPushButton("Tambah File")
        add_btn.clicked.connect(self._browse_files)
        clear_btn = QPushButton("Bersihkan")
        clear_btn.clicked.connect(self._clear_files)
        btn_row.addWidget(add_btn)
        btn_row.addWidget(clear_btn)
        btn_row.addStretch()
        main_layout.addLayout(btn_row)

        self.accept_types = accept_types
        self.drop_area = FileDropArea(accept_types=accept_types)
        main_layout.addWidget(self.drop_area, 1)

        self.output_row = OutputFolderRow()
        main_layout.addWidget(self.output_row)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        main_layout.addWidget(self.progress_bar)

        self.status_label = QLabel("Siap.")
        main_layout.addWidget(self.status_label)

        action_row = QHBoxLayout()
        self.start_btn = QPushButton("Mulai")
        self.start_btn.setObjectName("startButton")
        self.start_btn.clicked.connect(self._on_start_clicked)
        self.stop_btn = QPushButton("Stop")
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

        self.build_options(self.sidebar_layout)
        self.sidebar_layout.addStretch()

    # subclasses override this to add their option widgets
    def build_options(self, layout):
        pass

    def _browse_files(self):
        filters = []
        if 'image' in self.accept_types:
            filters.append("*.png *.jpg *.jpeg *.bmp *.webp *.gif")
        if 'pdf' in self.accept_types:
            filters.append("*.pdf")
        filter_str = "Files (" + " ".join(filters) + ")"
        files, _ = QFileDialog.getOpenFileNames(self, "Pilih File", "", filter_str)
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
            QMessageBox.warning(self, "Error", "Pilih folder output terlebih dahulu.")
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
        self.status_label.setText("Dihentikan oleh pengguna.")
        self._cleanup_thread()

    @Slot(int, str)
    def _on_progress(self, value, message):
        self.progress_bar.setValue(value)
        self.status_label.setText(message)

    @Slot(str)
    def _on_finished(self, message):
        self.progress_bar.setValue(100)
        self.status_label.setText(message)
        QMessageBox.information(self, "Selesai", message)
        self._cleanup_thread()

    @Slot(str)
    def _on_error(self, message):
        self.status_label.setText(message)
        QMessageBox.critical(self, "Error", message)
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
    def __init__(self, parent=None):
        super().__init__("Image to PDF", ['image'], parent)
        self.start_btn.setText("Gabungkan -> PDF")

    def build_options(self, layout):
        layout.addWidget(QLabel("<b>Opsi</b>"))
        grid = QGridLayout(); grid.setSpacing(6)
        grid.setColumnMinimumWidth(0, 110); grid.setColumnStretch(1, 1)

        grid.addWidget(QLabel("Kualitas:"), 0, 0)
        self.quality_combo = QComboBox()
        self.quality_combo.addItems(["Maksimum (100)", "Bagus (95)", "Baik (85)", "Sedang (75)", "Rendah (50)"])
        self.quality_combo.setCurrentIndex(2)
        grid.addWidget(self.quality_combo, 0, 1)

        grid.addWidget(QLabel("Output:"), 1, 0)
        self.output_combo = QComboBox()
        self.output_combo.addItems(["Satu PDF untuk semua gambar", "Satu PDF per gambar"])
        grid.addWidget(self.output_combo, 1, 1)

        grid.addWidget(QLabel("Target Ukuran:"), 2, 0)
        self.target_spin = QSpinBox()
        self.target_spin.setRange(0, 9999)
        self.target_spin.setSuffix(" MB")
        self.target_spin.setToolTip("0 = otomatis")
        grid.addWidget(self.target_spin, 2, 1)

        layout.addLayout(grid)

    def _on_start_clicked(self):
        img_paths = self._filtered_paths(IMAGE_EXT)
        if not img_paths:
            QMessageBox.warning(self, "Error", "Tambahkan file gambar terlebih dahulu.")
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
    def __init__(self, parent=None):
        super().__init__("PDF to Image", ['pdf'], parent)
        self.start_btn.setText("Konversi -> Gambar")

    def build_options(self, layout):
        layout.addWidget(QLabel("<b>Opsi</b>"))
        grid = QGridLayout(); grid.setSpacing(6)
        grid.setColumnMinimumWidth(0, 110); grid.setColumnStretch(1, 1)

        grid.addWidget(QLabel("Format:"), 0, 0)
        self.format_combo = QComboBox()
        self.format_combo.addItems(["PNG", "JPG", "WEBP"])
        grid.addWidget(self.format_combo, 0, 1)

        grid.addWidget(QLabel("Kualitas (DPI):"), 1, 0)
        self.dpi_combo = QComboBox()
        self.dpi_combo.addItems(["Rendah (72 DPI)", "Sedang (150 DPI)", "Baik (200 DPI)", "Tinggi (300 DPI)", "Maksimum (600 DPI)"])
        self.dpi_combo.setCurrentIndex(2)
        grid.addWidget(self.dpi_combo, 1, 1)

        grid.addWidget(QLabel("Halaman:"), 2, 0)
        self.pages_edit = QLineEdit()
        self.pages_edit.setPlaceholderText("Cth: 1,3,5-8 (kosong = semua)")
        grid.addWidget(self.pages_edit, 2, 1)

        layout.addLayout(grid)

    def _on_start_clicked(self):
        pdf_paths = self._filtered_paths(PDF_EXT)
        if not pdf_paths:
            QMessageBox.warning(self, "Error", "Tambahkan file PDF terlebih dahulu.")
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
    def __init__(self, parent=None):
        super().__init__("PDF Merger", ['pdf'], parent)
        self.start_btn.setText("Gabungkan PDF")

    def build_options(self, layout):
        layout.addWidget(QLabel("<b>Opsi</b>"))
        grid = QGridLayout(); grid.setSpacing(6)
        grid.setColumnMinimumWidth(0, 110); grid.setColumnStretch(1, 1)

        grid.addWidget(QLabel("Kualitas Gambar:"), 0, 0)
        self.quality_combo = QComboBox()
        self.quality_combo.addItems(["Pertahankan asli", "Optimal (85)", "Rendah (60)"])
        grid.addWidget(self.quality_combo, 0, 1)

        grid.addWidget(QLabel("Target Ukuran:"), 1, 0)
        self.target_spin = QSpinBox()
        self.target_spin.setRange(0, 9999)
        self.target_spin.setSuffix(" MB")
        self.target_spin.setToolTip("0 = otomatis")
        grid.addWidget(self.target_spin, 1, 1)

        layout.addLayout(grid)

        if not HAS_PIKEPDF:
            note = QLabel("Catatan: 'pikepdf' tidak terpasang — kompresi tidak\ntersedia, hanya merge struktural (via pypdfium2).")
            note.setStyleSheet("color: #d08770; font-size: 8pt;")
            layout.addWidget(note)

    def _on_start_clicked(self):
        pdf_paths = self._filtered_paths(PDF_EXT)
        if len(pdf_paths) < 2:
            QMessageBox.warning(self, "Error", "Tambahkan minimal 2 file PDF.")
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
    def __init__(self, parent=None):
        super().__init__("PDF Document Conversion", ['pdf'], parent)
        self.start_btn.setText("Konversi")

    def build_options(self, layout):
        layout.addWidget(QLabel("<b>Opsi</b>"))
        grid = QGridLayout(); grid.setSpacing(6)
        grid.setColumnMinimumWidth(0, 110); grid.setColumnStretch(1, 1)

        grid.addWidget(QLabel("Konversi ke:"), 0, 0)
        self.target_combo = QComboBox()
        items = ["TXT"]
        if HAS_DOCX:
            items.append("DOCX")
        if HAS_OPENPYXL:
            items.append("XLSX")
        self.target_combo.addItems(items)
        grid.addWidget(self.target_combo, 0, 1)

        layout.addLayout(grid)

        missing = []
        if not HAS_DOCX:
            missing.append("python-docx (untuk DOCX)")
        if not HAS_OPENPYXL:
            missing.append("openpyxl (untuk XLSX)")
        if missing:
            note = QLabel("Belum terpasang: " + ", ".join(missing))
            note.setStyleSheet("color: #d08770; font-size: 8pt;")
            note.setWordWrap(True)
            layout.addWidget(note)

    def _on_start_clicked(self):
        pdf_paths = self._filtered_paths(PDF_EXT)
        if not pdf_paths:
            QMessageBox.warning(self, "Error", "Tambahkan file PDF terlebih dahulu.")
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
        self.setWindowTitle(f"Macan PDF Tools — Standalone (v{APP_VERSION})")
        self.resize(1100, 760)

        central = QWidget()
        self.setCentralWidget(central)
        root = QHBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # left nav
        self.nav_list = QListWidget()
        self.nav_list.setObjectName("navList")
        self.nav_list.setFixedWidth(210)
        self.nav_list.addItems([
            "Image to PDF",
            "PDF to Image",
            "PDF Merger",
            "PDF Document Conversion",
        ])
        self.nav_list.currentRowChanged.connect(self._on_nav_changed)
        root.addWidget(self.nav_list)

        # stacked pages
        self.stack = QStackedWidget()
        self.stack.addWidget(ImageToPdfPage())
        self.stack.addWidget(PdfToImagePage())
        self.stack.addWidget(PdfMergerPage())
        self.stack.addWidget(PdfDocConversionPage())
        root.addWidget(self.stack, 1)

        self.nav_list.setCurrentRow(0)

        self._apply_stylesheet()
        self._check_dependencies()

    def _on_nav_changed(self, index):
        if index >= 0:
            self.stack.setCurrentIndex(index)

    def _check_dependencies(self):
        missing = []
        if not HAS_PDFIUM:
            missing.append("pypdfium2 (WAJIB untuk PDF to Image / Merger / Document Conversion)")
        if missing:
            QMessageBox.warning(self, "Dependency hilang",
                                 "Beberapa fitur tidak akan berfungsi:\n\n" + "\n".join(missing) +
                                 "\n\nInstall dengan:\npip install " +
                                 " ".join(m.split(" ")[0] for m in missing) + " --break-system-packages")

    def _apply_stylesheet(self):
        self.setStyleSheet("""
            QMainWindow, QWidget {
                background-color: #252525;
                color: #E0E0E0;
                font-family: Segoe UI, sans-serif;
            }
            QLabel { font-size: 10pt; }
            #navList {
                background-color: #333333;
                border: none;
                border-right: 1px solid #444444;
            }
            #navList::item {
                padding: 12px 10px;
                border-bottom: 1px solid #3a3a3a;
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
