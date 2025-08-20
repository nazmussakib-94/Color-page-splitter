"""
Color/BW Print Splitter – Desktop Helper App (Tkinter)

What it does
- Opens a PDF or Word (.docx) file
- Detects which pages contain color (vs. grayscale)
- Lets you review and toggle page classifications
- Exports two PDFs: one with color pages, one with black/white pages
- Also exports a CSV report listing: page number, classification, dominant color, percent color area, and a ratio vs. a typical A4 text page coverage baseline

Dependencies (install with pip)
    pip install pymupdf PyPDF2 pillow docx2pdf

Notes
- PDF handling & rendering: PyMuPDF (fitz)
- Writing PDFs: PyPDF2
- Optional: docx2pdf (Windows with MS Word installed, or macOS w/ Word)
- If docx2pdf is unavailable, please export your DOCX to PDF first and open the PDF.

Build a single-file EXE (Windows) (optional):
    pip install pyinstaller
    pyinstaller --noconsole --onefile --name PrintSplitter print_splitter.py

"""
from __future__ import annotations
import os
import sys
import tempfile
import threading
from dataclasses import dataclass
from typing import List, Tuple
import csv

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Third-party
try:
    import fitz  # PyMuPDF
except Exception as e:
    raise SystemExit("PyMuPDF (fitz) is required. Install with: pip install pymupdf")

try:
    from PyPDF2 import PdfReader, PdfWriter
except Exception as e:
    raise SystemExit("PyPDF2 is required. Install with: pip install PyPDF2")

try:
    from PIL import Image
except Exception:
    raise SystemExit("Pillow is required. Install with: pip install pillow")

# Optional – for DOCX -> PDF
try:
    from docx2pdf import convert as docx2pdf_convert
    DOCX2PDF_AVAILABLE = True
except Exception:
    DOCX2PDF_AVAILABLE = False


@dataclass
class PageInfo:
    index: int                 # 0-based index
    is_color: bool             # current classification
    detected_color: bool       # original auto-detect result (for reference)
    colored_ppm: int           # colored pixels per million
    colored_pct: float         # colored pixels as % of page area
    dominant_color: str        # e.g., Red/Green/Blue/Cyan/Magenta/Yellow/Gray
    mean_hsv: Tuple[float, float, float]  # (H,S,V) avg of colored pixels (H 0..360)
    cmyk_pct: Tuple[float, float, float, float]   # absolute page coverage % per channel
    cmyk_norm: Tuple[float, float, float, float]  # normalized CMYK % that sum to 100 if any ink360)


class PrintSplitterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Color/BW Print Splitter")
        self.geometry("980x600")
        self.minsize(760, 500)

        self.file_path: str | None = None
        self.temp_pdf_from_docx: str | None = None
        self.pages: List[PageInfo] = []

        # Detection parameters (tunable)
        self.color_delta = tk.IntVar(value=18)  # 0–100 (bigger = stricter color difference)
        self.min_color_ratio_ppm = tk.IntVar(value=0)  # colored pixels per million required
        self.render_dpi = tk.IntVar(value=120)  # rendering DPI per page
        # Baseline coverage for a "typical A4 text page" (industry rule-of-thumb ~5%)
        self.text_baseline_pct = tk.DoubleVar(value=5.0)

        self._build_ui()

    # ---------------------------- UI ----------------------------
    def _build_ui(self):
        topbar = ttk.Frame(self)
        topbar.pack(fill=tk.X, padx=10, pady=10)

        self.path_var = tk.StringVar(value="No file selected")
        ttk.Button(topbar, text="Open PDF / DOCX…", command=self.on_open).pack(side=tk.LEFT)
        ttk.Label(topbar, textvariable=self.path_var).pack(side=tk.LEFT, padx=10)

        # Controls frame
        ctl = ttk.LabelFrame(self, text="Detection Settings")
        ctl.pack(fill=tk.X, padx=10, pady=(0,10))

        s1 = ttk.Frame(ctl); s1.pack(fill=tk.X, padx=10, pady=4)
        ttk.Label(s1, text="Color sensitivity (ΔRGB):").pack(side=tk.LEFT)
        tk.Scale(s1, from_=0, to=100, orient=tk.HORIZONTAL, variable=self.color_delta, length=260).pack(side=tk.LEFT, padx=8)
        ttk.Label(s1, text="(higher = pickier)").pack(side=tk.LEFT)

        s2 = ttk.Frame(ctl); s2.pack(fill=tk.X, padx=10, pady=4)
        ttk.Label(s2, text="Min colored pixel ratio (ppm):").pack(side=tk.LEFT)
        tk.Scale(s2, from_=1, to=5000, orient=tk.HORIZONTAL, variable=self.min_color_ratio_ppm, length=260).pack(side=tk.LEFT, padx=8)
        ttk.Label(s2, text="(e.g., 800 ≈ 0.08%)").pack(side=tk.LEFT)

        s3 = ttk.Frame(ctl); s3.pack(fill=tk.X, padx=10, pady=4)
        ttk.Label(s3, text="Render DPI:").pack(side=tk.LEFT)
        tk.Scale(s3, from_=72, to=200, orient=tk.HORIZONTAL, variable=self.render_dpi, length=260).pack(side=tk.LEFT, padx=8)
        ttk.Label(s3, text="(higher = more accurate, slower)").pack(side=tk.LEFT)

        s4 = ttk.Frame(ctl); s4.pack(fill=tk.X, padx=10, pady=4)
        ttk.Label(s4, text="Baseline text coverage (% of page):").pack(side=tk.LEFT)
        tk.Spinbox(s4, from_=1.0, to=50.0, increment=0.5, textvariable=self.text_baseline_pct, width=6).pack(side=tk.LEFT, padx=8)
        ttk.Label(s4, text="(A4 text ~5%)").pack(side=tk.LEFT)

        
        # Save and Reset Settings Buttons
        def save_settings():
            with open("saved_settings.txt", "w") as f:
                f.write(f"{self.color_delta.get()},{self.min_color_ratio_ppm.get()},{self.render_dpi.get()},{self.text_baseline_pct.get()}")

        def reset_settings():
            self.color_delta.set(18)
            self.min_color_ratio_ppm.set(0)
            self.render_dpi.set(120)
            self.text_baseline_pct.set(5.0)

        ttk.Button(ctl, text="Save Settings", command=save_settings).pack(side=tk.RIGHT, padx=5)
        ttk.Button(ctl, text="Reset Settings", command=reset_settings).pack(side=tk.RIGHT)

# Actions
        actions = ttk.Frame(self)
        actions.pack(fill=tk.X, padx=10, pady=(0,10))
        self.analyze_btn = ttk.Button(actions, text="Analyze Pages", command=self.on_analyze, state=tk.DISABLED)
        self.analyze_btn.pack(side=tk.LEFT)
        ttk.Button(actions, text="Select All as Color", command=lambda: self._bulk_set(True)).pack(side=tk.LEFT, padx=5)
        ttk.Button(actions, text="Select All as BW", command=lambda: self._bulk_set(False)).pack(side=tk.LEFT)
        self.export_btn = ttk.Button(actions, text="Export Split PDFs + CSV…", command=self.on_export, state=tk.DISABLED)
        self.export_btn.pack(side=tk.RIGHT)

        # Table
        table_frame = ttk.Frame(self)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0,10))

        self.tree = ttk.Treeview(table_frame, columns=("page","status","auto","dominant","color%","rel"), show="headings")
        self.tree.heading("page", text="Page")
        self.tree.heading("status", text="Classification (double-click to toggle)")
        self.tree.heading("auto", text="Auto-detected")
        self.tree.heading("dominant", text="Dominant Color")
        self.tree.heading("color%", text="Color Area %")
        self.tree.heading("rel", text="× vs Text(%)")
        self.tree.column("page", width=60, anchor=tk.CENTER)
        self.tree.column("status", width=200)
        self.tree.column("auto", width=120, anchor=tk.CENTER)
        self.tree.column("dominant", width=140, anchor=tk.CENTER)
        self.tree.column("color%", width=110, anchor=tk.CENTER)
        self.tree.column("rel", width=120, anchor=tk.CENTER)
        self.tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)

        # --- Preview panel ---
        self.preview_frame = ttk.LabelFrame(table_frame, text="Preview")
        self.preview_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 0), ipadx=10, ipady=10)
        self.preview_label = ttk.Label(self.preview_frame, anchor=tk.CENTER, text="Select a page to preview")
        self.preview_label.pack(fill=tk.BOTH, expand=True)
        self.tree.bind("<Double-1>", self.on_toggle_row)
        self.tree.bind("<<TreeviewSelect>>", self._on_row_select)

        self.tree.bind("<<TreeviewSelect>>", self._on_row_select)

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        self.status_var = tk.StringVar(value="Ready")
        statusbar = ttk.Label(self, textvariable=self.status_var, anchor=tk.W)
        statusbar.pack(fill=tk.X, padx=10, pady=(0,10))

    # ---------------------------- Events ----------------------------
    def preview_selected_page(self):
        sel = self.tree.selection()
        if not sel:
            self.preview_label.configure(text="No page selected.")
            return
        try:
            idx = int(self.tree.set(sel[0], "page")) - 1
            self._show_preview(idx)
        except Exception as e:
            self.preview_label.configure(text=f"Preview failed: {e}")

    
    
    def _show_preview(self, index: int):
        try:
            import fitz
            from PIL import Image, ImageTk

            self.preview_label.configure(text=f"Loading preview for page {index+1}...")

            pdf_path = self.temp_pdf_from_docx if self.temp_pdf_from_docx else self.file_path
            if not pdf_path:
                self.preview_label.configure(text="No PDF path available.")
                return

            doc = fitz.open(pdf_path)
            page = doc.load_page(index)
            mat = fitz.Matrix(2.0, 2.0)  # 144 DPI
            pix = page.get_pixmap(matrix=mat, alpha=False)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

            # Resize image to fit within A4-proportioned preview area (approx 595x842 points at 72 DPI)
            preview_width = 500
            preview_height = int(preview_width * (11.7 / 8.27))  # A4 aspect ratio

            img = img.resize((preview_width, preview_height), Image.LANCZOS)

            self._preview_photo = ImageTk.PhotoImage(img)
            self.preview_label.configure(image=self._preview_photo, text="")
            self.preview_label.image = self._preview_photo
        except Exception as e:
            self.preview_label.configure(text=f"Preview error: {e}")

    def on_open(self):
        fp = filedialog.askopenfilename(
            title="Open PDF or DOCX",
            filetypes=[
                ("PDF files", "*.pdf"),
                ("Word (DOCX)", "*.docx"),
                ("All supported", "*.pdf *.docx"),
                ("All files", "*.*"),
            ],
        )
        if not fp:
            return

        self._cleanup_temp_docx_pdf()

        # If DOCX, convert to PDF first
        ext = os.path.splitext(fp)[1].lower()
        if ext == ".docx":
            if not DOCX2PDF_AVAILABLE:
                messagebox.showwarning(
                    "DOCX handling",
                    "docx2pdf is not available. Please export your DOCX to PDF first and open the PDF.\n\nInstall with: pip install docx2pdf",
                )
                return
            try:
                self.status_var.set("Converting DOCX to PDF…")
                self.update_idletasks()
                tempdir = tempfile.mkdtemp(prefix="printsplit_")
                out_pdf = os.path.join(tempdir, "converted.pdf")
                docx2pdf_convert(fp, out_pdf)
                self.temp_pdf_from_docx = out_pdf
                self.file_path = out_pdf
            except Exception as e:
                messagebox.showerror("DOCX Conversion Failed", f"Could not convert DOCX to PDF.\n{e}")
                self.status_var.set("Ready")
                return
        else:
            self.file_path = fp

        self.path_var.set(self.file_path)
        self.analyze_btn.config(state=tk.NORMAL)
        self.export_btn.config(state=tk.DISABLED)
        self.pages.clear()
        self._refresh_table()
        self.status_var.set("File loaded. Click ‘Analyze Pages’.")

    def on_analyze(self):
        if not self.file_path:
            return
        self.analyze_btn.config(state=tk.DISABLED)
        self.export_btn.config(state=tk.DISABLED)
        self.status_var.set("Analyzing… This may take a moment for long PDFs.")

        t = threading.Thread(target=self._analyze_worker, daemon=True)
        t.start()

    def on_toggle_row(self, event):
        item = self.tree.identify_row(event.y)
        if not item:
            return
        idx = int(self.tree.set(item, "page")) - 1
        for p in self.pages:
            if p.index == idx:
                p.is_color = not p.is_color
                break
        self._refresh_table(select_index=idx)

    def on_export(self):
        if not self.file_path or not self.pages:
            return
        base_dir = os.path.dirname(self.file_path)
        base_name = os.path.splitext(os.path.basename(self.file_path))[0]

        out_color = filedialog.asksaveasfilename(
            title="Save COLOR pages PDF as…",
            defaultextension=".pdf",
            initialdir=base_dir,
            initialfile=f"{base_name}_COLOR.pdf",
            filetypes=[("PDF", "*.pdf")],
        )
        if not out_color:
            return
        out_bw = filedialog.asksaveasfilename(
            title="Save BW pages PDF as…",
            defaultextension=".pdf",
            initialdir=base_dir,
            initialfile=f"{base_name}_BW.pdf",
            filetypes=[("PDF", "*.pdf")],
        )
        if not out_bw:
            return
        out_csv = filedialog.asksaveasfilename(
            title="Save CSV report as…",
            defaultextension=".csv",
            initialdir=base_dir,
            initialfile=f"{base_name}_PageColorReport.csv",
            filetypes=[("CSV", "*.csv")],
        )
        if not out_csv:
            return

        try:
            self._export_split_pdfs(out_color, out_bw)
            self._export_csv(out_csv)
            messagebox.showinfo("Done", f"Exported:\n• {out_color}\n• {out_bw}\n• {out_csv}")
        except Exception as e:
            messagebox.showerror("Export Failed", str(e))

    # ---------------------------- Core Logic ----------------------------
    def _analyze_worker(self):
        try:
            doc = fitz.open(self.file_path)
            dpi = max(72, int(self.render_dpi.get()))
            scale = dpi / 72.0
            mat = fitz.Matrix(scale, scale)
            delta = int(self.color_delta.get())
            ppm_threshold = max(1, int(self.min_color_ratio_ppm.get()))  # colored pixels per million

            pages: List[PageInfo] = []
            for i, page in enumerate(doc):
                # Render to pixmap (RGB)
                pix = page.get_pixmap(matrix=mat, colorspace=fitz.csRGB, alpha=False)
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

                stats = self._page_color_stats(img, delta=delta, ppm_threshold=ppm_threshold)
                pages.append(PageInfo(index=i,
                                      is_color=stats["is_color"],
                                      detected_color=stats["is_color"],
                                      colored_ppm=stats["colored_ppm"],
                                      colored_pct=stats["colored_pct"],
                                      dominant_color=stats["dominant_color"],
                                      mean_hsv=(stats["mean_h"], stats["mean_s"], stats["mean_v"]),
                                      cmyk_pct=(stats["c_pct"], stats["m_pct"], stats["y_pct"], stats["k_pct"]),
                                      cmyk_norm=(stats["c_norm"], stats["m_norm"], stats["y_norm"], stats["k_norm"])) )

            self.pages = pages
            self.after(0, self._on_analyze_done)
        except Exception as e:
            self.after(0, lambda: self._on_analyze_failed(e))

    def _on_analyze_done(self):
        self._refresh_table()
        self.analyze_btn.config(state=tk.NORMAL)
        self.export_btn.config(state=tk.NORMAL)
        col_pages = sum(1 for p in self.pages if p.is_color)
        bw_pages = len(self.pages) - col_pages
        self.status_var.set(f"Analysis complete: {col_pages} color, {bw_pages} BW.")

    def _on_analyze_failed(self, err: Exception):
        self.analyze_btn.config(state=tk.NORMAL)
        self.export_btn.config(state=tk.DISABLED)
        messagebox.showerror("Analyze Failed", str(err))
        self.status_var.set("Ready")

    def _export_split_pdfs(self, out_color: str, out_bw: str):
        reader = PdfReader(self.file_path)
        w_color = PdfWriter()
        w_bw = PdfWriter()

        for p in self.pages:
            if p.is_color:
                w_color.add_page(reader.pages[p.index])
            else:
                w_bw.add_page(reader.pages[p.index])

        with open(out_color, "wb") as f:
            w_color.write(f)
        with open(out_bw, "wb") as f:
            w_bw.write(f)

    def _export_csv(self, out_csv: str):
        baseline = max(0.1, float(self.text_baseline_pct.get()))
        with open(out_csv, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["page_number","classification","auto_detected","dominant_color","colored_ppm","colored_pct","relative_to_text_baseline","mean_h","mean_s","mean_v","C_pct","M_pct","Y_pct","K_pct","C_norm","M_norm","Y_norm","K_norm"])
            for p in self.pages:
                rel = (p.colored_pct / baseline) if baseline > 0 else 0.0
                writer.writerow([
                    p.index + 1,
                    "COLOR" if p.is_color else "BW",
                    "COLOR" if p.detected_color else "BW",
                    p.dominant_color,
                    p.colored_ppm,
                    round(p.colored_pct, 6),
                    round(rel, 4),
                    round(p.mean_hsv[0], 2),
                    round(p.mean_hsv[1], 4),
                    round(p.mean_hsv[2], 4),
                    round(p.cmyk_pct[0], 6),
                    round(p.cmyk_pct[1], 6),
                    round(p.cmyk_pct[2], 6),
                    round(p.cmyk_pct[3], 6),
                    round(p.cmyk_norm[0], 4),
                    round(p.cmyk_norm[1], 4),
                    round(p.cmyk_norm[2], 4),
                    round(p.cmyk_norm[3], 4),
                ])

    def _bulk_set(self, is_color: bool):
        if not self.pages:
            return
        for p in self.pages:
            p.is_color = is_color
        self._refresh_table()

    def _refresh_table(self, select_index: int | None = None):
        for i in self.tree.get_children():
            self.tree.delete(i)
        baseline = max(0.1, float(self.text_baseline_pct.get()))
        for p in self.pages:
            rel = (p.colored_pct / baseline) if baseline > 0 else 0.0
            self.tree.insert(
                "", tk.END,
                values=(p.index + 1,
                        "COLOR" if p.is_color else "BW",
                        "COLOR" if p.detected_color else "BW",
                        p.dominant_color,
                        f"{p.colored_pct:.3f}",
                        f"{rel:.2f}×"),
            )
        if select_index is not None:
            children = self.tree.get_children()
            if 0 <= select_index < len(children):
                self.tree.selection_set(children[select_index])
                self.tree.see(children[select_index])

    # ---------------------------- Color Stats ----------------------------
    @staticmethod
    def _page_color_stats(img: Image.Image, *, delta: int, ppm_threshold: int):
        """
        Compute color stats and CMYK coverage.
        Returns dict with keys:
          - is_color (bool) using ppm_threshold
          - colored_ppm (int), colored_pct (float 0-100)
          - dominant_color (str)
          - mean_h, mean_s, mean_v (over colored pixels)
          - c_pct, m_pct, y_pct, k_pct  (absolute page coverage % for each channel)
          - c_norm, m_norm, y_norm, k_norm (normalized % that sum to 100 when any ink present)
        """
        if img.mode != "RGB":
            img = img.convert("RGB")
        w, h = img.size
        total = max(1, w * h)

        # Downsample for speed
        max_side = 1800
        if max(w, h) > max_side:
            factor = max_side / float(max(w, h))
            new_size = (int(w * factor), int(h * factor))
            img = img.resize(new_size, Image.BILINEAR)
            w, h = img.size
            total = max(1, w * h)

        px = img.tobytes()
        thr = int(delta)
        WHITE_CUTOFF = 250

        colored = 0
        hsv_h_total = 0.0
        hsv_s_total = 0.0
        hsv_v_total = 0.0

        def rgb_to_hsv_deg(r, g, b):
            r1, g1, b1 = r/255.0, g/255.0, b/255.0
            mx = max(r1, g1, b1); mn = min(r1, g1, b1)
            diff = mx - mn
            if diff == 0:
                h = 0.0
            elif mx == r1:
                h = (60 * ((g1 - b1) / diff) + 360) % 360
            elif mx == g1:
                h = (60 * ((b1 - r1) / diff) + 120) % 360
            else:
                h = (60 * ((r1 - g1) / diff) + 240) % 360
            s = 0.0 if mx == 0 else diff / mx
            v = mx
            return h, s, v

        # Color detection over RGB space
        for i in range(0, len(px), 3):
            r = px[i]; g = px[i+1]; b = px[i+2]
            mx = r if r > g else g
            if b > mx: mx = b
            mn = r if r < g else g
            if b < mn: mn = b
            if mx < WHITE_CUTOFF and (mx - mn) > thr:
                colored += 1
                h, s, v = rgb_to_hsv_deg(r, g, b)
                hsv_h_total += h; hsv_s_total += s; hsv_v_total += v

        ppm = int(colored * 1_000_000 / total)
        pct = (colored / total) * 100.0
        is_color = ppm >= max(1, ppm_threshold)

        if colored == 0:
            dominant = "Gray"
            mean_h = 0.0; mean_s = 0.0; mean_v = 0.0
        else:
            mean_h = hsv_h_total / colored; mean_s = hsv_s_total / colored; mean_v = hsv_v_total / colored
            # simple mapping by hue/sat
            if mean_s < 0.15:
                dominant = "Gray"
            else:
                h = mean_h
                if (h >= 345 or h < 15): dominant = "Red"
                elif 15 <= h < 45:       dominant = "Orange"
                elif 45 <= h < 75:       dominant = "Yellow"
                elif 75 <= h < 165:      dominant = "Green"
                elif 165 <= h < 195:     dominant = "Cyan"
                elif 195 <= h < 255:     dominant = "Blue"
                elif 255 <= h < 300:     dominant = "Purple"
                else:                     dominant = "Magenta"

        # --- CMYK coverage ---
        # Use Pillow's device conversion; values in 0..255 per channel
        cmyk = img.convert("CMYK").tobytes()
        sum_c = sum_m = sum_y = sum_k = 0
        for i in range(0, len(cmyk), 4):
            sum_c += cmyk[i]
            sum_m += cmyk[i+1]
            sum_y += cmyk[i+2]
            sum_k += cmyk[i+3]
        denom = (w * h * 255.0)
        c_pct = (sum_c / denom) * 100.0
        m_pct = (sum_m / denom) * 100.0
        y_pct = (sum_y / denom) * 100.0
        k_pct = (sum_k / denom) * 100.0
        total_ink = c_pct + m_pct + y_pct + k_pct
        if total_ink > 1e-9:
            c_norm = (c_pct / total_ink) * 100.0
            m_norm = (m_pct / total_ink) * 100.0
            y_norm = (y_pct / total_ink) * 100.0
            k_norm = (k_pct / total_ink) * 100.0
        else:
            c_norm = m_norm = y_norm = k_norm = 0.0

        return {
            "is_color": is_color,
            "colored_ppm": ppm,
            "colored_pct": pct,
            "dominant_color": dominant,
            "mean_h": mean_h,
            "mean_s": mean_s,
            "mean_v": mean_v,
            "c_pct": c_pct,
            "m_pct": m_pct,
            "y_pct": y_pct,
            "k_pct": k_pct,
            "c_norm": c_norm,
            "m_norm": m_norm,
            "y_norm": y_norm,
            "k_norm": k_norm,
        }

    # ---------------------------- Utilities ----------------------------
    def _cleanup_temp_docx_pdf(self):
        self.temp_pdf_from_docx = None



    def _on_row_select(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        try:
            idx = int(self.tree.set(sel[0], "page")) - 1
            self._show_preview(idx)
        except Exception:
            pass

def main():
    app = PrintSplitterApp()
    app.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import traceback
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        tb = traceback.format_exc()
        messagebox.showerror("Unhandled Error", f"An unexpected error occurred:\n\n{e}\n\n{tb}")



    
    def _on_row_select(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        try:
            idx = int(self.tree.set(sel[0], "page")) - 1
            self._show_preview(idx)
        except Exception:
            pass


    
    
     