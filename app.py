import tkinter as tk
from tkinter import ttk, filedialog, colorchooser, messagebox
import fitz  # PyMuPDF
import pandas as pd
from PIL import Image, ImageTk
import os
import io
import subprocess
import math
import time
import json

# ─── System Font Discovery ───────────────────────────────

def get_system_fonts():
    """Return a dict of {family_name: font_file_path} from fc-list."""
    fonts = {}
    try:
        result = subprocess.run(
            ["fc-list", "--format=%{family}|%{file}\n"],
            capture_output=True, text=True, timeout=10
        )
        for line in result.stdout.splitlines():
            line = line.strip()
            if "|" not in line:
                continue
            family, path = line.split("|", 1)
            family = family.split(",")[0].strip()
            if family and path and path.lower().endswith((".ttf", ".otf")):
                if family not in fonts:
                    fonts[family] = path
    except Exception:
        pass
    return dict(sorted(fonts.items(), key=lambda x: x[0].lower()))


# Interaction modes
INTERACT_NONE = 0
INTERACT_MOVE = 1
INTERACT_RESIZE = 2
INTERACT_ROTATE = 3


class CertificateApp(tk.Tk):
    # --- Class fields ---
    names = []
    current_index = 0
    df = None
    preview_image = None
    text_color = (0.003, 0.105, 0.329)
    generating = False
    system_fonts = {}
    # Preview state
    _preview_scale = 1.0
    _preview_page_width = 1
    _preview_page_height = 1
    _preview_img_x = 0
    _preview_img_y = 0
    # Interaction state
    _interact_mode = INTERACT_NONE
    _handles_active = False
    _drag_start_x = 0
    _drag_start_y = 0
    _drag_orig_xoffset = 0
    _drag_orig_texty = 0
    _drag_orig_fontsize = 0
    _drag_orig_rotation = 0
    _drag_orig_bbox_center = (0, 0)
    _last_render_time = 0
    _text_bbox_canvas = None

    def __init__(self):
        super().__init__()
        self.title("Certificate Generator")
        self.geometry("1200x750")
        self.minsize(900, 600)
        self.configure(bg="#1e1e2e")

        self.system_fonts = get_system_fonts()

        self.style = ttk.Style(self)
        self.style.theme_use("clam")
        self._configure_styles()

        self.build_ui()

        self.bind("<Left>", lambda e: self.navigate(-1))
        self.bind("<Right>", lambda e: self.navigate(1))

        self._auto_detect_files()

    def _configure_styles(self):
        bg = "#1e1e2e"
        fg = "#cdd6f4"
        accent = "#89b4fa"
        entry_bg = "#313244"
        btn_bg = "#45475a"
        btn_active = "#585b70"

        self.style.configure("TFrame", background=bg)
        self.style.configure("TLabel", background=bg, foreground=fg, font=("Inter", 10))
        self.style.configure("Header.TLabel", background=bg, foreground=accent, font=("Inter", 13, "bold"))
        self.style.configure("Section.TLabel", background=bg, foreground="#a6adc8", font=("Inter", 9, "bold"))
        self.style.configure("TButton", background=btn_bg, foreground=fg, font=("Inter", 10), borderwidth=0, padding=6)
        self.style.map("TButton", background=[("active", btn_active)])
        self.style.configure("Accent.TButton", background=accent, foreground="#1e1e2e", font=("Inter", 11, "bold"), padding=10)
        self.style.map("Accent.TButton", background=[("active", "#74c7ec")])
        self.style.configure("Nav.TButton", background=btn_bg, foreground=fg, font=("Inter", 10), padding=4)
        self.style.configure("TEntry", fieldbackground=entry_bg, foreground=fg, insertcolor=fg, borderwidth=0)
        self.style.configure("TSpinbox", fieldbackground=entry_bg, foreground=fg, insertcolor=fg, borderwidth=0, arrowcolor=fg)
        self.style.configure("TCombobox", fieldbackground=entry_bg, foreground=fg, borderwidth=0)
        self.style.map("TCombobox", fieldbackground=[("readonly", entry_bg)])
        self.style.configure("green.Horizontal.TProgressbar", troughcolor=entry_bg, background="#a6e3a1", borderwidth=0)
        self.style.configure("TRadiobutton", background=bg, foreground=fg, font=("Inter", 9))
        self.style.map("TRadiobutton", background=[("active", bg)])

    def build_ui(self):
        # Top bar with help button
        topbar = ttk.Frame(self)
        topbar.pack(fill=tk.X, padx=12, pady=(8, 0))
        ttk.Label(topbar, text="", style="TLabel").pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.help_btn = tk.Button(topbar, text="?", font=("Inter", 12, "bold"),
            bg="#45475a", fg="#cdd6f4", activebackground="#585b70", activeforeground="#cdd6f4",
            bd=0, width=3, height=1, cursor="hand2", command=self._start_tutorial)
        self.help_btn.pack(side=tk.RIGHT)

        main = ttk.Frame(self)
        main.pack(fill=tk.BOTH, expand=True, padx=12, pady=(4, 12))

        left = ttk.Frame(main, width=350)
        left.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 12))
        left.pack_propagate(False)

        right = ttk.Frame(main)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._build_settings_panel(left)
        self._build_preview_panel(right)

    # ─── Settings Panel ───────────────────────────────────────

    def _build_settings_panel(self, parent):
        self._settings_canvas = tk.Canvas(parent, bg="#1e1e2e", highlightthickness=0)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=self._settings_canvas.yview)
        self.settings_inner = ttk.Frame(self._settings_canvas)

        self.settings_inner.bind(
            "<Configure>",
            lambda e: self._settings_canvas.configure(scrollregion=self._settings_canvas.bbox("all"))
        )
        self._settings_canvas.create_window((0, 0), window=self.settings_inner, anchor="nw", width=336)
        self._settings_canvas.configure(yscrollcommand=scrollbar.set)

        self._settings_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        def _on_mousewheel_linux(event):
            if event.num == 4:
                self._settings_canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                self._settings_canvas.yview_scroll(1, "units")

        def _bind_scroll(event):
            self._settings_canvas.bind_all("<Button-4>", _on_mousewheel_linux)
            self._settings_canvas.bind_all("<Button-5>", _on_mousewheel_linux)

        def _unbind_scroll(event):
            self._settings_canvas.unbind_all("<Button-4>")
            self._settings_canvas.unbind_all("<Button-5>")

        self._settings_canvas.bind("<Enter>", _bind_scroll)
        self._settings_canvas.bind("<Leave>", _unbind_scroll)

        inner = self.settings_inner

        ttk.Label(inner, text="⚙  Settings", style="Header.TLabel").pack(anchor="w", pady=(0, 12))

        # --- File selectors ---
        ttk.Label(inner, text="FILES", style="Section.TLabel").pack(anchor="w", pady=(8, 4))

        self.template_var = tk.StringVar()
        self._file_row(inner, "Template PDF", self.template_var, self._browse_template)

        self.excel_var = tk.StringVar()
        self._file_row(inner, "Excel File", self.excel_var, self._browse_excel)

        self.outdir_var = tk.StringVar(value="sertifikalar")
        self._file_row(inner, "Output Dir", self.outdir_var, self._browse_outdir)

        ttk.Frame(inner, height=2).pack(fill=tk.X, pady=12)

        # --- Font options ---
        ttk.Label(inner, text="FONT", style="Section.TLabel").pack(anchor="w", pady=(0, 4))

        self.font_source_var = tk.StringVar(value="file")
        font_src_frame = ttk.Frame(inner)
        font_src_frame.pack(fill=tk.X, pady=2)
        ttk.Radiobutton(font_src_frame, text="Font File", variable=self.font_source_var, value="file",
                        command=self._on_font_source_change).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Radiobutton(font_src_frame, text="System Font", variable=self.font_source_var, value="system",
                        command=self._on_font_source_change).pack(side=tk.LEFT)

        # Container for font file / system font
        self.font_selector_container = ttk.Frame(inner)
        self.font_selector_container.pack(fill=tk.X, pady=2)

        self.font_var = tk.StringVar()
        self.font_file_frame = ttk.Frame(self.font_selector_container)
        self.font_file_frame.pack(fill=tk.X)
        ttk.Label(self.font_file_frame, text="File", width=11).pack(side=tk.LEFT)
        ttk.Entry(self.font_file_frame, textvariable=self.font_var, width=14).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        ttk.Button(self.font_file_frame, text="...", width=3, command=self._browse_font).pack(side=tk.LEFT)

        self.system_font_var = tk.StringVar()
        self.system_font_frame = ttk.Frame(self.font_selector_container)
        ttk.Label(self.system_font_frame, text="Family", width=11).pack(side=tk.LEFT)
        font_families = list(self.system_fonts.keys())
        self.system_font_combo = ttk.Combobox(self.system_font_frame, textvariable=self.system_font_var,
                                               values=font_families, state="readonly", width=18)
        self.system_font_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.system_font_var.trace_add("write", lambda *_: self._on_setting_change())

        ttk.Frame(inner, height=2).pack(fill=tk.X, pady=12)

        # --- Text options ---
        ttk.Label(inner, text="TEXT OPTIONS", style="Section.TLabel").pack(anchor="w", pady=(0, 4))

        self.fontsize_var = tk.StringVar(value="35.2")
        self._spin_row(inner, "Font Size", self.fontsize_var, 1.0, 100.0, 0.5)

        self.texty_var = tk.StringVar(value="307")
        self._spin_row(inner, "Text Y", self.texty_var, 0, 2000, 1)

        self.xoffset_var = tk.StringVar(value="-100")
        self._spin_row(inner, "X Offset", self.xoffset_var, -1000, 1000, 5)

        self.linespace_var = tk.StringVar(value="45")
        self._spin_row(inner, "Line Space", self.linespace_var, 0, 200, 1)

        self.rotation_var = tk.StringVar(value="0")
        self._spin_row(inner, "Rotation °", self.rotation_var, -180, 180, 1)

        # Snap to grid checkbox
        self.snap_rotation_var = tk.BooleanVar(value=False)
        snap_frame = ttk.Frame(inner)
        snap_frame.pack(fill=tk.X, pady=(0, 3))
        ttk.Checkbutton(snap_frame, text="Snap to 15°", variable=self.snap_rotation_var,
                        command=self._on_snap_rotation_toggle).pack(side=tk.LEFT, padx=(88, 0))
        self.style.configure("TCheckbutton", background="#1e1e2e", foreground="#cdd6f4", font=("Inter", 9))
        self.style.map("TCheckbutton", background=[("active", "#1e1e2e")])

        # Color picker row
        color_frame = ttk.Frame(inner)
        color_frame.pack(fill=tk.X, pady=3)
        ttk.Label(color_frame, text="Color", width=11).pack(side=tk.LEFT)
        self.color_swatch = tk.Canvas(color_frame, width=28, height=28, highlightthickness=1, highlightbackground="#585b70", cursor="hand2")
        self.color_swatch.pack(side=tk.LEFT, padx=(0, 6))
        self._update_swatch()
        self.color_swatch.bind("<Button-1>", lambda e: self.pick_color())
        self.color_hex_label = ttk.Label(color_frame, text=self._color_to_hex(), font=("Inter", 9))
        self.color_hex_label.pack(side=tk.LEFT)

        ttk.Frame(inner, height=2).pack(fill=tk.X, pady=12)

        # --- Text Alignment ---
        ttk.Label(inner, text="ALIGNMENT", style="Section.TLabel").pack(anchor="w", pady=(0, 4))

        self.alignment_var = tk.StringVar(value="center")
        align_frame = ttk.Frame(inner)
        align_frame.pack(fill=tk.X, pady=2)
        for val, label in [("left", "Left"), ("center", "Center"), ("right", "Right")]:
            ttk.Radiobutton(align_frame, text=label, variable=self.alignment_var, value=val,
                            command=self._on_setting_change).pack(side=tk.LEFT, padx=(0, 10))

        ttk.Frame(inner, height=2).pack(fill=tk.X, pady=12)

        # --- Text Splitting ---
        ttk.Label(inner, text="TEXT SPLITTING", style="Section.TLabel").pack(anchor="w", pady=(0, 4))

        self.split_mode_var = tk.StringVar(value="auto")
        split_frame = ttk.Frame(inner)
        split_frame.pack(fill=tk.X, pady=2)
        ttk.Radiobutton(split_frame, text="Auto", variable=self.split_mode_var, value="auto",
                        command=self._on_setting_change).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Radiobutton(split_frame, text="No Split", variable=self.split_mode_var, value="none",
                        command=self._on_setting_change).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Radiobutton(split_frame, text="Always", variable=self.split_mode_var, value="always",
                        command=self._on_setting_change).pack(side=tk.LEFT)

        self.split_threshold_var = tk.StringVar(value="19")
        self._spin_row(inner, "Threshold", self.split_threshold_var, 1, 100, 1)

        ttk.Frame(inner, height=2).pack(fill=tk.X, pady=12)

        # --- Excel column ---
        ttk.Label(inner, text="DATA", style="Section.TLabel").pack(anchor="w", pady=(0, 4))

        col_frame = ttk.Frame(inner)
        col_frame.pack(fill=tk.X, pady=3)
        ttk.Label(col_frame, text="Name Col", width=11).pack(side=tk.LEFT)
        self.column_var = tk.StringVar()
        self.column_combo = ttk.Combobox(col_frame, textvariable=self.column_var, state="readonly", width=18)
        self.column_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)

        ttk.Frame(inner, height=2).pack(fill=tk.X, pady=12)

        # --- Generate ---
        self.gen_btn = ttk.Button(inner, text="▶  Generate All", style="Accent.TButton", command=self.generate_all)
        self.gen_btn.pack(fill=tk.X, pady=(4, 6))

        self.progress_var = tk.IntVar(value=0)
        self.progress_bar = ttk.Progressbar(inner, variable=self.progress_var, maximum=100, style="green.Horizontal.TProgressbar")
        self.progress_bar.pack(fill=tk.X, pady=(0, 4))

        self.status_label = ttk.Label(inner, text="Ready", font=("Inter", 9))
        self.status_label.pack(anchor="w")

        ttk.Frame(inner, height=2).pack(fill=tk.X, pady=12)

        # --- Save / Import Settings ---
        ttk.Label(inner, text="SETTINGS", style="Section.TLabel").pack(anchor="w", pady=(0, 4))
        settings_btns = ttk.Frame(inner)
        settings_btns.pack(fill=tk.X, pady=2)
        ttk.Button(settings_btns, text="💾 Save", command=self._save_settings).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        ttk.Button(settings_btns, text="📂 Import", command=self._import_settings).pack(side=tk.LEFT, fill=tk.X, expand=True)

    def _file_row(self, parent, label, var, command):
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.X, pady=2)
        ttk.Label(frame, text=label, width=11).pack(side=tk.LEFT)
        entry = ttk.Entry(frame, textvariable=var, width=14)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        ttk.Button(frame, text="...", width=3, command=command).pack(side=tk.LEFT)

    def _spin_row(self, parent, label, var, from_, to, increment):
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.X, pady=3)
        ttk.Label(frame, text=label, width=11).pack(side=tk.LEFT)
        spin = ttk.Spinbox(frame, textvariable=var, from_=from_, to=to, increment=increment, width=10)
        spin.pack(side=tk.LEFT, fill=tk.X, expand=True)
        var.trace_add("write", lambda *_: self._on_setting_change())

    # ─── Preview Panel ────────────────────────────────────────

    def _build_preview_panel(self, parent):
        ttk.Label(parent, text="📄  Certificate Preview", style="Header.TLabel").pack(anchor="w", pady=(0, 8))

        self.preview_frame = ttk.Frame(parent)
        self.preview_frame.pack(fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(self.preview_frame, bg="#181825", highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)

        nav = ttk.Frame(parent)
        nav.pack(fill=tk.X, pady=(8, 0))

        self.prev_btn = ttk.Button(nav, text="◀  Prev", style="Nav.TButton", command=lambda: self.navigate(-1))
        self.prev_btn.pack(side=tk.LEFT)

        self.nav_label = ttk.Label(nav, text="No data loaded", font=("Inter", 10))
        self.nav_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.next_btn = ttk.Button(nav, text="Next  ▶", style="Nav.TButton", command=lambda: self.navigate(1))
        self.next_btn.pack(side=tk.RIGHT)

        # Bind resize
        self.canvas.bind("<Configure>", lambda e: self.render_preview())

        # Interactive handle events
        self.canvas.bind("<ButtonPress-1>", self._on_handle_press)
        self.canvas.bind("<B1-Motion>", self._on_handle_motion)
        self.canvas.bind("<ButtonRelease-1>", self._on_handle_release)
        self.canvas.bind("<Motion>", self._on_hover)

    # ─── Coordinate Conversion ────────────────────────────────

    def _pdf_to_canvas(self, pdf_x, pdf_y):
        """Convert PDF page coordinates to canvas pixel coordinates."""
        if not self.preview_image:
            return 0, 0
        img_w = self.preview_image.width()
        img_h = self.preview_image.height()
        img_left = self._preview_img_x - img_w / 2
        img_top = self._preview_img_y - img_h / 2
        cx = img_left + pdf_x * self._preview_scale
        cy = img_top + pdf_y * self._preview_scale
        return cx, cy

    def _canvas_to_pdf(self, cx, cy):
        """Convert canvas pixel coordinates to PDF page coordinates."""
        if not self.preview_image:
            return 0, 0
        img_w = self.preview_image.width()
        img_h = self.preview_image.height()
        img_left = self._preview_img_x - img_w / 2
        img_top = self._preview_img_y - img_h / 2
        pdf_x = (cx - img_left) / self._preview_scale
        pdf_y = (cy - img_top) / self._preview_scale
        return pdf_x, pdf_y

    # ─── Text Bounding Box ────────────────────────────────────

    def _compute_text_bbox_pdf(self):
        """Compute the text bounding box in PDF coordinates. Returns (x1, y1, x2, y2) or None."""
        if not self.names:
            return None
        font_path = self._get_active_font_path()
        if not font_path or not os.path.exists(font_path):
            return None
        template_path = self.template_var.get()
        if not template_path or not os.path.exists(template_path):
            return None

        name = self.names[self.current_index]
        font_size, text_y, x_offset, line_spacing, rotation = self._get_settings()

        try:
            custom_font = fitz.Font(fontfile=font_path)
        except Exception:
            return None

        try:
            doc = fitz.open(template_path)
            page_width = doc[0].rect.width
            doc.close()
        except Exception:
            return None

        lines = self._split_name(name)
        alignment = self.alignment_var.get()

        min_x = float('inf')
        max_x = float('-inf')

        for i, line in enumerate(lines):
            text_width = custom_font.text_length(line, fontsize=font_size)
            if alignment == "center":
                pos_x = ((page_width - text_width) / 2) + x_offset
            elif alignment == "left":
                pos_x = x_offset if x_offset >= 0 else 50 + x_offset
            else:
                pos_x = page_width - text_width + x_offset
            min_x = min(min_x, pos_x)
            max_x = max(max_x, pos_x + text_width)

        # Y bounds
        if len(lines) > 1:
            first_y = text_y - (line_spacing / 2)
            last_y = text_y - (line_spacing / 2) + (len(lines) - 1) * line_spacing
        else:
            first_y = text_y
            last_y = text_y

        ascent = font_size * 0.82
        descent = font_size * 0.22
        top = first_y - ascent
        bottom = last_y + descent

        return (min_x, top, max_x, bottom)

    # ─── Draw Handles ─────────────────────────────────────────

    def _draw_handles(self):
        """Draw bounding box, corner handles, and rotation handle over the text."""
        self.canvas.delete("handle")

        # Always compute bbox for click detection, but only draw if active
        bbox_pdf = self._compute_text_bbox_pdf()
        if bbox_pdf is None:
            self._text_bbox_canvas = None
            return

        x1_raw, y1_raw, x2_raw, y2_raw = bbox_pdf
        cx1_raw, cy1_raw = self._pdf_to_canvas(x1_raw, y1_raw)
        cx2_raw, cy2_raw = self._pdf_to_canvas(x2_raw, y2_raw)
        pad = 8
        self._text_bbox_canvas = (cx1_raw - pad, cy1_raw - pad, cx2_raw + pad, cy2_raw + pad)

        if not self._handles_active:
            return

        cx1, cy1, cx2, cy2 = self._text_bbox_canvas

        # Dashed bounding box
        self.canvas.create_rectangle(cx1, cy1, cx2, cy2,
            outline="#89b4fa", width=1.5, dash=(5, 3), tags="handle")

        # Corner handles (small filled squares)
        hs = 5
        corners = [
            (cx1, cy1), (cx2, cy1),  # top-left, top-right
            (cx1, cy2), (cx2, cy2),  # bottom-left, bottom-right
        ]
        for x, y in corners:
            self.canvas.create_rectangle(x - hs, y - hs, x + hs, y + hs,
                fill="#89b4fa", outline="#cdd6f4", width=1, tags=("handle", "h_corner"))

        # Edge midpoint handles (smaller)
        ms = 4
        mid_x = (cx1 + cx2) / 2
        mid_y = (cy1 + cy2) / 2
        for x, y in [(mid_x, cy1), (mid_x, cy2), (cx1, mid_y), (cx2, mid_y)]:
            self.canvas.create_rectangle(x - ms, y - ms, x + ms, y + ms,
                fill="#89b4fa", outline="#cdd6f4", width=1, tags=("handle", "h_corner"))

        # Rotation handle (circle above top center, connected by a line)
        rot_line_len = 28
        rot_y = cy1 - rot_line_len
        # Connecting line
        self.canvas.create_line(mid_x, cy1, mid_x, rot_y,
            fill="#89b4fa", width=1.5, tags="handle")
        # Circle
        rs = 7
        self.canvas.create_oval(mid_x - rs, rot_y - rs, mid_x + rs, rot_y + rs,
            fill="#1e1e2e", outline="#89b4fa", width=2, tags=("handle", "h_rotate"))

    # ─── Handle Interaction ───────────────────────────────────

    def _on_hover(self, event):
        """Change cursor based on what's under the mouse."""
        if self._text_bbox_canvas is None:
            self.canvas.config(cursor="arrow")
            return

        x, y = event.x, event.y
        cx1, cy1, cx2, cy2 = self._text_bbox_canvas

        # If handles are NOT active, only show hand cursor when hovering over text area
        if not self._handles_active:
            if cx1 <= x <= cx2 and cy1 <= y <= cy2:
                self.canvas.config(cursor="hand2")
            else:
                self.canvas.config(cursor="arrow")
            return

        mid_x = (cx1 + cx2) / 2
        rot_y = cy1 - 28

        # Check rotation handle
        if math.hypot(x - mid_x, y - rot_y) < 12:
            self.canvas.config(cursor="exchange")
            return

        # Check corner/edge handles
        hs = 10
        corners = [
            (cx1, cy1), (cx2, cy1), (cx1, cy2), (cx2, cy2),
            (mid_x, cy1), (mid_x, cy2), (cx1, (cy1+cy2)/2), (cx2, (cy1+cy2)/2),
        ]
        for hx, hy in corners:
            if abs(x - hx) < hs and abs(y - hy) < hs:
                self.canvas.config(cursor="sizing")
                return

        # Check inside bbox
        if cx1 <= x <= cx2 and cy1 <= y <= cy2:
            self.canvas.config(cursor="fleur")
            return

        self.canvas.config(cursor="arrow")

    def _on_handle_press(self, event):
        """Determine interaction mode based on click location."""
        if self._text_bbox_canvas is None:
            self._interact_mode = INTERACT_NONE
            self._handles_active = False
            self.render_preview()
            return

        x, y = event.x, event.y
        cx1, cy1, cx2, cy2 = self._text_bbox_canvas
        mid_x = (cx1 + cx2) / 2
        rot_y = cy1 - 28

        # If handles are NOT active yet, check if clicking inside text bbox to activate + start moving
        if not self._handles_active:
            if cx1 <= x <= cx2 and cy1 <= y <= cy2:
                self._handles_active = True
                # Set up drag state so dragging starts immediately
                self._drag_start_x = x
                self._drag_start_y = y
                try:
                    self._drag_orig_xoffset = float(self.xoffset_var.get())
                except ValueError:
                    self._drag_orig_xoffset = -100
                try:
                    self._drag_orig_texty = float(self.texty_var.get())
                except ValueError:
                    self._drag_orig_texty = 307
                try:
                    self._drag_orig_fontsize = float(self.fontsize_var.get())
                except ValueError:
                    self._drag_orig_fontsize = 35.2
                try:
                    self._drag_orig_rotation = float(self.rotation_var.get())
                except ValueError:
                    self._drag_orig_rotation = 0
                self._drag_orig_bbox_center = ((cx1 + cx2) / 2, (cy1 + cy2) / 2)
                self._interact_mode = INTERACT_MOVE
                self.render_preview()
            else:
                self._interact_mode = INTERACT_NONE
            return

        # Handles ARE active — check if clicking outside to deactivate
        # Expanded area includes rotation handle above the bbox
        in_handles_area = (cx1 - 15 <= x <= cx2 + 15 and rot_y - 15 <= y <= cy2 + 15)
        if not in_handles_area:
            self._handles_active = False
            self._interact_mode = INTERACT_NONE
            self.render_preview()
            return

        # Store drag start state
        self._drag_start_x = x
        self._drag_start_y = y
        try:
            self._drag_orig_xoffset = float(self.xoffset_var.get())
        except ValueError:
            self._drag_orig_xoffset = -100
        try:
            self._drag_orig_texty = float(self.texty_var.get())
        except ValueError:
            self._drag_orig_texty = 307
        try:
            self._drag_orig_fontsize = float(self.fontsize_var.get())
        except ValueError:
            self._drag_orig_fontsize = 35.2
        try:
            self._drag_orig_rotation = float(self.rotation_var.get())
        except ValueError:
            self._drag_orig_rotation = 0

        self._drag_orig_bbox_center = ((cx1 + cx2) / 2, (cy1 + cy2) / 2)

        # Check rotation handle first
        if math.hypot(x - mid_x, y - rot_y) < 14:
            self._interact_mode = INTERACT_ROTATE
            bcx, bcy = self._drag_orig_bbox_center
            self._drag_start_angle = math.degrees(math.atan2(x - bcx, -(y - bcy)))
            return

        # Check corner/edge handles
        hs = 12
        corners = [
            (cx1, cy1), (cx2, cy1), (cx1, cy2), (cx2, cy2),
            (mid_x, cy1), (mid_x, cy2), (cx1, (cy1+cy2)/2), (cx2, (cy1+cy2)/2),
        ]
        for hx, hy in corners:
            if abs(x - hx) < hs and abs(y - hy) < hs:
                self._interact_mode = INTERACT_RESIZE
                return

        # Inside bbox → move
        if cx1 - 5 <= x <= cx2 + 5 and cy1 - 5 <= y <= cy2 + 5:
            self._interact_mode = INTERACT_MOVE
            return

        self._interact_mode = INTERACT_NONE

    def _on_handle_motion(self, event):
        """Handle dragging for move/resize/rotate."""
        if self._interact_mode == INTERACT_NONE:
            return
        if self._preview_scale <= 0:
            return

        x, y = event.x, event.y
        dx_px = x - self._drag_start_x
        dy_px = y - self._drag_start_y

        if self._interact_mode == INTERACT_MOVE:
            dx_pdf = dx_px / self._preview_scale
            dy_pdf = dy_px / self._preview_scale
            self.xoffset_var.set(f"{self._drag_orig_xoffset + dx_pdf:.0f}")
            self.texty_var.set(f"{self._drag_orig_texty + dy_pdf:.0f}")

        elif self._interact_mode == INTERACT_RESIZE:
            # Drag distance from center determines scale
            bcx, bcy = self._drag_orig_bbox_center
            orig_dist = max(1, math.hypot(self._drag_start_x - bcx, self._drag_start_y - bcy))
            curr_dist = max(1, math.hypot(x - bcx, y - bcy))
            scale_factor = curr_dist / orig_dist
            new_size = max(1.0, min(100.0, self._drag_orig_fontsize * scale_factor))
            self.fontsize_var.set(f"{new_size:.1f}")

        elif self._interact_mode == INTERACT_ROTATE:
            bcx, bcy = self._drag_orig_bbox_center
            curr_angle = math.degrees(math.atan2(x - bcx, -(y - bcy)))
            delta = curr_angle - self._drag_start_angle
            new_rot = self._drag_orig_rotation - delta  # fixed: subtract for correct direction
            # Clamp to -180..180
            while new_rot > 180:
                new_rot -= 360
            while new_rot < -180:
                new_rot += 360
            # Snap to 15° grid if enabled
            if self.snap_rotation_var.get():
                new_rot = round(new_rot / 15) * 15
            self.rotation_var.set(f"{new_rot:.1f}")

        # Throttled re-render for real-time feedback
        self._throttled_render()

    def _on_handle_release(self, event):
        """Finalize interaction — do a clean re-render."""
        if self._interact_mode != INTERACT_NONE:
            self._interact_mode = INTERACT_NONE
            # Cancel any pending throttled render and do a clean one
            if hasattr(self, "_render_after_id"):
                self.after_cancel(self._render_after_id)
                del self._render_after_id
            self.render_preview()

    def _throttled_render(self):
        """Render with a ~40ms throttle for smooth interactive feedback."""
        now = time.time()
        elapsed = now - self._last_render_time

        # Cancel any pending render
        if hasattr(self, "_render_after_id"):
            self.after_cancel(self._render_after_id)
            del self._render_after_id

        if elapsed >= 0.04:  # ~25fps
            self._last_render_time = now
            self.render_preview()
        else:
            # Schedule for remaining time
            delay = int((0.04 - elapsed) * 1000)
            self._render_after_id = self.after(delay, self._do_deferred_render)

    def _do_deferred_render(self):
        if hasattr(self, "_render_after_id"):
            del self._render_after_id
        self._last_render_time = time.time()
        self.render_preview()

    # ─── Font Source Toggle ───────────────────────────────────

    def _on_font_source_change(self):
        self.font_file_frame.pack_forget()
        self.system_font_frame.pack_forget()
        if self.font_source_var.get() == "file":
            self.font_file_frame.pack(fill=tk.X, in_=self.font_selector_container)
        else:
            self.system_font_frame.pack(fill=tk.X, in_=self.font_selector_container)
        self._on_setting_change()

    def _get_active_font_path(self):
        if self.font_source_var.get() == "system":
            family = self.system_font_var.get()
            return self.system_fonts.get(family, "")
        else:
            return self.font_var.get()

    # ─── File Browsing ────────────────────────────────────────

    def _browse_template(self):
        path = filedialog.askopenfilename(title="Select Template PDF", filetypes=[("PDF files", "*.pdf")])
        if path:
            self.template_var.set(path)
            self.render_preview()

    def _browse_excel(self):
        path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.excel_var.set(path)
            self.load_excel()

    def _browse_font(self):
        path = filedialog.askopenfilename(title="Select Font File", filetypes=[("Font files", "*.ttf *.otf")])
        if path:
            self.font_var.set(path)
            self.render_preview()

    def _browse_outdir(self):
        path = filedialog.askdirectory(title="Select Output Directory")
        if path:
            self.outdir_var.set(path)

    # ─── Auto Detect ──────────────────────────────────────────

    def _auto_detect_files(self):
        base = os.path.dirname(os.path.abspath(__file__))

        template = os.path.join(base, "template.pdf")
        if os.path.exists(template):
            self.template_var.set(template)

        excel = os.path.join(base, "ROTALIST.xlsx")
        if os.path.exists(excel):
            self.excel_var.set(excel)

        font = os.path.join(base, "LibreBaskerville.ttf")
        if os.path.exists(font):
            self.font_var.set(font)

        outdir = os.path.join(base, "sertifikalar")
        self.outdir_var.set(outdir)

        if self.excel_var.get():
            self.load_excel()

    # ─── Excel Loading ────────────────────────────────────────

    def load_excel(self):
        path = self.excel_var.get()
        if not path or not os.path.exists(path):
            return
        try:
            self.df = pd.read_excel(path)
            cols = list(self.df.columns)
            self.column_combo["values"] = cols
            if "Column 2" in cols:
                self.column_var.set("Column 2")
            elif len(cols) > 0:
                self.column_var.set(cols[0])
            self.column_var.trace_add("write", lambda *_: self._on_column_change())
            self._refresh_names()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load Excel:\n{e}")

    def _on_column_change(self):
        self._refresh_names()

    def _refresh_names(self):
        if self.df is None or not self.column_var.get():
            return
        col = self.column_var.get()
        if col not in self.df.columns:
            return

        self.names = []
        for _, row in self.df.iterrows():
            raw = str(row[col])
            cleaned = raw.replace('i', 'İ').replace('ç', 'Ç').replace('ö', 'Ö').replace('ü', 'Ü').replace('ş', 'Ş')
            name = cleaned.upper().strip()
            self.names.append(name)

        self.current_index = 0
        self._update_nav_label()
        self.render_preview()

    # ─── Color ────────────────────────────────────────────────

    def _color_to_hex(self):
        r, g, b = self.text_color
        return f"#{int(r*255):02x}{int(g*255):02x}{int(b*255):02x}"

    def _update_swatch(self):
        self.color_swatch.delete("all")
        self.color_swatch.create_rectangle(2, 2, 26, 26, fill=self._color_to_hex(), outline="")

    def pick_color(self):
        initial = self._color_to_hex()
        result = colorchooser.askcolor(color=initial, title="Choose Text Color")
        if result and result[0]:
            r, g, b = result[0]
            self.text_color = (r / 255.0, g / 255.0, b / 255.0)
            self._update_swatch()
            self.color_hex_label.config(text=self._color_to_hex())
            self.render_preview()

    # ─── Navigation ───────────────────────────────────────────

    def navigate(self, direction):
        if not self.names:
            return
        focused = self.focus_get()
        if isinstance(focused, (ttk.Entry, tk.Entry, ttk.Spinbox, ttk.Combobox)):
            return
        self.current_index = (self.current_index + direction) % len(self.names)
        self._update_nav_label()
        self.render_preview()

    def _update_nav_label(self):
        if self.names:
            self.nav_label.config(text=f"{self.names[self.current_index]}    ({self.current_index + 1} / {len(self.names)})")
        else:
            self.nav_label.config(text="No data loaded")

    # ─── Setting Change ──────────────────────────────────────

    def _on_snap_rotation_toggle(self):
        """When snap is toggled on, snap the current rotation to nearest 15°."""
        if self.snap_rotation_var.get():
            try:
                rot = float(self.rotation_var.get())
                snapped = round(rot / 15) * 15
                self.rotation_var.set(f"{snapped:.1f}")
            except ValueError:
                pass
        self._on_setting_change()

    def _on_setting_change(self):
        if hasattr(self, "_setting_after_id"):
            self.after_cancel(self._setting_after_id)
        self._setting_after_id = self.after(300, self.render_preview)

    # ─── Text Splitting Logic ─────────────────────────────────

    def _split_name(self, name):
        mode = self.split_mode_var.get()
        words = name.split()

        if mode == "none":
            return [name]

        if mode == "always":
            if len(words) >= 3:
                return [" ".join(words[:-1]), words[-1]]
            elif len(words) == 2:
                return [words[0], words[1]]
            else:
                return [name]

        try:
            threshold = int(self.split_threshold_var.get())
        except ValueError:
            threshold = 19

        if len(name) >= threshold and len(words) >= 2:
            if len(words) >= 3:
                return [" ".join(words[:-1]), words[-1]]
            else:
                return [words[0], words[1]]
        else:
            return [name]

    # ─── Certificate Rendering (core logic) ───────────────────

    def _get_settings(self):
        try:
            font_size = float(self.fontsize_var.get())
        except ValueError:
            font_size = 35.2
        try:
            text_y = float(self.texty_var.get())
        except ValueError:
            text_y = 307
        try:
            x_offset = float(self.xoffset_var.get())
        except ValueError:
            x_offset = -100
        try:
            line_spacing = float(self.linespace_var.get())
        except ValueError:
            line_spacing = 45
        try:
            rotation = float(self.rotation_var.get())
        except ValueError:
            rotation = 0
        return font_size, text_y, x_offset, line_spacing, rotation

    def render_certificate(self, name):
        template_path = self.template_var.get()
        font_path = self._get_active_font_path()

        if not template_path or not os.path.exists(template_path):
            return None
        if not font_path or not os.path.exists(font_path):
            return None

        font_size, text_y, x_offset, line_spacing, rotation = self._get_settings()
        custom_font = fitz.Font(fontfile=font_path)

        doc = fitz.open(template_path)
        page = doc[0]
        page_width = page.rect.width

        lines = self._split_name(name)
        page.insert_font(fontname="f1", fontfile=font_path)

        alignment = self.alignment_var.get()

        for i, line in enumerate(lines):
            text_width = custom_font.text_length(line, fontsize=font_size)

            if alignment == "center":
                pos_x = ((page_width - text_width) / 2) + x_offset
            elif alignment == "left":
                pos_x = x_offset if x_offset >= 0 else 50 + x_offset
            else:
                pos_x = page_width - text_width + x_offset

            if len(lines) > 1:
                current_y = text_y - (line_spacing / 2) + (i * line_spacing)
            else:
                current_y = text_y

            if abs(rotation) > 0.1:
                pivot = fitz.Point(pos_x + text_width / 2, current_y - font_size / 3)
                morph = (pivot, fitz.Matrix(1, 0, 0, 1, 0, 0).prerotate(rotation))
                page.insert_text(
                    (pos_x, current_y), line,
                    fontname="f1", fontsize=font_size,
                    color=self.text_color, morph=morph
                )
            else:
                page.insert_text(
                    (pos_x, current_y), line,
                    fontname="f1", fontsize=font_size,
                    color=self.text_color
                )

        return doc

    # ─── Preview Rendering ────────────────────────────────────

    def render_preview(self):
        self.canvas.delete("all")

        template_path = self.template_var.get()
        if not template_path or not os.path.exists(template_path):
            self.canvas.create_text(
                self.canvas.winfo_width() // 2,
                self.canvas.winfo_height() // 2,
                text="Select a template PDF to preview",
                fill="#6c7086", font=("Inter", 12)
            )
            return

        if self.names:
            name = self.names[self.current_index]
            doc = self.render_certificate(name)
        else:
            doc = fitz.open(template_path)

        if doc is None:
            self.canvas.create_text(
                self.canvas.winfo_width() // 2,
                self.canvas.winfo_height() // 2,
                text="Cannot render preview\n(check template and font paths)",
                fill="#f38ba8", font=("Inter", 11)
            )
            return

        page = doc[0]
        canvas_w = self.canvas.winfo_width()
        canvas_h = self.canvas.winfo_height()

        if canvas_w < 10 or canvas_h < 10:
            doc.close()
            return

        page_rect = page.rect
        scale_x = canvas_w / page_rect.width
        scale_y = canvas_h / page_rect.height
        scale = min(scale_x, scale_y) * 0.95

        self._preview_scale = scale
        self._preview_page_width = page_rect.width
        self._preview_page_height = page_rect.height

        mat = fitz.Matrix(scale, scale)
        pix = page.get_pixmap(matrix=mat)
        doc.close()

        img = Image.open(io.BytesIO(pix.tobytes("png")))
        self.preview_image = ImageTk.PhotoImage(img)

        img_x = canvas_w // 2
        img_y = canvas_h // 2
        self._preview_img_x = img_x
        self._preview_img_y = img_y
        self.canvas.create_image(img_x, img_y, image=self.preview_image, anchor=tk.CENTER)

        # Draw selection handles over the text
        self._draw_handles()

    # ─── Batch Generation ─────────────────────────────────────

    def generate_all(self):
        if self.generating:
            return
        if not self.names:
            messagebox.showwarning("Warning", "No names loaded. Please load an Excel file first.")
            return

        outdir = self.outdir_var.get()
        if not outdir:
            messagebox.showwarning("Warning", "Please set an output directory.")
            return

        if not os.path.exists(outdir):
            os.makedirs(outdir)

        self.generating = True
        self.gen_btn.config(state="disabled")
        self.progress_var.set(0)
        self.progress_bar["maximum"] = len(self.names)
        self._generate_step(0)

    def _generate_step(self, index):
        if index >= len(self.names):
            self.generating = False
            self.gen_btn.config(state="normal")
            self.status_label.config(text=f"Done! {len(self.names)} certificates generated.")
            messagebox.showinfo("Complete", f"All {len(self.names)} certificates have been generated!")
            return

        name = self.names[index]
        self.status_label.config(text=f"Generating: {name}  ({index + 1}/{len(self.names)})")
        self.progress_var.set(index + 1)

        try:
            doc = self.render_certificate(name)
            if doc:
                clean_name = "".join(c for c in name if c.isalnum() or c in (' ', '_')).rstrip()
                output_path = os.path.join(self.outdir_var.get(), f"{clean_name}.pdf")

                if os.path.exists(output_path):
                    os.remove(output_path)
                doc.save(output_path)
                doc.close()
                print(f"Başarılı: {output_path}")
            else:
                print(f"!!! HATA: {name} — render failed")
        except Exception as e:
            print(f"!!! HATA: {name} -> {e}")

        self.after(10, self._generate_step, index + 1)


    # ─── Save / Import Settings ───────────────────────────────

    def _resolve_path(self, path):
        """If path doesn't exist, try looking in the app's directory."""
        if os.path.exists(path):
            return path
        base = os.path.dirname(os.path.abspath(__file__))
        local = os.path.join(base, os.path.basename(path))
        if os.path.exists(local):
            return local
        return path  # return original even if not found

    def _save_settings(self):
        path = filedialog.asksaveasfilename(
            title="Save Settings",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")],
            initialfile="cert_settings.json"
        )
        if not path:
            return

        settings = {
            "font_source": self.font_source_var.get(),
            "font_file": self.font_var.get(),
            "system_font": self.system_font_var.get(),
            "font_size": self.fontsize_var.get(),
            "text_y": self.texty_var.get(),
            "x_offset": self.xoffset_var.get(),
            "line_spacing": self.linespace_var.get(),
            "rotation": self.rotation_var.get(),
            "snap_rotation": self.snap_rotation_var.get(),
            "color_r": self.text_color[0],
            "color_g": self.text_color[1],
            "color_b": self.text_color[2],
            "alignment": self.alignment_var.get(),
            "split_mode": self.split_mode_var.get(),
            "split_threshold": self.split_threshold_var.get(),
            "template_path": self.template_var.get(),
            "excel_path": self.excel_var.get(),
            "output_dir": self.outdir_var.get(),
            "name_column": self.column_var.get(),
        }

        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(settings, f, indent=4, ensure_ascii=False)
            self.status_label.config(text=f"Settings saved to {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save settings:\n{e}")

    def _import_settings(self):
        path = filedialog.askopenfilename(
            title="Import Settings",
            filetypes=[("JSON files", "*.json")]
        )
        if not path:
            return

        try:
            with open(path, "r", encoding="utf-8") as f:
                s = json.load(f)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load settings:\n{e}")
            return

        # Apply settings with path resolution
        if "template_path" in s:
            self.template_var.set(self._resolve_path(s["template_path"]))
        if "excel_path" in s:
            self.excel_var.set(self._resolve_path(s["excel_path"]))
        if "output_dir" in s:
            self.outdir_var.set(s["output_dir"])
        if "font_source" in s:
            self.font_source_var.set(s["font_source"])
            self._on_font_source_change()
        if "font_file" in s:
            self.font_var.set(self._resolve_path(s["font_file"]))
        if "system_font" in s:
            self.system_font_var.set(s["system_font"])
        if "font_size" in s:
            self.fontsize_var.set(s["font_size"])
        if "text_y" in s:
            self.texty_var.set(s["text_y"])
        if "x_offset" in s:
            self.xoffset_var.set(s["x_offset"])
        if "line_spacing" in s:
            self.linespace_var.set(s["line_spacing"])
        if "rotation" in s:
            self.rotation_var.set(s["rotation"])
        if "snap_rotation" in s:
            self.snap_rotation_var.set(s["snap_rotation"])
        if "color_r" in s and "color_g" in s and "color_b" in s:
            self.text_color = (s["color_r"], s["color_g"], s["color_b"])
            self._update_swatch()
            self.color_hex_label.config(text=self._color_to_hex())
        if "alignment" in s:
            self.alignment_var.set(s["alignment"])
        if "split_mode" in s:
            self.split_mode_var.set(s["split_mode"])
        if "split_threshold" in s:
            self.split_threshold_var.set(s["split_threshold"])
        if "name_column" in s:
            self.column_var.set(s["name_column"])

        # Reload excel if path was set
        if self.excel_var.get():
            self.load_excel()

        self.status_label.config(text=f"Settings imported from {os.path.basename(path)}")
        self.render_preview()

    # ─── Tutorial Overlay ─────────────────────────────────────

    _tutorial_steps = [
        {
            "title": "Welcome! 👋",
            "text": "This is the Certificate Generator.\n\n"
                    "It reads names from an Excel file and\n"
                    "renders them onto a PDF template to\n"
                    "create personalized certificates."
        },
        {
            "title": "1. Load Files 📁",
            "text": "Start by loading your files in the\n"
                    "FILES section on the left panel:\n\n"
                    "• Template PDF — your certificate design\n"
                    "• Excel File — list of names\n"
                    "• Output Dir — where certificates are saved\n\n"
                    "Files in the app folder are auto-detected."
        },
        {
            "title": "2. Choose Font 🔤",
            "text": "Under FONT, pick your text font:\n\n"
                    "• Font File — browse for a .ttf / .otf file\n"
                    "• System Font — pick from installed fonts\n\n"
                    "Toggle between the two with the\n"
                    "radio buttons."
        },
        {
            "title": "3. Text Options ✏️",
            "text": "Adjust how the name appears:\n\n"
                    "• Font Size, Text Y, X Offset\n"
                    "• Line Spacing (for multi-line names)\n"
                    "• Rotation (with optional 15° snap)\n"
                    "• Color picker\n\n"
                    "All changes update the preview live."
        },
        {
            "title": "4. Interactive Preview 🖱️",
            "text": "Click on the text in the preview to\n"
                    "activate selection handles:\n\n"
                    "• Drag the text to reposition\n"
                    "• Drag corner handles to resize\n"
                    "• Drag the top circle to rotate\n\n"
                    "Click outside the box to deactivate.\n"
                    "Use ← → arrow keys to browse names."
        },
        {
            "title": "5. Alignment & Splitting 📐",
            "text": "• ALIGNMENT — Left / Center / Right\n\n"
                    "• TEXT SPLITTING modes:\n"
                    "  Auto — splits long names (≥ threshold)\n"
                    "  No Split — always single line\n"
                    "  Always — always first + last name\n"
        },
        {
            "title": "6. Generate & Export 🚀",
            "text": "Once you're happy with the preview:\n\n"
                    "• Click \"Generate All\" to batch-create\n"
                    "  all certificates as PDFs\n\n"
                    "• Use Save / Import to store your\n"
                    "  settings as a JSON file for reuse."
        },
        {
            "title": "You're all set! ✅",
            "text": "That's everything you need to know.\n\n"
                    "Click this \"?\" button anytime\n"
                    "to see this tutorial again.\n\n"
                    "Happy certificate making! 🎓"
        },
    ]

    _tutorial_index = 0

    def _start_tutorial(self):
        self._tutorial_index = 0
        self._show_tutorial_step()

    def _show_tutorial_step(self):
        # Remove previous overlay
        self.canvas.delete("tutorial")

        if self._tutorial_index >= len(self._tutorial_steps):
            return

        step = self._tutorial_steps[self._tutorial_index]
        cw = self.canvas.winfo_width()
        ch = self.canvas.winfo_height()

        # Semi-transparent overlay (dark rectangle)
        self.canvas.create_rectangle(0, 0, cw, ch,
            fill="#11111b", stipple="gray50", outline="", tags="tutorial")

        # Card background
        card_w = 360
        card_h = 280
        cx = cw // 2
        cy = ch // 2
        x1 = cx - card_w // 2
        y1 = cy - card_h // 2
        x2 = cx + card_w // 2
        y2 = cy + card_h // 2

        # Rounded card (approximate with rectangle)
        self.canvas.create_rectangle(x1, y1, x2, y2,
            fill="#1e1e2e", outline="#89b4fa", width=2, tags="tutorial")

        # Title
        self.canvas.create_text(cx, y1 + 30,
            text=step["title"], fill="#89b4fa",
            font=("Inter", 15, "bold"), tags="tutorial")

        # Body text
        self.canvas.create_text(cx, cy + 5,
            text=step["text"], fill="#cdd6f4",
            font=("Inter", 10), justify=tk.LEFT, width=card_w - 40, tags="tutorial")

        # Footer: step counter + click prompt
        total = len(self._tutorial_steps)
        idx = self._tutorial_index + 1
        if idx < total:
            footer = f"Step {idx}/{total}  —  Click to continue"
        else:
            footer = f"Step {idx}/{total}  —  Click to close"
        self.canvas.create_text(cx, y2 - 20,
            text=footer, fill="#6c7086",
            font=("Inter", 9), tags="tutorial")

        # Bind click on overlay to advance
        self.canvas.tag_bind("tutorial", "<Button-1>", self._tutorial_next)

    def _tutorial_next(self, event):
        self._tutorial_index += 1
        if self._tutorial_index >= len(self._tutorial_steps):
            self.canvas.delete("tutorial")
        else:
            self._show_tutorial_step()


if __name__ == "__main__":
    app = CertificateApp()
    app.mainloop()
