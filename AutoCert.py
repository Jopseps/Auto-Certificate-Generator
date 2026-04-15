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
import locale

TRANSLATIONS = {
    "en": {
        "title": "Certificate Generator", "settings": "⚙  Settings", "files": "FILES", "tpl_pdf": "Template PDF", "excel": "Excel File", "out_dir": "Output Dir", "font": "FONT", "font_file": "Font File", "sys_font": "System Font", "file": "File", "family": "Family", "text_ops": "TEXT OPTIONS", "font_size": "Font Size", "text_y": "Text Y", "x_off": "X Offset", "line_sp": "Line Space", "rot": "Rotation °", "snap": "Snap to 15°", "color": "Color", "align": "ALIGNMENT", "a_left": "Left", "a_center": "Center", "a_right": "Right", "split": "TEXT SPLITTING", "s_auto": "Auto", "s_none": "No Split", "s_always": "Always", "s_thresh": "Threshold", "data": "DATA", "name_col": "Name Col", "gen_all": "▶  Generate All", "ready": "Ready", "settings_hdr": "EXTRAS", "save": "💾 Save", "import": "📂 Import", "preview_hdr": "📄  Certificate Preview", "prev": "◀  Prev", "next": "Next  ▶", "no_data": "No data loaded", "sel_tpl": "Select a template PDF to preview", "fail_render": "Cannot render preview\n(check template and font paths)", "done_gen": "Done! {count} certificates generated.", "gen_msg": "Generating: {name}  ({idx}/{total})", "sav_msg": "Settings saved to {path}", "imp_msg": "Settings imported from {path}",
        "tut1_t": "Welcome! 👋", "tut1_b": "This is the Certificate Generator.\n\nIt reads names from an Excel file and\nrenders them onto a PDF template to\ncreate personalized certificates.", "tut2_t": "1. Load Files 📁", "tut2_b": "Start by loading your files in the\nFILES section on the left panel:\n\n• Template PDF — your certificate design\n• Excel File — list of names\n• Output Dir — where certificates are saved\n\nFiles in the app folder are auto-detected.", "tut3_t": "2. Choose Font 🔤", "tut3_b": "Under FONT, pick your text font:\n\n• Font File — browse for a .ttf / .otf file\n• System Font — pick from installed fonts\n\nToggle between the two with the\nradio buttons.", "tut4_t": "3. Text Options ✏️", "tut4_b": "Adjust how the name appears:\n\n• Font Size, Text Y, X Offset\n• Line Spacing (for multi-line names)\n• Rotation (with optional 15° snap)\n• Color picker\n\nAll changes update the preview live.", "tut5_t": "4. Interactive Preview 🖱️", "tut5_b": "Click on the text in the preview to\nactivate selection handles:\n\n• Drag the text to position, or rotate it.\n• Arrow Keys nudge text 1px (5px with Shift)\n• Press Ctrl+Z to Undo and Ctrl+Shift+Z to Redo\n\nClick outside the box to deactivate.\nUse ← → arrow keys to browse names.", "tut6_t": "5. Alignment & Splitting 📐", "tut6_b": "• ALIGNMENT — Left / Center / Right\n\n• TEXT SPLITTING modes:\n  Auto — splits long names (≥ threshold)\n  No Split — always single line\n  Always — always first + last name\n", "tut7_t": "6. Generate & Export 🚀", "tut7_b": "Once you're happy with the preview:\n\n• Click \"Generate All\" to batch-create\n  all certificates as PDFs\n\n• Use Save / Import to store your\n  settings as a JSON file for reuse.", "tut8_t": "You're all set! ✅", "tut8_b": "That's everything you need to know.\n\nClick this \"?\" button anytime\nto see this tutorial again.\n\nHappy certificate making! 🎓",
        "tut_step": "Step", "tut_cont": "Click to continue", "tut_close": "Click to close"
    },
    "tr": {
        "title": "Sertifika Oluşturucu", "settings": "⚙  Ayarlar", "files": "DOSYALAR", "tpl_pdf": "Şablon PDF", "excel": "Excel Dosyası", "out_dir": "Çıktı Klasörü", "font": "YAZI TİPİ", "font_file": "Yazı Tipi Dosyası", "sys_font": "Sistem Yazı Tipi", "file": "Dosya", "family": "Aile", "text_ops": "METİN AYARLARI", "font_size": "Yazı Boyutu", "text_y": "Metin Y", "x_off": "X Ofseti", "line_sp": "Satır Boşluğu", "rot": "Döndürme °", "snap": "15° Yasla", "color": "Renk", "align": "HİZALAMA", "a_left": "Sol", "a_center": "Orta", "a_right": "Sağ", "split": "METİN BÖLME", "s_auto": "Otomatik", "s_none": "Bölme Yok", "s_always": "Her Zaman", "s_thresh": "Eşik", "data": "VERİ", "name_col": "İsim Sütunu", "gen_all": "▶  Tümünü Oluştur", "ready": "Hazır", "settings_hdr": "EKSTRALAR", "save": "💾 Kaydet", "import": "📂 İçe Aktar", "preview_hdr": "📄  Sertifika Önizlemesi", "prev": "◀  Önceki", "next": "Sonraki  ▶", "no_data": "Veri yüklenmedi", "sel_tpl": "Önizleme için bir şablon PDF seçin", "fail_render": "Önizleme oluşturulamadı\n(şablon ve font yollarını kontrol edin)", "done_gen": "Bitti! {count} sertifika oluşturuldu.", "gen_msg": "Oluşturuluyor: {name}  ({idx}/{total})", "sav_msg": "Ayarlar {path} konumuna kaydedildi", "imp_msg": "Ayarlar {path} konumundan içe aktarıldı",
        "tut1_t": "Hoş Geldiniz! 👋", "tut1_b": "Sertifika Oluşturucu'ya hoş geldiniz.\n\nBir Excel dosyasından isimleri okur ve\nkişiselleştirilmiş sertifikalar oluşturmak\niçin bunları şablonun üzerine yazar.", "tut2_t": "1. Dosyaları Yükle 📁", "tut2_b": "Sol paneldeki DOSYALAR menüsünden\ndosyalarınızı yükleyerek başlayın:\n\n• Şablon PDF — sertifika tasarımınız\n• Excel Dosyası — isim listesi\n• Çıktı Klasörü — kaydedileceği konum\n\nKlasördeki dosyalar otomatik bulunur.", "tut3_t": "2. Yazı Tipi Seçimi 🔤", "tut3_b": "YAZI TİPİ altından fontunuzu seçin:\n\n• Yazı Tipi Dosyası — bir .ttf/.otf dosyası\n• Sistem Yazı Tipi — yüklü fontlardan biri\n\nDüğmeler ile iki format arasında\ngeçiş yapabilirsiniz.", "tut4_t": "3. Metin Ayarları ✏️", "tut4_b": "Yazının duruşunu yapılandırın:\n\n• Yazı Boyutu, Y, X Ofseti\n• Satır Boşluğu (çok satırlı isimler için)\n• Döndürme (isteğe bağlı 15° yaslama)\n• Renk seçici\n\nDeğişiklikler anında önizlemeyi günceller.", "tut5_t": "4. Etkileşimli Önizleme 🖱️", "tut5_b": "Kontrolleri görmek için önizlemedeki\nmetne doğrudan tıklayın:\n\n• Konumlandırmak veya döndürmek için sürükle\n• Yön Tuşları metni 1px kaydırır (Shift ile 5px)\n• Geri almak için Ctrl+Z (Yinelemek için Ctrl+Shift+Z)\n\nKapatmak için dışarıya tıklayın.\nİsimleri ← → ile gezin.", "tut6_t": "5. Hizalama & Bölme 📐", "tut6_b": "• HİZALAMA — Sol / Orta / Sağ\n\n• METİN BÖLME modları:\n  Otomatik — uzun isimleri (≥ eşik) böler\n  Bölme Yok — her zaman tek satırda tutar\n  Her Zaman — her zaman isim + soyisim\n", "tut7_t": "6. Dışa Aktarım 🚀", "tut7_b": "Önizlemeden memnun olduğunuzda:\n\n• Bütün sertifikaları PDF yazdırmak\n  için \"Tümünü Oluştur\"a tıklayın\n\n• Tüm projenizin ayarlarını JSON'a\n  Kaydet / İçe Aktar yapabilirsiniz.", "tut8_t": "Her Şey Hazır! ✅", "tut8_b": "Bilmeniz gerekenler bu kadar.\n\nİstediğiniz bir zaman sağ üstteki \"?\"\nbutonuna tıklayarak tekrar okuyun.\n\nİyi sertifikalamalar! 🎓",
        "tut_step": "Adım", "tut_cont": "Devam etmek için tıklayın", "tut_close": "Kapatmak için tıklayın"
    }
}

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
        # Internal language state & string catalog
        self._lang_widgets = []
        self._lang = self._detect_language()

        self.title(self._("title"))
        self.geometry("1200x750")
        self.minsize(900, 600)

        self.system_fonts = get_system_fonts()

        self.style = ttk.Style(self)
        self.style.theme_use("clam")
        self._configure_styles()

        self.build_ui()

        # Undo / Redo system
        self._undo_stack = []
        self._redo_stack = []
        self._is_undoing = False
        self._last_stable_snapshot = self._get_snapshot()

        self.bind("<Left>", self._on_left_arrow)
        self.bind("<Right>", self._on_right_arrow)
        self.bind("<Up>", self._on_up_arrow)
        self.bind("<Down>", self._on_down_arrow)
        self.bind("<Shift-Left>", self._on_left_arrow)
        self.bind("<Shift-Right>", self._on_right_arrow)
        self.bind("<Shift-Up>", self._on_up_arrow)
        self.bind("<Shift-Down>", self._on_down_arrow)
        self.bind("<Control-z>", self.undo)
        self.bind("<Control-Z>", self.redo)
        self.bind("<Control-Shift-Z>", self.redo)
        self.bind("<Control-Shift-z>", self.redo)

        self._auto_detect_files()
        self.bind_all("<Button-1>", self._on_global_click, add="+")
        
    def _on_global_click(self, event):
        if not isinstance(event.widget, (ttk.Entry, ttk.Spinbox, tk.Entry, tk.Spinbox)):
            self.focus_set()
        
    def _(self, key):
        return TRANSLATIONS.get(self._lang, TRANSLATIONS["en"]).get(key, key)

    def _add_l(self, widget, key, is_text=True):
        if is_text:
            widget.config(text=self._(key))
        self._lang_widgets.append((widget, key, is_text))
        return widget

    def _detect_language(self):
        try:
            lang = locale.getlocale()[0]
            if not lang:
                import os
                lang = os.environ.get('LANG', '')
            if lang and lang.lower().startswith('tr'):
                return "tr"
        except Exception:
            pass
        return "en"

    def _configure_styles(self):
        if not hasattr(self, "theme"):
            self.theme = "dark"

        if self.theme == "dark":
            self.c_bg = "#1e1e2e"
            self.c_fg = "#cdd6f4"
            self.c_accent = "#89b4fa"
            self.c_entry_bg = "#313244"
            self.c_btn_bg = "#45475a"
            self.c_btn_active = "#585b70"
            self.c_canvas_bg = "#181825"
            self.c_text_inactive = "#6c7086"
            self.c_error = "#f38ba8"
            self.c_overlay = "#11111b"
            self.c_success = "#a6e3a1"
            self.c_handle_fill = "#89b4fa"
            self.c_handle_outline = "#cdd6f4"
        else:
            self.c_bg = "#eff1f5"
            self.c_fg = "#4c4f69"
            self.c_accent = "#1e66f5"
            self.c_entry_bg = "#ccd0da"
            self.c_btn_bg = "#bcc0cc"
            self.c_btn_active = "#acb0be"
            self.c_canvas_bg = "#e6e9ef"
            self.c_text_inactive = "#7c7f93"
            self.c_error = "#d20f39"
            self.c_overlay = "#dce0e8"
            self.c_success = "#40a02b"
            self.c_handle_fill = "#93b5fa"
            self.c_handle_outline = "#78a3f9"

        self.style.configure("TFrame", background=self.c_bg)
        self.style.configure("TLabel", background=self.c_bg, foreground=self.c_fg, font=("Inter", 10))
        self.style.configure("Header.TLabel", background=self.c_bg, foreground=self.c_accent, font=("Inter", 13, "bold"))
        self.style.configure("Section.TLabel", background=self.c_bg, foreground=self.c_text_inactive, font=("Inter", 9, "bold"))
        self.style.configure("TButton", background=self.c_btn_bg, foreground=self.c_fg, font=("Inter", 10), borderwidth=0, padding=6)
        self.style.map("TButton", background=[("active", self.c_btn_active)])
        self.style.configure("Accent.TButton", background=self.c_accent, foreground=self.c_bg, font=("Inter", 11, "bold"), padding=10)
        self.style.map("Accent.TButton", background=[("active", self.c_handle_outline)])
        self.style.configure("Nav.TButton", background=self.c_btn_bg, foreground=self.c_fg, font=("Inter", 10), padding=4)
        self.style.configure("TEntry", fieldbackground=self.c_entry_bg, foreground=self.c_fg, insertcolor=self.c_fg, borderwidth=0)
        self.style.configure("TSpinbox", fieldbackground=self.c_entry_bg, foreground=self.c_fg, insertcolor=self.c_fg, borderwidth=0, arrowcolor=self.c_fg)
        self.style.configure("TCombobox", fieldbackground=self.c_entry_bg, foreground=self.c_fg, borderwidth=0)
        self.style.map("TCombobox", fieldbackground=[("readonly", self.c_entry_bg)])
        self.style.configure("green.Horizontal.TProgressbar", troughcolor=self.c_entry_bg, background=self.c_success, borderwidth=0)
        self.style.configure("TRadiobutton", background=self.c_bg, foreground=self.c_fg, font=("Inter", 9))
        self.style.map("TRadiobutton", background=[("active", self.c_bg)])
        self.style.configure("TCheckbutton", background=self.c_bg, foreground=self.c_fg, font=("Inter", 9))
        self.style.map("TCheckbutton", background=[("active", self.c_bg)])
        self.configure(bg=self.c_bg)

    def build_ui(self):
        # Top bar with buttons
        topbar = ttk.Frame(self)
        topbar.pack(fill=tk.X, padx=12, pady=(8, 0))
        ttk.Label(topbar, text="", style="TLabel").pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.help_btn = tk.Button(topbar, text="?", font=("Inter", 12, "bold"),
            bg=self.c_btn_bg, fg=self.c_fg, activebackground=self.c_btn_active, activeforeground=self.c_fg,
            bd=0, width=3, height=1, cursor="hand2", command=self._start_tutorial)
        self.help_btn.pack(side=tk.RIGHT)

        self.theme_btn = tk.Button(topbar, text="☀️", font=("Inter", 12),
            bg=self.c_btn_bg, fg=self.c_fg, activebackground=self.c_btn_active, activeforeground=self.c_fg,
            bd=0, width=3, height=1, cursor="hand2", command=self.toggle_theme)
        self.theme_btn.pack(side=tk.RIGHT, padx=(0, 6))
    
        self.lang_btn = tk.Button(topbar, text="EN", font=("Inter", 10, "bold"),
            bg=self.c_btn_bg, fg=self.c_fg, activebackground=self.c_btn_active, activeforeground=self.c_fg,
            bd=0, width=3, height=1, cursor="hand2", command=self.toggle_lang)
        self.lang_btn.pack(side=tk.RIGHT, padx=(0, 6))
        self._update_lang_btn_text()

        self.redo_btn = tk.Button(topbar, text="⟳", font=("Inter", 14, "bold"),
            bg=self.c_btn_bg, fg=self.c_fg, activebackground=self.c_btn_active, activeforeground=self.c_fg,
            bd=0, width=3, height=1, cursor="hand2", command=self.redo)
        self.redo_btn.pack(side=tk.RIGHT, padx=(0, 21))

        self.undo_btn = tk.Button(topbar, text="⟲", font=("Inter", 14, "bold"),
            bg=self.c_btn_bg, fg=self.c_fg, activebackground=self.c_btn_active, activeforeground=self.c_fg,
            bd=0, width=3, height=1, cursor="hand2", command=self.undo)
        self.undo_btn.pack(side=tk.RIGHT, padx=(0, 6))

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
        self._settings_canvas = tk.Canvas(parent, bg=self.c_bg, highlightthickness=0)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=self._settings_canvas.yview)
        self.settings_inner = ttk.Frame(self._settings_canvas)

        self.settings_inner.bind(
            "<Configure>",
            lambda e: self._settings_canvas.configure(scrollregion=self._settings_canvas.bbox("all"))
        )
        self._settings_canvas.bind(
            "<Configure>",
            lambda e: self._settings_canvas.itemconfig(self._settings_window_id, width=e.width)
        )
        self._settings_window_id = self._settings_canvas.create_window((0, 0), window=self.settings_inner, anchor="nw", width=336)
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

        self._add_l(ttk.Label(inner, style="Header.TLabel"), "settings").pack(anchor="w", pady=(0, 4))

        # --- File selectors ---
        self._add_l(ttk.Label(inner, style="Section.TLabel"), "files").pack(anchor="w", pady=(4, 2))

        self.template_var = tk.StringVar()
        self._file_row(inner, "tpl_pdf", self.template_var, self._browse_template)

        self.excel_var = tk.StringVar()
        self._file_row(inner, "excel", self.excel_var, self._browse_excel)

        self.outdir_var = tk.StringVar(value="sertifikalar")
        self._file_row(inner, "out_dir", self.outdir_var, self._browse_outdir)

        ttk.Frame(inner, height=2).pack(fill=tk.X, pady=6)

        # --- Font options ---
        self._add_l(ttk.Label(inner, style="Section.TLabel"), "font").pack(anchor="w", pady=(0, 2))

        self.font_source_var = tk.StringVar(value="system")
        font_src_frame = ttk.Frame(inner)
        font_src_frame.pack(fill=tk.X, pady=1)
        self._add_l(ttk.Radiobutton(font_src_frame, variable=self.font_source_var, value="file", command=self._on_font_source_change), "font_file").pack(side=tk.LEFT, padx=(0, 12))
        self._add_l(ttk.Radiobutton(font_src_frame, variable=self.font_source_var, value="system", command=self._on_font_source_change), "sys_font").pack(side=tk.LEFT)

        self.font_selector_container = ttk.Frame(inner)
        self.font_selector_container.pack(fill=tk.X, pady=1)

        self.font_var = tk.StringVar()
        self.font_file_frame = ttk.Frame(self.font_selector_container)
        self._add_l(ttk.Label(self.font_file_frame, width=11), "file").pack(side=tk.LEFT)
        ttk.Entry(self.font_file_frame, textvariable=self.font_var, width=14).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        ttk.Button(self.font_file_frame, text="...", width=3, command=self._browse_font).pack(side=tk.LEFT)

        self.system_font_var = tk.StringVar()
        self.system_font_frame = ttk.Frame(self.font_selector_container)
        self._add_l(ttk.Label(self.system_font_frame, width=11), "family").pack(side=tk.LEFT)
        font_families = list(self.system_fonts.keys())
        # Set default system font to Times New Roman or a basic fallback
        if "Times New Roman" in font_families:
            self.system_font_var.set("Times New Roman")
        else:
            basic_fonts = ["Liberation Serif", "DejaVu Serif", "Noto Serif", "Arial", "Liberation Sans"]
            for font in basic_fonts:
                if font in font_families:
                    self.system_font_var.set(font)
                    break
            else:
                if font_families:
                    self.system_font_var.set(font_families[0])
        self.system_font_combo = ttk.Combobox(self.system_font_frame, textvariable=self.system_font_var, values=font_families, state="readonly", width=18)
        self.system_font_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.system_font_var.trace_add("write", lambda *_: self._on_setting_change())
        
        # Initialize with system font selected
        self._on_font_source_change()

        ttk.Frame(inner, height=2).pack(fill=tk.X, pady=6)

        # --- Text options ---
        self._add_l(ttk.Label(inner, style="Section.TLabel"), "text_ops").pack(anchor="w", pady=(0, 2))

        self.fontsize_var = tk.StringVar(value="35.2")
        self._spin_row(inner, "font_size", self.fontsize_var, 1.0, 100.0, 0.5)

        self.texty_var = tk.StringVar(value="307")
        self._spin_row(inner, "text_y", self.texty_var, 0, 2000, 1)

        self.xoffset_var = tk.StringVar(value="-100")
        self._spin_row(inner, "x_off", self.xoffset_var, -1000, 1000, 5)

        self.linespace_var = tk.StringVar(value="45")
        self._spin_row(inner, "line_sp", self.linespace_var, 0, 200, 1)

        self.rotation_var = tk.StringVar(value="0.0")
        self._spin_row(inner, "rot", self.rotation_var, -180, 180, 1)

        self.snap_rotation_var = tk.BooleanVar(value=False)
        snap_frame = ttk.Frame(inner)
        snap_frame.pack(fill=tk.X, pady=(0, 1))
        self._add_l(ttk.Checkbutton(snap_frame, variable=self.snap_rotation_var, command=self._on_snap_rotation_toggle), "snap").pack(side=tk.LEFT, padx=(88, 0))

        color_frame = ttk.Frame(inner)
        color_frame.pack(fill=tk.X, pady=1)
        self._add_l(ttk.Label(color_frame, width=11), "color").pack(side=tk.LEFT)
        self.color_swatch = tk.Canvas(color_frame, width=28, height=28, highlightthickness=1, highlightbackground=self.c_btn_active, cursor="hand2")
        self.color_swatch.pack(side=tk.LEFT, padx=(0, 6))
        self._update_swatch()
        self.color_swatch.bind("<Button-1>", lambda e: self.pick_color())
        self.color_hex_label = ttk.Label(color_frame, text=self._color_to_hex(), font=("Inter", 9))
        self.color_hex_label.pack(side=tk.LEFT)

        ttk.Frame(inner, height=2).pack(fill=tk.X, pady=6)

        # --- Text Alignment ---
        self._add_l(ttk.Label(inner, style="Section.TLabel"), "align").pack(anchor="w", pady=(0, 2))

        self.alignment_var = tk.StringVar(value="center")
        align_frame = ttk.Frame(inner)
        align_frame.pack(fill=tk.X, pady=1)
        for val, key in [("left", "a_left"), ("center", "a_center"), ("right", "a_right")]:
            self._add_l(ttk.Radiobutton(align_frame, variable=self.alignment_var, value=val, command=self._on_setting_change), key).pack(side=tk.LEFT, padx=(0, 10))

        ttk.Frame(inner, height=2).pack(fill=tk.X, pady=6)

        # --- Text Splitting ---
        self._add_l(ttk.Label(inner, style="Section.TLabel"), "split").pack(anchor="w", pady=(0, 2))

        self.split_mode_var = tk.StringVar(value="auto")
        split_frame = ttk.Frame(inner)
        split_frame.pack(fill=tk.X, pady=1)
        self._add_l(ttk.Radiobutton(split_frame, variable=self.split_mode_var, value="auto", command=self._on_setting_change), "s_auto").pack(side=tk.LEFT, padx=(0, 8))
        self._add_l(ttk.Radiobutton(split_frame, variable=self.split_mode_var, value="none", command=self._on_setting_change), "s_none").pack(side=tk.LEFT, padx=(0, 8))
        self._add_l(ttk.Radiobutton(split_frame, variable=self.split_mode_var, value="always", command=self._on_setting_change), "s_always").pack(side=tk.LEFT)

        self.split_threshold_var = tk.StringVar(value="19")
        self._spin_row(inner, "s_thresh", self.split_threshold_var, 1, 100, 1)

        ttk.Frame(inner, height=2).pack(fill=tk.X, pady=6)

        # --- Excel column ---
        self._add_l(ttk.Label(inner, style="Section.TLabel"), "data").pack(anchor="w", pady=(0, 2))

        col_frame = ttk.Frame(inner)
        col_frame.pack(fill=tk.X, pady=1)
        self._add_l(ttk.Label(col_frame, width=11), "name_col").pack(side=tk.LEFT)
        self.column_var = tk.StringVar()
        self.column_combo = ttk.Combobox(col_frame, textvariable=self.column_var, state="readonly", width=18)
        self.column_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)

        ttk.Frame(inner, height=2).pack(fill=tk.X, pady=6)

        # --- Generate ---
        self.gen_btn = self._add_l(ttk.Button(inner, style="Accent.TButton", command=self.generate_all), "gen_all")
        self.gen_btn.pack(fill=tk.X, pady=(2, 6))

        self.progress_var = tk.IntVar(value=0)
        self.progress_bar = ttk.Progressbar(inner, variable=self.progress_var, maximum=100, style="green.Horizontal.TProgressbar")

        self.status_label = self._add_l(ttk.Label(inner, font=("Inter", 9)), "ready")
        self.status_label.pack(anchor="w")

        ttk.Frame(inner, height=2).pack(fill=tk.X, pady=6)

        # --- Save / Import Settings ---
        self._add_l(ttk.Label(inner, style="Section.TLabel"), "settings_hdr").pack(anchor="w", pady=(0, 2))
        settings_btns = ttk.Frame(inner)
        settings_btns.pack(fill=tk.X, pady=1)
        self._add_l(ttk.Button(settings_btns, command=self._save_settings), "save").pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        self._add_l(ttk.Button(settings_btns, command=self._import_settings), "import").pack(side=tk.LEFT, fill=tk.X, expand=True)

    def _file_row(self, parent, key, var, command):
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.X, pady=1)
        self._add_l(ttk.Label(frame, width=11), key).pack(side=tk.LEFT)
        entry = ttk.Entry(frame, textvariable=var, width=14)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        ttk.Button(frame, text="...", width=3, command=command).pack(side=tk.LEFT)

    def _spin_row(self, parent, key, var, from_, to, increment):
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.X, pady=1)
        self._add_l(ttk.Label(frame, width=11), key).pack(side=tk.LEFT)
        spin = ttk.Spinbox(frame, textvariable=var, from_=from_, to=to, increment=increment, width=10)
        spin.pack(side=tk.LEFT, fill=tk.X, expand=True)
        var.trace_add("write", lambda *_: self._on_setting_change())

    # ─── Preview Panel ────────────────────────────────────────

    def _build_preview_panel(self, parent):
        self._add_l(ttk.Label(parent, style="Header.TLabel"), "preview_hdr").pack(anchor="w", pady=(0, 8))

        self.preview_frame = ttk.Frame(parent)
        self.preview_frame.pack(fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(self.preview_frame, bg=self.c_canvas_bg, highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)

        nav = ttk.Frame(parent)
        nav.pack(fill=tk.X, pady=(8, 0))

        self.prev_btn = self._add_l(ttk.Button(nav, style="Nav.TButton", command=lambda: self.navigate(-1)), "prev")
        self.prev_btn.pack(side=tk.LEFT)

        self.nav_label = self._add_l(ttk.Label(nav, font=("Inter", 10)), "no_data")
        self.nav_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.next_btn = self._add_l(ttk.Button(nav, style="Nav.TButton", command=lambda: self.navigate(1)), "next")
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
            outline=self.c_handle_fill, width=1.5, dash=(5, 3), tags="handle")

        # Corner handles (small filled squares)
        hs = 5
        corners = [
            (cx1, cy1), (cx2, cy1),  # top-left, top-right
            (cx1, cy2), (cx2, cy2),  # bottom-left, bottom-right
        ]
        for x, y in corners:
            self.canvas.create_rectangle(x - hs, y - hs, x + hs, y + hs,
                fill=self.c_handle_fill, outline=self.c_handle_outline, width=1, tags=("handle", "h_corner"))

        # Edge midpoint handles (smaller)
        ms = 4
        mid_x = (cx1 + cx2) / 2
        mid_y = (cy1 + cy2) / 2
        for x, y in [(mid_x, cy1), (mid_x, cy2), (cx1, mid_y), (cx2, mid_y)]:
            self.canvas.create_rectangle(x - ms, y - ms, x + ms, y + ms,
                fill=self.c_handle_fill, outline=self.c_handle_outline, width=1, tags=("handle", "h_corner"))

        # Rotation handle (circle above top center, connected by a line)
        rot_line_len = 28
        rot_y = cy1 - rot_line_len
        # Connecting line
        self.canvas.create_line(mid_x, cy1, mid_x, rot_y,
            fill=self.c_handle_fill, width=1.5, tags="handle")
        # Circle
        rs = 7
        self.canvas.create_oval(mid_x - rs, rot_y - rs, mid_x + rs, rot_y + rs,
            fill=self.c_bg, outline=self.c_handle_fill, width=2, tags=("handle", "h_rotate"))

    # ─── Handle Interaction ───────────────────────────────────

    def _on_hover(self, event):
        """Change cursor based on what's under the mouse."""
        if self.canvas.find_withtag("tutorial"):
            self.canvas.config(cursor="arrow")
            return

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
        if self.canvas.find_withtag("tutorial"):
            return

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
        if self.canvas.find_withtag("tutorial"):
            return
            
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
        if self.canvas.find_withtag("tutorial"):
            return
            
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

        if elapsed >= 0.016:  # ~60fps
            self._last_render_time = now
            self.render_preview()
        else:
            # Schedule for remaining time
            delay = int((0.016 - elapsed) * 1000)
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

    # ─── Keyboard Nudging & Undo/Redo ────────────────────────

    def _on_left_arrow(self, event):
        shift = bool(event.state & 0x0001)
        if self._handles_active:
            self._nudge_text(-5 if shift else -1, 0)
        else:
            self.navigate(-1)

    def _on_right_arrow(self, event):
        shift = bool(event.state & 0x0001)
        if self._handles_active:
            self._nudge_text(5 if shift else 1, 0)
        else:
            self.navigate(1)
            
    def _on_up_arrow(self, event):
        shift = bool(event.state & 0x0001)
        if self._handles_active:
            self._nudge_text(0, -5 if shift else -1)
            
    def _on_down_arrow(self, event):
        shift = bool(event.state & 0x0001)
        if self._handles_active:
            self._nudge_text(0, 5 if shift else 1)

    def _nudge_text(self, dx, dy):
        try:
            if dx != 0:
                curr_x = float(self.xoffset_var.get())
                self.xoffset_var.set(f"{curr_x + dx:.0f}")
            if dy != 0:
                curr_y = float(self.texty_var.get())
                self.texty_var.set(f"{curr_y + dy:.0f}")
        except ValueError:
            pass

    def _get_snapshot(self):
        try:
            return {
                "font_size": self.fontsize_var.get(),
                "text_y": self.texty_var.get(),
                "x_offset": self.xoffset_var.get(),
                "line_spacing": self.linespace_var.get(),
                "rotation": self.rotation_var.get()
            }
        except Exception:
            return None

    def _apply_snapshot(self, snap):
        if not snap: return
        self._is_undoing = True
        self.fontsize_var.set(snap["font_size"])
        self.texty_var.set(snap["text_y"])
        self.xoffset_var.set(snap["x_offset"])
        self.linespace_var.set(snap["line_spacing"])
        self.rotation_var.set(snap["rotation"])
        self._is_undoing = False
        self.render_preview()
        self._last_stable_snapshot = self._get_snapshot()

    def _push_undo(self, snapshot):
        if not snapshot: return
        if self._undo_stack and self._undo_stack[-1] == snapshot:
            return
        self._undo_stack.append(snapshot)
        if len(self._undo_stack) > 50:
            self._undo_stack.pop(0)
        self._redo_stack.clear()

    def undo(self, event=None):
        if not self._undo_stack:
            return
        curr_snap = self._get_snapshot()
        prev_snap = self._undo_stack.pop()
        self._redo_stack.append(curr_snap)
        self._apply_snapshot(prev_snap)

    def redo(self, event=None):
        if not self._redo_stack:
            return
        curr_snap = self._get_snapshot()
        next_snap = self._redo_stack.pop()
        self._undo_stack.append(curr_snap)
        self._apply_snapshot(next_snap)

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
        if getattr(self, "_is_undoing", False):
            return
        
        # Skip rendering if canvas doesn't exist yet (during initialization)
        if not hasattr(self, "canvas"):
            return
            
        # Capture pre-change snapshot
        if not hasattr(self, "_setting_start_snapshot"):
            self._setting_start_snapshot = getattr(self, "_last_stable_snapshot", self._get_snapshot())
            
        # Restart the commit timer
        if hasattr(self, "_setting_after_id"):
            self.after_cancel(self._setting_after_id)
        self._setting_after_id = self.after(400, self._finalize_setting_change)
        
        # Immediate visual feedback, throttled
        current = time.time()
        if not hasattr(self, "_last_render_time"): self._last_render_time = 0
        if current - self._last_render_time >= 0.016:
            self.render_preview()
            self._last_render_time = time.time()

    def _finalize_setting_change(self):
        # One last render to ensure we didn't drop the final frame
        self.render_preview()
        curr = self._get_snapshot()
        if hasattr(self, "_setting_start_snapshot"):
            if curr and self._setting_start_snapshot != curr:
                self._push_undo(self._setting_start_snapshot)
            del self._setting_start_snapshot
        self._last_stable_snapshot = curr

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
                text=self._("sel_tpl"),
                fill=self.c_text_inactive, font=("Inter", 12)
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
                text=self._("fail_render"),
                fill=self.c_error, font=("Inter", 11)
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
        self.progress_bar.pack(fill=tk.X, pady=(0, 4), before=self.status_label)
        self._generate_step(0)

    def _generate_step(self, index):
        if index >= len(self.names):
            self.generating = False
            self.gen_btn.config(state="normal")
            #Fself.progress_bar.pack_forget()
            self.status_label.config(text=self._("done_gen").format(count=len(self.names)))
            messagebox.showinfo("Complete", f"All {len(self.names)} certificates have been generated!")
            return

        name = self.names[index]
        self.status_label.config(text=self._("gen_msg").format(name=name, idx=index+1, total=len(self.names)))
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
            self.status_label.config(text=self._("sav_msg").format(path=os.path.basename(path)))
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

        self.status_label.config(text=self._("imp_msg").format(path=os.path.basename(path)))
        self.render_preview()

    # ─── Tutorial Overlay ─────────────────────────────────────

    _tutorial_index = 0

    def _start_tutorial(self):
        self._tutorial_index = 0
        self._show_tutorial_step()

    def _show_tutorial_step(self):
        # Remove previous overlay
        self.canvas.delete("tutorial")

        if self._tutorial_index >= 8:
            return

        title_text = self._(f"tut{self._tutorial_index + 1}_t")
        body_text = self._(f"tut{self._tutorial_index + 1}_b")
        cw = self.canvas.winfo_width()
        ch = self.canvas.winfo_height()

        # Semi-transparent overlay (dark rectangle)
        self.canvas.create_rectangle(0, 0, cw, ch,
            fill=self.c_overlay, stipple="gray50", outline="", tags="tutorial")

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
            fill=self.c_bg, outline=self.c_accent, width=2, tags="tutorial")

        # Title
        self.canvas.create_text(cx, y1 + 30,
            text=title_text, fill=self.c_accent,
            font=("Inter", 15, "bold"), tags="tutorial")

        # Body text
        self.canvas.create_text(cx, cy + 5,
            text=body_text, fill=self.c_fg,
            font=("Inter", 10), justify=tk.LEFT, width=card_w - 40, tags="tutorial")

        # Footer: step counter + click prompt
        total = 8
        idx = self._tutorial_index + 1
        if idx < total:
            footer = f"{self._('tut_step')} {idx}/{total}  —  {self._('tut_cont')}"
        else:
            footer = f"{self._('tut_step')} {idx}/{total}  —  {self._('tut_close')}"
        self.canvas.create_text(cx, y2 - 20,
            text=footer, fill=self.c_text_inactive,
            font=("Inter", 9), tags="tutorial")

        # Bind click on overlay to advance
        self.canvas.tag_bind("tutorial", "<Button-1>", self._tutorial_next)

    def _tutorial_next(self, event):
        self._tutorial_index += 1
        if self._tutorial_index >= 8:
            self.canvas.delete("tutorial")
        else:
            self._show_tutorial_step()

    # ─── Language & Theme Toggle ──────────────────────────────

    def toggle_theme(self):
        self.theme = "light" if self.theme == "dark" else "dark"
        self._configure_styles()
        self._apply_theme()

    def toggle_lang(self):
        self._lang = "tr" if self._lang == "en" else "en"
        self._update_lang_btn_text()
        self._apply_strings()
        
    def _update_lang_btn_text(self):
        self.lang_btn.config(text="EN" if self._lang == "en" else "TR")

    def _apply_strings(self):
        self.title(self._("title"))
        for widget, key, is_text in self._lang_widgets:
            if is_text:
                widget.config(text=self._(key))
        
        self.system_font_combo.set(self.system_font_var.get())
        if self.canvas.find_withtag("tutorial"):
            self._show_tutorial_step()
        self.render_preview()

    def _apply_theme(self):
        self.configure(bg=self.c_bg)
        self._settings_canvas.configure(bg=self.c_bg)
        self.canvas.configure(bg=self.c_canvas_bg)

        for btn in (self.help_btn, self.lang_btn, self.undo_btn, self.redo_btn):
            btn.config(bg=self.c_btn_bg, fg=self.c_fg,
                       activebackground=self.c_btn_active, activeforeground=self.c_fg)

        self.theme_btn.config(text="☀️" if self.theme == "dark" else "🌙",
                              bg=self.c_btn_bg, fg=self.c_fg,
                              activebackground=self.c_btn_active, activeforeground=self.c_fg)

        self.color_swatch.config(highlightbackground=self.c_btn_active)

        if self.canvas.find_withtag("tutorial"):
            self._show_tutorial_step()

        self.render_preview()


if __name__ == "__main__":
    app = CertificateApp()
    app.mainloop()
