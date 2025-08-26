#!/usr/bin/env python3
import json
import locale
import os
import subprocess
import sys
import threading
import threading
import tkinter as tk

from pathlib import Path
from tkinter import ttk
from tkinter import filedialog, messagebox

def get_base_dir():
    """
    Return the base directory for bundled resources:
      - When frozen (bundled), return the bundle’s temp extraction directory.
      - Otherwise, return the directory of this source file.
    Works on macOS, Windows, and Linux.
    """
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent

APP_DIR = get_base_dir()
LOCALES_DIR = APP_DIR / "locales"
DEFAULT_LOCALE = "en"

EXCEL_DIFF_SCRIPT = str((APP_DIR / "excel_diff.py").resolve())

# Appearance toggles
GROUP_BORDERS = True          # True: use ttk.Labelframe borders; False: use plain ttk.Frame
SHOW_GROUP_TITLES = False      # True: show titles; False: hide titles (independent of borders)

# Consistent padding and widths
GROUP_PAD_Y = 8
GROUP_INNER_PAD = dict(padx=0, pady=(0, 8))
LABEL_WIDTH = 28
if os.name == "nt":
    LABEL_WIDTH = 40  # slightly wider for Windows default fonts
PATH_ENTRY_WIDTH = 72      # unified width for file/pattern/extra header fields
SMALL_ENTRY_WIDTH = 10     # unified width for letter and numeric controls
BUTTON_PADX = 10           # base left padding for buttons
BUTTON_PAD = (BUTTON_PADX, BUTTON_PADX + 6)  # a bit more on the right for breathing space
BUTTON_WIDTH = 8          # consistent width across all buttons (in text units)

# ---------- i18n helpers (robust, no deprecated getdefaultlocale) ----------

def _normalize_locale_tag(tag: str) -> str:
    """
    Normalize locale tag to lang[_REGION] with underscores.
    Examples:
      'en-GB.UTF-8' -> 'en_GB'
      'pl_PL@latin' -> 'pl_PL'
      'hr' -> 'hr'
    """
    if not tag:
        return ""
    s = str(tag)
    s = s.split(".", 1)[0]     # drop codeset
    s = s.split("@", 1)[0]     # drop modifier
    s = s.strip().replace("-", "_")
    if not s:
        return ""
    parts = s.split("_")
    if len(parts) == 1:
        return parts[0].lower()
    lang = parts[0].lower()
    region = parts[1].upper()
    return f"{lang}_{region}"

def _gather_env_locales():
    """
    Collect candidate locales from env (LC_ALL, LC_MESSAGES, LANG).
    Returns list like ['en_GB', 'en'] without duplicates, preserving order.
    """
    seen = set()
    out = []
    for var in ("LC_ALL", "LC_MESSAGES", "LANG"):
        val = os.environ.get(var)
        if not val:
            continue
        if isinstance(val, str):
            candidates = val.split(":") if ":" in val else [val]
        else:
            candidates = [str(val)]
        for cand in candidates:
            norm = _normalize_locale_tag(cand)
            if norm and norm not in seen:
                seen.add(norm)
                out.append(norm)
                base = norm.split("_", 1)[0]
                if base and base not in seen:
                    seen.add(base)
                    out.append(base)
    return out

def detect_locale_chain():
    """
    Preference list without using deprecated APIs:
      1) Current LC_CTYPE via setlocale/getlocale (variant + base).
      2) Env locales (variant + base).
      3) DEFAULT_LOCALE.
    """
    prefs = []
    seen = set()

    try:
        locale.setlocale(locale.LC_CTYPE, "")
        lang_tuple = locale.getlocale()
        if lang_tuple and lang_tuple[0]:
            norm = _normalize_locale_tag(lang_tuple[0])
            if norm and norm not in seen:
                seen.add(norm)
                prefs.append(norm)
                base = norm.split("_", 1)[0]
                if base and base not in seen:
                    seen.add(base)
                    prefs.append(base)
    except Exception:
        pass

    for env_loc in _gather_env_locales():
        if env_loc not in seen:
            seen.add(env_loc)
            prefs.append(env_loc)

    if DEFAULT_LOCALE not in seen:
        prefs.append(DEFAULT_LOCALE)

    return prefs

def load_labels():
    """
    Load and merge locale files, from base to variant.
    This ensures that specific locales (e.g., en_GB) override general ones (e.g., en).
    """
    merged = {}
    # Reverse the chain to load from base to variant (e.g., 'en' then 'en_IE')
    for cand in reversed(detect_locale_chain()):
        f = LOCALES_DIR / f"{cand}_gui.json"
        if f.exists():
            try:
                with f.open("r", encoding="utf-8") as fh:
                    data = json.load(fh)
                    merged.update(data)
            except json.JSONDecodeError as e:
                # This will catch malformed JSON and tell you exactly where the error is.
                print(f"Error: Could not parse '{f.name}'. It may contain invalid JSON. Details: {e}", file=sys.stderr)
            except Exception as e:
                # Catch other potential file reading errors.
                print(f"Error: Could not read file '{f.name}'. Details: {e}", file=sys.stderr)
    return merged

# Initialize labels and helper BEFORE class definition
LABELS = load_labels()
def L(key, default_text):
    return LABELS.get(key, default_text)

# ---------- GUI ----------

if os.name == "nt":
    try:
        # Make process DPI-aware on Windows so fonts and geometry scale correctly
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(2)  # 2=Per-Monitor DPI awareness if available; fallback is okay
    except Exception:
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)  # system DPI as fallback
        except Exception:
            pass

class ExcelDiffGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(L("app_title", "Excel Diff"))
        # Start reasonably sized; will be adjusted with geometry("") after layout
        if os.name == "nt":
            self.geometry("980x780")
        else:
            self.geometry("940x760")

        # Window/app icon (from ./icons/icon.png)
        try:
            icon_path = APP_DIR / "icons" / "icon.png"
            if icon_path.exists():
                # Keep a reference so it's not garbage-collected
                self._app_icon = tk.PhotoImage(file=str(icon_path))
                # Apply to this window and future toplevels (True), or only this window (False)
                self.iconphoto(True, self._app_icon)
        except Exception:
            # Icon is optional; ignore if unsupported or missing
            pass

        # Styles for status label: default uses theme color; error uses red
        self.style = ttk.Style(self)
        self.style.configure("StatusDefault.TLabel")  # no explicit foreground -> theme default
        self.style.configure("StatusError.TLabel", foreground="#fc6666")  # red on errors

        self.home = Path.home()
        self.start_dirs = {"original": self.home, "modified": self.home, "output": self.home}

        # State variables
        self.var_compare_folders = tk.BooleanVar(value=False)
        self.var_original = tk.StringVar(value=str(self.home))
        self.var_modified = tk.StringVar(value=str(self.home))
        self.var_wspattern = tk.StringVar(value=".*")
        self.var_row_offset = tk.IntVar(value=0)
        self.var_source_col = tk.StringVar(value="")
        self.var_target_col = tk.StringVar(value="")

        self.var_extract_extra = tk.BooleanVar(value=False)
        self.var_extra_col = tk.StringVar(value="")
        self.var_extra_header = tk.StringVar(value="")

        self.var_nolimits = tk.BooleanVar(value=False)
        self.var_tolerate = tk.IntVar(value=0)   # 0..35 or 0..100
        self.var_realign = tk.IntVar(value=0)    # 0..15 or very large for "no limit"

        self.var_include_identical_pairs = tk.BooleanVar(value=True)
        self.var_output_html = tk.StringVar(value=str(self.home / "diff_report.html"))

        # Status bar text (centered)
        self.var_status = tk.StringVar(value="")

        self._build_ui()
        self._bind_logic()
        self._toggle_extra()
        self._toggle_limits()

        # Resize to fit content precisely (not bigger than needed)
        self.update_idletasks()
        self.geometry("")  # let Tk compute best size

    # ---- utilities ----

    def _group_frame(self, parent, title=None):
        """
        Creates a group container. Honors two independent toggles:
        - GROUP_BORDERS: choose between Labelframe (borders) and Frame (no borders).
        - SHOW_GROUP_TITLES: show or hide the group title text independently.
        """
        if GROUP_BORDERS:
            # Use Labelframe; title is controlled by SHOW_GROUP_TITLES via text
            lf_text = (title or "") if SHOW_GROUP_TITLES else ""
            frm = ttk.Labelframe(parent, text=lf_text)
            return frm
        else:
            # Use plain Frame; optionally add a bold header label if SHOW_GROUP_TITLES is True
            frm = ttk.Frame(parent)
            if SHOW_GROUP_TITLES and title:
                hdr = ttk.Label(frm, text=title, font=("", 10, "bold"))
                hdr.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 6))
            return frm

    def _first_content_row(self):
        """
        Returns the starting row index for content inside a group frame,
        depending on whether a manual header line was added (borders off + titles on).
        """
        if GROUP_BORDERS:
            return 0
        else:
            return 1 if SHOW_GROUP_TITLES else 0

    def _right_label(self, parent, text):
        lbl = ttk.Label(parent, text=text, anchor="e", justify="right", width=LABEL_WIDTH)
        return lbl

    def _button(self, parent, text, command):
        # Consistent button width and styling
        return ttk.Button(parent, text=text, command=command, width=BUTTON_WIDTH)

    # ---- layout ----

    def _build_ui(self):
        root = ttk.Frame(self, padding=12)
        root.grid(row=0, column=0, sticky="nsew")
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        root.columnconfigure(0, weight=1)
    
        # Group 1: Basics
        grp1 = self._group_frame(root, L("grp1_title", "Basics"))
        grp1.grid(row=0, column=0, sticky="ew", pady=(0, GROUP_PAD_Y))
        for c in range(3):
            grp1.columnconfigure(c, weight=1 if c == 1 else 0)
    
        row = self._first_content_row()
    
        # Compare folders (checkbox in same column as text entries → column 1)
        self.chk_compare = ttk.Checkbutton(
            grp1,
            text=L("compare_folders", "Compare folders"),
            variable=self.var_compare_folders
        )
        self.chk_compare.grid(row=row, column=1, sticky="w", **GROUP_INNER_PAD)
        row += 1
    
        # Original
        self._right_label(grp1, L("original", "Original")).grid(row=row, column=0, sticky="e", **GROUP_INNER_PAD)
        self.ent_original = ttk.Entry(grp1, textvariable=self.var_original, width=PATH_ENTRY_WIDTH)
        self.ent_original.grid(row=row, column=1, sticky="ew", **GROUP_INNER_PAD)
        self.btn_browse_original = self._button(grp1, L("browse", "Browse"), self._browse_original)
        self.btn_browse_original.grid(row=row, column=2, padx=BUTTON_PAD, pady=(0, 8), sticky="w")
        row += 1
    
        # Modified
        self._right_label(grp1, L("modified", "Modified")).grid(row=row, column=0, sticky="e", **GROUP_INNER_PAD)
        self.ent_modified = ttk.Entry(grp1, textvariable=self.var_modified, width=PATH_ENTRY_WIDTH)
        self.ent_modified.grid(row=row, column=1, sticky="ew", **GROUP_INNER_PAD)
        self.btn_browse_modified = self._button(grp1, L("browse", "Browse"), self._browse_modified)
        self.btn_browse_modified.grid(row=row, column=2, padx=BUTTON_PAD, pady=(0, 8), sticky="w")
        row += 1
    
        # Worksheet name pattern
        self._right_label(grp1, L("worksheet_pattern", "Worksheet name pattern")).grid(row=row, column=0, sticky="e", **GROUP_INNER_PAD)
        self.ent_wspattern = ttk.Entry(grp1, textvariable=self.var_wspattern, width=PATH_ENTRY_WIDTH)
        self.ent_wspattern.grid(row=row, column=1, sticky="ew", **GROUP_INNER_PAD)
        row += 1
    
        # Header row offset (numeric, same length as other numeric/letter fields)
        self._right_label(grp1, L("row_offset", "Header row offset")).grid(row=row, column=0, sticky="e", **GROUP_INNER_PAD)
        self.spn_row_offset = ttk.Spinbox(grp1, from_=0, to=1_000_000, textvariable=self.var_row_offset, width=SMALL_ENTRY_WIDTH)
        self.spn_row_offset.grid(row=row, column=1, sticky="w", **GROUP_INNER_PAD)
        row += 1
    
        # Source column (letter)
        self._right_label(grp1, L("source_col", "Source column (letter)")).grid(row=row, column=0, sticky="e", **GROUP_INNER_PAD)
        self.ent_source = ttk.Entry(grp1, textvariable=self.var_source_col, width=SMALL_ENTRY_WIDTH)
        self.ent_source.grid(row=row, column=1, sticky="w", **GROUP_INNER_PAD)
        row += 1
    
        # Target column (letter)
        self._right_label(grp1, L("target_col", "Target column (letter)")).grid(row=row, column=0, sticky="e", **GROUP_INNER_PAD)
        self.ent_target = ttk.Entry(grp1, textvariable=self.var_target_col, width=SMALL_ENTRY_WIDTH)
        self.ent_target.grid(row=row, column=1, sticky="w", **GROUP_INNER_PAD)
        row += 1
    
        # Group 2: Matching strategy
        grp2 = self._group_frame(root, L("grp2_title", "Matching strategy"))
        grp2.grid(row=1, column=0, sticky="ew", pady=(0, GROUP_PAD_Y))
        for c in range(3):
            grp2.columnconfigure(c, weight=1 if c == 1 else 0)
        row2 = self._first_content_row()
    
        # Guidance text
        instr = L("no_same_row_found", "If no match is found in the same row:")
        self._right_label(grp2, instr).grid(row=row2, column=0, sticky="e", **GROUP_INNER_PAD)
        ttk.Label(grp2, text="").grid(row=row2, column=1, sticky="w", **GROUP_INNER_PAD)
        row2 += 1
    
        # Similarity tolerance (%)
        self._right_label(grp2, L("tolerate", "Similarity tolerance (%)")).grid(row=row2, column=0, sticky="e", **GROUP_INNER_PAD)
        self.spn_tolerate = ttk.Spinbox(grp2, from_=0, to=35, textvariable=self.var_tolerate, width=SMALL_ENTRY_WIDTH)
        self.spn_tolerate.grid(row=row2, column=1, sticky="w", **GROUP_INNER_PAD)
        row2 += 1
    
        # Realign search window (rows)
        self._right_label(grp2, L("realign", "Realign search window (rows)")).grid(row=row2, column=0, sticky="e", **GROUP_INNER_PAD)
        self.spn_realign = ttk.Spinbox(grp2, from_=0, to=15, textvariable=self.var_realign, width=SMALL_ENTRY_WIDTH)
        self.spn_realign.grid(row=row2, column=1, sticky="w", **GROUP_INNER_PAD)
        row2 += 1
    
        # No limits checkbox (same column as inputs)
        self.chk_nolimits = ttk.Checkbutton(
            grp2,
            text=L("nolimits", "No limits on similarity/realign search"),
            variable=self.var_nolimits,
            command=self._toggle_limits
        )
        self.chk_nolimits.grid(row=row2, column=1, sticky="w", **GROUP_INNER_PAD)
        row2 += 1
    
        # Group 3: Extra column
        grp3 = self._group_frame(root, L("grp3_title", "Extra column"))
        grp3.grid(row=2, column=0, sticky="ew", pady=(0, GROUP_PAD_Y))
        for c in range(3):
            grp3.columnconfigure(c, weight=1 if c == 1 else 0)
        row3 = self._first_content_row()
    
        # Include extra (checkbox same column as entries)
        self.chk_extra = ttk.Checkbutton(
            grp3,
            text=L("extract_extra", "Include text from an additional column"),
            variable=self.var_extract_extra,
            command=self._toggle_extra
        )
        self.chk_extra.grid(row=row3, column=1, sticky="w", **GROUP_INNER_PAD)
        row3 += 1
    
        # Extra column letter
        self._right_label(grp3, L("extra_col", "Extra column (letter)")).grid(row=row3, column=0, sticky="e", **GROUP_INNER_PAD)
        self.ent_extra_col = ttk.Entry(grp3, textvariable=self.var_extra_col, width=SMALL_ENTRY_WIDTH)
        self.ent_extra_col.grid(row=row3, column=1, sticky="w", **GROUP_INNER_PAD)
        row3 += 1
    
        # Extra column header (keep within column; do not stretch to the right border)
        self._right_label(grp3, L("extra_header", "Extra column output header")).grid(row=row3, column=0, sticky="e", **GROUP_INNER_PAD)
        self.ent_extra_header = ttk.Entry(grp3, textvariable=self.var_extra_header, width=PATH_ENTRY_WIDTH)
        # Important: do not use sticky="ew" so it doesn't stretch beyond the button column
        self.ent_extra_header.grid(row=row3, column=1, sticky="w", **GROUP_INNER_PAD)
        row3 += 1
    
        # Group 4: Output + Run
        grp4 = self._group_frame(root, L("grp4_title", "Output"))
        grp4.grid(row=3, column=0, sticky="ew", pady=(0, GROUP_PAD_Y))
        for c in range(3):
            grp4.columnconfigure(c, weight=1 if c == 1 else 0)
        row4 = self._first_content_row()
    
        # Output HTML
        self._right_label(grp4, L("output_html", "Output HTML file")).grid(row=row4, column=0, sticky="e", **GROUP_INNER_PAD)
        self.ent_output = ttk.Entry(grp4, textvariable=self.var_output_html, width=PATH_ENTRY_WIDTH)
        self.ent_output.grid(row=row4, column=1, sticky="ew", **GROUP_INNER_PAD)
        self.btn_browse_output = self._button(grp4, L("browse", "Browse"), self._browse_output)
        self.btn_browse_output.grid(row=row4, column=2, padx=BUTTON_PAD, pady=(0, 8), sticky="w")
        row4 += 1
    
        # NEW: Include identical-target pairs checkbox (checked by default)
        self.chk_include_identical = ttk.Checkbutton(
            grp4,
            text=L("include_identical_pairs",
                   "Include rows with no target changes when rows are misaligned or the source changed"),
            variable=self.var_include_identical_pairs
        )
        self.chk_include_identical.grid(row=row4, column=1, sticky="w", **GROUP_INNER_PAD)
        row4 += 1

        # Run Diff button directly under the last Browse (column 2), consistent padding and width
        self.btn_run = self._button(grp4, L("run_diff", "Run diff"), self._run_diff)
        self.btn_run.grid(row=row4, column=2, sticky="w", padx=BUTTON_PAD, pady=(0, 0))
    
        # Group 5: Status + Exit
        grp5 = self._group_frame(root, L("grp5_title", "Status"))
        grp5.grid(row=4, column=0, sticky="nsew")
        for c in range(3):
            grp5.columnconfigure(c, weight=1 if c == 1 else 0)
        row5 = self._first_content_row()
    
        # Status bar (centered text) with style reference kept for color changes
        status_lbl = ttk.Label(grp5, textvariable=self.var_status, anchor="center", justify="center", style="StatusDefault.TLabel")
        status_lbl.grid(row=row5, column=0, columnspan=3, sticky="ew", pady=(8, 8))
        self.status_lbl = status_lbl
        row5 += 1
    
        # Exit under Run diff: place Exit in column 2 with same padding and width
        self.btn_exit = self._button(grp5, L("exit", "Exit"), self.destroy)
        self.btn_exit.grid(row=row5, column=2, sticky="e", padx=BUTTON_PAD, pady=(0, 0))

#        # Debug button for locales
#        btn_tmp = self._button(grp5, "Diag info", self._diag_locales)
#        btn_tmp.grid(row=row5, column=0, sticky="w", padx=BUTTON_PAD, pady=(0, 0))
    
        # -----------------------------
        # Normalize button widths dynamically based on longest label
        # -----------------------------
        buttons = [
            self.btn_browse_original,
            self.btn_browse_modified,
            self.btn_browse_output,
            self.btn_run,
            self.btn_exit,
        ]
        # Determine the longest label (character count is appropriate since width is in text units)
        longest = max(len(b.cget("text")) for b in buttons if b is not None)
        # Ensure at least the initial BUTTON_WIDTH
        final_width = max(BUTTON_WIDTH, longest)
        for b in buttons:
            if b is not None:
                b.configure(width=final_width)

    # ---- interactions ----

    def _bind_logic(self):
        self.ent_original.bind("<FocusOut>", lambda e: self._sync_start_dirs())
        self.ent_modified.bind("<FocusOut>", lambda e: self._sync_start_dirs())
        self.ent_output.bind("<FocusOut>", lambda e: self._sync_start_dirs())

    def _toggle_extra(self):
        state = "normal" if self.var_extract_extra.get() else "disabled"
        if hasattr(self, "ent_extra_col"):
            self.ent_extra_col.configure(state=state)
        if hasattr(self, "ent_extra_header"):
            self.ent_extra_header.configure(state=state)

    def _toggle_limits(self):
        nolim = self.var_nolimits.get()
        # Update spin ranges
        try:
            self.spn_tolerate.configure(to=100 if nolim else 35)
        except Exception:
            pass
        try:
            self.spn_realign.configure(to=999999 if nolim else 15)
        except Exception:
            pass

    # ----- File/folder choosers with remembered starts -----

    def _browse_original(self):
        start = Path(self.var_original.get()).expanduser()
        start_dir = start if start.is_dir() else (start.parent if start.exists() else self.start_dirs["original"])
        if self.var_compare_folders.get():
            path = filedialog.askdirectory(initialdir=str(start_dir), title=L("choose_original_dir", "Choose original folder"))
            if not path: return
            self.var_original.set(path)
            self.start_dirs["original"] = Path(path)
            self._sync_start_dirs()
        else:
            path = filedialog.askopenfilename(
                initialdir=str(start_dir),
                title=L("choose_original_file", "Choose original file"),
                filetypes=[(L("excel_files", "Excel files"), "*.xlsx *.xls"), (L("all_files", "All files"), "*.*")]
            )
            if not path: return
            self.var_original.set(path)
            self.start_dirs["original"] = Path(path).parent
            self._sync_start_dirs()

    def _browse_modified(self):
        cur_mod = Path(self.var_modified.get()).expanduser()
        if not cur_mod.exists() or str(cur_mod) == str(self.home):
            start_default = Path(self.var_original.get()).expanduser()
            start_default = start_default if start_default.is_dir() else start_default.parent
        else:
            start_default = cur_mod if cur_mod.is_dir() else cur_mod.parent
        start_dir = start_default if start_default.exists() else self.home

        if self.var_compare_folders.get():
            path = filedialog.askdirectory(initialdir=str(start_dir), title=L("choose_modified_dir", "Choose modified folder"))
            if not path: return
            self.var_modified.set(path)
            self.start_dirs["modified"] = Path(path)
        else:
            path = filedialog.askopenfilename(
                initialdir=str(start_dir),
                title=L("choose_modified_file", "Choose modified file"),
                filetypes=[(L("excel_files", "Excel files"), "*.xlsx *.xls"), (L("all_files", "All files"), "*.*")]
            )
            if not path: return
            self.var_modified.set(path)
            self.start_dirs["modified"] = Path(path).parent

    def _browse_output(self):
        start = Path(self.var_output_html.get()).expanduser()
        start_dir = start.parent if start else self.start_dirs["output"]
        path = filedialog.asksaveasfilename(
            initialdir=str(start_dir),
            initialfile=Path(self.var_output_html.get()).name,
            title=L("choose_output_file", "Choose output HTML file"),
            defaultextension=".html",
            filetypes=[("HTML", "*.html"), (L("all_files", "All files"), "*.*")]
        )
        if not path: return
        self.var_output_html.set(path)
        self.start_dirs["output"] = Path(path).parent

    def _sync_start_dirs(self):
        orig = Path(self.var_original.get()).expanduser()
        mod = Path(self.var_modified.get()).expanduser()
        outp = Path(self.var_output_html.get()).expanduser()
        if orig.exists():
            self.start_dirs["original"] = orig if orig.is_dir() else orig.parent
            if not mod.exists() or str(mod) == str(self.home):
                self.start_dirs["modified"] = self.start_dirs["original"]
        if mod.exists():
            self.start_dirs["modified"] = mod if mod.is_dir() else mod.parent
        if outp:
            self.start_dirs["output"] = outp.parent

    # ---- validation/run ----

    def _validate(self):
        compare_dirs = self.var_compare_folders.get()
        orig = Path(self.var_original.get()).expanduser()
        mod = Path(self.var_modified.get()).expanduser()
        out_html = Path(self.var_output_html.get()).expanduser()

        if not self.var_original.get().strip():
            self.var_original.set(str(self.home))
            orig = self.home
        if not self.var_modified.get().strip():
            self.var_modified.set(str(self.home))
            mod = self.home

        if compare_dirs:
            if not orig.exists() or not orig.is_dir():
                self._set_status(L("err_orig_dir", "“Original” must be an existing folder."), error=True)
                return None
            if not mod.exists() or not mod.is_dir():
                self._set_status(L("err_mod_dir", "“Modified” must be an existing folder."), error=True)
                return None
        else:
            if not orig.exists() or not orig.is_file():
                self._set_status(L("err_orig_file", "“Original” must be an existing file."), error=True)
                return None
            if not mod.exists() or not mod.is_file():
                self._set_status(L("err_mod_file", "“Modified” must be an existing file."), error=True)
                return None

        source = self.var_source_col.get().strip()
        target = self.var_target_col.get().strip()
        if not source or not target:
            self._set_status(L("err_cols", "Enter both source and target column letters."), error=True)
            return None

        extra_args = []
        if self.var_extract_extra.get():
            extra_col = self.var_extra_col.get().strip()
            if not extra_col:
                self._set_status(L("err_extra_col", "Extra column is enabled but not set."), error=True)
                return None
            extra_args.extend(["--extra_column", extra_col])
            header = self.var_extra_header.get().strip()
            if header:
                extra_args.extend(["--extra_header", header])

        if not out_html.parent.exists():
            try:
                out_html.parent.mkdir(parents=True, exist_ok=True)
            except Exception:
                self._set_status(L("err_output_dir", "Cannot create the output folder."), error=True)
                return None

        args = [sys.executable, EXCEL_DIFF_SCRIPT,
                "--original", str(orig),
                "--modified", str(mod),
                "--output", str(out_html),
                "--source", source,
                "--target", target]

        if compare_dirs:
            args.append("--dir")

        wsp = self.var_wspattern.get().strip()
        if wsp:
            args.extend(["--wspattern", wsp])

        # Header row offset
        try:
            off = int(self.var_row_offset.get())
        except Exception:
            off = 0
        if off > 0:
            args.extend(["--row-offset", str(off)])

        # Limits and counters
        nolim = self.var_nolimits.get()
        if nolim:
            args.append("--nocap")

        try:
            tol = max(0, int(self.var_tolerate.get()))
        except Exception:
            tol = 0
        try:
            rea = max(0, int(self.var_realign.get()))
        except Exception:
            rea = 0

        if tol > 0:
            args.extend(["--tolerate", str(tol)])
        if rea > 0:
            args.extend(["--realign", str(rea)])

        args.extend(extra_args)
        if not self.var_include_identical_pairs.get():
            args.append("--omit_identical")
        return args, out_html

    def _run_diff(self):
        self._set_status("")  # clear
        data = self._validate()
        if not data:
            return
        args, out_html = data
    
        # Build CLI args for excel_diff.py (strip the interpreter/script)
        # Current args are like: [sys.executable, EXCEL_DIFF_SCRIPT, --flag, value, ...]
        # Keep only flags for argparse in excel_diff.main()
        cli_args = args[:]
        try:
            first_flag_idx = next(
                i for i, a in enumerate(cli_args)
                if a.startswith("--") or a in ("--dir", "--nocap")
            )
            script_argv = cli_args[first_flag_idx:]
        except StopIteration:
            script_argv = []
    
        try:
            if getattr(sys, "frozen", False):
                # In a frozen build, run excel_diff.main() in-process to avoid re-spawning the GUI exe.
                def worker():
                    ok = False
                    try:
                        import excel_diff as _excel_diff
                        old_argv = sys.argv
                        try:
                            sys.argv = ["excel_diff.py"] + script_argv
                            _excel_diff.main()
                            ok = True
                        except SystemExit as e:
                            try:
                                ok = (int(getattr(e, "code", 0)) == 0)
                            except Exception:
                                ok = False
                        finally:
                            sys.argv = old_argv
                    except Exception:
                        ok = False
                    # Marshal UI updates back to the main thread
                    self.after(0, lambda: self._after_diff(ok, out_html))
                threading.Thread(target=worker, daemon=True).start()
            else:
                # Source/dev run: use the system Python interpreter to invoke the helper script
                completed = subprocess.run(args, check=False)
                ok = (completed.returncode == 0)
                self._after_diff(ok, out_html)
        except Exception:
            self._set_status(L("err_run", "Failed to run the diff script."), error=True)
            return
    
    def _after_diff(self, ok, out_html):
        if not ok:
            self._set_status(L("err_run_rc", "The diff script returned a non-zero status."), error=True)
            return
        self._set_status(L("done_status", "Diff finished successfully."), error=False)
        if messagebox.askyesno(L("done_title", "Done"), L("open_now", "Open the HTML report now?")):
            self._open_file(out_html)

    def _set_status(self, text, error=False):
        self.var_status.set(text)
        # switch label style to show red on errors, default otherwise
        try:
            self.status_lbl.configure(style="StatusError.TLabel" if error else "StatusDefault.TLabel")
        except Exception:
            # Fallback direct color if theme ignores styles
            self.status_lbl.configure(foreground="#b00000" if error else "")

    def _open_file(self, path: Path):
        try:
            if sys.platform.startswith("darwin"):
                subprocess.run(["open", str(path)])
            elif os.name == "nt":
                os.startfile(str(path))  # type: ignore[attr-defined]
            else:
                subprocess.run(["xdg-open", str(path)])
        except Exception:
            messagebox.showinfo(L("open_fail_title", "Info"), L("open_fail", "Cannot open the file automatically."))

#    def _diag_locales(self):
#        try:
#            msgs = []
#            msgs.append(f"APP_DIR={APP_DIR}")
#            msgs.append(f"LOCALES_DIR exists={LOCALES_DIR.exists()} path={LOCALES_DIR}")
#            chain = detect_locale_chain()
#            msgs.append(f"detect_locale_chain={chain}")
#            available = []
#            chosen = None
#            if LOCALES_DIR.exists():
#                # Show exact filenames present in the bundle’s locales dir
#                files = sorted(p.name for p in LOCALES_DIR.iterdir() if p.is_file())
#                available = files
#                msgs.append(f"locales files={files}")
#                # Determine which candidate actually matches first
#                for cand in chain:
#                    f = LOCALES_DIR / f"{cand}_gui.json"
#                    if f.exists():
#                        chosen = f.name
#                        break
#            msgs.append(f"locale_choice={chosen if chosen else 'None'}")
#            msgs.append(f"LABELS keys={len(LABELS)}")
#            self._set_status("\n".join(msgs), error=False)
#        except Exception as e:
#            self._set_status(f"Locales diag error: {e}", error=True)

def main():
    app = ExcelDiffGUI()
    app.mainloop()

if __name__ == "__main__":
    main()
