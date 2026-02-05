"""
Deps:
  pip install discord.py ttkbootstrap pygame-ce
  (pygame-ce provides 'import pygame' and works on Python 3.14)

Build (PyInstaller):
  py -3.14 -m PyInstaller --noconfirm --clean --onefile --windowed --name TalkAlert --icon TalkAlert.ico ^
    --collect-all ttkbootstrap --collect-all pygame --collect-all discord TalkAlert.py
"""

from __future__ import annotations

import asyncio
import json
import os
import sys
import threading
import time
import urllib.parse
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import tkinter.font as tkfont

import ttkbootstrap as tb
from ttkbootstrap.constants import *


APP_NAME = "TalkAlert"
FORCE_EXIT = True  # close app -> ensure process exits

CONFIG_DIR = Path(os.environ.get("APPDATA", str(Path.home()))) / APP_NAME
CONFIG_PATH = CONFIG_DIR / "config.json"

ALLOWED_AUDIO = (".wav", ".mp3")

PUSHOVER_API_URL = "https://api.pushover.net/1/messages.json"

def bind_edit_context_menu(widget: tk.Widget):
    """Attach a right-click context menu (Cut/Copy/Paste/Select All) to Entry/Text widgets."""
    menu = tk.Menu(widget, tearoff=0)

    def _cut():
        try:
            widget.event_generate("<<Cut>>")
        except Exception:
            pass

    def _copy():
        try:
            widget.event_generate("<<Copy>>")
        except Exception:
            pass

    def _paste():
        try:
            widget.event_generate("<<Paste>>")
        except Exception:
            pass

    def _select_all():
        try:
            widget.focus_set()
        except Exception:
            pass
        # Entry-like
        try:
            widget.selection_range(0, tk.END)  # type: ignore[attr-defined]
            widget.icursor(tk.END)  # type: ignore[attr-defined]
            return
        except Exception:
            pass
        # Text-like
        try:
            widget.tag_add("sel", "1.0", "end-1c")  # type: ignore[attr-defined]
            widget.mark_set("insert", "end-1c")  # type: ignore[attr-defined]
        except Exception:
            pass

    menu.add_command(label="切り取り", command=_cut)
    menu.add_command(label="コピー", command=_copy)
    menu.add_command(label="貼り付け", command=_paste)
    menu.add_separator()
    menu.add_command(label="全選択", command=_select_all)

    def _popup(event):
        try:
            widget.focus_set()
        except Exception:
            pass

        # Enable/disable Cut/Copy based on selection (best-effort)
        has_sel = False
        try:
            if hasattr(widget, "selection_present"):
                has_sel = bool(widget.selection_present())  # type: ignore[attr-defined]
        except Exception:
            has_sel = False
        state = "normal" if has_sel else "disabled"
        try:
            menu.entryconfig("切り取り", state=state)
            menu.entryconfig("コピー", state=state)
        except Exception:
            pass

        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            try:
                menu.grab_release()
            except Exception:
                pass

    # Windows/Linux right click
    widget.bind("<Button-3>", _popup, add="+")
    # macOS Ctrl+Click fallback
    widget.bind("<Control-Button-1>", _popup, add="+")

def resource_path(rel: str) -> Path:
    """Return absolute path to a resource (works for PyInstaller onefile)."""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / rel


# Optional deps
try:
    import pygame  # pygame-ce works as 'pygame'
    _HAS_PYGAME = True
except Exception:
    pygame = None
    _HAS_PYGAME = False

try:
    import discord
    _HAS_DISCORD = True
except Exception:
    discord = None
    _HAS_DISCORD = False

try:
    import pystray
    from PIL import Image, ImageDraw
    _HAS_TRAY = True
except Exception:
    pystray = None
    Image = None
    ImageDraw = None
    _HAS_TRAY = False


@dataclass
class Rule:
    name: str
    user_id: str
    sound_path: str
    volume: int = 100  # 0-100

    pushover_sound: str = ""  # empty -> device default
    @property
    def sound_filename(self) -> str:
        try:
            return Path(self.sound_path).name
        except Exception:
            return self.sound_path


class TalkAlertApp(tb.Window):
    def __init__(self):
        super().__init__(
            title=APP_NAME,
            themename="flatly",
            size=(720, 630),  # ↑ footer text visible on high-DPI

            resizable=(False, False),
        )

        self.protocol("WM_DELETE_WINDOW", lambda: self._confirm_exit())

        self._apply_app_icon()

        # runtime state
        self.stop_event = threading.Event()
        self._after_ids = set()

        self.rules: List[Rule] = []
        self.muted = False
        self.bot_token: str = ""  # loaded from config
        self.tray_on_minimize: bool = True  # minimize -> tray (default ON)
        self._in_tray = False
        self._tray_icon = None
        self._tray_thread = None

        # pushover (iOS push)
        self.pushover_enabled: bool = False
        self.pushover_user_key: str = ""
        self.pushover_app_token: str = ""
        self.pushover_push_when_muted: bool = True
        self.pushover_include_message: bool = True  # include message text in push

        # discord runtime
        self._discord_client = None
        self._discord_loop: Optional[asyncio.AbstractEventLoop] = None
        self._bot_thread: Optional[threading.Thread] = None
        self._bot_thread_lock = threading.Lock()

        # bot status UI
        self._bot_state = "offline"  # offline | connecting | online
        self._blink_on = True

        # audio init
        self._audio_ready = False
        self._init_audio()
        self._now_playing_rule_id: Optional[str] = None
        self._now_playing_volume: int = 100

        # fonts (thin & clean)
        self._setup_fonts()

        # load config (rules/mute/token)
        self._load_config()

        # build UI
        self._build_ui()

        # minimize -> tray
        self.bind("<Unmap>", self._on_unmap)
        self._start_minimize_watchdog()

        # fill table
        self._refresh_table()

        # auto-start bot if token exists
        self._auto_start_bot()

        self.bind("<Escape>", lambda _e: self.on_close())

    # ------------------------------------------------------------
    # Fonts / style
    # ------------------------------------------------------------
    def _setup_fonts(self):
        preferred = [
            "Meiryo UI",
        ]
        available = set(tkfont.families(self))
        family = next((f for f in preferred if f in available), "Segoe UI")

        self.font_base = (family, 10)
        self.font_small = (family, 9)
        self.font_head = (family, 11)  # not bold

        try:
            self.style.configure("TLabel", font=self.font_base)
            self.style.configure("TButton", font=self.font_base)
            self.style.configure("TEntry", font=self.font_base)

            # Table font (change here if you want different)
            self.style.configure("Treeview", font=("Meiryo UI", 10), rowheight=26)
            self.style.configure("Treeview.Heading", font=("Meiryo UI", 9))
        except Exception:
            pass

    def _apply_app_icon(self):
        """Set window/taskbar icon to TalkAlert.ico (and improve grouping on Windows)."""
        if os.name == "nt":
            try:
                import ctypes  # noqa
                ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_NAME)
            except Exception:
                pass

        # Prefer ICO for Windows taskbar; also try PNG for better scaling.
        ico_candidates = [
            "TalkAlert.ico",
            "TalkAlert_icon_clean_multi.ico",
            "TalkAlert_icon_big_multi.ico",
        ]
        png_candidates = [
            "TalkAlert_icon_256.png",
            "TalkAlert_icon_clean_512.png",
            "TalkAlert_icon_clean_1024.png",
            "TalkAlert_icon_1024.png",
        ]

        try:
            for fn in ico_candidates:
                p = resource_path(fn)
                if p.exists():
                    try:
                        self.iconbitmap(str(p))
                        break
                    except Exception:
                        pass
        except Exception:
            pass

        # iconphoto (PNG) helps on some setups; keep a reference to avoid GC.
        try:
            for fn in png_candidates:
                p = resource_path(fn)
                if p.exists():
                    try:
                        self._tk_icon_img = tk.PhotoImage(file=str(p))
                        self.iconphoto(True, self._tk_icon_img)
                        break
                    except Exception:
                        pass
        except Exception:
            pass


    # ------------------------------------------------------------
    # UI building
    # ------------------------------------------------------------
    def _build_ui(self):
        root = tb.Frame(self, padding=12)
        root.pack(fill=BOTH, expand=YES)

        header = tb.Frame(root)
        header.pack(fill=X)

        bg = self.style.colors.bg
        self.dot_canvas = tk.Canvas(header, width=16, height=16, highlightthickness=0, bg=bg)
        self.dot_canvas.pack(side=LEFT, padx=(0, 6), pady=(2, 0))
        self._dot_id = self.dot_canvas.create_oval(3, 3, 13, 13, fill="#e74c3c", outline="")

        self.status_var = tk.StringVar(value="Bot: offline")
        self.lbl_status = tb.Label(
            header,
            textvariable=self.status_var,
            font=self.font_head,
            bootstyle="secondary",
            width=38,  # fixed width -> no jitter
            anchor="w",
        )
        self.lbl_status.pack(side=LEFT)

        # Right: Settings + Mute
        self.btn_settings = tb.Button(
            header,
            text="⚙",
            command=self.open_settings,
            bootstyle="secondary",
            width=3,
        )
        self.btn_settings.pack(side=RIGHT)

        self.btn_mute = tb.Button(
            header,
            text="🔇 Mute: OFF",
            command=self.toggle_mute,
            bootstyle="warning",
            width=14,
        )
        self.btn_mute.pack(side=RIGHT, padx=(0, 10))

        # Rules card
        rules_card = tb.Frame(root, padding=(12, 10))
        rules_card.pack(fill=BOTH, expand=YES, pady=(12, 0))

        top = tb.Frame(rules_card)
        top.pack(fill=X)
        tb.Label(top, text="Rules", font=self.font_head, bootstyle="secondary").pack(side=LEFT)
        ttk.Separator(rules_card).pack(fill=X, pady=(8, 10))

        body = tb.Frame(rules_card)
        body.pack(fill=BOTH, expand=YES)

        self.tree = ttk.Treeview(body, columns=("name", "user_id", "sound", "vol", "push"), show="headings", height=8)
        # Headings (Nameはクリックでソート)
        self.tree.heading("name", text="Name", anchor=tk.W)
        self.tree.heading("user_id", text="UserID", anchor=tk.W)
        self.tree.heading("sound", text="Sound", anchor=tk.W)
        self.tree.heading("vol", text="Vol", anchor=tk.W)
        self.tree.heading("push", text="Push音", anchor=tk.W)

        self.tree.column("name", width=150, anchor=tk.W)
        self.tree.column("user_id", width=190, anchor=tk.W)
        self.tree.column("sound", width=150, anchor=tk.W)
        self.tree.column("vol", width=30, anchor=tk.W)
        self.tree.column("push", width=90, anchor=tk.W)

        vsb = ttk.Scrollbar(body, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns", padx=(8, 0))
        body.rowconfigure(0, weight=1)
        body.columnconfigure(0, weight=1)

        self.tree.bind("<<TreeviewSelect>>", self._load_selected_to_form)

        # クリックで Name ソート（▲▼）
        self._name_sort_active = False
        self._name_sort_asc = True
        self._update_name_heading()

        # ドラッグで並び替え（空行は作らない）
        self._drag_iid = None
        self._dragging = False
        self._drag_win = None
        self.tree.configure(cursor="hand2")
        self.tree.bind("<ButtonPress-1>", self._on_tree_press, add="+")
        self.tree.bind("<B1-Motion>", self._on_tree_motion, add="+")
        self.tree.bind("<ButtonRelease-1>", self._on_tree_release, add="+")

        # Form
        form = tb.Frame(root, padding=(12, 10))
        form.pack(fill=X, pady=(8, 0))

        # Row 0: labels
        tb.Label(form, text="Name (任意)", bootstyle="secondary").grid(row=0, column=0, sticky=W)
        tb.Label(form, text="UserID", bootstyle="secondary").grid(row=0, column=1, sticky=W)
        tb.Label(form, text="Sound (wav/mp3)", bootstyle="secondary").grid(row=0, column=2, sticky=W)

        # Row 1: entries + Browse/Test
        self.entry_name = tb.Entry(form)
        self.entry_id = tb.Entry(form)
        self.entry_sound = tb.Entry(form)

        self.entry_name.grid(row=1, column=0, sticky=EW, padx=(0, 10), pady=(4, 0))
        self.entry_id.grid(row=1, column=1, sticky=EW, padx=(0, 10), pady=(4, 0))
        self.entry_sound.grid(row=1, column=2, sticky=EW, padx=(0, 10), pady=(4, 0))

        # Right-click edit menu
        bind_edit_context_menu(self.entry_name)
        bind_edit_context_menu(self.entry_id)
        bind_edit_context_menu(self.entry_sound)

        btn_browse = tb.Button(form, text="Browse", command=self.browse_sound, bootstyle="secondary", width=8)
        btn_browse.grid(row=1, column=3, sticky=E, pady=(4, 0))

        btn_test = tb.Button(form, text="▶ Test", command=self.test_form, bootstyle="secondary", width=8)
        btn_test.grid(row=1, column=4, sticky=E, pady=(4, 0), padx=(8, 0))

        # Row 2/3: Pushover push sound (per rule) + Volume
        tb.Label(form, text="Push音(Pushover)", bootstyle="secondary").grid(row=2, column=0, sticky=W, pady=(10, 0))
        tb.Label(form, text="Volume", bootstyle="secondary").grid(row=2, column=2, sticky=W, pady=(10, 0))

        self.var_pushover_sound = tk.StringVar(value="")
        self.entry_pushover_sound = tb.Entry(form, textvariable=self.var_pushover_sound)
        self.entry_pushover_sound.grid(row=3, column=0, columnspan=1, sticky=EW, padx=(0, 10), pady=(4, 0))

        bind_edit_context_menu(self.entry_pushover_sound)

        self.var_volume = tk.IntVar(value=100)
        self.scale_volume = tb.Scale(
            form,
            from_=0,
            to=100,
            length=300,
            orient=HORIZONTAL,
            variable=self.var_volume,
            command=lambda _v: self._update_volume_label(),
        )
        self.scale_volume.grid(row=3, column=2, columnspan=2, sticky=EW, pady=(4, 0), padx=(0, 10))
        self.lbl_volume = tb.Label(form, text="100%", bootstyle="secondary", width=5, anchor=E)
        self.lbl_volume.grid(row=3, column=4, sticky=W, pady=(4, 0))
        self._update_volume_label()

        # Column weights
        form.columnconfigure(0, weight=1)
        form.columnconfigure(1, weight=1)
        form.columnconfigure(2, weight=2)

        # Actions
        actions = tb.Frame(root)
        actions.pack(fill=X, pady=(10, 0))

        tb.Button(actions, text="＋ Add", command=self.add_rule, bootstyle="success", width=12).pack(side=LEFT)
        tb.Button(actions, text="⟳ Update", command=self.update_rule, bootstyle="primary", width=12).pack(side=LEFT, padx=(10, 0))
        tb.Button(actions, text="🗑 Remove", command=self.remove_rule, bootstyle="danger", width=12).pack(side=LEFT, padx=(10, 0))

        # Footer note (keep)
##note = (
##            "※ Discordの“通常通知音”はDiscord側が鳴らしている音です。TalkAlertからは止められません。\n"
##            "   TalkAlertの音だけにしたい場合は Discord の通知設定で「メッセージ通知音」をOFFにしてください。"
##        )
##        tb.Label(root, text=note, bootstyle="secondary", justify=LEFT, font=self.font_small, wraplength=690).pack(fill=X, pady=(8, 0))
        tb.Label(root, text=f"Config: {CONFIG_PATH}", bootstyle="secondary", font=self.font_small).pack(fill=X, pady=(4, 0))

        # initial bot status display
        if self.bot_token:
            self._set_bot_state("offline", "Bot: offline")
        else:
            self._set_bot_state("offline", "Bot: TOKEN未設定（⚙で設定）")

        self._tick_blink()

        # ------------------------------------------------------------
    # Tray (minimize to tray)
    # ------------------------------------------------------------
    def _build_tray_image(self):
        """Create tray icon. Prefer bundled ICO/PNG so it matches taskbar icon."""
        if Image is None:
            return None

        # Try to load from bundled files first
        candidates = [
            "TalkAlert.ico",
            "TalkAlert_icon_clean_multi.ico",
            "TalkAlert_icon_big_multi.ico",
            "TalkAlert_icon_clean_1024.png",
            "TalkAlert_icon_1024.png",
            "TalkAlert_icon_256.png",
        ]
        for fn in candidates:
            try:
                p = resource_path(fn)
                if not p.exists():
                    continue
                img = Image.open(p)

                # If ICO has multiple sizes, pick the largest when possible.
                try:
                    if hasattr(img, "sizes") and img.sizes:
                        best = max(img.sizes, key=lambda s: s[0] * s[1])
                        if hasattr(img, "getimage"):
                            img = img.getimage(best)
                except Exception:
                    pass

                img = img.convert("RGBA").resize((64, 64))
                return img
            except Exception:
                continue

        # Fallback: draw a simple icon
        if ImageDraw is None:
            return None
        size = 64
        img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        pad = 8
        d.rounded_rectangle(
            (pad, pad, size - pad, size - pad),
            radius=14,
            fill=(255, 255, 255, 255),
            outline=(180, 180, 180, 255),
            width=2,
        )
        d.rounded_rectangle((18, 22, 44, 34), radius=7, fill=(220, 220, 220, 255))
        d.rounded_rectangle((22, 36, 48, 48), radius=7, fill=(200, 200, 200, 255))
        d.ellipse((40, 18, 50, 28), fill=(46, 204, 113, 255))
        return img
    def _toggle_visibility_from_tray(self):
        """Toggle window visibility from tray (double-click)."""
        try:
            # If visible, hide to tray (force)
            if self.winfo_viewable() and self.state() != "iconic":
                if not _HAS_TRAY:
                    return
                if not self._start_tray():
                    return
                self._in_tray = True
                self.withdraw()
                return
        except Exception:
            pass

        # Otherwise, show
        self._show_from_tray()





    def _ensure_tray_icon(self) -> bool:
        if not _HAS_TRAY:
            return False
        if self._tray_icon is not None:
            return True

        image = self._build_tray_image()
        if image is None:
            return False

        def on_open(_icon=None, _item=None):
            # Open acts as show/hide toggle (double-click too)
            self._ui_call(self._toggle_visibility_from_tray)

        def on_mute(_icon=None, _item=None):
            self._ui_call(self.toggle_mute)

        def on_exit(_icon=None, _item=None):
            self._ui_call(self._confirm_exit)

        menu = pystray.Menu(
            pystray.MenuItem("Open", on_open, default=True),
            pystray.MenuItem("Mute", on_mute, checked=lambda _i: bool(self.muted)),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Exit", on_exit),
        )

        try:
            self._tray_icon = pystray.Icon(APP_NAME, image, APP_NAME, menu)
            try:
                self._tray_icon.on_activate = on_open
            except Exception:
                pass
        except Exception:
            self._tray_icon = None
            return False

        return True


    def _start_tray(self) -> bool:
        """Start tray icon. Returns True if started/visible."""
        if not self._ensure_tray_icon():
            return False

        # Prefer run_detached when available (more reliable on Windows).
        try:
            if hasattr(self._tray_icon, "run_detached"):
                self._tray_icon.run_detached()
                return True
        except Exception:
            # Fall back to threaded run()
            pass

        if self._tray_thread and self._tray_thread.is_alive():
            return True

        def run_icon():
            try:
                self._tray_icon.run()
            except Exception:
                pass

        self._tray_thread = threading.Thread(target=run_icon, daemon=True)
        self._tray_thread.start()
        return True

    def _stop_tray(self):
        try:
            if self._tray_icon:
                self._tray_icon.stop()
        except Exception:
            pass

        t = self._tray_thread
        if t and t.is_alive():
            try:
                t.join(timeout=1.5)
            except Exception:
                pass

        self._tray_thread = None
        self._tray_icon = None

    def _hide_to_tray(self):
        if self._in_tray:
            return
        if not self.tray_on_minimize:
            return
        if not _HAS_TRAY:
            return  # tray feature unavailable

        # Start tray first; if it fails, do not withdraw (avoid "vanishing" app)
        started = False
        try:
            self._in_tray = True
            started = self._start_tray()
        except Exception:
            started = False

        if not started:
            self._in_tray = False
            return

        # Optional: notify so user realizes it went to tray (may be hidden under ^)
        try:
            if self._tray_icon and hasattr(self._tray_icon, "notify"):
                self._tray_icon.notify("TalkAlert", "タスクトレイに常駐しました（^ に隠れている場合があります）")
        except Exception:
            pass

        try:
            self.withdraw()  # hide from taskbar
        except Exception:
            pass

    def _show_from_tray(self):
        try:
            self._in_tray = False
            self.deiconify()
            self.state("normal")
            self.update_idletasks()
            self.lift()
            self.focus_force()
        except Exception:
            pass



        self._in_tray = False
        try:
            self.deiconify()
            self.state("normal")
            self.lift()
            self.focus_force()
        except Exception:
            pass

        self._stop_tray()

    def _on_unmap(self, _evt=None):
        # Triggered when window is minimized (iconic) or withdrawn.
        # On some systems, state() is not yet updated at the event timing, so re-check after a short delay.
        def late_check():
            try:
                if self.tray_on_minimize and self.state() == "iconic" and not self._in_tray:
                    self._hide_to_tray()
            except Exception:
                pass
        try:
            self.after(80, late_check)
        except Exception:
            pass

    def _start_minimize_watchdog(self):
        """Fallback watcher: ensures minimize-to-tray works even if events are missed."""
        def poll():
            try:
                if self.tray_on_minimize and not self._in_tray and self.state() == "iconic":
                    self._hide_to_tray()
            except Exception:
                pass
            finally:
                aid = self.after(350, poll)
                self._after_ids.add(aid)

        aid = self.after(350, poll)
        self._after_ids.add(aid)



# ------------------------------------------------------------
    # Settings dialog (Token)
    # ------------------------------------------------------------
    def open_settings(self):
        dlg = tb.Toplevel(self)
        dlg.title("Settings")
        dlg.resizable(False, False)
        dlg.transient(self)
        dlg.grab_set()

        frame = tb.Frame(dlg, padding=16)
        frame.pack(fill=BOTH, expand=YES)

        tb.Label(frame, text="Discord Bot Token", font=self.font_head).pack(anchor="w")

        hint = (
            "ここにBotのTOKENを入力してください。\n"
            "入力したTOKENはPC内（AppData）の設定ファイルに保存されます。"
        )
        tb.Label(frame, text=hint, bootstyle="secondary", justify=LEFT, font=self.font_small).pack(anchor="w", pady=(6, 10))

        token_var = tk.StringVar(value=self.bot_token)
        show_var = tk.BooleanVar(value=False)

        entry = tb.Entry(frame, textvariable=token_var, show="•", width=62)
        entry.pack(fill=X)
        entry.focus_set()

        bind_edit_context_menu(entry)

        def toggle_show():
            entry.configure(show="" if show_var.get() else "•")

        tb.Checkbutton(
            frame,
            text="表示する",
            variable=show_var,
            command=toggle_show,
            bootstyle="secondary",
        ).pack(anchor="w", pady=(8, 0))

        tray_var = tk.BooleanVar(value=self.tray_on_minimize)
        tb.Checkbutton(
            frame,
            text="最小化時にタスクトレイに常駐する",
            variable=tray_var,
            bootstyle="secondary",
        ).pack(anchor="w", pady=(6, 0))


        # ---- Pushover ----
        tb.Separator(frame).pack(fill=X, pady=(14, 10))
        tb.Label(frame, text="Pushover (iOS Push)", font=self.font_head).pack(anchor="w")
        tb.Label(
            frame,
            text="iPhoneにPushoverアプリを入れて、User Key と Application Token を設定すると\n対象ユーザー発言時にプッシュ通知が届きます（PCでTalkAlertが起動している間のみ）。",
            bootstyle="secondary",
            justify=LEFT,
            font=self.font_small,
        ).pack(anchor="w", pady=(6, 10))

        po_enabled_var = tk.BooleanVar(value=self.pushover_enabled)
        po_when_muted_var = tk.BooleanVar(value=self.pushover_push_when_muted)
        po_include_msg_var = tk.BooleanVar(value=self.pushover_include_message)
        po_user_var = tk.StringVar(value=self.pushover_user_key)
        po_app_var = tk.StringVar(value=self.pushover_app_token)

        tb.Checkbutton(frame, text="Push通知を有効にする", variable=po_enabled_var, bootstyle="secondary").pack(anchor="w")

        row = tb.Frame(frame)
        row.pack(fill=X, pady=(8, 0))
        tb.Label(row, text="User Key", width=10, font=self.font_small).pack(side=LEFT)
        entry_po_user = tb.Entry(row, textvariable=po_user_var, width=55)
        entry_po_user.pack(side=LEFT, fill=X, expand=YES)
        bind_edit_context_menu(entry_po_user)

        row2 = tb.Frame(frame)
        row2.pack(fill=X, pady=(6, 0))
        tb.Label(row2, text="App Token", width=10, font=self.font_small).pack(side=LEFT)
        entry_po_app = tb.Entry(row2, textvariable=po_app_var, width=55, show="•")
        entry_po_app.pack(side=LEFT, fill=X, expand=YES)
        bind_edit_context_menu(entry_po_app)

        tb.Label(
            frame,
            text="通知音はルールごとに指定できます。\nメイン画面の『Pushover用Push音』に Pushover のサウンド名を入力してください（空欄なら端末デフォルト）。",
            bootstyle="secondary",
            justify=LEFT,
            font=self.font_small,
        ).pack(anchor="w", pady=(6, 0))

        tb.Checkbutton(frame, text="Mute中もPushを送る", variable=po_when_muted_var, bootstyle="secondary").pack(anchor="w", pady=(6, 0))
        tb.Checkbutton(frame, text="Pushにメッセージ本文も含める", variable=po_include_msg_var, bootstyle="secondary").pack(anchor="w", pady=(2, 0))

        test_row = tb.Frame(frame)
        test_row.pack(fill=X, pady=(8, 0))

        def do_test_push():
            u = po_user_var.get().strip()
            t = po_app_var.get().strip()
            if not (u and t):
                status.set("PushoverのUser Key / App Token を入力してください。")
                return
            status.set("テスト通知を送信中…")

            def worker():
                ok, err = self._pushover_request_sync(t, u, APP_NAME, "TalkAlert テスト通知です。\n（通知音は端末デフォルト。音を変える場合はメイン画面の『Pushover用Push音』で設定）")
                self._ui_call(lambda: status.set("テスト通知: 送信OK" if ok else f"テスト通知: 送信失敗 ({err})"))

            threading.Thread(target=worker, daemon=True).start()

        tb.Button(test_row, text="Test Push", command=do_test_push, bootstyle="info", width=10).pack(side=LEFT)


        status = tk.StringVar(value="")
        tb.Label(frame, textvariable=status, bootstyle="secondary").pack(anchor="w", pady=(8, 0))

        btns = tb.Frame(frame)
        btns.pack(fill=X, pady=(8, 0))

        def do_save():
            tok = token_var.get().strip()
            self.tray_on_minimize = bool(tray_var.get())

            # Pushover settings
            self.pushover_enabled = bool(po_enabled_var.get())
            self.pushover_push_when_muted = bool(po_when_muted_var.get())
            self.pushover_include_message = bool(po_include_msg_var.get())
            self.pushover_user_key = po_user_var.get().strip()
            self.pushover_app_token = po_app_var.get().strip()

            # If token field is empty: keep existing token (do not overwrite),
            # but still save other settings.
            if tok:
                if len(tok) < 20:
                    status.set("TOKENが短すぎます。")
                    return
                token_changed = (tok != self.bot_token)
                self.bot_token = tok
            else:
                token_changed = False

            self._save_config()

            # Apply immediately
            if self.bot_token:
                if token_changed:
                    self._restart_bot_async()
                else:
                    # ensure running
                    self._auto_start_bot()
            else:
                self._stop_bot_async()
                self._set_bot_state("offline", "Bot: TOKEN未設定（⚙で設定）")

            dlg.destroy()

        def do_clear():
            if messagebox.askyesno(APP_NAME, "保存済みTOKENを削除しますか？（Botは停止します）"):
                self.bot_token = ""
                self.tray_on_minimize = bool(tray_var.get())
                # Pushover settings (also saved)
                self.pushover_enabled = bool(po_enabled_var.get())
                self.pushover_push_when_muted = bool(po_when_muted_var.get())
                self.pushover_include_message = bool(po_include_msg_var.get())
                self.pushover_user_key = po_user_var.get().strip()
                self.pushover_app_token = po_app_var.get().strip()
                self._save_config()
                self._stop_bot_async()
                self._set_bot_state("offline", "Bot: TOKEN未設定（⚙で設定）")
                dlg.destroy()

        tb.Button(btns, text="Save", command=do_save, bootstyle="success", width=10).pack(side=LEFT)
        tb.Button(btns, text="Clear", command=do_clear, bootstyle="danger", width=10).pack(side=LEFT, padx=(10, 0))
        tb.Button(btns, text="Close", command=dlg.destroy, bootstyle="secondary", width=10).pack(side=RIGHT)

        dlg.bind("<Return>", lambda _e: do_save())

    # ------------------------------------------------------------
    # Config
    # ------------------------------------------------------------
    def _ensure_config_dir(self):
        try:
            CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass

    def _load_config(self):
        self._ensure_config_dir()
        if not CONFIG_PATH.exists():
            return
        try:
            data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            self.muted = bool(data.get("mute", False))
            self.bot_token = str(data.get("token", "") or "").strip()
            self.tray_on_minimize = bool(data.get("tray_on_minimize", True))

            # pushover (iOS push)
            self.pushover_enabled = bool(data.get("pushover_enabled", False))
            self.pushover_user_key = str(data.get("pushover_user_key", "") or "").strip()
            self.pushover_app_token = str(data.get("pushover_app_token", "") or "").strip()
            # legacy: older versions had a global pushover_sound; migrate to per-rule default when present
            legacy_po_sound = str(data.get("pushover_sound", "") or "").strip()
            self.pushover_push_when_muted = bool(data.get("pushover_push_when_muted", True))
            self.pushover_include_message = bool(data.get("pushover_include_message", True))

            self.rules = []
            for r in data.get("rules", []):
                user_id = str(r.get("user_id", "")).strip()
                if not user_id:
                    continue

                po_sound = str(r.get("pushover_sound", "") or "").strip()
                if not po_sound and legacy_po_sound:
                    po_sound = legacy_po_sound

                self.rules.append(
                    Rule(
                        name=str(r.get("name", "")).strip(),
                        user_id=user_id,
                        sound_path=str(r.get("sound_path", "")).strip(),
                        volume=max(0, min(100, int(r.get("volume", 100) or 100))),
                        pushover_sound=po_sound,
                    )
                )
        except Exception:
            self.rules = []

    def _save_config(self):
        self._ensure_config_dir()
        try:
            data = {
                "mute": self.muted,
                "token": self.bot_token,
                "tray_on_minimize": self.tray_on_minimize,

                # pushover (iOS push)
                "pushover_enabled": self.pushover_enabled,
                "pushover_user_key": self.pushover_user_key,
                "pushover_app_token": self.pushover_app_token,
                "pushover_push_when_muted": self.pushover_push_when_muted,
                "pushover_include_message": self.pushover_include_message,

                "rules": [
                    {
                        "name": r.name,
                        "user_id": r.user_id,
                        "sound_path": r.sound_path,
                        "volume": int(getattr(r, "volume", 100)),
                        "pushover_sound": str(getattr(r, "pushover_sound", "") or "").strip(),
                    }
                    for r in self.rules
                ],
            }
            CONFIG_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass

    def _refresh_table(self):
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        for r in self.rules:
            self.tree.insert("", "end", iid=r.user_id, values=(r.name, r.user_id, r.sound_filename, f"{int(getattr(r, 'volume', 100))}%", getattr(r, 'pushover_sound', '')))

        if self.muted:
            self.btn_mute.configure(text="🔇 Mute: ON", bootstyle="danger")
        else:
            self.btn_mute.configure(text="🔇 Mute: OFF", bootstyle="warning")

    


    # ------------------------------------------------------------
    # Sort / Drag & Drop for rules table
    # ------------------------------------------------------------
    def _update_name_heading(self):
        try:
            base = "Name"
            if getattr(self, "_name_sort_active", False):
                base = base + (" ▲" if getattr(self, "_name_sort_asc", True) else " ▼")
            self.tree.heading("name", text=base, anchor=tk.W, command=self._on_name_heading_click)
        except Exception:
            pass

    def _on_name_heading_click(self):
        # toggle sort order (ascending/descending)
        try:
            if not getattr(self, "_name_sort_active", False):
                self._name_sort_active = True
                self._name_sort_asc = True
            else:
                self._name_sort_asc = not getattr(self, "_name_sort_asc", True)

            asc = getattr(self, "_name_sort_asc", True)
            self.rules.sort(key=lambda r: (r.name or "").casefold())
            if not asc:
                self.rules.reverse()

            self._save_config()
            self._refresh_table()
            self._update_name_heading()
        except Exception:
            pass

    def _sync_rules_from_tree_order(self):
        """Treeviewの表示順を self.rules に反映する。"""
        try:
            order = list(self.tree.get_children())
            by_id = {str(r.user_id): r for r in self.rules}
            new_rules = []
            for iid in order:
                if iid in by_id:
                    new_rules.append(by_id[iid])
            # 念のため：Treeにいないものがあれば末尾に残す
            for r in self.rules:
                if str(r.user_id) not in order:
                    new_rules.append(r)
            self.rules = new_rules
        except Exception:
            pass

    def _show_drag_hint(self, x_root: int, y_root: int, text_: str):
        try:
            if self._drag_win is None:
                w = tk.Toplevel(self)
                w.overrideredirect(True)
                try:
                    w.attributes("-topmost", True)
                except Exception:
                    pass
                lbl = tk.Label(w, text=text_, padx=8, pady=3, relief="solid", borderwidth=1)
                lbl.pack()
                self._drag_win = w
            self._drag_win.geometry(f"+{x_root + 12}+{y_root + 12}")
        except Exception:
            self._drag_win = None

    def _hide_drag_hint(self):
        try:
            if self._drag_win is not None:
                self._drag_win.destroy()
        except Exception:
            pass
        self._drag_win = None

    def _on_tree_press(self, event):
        # クリックで通常選択できるように、ここでは「ドラッグ候補」をセットするだけ。
        # ある程度マウスが動いた時だけドラッグ開始にする（閾値あり）。
        try:
            if self.tree.identify_region(event.x, event.y) == "heading":
                return

            iid = self.tree.identify_row(event.y)
            if not iid:
                # 空白クリック：候補解除
                self._drag_candidate_iid = None
                self._dragging = False
                return

            # 通常の選択は必ず行う（Remove等が効く）
            try:
                self.tree.selection_set(iid)
                self.tree.focus(iid)
            except Exception:
                pass

            self._drag_candidate_iid = iid
            self._drag_iid = iid
            self._dragging = False
            self._drag_started = False
            self._press_x_root = event.x_root
            self._press_y_root = event.y_root
            self._press_x = event.x
            self._press_y = event.y

            # つかむ感はカーソルで（押下時は手のまま、ドラッグ開始でfleurへ）
            try:
                self.tree.configure(cursor="hand2")
            except Exception:
                pass
        except Exception:
            self._drag_candidate_iid = None
            self._dragging = False
            self._drag_started = False

    def _on_tree_motion(self, event):
        # 押下中に一定距離動いたらドラッグ開始
        iid = getattr(self, "_drag_candidate_iid", None)
        if not iid:
            return

        # まだドラッグ開始していない場合は閾値判定
        if not getattr(self, "_drag_started", False):
            try:
                dx = abs(event.x_root - getattr(self, "_press_x_root", event.x_root))
                dy = abs(event.y_root - getattr(self, "_press_y_root", event.y_root))
            except Exception:
                dx = dy = 0
            if max(dx, dy) < 6:
                return

            # ドラッグ開始
            self._drag_started = True
            self._dragging = True

            # ソート表示は一旦解除（以降は手動順）
            self._name_sort_active = False
            self._update_name_heading()

            # visual "grab" feel
            try:
                self.tree.configure(cursor="fleur")
            except Exception:
                pass

            try:
                name = self.tree.set(iid, "name") or (self.tree.item(iid, "values") or [""])[0]
            except Exception:
                name = ""
            self._show_drag_hint(event.x_root, event.y_root, f"↕ {name}")

        if not getattr(self, "_dragging", False):
            return

        # hint follow
        try:
            if self._drag_win:
                self._show_drag_hint(event.x_root, event.y_root, self._drag_win.winfo_children()[0].cget("text"))
        except Exception:
            pass

        # move item in-tree (no placeholder rows)
        try:
            target = self.tree.identify_row(event.y)
            if target and target != iid:
                idx = self.tree.index(target)
                self.tree.move(iid, "", idx)
            elif not target:
                self.tree.move(iid, "", "end")
        except Exception:
            pass

    def _on_tree_release(self, _event):
        # クリックだけなら何もしない（選択が残る）
        try:
            started = getattr(self, "_drag_started", False)
            dragging = getattr(self, "_dragging", False)
        except Exception:
            started = dragging = False

        # 後片付け（候補は常に解除）
        self._drag_candidate_iid = None
        self._dragging = False
        self._drag_started = False

        if not started or not dragging:
            # 普通クリック：ヒントも出てないはずだが念のため
            self._hide_drag_hint()
            try:
                self.tree.configure(cursor="hand2")
            except Exception:
                pass
            return

        # ドラッグ確定：順番保存（リフレッシュで選択が消えるので、Treeはそのまま）
        self._hide_drag_hint()
        try:
            self.tree.configure(cursor="hand2")
        except Exception:
            pass

        try:
            self._sync_rules_from_tree_order()
            self._save_config()
            # ドラッグした行を選択状態のままに
            if getattr(self, "_drag_iid", None):
                try:
                    self.tree.selection_set(self._drag_iid)
                    self.tree.focus(self._drag_iid)
                except Exception:
                    pass
        except Exception:
            pass

    def _update_volume_label(self):
        try:
            v = int(self.var_volume.get())
        except Exception:
            v = 100
        v = max(0, min(100, v))
        try:
            self.lbl_volume.configure(text=f"{v}%")
        except Exception:
            pass

        # 再生中でも音量をリアルタイム反映（Test含む）
        # Only apply when the currently playing sound belongs to the selected rule.
        try:
            if not self._audio_ready:
                return
            if self.muted:
                return
            if not pygame.mixer.music.get_busy():
                return
            selected = self._get_selected_user_id()
            if selected and selected == self._now_playing_rule_id:
                pygame.mixer.music.set_volume(v / 100.0)
                self._now_playing_volume = v
        except Exception:
            pass


    def _get_selected_user_id(self) -> Optional[str]:
        sel = self.tree.selection()
        return sel[0] if sel else None

    def _load_selected_to_form(self, _evt=None):
        uid = self._get_selected_user_id()
        if not uid:
            return
        r = self._find_rule(uid)
        if not r:
            return
        self.entry_name.delete(0, "end")
        self.entry_name.insert(0, r.name)
        self.entry_id.delete(0, "end")
        self.entry_id.insert(0, r.user_id)
        self.entry_sound.delete(0, "end")
        self.entry_sound.insert(0, r.sound_path)
        try:
            self.var_pushover_sound.set(str(getattr(r, "pushover_sound", "") or ""))
        except Exception:
            self.var_pushover_sound.set("")
        try:
            self.var_volume.set(int(getattr(r, "volume", 100)))
        except Exception:
            self.var_volume.set(100)
        self._update_volume_label()

    def _find_rule(self, user_id: str) -> Optional[Rule]:
        for r in self.rules:
            if r.user_id == user_id:
                return r
        return None

    # ------------------------------------------------------------
    # UI actions
    # ------------------------------------------------------------
    def browse_sound(self):
        path = filedialog.askopenfilename(filetypes=[("Audio", "*.wav *.mp3")])
        if path:
            self.entry_sound.delete(0, "end")
            self.entry_sound.insert(0, path)

    def add_rule(self):
        name = self.entry_name.get().strip()
        user_id = str(self.entry_id.get()).strip()
        sound = self.entry_sound.get().strip()
        push_sound = (self.var_pushover_sound.get() or "").strip()

        if not user_id:
            messagebox.showinfo(APP_NAME, "UserID を入力してください。")
            return
        if self._find_rule(user_id) is not None:
            messagebox.showinfo(APP_NAME, "既に登録済みです。")
            return
        if not sound or not sound.lower().endswith(ALLOWED_AUDIO):
            messagebox.showinfo(APP_NAME, "Sound は .wav または .mp3 を指定してください。")
            return

        vol = max(0, min(100, int(self.var_volume.get() or 100)))
        self.rules.append(Rule(name=name, user_id=user_id, sound_path=sound, volume=vol, pushover_sound=push_sound))
        self._save_config()
        self._refresh_table()

        self.entry_id.delete(0, "end")
        self.entry_sound.delete(0, "end")
        try:
            self.var_pushover_sound.set("")
        except Exception:
            pass
        self.entry_id.focus_set()

    def update_rule(self):
        selected = self._get_selected_user_id()
        if not selected:
            messagebox.showinfo(APP_NAME, "更新する行を選択してください。")
            return

        name = self.entry_name.get().strip()
        user_id = str(self.entry_id.get()).strip()
        sound = self.entry_sound.get().strip()
        push_sound = (self.var_pushover_sound.get() or "").strip()

        if not user_id:
            messagebox.showinfo(APP_NAME, "UserID を入力してください。")
            return
        if not sound or not sound.lower().endswith(ALLOWED_AUDIO):
            messagebox.showinfo(APP_NAME, "Sound は .wav または .mp3 を指定してください。")
            return
        if user_id != selected and self._find_rule(user_id) is not None:
            messagebox.showinfo(APP_NAME, "既に登録済みです。")
            return

        r = self._find_rule(selected)
        if not r:
            return

        old_id = r.user_id
        r.name = name
        r.user_id = user_id
        r.sound_path = sound
        try:
            r.pushover_sound = push_sound
        except Exception:
            pass
        try:
            r.volume = max(0, min(100, int(self.var_volume.get() or 100)))
        except Exception:
            r.volume = 100

        if old_id != user_id:
            self.rules = [x for x in self.rules if (x is r) or (x.user_id != old_id)]

        self._save_config()
        self._refresh_table()
        try:
            self.tree.selection_set(user_id)
            self.tree.focus(user_id)
        except Exception:
            pass

    def remove_rule(self):
        selected = self._get_selected_user_id()
        if not selected:
            messagebox.showinfo(APP_NAME, "削除する行を選択してください。")
            return
        self.rules = [r for r in self.rules if r.user_id != selected]
        self._save_config()
        self._refresh_table()

        self.entry_id.delete(0, "end")
        self.entry_sound.delete(0, "end")
        try:
            self.var_pushover_sound.set("")
        except Exception:
            pass

    def test_form(self):
        """現在フォームに入力されている内容を再生してテストする（Update前でもOK）。"""
        sound = self.entry_sound.get().strip()
        if not sound or not sound.lower().endswith(ALLOWED_AUDIO):
            messagebox.showinfo(APP_NAME, "Sound は .wav または .mp3 を指定してください。")
            return
        try:
            vol = max(0, min(100, int(self.var_volume.get() or 100)))
        except Exception:
            vol = 100

        # selectionがあればそのルールID、なければテスト用IDで再生
        rid = self._get_selected_user_id() or "__test__"
        self._play_sound(sound, vol, rule_id=rid)

    # 互換：古いUIから呼ばれても動くように残す
    def test_selected(self):
        return self.test_form()

    def toggle_mute(self):
        self.muted = not self.muted
        # Stop current playback immediately when mute is toggled
        try:
            self._stop_sound()
        except Exception:
            pass
        self._save_config()
        self._refresh_table()


    # ------------------------------------------------------------
    # Audio
    # ------------------------------------------------------------
    def _init_audio(self):
        if not _HAS_PYGAME:
            self._audio_ready = False
            return
        try:
            pygame.mixer.init()
            self._audio_ready = True
        except Exception:
            self._audio_ready = False

    def _play_sound(self, path: str, volume: int = 100, rule_id: Optional[str] = None):
        if self.muted:
            return
        if not self._audio_ready:
            messagebox.showerror(
                APP_NAME,
                "音声再生に必要な pygame が利用できません。\n"
                "Python 3.14 の場合は pygame-ce を推奨: pip install pygame-ce"
            )
            return
        try:
            pygame.mixer.music.stop()
            pygame.mixer.music.load(path)
            try:
                v = max(0, min(100, int(volume)))
            except Exception:
                v = 100
            pygame.mixer.music.set_volume(v / 100.0)
            self._now_playing_rule_id = rule_id
            self._now_playing_volume = v
            pygame.mixer.music.play()
        except Exception as e:
            messagebox.showerror(APP_NAME, f"音声ファイルを再生できません。\n{e}")

    
    def _stop_sound(self):
        if not self._audio_ready:
            return
        try:
            pygame.mixer.music.stop()
        except Exception:
            pass
        self._now_playing_rule_id = None

# ------------------------------------------------------------
    # Bot (auto)
    # ------------------------------------------------------------
    def _auto_start_bot(self):
        if not _HAS_DISCORD:
            self._set_bot_state("offline", "Bot: discord.py が未インストール")
            return
        if not self.bot_token:
            self._set_bot_state("offline", "Bot: TOKEN未設定（⚙で設定）")
            return

        self._start_bot_async()

    def _start_bot_async(self):
        with self._bot_thread_lock:
            if self._bot_thread and self._bot_thread.is_alive():
                return
            self._set_bot_state("connecting", "Bot: connecting...")
            self._bot_thread = threading.Thread(target=self._run_bot_thread, daemon=True)
            self._bot_thread.start()

    def _stop_bot_async(self):
        threading.Thread(target=self._stop_bot, daemon=True).start()

    def _restart_bot_async(self):
        def worker():
            self._stop_bot()
            time.sleep(0.4)
            self._auto_start_bot()
        threading.Thread(target=worker, daemon=True).start()

    def _stop_bot(self):
        try:
            if self._discord_client and self._discord_loop and self._discord_loop.is_running():
                fut = asyncio.run_coroutine_threadsafe(self._discord_client.close(), self._discord_loop)
                try:
                    fut.result(timeout=2.5)
                except Exception:
                    pass
        except Exception:
            pass

        t = self._bot_thread
        if t and t.is_alive():
            try:
                t.join(timeout=2.5)
            except Exception:
                pass

        self._discord_client = None
        self._discord_loop = None
        self._bot_thread = None

    def _run_bot_thread(self):
        try:
            intents = discord.Intents.default()
            intents.message_content = True
            intents.messages = True
            intents.guilds = True

            client = discord.Client(intents=intents)
            self._discord_client = client

            @client.event
            async def on_ready():
                self._ui_call(lambda: self._set_bot_state("online", f"Bot: online ({client.user})"))

            @client.event
            async def on_disconnect():
                self._ui_call(lambda: self._set_bot_state("offline", "Bot: disconnected"))

            @client.event
            async def on_resumed():
                self._ui_call(lambda: self._set_bot_state("online", f"Bot: online ({client.user})"))

            @client.event
            async def on_message(message):
                try:
                    if message.author.bot:
                        return
                    uid = str(message.author.id)
                    r = self._find_rule(uid)
                    if r is None:
                        return
                    self._ui_call(lambda: self._play_sound(r.sound_path, int(getattr(r, 'volume', 100)), rule_id=r.user_id))
                    # Pushover push (optional)
                    try:
                        if self.pushover_enabled and self.pushover_app_token and self.pushover_user_key:
                            if (not self.muted) or self.pushover_push_when_muted:
                                who = (r.name or getattr(message.author, 'display_name', '') or 'User')
                                where = "DM"
                                try:
                                    if message.guild and message.channel:
                                        where = f"{message.guild.name} / #{message.channel.name}"
                                except Exception:
                                    pass
                                if self.pushover_include_message:
                                    text = (getattr(message, 'clean_content', '') or getattr(message, 'content', '') or '').strip()
                                    if not text:
                                        text = "(textなし)"
                                    msg = f"{who} @ {where}: {text}"
                                else:
                                    msg = f"{who} @ {where}"
                                jump = getattr(message, 'jump_url', None)
                                await self._pushover_send_async(APP_NAME, msg, url=jump, sound=getattr(r, 'pushover_sound', ''))
                    except Exception:
                        pass
                except Exception:
                    pass

            self._discord_loop = asyncio.new_event_loop()
            asyncio.set_event_loop(self._discord_loop)
            client.run(self.bot_token)
        except Exception as e:
            self._ui_call(lambda: self._set_bot_state("offline", f"Bot: start failed ({e})"))

    def _ui_call(self, fn):
        try:
            self.after(0, fn)
        except Exception:
            pass


    # ------------------------------------------------------------
    # Pushover (iOS push)
    # ------------------------------------------------------------
    def _pushover_request_sync(
        self,
        app_token: str,
        user_key: str,
        title: str,
        message: str,
        url: Optional[str] = None,
        url_title: str = "Open in Discord",
        sound: str = "",
    ) -> tuple[bool, str]:
        """Send a Pushover notification (blocking). Returns (ok, error_message)."""
        app_token = (app_token or "").strip()
        user_key = (user_key or "").strip()
        if not (app_token and user_key):
            return False, "PushoverのApp Token / User Key が未設定です。"

        params = {
            "token": app_token,
            "user": user_key,
            "title": title,
            "message": message,
        }
        if url:
            params["url"] = url
            params["url_title"] = url_title

        # Optional sound (Pushover built-in sound name)
        if sound:
            params["sound"] = sound

        data = urllib.parse.urlencode(params).encode("utf-8")
        req = urllib.request.Request(PUSHOVER_API_URL, data=data, method="POST")
        try:
            with urllib.request.urlopen(req, timeout=10) as resp:
                body = resp.read().decode("utf-8", "replace")
                if getattr(resp, "status", 200) != 200:
                    return False, f"HTTP {getattr(resp, 'status', '?')}: {body}"
        except Exception as e:
            return False, str(e)

        try:
            j = json.loads(body)
            if int(j.get("status", 0)) == 1:
                return True, ""
            errs = j.get("errors") or j.get("error") or body
            return False, str(errs)
        except Exception:
            # if response isn't json
            return True, ""

    def _pushover_send_sync(
        self,
        title: str,
        message: str,
        url: Optional[str] = None,
        url_title: str = "Open in Discord",
        sound: str = "",
    ) -> tuple[bool, str]:
        return self._pushover_request_sync(
            self.pushover_app_token,
            self.pushover_user_key,
            title,
            message,
            url=url,
            url_title=url_title,
            sound=sound,
        )

    async def _pushover_send_async(
        self,
        title: str,
        message: str,
        url: Optional[str] = None,
        url_title: str = "Open in Discord",
        sound: str = "",
    ) -> tuple[bool, str]:
        return await asyncio.to_thread(self._pushover_send_sync, title, message, url, url_title, sound)


# ------------------------------------------------------------
    # Bot status indicator (blink dot only; no layout shift)
    # ------------------------------------------------------------
    def _set_dot_color(self, color: str):
        try:
            self.dot_canvas.itemconfigure(self._dot_id, fill=color)
        except Exception:
            pass

    def _set_bot_state(self, state: str, text: str):
        self._bot_state = state
        self.status_var.set(text)

        if state == "online":
            self._set_dot_color("#2ecc71")
        elif state == "connecting":
            self._set_dot_color("#2ecc71")
        else:
            self._set_dot_color("#e74c3c")

    def _tick_blink(self):
        try:
            if self._bot_state in ("connecting", "offline"):
                self._blink_on = not self._blink_on
                if self._bot_state == "connecting":
                    self._set_dot_color("#2ecc71" if self._blink_on else self.style.colors.bg)
                else:
                    self._set_dot_color("#e74c3c" if self._blink_on else self.style.colors.bg)
            else:
                self._set_dot_color("#2ecc71")
        except Exception:
            pass

        aid = self.after(450, self._tick_blink)
        self._after_ids.add(aid)

    def _cancel_afters(self):
        for aid in list(self._after_ids):
            try:
                self.after_cancel(aid)
            except Exception:
                pass
        self._after_ids.clear()
    def _confirm_exit(self):
        # Single confirmation dialog is handled in on_close().
        self.on_close()



    # ------------------------------------------------------------
    # Close / cleanup
    # ------------------------------------------------------------
    def _cleanup(self):
        self.stop_event.set()
        self._cancel_afters()
        self._save_config()

        # tray icon
        try:
            self._stop_tray()
        except Exception:
            pass

        try:
            if self._audio_ready:
                pygame.mixer.music.stop()
                pygame.mixer.quit()
        except Exception:
            pass

        try:
            self._stop_bot()
        except Exception:
            pass

    def on_close(self):
        # Click [X] -> confirm to stop monitoring and exit
        was_in_tray = bool(getattr(self, "_in_tray", False))

        # Ensure dialog appears
        try:
            if was_in_tray:
                self._show_from_tray()
            else:
                if self.state() == "iconic":
                    self.deiconify()
                    self.state("normal")
                self.lift()
                self.focus_force()
        except Exception:
            pass

        if not messagebox.askyesno(APP_NAME, "監視を終了してアプリを終了しますか？"):
            # If it was in tray, return to tray
            if was_in_tray:
                try:
                    self.after(0, self._hide_to_tray)
                except Exception:
                    pass
            return

        try:
            self._cleanup()
        finally:
            try:
                self.quit()
            except Exception:
                pass
            try:
                self.destroy()
            except Exception:
                pass
            if FORCE_EXIT:
                os._exit(0)



if __name__ == "__main__":
    TalkAlertApp().mainloop()
