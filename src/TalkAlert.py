"""
TalkAlert (Distribution build - Token via Settings ⚙)
- White UI, thin/clean fonts
- Bot auto-runs on launch IF token is set
- Token is NOT embedded in code; set it via ⚙ Settings dialog (saved to %APPDATA%\TalkAlert\config.json)
- Rules: Name (memo), UserID, Sound (wav/mp3)  ※Table shows filename only; Sound column is right-aligned
- Prevent duplicate UserID registration ("既に登録済みです")
- Bot status indicator: fixed-size dot (green/red) + blinking without layout shift + status text
- Note about Discord's own notification sound is kept in UI
- Close -> ensures process exits (no lingering process)

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

        self.tree = ttk.Treeview(body, columns=("name", "user_id", "sound", "vol"), show="headings", height=8)
        self.tree.heading("name", text="Name", anchor=tk.W)
        self.tree.heading("user_id", text="UserID", anchor=tk.W)
        self.tree.heading("sound", text="Sound", anchor=tk.W)
        self.tree.heading("vol", text="Vol", anchor=tk.W)


        self.tree.column("name", width=170, anchor=tk.W)
        self.tree.column("user_id", width=190, anchor=tk.W)
        self.tree.column("sound", width=190, anchor=tk.W)
        self.tree.column("vol", width=40, anchor=tk.W)

        vsb = ttk.Scrollbar(body, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns", padx=(8, 0))
        body.rowconfigure(0, weight=1)
        body.columnconfigure(0, weight=1)

        self.tree.bind("<<TreeviewSelect>>", self._load_selected_to_form)

        # Form
        form = tb.Frame(root, padding=(12, 10))
        form.pack(fill=X, pady=(8, 0))

        tb.Label(form, text="Name (任意)", bootstyle="secondary").grid(row=0, column=0, sticky=W)
        tb.Label(form, text="UserID", bootstyle="secondary").grid(row=0, column=1, sticky=W)
        tb.Label(form, text="Sound (.wav/.mp3)", bootstyle="secondary").grid(row=0, column=2, sticky=W)

        self.entry_name = tb.Entry(form)
        self.entry_id = tb.Entry(form)
        self.entry_sound = tb.Entry(form)

        self.entry_name.grid(row=1, column=0, sticky=EW, padx=(0, 10), pady=(4, 0))
        self.entry_id.grid(row=1, column=1, sticky=EW, padx=(0, 10), pady=(4, 0))
        self.entry_sound.grid(row=1, column=2, sticky=EW, padx=(0, 10), pady=(4, 0))

        btn_browse = tb.Button(form, text="Browse", command=self.browse_sound, bootstyle="secondary", width=10)
        btn_browse.grid(row=1, column=3, sticky=E, pady=(4, 0))

        # Volume (per sound)
        self.var_volume = tk.IntVar(value=100)
        tb.Label(form, text="Volume", bootstyle="secondary").grid(row=2, column=0, sticky=W, pady=(10, 0))
        self.scale_volume = tb.Scale(form, from_=0, to=100, orient=HORIZONTAL, variable=self.var_volume, command=lambda _v: self._update_volume_label())
        self.scale_volume.grid(row=2, column=1, columnspan=2, sticky=EW, pady=(10, 0), padx=(0, 10))
        self.lbl_volume = tb.Label(form, text="100%", bootstyle="secondary", width=5, anchor=E)
        self.lbl_volume.grid(row=2, column=3, sticky=E, pady=(10, 0))
        self._update_volume_label()

        form.columnconfigure(0, weight=1)
        form.columnconfigure(1, weight=1)
        form.columnconfigure(2, weight=2)

        # Actions
        actions = tb.Frame(root)
        actions.pack(fill=X, pady=(10, 0))

        tb.Button(actions, text="＋ Add", command=self.add_rule, bootstyle="success", width=12).pack(side=LEFT)
        tb.Button(actions, text="⟳ Update", command=self.update_rule, bootstyle="primary", width=12).pack(side=LEFT, padx=(10, 0))
        tb.Button(actions, text="🗑 Remove", command=self.remove_rule, bootstyle="danger", width=12).pack(side=LEFT, padx=(10, 0))
        tb.Button(actions, text="▶ Test", command=self.test_selected, bootstyle="secondary", width=12).pack(side=RIGHT)

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

        status = tk.StringVar(value="")
        tb.Label(frame, textvariable=status, bootstyle="secondary").pack(anchor="w", pady=(8, 0))

        btns = tb.Frame(frame)
        btns.pack(fill=X, pady=(8, 0))

        def do_save():
            tok = token_var.get().strip()
            self.tray_on_minimize = bool(tray_var.get())

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

            self.rules = []
            for r in data.get("rules", []):
                user_id = str(r.get("user_id", "")).strip()
                if not user_id:
                    continue
                self.rules.append(
                    Rule(
                        name=str(r.get("name", "")).strip(),
                        user_id=user_id,
                        sound_path=str(r.get("sound_path", "")).strip(),
                        volume=max(0, min(100, int(r.get("volume", 100) or 100))),
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
                "rules": [{"name": r.name, "user_id": r.user_id, "sound_path": r.sound_path, "volume": int(getattr(r, "volume", 100))} for r in self.rules],
            }
            CONFIG_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass

    # ------------------------------------------------------------
    # Table helpers
    # ------------------------------------------------------------
    def _refresh_table(self):
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        for r in self.rules:
            self.tree.insert("", "end", iid=r.user_id, values=(r.name, r.user_id, r.sound_filename, f"{int(getattr(r, 'volume', 100))}%"))

        if self.muted:
            self.btn_mute.configure(text="🔇 Mute: ON", bootstyle="danger")
        else:
            self.btn_mute.configure(text="🔇 Mute: OFF", bootstyle="warning")

    
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
        self.rules.append(Rule(name=name, user_id=user_id, sound_path=sound, volume=vol))
        self._save_config()
        self._refresh_table()

        self.entry_id.delete(0, "end")
        self.entry_sound.delete(0, "end")
        self.entry_id.focus_set()

    def update_rule(self):
        selected = self._get_selected_user_id()
        if not selected:
            messagebox.showinfo(APP_NAME, "更新する行を選択してください。")
            return

        name = self.entry_name.get().strip()
        user_id = str(self.entry_id.get()).strip()
        sound = self.entry_sound.get().strip()

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

    def test_selected(self):
        selected = self._get_selected_user_id()
        if not selected:
            messagebox.showinfo(APP_NAME, "テストする行を選択してください。")
            return
        r = self._find_rule(selected)
        if not r:
            return
        self._play_sound(r.sound_path, int(self.var_volume.get() or 100), rule_id=r.user_id)

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
