# timelyads_pro_designer.py
# TimelyAds Pro — Designer edition (visual upgrade + funcionalidades mantidas)
# - Um cadeado global com PIN 4510
# - Editor de horários por mídia
# - Mic com ducking (sounddevice + pycaw opcionais)
# - Export/Import, scheduler, play agora, criar/editar playlists
#
# Instale dependências quando necessário:
# pip install pygame sounddevice numpy pycaw comtypes Pillow

import os
import json
import time
import shutil
import threading
import traceback
import pathlib
from datetime import datetime

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog, Menu

# audio playback
import pygame
from pygame import mixer

# optional: pycaw for Windows session volume control
PYCAW_OK = True
try:
    from comtypes import CLSCTX_ALL
    from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
except Exception:
    PYCAW_OK = False

# optional: sounddevice + numpy for low-latency mic passthrough
SOUND_OK = True
try:
    import sounddevice as sd
    import numpy as np
except Exception:
    SOUND_OK = False

# ------------------------ Settings / Paths ------------------------
BASE_DIR = pathlib.Path(__file__).parent.resolve()
PLAYLISTS_JSON = str(BASE_DIR / "playlists.json")
CONFIG_JSON = str(BASE_DIR / "timelyads_config.json")

APP_TITLE = "TimelyAds Pro — Designer"
SCHEDULE_PIN = "4510"

# Color palette (refined)
BG = "#071016"
PANEL = "#0e1720"
CARD = "#0f1a23"
ACCENT = "#00d6a6"     # neon-green accent
ACCENT_DIM = "#00b788"
TEXT = "#cfeee2"
MUTED = "#7ea99a"
SEPARATOR = "#12232b"

# Fonts (fall back to common fonts)
DEFAULT_FONT = ("Segoe UI", 10)
TITLE_FONT = ("Segoe UI Semibold", 18)
HEADER_FONT = ("Segoe UI", 12, "bold")
H2_FONT = ("Segoe UI", 11, "bold")
TREE_HEADING_FONT = ("Segoe UI", 11, "bold")
TREE_FONT = ("Segoe UI", 10)

# ------------------------ Utility Helpers ------------------------
def safe_load_json(path, default):
    try:
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        return default
    return default

def safe_save_json(path, data):
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
            return True
    except Exception as e:
        print("save json error:", e)
        return False

# ------------------------ Main App ------------------------
class TimelyAdsApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1220x740")
        self.minsize(960, 600)
        self.configure(bg=BG)

        # state
        self.playlist_file = PLAYLISTS_JSON
        self.playlists = safe_load_json(self.playlist_file, {})
        self.current_playlist = None

        # playback / audio state
        self._playback_thread = None
        self._playback_lock = threading.Lock()
        self._is_playing = False

        # pycaw duck state
        self._saved_sessions = {}
        self._duck_active = False

        # mic state
        self._mic_active = False
        self._mic_stream = None
        self._mic_gain = 1.0
        self._mic_input_device = None
        self._mic_output_device = None

        # global lock (True = locked)
        self._global_locked = True

        # load saved config
        self._load_config()

        # init audio mixer
        self._init_mixer()

        # build UI
        self._setup_styles()
        self._build_layout()
        self._after_ui_setup()

        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------ Config / Persistence ------------------------
    def _load_config(self):
        cfg = safe_load_json(CONFIG_JSON, {})
        self._mic_input_device = cfg.get("mic_input_device")
        self._mic_output_device = cfg.get("mic_output_device")
        self._global_locked = bool(cfg.get("global_locked", True))

    def _save_config(self):
        cfg = {
            "mic_input_device": self._mic_input_device,
            "mic_output_device": self._mic_output_device,
            "global_locked": self._global_locked
        }
        safe_save_json(CONFIG_JSON, cfg)

    def _load_playlists(self):
        self.playlists = safe_load_json(self.playlist_file, {})

    def _save_playlists(self):
        safe_save_json(self.playlist_file, self.playlists)

    # ------------------------ Audio init ------------------------
    def _init_mixer(self):
        try:
            pygame.mixer.pre_init(44100, -16, 2, 512)
            pygame.init()
            mixer.init()
        except Exception as e:
            messagebox.showwarning("Áudio", f"Erro iniciando mixer: {e}")

    # ------------------------ Styles ------------------------
    def _setup_styles(self):
        style = ttk.Style(self)
        # base theme
        try:
            style.theme_use("clam")
        except Exception:
            pass

        # General
        style.configure("App.TFrame", background=BG)
        style.configure("Card.TFrame", background=CARD, relief="flat")
        style.configure("Accent.TLabel", background=BG, foreground=ACCENT, font=HEADER_FONT)
        style.configure("Title.TLabel", background=BG, foreground=ACCENT, font=TITLE_FONT)
        style.configure("Muted.TLabel", background=BG, foreground=MUTED, font=DEFAULT_FONT)

        # Buttons
        style.configure("Neon.TButton",
                        background=BG, foreground=ACCENT,
                        borderwidth=1, focusthickness=0, focuscolor="",
                        padding=(10, 8), font=DEFAULT_FONT)
        style.map("Neon.TButton",
                  background=[("active", PANEL), ("!active", BG)],
                  foreground=[("active", ACCENT_DIM), ("!active", ACCENT)])

        style.configure("Primary.TButton",
                        background=ACCENT, foreground=BG,
                        padding=(10, 8), font=("Segoe UI", 10, "bold"))
        style.map("Primary.TButton", background=[("active", ACCENT_DIM)])

        # Treeview
        style.configure("Treeview",
                        background=CARD, fieldbackground=CARD, foreground=TEXT,
                        rowheight=34, font=TREE_FONT)
        style.configure("Treeview.Heading",
                        background=PANEL, foreground=ACCENT, font=TREE_HEADING_FONT)
        style.map("Treeview", background=[("selected", PANEL)])

        # Scales
        style.configure("TScale", background=BG)

    # ------------------------ UI Layout ------------------------
    def _build_layout(self):
        # Grid config
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(1, weight=1)

        # Header
        header = ttk.Frame(self, style="App.TFrame")
        header.grid(row=0, column=0, columnspan=3, sticky="ew", padx=16, pady=(12,8))
        header.columnconfigure(0, weight=1)

        # Title area
        title = ttk.Label(header, text=APP_TITLE, style="Title.TLabel")
        title.grid(row=0, column=0, sticky="w")

        # Right side header controls: lock, gear, clock
        ctrl_frame = ttk.Frame(header, style="App.TFrame")
        ctrl_frame.grid(row=0, column=1, sticky="e")
        self._global_lock_btn = ttk.Button(ctrl_frame, text="", width=10,
                                           style="Neon.TButton", command=self._toggle_global_lock)
        self._update_global_lock_btn()
        self._global_lock_btn.grid(row=0, column=0, padx=(0,8))

        self._gear_btn = ttk.Button(ctrl_frame, text="⚙", width=4, style="Neon.TButton", command=self._open_main_menu)
        self._gear_btn.grid(row=0, column=1, padx=(0,8))

        self._clock_label = ttk.Label(ctrl_frame, text="--:--", style="Accent.TLabel")
        self._clock_label.grid(row=0, column=2, padx=(0,2))

        # Main panels: Left playlist, Center medias, Right rules
        # Left - Playlists
        left = ttk.Frame(self, style="App.TFrame")
        left.grid(row=1, column=0, sticky="nsw", padx=(16,8), pady=12)
        left.columnconfigure(0, weight=1)
        ttk.Label(left, text="Playlists", style="Accent.TLabel").grid(row=0, column=0, sticky="w", pady=(0,6))
        left_card = ttk.Frame(left, style="Card.TFrame", padding=10)
        left_card.grid(row=1, column=0, sticky="nsew")
        left_card.grid_propagate(False)
        left_card.config(width=220, height=520)

        self.playlist_list = tk.Listbox(left_card, bg=CARD, fg=TEXT, bd=0, highlightthickness=0,
                                       selectbackground=PANEL, font=DEFAULT_FONT, activestyle="none", width=28)
        self.playlist_list.pack(fill="both", expand=True)
        self.playlist_list.bind("<<ListboxSelect>>", self._on_select_playlist)

        ttk.Separator(left, orient="horizontal").grid(row=2, column=0, sticky="ew", pady=(8,8))
        self._toggle_playlist_btn = ttk.Button(left, text="Alternar ON/OFF", style="Neon.TButton", command=self._toggle_current_playlist)
        self._toggle_playlist_btn.grid(row=3, column=0, sticky="ew", pady=(4,0))

        # Center - Media list
        center = ttk.Frame(self, style="App.TFrame")
        center.grid(row=1, column=1, sticky="nsew", padx=(8,8), pady=12)
        center.columnconfigure(0, weight=1)
        ttk.Label(center, text="Mídias", style="Accent.TLabel").grid(row=0, column=0, sticky="w", pady=(0,8))
        center_card = ttk.Frame(center, style="Card.TFrame")
        center_card.grid(row=1, column=0, sticky="nsew")
        center_card.rowconfigure(0, weight=1)
        center_card.columnconfigure(0, weight=1)

        cols = ("name", "time", "repeat", "play")
        self.tree = ttk.Treeview(center_card, columns=cols, show="headings", selectmode="browse", style="Treeview")
        self.tree.heading("name", text="Nome")
        self.tree.heading("time", text="Horários")
        self.tree.heading("repeat", text="Repetir")
        self.tree.heading("play", text="▶")
        # column sizes
        self.tree.column("name", anchor="w", width=640)
        self.tree.column("time", anchor="center", width=140)
        self.tree.column("repeat", anchor="center", width=90)
        self.tree.column("play", anchor="center", width=50)
        self.tree.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)

        # scrollbar
        vs = ttk.Scrollbar(center_card, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vs.set)
        vs.grid(row=0, column=1, sticky="ns")

        # Bindings: click on time column opens schedule editor; double-click opens editor
        self.tree.bind("<Button-1>", self._on_tree_click)
        self.tree.bind("<Double-1>", self._on_tree_double_click)
        self.tree.bind("<Button-3>", self._on_tree_right_click)

        # Right - Rules
        right = ttk.Frame(self, style="App.TFrame")
        right.grid(row=1, column=2, sticky="ns", padx=(8,16), pady=12)
        ttk.Label(right, text="Regras de repetição", style="Accent.TLabel").grid(row=0, column=0, sticky="w", pady=(0,6))
        right_card = ttk.Frame(right, style="Card.TFrame", padding=12)
        right_card.grid(row=1, column=0, sticky="nsew")

        ttk.Label(right_card, text="Repetir", style="Muted.TLabel").grid(row=0, column=0, sticky="w")
        self.repeat_cb = ttk.Combobox(right_card, values=[f"{i} vez" if i==1 else f"{i} vezes" for i in range(1,11)], state="readonly", width=12)
        self.repeat_cb.current(2)
        self.repeat_cb.grid(row=0, column=1, sticky="e")

        ttk.Label(right_card, text="Distribuição ao longo do dia", style="Muted.TLabel").grid(row=1, column=0, columnspan=2, sticky="w", pady=(12,4))
        self.distrib_scale = ttk.Scale(right_card, from_=6, to=18, orient="horizontal")
        self.distrib_scale.set(12)
        self.distrib_scale.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(0,4))
        ttk.Label(right_card, text="06:00", style="Muted.TLabel").grid(row=3, column=0, sticky="w")
        ttk.Label(right_card, text="18:00", style="Muted.TLabel").grid(row=3, column=1, sticky="e")

        # Footer actions
        footer = ttk.Frame(self, style="App.TFrame")
        footer.grid(row=2, column=0, columnspan=3, sticky="ew", padx=16, pady=(8,16))
        for i in range(6):
            footer.columnconfigure(i, weight=1)

        self.btn_create = ttk.Button(footer, text="➕ Criar Playlist", style="Neon.TButton", command=self._create_playlist)
        self.btn_create.grid(row=0, column=0, sticky="ew", padx=6, pady=6)
        self.btn_add = ttk.Button(footer, text="＋ Adicionar Mídia", style="Neon.TButton", command=self._add_media)
        self.btn_add.grid(row=0, column=1, sticky="ew", padx=6, pady=6)
        self.btn_play = ttk.Button(footer, text="▶ Tocar Agora", style="Primary.TButton", command=self._play_selected_media)
        self.btn_play.grid(row=0, column=2, sticky="ew", padx=6, pady=6)
        self.btn_gen = ttk.Button(footer, text="⚡ Gerar Novo", style="Neon.TButton", command=self._generate_schedule)
        self.btn_gen.grid(row=0, column=3, sticky="ew", padx=6, pady=6)

        # Mic controls compact
        mic_frame = ttk.Frame(footer, style="App.TFrame")
        mic_frame.grid(row=0, column=4, sticky="ew", padx=6, pady=6)
        mic_frame.columnconfigure(0, weight=1)
        self._mic_btn = ttk.Button(mic_frame, text="🎙 Mic", style="Neon.TButton", command=self._toggle_mic)
        self._mic_btn.grid(row=0, column=0, sticky="ew")
        # mic volume inline
        self._mic_vol = tk.DoubleVar(value=100.0)
        self._mic_vol_scale = ttk.Scale(footer, from_=0, to=200, variable=self._mic_vol, orient="horizontal", command=self._on_mic_vol_change)
        self._mic_vol_scale.grid(row=0, column=5, sticky="ew", padx=6, pady=6)
        self._mic_label = ttk.Label(footer, text="Mic: off", background=BG, foreground=MUTED)
        self._mic_label.grid(row=1, column=5, sticky="e", padx=6)

        # Context menu for tree
        self._ctx_menu = Menu(self, tearoff=0)
        self._ctx_menu.add_command(label="Definir horário...", command=self._ctx_set_time)
        self._ctx_menu.add_command(label="Definir repetições...", command=self._ctx_set_repeat)
        self._ctx_menu.add_separator()
        self._ctx_menu.add_command(label="Remover", command=self._ctx_remove_item)

    # After UI: load data and start ticks
    def _after_ui_setup(self):
        # ensure consistent data structure, migrate old fields
        self._load_playlists()
        for pl in self.playlists.values():
            for m in pl.get("files", []):
                if isinstance(m, dict):
                    if "times" not in m:
                        if "time" in m:
                            m["times"] = [m.get("time")]
                            m.pop("time", None)
                        else:
                            m["times"] = []
                    if "repeats" not in m:
                        m["repeats"] = 1

        if not self.playlists:
            # create demo playlists
            self.playlists = {
                "FM": {"files": [], "time": "00:00", "repeats": 1, "active": True},
                "Dia a Dia": {"files": [], "time": "00:00", "repeats": 1, "active": True},
            }
        self._refresh_playlist_list()
        self._refresh_media_table()
        # clock and schedule ticks
        self._clock_tick()
        self._schedule_tick()

    # ------------------------ UI helpers ------------------------
    def _update_global_lock_btn(self):
        if self._global_locked:
            self._global_lock_btn.config(text="🔒 LOCK", style="Neon.TButton")
        else:
            self._global_lock_btn.config(text="🔓 UNLOCK", style="Neon.TButton")

    def _clock_tick(self):
        self._clock_label.config(text=time.strftime("%H:%M"))
        self.after(1000, self._clock_tick)

    # ------------------------ Playlist helpers ------------------------
    def _refresh_playlist_list(self):
        self.playlist_list.delete(0, tk.END)
        for name, data in self.playlists.items():
            tag = "ON" if data.get("active", True) else "OFF"
            self.playlist_list.insert(tk.END, f"{name}   [{tag}]")
        names = list(self.playlists.keys())
        if not names:
            self.current_playlist = None
            return
        if self.current_playlist not in names:
            self.current_playlist = names[0]
        idx = names.index(self.current_playlist)
        self.playlist_list.selection_clear(0, tk.END)
        self.playlist_list.selection_set(idx)
        self.playlist_list.see(idx)

    def _on_select_playlist(self, _evt=None):
        sel = self.playlist_list.curselection()
        if not sel:
            return
        self.current_playlist = list(self.playlists.keys())[sel[0]]
        self._refresh_media_table()

    def _toggle_current_playlist(self):
        if not self.current_playlist:
            messagebox.showwarning("Aviso", "Nenhuma playlist selecionada.")
            return
        self.playlists[self.current_playlist]["active"] = not self.playlists[self.current_playlist].get("active", True)
        self._refresh_playlist_list()
        self._save_playlists()

    # ------------------------ Media table ------------------------
    def _refresh_media_table(self):
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        if not self.current_playlist:
            return
        items = self.playlists[self.current_playlist]["files"]
        for it in items:
            if isinstance(it, dict):
                name = os.path.basename(it.get("path", ""))
                times = it.get("times", [])
                time_label = times[0] if len(times) == 1 else ("Múltiplos" if times else "—")
                repeats = it.get("repeats", 1)
            else:
                name = os.path.basename(it)
                time_label = "—"
                repeats = 1
            self.tree.insert("", "end", values=(name, time_label, repeats, "▶"))

    def _get_selected_media_index(self):
        sel = self.tree.selection()
        if not sel:
            return None
        return self.tree.index(sel[0])

    # ------------------------ Tree events ------------------------
    def _on_tree_right_click(self, event):
        row = self.tree.identify_row(event.y)
        if row:
            self.tree.selection_set(row)
            self._ctx_menu.tk_popup(event.x_root, event.y_root)

    def _on_tree_click(self, event):
        # if click in time column -> open editor (with PIN if locked)
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        col = self.tree.identify_column(event.x)  # '#1' '#2' ...
        row = self.tree.identify_row(event.y)
        if not row:
            return
        idx = list(self.tree.get_children()).index(row)
        if col == "#2":  # time column
            # if global locked -> request PIN first
            if self._global_locked:
                pin = simpledialog.askstring("Senha", "Digite a senha para editar (PIN):", show="*")
                if pin != SCHEDULE_PIN:
                    messagebox.showerror("Erro", "Senha incorreta.")
                    return
                self._global_locked = False
                self._update_global_lock_btn()
                self._save_config()
            self._open_schedule_editor(idx)

    def _on_tree_double_click(self, event):
        row = self.tree.identify_row(event.y)
        if not row:
            return
        idx = list(self.tree.get_children()).index(row)
        if self._global_locked:
            pin = simpledialog.askstring("Senha", "Digite a senha para editar (PIN):", show="*")
            if pin != SCHEDULE_PIN:
                messagebox.showerror("Erro", "Senha incorreta.")
                return
            self._global_locked = False
            self._update_global_lock_btn()
            self._save_config()
        self._open_schedule_editor(idx)

    # ------------------------ Buttons / Actions ------------------------
    def _create_playlist(self):
        name = simpledialog.askstring("Nova Playlist", "Nome da playlist:")
        if not name:
            return
        if name in self.playlists:
            messagebox.showwarning("Aviso", "Playlist já existe.")
            return
        self.playlists[name] = {"files": [], "time": "00:00", "repeats": 1, "active": True}
        self.current_playlist = name
        self._refresh_playlist_list()
        self._refresh_media_table()
        self._save_playlists()

    def _add_media(self):
        if not self.current_playlist:
            messagebox.showwarning("Aviso", "Selecione uma playlist primeiro.")
            return
        files = filedialog.askopenfilenames(title="Selecione arquivos de áudio", filetypes=[("Áudio", "*.mp3 *.wav *.ogg *.flac"), ("Todos","*.*")])
        if not files:
            return
        for p in files:
            self.playlists[self.current_playlist]["files"].append({"path": p, "times": [], "repeats": 1})
        self._refresh_media_table()
        self._save_playlists()

    def _play_selected_media(self):
        idx = self._get_selected_media_index()
        if idx is None:
            messagebox.showwarning("Aviso", "Selecione um item para tocar.")
            return
        media = self.playlists[self.current_playlist]["files"][idx]
        path = media.get("path")
        if not path or not os.path.exists(path):
            messagebox.showerror("Erro", "Arquivo não encontrado.")
            return
        try:
            repeats = int(simpledialog.askstring("Repetições", "Quantas vezes? (1-50)", initialvalue=str(media.get("repeats",1))) or "1")
        except Exception:
            repeats = 1
        media["repeats"] = max(1, min(50, repeats))
        self._save_playlists()
        # duck others (pycaw) and play async
        self.duck_all_sessions(target=0.06, exclude_pids={os.getpid()}, steps=6, step_ms=120)
        self.play_media_async(path, media["repeats"])

    def _generate_schedule(self):
        msg = f"Gerar agenda (mock)\n\nPlaylist: {self.current_playlist}\nRepetir global: {self._get_repeat_global()}x\nDistribuição: ~{int(self.distrib_scale.get())}h"
        messagebox.showinfo("Gerar Novo", msg)

    # ------------------------ Playback worker ------------------------
    def play_media_async(self, path, repeats=1):
        if self._is_playing:
            messagebox.showinfo("Info", "Já está tocando outro arquivo. Aguarde.")
            return
        th = threading.Thread(target=self._playback_worker, args=(path, repeats), daemon=True)
        th.start()
        self._playback_thread = th

    def _playback_worker(self, path, repeats):
        self._is_playing = True
        try:
            with self._playback_lock:
                mixer.music.load(path)
                for _ in range(repeats):
                    mixer.music.play()
                    while mixer.music.get_busy():
                        time.sleep(0.08)
        except Exception as e:
            print("Playback error:", e)
        finally:
            self._is_playing = False
            # restore volumes smoothly
            self.after(150, lambda: self.restore_all_sessions(steps=10, step_ms=150))

    # ------------------------ pycaw duck helpers ------------------------
    def _get_all_audio_sessions(self):
        if not PYCAW_OK:
            return []
        sessions = []
        try:
            all_sessions = AudioUtilities.GetAllSessions()
            for s in all_sessions:
                pid = getattr(s.Process, "pid", None) if getattr(s, "Process", None) else None
                vol_iface = None
                try:
                    vol_iface = s._ctl.QueryInterface(ISimpleAudioVolume)
                except Exception:
                    vol_iface = None
                sessions.append((pid, s, vol_iface))
        except Exception:
            return []
        return sessions

    def duck_all_sessions(self, target=0.06, exclude_pids=None, steps=6, step_ms=120):
        if not PYCAW_OK:
            print("pycaw not available")
            return False
        if exclude_pids is None:
            exclude_pids = set()
        sessions = self._get_all_audio_sessions()
        if not self._duck_active:
            self._saved_sessions = {}
        to_duck = []
        for pid, s, vol in sessions:
            if vol is None:
                continue
            if pid in exclude_pids or pid == os.getpid():
                continue
            key = pid if pid is not None else id(s)
            try:
                orig = vol.GetMasterVolume()
            except Exception:
                orig = 1.0
            if key not in self._saved_sessions:
                self._saved_sessions[key] = float(orig)
            to_duck.append((key, vol))
        if not to_duck:
            self._duck_active = True
            return True
        def step(i):
            t = i / float(steps)
            for key, vol in to_duck:
                orig = self._saved_sessions.get(key, 1.0)
                new = float(orig) + (float(target) - float(orig)) * t
                try:
                    vol.SetMasterVolume(max(0.0, min(1.0, new)), None)
                except Exception:
                    pass
            if i < steps:
                self.after(step_ms, lambda: step(i+1))
            else:
                self._duck_active = True
        self.after(0, lambda: step(1))
        return True

    def restore_all_sessions(self, steps=10, step_ms=150):
        if not PYCAW_OK:
            return False
        if not self._saved_sessions:
            return True
        sessions = self._get_all_audio_sessions()
        vol_map = {}
        for pid, s, vol in sessions:
            if vol is None:
                continue
            key = pid if pid is not None else id(s)
            vol_map[key] = vol
        def step(i):
            t = i / float(steps)
            for key, orig in list(self._saved_sessions.items()):
                voliface = vol_map.get(key)
                if voliface is None:
                    continue
                try:
                    cur = voliface.GetMasterVolume()
                except Exception:
                    cur = orig
                new = float(cur) + (float(orig) - float(cur)) * t
                try:
                    voliface.SetMasterVolume(max(0.0, min(1.0, new)), None)
                except Exception:
                    pass
            if i < steps:
                self.after(step_ms, lambda: step(i+1))
            else:
                for key, orig in list(self._saved_sessions.items()):
                    voliface = vol_map.get(key)
                    if voliface:
                        try:
                            voliface.SetMasterVolume(float(orig), None)
                        except Exception:
                            pass
                self._saved_sessions = {}
                self._duck_active = False
        self.after(0, lambda: step(1))
        return True

    # ------------------------ Mic passthrough (low-latency) ------------------------
    def _toggle_mic(self):
        if not SOUND_OK:
            messagebox.showwarning("Mic", "Instale 'sounddevice' e 'numpy' para usar o microfone.")
            return
        if not self._mic_active:
            # duck others
            self.duck_all_sessions(target=0.03, exclude_pids={os.getpid()}, steps=6, step_ms=60)
            self._start_mic()
        else:
            self._stop_mic()
            self.restore_all_sessions(steps=8, step_ms=120)

    def _start_mic(self):
        if self._mic_active:
            return
        # parameters tuned for low latency
        samplerate = 44100
        channels_in = 1
        channels_out = 2
        blocksize = 256
        latency = 'low'

        def callback(indata, outdata, frames, time_info, status):
            try:
                data = indata * (self._mic_gain)
                # mono to stereo
                if data.ndim == 2 and data.shape[1] == 1:
                    outdata[:,0] = data[:,0]
                    outdata[:,1] = data[:,0]
                else:
                    outdata[:] = data
            except Exception:
                outdata.fill(0)

        try:
            device_pair = None
            if self._mic_input_device is not None or self._mic_output_device is not None:
                in_dev = self._mic_input_device
                out_dev = self._mic_output_device
                device_pair = (in_dev, out_dev)
            self._mic_stream = sd.Stream(samplerate=samplerate, blocksize=blocksize, device=device_pair,
                                         channels=(channels_in, channels_out), callback=callback, latency=latency)
            self._mic_stream.start()
            self._mic_active = True
            self._mic_btn.config(text="🎙 Mic (ON)")
            self._mic_label.config(text="Mic: on")
        except Exception as e:
            messagebox.showerror("Mic", f"Erro iniciando microfone: {e}")
            self._mic_stream = None
            self._mic_active = False

    def _stop_mic(self):
        if not self._mic_active:
            return
        try:
            if self._mic_stream:
                self._mic_stream.stop()
                self._mic_stream.close()
        except Exception:
            pass
        self._mic_stream = None
        self._mic_active = False
        self._mic_btn.config(text="🎙 Mic")
        self._mic_label.config(text="Mic: off")

    def _on_mic_vol_change(self, _v):
        try:
            v = float(self._mic_vol.get())
            self._mic_gain = max(0.0, v / 100.0)
        except Exception:
            self._mic_gain = 1.0

    # ------------------------ Global lock logic (single lock) ------------------------
    def _toggle_global_lock(self):
        if self._global_locked:
            # ask PIN to unlock
            pin = simpledialog.askstring("Senha", "Digite a senha para liberar edição (PIN):", show="*")
            if pin != SCHEDULE_PIN:
                messagebox.showerror("Senha incorreta", "PIN inválido.")
                return
            self._global_locked = False
            messagebox.showinfo("Liberado", "Edição liberada para todos os anúncios.")
        else:
            # lock again (no PIN)
            self._global_locked = True
            messagebox.showinfo("Bloqueado", "Edição bloqueada novamente.")
        self._update_global_lock_btn()
        self._save_config()

    # ------------------------ Schedule editor ------------------------
    def _open_schedule_editor(self, idx):
        try:
            media = self.playlists[self.current_playlist]["files"][idx]
        except Exception:
            return
        dlg = tk.Toplevel(self)
        dlg.title("Editor de Horários")
        dlg.geometry("520x420")
        dlg.transient(self)
        dlg.grab_set()
        dlg.configure(bg=BG)

        header = ttk.Label(dlg, text=os.path.basename(media.get("path", "")), style="Accent.TLabel")
        header.pack(fill="x", padx=12, pady=(12,6))

        listbox = tk.Listbox(dlg, bg=CARD, fg=TEXT, height=10, font=DEFAULT_FONT)
        listbox.pack(fill="both", expand=True, padx=12, pady=(6,6))
        for t in media.get("times", []):
            listbox.insert(tk.END, t)

        ctrl = ttk.Frame(dlg, style="App.TFrame")
        ctrl.pack(fill="x", padx=12, pady=(6,12))
        ctrl.columnconfigure(0, weight=1); ctrl.columnconfigure(1, weight=1); ctrl.columnconfigure(2, weight=1)

        def add_time():
            new = simpledialog.askstring("Adicionar horário", "Horário (HH:MM):", parent=dlg, initialvalue="12:00")
            if not new: return
            try:
                time.strptime(new, "%H:%M")
            except Exception:
                messagebox.showerror("Erro", "Formato inválido (HH:MM).")
                return
            media.setdefault("times", []).append(new)
            listbox.insert(tk.END, new)
            self._save_playlists()
            self._refresh_media_table()

        def edit_time():
            sel = listbox.curselection()
            if not sel: return
            cur = listbox.get(sel[0])
            new = simpledialog.askstring("Editar horário", "Horário (HH:MM):", parent=dlg, initialvalue=cur)
            if not new: return
            try:
                time.strptime(new, "%H:%M")
            except Exception:
                messagebox.showerror("Erro", "Formato inválido (HH:MM).")
                return
            media["times"][sel[0]] = new
            listbox.delete(sel[0]); listbox.insert(sel[0], new)
            self._save_playlists()
            self._refresh_media_table()

        def remove_time():
            sel = listbox.curselection()
            if not sel: return
            if not messagebox.askyesno("Remover", "Remover horário selecionado?"): return
            media["times"].pop(sel[0])
            listbox.delete(sel[0])
            self._save_playlists()
            self._refresh_media_table()

        ttk.Button(ctrl, text="＋ Adicionar", style="Neon.TButton", command=add_time).grid(row=0, column=0, sticky="ew", padx=6)
        ttk.Button(ctrl, text="✎ Editar", style="Neon.TButton", command=edit_time).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(ctrl, text="🗑 Remover", style="Neon.TButton", command=remove_time).grid(row=0, column=2, sticky="ew", padx=6)

        bottom = ttk.Frame(dlg, style="App.TFrame")
        bottom.pack(fill="x", padx=12, pady=(0,12))
        bottom.columnconfigure(0, weight=1)
        ttk.Label(bottom, text="Repetições (loop):", background=BG, foreground=TEXT).grid(row=0, column=0, sticky="w")
        repeats = tk.IntVar(value=media.get("repeats", 1))
        spin = tk.Spinbox(bottom, from_=1, to=100, textvariable=repeats, width=6)
        spin.grid(row=0, column=1, sticky="e")
        def save_and_close():
            media["repeats"] = int(repeats.get())
            self._save_playlists()
            self._refresh_media_table()
            dlg.destroy()
        ttk.Button(bottom, text="Salvar", style="Primary.TButton", command=save_and_close).grid(row=0, column=2, sticky="e", padx=8)

    # ------------------------ Context actions (menu) ------------------------
    def _ctx_set_time(self):
        # respect global lock
        if self._global_locked:
            pin = simpledialog.askstring("Senha", "Digite a senha para editar (PIN):", show="*")
            if pin != SCHEDULE_PIN:
                messagebox.showerror("Senha", "PIN incorreto.")
                return
            self._global_locked = False; self._update_global_lock_btn(); self._save_config()
        idx = self._get_selected_media_index()
        if idx is not None:
            self._open_schedule_editor(idx)

    def _ctx_set_repeat(self):
        if self._global_locked:
            pin = simpledialog.askstring("Senha", "Digite a senha para editar (PIN):", show="*")
            if pin != SCHEDULE_PIN:
                messagebox.showerror("Senha", "PIN incorreto.")
                return
            self._global_locked = False; self._update_global_lock_btn(); self._save_config()
        idx = self._get_selected_media_index()
        if idx is None:
            return
        media = self.playlists[self.current_playlist]["files"][idx]
        try:
            new_r = int(simpledialog.askstring("Repetições", "Quantas vezes? (1-100):", initialvalue=str(media.get("repeats",1))) or "1")
        except Exception:
            return
        media["repeats"] = max(1, min(100, new_r))
        self._save_playlists()
        self._refresh_media_table()

    def _ctx_remove_item(self):
        if self._global_locked:
            pin = simpledialog.askstring("Senha", "Digite a senha para editar (PIN):", show="*")
            if pin != SCHEDULE_PIN:
                messagebox.showerror("Senha", "PIN incorreto.")
                return
            self._global_locked = False; self._update_global_lock_btn(); self._save_config()
        idx = self._get_selected_media_index()
        if idx is None:
            return
        if not messagebox.askyesno("Remover", "Deseja remover o item selecionado?"): return
        self.playlists[self.current_playlist]["files"].pop(idx)
        self._save_playlists()
        self._refresh_media_table()

    # ------------------------ Import / Export ------------------------
    def _export_playlist(self):
        if not self.current_playlist:
            messagebox.showwarning("Aviso", "Selecione uma playlist para exportar.")
            return
        folder = filedialog.askdirectory(title="Exportar para pasta")
        if not folder: return
        pl = self.playlists[self.current_playlist]
        export_folder = os.path.join(folder, f"export_{self.current_playlist}")
        os.makedirs(export_folder, exist_ok=True)
        exported = []
        for m in pl.get("files", []):
            src = m.get("path")
            if not src or not os.path.exists(src): continue
            try:
                dst = os.path.join(export_folder, os.path.basename(src))
                shutil.copy2(src, dst)
                exported.append({"path": os.path.basename(src), "times": m.get("times", []), "repeats": m.get("repeats",1)})
            except Exception:
                pass
        if not exported:
            messagebox.showwarning("Exportar", "Nenhum arquivo exportado.")
            shutil.rmtree(export_folder, ignore_errors=True)
            return
        cfg = {"playlist": {"files": exported, "time": pl.get("time", "00:00"), "repeats": pl.get("repeats",1)}, "metadata": {"playlist_name": self.current_playlist, "export_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}}
        with open(os.path.join(export_folder, "playlist_config.json"), "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=4, ensure_ascii=False)
        messagebox.showinfo("Exportar", f"Exportado em: {export_folder}")

    def _import_playlist(self):
        folder = filedialog.askdirectory(title="Selecione pasta exportada")
        if not folder: return
        cfg_path = os.path.join(folder, "playlist_config.json")
        if not os.path.exists(cfg_path):
            messagebox.showerror("Importar", "playlist_config.json não encontrado.")
            return
        try:
            with open(cfg_path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            messagebox.showerror("Importar", f"Erro lendo arquivo: {e}")
            return
        name = data.get("metadata", {}).get("playlist_name", f"import_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
        base = name; i=1
        while name in self.playlists:
            name = f"{base}_{i}"; i+=1
        imported = []
        for entry in data.get("playlist", {}).get("files", []):
            fname = entry.get("path")
            path = os.path.join(folder, fname)
            if os.path.exists(path):
                imported.append({"path": path, "times": entry.get("times", []), "repeats": entry.get("repeats",1)})
        if not imported:
            messagebox.showwarning("Importar", "Nenhum arquivo válido encontrado para importar.")
            return
        self.playlists[name] = {"files": imported, "time": data.get("playlist", {}).get("time", "00:00"), "repeats": data.get("playlist", {}).get("repeats",1), "active": True}
        self.current_playlist = name
        self._refresh_playlist_list()
        self._refresh_media_table()
        self._save_playlists()

    # ------------------------ Scheduler tick ------------------------
    def _schedule_tick(self):
        now = datetime.now().strftime("%H:%M")
        for pl_name, pl in self.playlists.items():
            if not pl.get("active", True): continue
            if pl.get("time") == now:
                self._play_playlist(pl_name)
            for m in pl.get("files", []):
                for t in m.get("times", []):
                    if t == now:
                        self.play_media_async(m.get("path"), m.get("repeats",1))
        self.after(60000, self._schedule_tick)

    def _play_playlist(self, playlist_name):
        for m in self.playlists[playlist_name]["files"]:
            self.duck_all_sessions(target=0.06, exclude_pids={os.getpid()}, steps=6, step_ms=120)
            self.play_media_async(m.get("path"), m.get("repeats",1))

    # ------------------------ Misc ------------------------
    def _get_repeat_global(self):
        try:
            return int(self.repeat_cb.get().split()[0])
        except Exception:
            return 3

    def _get_selected_media_index(self):
        sel = self.tree.selection()
        if not sel: return None
        return self.tree.index(sel[0])

    def _on_close(self):
        try:
            self._stop_mic()
        except Exception:
            pass
        try:
            self.restore_all_sessions(steps=1, step_ms=10)
        except Exception:
            pass
        self._save_playlists()
        self._save_config()
        try:
            mixer.quit()
        except Exception:
            pass
        self.destroy()

    # ------------------------ Menu & debug ------------------------
    def _open_main_menu(self):
        menu = Menu(self, tearoff=0, bg=PANEL, fg=TEXT)
        pm = Menu(menu, tearoff=0, bg=PANEL, fg=TEXT)
        pm.add_command(label="Nova Playlist", command=self._create_playlist)
        pm.add_command(label="Renomear Playlist", command=self._rename_playlist)
        pm.add_command(label="Excluir Playlist", command=self._delete_playlist)
        pm.add_separator()
        pm.add_command(label="Exportar Playlist", command=self._export_playlist)
        pm.add_command(label="Importar Playlist", command=self._import_playlist)
        menu.add_cascade(label="Playlist", menu=pm)

        cm = Menu(menu, tearoff=0, bg=PANEL, fg=TEXT)
        cm.add_command(label="Ligar/Desligar Playlist", command=self._toggle_current_playlist)
        menu.add_cascade(label="Configurações", menu=cm)

        menu.add_command(label="Config Mic", command=self._open_mic_config)
        menu.add_separator()
        menu.add_command(label="Salvar Tudo", command=self._save_playlists)

        try:
            menu.tk_popup(self._gear_btn.winfo_rootx(), self._gear_btn.winfo_rooty() + self._gear_btn.winfo_height())
        finally:
            menu.grab_release()

    def _rename_playlist(self):
        if not self.current_playlist:
            messagebox.showwarning("Aviso", "Selecione uma playlist primeiro.")
            return
        new = simpledialog.askstring("Renomear", "Novo nome:", initialvalue=self.current_playlist)
        if not new or new == self.current_playlist:
            return
        self.playlists[new] = self.playlists.pop(self.current_playlist)
        self.current_playlist = new
        self._refresh_playlist_list()
        self._save_playlists()

    def _delete_playlist(self):
        if not self.current_playlist:
            messagebox.showwarning("Aviso", "Selecione uma playlist primeiro.")
            return
        if not messagebox.askyesno("Confirmar", f"Excluir playlist '{self.current_playlist}'?"):
            return
        self.playlists.pop(self.current_playlist, None)
        self.current_playlist = None
        self._refresh_playlist_list()
        self._refresh_media_table()
        self._save_playlists()

    # ------------------------ Mic config dialog ------------------------
    def _open_mic_config(self):
        if not SOUND_OK:
            messagebox.showwarning("Config Mic", "Instale 'sounddevice' e 'numpy' para configurar microfones.")
            return
        try:
            devices = sd.query_devices()
        except Exception as e:
            messagebox.showerror("Dispositivos", f"Erro listando dispositivos: {e}")
            return

        dlg = tk.Toplevel(self)
        dlg.title("Config Mic")
        dlg.geometry("760x420")
        dlg.transient(self)
        dlg.grab_set()
        dlg.configure(bg=BG)

        left = ttk.Frame(dlg, style="App.TFrame")
        left.grid(row=0, column=0, sticky="nsew", padx=12, pady=12)
        left.columnconfigure(0, weight=1)
        ttk.Label(left, text="Entradas (mic)", style="Accent.TLabel").grid(row=0, column=0, sticky="w")
        input_list = tk.Listbox(left, bg=CARD, fg=TEXT)
        input_list.grid(row=1, column=0, sticky="nsew", pady=(6,0))
        left.rowconfigure(1, weight=1)

        right = ttk.Frame(dlg, style="App.TFrame")
        right.grid(row=0, column=1, sticky="nsew", padx=12, pady=12)
        right.columnconfigure(0, weight=1)
        ttk.Label(right, text="Saídas (alto-falante)", style="Accent.TLabel").grid(row=0, column=0, sticky="w")
        output_list = tk.Listbox(right, bg=CARD, fg=TEXT)
        output_list.grid(row=1, column=0, sticky="nsew", pady=(6,0))
        right.rowconfigure(1, weight=1)

        # populate
        input_indices = []
        output_indices = []
        try:
            devices = sd.query_devices()
            for i, d in enumerate(devices):
                name = f"{i}: {d['name']} (in:{d.get('max_input_channels',0)} out:{d.get('max_output_channels',0)})"
                if d.get('max_input_channels', 0) > 0:
                    input_list.insert(tk.END, name); input_indices.append(i)
                if d.get('max_output_channels', 0) > 0:
                    output_list.insert(tk.END, name); output_indices.append(i)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível listar dispositivos: {e}")
            dlg.destroy(); return

        # preselect
        try:
            if self._mic_input_device in input_indices:
                input_list.selection_set(input_indices.index(self._mic_input_device))
                input_list.see(input_indices.index(self._mic_input_device))
        except Exception:
            pass
        try:
            if self._mic_output_device in output_indices:
                output_list.selection_set(output_indices.index(self._mic_output_device))
                output_list.see(output_indices.index(self._mic_output_device))
        except Exception:
            pass

        bottom = ttk.Frame(dlg, style="App.TFrame")
        bottom.grid(row=1, column=0, columnspan=2, sticky="ew", padx=12, pady=(6,12))
        bottom.columnconfigure(0, weight=1)

        def on_update():
            input_list.delete(0, tk.END); output_list.delete(0, tk.END)
            try:
                devs = sd.query_devices()
                for i, d in enumerate(devs):
                    name = f"{i}: {d['name']} (in:{d.get('max_input_channels',0)} out:{d.get('max_output_channels',0)})"
                    if d.get('max_input_channels',0) > 0:
                        input_list.insert(tk.END, name); input_indices.append(i)
                    if d.get('max_output_channels',0) > 0:
                        output_list.insert(tk.END, name); output_indices.append(i)
                messagebox.showinfo("Atualizado", "Lista de dispositivos atualizada.")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro atualizando: {e}")

        def on_save():
            sel_in = input_list.curselection()
            sel_out = output_list.curselection()
            new_in = input_indices[sel_in[0]] if sel_in else None
            new_out = output_indices[sel_out[0]] if sel_out else None
            self._mic_input_device = new_in
            self._mic_output_device = new_out
            self._save_config()
            messagebox.showinfo("Config Mic", f"Salvo. Entrada: {new_in} Saída: {new_out}")
            dlg.destroy()

        ttk.Button(bottom, text="Atualizar", style="Neon.TButton", command=on_update).grid(row=0, column=0, sticky="w", padx=(0,6))
        ttk.Button(bottom, text="Salvar", style="Primary.TButton", command=on_save).grid(row=0, column=1, sticky="e", padx=(6,0))

# ------------------------ Run ------------------------
if __name__ == "__main__":
    app = TimelyAdsApp()
    app.mainloop()
