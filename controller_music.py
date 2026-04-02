import os
import json
import asyncio
import threading
import ctypes
import websockets
import customtkinter as ctk
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume


APP_NAME = "ControllerMusic"
APPDATA_DIR = os.path.join(os.getenv("APPDATA"), APP_NAME)
SETTINGS_PATH = os.path.join(APPDATA_DIR, "settings.json")


DEFAULT_SETTINGS = {
    "theme": "dark",
    "accent": "#00aaff",
    "apps": ["spotify.exe"],
    "mute_volume": 5,
    "fade_down": 0.1,
    "fade_up": 0.1,
    "fades_enabled": True,
    "auto_start_monitoring": False,
    "auto_stop_on_disconnect": True,
    "only_duck_on_rx": False,
    "restore_mode": "previous"
}


class SettingsManager:
    def __init__(self, path=SETTINGS_PATH):
        self.path = path
        self.data = {}
        self.ensure_exists()
        self.load()

    def ensure_exists(self):
        if not os.path.isdir(APPDATA_DIR):
            os.makedirs(APPDATA_DIR, exist_ok=True)
        if not os.path.isfile(self.path):
            with open(self.path, "w", encoding="utf-8") as f:
                json.dump(DEFAULT_SETTINGS, f, indent=2)

    def load(self):
        try:
            with open(self.path, "r", encoding="utf-8") as f:
                loaded = json.load(f)
            self.data = DEFAULT_SETTINGS | loaded
        except Exception:
            self.data = DEFAULT_SETTINGS.copy()
            self.save()

    def save(self):
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump(self.data, f, indent=2)

    def get(self, key):
        return self.data.get(key, DEFAULT_SETTINGS.get(key))

    def set(self, key, value):
        self.data[key] = value
        self.save()


class MultiAppVolumeController:
    def __init__(self):
        self._lock = asyncio.Lock()

    def _find_sessions(self, names):
        names = [n.lower() for n in names]
        sessions = AudioUtilities.GetAllSessions()
        result = []
        for s in sessions:
            if s.Process:
                name = s.Process.name().lower()
                if name in names:
                    result.append(s._ctl.QueryInterface(ISimpleAudioVolume))
        return result

    def get_average_volume(self, names):
        try:
            ctypes.windll.ole32.CoInitialize(0)
            vols = self._find_sessions(names)
            if not vols:
                return 100
            values = [v.GetMasterVolume() * 100 for v in vols]
            return sum(values) / len(values)
        finally:
            ctypes.windll.ole32.CoUninitialize()

    def set_volume(self, names, percent):
        try:
            ctypes.windll.ole32.CoInitialize(0)
            vols = self._find_sessions(names)
            for v in vols:
                v.SetMasterVolume(max(0, min(1, percent / 100)), None)
        finally:
            ctypes.windll.ole32.CoUninitialize()

    async def fade_to(self, names, target, duration):
        async with self._lock:
            start = self.get_average_volume(names)
            steps = max(1, int(duration * 60))
            step = (target - start) / steps
            delay = duration / steps
            for i in range(steps):
                self.set_volume(names, start + step * (i + 1))
                await asyncio.sleep(delay)


class TransmissionState:
    def __init__(self):
        self.tx = False
        self.rx = set()

    def active(self):
        return self.tx or bool(self.rx)

    def update(self, event_type, value):
        if event_type == "kTxBegin":
            self.tx = True
        elif event_type == "kTxEnd":
            self.tx = False
        elif event_type == "kRxBegin":
            self.rx.add(value.get("pFrequencyHz"))
        elif event_type == "kRxEnd":
            self.rx.discard(value.get("pFrequencyHz"))


class TrackAudioMonitor:
    def __init__(self, uri, on_start, on_end, on_status, on_disconnect, only_duck_on_rx):
        self.uri = uri
        self.state = TransmissionState()
        self.on_start = on_start
        self.on_end = on_end
        self.on_status = on_status
        self.on_disconnect = on_disconnect
        self.only_duck_on_rx = only_duck_on_rx
        self.running = False
        self.listener_task = None

    async def connect(self):
        self.running = True
        while self.running:
            try:
                self.on_status("Connecting…")
                async with websockets.connect(self.uri) as ws:
                    self.on_status("Connected")
                    self.listener_task = asyncio.create_task(self.listen(ws))
                    await self.listener_task
            except Exception:
                if not self.running:
                    break
                self.on_disconnect()
                break

    async def listen(self, ws):
        while self.running:
            msg = await ws.recv()
            data = json.loads(msg)
            await self.handle_event(data)

    async def handle_event(self, event):
        event_type = event.get("type")
        value = event.get("value", {})
        before = self.state.active()
        self.state.update(event_type, value)
        after = self.state.active()

        if self.only_duck_on_rx:
            before = bool(self.state.rx)
            after = bool(self.state.rx)

        if not before and after:
            await self.on_start()
        elif before and not after:
            await self.on_end()

    def stop(self):
        self.running = False
        if self.listener_task:
            self.listener_task.cancel()
            self.listener_task = None
        self.state = TransmissionState()


class AudioDuckController:
    def __init__(self, settings: SettingsManager, uri="ws://localhost:49080/ws"):
        self.settings = settings
        self.audio = MultiAppVolumeController()
        self.uri = uri
        self.apps = self._parse_apps(self.settings.get("apps"))
        self.target = self.settings.get("mute_volume")
        self.fade_down = self.settings.get("fade_down")
        self.fade_up = self.settings.get("fade_up")
        self.original = 100
        self.monitor = None
        self.fade_task = None
        self.fades_enabled = self.settings.get("fades_enabled")

    def _parse_apps(self, apps):
        if isinstance(apps, list):
            return [a.lower() for a in apps]
        if isinstance(apps, str):
            parts = [p.strip().lower() for p in apps.replace(";", ",").split(",") if p.strip()]
            return parts or ["spotify.exe"]
        return ["spotify.exe"]

    async def start(self, status_callback, disconnect_callback):
        self.monitor = TrackAudioMonitor(
            self.uri,
            self.lower_volume,
            self.restore_volume,
            status_callback,
            disconnect_callback,
            self.settings.get("only_duck_on_rx")
        )
        await self.monitor.connect()

    def cancel_fade(self):
        if self.fade_task and not self.fade_task.done():
            self.fade_task.cancel()
        self.fade_task = None

    def force_restore(self):
        if self.settings.get("restore_mode") == "fixed":
            fixed = self.settings.get("mute_volume")
            self.audio.set_volume(self.apps, fixed)
        else:
            self.audio.set_volume(self.apps, self.original)

    def stop(self):
        if self.monitor:
            self.monitor.stop()
        self.cancel_fade()
        self.force_restore()

    async def lower_volume(self):
        self.cancel_fade()
        self.original = self.audio.get_average_volume(self.apps)
        if self.fades_enabled:
            self.fade_task = asyncio.create_task(
                self.audio.fade_to(self.apps, self.target, self.fade_down)
            )
            await self.fade_task
        else:
            self.audio.set_volume(self.apps, self.target)

    async def restore_volume(self):
        self.cancel_fade()
        if self.fades_enabled:
            target = self.original if self.settings.get("restore_mode") == "previous" else self.settings.get("mute_volume")
            self.fade_task = asyncio.create_task(
                self.audio.fade_to(self.apps, target, self.fade_up)
            )
            await self.fade_task
        else:
            self.force_restore()

    async def test_duck(self):
        await self.lower_volume()
        await asyncio.sleep(1)
        await self.restore_volume()

    def refresh_from_settings(self):
        self.apps = self._parse_apps(self.settings.get("apps"))
        self.target = self.settings.get("mute_volume")
        self.fade_down = self.settings.get("fade_down")
        self.fade_up = self.settings.get("fade_up")
        self.fades_enabled = self.settings.get("fades_enabled")

class SettingsWindow(ctk.CTkToplevel):
    def __init__(self, master, settings: SettingsManager, controller: AudioDuckController):
        super().__init__(master)
        self.settings = settings
        self.controller = controller
        self.title("Settings")
        self.geometry("520x480")
        self.resizable(False, False)

        self.tabview = ctk.CTkTabview(self)
        self.tabview.pack(fill="both", expand=True, padx=10, pady=10)

        general_tab = self.tabview.add("General")
        appearance_tab = self.tabview.add("Appearance")
        audio_tab = self.tabview.add("Audio")
        apps_tab = self.tabview.add("Apps")

        self.build_general_tab(general_tab)
        self.build_appearance_tab(appearance_tab)
        self.build_audio_tab(audio_tab)
        self.build_apps_tab(apps_tab)

    def build_general_tab(self, parent):
        auto_start_var = ctk.BooleanVar(value=self.settings.get("auto_start_monitoring"))
        auto_stop_var = ctk.BooleanVar(value=self.settings.get("auto_stop_on_disconnect"))

        auto_start = ctk.CTkCheckBox(
            parent, text="Auto start monitoring on launch",
            variable=auto_start_var,
            command=lambda: self.settings.set("auto_start_monitoring", auto_start_var.get())
        )
        auto_start.pack(anchor="w", pady=5, padx=10)

        auto_stop = ctk.CTkCheckBox(
            parent, text="Auto stop on TrackAudio disconnect",
            variable=auto_stop_var,
            command=lambda: self.settings.set("auto_stop_on_disconnect", auto_stop_var.get())
        )
        auto_stop.pack(anchor="w", pady=5, padx=10)

    def build_appearance_tab(self, parent):
        theme_var = ctk.StringVar(value=self.settings.get("theme"))

        theme_label = ctk.CTkLabel(parent, text="Theme")
        theme_label.pack(anchor="w", padx=10, pady=(10, 0))

        theme_option = ctk.CTkOptionMenu(
            parent,
            values=["light", "dark", "system"],
            variable=theme_var,
            command=self.on_theme_change
        )
        theme_option.pack(anchor="w", padx=10, pady=5)

    def on_theme_change(self, value):
        self.settings.set("theme", value)
        if value == "light":
            ctk.set_appearance_mode("Light")
        elif value == "dark":
            ctk.set_appearance_mode("Dark")
        else:
            ctk.set_appearance_mode("System")

    def build_audio_tab(self, parent):
        mute_var = ctk.DoubleVar(value=self.settings.get("mute_volume"))
        fade_down_var = ctk.DoubleVar(value=self.settings.get("fade_down"))
        fade_up_var = ctk.DoubleVar(value=self.settings.get("fade_up"))
        fades_enabled_var = ctk.BooleanVar(value=self.settings.get("fades_enabled"))

        mute_label = ctk.CTkLabel(parent, text="Mute Volume (%)")
        mute_label.pack(anchor="w", padx=10, pady=(10, 0))
        mute_slider = ctk.CTkSlider(
            parent, from_=0, to=100, variable=mute_var,
            command=lambda v: self.on_mute_change(mute_var.get())
        )
        mute_slider.pack(fill="x", padx=10, pady=5)

        fades_toggle = ctk.CTkCheckBox(
            parent, text="Enable fades",
            variable=fades_enabled_var,
            command=lambda: self.on_fades_toggle(fades_enabled_var.get())
        )
        fades_toggle.pack(anchor="w", padx=10, pady=5)

        fade_down_label = ctk.CTkLabel(parent, text="Fade Down (s)")
        fade_down_label.pack(anchor="w", padx=10, pady=(10, 0))
        fade_down_slider = ctk.CTkSlider(
            parent, from_=0.1, to=2.0, variable=fade_down_var,
            command=lambda v: self.on_fade_down_change(fade_down_var.get())
        )
        fade_down_slider.pack(fill="x", padx=10, pady=5)

        fade_up_label = ctk.CTkLabel(parent, text="Fade Up (s)")
        fade_up_label.pack(anchor="w", padx=10, pady=(10, 0))
        fade_up_slider = ctk.CTkSlider(
            parent, from_=0.1, to=2.0, variable=fade_up_var,
            command=lambda v: self.on_fade_up_change(fade_up_var.get())
        )
        fade_up_slider.pack(fill="x", padx=10, pady=5)

        if not fades_enabled_var.get():
            fade_down_slider.configure(state="disabled")
            fade_up_slider.configure(state="disabled")

        self._fade_down_slider = fade_down_slider
        self._fade_up_slider = fade_up_slider

    def on_mute_change(self, value):
        self.settings.set("mute_volume", float(value))
        self.controller.refresh_from_settings()

    def on_fades_toggle(self, enabled):
        self.settings.set("fades_enabled", bool(enabled))
        self.controller.refresh_from_settings()
        state = "normal" if enabled else "disabled"
        self._fade_down_slider.configure(state=state)
        self._fade_up_slider.configure(state=state)

    def on_fade_down_change(self, value):
        self.settings.set("fade_down", float(value))
        self.controller.refresh_from_settings()

    def on_fade_up_change(self, value):
        self.settings.set("fade_up", float(value))
        self.controller.refresh_from_settings()

    def build_apps_tab(self, parent):
        apps_str = ",".join(self.settings.get("apps"))
        apps_var = ctk.StringVar(value=apps_str)

        label = ctk.CTkLabel(parent, text="Apps to duck (exe names, separated by , or ;)")
        label.pack(anchor="w", padx=10, pady=(10, 0))

        entry = ctk.CTkEntry(parent, textvariable=apps_var)
        entry.pack(fill="x", padx=10, pady=5)

        def save_apps():
            raw = apps_var.get()
            parts = [p.strip().lower() for p in raw.replace(";", ",").split(",") if p.strip()]
            if not parts:
                parts = ["spotify.exe"]
            self.settings.set("apps", parts)
            self.controller.refresh_from_settings()

        save_btn = ctk.CTkButton(parent, text="Apply", command=save_apps)
        save_btn.pack(anchor="e", padx=10, pady=5)

        quick_frame = ctk.CTkFrame(parent)
        quick_frame.pack(fill="x", padx=10, pady=10)

        def add_app(name):
            current = [p.strip().lower() for p in apps_var.get().replace(";", ",").split(",") if p.strip()]
            if name not in current:
                current.append(name)
            apps_var.set(",".join(current))
            save_apps()

        buttons = [
            ("Spotify", "spotify.exe"),
            ("Chrome", "chrome.exe"),
            ("Firefox", "firefox.exe"),
            ("Edge", "msedge.exe"),
            ("Discord", "discord.exe"),
            ("VLC", "vlc.exe"),
        ]

        for text, exe in buttons:
            btn = ctk.CTkButton(quick_frame, text=text, width=80,
                                command=lambda e=exe: add_app(e))
            btn.pack(side="left", padx=5, pady=5)


class DuckingUI(ctk.CTk):
    def __init__(self, settings: SettingsManager):
        super().__init__()
        self.settings = settings
        self.apply_theme()
        self.title("Controller Music")
        self.geometry("480x650")
        self.resizable(False, False)

        self.controller = AudioDuckController(self.settings)
        self.monitor_thread = None

        top_frame = ctk.CTkFrame(self, fg_color="transparent")
        top_frame.pack(fill="x", pady=(5, 0))

        title_label = ctk.CTkLabel(top_frame, text="🎵 Controller Music",
                                   font=ctk.CTkFont(size=26, weight="bold"))
        title_label.pack(side="left", padx=10, pady=10)

        settings_btn = ctk.CTkButton(top_frame, text="⚙", width=32,
                                     command=self.open_settings)
        settings_btn.pack(side="right", padx=10, pady=10)

        self.status_label = ctk.CTkLabel(self, text="Ready",
                                         font=ctk.CTkFont(size=14, weight="bold"),
                                         text_color="#00aaff")
        self.status_label.pack(pady=5)

        self.vol_var = ctk.DoubleVar(value=self.settings.get("mute_volume"))
        vol_frame = ctk.CTkFrame(self, fg_color="transparent")
        vol_frame.pack(pady=5)
        ctk.CTkLabel(vol_frame, text="Mute Volume (%)").pack()
        self.vol_slider = ctk.CTkSlider(
            vol_frame, from_=0, to=100,
            variable=self.vol_var,
            command=lambda v: self.on_ui_mute_change()
        )
        self.vol_slider.pack(pady=5)
        self.vol_label = ctk.CTkLabel(vol_frame, text=f"{self.vol_var.get():.1f}%")
        self.vol_label.pack()

        self.fade_toggle = ctk.CTkSwitch(self, text="Enable Fades",
                                         command=self.on_ui_fades_toggle)
        if self.settings.get("fades_enabled"):
            self.fade_toggle.select()
        else:
            self.fade_toggle.deselect()
        self.fade_toggle.pack(pady=10)

        self.fade_down_var = ctk.DoubleVar(value=self.settings.get("fade_down"))
        fade_down_frame = ctk.CTkFrame(self, fg_color="transparent")
        fade_down_frame.pack(pady=5)
        ctk.CTkLabel(fade_down_frame, text="Fade Down (s)").pack()
        self.fade_down_slider = ctk.CTkSlider(
            fade_down_frame, from_=0.1, to=2.0,
            variable=self.fade_down_var,
            command=lambda v: self.on_ui_fade_down_change()
        )
        self.fade_down_slider.pack(pady=5)
        self.fade_down_label = ctk.CTkLabel(fade_down_frame,
                                            text=f"{self.fade_down_var.get():.2f}s")
        self.fade_down_label.pack()

        self.fade_up_var = ctk.DoubleVar(value=self.settings.get("fade_up"))
        fade_up_frame = ctk.CTkFrame(self, fg_color="transparent")
        fade_up_frame.pack(pady=5)
        ctk.CTkLabel(fade_up_frame, text="Fade Up (s)").pack()
        self.fade_up_slider = ctk.CTkSlider(
            fade_up_frame, from_=0.1, to=2.0,
            variable=self.fade_up_var,
            command=lambda v: self.on_ui_fade_up_change()
        )
        self.fade_up_slider.pack(pady=5)
        self.fade_up_label = ctk.CTkLabel(fade_up_frame,
                                          text=f"{self.fade_up_var.get():.2f}s")
        self.fade_up_label.pack()

        if not self.settings.get("fades_enabled"):
            self.fade_down_slider.configure(state="disabled")
            self.fade_up_slider.configure(state="disabled")

        self.start_btn = ctk.CTkButton(self, text="Start Monitoring",
                                       command=self.start_monitor)
        self.start_btn.pack(pady=10)

        self.stop_btn = ctk.CTkButton(self, text="Stop Monitoring",
                                      command=self.stop_monitor,
                                      state="disabled")
        self.stop_btn.pack()

        self.test_btn = ctk.CTkButton(self, text="Test",
                                      command=self.run_test_duck)
        self.test_btn.pack(pady=20)

        self.controller.refresh_from_settings()
        self.update_ui_from_settings()

        if self.settings.get("auto_start_monitoring"):
            self.start_monitor()

    def apply_theme(self):
        theme = self.settings.get("theme")
        if theme == "light":
            ctk.set_appearance_mode("Light")
        elif theme == "dark":
            ctk.set_appearance_mode("Dark")
        else:
            ctk.set_appearance_mode("System")

    def open_settings(self):
        SettingsWindow(self, self.settings, self.controller)

    def update_ui_from_settings(self):
        self.vol_var.set(self.settings.get("mute_volume"))
        self.fade_down_var.set(self.settings.get("fade_down"))
        self.fade_up_var.set(self.settings.get("fade_up"))
        self.vol_label.configure(text=f"{self.vol_var.get():.1f}%")
        self.fade_down_label.configure(text=f"{self.fade_down_var.get():.2f}s")
        self.fade_up_label.configure(text=f"{self.fade_up_var.get():.2f}s")

    def on_ui_mute_change(self):
        value = self.vol_var.get()
        self.vol_label.configure(text=f"{value:.1f}%")
        self.settings.set("mute_volume", float(value))
        self.controller.refresh_from_settings()

    def on_ui_fades_toggle(self):
        enabled = self.fade_toggle.get() == 1
        self.settings.set("fades_enabled", enabled)
        self.controller.refresh_from_settings()
        state = "normal" if enabled else "disabled"
        self.fade_down_slider.configure(state=state)
        self.fade_up_slider.configure(state=state)

    def on_ui_fade_down_change(self):
        value = self.fade_down_var.get()
        self.fade_down_label.configure(text=f"{value:.2f}s")
        self.settings.set("fade_down", float(value))
        self.controller.refresh_from_settings()

    def on_ui_fade_up_change(self):
        value = self.fade_up_var.get()
        self.fade_up_label.configure(text=f"{value:.2f}s")
        self.settings.set("fade_up", float(value))
        self.controller.refresh_from_settings()

    def start_monitor(self):
        if self.monitor_thread and self.monitor_thread.is_alive():
            return
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.status_label.configure(text="Starting…", text_color="#ffff00")

        def run():
            asyncio.run(self.controller.start(self.update_status, self.handle_disconnect))

        self.monitor_thread = threading.Thread(target=run, daemon=True)
        self.monitor_thread.start()

    def stop_monitor(self):
        self.controller.stop()
        self.status_label.configure(text="Stopped", text_color="#ff4444")
        self.start_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")

    def run_test_duck(self):
        def run():
            asyncio.run(self.controller.test_duck())
        threading.Thread(target=run, daemon=True).start()

    def update_status(self, msg):
        color = "#00ff00" if "Connected" in msg else "#ff4444"
        self.status_label.configure(text=msg, text_color=color)

    def handle_disconnect(self):
        if self.settings.get("auto_stop_on_disconnect"):
            self.controller.stop()
            self.status_label.configure(text="Lost connection", text_color="#ff4444")
            self.start_btn.configure(state="normal")
            self.stop_btn.configure(state="disabled")
        else:
            self.status_label.configure(text="Lost connection", text_color="#ff4444")

    def on_close(self):
        self.controller.stop()
        self.destroy()


if __name__ == "__main__":
    settings = SettingsManager()
    app = DuckingUI(settings)
    app.protocol("WM_DELETE_WINDOW", app.on_close)
    app.mainloop()
