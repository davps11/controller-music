import asyncio
import threading
import json
import websockets
import ctypes
import customtkinter as ctk
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume


class SpotifyVolumeController:
    def __init__(self):
        self._lock = asyncio.Lock()

    def _find_spotify(self):
        sessions = AudioUtilities.GetAllSessions()
        for s in sessions:
            if s.Process and s.Process.name().lower() == "spotify.exe":
                return s._ctl.QueryInterface(ISimpleAudioVolume)
        return None

    def get_volume(self):
        try:
            ctypes.windll.ole32.CoInitialize(0)
            vol = self._find_spotify()
            if vol:
                return vol.GetMasterVolume() * 100
        finally:
            ctypes.windll.ole32.CoUninitialize()
        return 100

    def set_volume(self, percent):
        try:
            ctypes.windll.ole32.CoInitialize(0)
            vol = self._find_spotify()
            if vol:
                vol.SetMasterVolume(max(0, min(1, percent / 100)), None)
        finally:
            ctypes.windll.ole32.CoUninitialize()

    async def fade_to(self, target, duration):
        async with self._lock:
            start = self.get_volume()
            steps = max(1, int(duration * 60))
            step = (target - start) / steps
            delay = duration / steps
            for i in range(steps):
                self.set_volume(start + step * (i + 1))
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
    def __init__(self, uri, on_start, on_end, on_status):
        self.uri = uri
        self.state = TransmissionState()
        self.on_start = on_start
        self.on_end = on_end
        self.on_status = on_status
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
            except Exception as e:
                if not self.running:
                    break
                self.on_status(f"Disconnected: {e}")
                await asyncio.sleep(2)

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
    def __init__(self, uri="ws://localhost:49080/ws"):
        self.spotify = SpotifyVolumeController()
        self.uri = uri
        self.target = 10
        self.fade_down = 0.4
        self.fade_up = 0.4
        self.original = 100
        self.monitor = None

    async def start(self, status_callback):
        self.monitor = TrackAudioMonitor(
            self.uri,
            self.lower_volume,
            self.restore_volume,
            status_callback
        )
        await self.monitor.connect()

    def force_restore(self):
        self.spotify.set_volume(self.original)

    def stop(self):
        if self.monitor:
            self.monitor.stop()
        self.force_restore()

    async def lower_volume(self):
        self.original = self.spotify.get_volume()
        await self.spotify.fade_to(self.target, self.fade_down)

    async def restore_volume(self):
        await self.spotify.fade_to(self.original, self.fade_up)


class DuckingUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Controller Music")
        self.geometry("420x500")
        self.resizable(False, False)
        self.configure(fg_color="#1a1a1a")

        self.controller = AudioDuckController()
        self.monitor_thread = None

        ctk.CTkLabel(self, text="🎵 Controller Music",
                     font=ctk.CTkFont(size=26, weight="bold")).pack(pady=10)

        self.status_label = ctk.CTkLabel(self, text="Ready",
                                         font=ctk.CTkFont(size=14, weight="bold"),
                                         text_color="#00aaff")
        self.status_label.pack(pady=5)

        self.vol_var = ctk.DoubleVar(value=5)
        vol_frame = ctk.CTkFrame(self, fg_color="transparent")
        vol_frame.pack(pady=5)
        ctk.CTkLabel(vol_frame, text="Mute Volume (%)").pack()
        self.vol_slider = ctk.CTkSlider(vol_frame, from_=0, to=100,
                                        variable=self.vol_var,
                                        command=lambda v: self.update_values())
        self.vol_slider.pack(pady=5)
        self.vol_label = ctk.CTkLabel(vol_frame, text=f"{self.vol_var.get():.1f}%")
        self.vol_label.pack()

        self.fade_down_var = ctk.DoubleVar(value=0.1)
        fade_down_frame = ctk.CTkFrame(self, fg_color="transparent")
        fade_down_frame.pack(pady=5)
        ctk.CTkLabel(fade_down_frame, text="Fade Down (s)").pack()
        self.fade_down_slider = ctk.CTkSlider(fade_down_frame, from_=0.1, to=2.0,
                                              variable=self.fade_down_var,
                                              command=lambda v: self.update_values())
        self.fade_down_slider.pack(pady=5)
        self.fade_down_label = ctk.CTkLabel(fade_down_frame,
                                            text=f"{self.fade_down_var.get():.2f}s")
        self.fade_down_label.pack()

        self.fade_up_var = ctk.DoubleVar(value=0.1)
        fade_up_frame = ctk.CTkFrame(self, fg_color="transparent")
        fade_up_frame.pack(pady=5)
        ctk.CTkLabel(fade_up_frame, text="Fade Up (s)").pack()
        self.fade_up_slider = ctk.CTkSlider(fade_up_frame, from_=0.1, to=2.0,
                                            variable=self.fade_up_var,
                                            command=lambda v: self.update_values())
        self.fade_up_slider.pack(pady=5)
        self.fade_up_label = ctk.CTkLabel(fade_up_frame,
                                          text=f"{self.fade_up_var.get():.2f}s")
        self.fade_up_label.pack()

        self.start_btn = ctk.CTkButton(self, text="Start Monitoring",
                                       command=self.start_monitor)
        self.start_btn.pack(pady=10)

        self.stop_btn = ctk.CTkButton(self, text="Stop Monitoring",
                                      command=self.stop_monitor,
                                      state="disabled")
        self.stop_btn.pack()

        self.update_values()

    def update_values(self):
        self.controller.target = self.vol_var.get()
        self.controller.fade_down = self.fade_down_var.get()
        self.controller.fade_up = self.fade_up_var.get()
        self.vol_label.configure(text=f"{self.vol_var.get():.1f}%")
        self.fade_down_label.configure(text=f"{self.fade_down_var.get():.2f}s")
        self.fade_up_label.configure(text=f"{self.fade_up_var.get():.2f}s")

    def start_monitor(self):
        if self.monitor_thread and self.monitor_thread.is_alive():
            return
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.status_label.configure(text="Starting…", text_color="#ffff00")

        def run():
            asyncio.run(self.controller.start(self.update_status))

        self.monitor_thread = threading.Thread(target=run, daemon=True)
        self.monitor_thread.start()

    def stop_monitor(self):
        self.controller.stop()
        self.status_label.configure(text="Stopped", text_color="#ff4444")
        self.start_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")

    def update_status(self, msg):
        color = "#00ff00" if "Connected" in msg else "#ff4444"
        self.status_label.configure(text=msg, text_color=color)

    def on_close(self):
        self.controller.stop()
        self.destroy()


if __name__ == "__main__":
    ctk.set_appearance_mode("Dark")
    app = DuckingUI()
    app.protocol("WM_DELETE_WINDOW", app.on_close)
    app.mainloop()
