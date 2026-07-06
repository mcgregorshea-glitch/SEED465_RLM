"""
Fake RLM Vivigo Hub Simulator
Speaks the COBS/CRC-32 binary protocol that backend.py expects.
Run with: python sim/fake_hub.py --port COM100
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import queue
import sys
import os
import argparse
import time
import math
import json

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from hub_server import HubServer
from hub_logic import (
    STATE_NAMES, STATE_COLORS, STATE_READY,
    SineParams, get_sine_params, set_sine_params,
    get_all_sine_params, load_sine_params, reset_sine_params, _SIN,
)


BG        = '#1a1a2e'
BG2       = '#16213e'
PANEL_BG  = '#0f3460'
ACCENT    = '#e94560'
TEXT      = '#e0e0e0'
TEXT_DIM  = '#888888'
GREEN     = '#4CAF50'
AMBER     = '#FF9800'
RED_C     = '#F44336'

FONT_BODY  = ('Consolas', 10)
FONT_MONO  = ('Consolas', 11)
FONT_LG    = ('Consolas', 14, 'bold')
FONT_TITLE = ('Consolas', 18, 'bold')


class TelemetryRow:
    """One labelled value row in the telemetry panel."""
    def __init__(self, parent, label: str, unit: str, row: int, key: str,
                 open_editor_cb):
        tk.Label(parent, text=label, font=FONT_BODY, bg=BG2, fg=TEXT_DIM,
                 anchor='w', width=14).grid(row=row, column=0, sticky='w', padx=(8, 2), pady=1)
        self.var = tk.StringVar(value='---')
        tk.Label(parent, textvariable=self.var, font=FONT_MONO, bg=BG2, fg=TEXT,
                 anchor='e', width=10).grid(row=row, column=1, sticky='e', padx=(2, 2))
        tk.Label(parent, text=unit, font=FONT_BODY, bg=BG2, fg=TEXT_DIM,
                 anchor='w', width=5).grid(row=row, column=2, sticky='w', padx=(2, 4))
        tk.Button(parent, text='⚙', font=FONT_BODY, bg=PANEL_BG, fg=TEXT_DIM,
                  relief=tk.FLAT, cursor='hand2', width=2,
                  command=lambda: open_editor_cb(key, label)
                  ).grid(row=row, column=3, sticky='w', padx=(0, 8))

    def update(self, value: float, fmt: str = '.3f'):
        self.var.set(f'{value:{fmt}}')


class WaveformEditorDialog:
    """Per-variable sine waveform editor (amplitude, frequency, phase)."""

    # Maps key → editor instance so only one editor per variable exists.
    _open_editors: dict = {}

    def __init__(self, parent, key: str, label: str):
        if key in self._open_editors:
            self._open_editors[key]._win.lift()
            self._open_editors[key]._win.focus_force()
            return

        self._key = key
        self._label = label
        self._default_amp, self._default_period = _SIN[key]

        self._win = tk.Toplevel(parent)
        self._win.title(f'Waveform Editor — {label}')
        self._win.configure(bg=BG)
        self._win.resizable(False, False)
        self._win.protocol('WM_DELETE_WINDOW', self._on_close)
        self._open_editors[key] = self

        self._build(self._win)
        self._refresh_canvas()

    # ── UI construction ───────────────────────────────────────────────────────

    def _build(self, win):
        p = get_sine_params(self._key)

        # Canvas preview
        canvas_frame = tk.Frame(win, bg=BG, pady=6)
        canvas_frame.pack(fill=tk.X, padx=12)
        self._canvas = tk.Canvas(canvas_frame, width=300, height=80,
                                 bg='#0d0d0d', highlightthickness=1,
                                 highlightbackground='#333')
        self._canvas.pack()

        # Slider rows
        sliders_frame = tk.Frame(win, bg=BG)
        sliders_frame.pack(fill=tk.X, padx=12, pady=4)

        amp_max = self._default_amp * 2
        self._amp_var, _ = self._make_slider_row(
            sliders_frame, 0, 'Amplitude', 0.0, amp_max, p.amplitude, 4)

        default_freq = 1.0 / self._default_period
        freq_init = 1.0 / p.period if p.period > 0 else default_freq
        self._freq_var, _ = self._make_slider_row(
            sliders_frame, 1, 'Frequency (Hz)', 0.0, default_freq * 2, freq_init, 3)

        phase_deg = math.degrees(p.phase)
        self._phase_var, _ = self._make_slider_row(
            sliders_frame, 2, 'Phase (°)', -180.0, 180.0, phase_deg, 1)

        # Buttons
        btn_frame = tk.Frame(win, bg=BG, pady=6)
        btn_frame.pack(fill=tk.X, padx=12)

        for text, cmd in [
            ('Reset to Default', self._on_reset),
            ('Close',            self._on_close),
        ]:
            tk.Button(btn_frame, text=text, font=FONT_BODY, bg=PANEL_BG, fg=TEXT,
                      relief=tk.FLAT, padx=8, pady=3, cursor='hand2',
                      command=cmd).pack(side=tk.LEFT, padx=4)

    def _make_slider_row(self, parent, row, label, lo, hi, init, digits):
        tk.Label(parent, text=label, font=FONT_BODY, bg=BG, fg=TEXT_DIM,
                 width=16, anchor='w').grid(row=row, column=0, sticky='w', pady=3)

        var = tk.DoubleVar(value=init)
        slider = tk.Scale(parent, from_=lo, to=hi, resolution=(hi - lo) / 1000,
                          orient=tk.HORIZONTAL, variable=var,
                          bg=BG, fg=TEXT, troughcolor=BG2, highlightthickness=0,
                          length=180, showvalue=False,
                          command=lambda _: self._on_slider_change())
        slider.grid(row=row, column=1, padx=6)

        entry = tk.Entry(parent, textvariable=var, font=FONT_MONO,
                         bg=BG2, fg=TEXT, insertbackground=TEXT,
                         width=8, relief=tk.FLAT)
        entry.grid(row=row, column=2, padx=4)
        entry.bind('<Return>', lambda _: self._on_slider_change())
        entry.bind('<FocusOut>', lambda _: self._on_slider_change())

        return var, slider

    # ── callbacks ─────────────────────────────────────────────────────────────

    def _on_slider_change(self):
        try:
            amp = float(self._amp_var.get())
            freq = float(self._freq_var.get())
            phase_deg = float(self._phase_var.get())
        except (tk.TclError, ValueError):
            return

        period = 1.0 / freq if freq > 0 else 100.0
        set_sine_params(self._key, SineParams(
            amplitude=amp,
            period=period,
            phase=math.radians(phase_deg),
        ))
        self._refresh_canvas()

    def _on_reset(self):
        amp, period = _SIN[self._key]
        self._amp_var.set(round(amp, 6))
        self._freq_var.set(round(1.0 / period, 6))
        self._phase_var.set(0.0)
        self._on_slider_change()

    def _sync_sliders_from_params(self):
        p = get_sine_params(self._key)
        self._amp_var.set(round(p.amplitude, 6))
        self._freq_var.set(round(1.0 / p.period if p.period > 0 else 0.01, 6))
        self._phase_var.set(round(math.degrees(p.phase), 2))
        self._refresh_canvas()

    def _on_close(self):
        self._open_editors.pop(self._key, None)
        self._win.destroy()

    # ── canvas preview ────────────────────────────────────────────────────────

    def _refresh_canvas(self):
        c = self._canvas
        c.delete('all')
        w, h = 300, 80
        mid = h // 2

        # Baseline
        c.create_line(0, mid, w, mid, fill='#333', dash=(4, 4))

        try:
            amp = float(self._amp_var.get())
            freq = float(self._freq_var.get())
            phase_deg = float(self._phase_var.get())
        except (tk.TclError, ValueError):
            return

        if freq <= 0:
            return

        phase = math.radians(phase_deg)
        amp_max = self._default_amp * 2
        scale_y = (h / 2 - 6) / amp_max
        display_s = self._default_period  # 1 cycle at default frequency

        points = []
        for px in range(w):
            t = (px / w) * display_s
            y = amp * math.sin(2 * math.pi * freq * t + phase)
            points.extend([px, mid - y * scale_y])

        if len(points) >= 4:
            c.create_line(points, fill=ACCENT, width=2, smooth=True)


class FakeHubApp:
    def __init__(self, root: tk.Tk, port: str):
        self.root = root
        self.port = port
        root.title(f'Fake RLM Vivigo Hub  —  {port}')
        root.geometry('580x720')
        root.configure(bg=BG)
        root.resizable(False, True)

        self._build_ui()

        self.server = HubServer(port=port)
        self.server.start()

        root.after(100, self._update_loop)

    # ── UI construction ───────────────────────────────────────────────────────

    def _build_ui(self):
        # ── Title bar ────────────────────────────────────────────────────────
        title_frame = tk.Frame(self.root, bg=PANEL_BG, pady=8)
        title_frame.pack(fill=tk.X)
        tk.Label(title_frame, text='RLM VIVIGO HUB  SIMULATOR',
                 font=FONT_TITLE, bg=PANEL_BG, fg=ACCENT).pack()
        tk.Label(title_frame, text=f'Port: {self.port}  |  115200 baud',
                 font=FONT_BODY, bg=PANEL_BG, fg=TEXT_DIM).pack()

        # ── Connection status ─────────────────────────────────────────────────
        status_frame = tk.Frame(self.root, bg=BG, pady=4)
        status_frame.pack(fill=tk.X, padx=12)
        self.status_dot = tk.Canvas(status_frame, width=14, height=14, bg=BG, highlightthickness=0)
        self.status_dot.pack(side=tk.LEFT)
        self._dot = self.status_dot.create_oval(2, 2, 12, 12, fill='grey', outline='')
        self.status_lbl = tk.Label(status_frame, text='Waiting for connection...',
                                   font=FONT_BODY, bg=BG, fg=TEXT_DIM)
        self.status_lbl.pack(side=tk.LEFT, padx=6)

        # ── State indicator ───────────────────────────────────────────────────
        state_outer = tk.Frame(self.root, bg=BG, pady=6)
        state_outer.pack(fill=tk.X, padx=12)
        tk.Label(state_outer, text='HUB STATE', font=FONT_BODY,
                 bg=BG, fg=TEXT_DIM).pack(anchor='w')
        self.state_canvas = tk.Canvas(state_outer, height=52, bg=BG2,
                                      highlightthickness=1, highlightbackground='#333')
        self.state_canvas.pack(fill=tk.X)
        self._state_rect  = self.state_canvas.create_rectangle(0, 0, 600, 52,
                                                                fill='#1a1a2e', outline='')
        self._state_label = self.state_canvas.create_text(290, 26, text='WPT Ready',
                                                           font=FONT_LG, fill=GREEN)
        self.state_canvas.bind('<Configure>',
            lambda e: self.state_canvas.coords(self._state_rect, 0, 0, e.width, e.height))

        # ── Controls ──────────────────────────────────────────────────────────
        ctrl_frame = tk.LabelFrame(self.root, text='Controls', font=FONT_BODY,
                                   bg=BG, fg=TEXT_DIM, bd=1, relief=tk.GROOVE)
        ctrl_frame.pack(fill=tk.X, padx=12, pady=4)

        self.btn_var = tk.BooleanVar(value=False)
        btn_check = tk.Checkbutton(ctrl_frame, text='Simulate Button Press  (V_BST: 40 V → 80 V)',
                                   variable=self.btn_var, command=self._on_button_toggle,
                                   font=FONT_BODY, bg=BG, fg=TEXT,
                                   selectcolor='#333', activebackground=BG,
                                   activeforeground=TEXT)
        btn_check.pack(anchor='w', padx=8, pady=4)

        profile_row = tk.Frame(ctrl_frame, bg=BG)
        profile_row.pack(anchor='w', padx=8, pady=(0, 6))
        tk.Label(profile_row, text='Telemetry Profile:', font=FONT_BODY,
                 bg=BG, fg=TEXT_DIM).pack(side=tk.LEFT, padx=(0, 6))
        tk.Button(profile_row, text='Save Profile…', font=FONT_BODY,
                  bg=PANEL_BG, fg=TEXT, relief=tk.FLAT, padx=8, pady=2,
                  cursor='hand2', command=self._on_save_profile).pack(side=tk.LEFT, padx=(0, 4))
        tk.Button(profile_row, text='Load Profile…', font=FONT_BODY,
                  bg=PANEL_BG, fg=TEXT, relief=tk.FLAT, padx=8, pady=2,
                  cursor='hand2', command=self._on_load_profile).pack(side=tk.LEFT)

        # ── Telemetry readout ─────────────────────────────────────────────────
        tlm_outer = tk.LabelFrame(self.root, text='Live Telemetry  (5 Hz)', font=FONT_BODY,
                                  bg=BG, fg=TEXT_DIM, bd=1, relief=tk.GROOVE)
        tlm_outer.pack(fill=tk.X, padx=12, pady=4)

        tlm_grid = tk.Frame(tlm_outer, bg=BG2)
        tlm_grid.pack(fill=tk.X, padx=4, pady=4)

        rows = [
            ('V_SYS',  'V',  'v_sys',  '.3f'),
            ('I_SYS',  'A',  'i_sys',  '.4f'),
            ('V_BST',  'V',  'v_bst',  '.2f'),
            ('I_BST',  'A',  'i_bst',  '.4f'),
            ('T_BST',  '°C', 't_bst',  '.1f'),
            ('T_INV',  '°C', 't_inv',  '.1f'),
            ('V_REC',  'V',  'v_rec',  '.3f'),
            ('T_REC',  '°C', 't_rec',  '.1f'),
        ]
        self._tlm_rows: dict[str, tuple] = {}
        for i, (label, unit, key, fmt) in enumerate(rows):
            row_w = TelemetryRow(tlm_grid, label, unit, i, key, self._open_waveform_editor)
            self._tlm_rows[key] = (row_w, fmt)

        # ── Command log ───────────────────────────────────────────────────────
        log_outer = tk.LabelFrame(self.root, text='Command Log', font=FONT_BODY,
                                  bg=BG, fg=TEXT_DIM, bd=1, relief=tk.GROOVE)
        log_outer.pack(fill=tk.BOTH, expand=True, padx=12, pady=4)

        log_frame = tk.Frame(log_outer, bg=BG)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        self.log_text = tk.Text(log_frame, font=FONT_BODY, bg='#0d0d0d', fg=TEXT,
                                state=tk.DISABLED, wrap=tk.WORD, relief=tk.FLAT,
                                insertbackground=TEXT)
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.log_text.tag_configure('cmd',  foreground='#64B5F6')
        self.log_text.tag_configure('err',  foreground=RED_C)
        self.log_text.tag_configure('info', foreground=TEXT_DIM)

    # ── update loop ───────────────────────────────────────────────────────────

    def _update_loop(self):
        self._drain_log()
        self._refresh_state()
        self._refresh_telemetry()
        self._refresh_status()
        self.root.after(100, self._update_loop)

    def _drain_log(self):
        count = 0
        try:
            while count < 20:
                msg = self.server.log_queue.get_nowait()
                tag = 'cmd' if msg.startswith('CMD:') else 'err' if 'error' in msg.lower() else 'info'
                ts = time.strftime('%H:%M:%S')
                self.log_text.configure(state=tk.NORMAL)
                self.log_text.insert(tk.END, f'[{ts}] {msg}\n', tag)
                self.log_text.see(tk.END)
                self.log_text.configure(state=tk.DISABLED)
                count += 1
        except queue.Empty:
            pass

    def _refresh_state(self):
        s = self.server.state.state
        name  = STATE_NAMES.get(s, '?')
        color = STATE_COLORS.get(s, '#888888')
        self.state_canvas.itemconfig(self._state_label, text=name, fill=color)

    def _refresh_telemetry(self):
        vals = self.server.state.get_real_values()
        for key, (row_w, fmt) in self._tlm_rows.items():
            if key in vals:
                row_w.update(vals[key], fmt)

    def _refresh_status(self):
        if self.server.connected:
            self.status_dot.itemconfig(self._dot, fill=GREEN)
            self.status_lbl.config(text=f'Connected  —  {self.port}', fg=GREEN)
        else:
            self.status_dot.itemconfig(self._dot, fill=AMBER)
            self.status_lbl.config(text='Waiting for connection...', fg=AMBER)

    # ── control callbacks ─────────────────────────────────────────────────────

    def _on_button_toggle(self):
        self.server.state.button_pressed = self.btn_var.get()

    def _open_waveform_editor(self, key: str, label: str):
        WaveformEditorDialog(self.root, key, label)

    def _on_save_profile(self):
        path = filedialog.asksaveasfilename(
            title='Save Telemetry Profile',
            defaultextension='.json',
            filetypes=[('JSON files', '*.json'), ('All files', '*.*')],
        )
        if not path:
            return
        data = {
            k: {'amplitude': v.amplitude, 'period': v.period, 'phase': v.phase}
            for k, v in get_all_sine_params().items()
        }
        try:
            with open(path, 'w') as f:
                json.dump(data, f, indent=2)
        except OSError as e:
            messagebox.showerror('Save Failed', str(e))

    def _on_load_profile(self):
        path = filedialog.askopenfilename(
            title='Load Telemetry Profile',
            filetypes=[('JSON files', '*.json'), ('All files', '*.*')],
        )
        if not path:
            return
        try:
            with open(path) as f:
                data = json.load(f)
            load_sine_params(data)
        except (OSError, json.JSONDecodeError, KeyError) as e:
            messagebox.showerror('Load Failed', str(e))
            return
        for k, editor in list(WaveformEditorDialog._open_editors.items()):
            if k in data:
                editor._sync_sliders_from_params()


# ── entry point ───────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description='Fake RLM Vivigo Hub Simulator')
    parser.add_argument('--port', default='COM100', help='Virtual COM port to listen on')
    args = parser.parse_args()

    root = tk.Tk()
    app = FakeHubApp(root, args.port)

    def on_close():
        app.server.stop()
        root.destroy()

    root.protocol('WM_DELETE_WINDOW', on_close)
    root.mainloop()


if __name__ == '__main__':
    main()
