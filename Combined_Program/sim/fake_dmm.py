"""
Fake SCPI Digital Multimeter Simulator.

Unlike the printer/hub sims, real DMMs are addressed over the network (see
sender_components/dmm_manager.py: PyVISA TCPIP resource, falling back to a
raw `TCPIP::<ip>::5025::SOCKET`) rather than a serial COM port — so there is
no virtual-serial-port pair here. Instead this binds to a loopback address
whose last octet stands in for the "port number", and the real app is
pointed at it by typing "127.0.0" into the existing DMM IP Prefix field in
the G-Code Sender's DMM panel — no seed-control-center code changes required.

dmm_manager.py's DMM_LIST hard-codes six instruments at IDs 20-25 (VINV,
VINP, IINP, VSYS, SAUX, SINV); run one instance of this script per ID to
connect all six. See launchers/{linux,windows}/run_fake_dmm_simulator.*
for a script that launches all six at once.

Run with: python sim/fake_dmm.py --id 20 --name VINV
"""
import argparse
import os
import sys
import time
import queue

import tkinter as tk
from tkinter import ttk

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from dmm_server import DmmServer

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


class FakeDmmApp:
    def __init__(self, root: tk.Tk, host: str, port: int, name: str):
        self.root = root
        self.host = host
        self.port = port
        self.name = name
        root.title(f'Fake DMM Simulator — {name}  ({host}:{port})')
        root.geometry('520x620')
        root.configure(bg=BG)
        root.resizable(False, True)

        self._build_ui()

        self.server = DmmServer(host=host, port=port)
        self.server.start()

        root.after(100, self._update_loop)

    # ── UI construction ──────────────────────────────────────────────────────

    def _build_ui(self):
        title_frame = tk.Frame(self.root, bg=PANEL_BG, pady=8)
        title_frame.pack(fill=tk.X)
        tk.Label(title_frame, text='SCPI DMM  SIMULATOR',
                 font=FONT_TITLE, bg=PANEL_BG, fg=ACCENT).pack()
        tk.Label(title_frame, text=f'{self.name}  |  {self.host}:{self.port}  (SCPI/TCP)',
                 font=FONT_BODY, bg=PANEL_BG, fg=TEXT_DIM).pack()

        status_frame = tk.Frame(self.root, bg=BG, pady=4)
        status_frame.pack(fill=tk.X, padx=12)
        self.status_dot = tk.Canvas(status_frame, width=14, height=14, bg=BG, highlightthickness=0)
        self.status_dot.pack(side=tk.LEFT)
        self._dot = self.status_dot.create_oval(2, 2, 12, 12, fill='grey', outline='')
        self.status_lbl = tk.Label(status_frame, text='Waiting for connection...',
                                   font=FONT_BODY, bg=BG, fg=TEXT_DIM)
        self.status_lbl.pack(side=tk.LEFT, padx=6)

        # ── Live reading ─────────────────────────────────────────────────────
        reading_outer = tk.Frame(self.root, bg=BG, pady=6)
        reading_outer.pack(fill=tk.X, padx=12)
        tk.Label(reading_outer, text='LAST AVERAGED READING', font=FONT_BODY,
                 bg=BG, fg=TEXT_DIM).pack(anchor='w')
        self.reading_canvas = tk.Canvas(reading_outer, height=56, bg=BG2,
                                        highlightthickness=1, highlightbackground='#333')
        self.reading_canvas.pack(fill=tk.X)
        self._reading_text = self.reading_canvas.create_text(
            10, 28, anchor='w', text='---', font=FONT_LG, fill=TEXT)
        self.reading_canvas.bind('<Configure>', lambda e: None)

        # ── Controls ─────────────────────────────────────────────────────────
        ctrl_frame = tk.LabelFrame(self.root, text='Simulated Signal', font=FONT_BODY,
                                   bg=BG, fg=TEXT_DIM, bd=1, relief=tk.GROOVE)
        ctrl_frame.pack(fill=tk.X, padx=12, pady=4)

        self.base_var = tk.DoubleVar(value=5.000)
        self._make_slider_row(ctrl_frame, 0, 'Base Value', 0.0, 400.0, self.base_var,
                               lambda v: self._on_base_change(), resolution=0.01)

        self.noise_var = tk.DoubleVar(value=0.5)
        self._make_slider_row(ctrl_frame, 1, 'Noise (%)', 0.0, 10.0, self.noise_var,
                               lambda v: self._on_noise_change(), resolution=0.01)

        # ── Instrument state ─────────────────────────────────────────────────
        state_outer = tk.LabelFrame(self.root, text='Instrument State', font=FONT_BODY,
                                    bg=BG, fg=TEXT_DIM, bd=1, relief=tk.GROOVE)
        state_outer.pack(fill=tk.X, padx=12, pady=4)

        grid = tk.Frame(state_outer, bg=BG2)
        grid.pack(fill=tk.X, padx=4, pady=4)

        self._mode_var = tk.StringVar(value='---')
        self._samp_var = tk.StringVar(value='---')
        self._aver_var = tk.StringVar(value='---')
        rows = [
            ('Mode (CONF):', self._mode_var),
            ('Sample Count Target:', self._samp_var),
            ('Averaging Count:', self._aver_var),
        ]
        for i, (label, var) in enumerate(rows):
            tk.Label(grid, text=label, font=FONT_BODY, bg=BG2, fg=TEXT_DIM,
                     anchor='w', width=20).grid(row=i, column=0, sticky='w', padx=(8, 2), pady=2)
            tk.Label(grid, textvariable=var, font=FONT_MONO, bg=BG2, fg=TEXT,
                     anchor='w').grid(row=i, column=1, sticky='w', padx=(2, 8))

        # ── Command log ──────────────────────────────────────────────────────
        log_outer = tk.LabelFrame(self.root, text='SCPI Command Log', font=FONT_BODY,
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

        self.log_text.tag_configure('rx', foreground='#64B5F6')
        self.log_text.tag_configure('tx', foreground=GREEN)
        self.log_text.tag_configure('info', foreground=TEXT_DIM)

    def _make_slider_row(self, parent, row, label, lo, hi, var, on_change, resolution=None):
        tk.Label(parent, text=label, font=FONT_BODY, bg=BG, fg=TEXT_DIM,
                 width=14, anchor='w').grid(row=row, column=0, sticky='w', padx=8, pady=4)
        slider = tk.Scale(parent, from_=lo, to=hi, resolution=resolution or (hi - lo) / 1000,
                          orient=tk.HORIZONTAL, variable=var,
                          bg=BG, fg=TEXT, troughcolor=BG2, highlightthickness=0,
                          length=220, showvalue=False, command=lambda _: on_change(var.get()))
        slider.grid(row=row, column=1, padx=6)
        entry = tk.Entry(parent, textvariable=var, font=FONT_MONO,
                         bg=BG2, fg=TEXT, insertbackground=TEXT,
                         width=8, relief=tk.FLAT)
        entry.grid(row=row, column=2, padx=(4, 8))
        entry.bind('<Return>', lambda _: on_change(var.get()))
        entry.bind('<FocusOut>', lambda _: on_change(var.get()))

    # ── update loop ──────────────────────────────────────────────────────────

    def _update_loop(self):
        self._drain_log()
        self._refresh_state()
        self._refresh_status()
        self.root.after(100, self._update_loop)

    def _drain_log(self):
        count = 0
        try:
            while count < 20:
                msg = self.server.log_queue.get_nowait()
                tag = 'rx' if msg.startswith('RX:') else 'tx' if msg.startswith('TX:') else 'info'
                ts = time.strftime('%H:%M:%S')
                self.log_text.configure(state=tk.NORMAL)
                self.log_text.insert(tk.END, f'[{ts}] {msg}\n', tag)
                self.log_text.see(tk.END)
                self.log_text.configure(state=tk.DISABLED)
                count += 1
        except queue.Empty:
            pass

    def _refresh_state(self):
        s = self.server.state
        self._mode_var.set(s.mode)
        self._samp_var.set(str(s.sample_target))
        self._aver_var.set(f'{s.current_aver_count()} / {s.sample_target}')
        if s.current_aver_count() >= s.sample_target and s._trigger_time is not None:
            self.reading_canvas.itemconfig(self._reading_text, text=f'{s.averaged_reading():.6f}')

    def _refresh_status(self):
        if self.server.connected:
            self.status_dot.itemconfig(self._dot, fill=GREEN)
            self.status_lbl.config(text=f'Connected  —  {self.host}:{self.port}', fg=GREEN)
        else:
            self.status_dot.itemconfig(self._dot, fill=AMBER)
            self.status_lbl.config(text='Waiting for connection...', fg=AMBER)

    # ── control callbacks ────────────────────────────────────────────────────

    def _on_base_change(self):
        try:
            self.server.state.base_value = float(self.base_var.get())
        except (tk.TclError, ValueError):
            pass

    def _on_noise_change(self):
        try:
            self.server.state.noise_pct = float(self.noise_var.get())
        except (tk.TclError, ValueError):
            pass


def main():
    parser = argparse.ArgumentParser(description='Fake SCPI DMM Simulator')
    parser.add_argument('--id', type=int, default=50,
                        help='Last IP octet to bind on 127.0.0.<id> (default: 50)')
    parser.add_argument('--host', default=None,
                        help='Full bind address, overrides --id (e.g. 192.168.0.20)')
    parser.add_argument('--tcp-port', type=int, default=5025,
                        help='TCP port to listen on (dmm_manager.py always uses 5025)')
    parser.add_argument('--name', default='DMM', help='Display label only')
    args = parser.parse_args()

    host = args.host or f'127.0.0.{args.id}'

    root = tk.Tk()
    app = FakeDmmApp(root, host, args.tcp_port, args.name)

    def on_close():
        app.server.stop()
        root.destroy()

    root.protocol('WM_DELETE_WINDOW', on_close)
    root.mainloop()


if __name__ == '__main__':
    main()
