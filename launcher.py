"""Standard-library-only desktop setup and launcher; no financial files are read here."""
from __future__ import annotations

import hashlib
import os
from pathlib import Path
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import webbrowser

CODE = Path(__file__).resolve().parent
HIDDEN = {"creationflags": subprocess.CREATE_NO_WINDOW} if os.name == "nt" else {}


def enable_high_dpi():
    """Opt in before Tk creates any windows, so Windows does not bitmap-stretch them."""
    if os.name == "nt":
        import ctypes
        try:
            ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
        except (AttributeError, OSError):
            try:
                ctypes.windll.shcore.SetProcessDpiAwareness(1)
            except (AttributeError, OSError):
                ctypes.windll.user32.SetProcessDPIAware()


def configured_home(code=CODE):
    """Read configuration only, never inspect the selected private directory."""
    if os.environ.get("HARCAMA_HOME"):
        return os.environ["HARCAMA_HOME"]
    cfg = code / "home.txt"
    if cfg.exists():
        value = next((s.strip() for s in cfg.read_text(encoding="utf-8-sig").splitlines()
                      if s.strip() and not s.lstrip().startswith("#")), "")
        if value:
            return os.path.expandvars(value)
    return str(Path.home() / "Harcamalarim")


def validate_home(value, code=CODE):
    if not value.strip():
        raise ValueError("Verilerinizi saklamak için bir klasör seçin.")
    path = Path(os.path.expandvars(value.strip())).expanduser()
    if not path.is_absolute():
        raise ValueError("Tam bir klasör yolu seçin; Klasör seç düğmesini kullanabilirsiniz.")
    path = path.resolve()
    if path == Path(path.anchor) or path.is_relative_to(code.resolve()):
        raise ValueError("Kod klasörünün dışında, size ait bir veri klasörü seçin.")
    return path


def save_home(path, code=CODE):
    # Atomic update: a failed write must not truncate the existing configuration.
    cfg = code / "home.txt"
    temp = code / "home.txt.tmp"
    temp.write_text(f"# Private data directory; do not commit.\n{path}\n", encoding="utf-8")
    temp.replace(cfg)


def environment_python(code=CODE):
    return code / ".venv" / ("Scripts/python.exe" if os.name == "nt" else "bin/python")


def prepare_environment(report, code=CODE):
    """Install once, then re-check only when requirements or installed packages change."""
    python = environment_python(code)
    env = dict(os.environ, PYTHONIOENCODING="utf-8", PIP_DISABLE_PIP_VERSION_CHECK="1")

    def run(args, timeout):
        result = subprocess.run(args, cwd=code, env=env, capture_output=True, text=True,
                                encoding="utf-8", errors="replace", timeout=timeout, **HIDDEN)
        if result.returncode:
            # Installer output only; application logs may contain financial data.
            raise RuntimeError((result.stdout + "\n" + result.stderr).strip()[-5000:])

    if not python.exists():
        report("Uygulama ortamı hazırlanıyor…")
        run([sys.executable, "-m", "venv", str(code / ".venv")], 180)
    requirements = code / "requirements.txt"
    digest = hashlib.sha256(requirements.read_bytes()).hexdigest()
    marker = code / ".venv" / "fintracker-requirements.sha256"
    if marker.exists() and marker.read_text(encoding="ascii") == digest:
        try:
            run([str(python), "-c", "import pdfplumber, openpyxl"], 30)
            report("Gerekli bileşenler hazır.")
            return python
        except RuntimeError:
            pass
    report("Gerekli bileşenler indiriliyor… İlk kurulum için internet gerekir.")
    run([str(python), "-m", "pip", "install", "-r", str(requirements)], 600)
    run([str(python), "-c", "import pdfplumber, openpyxl"], 30)
    marker.write_text(digest, encoding="ascii")
    return python


class Launcher:
    def __init__(self, root):
        self.root = root
        self.events = queue.Queue()
        self.process = None
        self.url = None
        self.busy = False
        self.root.title("Fintracker • Başlangıç")
        scale = root.winfo_fpixels("1i") / 96
        px = lambda n: round(n * scale)
        self.root.geometry(f"{px(760)}x{px(650)}")
        self.root.minsize(px(700), px(620))
        self.root.configure(bg="#f4f6f3")
        style = ttk.Style(root)
        style.theme_use("clam")
        style.configure("TFrame", background="#f4f6f3")
        style.configure("Card.TFrame", background="white")
        style.configure("TLabel", background="#f4f6f3", foreground="#193b30", font=("Segoe UI", 11))
        style.configure("Brand.TLabel", font=("Segoe UI", 19, "bold"), foreground="#24372f", padding=0)
        style.configure("BrandDot.TLabel", font=("Segoe UI", 19, "bold"), foreground="#d18a24", padding=0)
        style.configure("Muted.TLabel", foreground="#65746b")
        style.configure("Card.TLabel", background="white")
        style.configure("Caption.TLabel", font=("Segoe UI", 9, "bold"), foreground="#56816c")
        style.configure("TButton", font=("Segoe UI", 11, "bold"), padding=(px(18), px(12)),
                        background="#e7eee8", foreground="#254d3e", borderwidth=0,
                        relief="flat", focusthickness=2, focuscolor="#80aa94")
        style.map("TButton", background=[("disabled", "#edf1ed"), ("active", "#d7e6dc")],
                  foreground=[("disabled", "#718077")])
        style.configure("Primary.TButton", background="#176b50", foreground="white")
        style.map("Primary.TButton", background=[("disabled", "#dae8df"), ("active", "#10543e")],
                  foreground=[("disabled", "#648171"), ("!disabled", "white")])
        style.configure("TEntry", padding=px(10), borderwidth=1, relief="flat",
                        fieldbackground="#f7f9f6", bordercolor="#dce5dd")
        style.configure("TProgressbar", background="#258563", troughcolor="#e5ebe4",
                        borderwidth=0, thickness=px(4), lightcolor="#258563", darkcolor="#258563")
        frame = ttk.Frame(root, padding=px(32))
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="KİŞİSEL FİNANS • BİLGİSAYARINIZDA", style="Caption.TLabel").pack(anchor="w")
        wordmark = ttk.Frame(frame)
        wordmark.pack(anchor="w", pady=(px(8), px(8)))
        ttk.Label(wordmark, text="Fintracker", style="Brand.TLabel").pack(side="left")
        ttk.Label(wordmark, text=".", style="BrandDot.TLabel").pack(side="left")
        ttk.Label(frame, text="Harcamalarınıza daha net bir bakış.", style="Muted.TLabel").pack(anchor="w", pady=(0, px(24)))
        self.setup_frame = ttk.Frame(frame, style="Card.TFrame", padding=px(24))
        self.setup_frame.pack(fill="x")
        ttk.Label(self.setup_frame, text="Verileriniz nerede saklansın?", style="Card.TLabel",
                  font=("Segoe UI", 15, "bold")).pack(anchor="w")
        ttk.Label(self.setup_frame, text="Ekstreleriniz ve raporlarınız sadece seçtiğiniz klasörde tutulur.",
                  style="Card.TLabel", foreground="#65746b").pack(anchor="w", pady=(px(6), px(18)))
        row = ttk.Frame(self.setup_frame, style="Card.TFrame")
        row.pack(fill="x")
        self.folder = tk.StringVar(value=configured_home())
        self.entry = ttk.Entry(row, textvariable=self.folder, font=("Segoe UI", 10))
        self.entry.pack(side="left", fill="x", expand=True, ipady=px(3))
        self.browse = ttk.Button(row, text="Klasör seç", command=self.choose)
        self.browse.pack(side="left", padx=(px(10), 0))
        ttk.Label(self.setup_frame, text="Gerekli bileşenleri sizin için hazırlayıp paneli açacağız.",
                  style="Card.TLabel", foreground="#65746b").pack(anchor="w", pady=(px(22), 0))
        buttons = ttk.Frame(self.setup_frame, style="Card.TFrame")
        buttons.pack(fill="x", pady=(px(16), px(12)))
        self.start_button = ttk.Button(buttons, text="Kendi verilerimle başla", style="Primary.TButton",
                                       command=lambda: self.start(False))
        self.start_button.pack(side="left")
        self.demo_button = ttk.Button(buttons, text="Önce demoyu dene", command=lambda: self.start(True))
        self.demo_button.pack(side="left", padx=px(10))
        ttk.Label(self.setup_frame, text="Sadece göz atmak mı istiyorsunuz? Demo tamamen yapay veriler kullanır.",
                  style="Card.TLabel", foreground="#65746b", font=("Segoe UI", 9)).pack(anchor="w")
        self.running_frame = ttk.Frame(frame, style="Card.TFrame", padding=px(24))
        ttk.Label(self.running_frame, text="●  HAZIR", style="Card.TLabel", foreground="#258563",
                  font=("Segoe UI", 10, "bold")).pack(anchor="w")
        ttk.Label(self.running_frame, text="Paneliniz açık.", style="Card.TLabel",
                  font=("Segoe UI", 24, "bold")).pack(anchor="w", pady=(px(12), px(8)))
        ttk.Label(self.running_frame, text="Harcamalarınızı tarayıcıdan takip edebilirsiniz.\nBu pencereyi küçültüp çalışmaya devam edin.",
                  style="Card.TLabel", foreground="#65746b").pack(anchor="w")
        controls = ttk.Frame(self.running_frame, style="Card.TFrame")
        controls.pack(fill="x", pady=(px(24), 0))
        self.open_button = ttk.Button(controls, text="Paneli aç  →", style="Primary.TButton", command=self.open_panel)
        self.open_button.pack(side="left")
        self.stop_button = ttk.Button(controls, text="Durdur ve geri dön", command=self.stop)
        self.stop_button.pack(side="left", padx=px(10))
        self.progress = ttk.Progressbar(frame, mode="indeterminate")
        self.progress.pack(fill="x", pady=(px(24), px(12)))
        self.status = tk.StringVar(value="Hazır. İlk kurulum birkaç dakika sürebilir.")
        ttk.Label(frame, textvariable=self.status, style="Muted.TLabel", wraplength=px(620)).pack(anchor="w")
        root.protocol("WM_DELETE_WINDOW", self.close)
        root.after(100, self.poll)

    def choose(self):
        path = filedialog.askdirectory(title="Özel veri klasörünüzü seçin", parent=self.root)
        if path:
            self.folder.set(path)

    def set_busy(self, value):
        self.busy = value
        for widget in (self.entry, self.browse, self.start_button, self.demo_button):
            widget.configure(state="disabled" if value else "normal")

    def start(self, demo):
        if self.busy:
            return
        path = None
        if not demo:
            try:
                path = validate_home(self.folder.get())
            except (ValueError, OSError) as exc:
                messagebox.showerror("Veri klasörü", str(exc), parent=self.root)
                return
        self.set_busy(True)
        self.progress.start()
        self.status.set("Başlatılıyor…")
        threading.Thread(target=self.worker, args=(demo, path), daemon=True).start()

    def worker(self, demo, path):
        try:
            python = prepare_environment(lambda s: self.events.put(("status", s)))
            env = dict(os.environ, PYTHONIOENCODING="utf-8")
            args = [str(python), "-u", str(CODE / "app.py"), "--no-browser"]
            if demo:
                args.append("--demo")
                env["HARCAMA_HOME"] = str(CODE / "demo")
            else:
                for name in ("data", "Belgeler"):
                    (path / name).mkdir(parents=True, exist_ok=True)
                save_home(path)
                env["HARCAMA_HOME"] = str(path)
            self.events.put(("status", "Demo hazırlanıyor…" if demo else "Panel açılıyor…"))
            self.process = subprocess.Popen(args, cwd=CODE, env=env, stdout=subprocess.PIPE,
                                            stderr=subprocess.STDOUT, text=True, encoding="utf-8",
                                            errors="replace", **HIDDEN)
            ready = False
            for line in self.process.stdout:
                # Only forward the startup URL; never surface private pipeline output here.
                if line.startswith("Harcamalarım is running at "):
                    url = line.split(" at ", 1)[1].split()[0]
                    if url.startswith("http://127.0.0.1:"):
                        ready = True
                        self.events.put(("ready", url))
            code = self.process.wait()
            self.events.put(("stopped", None) if ready else (
                "error", f"Uygulama başlatılamadı (çıkış kodu: {code}). Portları ve kurulumunuzu kontrol edip yeniden deneyin."))
        except Exception as exc:
            self.events.put(("error", str(exc)))

    def poll(self):
        try:
            while True:
                kind, value = self.events.get_nowait()
                if kind == "status":
                    self.status.set(value)
                elif kind == "ready":
                    self.progress.stop()
                    self.progress.configure(mode="determinate", value=100)
                    self.setup_frame.pack_forget()
                    self.running_frame.pack(fill="x", before=self.progress)
                    self.url = value
                    self.status.set("Panel hazır. Uygulama açık kaldığı sürece bu pencereyi açık tutun.")
                    self.open_button.configure(state="normal")
                    self.stop_button.configure(state="normal")
                    self.open_panel()
                elif kind in ("error", "stopped"):
                    self.progress.stop()
                    self.progress.configure(mode="indeterminate", value=0)
                    self.running_frame.pack_forget()
                    self.setup_frame.pack(fill="x", before=self.progress)
                    self.url = None
                    self.process = None
                    self.set_busy(False)
                    self.open_button.configure(state="disabled")
                    self.stop_button.configure(state="disabled")
                    self.status.set("Kurulum tamamlanamadı. Yeniden deneyebilirsiniz." if kind == "error" else "Uygulama durduruldu.")
                    if kind == "error":
                        messagebox.showerror("Başlatma tamamlanamadı", value, parent=self.root)
        except queue.Empty:
            pass
        self.root.after(100, self.poll)

    def open_panel(self):
        if self.url:
            webbrowser.open(self.url)

    def stop(self):
        if self.process and self.process.poll() is None:
            self.process.terminate()

    def close(self):
        if self.busy and not self.url:
            messagebox.showinfo("Hazırlık sürüyor", "Hazırlık tamamlandığında pencereyi kapatabilirsiniz.", parent=self.root)
            return
        if self.url and not messagebox.askokcancel(
                "Uygulama kapatılsın mı?",
                "Pencereyi kapatmak yerel sunucuyu da durdurur. Tarayıcıdan dosya eklemeye devam etmek için İptal'e basıp bu pencereyi küçültün.",
                parent=self.root):
            return
        self.stop()
        self.root.destroy()


if __name__ == "__main__":
    enable_high_dpi()
    window = tk.Tk()
    Launcher(window)
    window.mainloop()
