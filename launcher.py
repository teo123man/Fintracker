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
        self.root.geometry("740x620")
        self.root.minsize(700, 600)
        self.root.configure(bg="#f5f3ee")
        style = ttk.Style(root)
        style.theme_use("clam")
        style.configure("TFrame", background="#f5f3ee")
        style.configure("TLabel", background="#f5f3ee", foreground="#24372f", font=("Segoe UI", 11))
        style.configure("Title.TLabel", font=("Segoe UI", 26, "bold"))
        style.configure("TButton", font=("Segoe UI", 11), padding=(14, 10))
        style.configure("Primary.TButton", background="#176b50", foreground="white")
        style.map("Primary.TButton", background=[("active", "#10543e"), ("disabled", "#becbc5")])
        frame = ttk.Frame(root, padding=30)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Fintracker", style="Title.TLabel").pack(anchor="w")
        ttk.Label(frame, text="Harcamalarınız için hazır bir başlangıç.").pack(anchor="w", pady=(4, 22))
        ttk.Label(frame, text="1   Veri klasörünüzü seçin", font=("Segoe UI", 12, "bold")).pack(anchor="w")
        ttk.Label(frame, text="Ekstreleriniz ve raporlarınız bu bilgisayarda saklanır.").pack(anchor="w", pady=(6, 10))
        row = ttk.Frame(frame)
        row.pack(fill="x")
        self.folder = tk.StringVar(value=configured_home())
        self.entry = ttk.Entry(row, textvariable=self.folder, font=("Segoe UI", 10))
        self.entry.pack(side="left", fill="x", expand=True, ipady=7)
        self.browse = ttk.Button(row, text="Klasör seç", command=self.choose)
        self.browse.pack(side="left", padx=(8, 0))
        ttk.Label(frame, text="2   Başlayın", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(24, 8))
        ttk.Label(frame, text="İlk açılışta gerekli bileşenler otomatik kurulur.\nArdından tarayıcıda ekstrelerinizi seçebilirsiniz.").pack(anchor="w")
        buttons = ttk.Frame(frame)
        buttons.pack(fill="x", pady=(16, 10))
        self.start_button = ttk.Button(buttons, text="Kendi verilerimle başla", style="Primary.TButton",
                                       command=lambda: self.start(False))
        self.start_button.pack(side="left")
        self.demo_button = ttk.Button(buttons, text="Önce demoyu dene", command=lambda: self.start(True))
        self.demo_button.pack(side="left", padx=10)
        ttk.Label(frame, text="Demo, tamamen yapay verilerle açılır.", font=("Segoe UI", 9)).pack(anchor="w")
        self.progress = ttk.Progressbar(frame, mode="indeterminate")
        self.progress.pack(fill="x", pady=(20, 8))
        self.status = tk.StringVar(value="Hazır. İlk kurulum birkaç dakika sürebilir.")
        ttk.Label(frame, textvariable=self.status, wraplength=620).pack(anchor="w")
        controls = ttk.Frame(frame)
        controls.pack(fill="x", pady=(14, 0))
        self.open_button = ttk.Button(controls, text="Paneli tekrar aç", state="disabled", command=self.open_panel)
        self.open_button.pack(side="left")
        self.stop_button = ttk.Button(controls, text="Uygulamayı durdur", state="disabled", command=self.stop)
        self.stop_button.pack(side="left", padx=10)
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
                    self.url = value
                    self.status.set("Panel hazır. Uygulama açık kaldığı sürece bu pencereyi açık tutun.")
                    self.open_button.configure(state="normal")
                    self.stop_button.configure(state="normal")
                    self.open_panel()
                elif kind in ("error", "stopped"):
                    self.progress.stop()
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
        self.stop()
        self.root.destroy()


if __name__ == "__main__":
    window = tk.Tk()
    Launcher(window)
    window.mainloop()
