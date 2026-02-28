#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""Interface gráfica profissional para executar preencher_planilha_normalizado.py."""

from __future__ import annotations

import json
import shlex
import subprocess
import sys
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

try:
    from PIL import Image, ImageTk
except ImportError:  # pragma: no cover - fallback visual
    Image = None
    ImageTk = None

SCRIPT_BASE = Path(__file__).with_name("preencher_planilha_normalizado.py")
SETTINGS_PATH = Path(__file__).with_name("gui_preencher_planilha.settings.json")
DATA_ROOT = Path(__file__).with_name("DADOS_MXM")
DEFAULT_PDF_DIR = DATA_ROOT / "ENTRADA_PDFS"
DEFAULT_CACHE_PATH = DEFAULT_PDF_DIR / "nomes_itens_cache.json"


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("MXM • Controle Material IP")
        self.geometry("1120x780")
        self.minsize(980, 700)
        self.configure(bg="#e9edf3")

        self._proc: subprocess.Popen[str] | None = None
        self._worker_thread: threading.Thread | None = None
        self._logo_imgs: list[object] = []
        self._run_started_at: datetime | None = None

        initial_pdf_dir = DEFAULT_PDF_DIR if DEFAULT_PDF_DIR.exists() else Path.cwd()
        self.var_pdf_dir = tk.StringVar(value=str(initial_pdf_dir))
        self.var_excel = tk.StringVar(value="")
        self.var_cache = tk.StringVar(value=str(DEFAULT_CACHE_PATH if DEFAULT_PDF_DIR.exists() else (Path.cwd() / "nomes_itens_cache.json")))
        self.var_command_preview = tk.StringVar(value="")

        self.var_dry_run = tk.BooleanVar(value=False)
        self.var_skip_export = tk.BooleanVar(value=False)
        self.var_skip_rename = tk.BooleanVar(value=False)
        self.var_open_mxm = tk.BooleanVar(value=True)
        self.var_auto_clear = tk.BooleanVar(value=False)

        self.var_last_exit = tk.StringVar(value="-")
        self.var_elapsed = tk.StringVar(value="00:00")
        self.var_status = tk.StringVar(value="Pronto")

        self._load_settings()

        self._build_style()
        self._build_layout()
        self._load_logos()
        self._wire_updates()
        self._update_command_preview()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(500, self._tick_elapsed)

    def _build_style(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure("App.TFrame", background="#ffffff")
        style.configure("Header.TFrame", background="#ffffff")
        style.configure("Footer.TFrame", background="#f8fafc")

        style.configure("Title.TLabel", font=("Segoe UI", 21, "bold"), foreground="#0f172a", background="#ffffff")
        style.configure("Subtitle.TLabel", font=("Segoe UI", 10), foreground="#475569", background="#ffffff")
        style.configure("Section.TLabelframe", background="#ffffff")
        style.configure("Section.TLabelframe.Label", font=("Segoe UI", 10, "bold"), foreground="#0f172a", background="#ffffff")
        style.configure("Field.TLabel", background="#ffffff", foreground="#0f172a", font=("Segoe UI", 10))
        style.configure("Field.TEntry", padding=6)
        style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"), padding=9)
        style.configure("Secondary.TButton", font=("Segoe UI", 10), padding=8)
        style.configure("Status.TLabel", font=("Segoe UI", 10), foreground="#334155", background="#f8fafc")
        style.configure("KPI.TLabel", font=("Segoe UI", 9, "bold"), foreground="#0f172a", background="#f8fafc")
        style.configure("TCheckbutton", background="#ffffff", foreground="#0f172a", font=("Segoe UI", 10))

    def _build_layout(self) -> None:
        container = ttk.Frame(self, padding=20, style="App.TFrame")
        container.pack(fill="both", expand=True, padx=18, pady=18)

        header = ttk.Frame(container, style="Header.TFrame")
        header.pack(fill="x", pady=(4, 14))

        self.fm_label = ttk.Label(header, style="Field.TLabel")
        self.fm_label.pack(side="left", padx=(0, 14))

        title_wrap = ttk.Frame(header, style="Header.TFrame")
        title_wrap.pack(side="left", fill="x", expand=True)
        ttk.Label(title_wrap, text="Processador MXM", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            title_wrap,
            text="Interface interativa com execução monitorada, logs em tempo real e diagnóstico de falhas.",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(3, 0))

        self.vision_label = ttk.Label(header, style="Field.TLabel")
        self.vision_label.pack(side="right", padx=(14, 0))

        form_card = ttk.LabelFrame(container, text="Arquivos e diretórios", padding=14, style="Section.TLabelframe")
        form_card.pack(fill="x")
        form_card.columnconfigure(1, weight=1)

        self._add_path_row(form_card, 0, "Pasta dos PDFs", self.var_pdf_dir, self._choose_pdf_dir)
        self._add_path_row(form_card, 1, "Planilha Excel (.xlsx)", self.var_excel, self._choose_excel)
        self._add_path_row(form_card, 2, "Cache manual (JSON)", self.var_cache, self._choose_cache)

        actions_inline = ttk.Frame(form_card, style="Header.TFrame")
        actions_inline.grid(row=3, column=1, sticky="w", pady=(6, 0))
        ttk.Button(actions_inline, text="Detectar Excel automaticamente", style="Secondary.TButton", command=self._auto_detect_excel).pack(side="left")
        ttk.Button(actions_inline, text="Verificar dependências", style="Secondary.TButton", command=self._check_dependencies).pack(side="left", padx=(8, 0))

        options_card = ttk.LabelFrame(container, text="Opções", padding=14, style="Section.TLabelframe")
        options_card.pack(fill="x", pady=(12, 0))
        ttk.Checkbutton(options_card, text="Dry-run (não salva Excel / não renomeia)", variable=self.var_dry_run).pack(anchor="w")
        ttk.Checkbutton(options_card, text="Pular exportação de abas para PDF", variable=self.var_skip_export).pack(anchor="w")
        ttk.Checkbutton(options_card, text="Pular renomeação/movimentação de PDFs", variable=self.var_skip_rename).pack(anchor="w")
        ttk.Checkbutton(options_card, text="Abrir pasta MXM ao finalizar", variable=self.var_open_mxm).pack(anchor="w")
        ttk.Checkbutton(options_card, text="Limpar log automaticamente antes de executar", variable=self.var_auto_clear).pack(anchor="w")

        preview_card = ttk.LabelFrame(container, text="Comando", padding=12, style="Section.TLabelframe")
        preview_card.pack(fill="x", pady=(12, 0))
        self.command_entry = ttk.Entry(preview_card, textvariable=self.var_command_preview, state="readonly")
        self.command_entry.pack(fill="x")

        actions = ttk.Frame(container, style="Header.TFrame")
        actions.pack(fill="x", pady=(12, 0))

        self.btn_start = ttk.Button(actions, text="Executar processamento", style="Primary.TButton", command=self._run_script)
        self.btn_start.pack(side="left")

        self.btn_stop = ttk.Button(actions, text="Parar execução", style="Secondary.TButton", command=self._stop_script)
        self.btn_stop.pack(side="left", padx=(8, 0))
        self.btn_stop.state(["disabled"])

        ttk.Button(actions, text="Limpar log", style="Secondary.TButton", command=self._clear_log).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Salvar log", style="Secondary.TButton", command=self._save_log).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Abrir pasta MXM", style="Secondary.TButton", command=self._open_mxm_folder).pack(side="left", padx=(8, 0))

        self.progress = ttk.Progressbar(actions, mode="indeterminate", length=180)
        self.progress.pack(side="right")

        log_frame = ttk.LabelFrame(container, text="Log de execução", padding=10, style="Section.TLabelframe")
        log_frame.pack(fill="both", expand=True, pady=(12, 0))

        self.log = tk.Text(log_frame, height=20, bg="#0b1220", fg="#dbe7ff", insertbackground="#dbe7ff", relief="flat")
        self.log.pack(fill="both", expand=True, side="left")

        scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        scroll.pack(side="right", fill="y")
        self.log.configure(yscrollcommand=scroll.set)

        footer = ttk.Frame(container, style="Footer.TFrame", padding=(12, 8))
        footer.pack(fill="x", pady=(10, 0))

        ttk.Label(footer, text="Status:", style="KPI.TLabel").pack(side="left")
        ttk.Label(footer, textvariable=self.var_status, style="Status.TLabel").pack(side="left", padx=(5, 16))
        ttk.Label(footer, text="Duração:", style="KPI.TLabel").pack(side="left")
        ttk.Label(footer, textvariable=self.var_elapsed, style="Status.TLabel").pack(side="left", padx=(5, 16))
        ttk.Label(footer, text="Último retorno:", style="KPI.TLabel").pack(side="left")
        ttk.Label(footer, textvariable=self.var_last_exit, style="Status.TLabel").pack(side="left", padx=(5, 0))

        ttk.Label(footer, text="GUI 3.0", style="Status.TLabel").pack(side="right")

    def _wire_updates(self) -> None:
        for variable in (
            self.var_pdf_dir,
            self.var_excel,
            self.var_cache,
            self.var_dry_run,
            self.var_skip_export,
            self.var_skip_rename,
        ):
            variable.trace_add("write", lambda *_: self._update_command_preview())

    def _add_path_row(self, parent: ttk.LabelFrame, row: int, label: str, variable: tk.StringVar, callback) -> None:
        ttk.Label(parent, text=label, style="Field.TLabel").grid(row=row, column=0, sticky="w", padx=(0, 10), pady=7)
        ttk.Entry(parent, textvariable=variable, style="Field.TEntry").grid(row=row, column=1, sticky="ew", pady=7)
        ttk.Button(parent, text="Selecionar", style="Secondary.TButton", command=callback).grid(row=row, column=2, padx=(10, 0), pady=7)

    def _load_logos(self) -> None:
        fm = self._load_logo([Path("FM.jpg"), Path("FM.JPG")])
        vision = self._load_logo([Path("VISION.JPG"), Path("VISION.jpg")])

        if fm is not None:
            self.fm_label.configure(image=fm)
            self.fm_label.image = fm
        else:
            self.fm_label.configure(text="FM.jpg", style="Field.TLabel")

        if vision is not None:
            self.vision_label.configure(image=vision)
            self.vision_label.image = vision
        else:
            self.vision_label.configure(text="VISION.JPG", style="Field.TLabel")

    def _load_logo(self, candidates: list[Path]):
        if Image is None or ImageTk is None:
            return None

        for file_path in candidates:
            if not file_path.exists():
                continue
            try:
                img = Image.open(file_path)
                img = img.resize((132, 60), Image.Resampling.LANCZOS)
                tkimg = ImageTk.PhotoImage(img)
                self._logo_imgs.append(tkimg)
                return tkimg
            except Exception:
                continue
        return None

    def _choose_pdf_dir(self) -> None:
        selected = filedialog.askdirectory(initialdir=self.var_pdf_dir.get() or str(Path.cwd()))
        if selected:
            self.var_pdf_dir.set(selected)
            self.var_cache.set(str(Path(selected) / "nomes_itens_cache.json"))

    def _choose_excel(self) -> None:
        selected = filedialog.askopenfilename(
            initialdir=self.var_pdf_dir.get() or str(Path.cwd()),
            filetypes=[("Excel files", "*.xlsx")],
        )
        if selected:
            self.var_excel.set(selected)

    def _choose_cache(self) -> None:
        selected = filedialog.asksaveasfilename(
            initialdir=self.var_pdf_dir.get() or str(Path.cwd()),
            defaultextension=".json",
            filetypes=[("JSON", "*.json")],
        )
        if selected:
            self.var_cache.set(selected)

    def _auto_detect_excel(self) -> None:
        pdf_dir = Path(self.var_pdf_dir.get().strip() or ".")
        if not pdf_dir.exists():
            messagebox.showerror("Pasta inválida", "A pasta informada para PDFs não existe.")
            return

        base_excel_dir = DATA_ROOT / "BASE_EXCEL"
        excels = []
        if base_excel_dir.exists():
            excels.extend(sorted([p for p in base_excel_dir.glob("*.xlsx") if not p.name.startswith("~$")]))
        excels.extend(sorted([p for p in pdf_dir.glob("*.xlsx") if not p.name.startswith("~$")]))
        if not excels:
            messagebox.showwarning("Excel não encontrado", "Nenhum arquivo .xlsx foi encontrado na pasta de PDFs.")
            return

        self.var_excel.set(str(excels[0]))
        self._append_log(f"[INFO] Excel detectado automaticamente: {excels[0]}\n")

    def _check_dependencies(self) -> None:
        cmd = [sys.executable, "-c", "import pdfplumber,openpyxl,PIL;print('Dependencias OK')"]
        try:
            out = subprocess.check_output(cmd, cwd=str(Path.cwd()), text=True, stderr=subprocess.STDOUT)
            messagebox.showinfo("Dependências", out.strip())
        except subprocess.CalledProcessError as exc:
            messagebox.showerror("Dependências", f"Falha ao validar dependências:\n\n{exc.output}")

    def _load_settings(self) -> None:
        if not SETTINGS_PATH.exists():
            return
        try:
            data = json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
        except Exception:
            return

        self.var_pdf_dir.set(data.get("pdf_dir", self.var_pdf_dir.get()))
        self.var_excel.set(data.get("excel", self.var_excel.get()))
        self.var_cache.set(data.get("cache", self.var_cache.get()))
        self.var_dry_run.set(bool(data.get("dry_run", self.var_dry_run.get())))
        self.var_skip_export.set(bool(data.get("skip_export", self.var_skip_export.get())))
        self.var_skip_rename.set(bool(data.get("skip_rename", self.var_skip_rename.get())))
        self.var_open_mxm.set(bool(data.get("open_mxm", self.var_open_mxm.get())))
        self.var_auto_clear.set(bool(data.get("auto_clear", self.var_auto_clear.get())))

    def _save_settings(self) -> None:
        data = {
            "pdf_dir": self.var_pdf_dir.get().strip(),
            "excel": self.var_excel.get().strip(),
            "cache": self.var_cache.get().strip(),
            "dry_run": self.var_dry_run.get(),
            "skip_export": self.var_skip_export.get(),
            "skip_rename": self.var_skip_rename.get(),
            "open_mxm": self.var_open_mxm.get(),
            "auto_clear": self.var_auto_clear.get(),
        }
        try:
            SETTINGS_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass

    def _build_command(self) -> list[str]:
        command = [
            sys.executable,
            "-u",  # log em tempo real sem buffering
            str(SCRIPT_BASE),
            "--pdf-dir",
            self.var_pdf_dir.get().strip() or str(Path.cwd()),
        ]
        excel = self.var_excel.get().strip()
        cache = self.var_cache.get().strip()

        if excel:
            command.extend(["--excel", excel])
        if cache:
            command.extend(["--manual-cache", cache])
        if self.var_dry_run.get():
            command.append("--dry-run")
        if self.var_skip_export.get():
            command.append("--skip-export")
        if self.var_skip_rename.get():
            command.append("--skip-rename")

        command.append("--non-interactive")
        return command

    def _update_command_preview(self) -> None:
        cmd = self._build_command()
        self.var_command_preview.set(" ".join(shlex.quote(x) for x in cmd))

    def _validate_inputs(self) -> bool:
        pdf_dir = Path(self.var_pdf_dir.get().strip() or ".")
        if not pdf_dir.exists():
            messagebox.showerror("Pasta inválida", "A pasta de PDFs informada não existe.")
            return False

        excel_path = self.var_excel.get().strip()
        if excel_path and Path(excel_path).suffix.lower() != ".xlsx":
            messagebox.showerror("Excel inválido", "Selecione um arquivo Excel com extensão .xlsx.")
            return False

        if not SCRIPT_BASE.exists():
            messagebox.showerror("Script não encontrado", f"Arquivo não encontrado: {SCRIPT_BASE}")
            return False

        return True

    def _set_running_state(self, running: bool) -> None:
        if running:
            self.btn_start.state(["disabled"])
            self.btn_stop.state(["!disabled"])
            self.progress.start(10)
            self.var_status.set("Executando...")
            self._run_started_at = datetime.now()
        else:
            self.btn_start.state(["!disabled"])
            self.btn_stop.state(["disabled"])
            self.progress.stop()
            self._run_started_at = None

    def _tick_elapsed(self) -> None:
        if self._run_started_at:
            delta = datetime.now() - self._run_started_at
            total = int(delta.total_seconds())
            mm, ss = divmod(total, 60)
            hh, mm = divmod(mm, 60)
            self.var_elapsed.set(f"{hh:02d}:{mm:02d}:{ss:02d}")
        self.after(500, self._tick_elapsed)

    def _run_script(self) -> None:
        if self._proc is not None:
            messagebox.showwarning("Execução em andamento", "Já existe uma execução em andamento.")
            return

        if not self._validate_inputs():
            return

        if self.var_auto_clear.get():
            self._clear_log()

        self.var_last_exit.set("-")
        self._save_settings()
        command = self._build_command()
        self._append_log(f"\n[{self._timestamp()}] >>> Executando:\n{' '.join(shlex.quote(x) for x in command)}\n\n")

        self._set_running_state(True)
        self._worker_thread = threading.Thread(target=self._worker, args=(command,), daemon=True)
        self._worker_thread.start()

    def _stop_script(self) -> None:
        if self._proc is None:
            return
        try:
            self._proc.terminate()
            self._append_log(f"\n[{self._timestamp()}] [AVISO] Solicitação de parada enviada.\n")
            self.after(3000, self._kill_if_running)
        except Exception as exc:
            self._append_log(f"\n[{self._timestamp()}] [ERRO] Não foi possível parar: {exc}\n")

    def _kill_if_running(self) -> None:
        if self._proc is None:
            return
        if self._proc.poll() is None:
            try:
                self._proc.kill()
                self._append_log(f"\n[{self._timestamp()}] [AVISO] Processo forçado a encerrar.\n")
            except Exception as exc:
                self._append_log(f"\n[{self._timestamp()}] [ERRO] Falha ao forçar encerramento: {exc}\n")

    def _worker(self, command: list[str]) -> None:
        try:
            self._proc = subprocess.Popen(
                command,
                cwd=str(Path.cwd()),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
            )

            assert self._proc.stdout is not None
            for line in self._proc.stdout:
                self.after(0, self._append_log, line)

            returncode = self._proc.wait()
            self.after(0, self._finish_run, returncode)
        except Exception as exc:
            self.after(0, self._append_log, f"\n[{self._timestamp()}] [ERRO] Falha ao executar: {exc}\n")
            self.after(0, self._finish_run, 1)

    def _finish_run(self, returncode: int) -> None:
        self._proc = None
        self._set_running_state(False)
        self.var_last_exit.set(str(returncode))

        if returncode == 0:
            self.var_status.set("Concluído com sucesso")
            self._append_log(f"\n[{self._timestamp()}] [OK] Processamento finalizado com sucesso.\n")
            if self.var_open_mxm.get():
                self._open_mxm_folder()
        else:
            self.var_status.set("Finalizado com erro")
            self._append_log(f"\n[{self._timestamp()}] [ERRO] Execução finalizada com código {returncode}.\n")
            messagebox.showerror(
                "Execução com erro",
                "O processamento terminou com erro. Verifique o log para detalhes.",
            )

    def _open_mxm_folder(self) -> None:
        base = Path(self.var_pdf_dir.get().strip() or ".")
        candidatos = [DATA_ROOT / "MXM", base / "MXM"]
        destino = next((p for p in candidatos if p.exists()), None)
        if destino is None:
            self._append_log(f"[{self._timestamp()}] [AVISO] Pasta MXM ainda não existe: {candidatos[0]}\n")
            return

        try:
            if sys.platform.startswith("win"):
                subprocess.Popen(["explorer", str(destino)])
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(destino)])
            else:
                subprocess.Popen(["xdg-open", str(destino)])
        except Exception as exc:
            self._append_log(f"[{self._timestamp()}] [ERRO] Não foi possível abrir a pasta MXM: {exc}\n")

    def _save_log(self) -> None:
        path = filedialog.asksaveasfilename(
            defaultextension=".log",
            filetypes=[("Log", "*.log"), ("Text", "*.txt")],
            initialfile=f"mxm_exec_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log",
        )
        if not path:
            return
        try:
            Path(path).write_text(self.log.get("1.0", "end"), encoding="utf-8")
            messagebox.showinfo("Log salvo", f"Log salvo em:\n{path}")
        except Exception as exc:
            messagebox.showerror("Erro ao salvar log", str(exc))

    def _timestamp(self) -> str:
        return datetime.now().strftime("%H:%M:%S")

    def _append_log(self, text: str) -> None:
        self.log.insert("end", text)
        self.log.see("end")

    def _clear_log(self) -> None:
        self.log.delete("1.0", "end")

    def _on_close(self) -> None:
        if self._proc is not None:
            if not messagebox.askyesno(
                "Encerrar aplicação",
                "Existe uma execução em andamento. Deseja encerrar a aplicação assim mesmo?",
            ):
                return
            try:
                self._proc.terminate()
            except Exception:
                pass

        self._save_settings()
        self.destroy()


if __name__ == "__main__":
    app = App()
    app.mainloop()
