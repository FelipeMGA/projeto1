#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""Instalador do processador MXM (CLI + GUI)."""

from __future__ import annotations

import argparse
import os
import platform
import shutil
import subprocess
import sys
from pathlib import Path


IS_WINDOWS = platform.system().lower().startswith("win")
RUNTIME_PACKAGES = ["pdfplumber", "openpyxl", "pillow"]


def run(cmd: list[str], cwd: Path) -> None:
    print("[RUN]", " ".join(cmd))
    subprocess.check_call(cmd, cwd=str(cwd))


def python_venv_exec(venv_dir: Path) -> Path:
    if IS_WINDOWS:
        return venv_dir / "Scripts" / "python.exe"
    return venv_dir / "bin" / "python"


def _desktop_dir() -> Path | None:
    home = Path.home()
    desk = home / "Desktop"
    if desk.exists():
        return desk
    one_drive_desk = home / "OneDrive" / "Desktop"
    if one_drive_desk.exists():
        return one_drive_desk
    return None


def preparar_estrutura_padrao(repo_dir: Path) -> dict[str, Path]:
    pasta_dados = repo_dir / "DADOS_MXM"
    pasta_entrada = pasta_dados / "ENTRADA_PDFS"
    pasta_mxm = pasta_dados / "MXM"
    pasta_folhas = pasta_dados / "FOLHAS DAS EQUIPES"
    pasta_excel = pasta_dados / "BASE_EXCEL"
    pasta_backup = pasta_dados / "BACKUP_EXCEL"

    for p in (pasta_dados, pasta_entrada, pasta_mxm, pasta_folhas, pasta_excel, pasta_backup):
        p.mkdir(parents=True, exist_ok=True)

    excel_destino = pasta_excel / "CONTROLE_MATERIAL_IP.xlsx"
    if not excel_destino.exists():
        candidatos = sorted([p for p in repo_dir.glob("*.xlsx") if not p.name.startswith("~$")])
        if candidatos:
            shutil.copy2(candidatos[0], excel_destino)
            print(f"[INFO] Excel base copiado para: {excel_destino}")

    return {
        "dados": pasta_dados,
        "entrada": pasta_entrada,
        "mxm": pasta_mxm,
        "folhas": pasta_folhas,
        "excel": pasta_excel,
        "backup": pasta_backup,
    }


def criar_atalhos_desktop(pastas: dict[str, Path]) -> None:
    if not IS_WINDOWS:
        return

    desktop = _desktop_dir()
    if not desktop:
        print("[AVISO] Área de trabalho não encontrada para criar atalhos.")
        return

    atalhos = {
        "MXM - Pasta MXM.bat": pastas["mxm"],
        "MXM - Pasta Folhas das Equipes.bat": pastas["folhas"],
        "MXM - Pasta Base Excel.bat": pastas["excel"],
    }

    for nome, destino in atalhos.items():
        arquivo = desktop / nome
        arquivo.write_text(
            "@echo off\n"
            f'explorer "{destino}"\n',
            encoding="utf-8",
        )
        print(f"[OK] Atalho criado na área de trabalho: {arquivo}")


def criar_atalhos_execucao(repo_dir: Path, venv_dir: Path, pastas: dict[str, Path]) -> None:
    py_exec = python_venv_exec(venv_dir)
    pdf_dir = pastas["entrada"]
    excel_dir = pastas["excel"]

    if IS_WINDOWS:
        bat = repo_dir / "executar_gui.bat"
        bat.write_text(
            "@echo off\n"
            f'"{py_exec}" "{repo_dir / "gui_preencher_planilha.py"}"\n',
            encoding="utf-8",
        )
        print(f"[OK] Atalho criado: {bat}")

        conf = repo_dir / "abrir_pastas_mxm.bat"
        conf.write_text(
            "@echo off\n"
            f'explorer "{pdf_dir}"\n'
            f'explorer "{excel_dir}"\n',
            encoding="utf-8",
        )
        print(f"[OK] Atalho criado: {conf}")
        return

    sh = repo_dir / "executar_gui.sh"
    sh.write_text(
        "#!/usr/bin/env bash\n"
        "set -euo pipefail\n"
        f'"{py_exec}" "{repo_dir / "gui_preencher_planilha.py"}"\n',
        encoding="utf-8",
    )
    os.chmod(sh, 0o755)
    print(f"[OK] Atalho criado: {sh}")


def validar_pre_requisitos(py_exec: Path) -> None:
    try:
        run([str(py_exec), "-c", "import tkinter; print('tkinter ok')"], Path.cwd())
    except Exception:
        print("[AVISO] tkinter não disponível no Python atual. A GUI pode não abrir.")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Instala dependências e configura o ambiente do MXM.")
    parser.add_argument("--repo-dir", type=Path, default=Path(__file__).resolve().parents[1], help="Diretório do projeto")
    parser.add_argument("--venv-path", type=Path, default=Path(".venv_mxm"), help="Pasta do ambiente virtual")
    parser.add_argument("--skip-pip", action="store_true", help="Não instala pacotes com pip")
    parser.add_argument("--run-gui", action="store_true", help="Abre a GUI automaticamente ao final da instalação")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    repo_dir = args.repo_dir.resolve()
    repo_default = Path(__file__).resolve().parents[1]

    if not (repo_dir / "requirements-mxm.txt").exists():
        print(
            f"[AVISO] requirements-mxm.txt não encontrado em {repo_dir}. "
            f"Usando diretório padrão do script: {repo_default}"
        )
        repo_dir = repo_default

    req_exists = (repo_dir / "requirements-mxm.txt").exists()
    if not req_exists:
        print(
            f"[AVISO] requirements-mxm.txt não encontrado em {repo_dir}. "
            "Instalador vai usar lista interna de pacotes (pdfplumber, openpyxl, pillow)."
        )

    venv_dir = (repo_dir / args.venv_path).resolve() if not args.venv_path.is_absolute() else args.venv_path

    print("[INFO] Repositório:", repo_dir)
    print("[INFO] Ambiente virtual:", venv_dir)

    run([sys.executable, "-m", "venv", str(venv_dir)], repo_dir)
    py_exec = python_venv_exec(venv_dir)

    if not args.skip_pip:
        run([str(py_exec), "-m", "pip", "install", "--upgrade", "pip"], repo_dir)

        if req_exists:
            try:
                run([str(py_exec), "-m", "pip", "install", "-r", "requirements-mxm.txt"], repo_dir)
            except subprocess.CalledProcessError:
                print(
                    "[AVISO] Falha ao instalar via requirements-mxm.txt. "
                    "Tentando instalação direta dos pacotes essenciais..."
                )
                run([str(py_exec), "-m", "pip", "install", *RUNTIME_PACKAGES], repo_dir)
        else:
            run([str(py_exec), "-m", "pip", "install", *RUNTIME_PACKAGES], repo_dir)

        if IS_WINDOWS:
            try:
                run([str(py_exec), "-m", "pip", "install", "pywin32"], repo_dir)
            except subprocess.CalledProcessError:
                print("[AVISO] Falha ao instalar pywin32 (exportação PDF no Windows pode não funcionar).")

    pastas = preparar_estrutura_padrao(repo_dir)
    validar_pre_requisitos(py_exec)
    criar_atalhos_execucao(repo_dir, venv_dir, pastas)
    criar_atalhos_desktop(pastas)

    print("\n[SUCESSO] Instalação concluída.")
    print("- GUI: execute 'executar_gui.bat' (Windows) ou './executar_gui.sh' (Linux/macOS).")
    print(f"- CLI: {py_exec} {repo_dir / 'preencher_planilha_normalizado.py'} --help")
    print(f"- Pasta de entrada padrão: {pastas['entrada']}")
    print(f"- Pasta de saída MXM: {pastas['mxm']}")
    print(f"- Pasta de folhas: {pastas['folhas']}")
    print(f"- Pasta Excel base: {pastas['excel']}")

    if args.run_gui:
        print("[INFO] Abrindo GUI...")
        run([str(py_exec), str(repo_dir / "gui_preencher_planilha.py")], repo_dir)


if __name__ == "__main__":
    main()
