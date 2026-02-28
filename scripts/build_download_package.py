#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""Gera pacote ZIP pronto para download/distribuição do MXM."""

from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

EXCLUDES = {
    ".git",
    ".venv_mxm",
    "__pycache__",
    "node_modules",
    "dist",
}

EXCLUDE_SUFFIXES = {
    ".pyc",
    ".pyo",
    ".zip",
}

INCLUDE_FILES = [
    "preencher_planilha_normalizado.py",
    "gui_preencher_planilha.py",
    "requirements-mxm.txt",
    "instalar_mxm.bat",
    "instalar_mxm.sh",
    "scripts/install_mxm_tool.py",
    "scripts/instalar_windows.bat",
    "scripts/instalar_linux.sh",
    "scripts/build_windows_installer.ps1",
    "installer/windows/MXM_Setup.iss",
    "docs/INSTALACAO_MXM.md",
    "docs/PRONTO_PARA_INSTALAR.md",
]

OPTIONAL_FILES = [
    "FM.jpg",
    "FM.JPG",
    "VISION.jpg",
    "VISION.JPG",
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Empacota o projeto MXM em arquivo ZIP para download.")
    parser.add_argument("--repo-dir", type=Path, default=Path(__file__).resolve().parents[1], help="Diretório do projeto")
    parser.add_argument("--output-dir", type=Path, default=Path("dist"), help="Diretório onde o ZIP será gerado")
    parser.add_argument("--name", type=str, default="", help="Nome base do ZIP (sem extensão)")
    return parser.parse_args()


def should_skip(path: Path) -> bool:
    parts = set(path.parts)
    if parts & EXCLUDES:
        return True
    if path.suffix.lower() in EXCLUDE_SUFFIXES:
        return True
    return False


def main() -> None:
    args = parse_args()
    repo_dir = args.repo_dir.resolve()
    output_dir = (repo_dir / args.output_dir).resolve() if not args.output_dir.is_absolute() else args.output_dir
    output_dir.mkdir(parents=True, exist_ok=True)

    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    base_name = args.name.strip() or f"MXM_PRONTO_PARA_DOWNLOAD_{stamp}"
    zip_path = output_dir / f"{base_name}.zip"

    with ZipFile(zip_path, "w", compression=ZIP_DEFLATED) as zf:
        for rel in INCLUDE_FILES:
            file_path = repo_dir / rel
            if not file_path.exists():
                raise FileNotFoundError(f"Arquivo obrigatório não encontrado: {file_path}")
            zf.write(file_path, arcname=rel)

        for rel in OPTIONAL_FILES:
            file_path = repo_dir / rel
            if file_path.exists() and file_path.is_file() and not should_skip(file_path.relative_to(repo_dir)):
                zf.write(file_path, arcname=rel)

    print(f"[OK] Pacote pronto para download: {zip_path}")


if __name__ == "__main__":
    main()
