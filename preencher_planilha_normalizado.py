#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Script multi-PDF para preencher o CONTROLE MATERIAL IP a partir de requisições MXM em PDF.

Principais features:

- Lê TODOS os PDFs "originais" da pasta (mais recentes primeiro), ignorando os já renomeados
  no padrão "DD-MM-AAAA - SETOR - OBS.pdf" ou "AAAA-MM-DD - SETOR - OBS.pdf".

- Para cada PDF:
    • Identifica:
        - Setor (palavra após "NI")
        - Data de entrega (campo "Data:" do cabeçalho / ENTREGA / primeira data dd/mm/aaaa)
        - Número da requisição (6 dígitos) → MXM
        - OBS (linha que começa com "OBS")

    • Lê itens:
        - Código (ABC.DEF.GHI000)
        - Descrição completa (linhas após o código até QNT/QTD...)
        - Quantidade requisitada (QNT/QTD REQUISITADA / REQ, aceita 1.000 / 10,0 etc)
        - Normaliza a descrição:
            - Regras fixas (braços, plaqueta, luminária Unicoba, cabos, conector, relé, etc.)
            - Itens especiais por código:
                 MAT.TRA.DESC001 → LÂMPADA MVM 150/250/400 RETIRADO (divide QNT)
                 MAT.TRA.DESC002 → REATOR MVM 150/250/400 RETIRADO (divide QNT)
                 MAT.TRA.DESC003 → RELÉ RETIRADO (25) + BASE RETIRADA (5)
                 MAT.TRA.DESC009 → LED PREFEITURA 80 RETIRADO
                 CNE.PER.BIME001 → CONECTOR TORÇÃO
                 CNE.PER.BIME014 → CONECTOR PERFURANTE SUBTERRÂNEO
                 ISL.FIS.AUTF003 → FITA ISOLANTE AUTOFUSÃO
            - Cabo quadriplex:
                 descrições com "quadrup" / "quadriplex" → CABO QUADRIPLEX
            - Itens "novos": gera nome curto e coerente (máx. 35 caracteres) interpretando
              a descrição (fita, cabo, conector, relé, etc.).

        - Agrupa itens SOMENTE se código E descrição forem iguais (nunca junta códigos diferentes).

    • Localiza a aba no Excel:
        - Ignora abas que terminam com "LED"
        - Procura título que contenha o nome do setor (DIOGO, ROGER, etc.)
        - Usa fallback com normalização de nome (tirando NI, OBRA, GARANTIA etc.)

    • Na aba do setor:
        - Limpa as colunas ITEM / CÓDIGO / SAÍDA da linha 5 à 40
        - Preenche itens a partir da linha 5
        - Ordena itens por ITEM (A → Z)
        - Atualiza:
            A2 = "<SETOR>" (sem "NI")
            A3 = "MXM: <número>"
            B3 = OBS do PDF
            Campo "DATA: dd/mm/aaaa" no cabeçalho

        - Insere logotipo em A1:
            FM.*     → abas MATEUS, TIAGO, ROGER, DAMILSON, DIOGO, RENATO, JRNE 5N, JRNE  5N
            VISION.* → abas GEOVANE, ALYSSON, EDVALDO

- Ao final:
    - Salva diretamente no Excel ORIGINAL.
    - Depois de salvar com sucesso, renomeia cada PDF processado para:
          "DD-MM-AAAA - SETOR - OBS.pdf"
      (sem sobrescrever arquivos existentes).

Modo simulação:
    python preencher_planilha_normalizado.py --dry-run
    (não salva Excel nem renomeia PDFs, só mostra o que faria)

Flags opcionais:
    --skip-export  -> pula exportação das abas modificadas para PDF
    --skip-rename  -> não renomeia/move os PDFs de entrada
    --pdf-dir DIR  -> pasta onde estão os PDFs de entrada (default: .)
    --excel PATH   -> caminho explícito do Excel base (default: busca em pdf-dir)
    --manual-cache -> arquivo JSON para guardar nomes manuais de itens novos
    --non-interactive -> não pede input para itens novos (usa sugestão automática)

Requisitos:
    pip install pdfplumber openpyxl pillow
"""

import argparse
import json
import math
import os
import re
import shutil
import sys
import unicodedata
from pathlib import Path
from datetime import datetime
from collections import defaultdict

import pdfplumber
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string

# ==========================
# Configurações gerais
# ==========================

# DRY_RUN = True se for chamado com argumento --dry-run
DRY_RUN = False

# Flags opcionais
SKIP_EXPORT = False  # não exporta abas para PDF
SKIP_RENAME = False  # não move/renomeia PDFs originais

# Diretórios / caminhos configuráveis
PDF_DIR = Path(".")
EXCEL_PATH = None
MANUAL_CACHE_PATH = None
INTERACTIVE_INPUT = True

# Setores por tipo de logo
SETORES_FM = {
    "MATEUS",
    "MATHEUS",
    "TIAGO",
    "ROGER",
    "DAMILSON",
    "DIOGO",
    "RENATO",
    "JRNE 5N",
    "JRNE  5N",
    "JRNE 7D",
}
SETORES_VISION = {"GEOVANE", "ALYSSON", "EDVALDO"}

# Regras fixas por código (descrição direta – SEM "FM" nos nomes)
REGRAS_CODIGO_FIXO = {
    "MAT.TRA.DESC009": "LED PREFEITURA 80 RETIRADO",
    "CNE.PER.BIME001": "CONECTOR TORÇÃO",
    "CNE.PER.BIME014": "CONECTOR PERFURANTE SUBTERRÂNEO",
    "ISL.FIS.AUTF003": "FITA ISOLANTE AUTOFUSÃO",
}

# Cache de colunas por aba: nome_aba -> (header_row, col_item, col_codigo, col_saida)
COL_CACHE = {}

# Cache para nomes informados manualmente para itens "novos"
NOME_MANUAL_CACHE = {}
_NOME_MANUAL_CACHE_SUJO = False


def _serializar_chave_manual(chave):
    desc_norm, code_norm = chave
    return f"{desc_norm}||{code_norm}"


def _desserializar_chave_manual(txt):
    if "||" in txt:
        a, b = txt.split("||", 1)
        return (a, b)
    return (txt, "")


def carregar_cache_manual():
    """Carrega o cache de nomes manuais de um arquivo JSON, se existir."""
    global NOME_MANUAL_CACHE

    if not MANUAL_CACHE_PATH:
        return
    if not MANUAL_CACHE_PATH.exists():
        return

    try:
        with MANUAL_CACHE_PATH.open("r", encoding="utf-8") as f:
            bruto = json.load(f)
    except Exception as exc:  # pragma: no cover - log defensivo
        print(f"[AVISO] Não foi possível ler o cache manual {MANUAL_CACHE_PATH}: {exc}")
        return

    if not isinstance(bruto, dict):
        print(f"[AVISO] Cache manual {MANUAL_CACHE_PATH} ignorado (formato inválido).")
        return

    restaurado = {}
    for k, v in bruto.items():
        if not isinstance(v, str):
            continue
        restaurado[_desserializar_chave_manual(k)] = v

    NOME_MANUAL_CACHE = restaurado
    if restaurado:
        print(f"[INFO] Cache manual carregado com {len(restaurado)} entradas de nomes.")


def salvar_cache_manual():
    """Persiste o cache manual em disco quando houve novas entradas."""

    global _NOME_MANUAL_CACHE_SUJO

    if not MANUAL_CACHE_PATH or not _NOME_MANUAL_CACHE_SUJO:
        return

    try:
        MANUAL_CACHE_PATH.parent.mkdir(parents=True, exist_ok=True)
        serializado = {_serializar_chave_manual(k): v for k, v in NOME_MANUAL_CACHE.items()}
        with MANUAL_CACHE_PATH.open("w", encoding="utf-8") as f:
            json.dump(serializado, f, ensure_ascii=False, indent=2)
        _NOME_MANUAL_CACHE_SUJO = False
    except Exception as exc:  # pragma: no cover - log defensivo
        print(f"[AVISO] Não foi possível salvar o cache manual em {MANUAL_CACHE_PATH}: {exc}")


# ==========================
# Argumentos CLI
# ==========================


def parse_args():
    """Processa argumentos de linha de comando e ajusta globais."""
    global DRY_RUN, SKIP_EXPORT, SKIP_RENAME, PDF_DIR, EXCEL_PATH, MANUAL_CACHE_PATH, INTERACTIVE_INPUT

    parser = argparse.ArgumentParser(
        description=(
            "Preenche o CONTROLE MATERIAL IP a partir de PDFs MXM, "
            "com opções para simular, pular exportação ou renomeação, "
            "e definir diretórios de entrada/planilha."
        )
    )
    parser.add_argument("--dry-run", action="store_true", help="Não salva Excel nem renomeia PDFs")
    parser.add_argument("--skip-export", action="store_true", help="Não exporta abas modificadas para PDF")
    parser.add_argument("--skip-rename", action="store_true", help="Não move/renomeia os PDFs de entrada")
    parser.add_argument("--pdf-dir", type=Path, default=Path("."), help="Pasta onde estão os PDFs a processar")
    parser.add_argument(
        "--excel",
        type=Path,
        default=None,
        help="Caminho explícito do Excel base (por padrão, busca em --pdf-dir)",
    )
    parser.add_argument(
        "--manual-cache",
        type=Path,
        default=None,
        help=(
            "Arquivo JSON para cachear nomes de itens digitados manualmente. "
            "Default: nomes_itens_cache.json dentro de --pdf-dir"
        ),
    )
    parser.add_argument(
        "--non-interactive",
        action="store_true",
        help="Não solicita nomes manuais para itens novos; usa sugestão automática",
    )

    args = parser.parse_args()

    DRY_RUN = args.dry_run
    SKIP_EXPORT = args.skip_export
    SKIP_RENAME = args.skip_rename
    PDF_DIR = args.pdf_dir.resolve()
    EXCEL_PATH = args.excel.resolve() if args.excel else None
    MANUAL_CACHE_PATH = (
        args.manual_cache.resolve()
        if args.manual_cache
        else (PDF_DIR / "nomes_itens_cache.json")
    )
    INTERACTIVE_INPUT = (not args.non_interactive) and sys.stdin.isatty()

    if not PDF_DIR.exists():
        raise FileNotFoundError(f"Diretório de PDFs não encontrado: {PDF_DIR}")

    if EXCEL_PATH and not EXCEL_PATH.exists():
        raise FileNotFoundError(f"Arquivo Excel especificado não encontrado: {EXCEL_PATH}")

    return args


# ==========================
# Utilitários de texto
# ==========================

def normalize_text(s):
    """Normaliza texto para comparação (minúsculo, sem acentos, trim)."""
    if not s:
        return ""
    s = str(s)
    s_nfkd = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s_nfkd if not unicodedata.combining(c))
    return s.lower().strip()


def sanitize_filename_part(s, max_len=40):
    """
    Limpa texto para uso em nome de arquivo:
    - tira acento
    - remove caracteres inválidos
    - limita tamanho
    """
    if not s:
        return ""
    s = str(s)
    s_nfkd = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s_nfkd if not unicodedata.combining(c))
    s = re.sub(r'[\\/:*?"<>|]', "-", s)
    s = re.sub(r"\s+", " ", s).strip()
    if len(s) > max_len:
        s = s[:max_len].rstrip()
    return s


def column_width_to_pixels(width):
    """Converte largura de coluna do Excel para pixels (aproximação)."""
    default_width = 8.43
    try:
        w = float(width if width is not None else default_width)
    except (TypeError, ValueError):
        w = default_width

    pixels = int(math.floor(((256 * w + math.floor(128 / 7)) / 256) * 7))
    return max(pixels, 1)


def points_to_pixels(points):
    """Converte pontos tipográficos (72 DPI) para pixels (96 DPI)."""
    try:
        return max(1, int(math.ceil(float(points) * 96 / 72)))
    except (TypeError, ValueError):
        return 1


def pixels_to_emu(pixels):
    """Converte pixels para EMU (English Metric Units) usados pelo Excel."""
    try:
        return int(max(0, float(pixels)) * 9525)
    except (TypeError, ValueError):
        return 0


def colocar_imagem_na_celula(ws, cell, img_path, padding_px=2):
    """
    Posiciona uma imagem totalmente dentro de uma célula específica,
    redimensionando-a proporcionalmente e centralizando-a.

    A estratégia abaixo evita anchors avançados que não estão se
    comportando conforme esperado na visualização do Excel do usuário.
    Em vez disso, definimos o anchor diretamente na célula e
    aplicamos deslocamentos (colOff/rowOff) manualmente para centralizar
    a imagem.
    """
    try:
        img = Image(str(img_path))
    except Exception as exc:
        print(f"[ERRO] Falha ao carregar a imagem '{img_path}': {exc}")
        return None

    col_letter, row_idx = coordinate_from_string(cell)
    row_number = int(row_idx)
    col_number = column_index_from_string(col_letter)

    col_width = ws.column_dimensions[col_letter].width
    if col_width is None:
        col_width = ws.sheet_format.defaultColWidth or 8.43

    row_height = ws.row_dimensions[row_number].height
    if row_height is None:
        row_height = ws.sheet_format.defaultRowHeight or 15

    cell_width_px = column_width_to_pixels(col_width)
    cell_height_px = points_to_pixels(row_height)

    max_width = max(1, cell_width_px - padding_px * 2)
    max_height = max(1, cell_height_px - padding_px * 2)

    if img.width and img.height:
        scale = min(max_width / img.width, max_height / img.height, 1.0)
        img.width = int(max(1, img.width * scale))
        img.height = int(max(1, img.height * scale))
    else:
        img.width = max_width
        img.height = max_height

    col_off_px = padding_px + max(0, (cell_width_px - img.width) // 2)
    row_off_px = padding_px + max(0, (cell_height_px - img.height) // 2)

    try:
        from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
        from openpyxl.drawing.xdr import XDRPositiveSize2D

        anchor = OneCellAnchor(
            _from=AnchorMarker(
                col=col_number - 1,
                colOff=pixels_to_emu(col_off_px),
                row=row_number - 1,
                rowOff=pixels_to_emu(row_off_px),
            ),
            ext=XDRPositiveSize2D(
                pixels_to_emu(img.width),
                pixels_to_emu(img.height),
            ),
        )
        img.anchor = anchor
    except Exception:
        # fallback para o anchor padrão na célula, mesmo que sem centralização precisa
        img.anchor = cell

    ws.add_image(img)
    return img


def limpar_sufixos_ruido(linha):
    """Remove sufixos colados ao código, como 'TIPODEES'."""
    if not linha:
        return ""
    return re.sub(r"tipodees", "", linha, flags=re.IGNORECASE)


def encurtar_descricao(desc, max_len=35):
    """
    Encurta descrição para no máximo max_len caracteres,
    tentando cortar em limite de palavra.
    Usado para itens "novos".
    """
    if not desc:
        return ""
    desc = desc.strip()
    if len(desc) <= max_len:
        return desc

    corte = desc.rfind(" ", 0, max_len)
    if corte == -1 or corte < max_len * 0.6:
        return desc[:max_len - 1].rstrip() + "…"
    return desc[:corte].rstrip() + "…"


def normalizar_nome_aba_ou_setor(nome):
    """
    Normaliza nome de aba/setor para comparação:
    - lowercase, sem acento
    - remove 'ni', 'obra', 'garantia', 'manutencao', 'led'
    - remove caracteres não alfanuméricos
    """
    n = normalize_text(nome or "")
    for lixo in [" ni ", " obra", " garantia", " manutencao", " led"]:
        n = n.replace(lixo, " ")
    n = n.replace("ni-", " ").replace("ni_", " ")
    n = re.sub(r"[^a-z0-9]+", "", n)
    return n


# ==========================
# Geração de nome inteligente para itens novos
# ==========================

def gerar_nome_generico_inteligente(desc, code=None):
    """
    Gera um nome curto, coerente e inteligente para itens "novos",
    interpretando a descrição completa e extraindo:
      - tipo (fita, cabo, conector, relé, etc.)
      - características principais (cor, bitola, comprimento, tipo PP/singelo, etc.)
    No final, garante no máximo 35 caracteres.
    Nunca adiciona "FM" em nomes.
    O parâmetro opcional code permite identificar itens como disjuntores
    mesmo que a descrição não cite explicitamente a palavra.
    """
    if not desc:
        return ""
    d = normalize_text(desc)
    codigo_norm = normalize_text(code or "")

    def _extrair_amperagem(texto):
        valores_validos = {"6", "10", "16", "20", "25", "30", "32", "40", "50", "63", "70"}
        m_amp = re.search(r"\b(6|10|16|20|25|30|32|40|50|63|70)(?:[\.,]0+)?\s*(?:a|amp|amps|ampere|amperes)?\b", texto)
        if m_amp:
            valor = m_amp.group(1)
            if valor in valores_validos:
                return f"{valor}A"
        return None

    # ---------- FITAS ----------
    if "fita" in d:
        partes = ["FITA"]
        if "isolant" in d or "isolacao" in d:
            partes.append("ISOLANTE")

        # autofusão
        if ("autofus" in d) or ("auto" in d and "fusa" in d) or ("self" in d and "fusing" in d):
            partes.append("AUTOFUSÃO")

        # cor
        if "preta" in d:
            partes.append("PRETA")
        elif "branca" in d:
            partes.append("BRANCA")
        elif "amarela" in d:
            partes.append("AMARELA")
        elif "vermelha" in d:
            partes.append("VERMELHA")

        # comprimento (20m, 10m etc.)
        m_len = re.search(r"(\d+)\s*m", d)
        if m_len:
            partes.append(f"{m_len.group(1)}M")

        nome = " ".join(partes)
        return encurtar_descricao(nome)

    # ---------- TOMADAS ----------
    if "tomada" in d:
        partes = ["TOMADA"]
        amp = _extrair_amperagem(d)
        if amp:
            partes.append(amp)
        nome = " ".join(partes)
        return encurtar_descricao(nome)

    # ---------- CABOS / CONDUTORES ----------
    if "cabo" in d or "condutor" in d:
        partes = ["CABO"]

        if "pp" in d:
            partes.append("PP")
        if "singelo" in d:
            partes.append("SINGELO")
        if "coax" in d:
            partes.append("COAXIAL")
        if "quadrup" in d or "quadriplex" in d:
            partes.append("QUADRIPLEX")

        if "aluminio" in d:
            partes.append("ALUMÍNIO")
        if "cobre" in d:
            partes.append("COBRE")

        m_bit = re.search(r"(\d+(?:,\d+)?)\s*mm2", d)
        if not m_bit:
            m_bit = re.search(r"(\d+(?:,\d+)?)\s*mm\b", d)
        if m_bit:
            bit = m_bit.group(1).replace(",", ",")
            partes.append(f"{bit}MM")

        m_vias = re.search(r"(\d+)\s*(vias|via|x)", d)
        if m_vias:
            partes.append(f"{m_vias.group(1)} VIAS")

        nome = " ".join(partes)
        return encurtar_descricao(nome)

    # ---------- DISJUNTORES ----------
    if "disjuntor" in d or d.startswith("dis.") or codigo_norm.startswith("dis"):
        partes = ["DISJUNTOR"]

        if "tripo" in d or re.search(r"\b3\s*(polos|p)\b", d):
            partes.append("TRIPOLAR")
        elif "bipo" in d or re.search(r"\b2\s*(polos|p)\b", d):
            partes.append("BIPOLAR")
        elif "mono" in d or re.search(r"\b1\s*(polos|p)\b", d):
            partes.append("MONOPOLAR")

        amp = _extrair_amperagem(d)
        if amp:
            partes.append(amp)

        nome = " ".join(partes)
        return encurtar_descricao(nome)

    # ---------- CONECTORES ----------
    if "conector" in d or "conectores" in d:
        partes = ["CONECTOR"]
        if "perfur" in d or "derivacao" in d:
            partes.append("PERFURANTE")
        if "bimet" in d:
            partes.append("BIMETÁLICO")
        if "subterr" in d:
            partes.append("SUBTERRÂNEO")
        if "torcao" in d:
            partes.append("TORÇÃO")
        # não adiciona "FM" em hipótese alguma
        nome = " ".join(partes)
        return encurtar_descricao(nome)

    # ---------- RELÉS ----------
    if "rele" in d or "relé" in desc.lower():
        partes = ["RELÉ"]
        if "fotoeletr" in d:
            partes.append("FOTOELÉTRICO")
        if "demape" in d:
            partes.append("DEMAPE")
        if "telegest" in d:
            partes.append("TELEGESTÃO")
        nome = " ".join(partes)
        return encurtar_descricao(nome)

    # ---------- REATOR / LÂMPADA / LUMINÁRIA ----------
    if "reator" in d:
        partes = ["REATOR"]
        m_w = re.search(r"(\d+)\s*w", d)
        if m_w:
            partes.append(f"{m_w.group(1)}W")
        if "mvm" in d:
            partes.append("MVM")
        nome = " ".join(partes)
        return encurtar_descricao(nome)

    if "lampada" in d or "lâmpada" in desc.lower():
        partes = ["LÂMPADA"]
        m_w = re.search(r"(\d+)\s*w", d)
        if m_w:
            partes.append(f"{m_w.group(1)}W")
        if "mvm" in d:
            partes.append("MVM")
        nome = " ".join(partes)
        return encurtar_descricao(nome)

    if "luminaria" in d or "luminária" in desc.lower():
        partes = ["LUMINÁRIA"]
        m_w = re.search(r"(\d+)\s*w", d)
        if m_w:
            partes.append(f"{m_w.group(1)}W")
        if "led" in d:
            partes.append("LED")
        nome = " ".join(partes)
        return encurtar_descricao(nome)

    # ---------- PLAQUETAS / ETIQUETAS ----------
    if "etiqueta" in d or "plaqueta" in d:
        partes = ["PLAQUETA"]
        if "aluminio" in d:
            partes.append("ALUMÍNIO")
        if "poste" in d:
            partes.append("POSTE")
        nome = " ".join(partes)
        return encurtar_descricao(nome)

    # ---------- PARAFUSOS / PORCAS / ARRUELAS ----------
    if "parafuso" in d:
        partes = ["PARAFUSO"]
        m_len = re.search(r"(\d+)\s*mm", d)
        if m_len:
            partes.append(f"{m_len.group(1)}MM")
        nome = " ".join(partes)
        return encurtar_descricao(nome)

    if "porca" in d:
        partes = ["PORCA"]
        m_m = re.search(r"m\s*[-]?\s*(\d+)", d)
        if m_m:
            partes.append(f"M{m_m.group(1)}")
        nome = " ".join(partes)
        return encurtar_descricao(nome)

    if "arruela" in d:
        partes = ["ARRUELA"]
        if "lisa" in d:
            partes.append("LISA")
        m_m = re.search(r"m\s*[-]?\s*(\d+)", d)
        if m_m:
            partes.append(f"M{m_m.group(1)}")
        nome = " ".join(partes)
        return encurtar_descricao(nome)

    # ---------- CABO DE AÇO ----------
    if "aco" in d and "cabo" in d:
        nome = "CABO DE AÇO"
        return encurtar_descricao(nome)

    # Fallback: encurta a descrição original
    return encurtar_descricao(desc)


def solicitar_nome_manual(desc_original, code, sugestao):
    """
    Pergunta ao usuário qual nome deve ser usado para um item sem regra conhecida.
    Usa cache para evitar repetir a pergunta para a mesma descrição/código.
    Se estiver em modo não interativo, usa automaticamente a sugestão.
    """
    global _NOME_MANUAL_CACHE_SUJO

    chave = (normalize_text(desc_original or ""), (code or "").upper())
    if chave in NOME_MANUAL_CACHE:
        return NOME_MANUAL_CACHE[chave]

    if not INTERACTIVE_INPUT:
        final = sugestao
        NOME_MANUAL_CACHE[chave] = final
        _NOME_MANUAL_CACHE_SUJO = True
        salvar_cache_manual()
        print(
            f"[INFO] Modo não interativo: usando sugestão automática para {code or '(sem código)'} -> {final}"
        )
        return final

    print("\n[INTERAÇÃO] Item novo identificado sem regra específica.")
    print(f"  Código: {code or '(sem código)'}")
    print(f"  Descrição original: {desc_original or '(vazia)'}")
    print(f"  Sugestão automática: {sugestao}")
    try:
        resposta = input("Digite o nome desejado (ENTER para aceitar a sugestão): ").strip()
    except EOFError:
        resposta = ""

    final = resposta or sugestao
    NOME_MANUAL_CACHE[chave] = final
    _NOME_MANUAL_CACHE_SUJO = True
    salvar_cache_manual()
    return final


# ==========================
# Normalização das descrições conhecidas
# ==========================

def normalizar_descricao_conhecida(desc):
    """
    Tenta aplicar regras de normalização que você definiu.
    Se bater alguma, retorna a descrição normalizada.
    Se não bater nenhuma, retorna None (item "novo").
    """
    if not desc:
        return None
    d = normalize_text(desc)

    # FITA AUTOFUSÃO (antes das outras fitas)
    if ("autofus" in d) or ("auto" in d and "fusa" in d) or ("self" in d and "fusing" in d):
        return "FITA ISOLANTE AUTOFUSÃO"

    # Braços
    if "braco para ip 005" in d:
        return "BRAÇO MÉDIO"
    if "braco para ip 010" in d:
        return "BRAÇO LONGO"
    if "braco para ip 013" in d or "braco tipo s" in d:
        return "BRAÇO TIPO S"

    # Adesivo / cola
    if "colas tecnicas" in d or "adesivo" in d:
        return "ADESIVO PARA PLAQUETA"

    # Plaqueta de alumínio
    if "etiqueta aluminio" in d:
        return "PLAQUETA PARA POSTE"

    # Luminárias LED Unicoba – só potência
    if "luminaria led publica" in d:
        m = re.search(r"(\d+)\s*w", d)
        if m:
            return f"LED UNICOBA {m.group(1)}W"

    # Cabo quadriplex
    if "quadrup" in d or "quadriplex" in d:
        return "CABO QUADRIPLEX"

    # Cinta – CINTA + medida
    if "cinta" in d:
        m = re.search(r"(\d+)\s*mm", d)
        if m:
            return f"CINTA {m.group(1)}MM"

    # Parafuso – PARAFUSO + medida
    if "parafuso" in d:
        m = re.search(r"(\d+)\s*mm", d)
        if m:
            return f"PARAFUSO {m.group(1)}MM"

    # Cabos específicos
    if "singelo" in d and "2,5" in d:
        return "CABO SINGELO VERDE 2,5MM"
    if "pp" in d and "2,5" in d:
        return "CABO PP 2,5MM"

    # Conectores (genéricos)
    if "perfuracao" in d or "derivacao" in d:
        return "CONECTOR PERFURANTE"
    if "torcao" in d:
        return "CONECTOR TORÇÃO"

    # Relés
    if "demape" in d:
        return "RELÉ DEMAPE"
    if "telegestao" in d or "modulo eletr" in d or "modulo eletronico" in d:
        return "RELÉ TELEGESTÃO"

    # Fita isolante preta 20m
    if "fita" in d and "isolant" in d and "preta" in d and "20m" in d:
        return "FITA ISOLANTE PRETA 20M"

    return None


def normalizar_descricao(desc, code=None):
    """
    Aplica regras conhecidas de normalização.
    Se não bater nenhuma regra, considera item "novo"
    e gera um nome inteligente baseado na descrição completa.
    (Função mantida por compatibilidade; hoje o controle de "novo" é feito em extract_items.)
    """
    if not desc:
        return ""
    conhecido = normalizar_descricao_conhecida(desc)
    if conhecido is not None:
        return conhecido
    return gerar_nome_generico_inteligente(desc, code)


# ==========================
# Arquivos (PDFs e Excel)
# ==========================

def find_pdfs():
    """
    Retorna PDFs 'originais' da pasta (mais recentes primeiro),
    IGNORANDO PDFs já renomeados no padrão:
        "dd-mm-aaaa - SETOR - OBS.pdf"
        "aaaa-mm-dd - SETOR - OBS.pdf"
    """
    pdfs = list(PDF_DIR.glob("*.pdf"))
    if not pdfs:
        raise FileNotFoundError("Nenhum PDF encontrado na pasta.")

    padrao_data_inicio = re.compile(r"^(\d{2}-\d{2}-\d{4}|\d{4}-\d{2}-\d{2})\s-\s")

    originais = []
    ignorados = []
    for p in pdfs:
        if padrao_data_inicio.match(p.name):
            ignorados.append(p)
        else:
            originais.append(p)

    if ignorados:
        print("\n[INFO] PDFs ignorados (já renomeados):")
        for p in ignorados:
            print(f"  - {p.name}")

    if not originais:
        raise FileNotFoundError(
            "Todos os PDFs na pasta já estão renomeados no padrão de saída.\n"
            "Nenhum PDF 'original' encontrado para processar."
        )

    originais.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return originais


def find_excel():
    """
    Encontra o arquivo .xlsx base na pasta.
    Evita usar arquivos que começam com 'CONTROLE MATERIAL IP - ATUALIZADO'
    como base quando possível.
    """
    if EXCEL_PATH:
        return EXCEL_PATH

    excels = sorted(PDF_DIR.glob("*.xlsx"))
    excels = [e for e in excels if not e.name.startswith("~$")]
    if not excels:
        raise FileNotFoundError("Nenhum arquivo .xlsx encontrado na pasta.")

    candidatos = [e for e in excels if not e.name.startswith("CONTROLE MATERIAL IP - ATUALIZADO")]
    if not candidatos:
        candidatos = excels

    return candidatos[0]


# ==========================
# Extração de dados do PDF
# ==========================

def extract_lines(pdf_path):
    """Extrai todas as linhas de texto do PDF."""
    linhas = []
    try:
        with pdfplumber.open(pdf_path) as doc:
            for i, page in enumerate(doc.pages, start=1):
                try:
                    texto = page.extract_text() or ""
                    linhas.extend(l.strip() for l in texto.splitlines() if l.strip())
                except Exception as e:
                    print(f"[ERRO] Falha ao extrair texto da página {i} do PDF {pdf_path.name}: {e}")
    except Exception as e:
        print(f"[ERRO] Não foi possível abrir o PDF {pdf_path.name}: {e}")
        return []

    if not linhas:
        print(f"[AVISO] Nenhuma linha foi extraída do PDF {pdf_path.name}.")

    return linhas


def extract_sector(lines):
    """Setor = palavra logo após 'NI' (ex.: 'NI DIOGO' → 'DIOGO')."""
    # Detecta setores JRNE 5N e JRNE 7D mesmo quando não vierem precedidos por "NI"
    for l in lines:
        norm_line = normalize_text(l)
        if re.search(r"\bjrne\s*5n\b", norm_line):
            return "JRNE 5N"
        if re.search(r"\bjrne\s*7d\b", norm_line):
            return "JRNE 7D"

    for l in lines:
        m = re.search(r"\bni\s+([a-zA-Z]+)", normalize_text(l))
        if m:
            return m.group(1).upper()
    return None


def extract_entrega_date(lines):
    """
    Tenta encontrar data de ENTREGA no formato dd/mm/aaaa.
    Prioriza:
        1) Linha de cabeçalho com "Hora" e "Data" (campo circulado)
        2) Linhas próximas da palavra 'entrega'
        3) Primeira data encontrada no PDF
    """
    date_pattern = re.compile(r"(\d{2}/\d{2}/\d{4})")

    # 1) Cabeçalho: linha com Hora/Data (ex.: "Hora: 23:00   Data: 17/12/2025")
    for line in lines:
        norm = normalize_text(line)
        if "hora" in norm and "data" in norm:
            m = date_pattern.search(line)
            if m:
                return m.group(1)

    # 2) Linhas próximas da palavra ENTREGA
    for i, line in enumerate(lines):
        norm = normalize_text(line)
        if "entrega" in norm:
            m = date_pattern.search(line)
            if m:
                return m.group(1)
            for j in range(i + 1, min(i + 4, len(lines))):
                m2 = date_pattern.search(lines[j])
                if m2:
                    return m2.group(1)

    for line in lines:
        m = date_pattern.search(line)
        if m:
            return m.group(1)

    return None


def extract_numero_requisicao(lines):
    """
    Número da requisição = primeiro número de 6 dígitos.
    Prioriza linhas com 'numero'.
    """
    num_pattern = re.compile(r"\b(\d{6})\b")

    for i, line in enumerate(lines):
        norm = normalize_text(line)
        if "numero" in norm:
            m = num_pattern.search(line)
            if m:
                return m.group(1)
            for j in range(i + 1, min(i + 4, len(lines))):
                m2 = num_pattern.search(lines[j])
                if m2:
                    return m2.group(1)

    for line in lines[:20]:
        m = num_pattern.search(line)
        if m:
            return m.group(1)

    return None


def extract_obs(lines):
    """OBS = linha que começa com 'OBS' (ignora maiúsc./minúsc.)."""
    for line in lines:
        norm = normalize_text(line)
        if norm.startswith("obs"):
            txt = re.sub(r"(?i)^obs[:\-\s]*", "", line).strip()
            return txt or None
    return None


def _parse_quantity(raw):
    """
    Converte texto numérico possivelmente com . ou , para int.
    Exemplos: '1.000' -> 1000, '10,0' -> 10, '5' -> 5
    """
    if not raw:
        return None
    raw = raw.strip()
    raw = raw.replace(" ", "")

    valor = raw.replace(".", "").replace(",", ".")
    try:
        return int(float(valor))
    except ValueError:
        return None


def extract_items(lines):
    """
    Extrai itens da requisição:
    - código no formato ABC.DEF.GHI000
    - descrição = linhas após o código até linha que mencione QNT/QTD
    - quantidade = número após 'QNT./QTD. REQUISITADA/REQ' com fallback
                   para primeiro bloco numérico da linha/bloco (aceita 1.000, 10,0, etc.)

    Marca itens que caíram em regra "conhecida" vs itens "novos".
    """
    items = []
    # Formato dos códigos: XXX.XXX.XXXXXXX
    # - primeiros 6 caracteres: apenas letras (dois blocos de 3)
    # - sufixo: exatamente 7 caracteres alfanuméricos
    # Usar comprimento fixo evita absorver sufixos colados como "TI" e retorna apenas
    # o código base (ex.: "DIS.DJT.0001041TI" → captura "DIS.DJT.0001041").
    code_re = re.compile(r"[A-Z]{3}\.[A-Z]{3}\.[A-Z0-9]{7}", re.IGNORECASE)
    cleaned_lines = [limpar_sufixos_ruido(l) for l in lines]
    idx = []

    for i, l in enumerate(cleaned_lines):
        canon = re.sub(r"\s+", "", l)
        sanitized = re.sub(r"[^A-Za-z0-9.]", "", l)
        if code_re.search(canon) or code_re.search(sanitized):
            idx.append(i)

    for k, start in enumerate(idx):
        end = idx[k + 1] if k + 1 < len(idx) else len(lines)
        block = cleaned_lines[start:end]

        canon_first = re.sub(r"\s+", "", block[0])
        sanitized_first = re.sub(r"[^A-Za-z0-9.]", "", block[0])
        m_code = code_re.search(canon_first) or code_re.search(sanitized_first)
        if not m_code:
            continue
        code = m_code.group(0).upper()

        # Descrição: linhas após o código até encontrar QNT/QTD
        desc_lines = []
        for ln in block[1:]:
            nln = normalize_text(ln)
            if "qnt" in nln or "qtd" in nln:
                break
            desc_lines.append(ln)
        raw_desc = " ".join(desc_lines).strip()

        # primeiro tenta regra conhecida
        conhecido = normalizar_descricao_conhecida(raw_desc)
        if conhecido is not None:
            desc = conhecido
            is_new = False
        else:
            sugestao = gerar_nome_generico_inteligente(raw_desc, code)
            desc = solicitar_nome_manual(raw_desc, code, sugestao)
            is_new = True

        # Quantidade
        qty = None
        for ln in block:
            norm_ln = normalize_text(ln)
            m = re.search(r"(qnt|qtd)[\. ]*(requisitada|req)?:[\-\s]*([\d\.,]+)", norm_ln)
            if m:
                q = _parse_quantity(m.group(3))
                if q is not None:
                    qty = q
                    break

        if qty is None:
            for ln in block:
                m2 = re.search(r"([\d\.,]+)", ln)
                if m2:
                    q = _parse_quantity(m2.group(1))
                    if q is not None:
                        qty = q
                        break

        if qty is None:
            continue

        items.append({"code": code, "desc": desc, "qty": qty, "is_new": is_new})

    return items


# ==========================
# Itens especiais + agrupamento
# ==========================

def processar_itens_especiais(items):
    """
    Trata códigos especiais:
        MAT.TRA.DESC001 -> 3 linhas LÂMPADA MVM 150/250/400 RETIRADO (divide quantidade)
        MAT.TRA.DESC002 -> 3 linhas REATOR MVM 150/250/400 RETIRADO (divide quantidade)
        MAT.TRA.DESC003 -> sempre 25 RELÉ RETIRADO + 5 BASE RETIRADA (ignora QNT do PDF)
        + REGRAS_CODIGO_FIXO para códigos com descrição fixa.
    """
    novos = []
    for it in items:
        code = it["code"]
        qty = it["qty"]

        # Lâmpadas MVM
        if code == "MAT.TRA.DESC001":
            base_q, sobra = divmod(qty, 3)
            q1, q2, q3 = base_q, base_q, base_q + sobra
            novos.extend([
                {"code": code, "desc": "LÂMPADA MVM 150W RETIRADO", "qty": q1},
                {"code": code, "desc": "LÂMPADA MVM 250W RETIRADO", "qty": q2},
                {"code": code, "desc": "LÂMPADA MVM 400W RETIRADO", "qty": q3},
            ])
            continue

        # Reatores MVM
        if code == "MAT.TRA.DESC002":
            base_q, sobra = divmod(qty, 3)
            q1, q2, q3 = base_q, base_q, base_q + sobra
            novos.extend([
                {"code": code, "desc": "REATOR MVM 150W RETIRADO", "qty": q1},
                {"code": code, "desc": "REATOR MVM 250W RETIRADO", "qty": q2},
                {"code": code, "desc": "REATOR MVM 400W RETIRADO", "qty": q3},
            ])
            continue

        # RELÉ + BASE RETIRADA – sempre 25 + 5 (total 30)
        if code == "MAT.TRA.DESC003":
            novos.append({"code": code, "desc": "RELÉ RETIRADO", "qty": 25})
            novos.append({"code": code, "desc": "BASE RETIRADA", "qty": 5})
            continue

        # Códigos com descrição fixa via tabela
        if code in REGRAS_CODIGO_FIXO:
            novos.append({"code": code, "desc": REGRAS_CODIGO_FIXO[code], "qty": qty})
            continue

        novos.append(it)

    return novos


def agrupar_por_descricao(items):
    """
    Agrupa SOMENTE itens com MESMO CÓDIGO e MESMA DESCRIÇÃO.
    (Nunca junta códigos diferentes!)
    """
    grupos = defaultdict(lambda: {"qty": 0})

    for it in items:
        chave = (it["code"], it["desc"])
        grupos[chave]["qty"] += it["qty"]

    saida = []
    for (code, desc), info in grupos.items():
        saida.append({
            "code": code,
            "desc": desc,
            "qty": info["qty"],
        })

    return saida


# ==========================
# Planilha (Excel)
# ==========================

def find_sheet(wb, sector):
    """
    Acha a aba do setor de forma simples e confiável:
    - ignora abas que terminam com 'LED'
    - considera que o nome do setor (ex.: 'DIOGO') aparece em algum lugar do título da aba
      (ex.: 'DIOGO', 'DIOGO OBRA', 'NI DIOGO GARANTIA', etc.).
    """
    if not sector:
        return None

    su = (sector or "").upper().strip()

    # 1) tentativa direta: substring do setor no título
    for ws in wb.worksheets:
        t = ws.title.upper().strip()
        if t.endswith("LED"):
            continue
        if su and su in t:
            return ws

    # 2) fallback: tenta bater nomes "limpos" (tirando NI, OBRA, GARANTIA...)
    key_setor = normalizar_nome_aba_ou_setor(sector)
    if not key_setor:
        return None

    for ws in wb.worksheets:
        t = ws.title.upper()
        if t.endswith("LED"):
            continue
        key_aba = normalizar_nome_aba_ou_setor(ws.title)
        if key_aba == key_setor or key_setor in key_aba or key_aba in key_setor:
            return ws

    return None


def find_columns(ws):
    """
    Encontra linha do cabeçalho (4) e colunas ITEM, CÓDIGO, SAÍDA.
    Retorna (header_row, col_item, col_codigo, col_saida).
    """
    header_row = 4
    desc_col = code_col = saida_col = None

    for c in range(1, 40):
        v = ws.cell(header_row, c).value
        n = normalize_text(v or "")
        if n == "item":
            desc_col = c
        if "cod" in n:
            code_col = c
        if "saida" in n:
            saida_col = c

    return header_row, desc_col, code_col, saida_col


def get_columns_cached(ws):
    """Retorna colunas da aba, usando cache para não recalcular toda hora."""
    if ws.title in COL_CACHE:
        return COL_CACHE[ws.title]
    cols = find_columns(ws)
    COL_CACHE[ws.title] = cols
    return cols


def fill_items(ws, items):
    """
    Limpa colunas ITEM/CÓDIGO/SAÍDA da linha 5 à 40
    e preenche a partir da linha 5 com os itens.
    """
    header, dc, cc, sc = get_columns_cached(ws)
    if not dc or not cc or not sc:
        print("[ERRO] Não achou colunas ITEM/CÓDIGO/SAÍDA na aba:", ws.title)
        return

    start = header + 1  # 5

    for r in range(start, 41):
        ws.cell(r, dc, "")
        ws.cell(r, cc, "")
        ws.cell(r, sc, "")

    for i, it in enumerate(items):
        r = start + i
        ws.cell(r, dc, it["desc"])
        ws.cell(r, cc, it["code"])
        ws.cell(r, sc, it["qty"])


def ordenar_itens(ws):
    """
    Ordena os itens da tabela pela coluna ITEM (A → Z),
    mantendo código e quantidade.
    """
    header_row, desc_col, code_col, saida_col = get_columns_cached(ws)
    if not desc_col or not code_col or not saida_col:
        print(f"[AVISO] Não foi possível ordenar itens na aba {ws.title}.")
        return

    start = header_row + 1
    dados = []
    row = start
    while row <= 200:
        item = ws.cell(row, desc_col).value
        code = ws.cell(row, code_col).value
        qty = ws.cell(row, saida_col).value
        if item is None or str(item).strip() == "":
            break
        dados.append((item, code, qty))
        row += 1

    if not dados:
        return

    dados_ordenados = sorted(dados, key=lambda x: normalize_text(x[0]))

    for r in range(start, start + len(dados)):
        ws.cell(r, desc_col, "")
        ws.cell(r, code_col, "")
        ws.cell(r, saida_col, "")

    for i, (item, code, qty) in enumerate(dados_ordenados):
        r = start + i
        ws.cell(r, desc_col, item)
        ws.cell(r, code_col, code)
        ws.cell(r, saida_col, qty)

    print(f"[INFO] Itens ordenados de A a Z na aba {ws.title}.")


def set_data_entrega(ws, data_entrega):
    """Atualiza o campo 'DATA: dd/mm/aaaa' no cabeçalho."""
    if not data_entrega:
        print("[AVISO] Nenhuma data de entrega encontrada no PDF.")
        return

    for row in range(1, 10):
        for col in range(1, 20):
            val = ws.cell(row, col).value
            if not val:
                continue
            n = normalize_text(val)
            if "data" in n:
                ws.cell(row, col, f"DATA: {data_entrega}")
                print(f"[INFO] Campo DATA atualizado para: {data_entrega}")
                return

    print("[AVISO] Não encontrei o campo DATA na planilha.")


def set_setor_header(ws, sector):
    """A2 = '<SETOR>' (sem 'NI')."""
    if not sector:
        return
    ws["A2"] = sector.upper()
    print(f"[INFO] A2 atualizado para: {sector.upper()}")


def set_mxm(ws, numero_mxm):
    """A3 = 'MXM: <número>'."""
    if not numero_mxm:
        print("[AVISO] Nenhum número de requisição encontrado para MXM.")
        return
    ws["A3"] = f"MXM: {numero_mxm}"
    print(f"[INFO] A3 atualizado para: MXM: {numero_mxm}")


def set_obs(ws, obs):
    """Atualiza B3 com OBS do PDF."""
    if not obs:
        print("[AVISO] Nenhuma OBS encontrada no PDF.")
        return
    obs_maiusculo = str(obs).upper()
    ws.cell(row=3, column=2, value=obs_maiusculo)
    print(f"[INFO] OBS atualizada em B3: {obs_maiusculo}")


# ==========================
# Logotipos
# ==========================

def limpar_imagens_da_aba(ws):
    """Remove imagens existentes na aba."""
    if hasattr(ws, "_images"):
        ws._images = []


def _normalizar_setor_logo(nome):
    """Normaliza nome do setor para comparação de logo."""
    base = normalize_text(nome or "")
    return re.sub(r"\s+", " ", base).strip()


SETORES_FM_NORM = {_normalizar_setor_logo(s) for s in SETORES_FM}
SETORES_VISION_NORM = {_normalizar_setor_logo(s) for s in SETORES_VISION}


def _tipo_logo_por_setor(nome_setor):
    """Retorna 'FM' ou 'VISION' conforme setor, ou None se não houver regra."""
    setor_norm = _normalizar_setor_logo(nome_setor)
    if setor_norm in SETORES_FM_NORM:
        return "FM"
    if setor_norm in SETORES_VISION_NORM:
        return "VISION"
    return None


def _encontrar_logo(prefixo, diretorios):
    """Procura logo com prefixo informado, priorizando FM.jpg/VISION.jpg."""
    prefixo_norm = prefixo.lower().strip()
    extensoes = [".jpg", ".jpeg", ".png"]

    candidatos_diretos = [
        f"{prefixo}.jpg",
        f"{prefixo}.JPG",
        f"{prefixo}.jpeg",
        f"{prefixo}.JPEG",
        f"{prefixo}.png",
        f"{prefixo}.PNG",
    ]

    # 1) Tentativa direta por nomes esperados (mais estável para operação)
    for base in diretorios:
        if not base or not base.exists():
            continue
        for nome_esperado in candidatos_diretos:
            candidato = base / nome_esperado
            if candidato.exists() and candidato.is_file():
                return candidato

    # 2) Fallback case-insensitive para arquivos equivalentes
    for base in diretorios:
        if not base or not base.exists():
            continue
        for item in base.iterdir():
            if not item.is_file():
                continue
            nome = item.name.lower()
            if not nome.startswith(prefixo_norm + "."):
                continue
            if item.suffix.lower() in extensoes:
                return item

    return None


def inserir_logotipo(ws):
    """
    Insere logo FM ou VISION em A1 conforme o setor.
    (Nomes dos itens não levam 'FM', apenas o arquivo de logo da empresa.)
    """
    img_path = None
    script_dir = Path(__file__).resolve().parent
    diretorios_busca = [PDF_DIR, script_dir / "DADOS_MXM", script_dir, Path.cwd()]
    tipo_logo = _tipo_logo_por_setor(ws.title)
    if tipo_logo:
        img_path = _encontrar_logo(tipo_logo, diretorios_busca)

    if not img_path:
        print(
            f"[AVISO] Logo não encontrado para a aba {ws.title}. "
            "Esperado: FM.jpg/ FM.png ou VISION.jpg/ VISION.png"
        )
        return

    limpar_imagens_da_aba(ws)

    try:
        img = colocar_imagem_na_celula(ws, "A1", img_path)
        if img:
            print(f"[INFO] Logo '{img_path.name}' inserido em A1 na aba {ws.title}.")
    except Exception as e:
        print(f"[ERRO] Falha ao inserir imagem na aba {ws.title}: {e}")


# ==========================
# Renomear PDFs
# ==========================

def renomear_pdf(pdf_path, data_entrega, sector, obs, numero_mxm=None):
    """
    Renomeia o PDF para 'DATA - SETOR - MXM <numero> - OBS.pdf' (quando houver MXM)
    e move para a pasta MXM,
    sem sobrescrever arquivos existentes.
    """
    try:
        if data_entrega:
            data_part = data_entrega.replace("/", "-")
        else:
            data_part = datetime.fromtimestamp(pdf_path.stat().st_mtime).strftime("%Y-%m-%d")

        setor_part = sanitize_filename_part(sector) if sector else ""
        obs_part = sanitize_filename_part(obs, max_len=60) if obs else ""
        mxm_part = sanitize_filename_part(f"MXM {numero_mxm}") if numero_mxm else ""

        parts = [p for p in (data_part, setor_part, mxm_part, obs_part) if p]
        if not parts:
            print(f"[AVISO] Dados insuficientes para renomear o PDF {pdf_path.name}.")
            return

        base_name = " - ".join(parts)
        destino = pdf_path.parent / "MXM"
        destino.mkdir(exist_ok=True)
        new_path = destino / (base_name + pdf_path.suffix)

        counter = 1
        while new_path.exists():
            new_path = destino / f"{base_name} ({counter}){pdf_path.suffix}"
            counter += 1

        pdf_path.rename(new_path)
        print(f"[INFO] PDF movido para MXM como: {new_path.name}")
    except Exception as e:
        print(f"[AVISO] Não foi possível mover/renomear o PDF {pdf_path.name}: {e}")


# ==========================
# Exportar abas para PDF
# ==========================

def _exportar_uma_aba_para_pdf(excel_app, workbook, sheet_name, destino, meta=None):
    """Exporta uma aba específica para PDF usando Excel via COM."""
    try:
        ws = workbook.Worksheets(sheet_name)
    except Exception:
        print(f"[AVISO] Não encontrei a aba '{sheet_name}' no Excel para exportar.")
        return

    meta = meta or {}
    data = meta.get("data")
    mxm = meta.get("mxm")

    nome_parts = [sanitize_filename_part(sheet_name) or "ABA"]
    if data:
        nome_parts.append(data.replace("/", "-"))
    if mxm:
        nome_parts.append(f"MXM-{mxm}")

    nome_base = " - ".join(nome_parts)
    destino.mkdir(exist_ok=True)
    pdf_path = destino / f"{nome_base}.pdf"
    counter = 1
    while pdf_path.exists():
        pdf_path = destino / f"{nome_base} ({counter}).pdf"
        counter += 1

    try:
        ws.ExportAsFixedFormat(0, str(pdf_path))
        print(f"[INFO] Aba '{sheet_name}' exportada para PDF em: {pdf_path}")
    except Exception as e:
        print(f"[AVISO] Falha ao exportar a aba '{sheet_name}' para PDF: {e}")


def exportar_abas_para_pdf(xlsx_path, sheet_meta):
    """
    Exporta as abas modificadas para PDF na pasta 'FOLHAS DAS EQUIPES'.
    Usa Excel via COM (win32com), portanto só funciona em Windows com Excel instalado.
    """
    if not sheet_meta:
        return

    if os.name != "nt":
        print("[AVISO] Exportação para PDF requer Windows + Excel; pulando etapa.")
        return

    xlsx_path = Path(xlsx_path).resolve()
    if not xlsx_path.exists():
        print(f"[AVISO] Arquivo Excel para exportação não encontrado: {xlsx_path}")
        return

    destino = xlsx_path.parent / "FOLHAS DAS EQUIPES"

    try:
        import win32com.client  # type: ignore
    except ImportError:
        print("[AVISO] win32com não disponível; exportação para PDF não será executada.")
        return

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb_com = excel.Workbooks.Open(str(xlsx_path))
        for nome, meta in sheet_meta.items():
            _exportar_uma_aba_para_pdf(excel, wb_com, nome, destino, meta)
    finally:
        try:
            wb_com.Close(False)
        except Exception:
            pass
        excel.Quit()


# ==========================
# Backup do Excel
# ==========================

def _backup_path(xlsx_path):
    return xlsx_path.parent / "BACKUP_EXCEL" / f"{xlsx_path.stem} - BACKUP{xlsx_path.suffix}"


def atualizar_backup_excel(xlsx_path):
    """Atualiza backup do Excel principal."""
    try:
        backup = _backup_path(Path(xlsx_path))
        backup.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(xlsx_path, backup)
        print(f"[INFO] Backup atualizado: {backup}")
        return backup
    except Exception as exc:
        print(f"[AVISO] Não foi possível atualizar backup do Excel: {exc}")
        return None


def restaurar_excel_por_backup(xlsx_path):
    """Restaura Excel principal a partir do backup, se disponível."""
    try:
        xlsx_path = Path(xlsx_path)
        backup = _backup_path(xlsx_path)
        if not backup.exists():
            return False
        shutil.copy2(backup, xlsx_path)
        print(f"[AVISO] Excel restaurado a partir do backup: {backup}")
        return True
    except Exception as exc:
        print(f"[ERRO] Falha ao restaurar Excel por backup: {exc}")
        return False


# ==========================
# Fluxo principal
# ==========================

def main():
    parse_args()
    carregar_cache_manual()

    pdf_paths = find_pdfs()
    xlsx_path = find_excel()

    print("\n================ INÍCIO DO PROCESSAMENTO ================")
    print("[INFO] Diretório de PDFs:", PDF_DIR)
    print("[INFO] Excel base:", xlsx_path)
    print("[INFO] PDFs a processar (originais):")
    for p in pdf_paths:
        dt = datetime.fromtimestamp(p.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")
        print(f"  - {p.name} (modificado em {dt})")

    if DRY_RUN:
        print("[MODO SIMULAÇÃO] DRY_RUN=True → Excel NÃO será salvo, PDFs NÃO serão renomeados e exportações serão puladas.")

    try:
        wb = load_workbook(xlsx_path)
    except Exception as exc:
        print(f"[AVISO] Falha ao abrir Excel principal ({xlsx_path}): {exc}")
        if restaurar_excel_por_backup(xlsx_path):
            wb = load_workbook(xlsx_path)
        else:
            raise

    processados = 0
    resumo_setores = defaultdict(int)
    novos_por_pdf = defaultdict(int)  # quantos itens "novos" em cada PDF
    pdfs_processados = []  # lista de dicts: {"path":..., "data":..., "setor":..., "obs":..., "mxm":...}
    abas_modificadas = {}

    for pdf in pdf_paths:
        print("\n" + "=" * 60)
        print(f"[INFO] Processando PDF: {pdf.name}")
        print("=" * 60)

        lines = extract_lines(pdf)
        if not lines:
            print(f"[ERRO] Nenhum texto extraído de {pdf.name}. Pulando este arquivo...")
            continue

        sector = extract_sector(lines)
        data_entrega = extract_entrega_date(lines)
        numero_mxm = extract_numero_requisicao(lines)
        obs = extract_obs(lines)
        items = extract_items(lines)

        # conta itens "novos" (sem regra conhecida)
        novos_count = sum(1 for it in items if it.get("is_new"))
        novos_por_pdf[pdf.name] = novos_count

        # remove flag is_new antes de seguir (resto do fluxo não precisa)
        for it in items:
            it.pop("is_new", None)

        items = processar_itens_especiais(items)
        items = agrupar_por_descricao(items)

        print(f"[INFO] Setor: {sector}")
        print(f"[INFO] Data de entrega: {data_entrega}")
        print(f"[INFO] Número MXM: {numero_mxm}")
        print(f"[INFO] OBS: {obs}")
        print(f"[INFO] Itens após tratamento/agrupamento: {len(items)}")
        print(f"[INFO] Itens novos (sem regra específica): {novos_count}")
        for it in items:
            print(f"  - {it['code']} | {it['desc']} | QNT {it['qty']}")

        if not sector:
            print("[ERRO] Setor não identificado nesse PDF. Pulando...")
            continue

        ws = find_sheet(wb, sector)
        if ws is None:
            print(f"[ERRO] Aba do setor '{sector}' não encontrada. Pulando este PDF...")
            print("Abas existentes:", ", ".join(wb.sheetnames))
            continue

        print(f"[INFO] Aba utilizada: {ws.title}")

        fill_items(ws, items)
        ordenar_itens(ws)
        set_setor_header(ws, sector)
        set_data_entrega(ws, data_entrega)
        set_mxm(ws, numero_mxm)
        set_obs(ws, obs)
        inserir_logotipo(ws)
        info_atual = abas_modificadas.get(ws.title, {})
        if data_entrega and not info_atual.get("data"):
            info_atual["data"] = data_entrega
        if numero_mxm and not info_atual.get("mxm"):
            info_atual["mxm"] = numero_mxm
        abas_modificadas[ws.title] = info_atual

        resumo_setores[sector] += len(items)
        processados += 1

        pdfs_processados.append({
            "path": pdf,
            "data": data_entrega,
            "setor": sector,
            "obs": obs,
            "mxm": numero_mxm
        })

    if processados == 0:
        print("\n[AVISO] Nenhum PDF foi processado. Nada será salvo.")
        print("================= FIM (SEM ALTERAÇÕES) =================")
        return

    if not DRY_RUN:
        try:
            wb.save(xlsx_path)
            atualizar_backup_excel(xlsx_path)
            print("\n[OK] Arquivo Excel atualizado com sucesso:", xlsx_path)

            # Exporta as abas modificadas para PDF
            if SKIP_EXPORT:
                print("[INFO] Exportação de abas para PDF pulada (--skip-export).")
            else:
                exportar_abas_para_pdf(xlsx_path, abas_modificadas)

            # Só agora, depois do Excel salvo, renomeia os PDFs processados
            if SKIP_RENAME:
                print("[INFO] Renomeação de PDFs pulada (--skip-rename).")
            else:
                for info in pdfs_processados:
                    renomear_pdf(info["path"], info["data"], info["setor"], info["obs"], info.get("mxm"))

        except PermissionError:
            print("\n[ERRO] Não foi possível salvar o arquivo original.")
            print("Possível causa: arquivo está aberto no Excel ou bloqueado.")
            input("Feche o arquivo e aperte ENTER para sair...")
    else:
        print("\n[MODO SIMULAÇÃO] DRY_RUN=True → Excel NÃO foi salvo e PDFs NÃO foram renomeados.")

    salvar_cache_manual()

    print("\n=========== RESUMO FINAL POR SETOR ==========")
    for setor, qt in resumo_setores.items():
        print(f"{setor:<10} -> {qt:>3} itens")

    print("\n===== ITENS NOVOS (SEM REGRA ESPECÍFICA) POR PDF =====")
    for pdf_name, qt in novos_por_pdf.items():
        print(f"{pdf_name}: {qt} itens novos")

    print("======================================================")
    print("===================== FIM ====================")


if __name__ == "__main__":
    main()
