# Instalação e uso do programa MXM

## Pré-requisitos

- Python 3.10+ instalado
- No Windows: Excel instalado (se quiser exportar abas para PDF)

## Instalação rápida

### Windows

1. Abra o Prompt de Comando na pasta do projeto.
2. Execute:

```bat
scripts\instalar_windows.bat
```

Depois disso, use:

```bat
executar_gui.bat
```

Se ocorrer erro, verifique o arquivo `instalar_mxm.log` na raiz do projeto.

### Linux/macOS

1. Abra o terminal na pasta do projeto.
2. Execute:

```bash
./scripts/instalar_linux.sh
```

Depois disso, use:

```bash
./executar_gui.sh
```

## Instalação manual (alternativa)

```bash
python -m venv .venv_mxm
```

Ativar venv:

- Windows:

```bat
.venv_mxm\Scripts\activate
```

- Linux/macOS:

```bash
source .venv_mxm/bin/activate
```

Instalar dependências:

```bash
pip install -r requirements-mxm.txt
```

## Uso via CLI

```bash
python preencher_planilha_normalizado.py --help
```

Exemplo:

```bash
python preencher_planilha_normalizado.py --pdf-dir "/caminho/pdfs" --excel "/caminho/base.xlsx"
```

## Observações

- A GUI executa em modo não interativo (`--non-interactive`) para não travar esperando input no terminal.
- Para logos, mantenha `FM.jpg` e `VISION.jpg` na pasta de trabalho ou em `--pdf-dir`.
