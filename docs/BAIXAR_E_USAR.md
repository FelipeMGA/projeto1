# Baixar e usar (pronto para download)

## 1) Gerar pacote ZIP pronto

Na raiz do projeto:

```bash
python scripts/build_download_package.py
```

Isso gera um ZIP em `dist/` com nome parecido com:

- `MXM_PRONTO_PARA_DOWNLOAD_YYYYMMDD_HHMM.zip`

## 2) Compartilhar / download

- Faça upload do ZIP gerado para Google Drive, OneDrive, SharePoint ou anexo interno.
- Quem baixar só precisa extrair o arquivo.

## 3) Instalar em 1 passo após baixar

### Windows

Abra a pasta extraída e dê duplo clique em:

- `instalar_mxm.bat`

### Linux/macOS

No terminal, dentro da pasta extraída:

```bash
./instalar_mxm.sh
```

## 4) Abrir depois da instalação

- Windows: `executar_gui.bat`
- Linux/macOS: `./executar_gui.sh`
