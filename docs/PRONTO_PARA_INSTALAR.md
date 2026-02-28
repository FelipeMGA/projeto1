# Pronto para instalar (1 passo)

## Windows (mais fácil)

Dê duplo clique em:

- `instalar_mxm.bat`

Ou no terminal:

```bat
instalar_mxm.bat
```

## Linux/macOS

No terminal:

```bash
./instalar_mxm.sh
```

## O que o instalador faz

1. Cria ambiente virtual `.venv_mxm`
2. Instala dependências (`requirements-mxm.txt`)
3. (Windows) tenta instalar `pywin32` para exportação PDF via Excel
4. Cria atalhos de execução da GUI (`executar_gui.bat` / `executar_gui.sh`)
5. Abre a GUI automaticamente ao final

## Se quiser só instalar (sem abrir GUI)

```bash
python scripts/install_mxm_tool.py --repo-dir .
```

## Se quiser reabrir depois

- Windows: `executar_gui.bat`
- Linux/macOS: `./executar_gui.sh`

## Se a janela do Windows fechar sozinha

Agora o `instalar_mxm.bat` **fica aberto ao final** e grava log em:

- `instalar_mxm.log` (na raiz do projeto)

Se der erro, abra esse arquivo e me envie o conteúdo.

- O instalador Windows agora detecta automaticamente a pasta do projeto (evita erro com caminhos acentuados do OneDrive).


> Correção aplicada: caminhos com espaço/acento (ex.: OneDrive + Área de Trabalho) agora são tratados corretamente no instalador Windows.

- O instalador principal (`instalar_mxm.bat`) agora passa `--repo-dir` explicitamente para evitar ambiguidades de diretório no Windows.


## Instalar como programa do Windows (Arquivos de Programas)

Se você quer instalar como software "normal" (com pasta em **Arquivos de Programas**):

1. Instale o **Inno Setup 6** no Windows.
2. No PowerShell, na raiz do projeto, execute:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/build_windows_installer.ps1
```

3. O setup será gerado a partir de `installer/windows/MXM_Setup.iss`.
4. Execute o `MXM_Processador_Setup.exe` gerado e siga o assistente.

Esse instalador cria entrada no menu iniciar, opção de desinstalar e instalação em Program Files.


## Pastas padrão e backup automático

A instalação agora cria automaticamente dentro da pasta do programa:

- `DADOS_MXM/ENTRADA_PDFS`
- `DADOS_MXM/MXM`
- `DADOS_MXM/FOLHAS DAS EQUIPES`
- `DADOS_MXM/BASE_EXCEL`
- `DADOS_MXM/BACKUP_EXCEL`

No Windows também são criados atalhos na área de trabalho para abrir:

- Pasta MXM
- Pasta Folhas das Equipes
- Pasta Base Excel

O Excel principal terá backup automático atualizado em `DADOS_MXM/BACKUP_EXCEL` após salvar.
Se o Excel principal estiver corrompido, o programa tenta restaurar pelo backup automaticamente.
