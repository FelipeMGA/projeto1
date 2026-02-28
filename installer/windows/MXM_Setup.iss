; Inno Setup script - gera instalador padrão para Program Files
[Setup]
AppId={{8E0EE47A-7698-4C3E-B36F-5CC59B7B7612}
AppName=MXM Processador
AppVersion=1.0.0
AppPublisher=VISION
DefaultDirName={autopf}\MXM Processador
DefaultGroupName=MXM Processador
UninstallDisplayIcon={app}\gui_preencher_planilha.py
Compression=lzma
SolidCompression=yes
OutputBaseFilename=MXM_Processador_Setup
WizardStyle=modern

[Files]
Source: "..\..\preencher_planilha_normalizado.py"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\..\gui_preencher_planilha.py"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\..\requirements-mxm.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\..\scripts\install_mxm_tool.py"; DestDir: "{app}\scripts"; Flags: ignoreversion
Source: "..\..\FM.jpg"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "..\..\VISION.jpg"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist

[Icons]
Name: "{group}\MXM Processador"; Filename: "{cmd}"; Parameters: "/k \"cd /d \"\"{app}\"\" && py scripts\\install_mxm_tool.py --repo-dir . --run-gui\""
Name: "{group}\Desinstalar MXM Processador"; Filename: "{uninstallexe}"
Name: "{autodesktop}\MXM Processador"; Filename: "{cmd}"; Parameters: "/k \"cd /d \"\"{app}\"\" && py scripts\\install_mxm_tool.py --repo-dir . --run-gui\""; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Criar atalho na área de trabalho"; GroupDescription: "Atalhos:"; Flags: unchecked

[Run]
Filename: "{cmd}"; Parameters: "/k \"cd /d \"\"{app}\"\" && py scripts\\install_mxm_tool.py --repo-dir . --run-gui\""; Description: "Executar configuração inicial e abrir GUI"; Flags: postinstall
