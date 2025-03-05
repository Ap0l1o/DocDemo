[Setup]
AppName=��Ŀ��������У�鹤��
AppVersion=1.0
DefaultDirName={autopf}\��Ŀ��������У�鹤��
DefaultGroupName=��Ŀ��������У�鹤��
OutputDir=.
OutputBaseFilename=��Ŀ��������У�鹤��_��װ��
PrivilegesRequired=lowest
SetupIconFile=img\logo.ico

[Languages]
Name: "chinesesimp"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
Source: "dist\doc_processor_gui\*.*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "img\logo.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "dist\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\��Ŀ��������У�鹤��"; Filename: "{app}\doc_processor_gui.exe"; IconFilename: "{app}\logo.jpg"
Name: "{autodesktop}\��Ŀ��������У�鹤��"; Filename: "{app}\doc_processor_gui.exe"; IconFilename: "{app}\logo.jpg"; Tasks: desktopicon

[Run]
Filename: "{app}\doc_processor_gui.exe"; Description: "{cm:LaunchProgram,��Ŀ��������У�鹤��}"; Flags: nowait postinstall skipifsilent