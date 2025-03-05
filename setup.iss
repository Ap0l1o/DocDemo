[Setup]
AppName=项目开发类金额校验工具
AppVersion=1.0
DefaultDirName={autopf}\项目开发类金额校验工具
DefaultGroupName=项目开发类金额校验工具
OutputDir=.
OutputBaseFilename=项目开发类金额校验工具_安装包
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
Name: "{group}\项目开发类金额校验工具"; Filename: "{app}\doc_processor_gui.exe"; IconFilename: "{app}\logo.jpg"
Name: "{autodesktop}\项目开发类金额校验工具"; Filename: "{app}\doc_processor_gui.exe"; IconFilename: "{app}\logo.jpg"; Tasks: desktopicon

[Run]
Filename: "{app}\doc_processor_gui.exe"; Description: "{cm:LaunchProgram,项目开发类金额校验工具}"; Flags: nowait postinstall skipifsilent