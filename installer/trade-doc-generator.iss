#pragma codepage 65001

#define MyAppName "外贸单据自动生成工具"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "CPH"
#define MyAppExeName "TradeDocGenerator.exe"

[Setup]
AppId={{92B3BB0B-C541-4E6A-92B0-0C20F55677D4}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\TradeDocGenerator
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=admin
OutputDir=..\dist-installer
OutputBaseFilename=TradeDocGenerator_Setup_{#MyAppVersion}
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: checkedonce

[Files]
Source: "..\dist\TradeDocGenerator\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\resources\installers\LibreOffice.msi"; DestDir: "{tmp}"; Flags: deleteafterinstall ignoreversion

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "msiexec.exe"; Parameters: "/i ""{tmp}\LibreOffice.msi"" /qn /norestart"; StatusMsg: "Installing LibreOffice..."; Flags: waituntilterminated; Check: not LibreOfficeInstalled
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[Code]
function LibreOfficeInstalled(): Boolean;
begin
  Result :=
    FileExists(ExpandConstant('{autopf}\LibreOffice\program\soffice.exe')) or
    FileExists(ExpandConstant('{commonpf}\LibreOffice\program\soffice.exe')) or
    FileExists(ExpandConstant('{commonpf32}\LibreOffice\program\soffice.exe'));
end;
