; Inno Setup script for Office MCP Server
; Produces a Windows installer from the PyInstaller --onedir output.
;
; Expected directory layout at compile time:
;   .github/mcp/dist/office-mcp-server/   <- PyInstaller --onedir output
;   .github/mcp/installer.iss             <- this file
;
; Usage (CI or local):
;   iscc /DAppVersion=1.0.0 installer.iss

#ifndef AppVersion
  #define AppVersion "0.0.0-dev"
#endif

[Setup]
AppName=Office MCP Server
AppVersion={#AppVersion}
AppPublisher=Architect-in-a-Box
DefaultDirName={autopf}\OfficeMCPServer
DefaultGroupName=Office MCP Server
OutputDir=dist
OutputBaseFilename=office-mcp-server-setup
Compression=lzma2/ultra64
SolidCompression=yes
ArchitecturesAllowed=x64compatible
; No admin rights needed — installs per-user by default
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
SetupIconFile=
UninstallDisplayIcon={app}\office-mcp-server.exe
WizardStyle=modern
DisableProgramGroupPage=yes
; Allow silent install for automated deployments
AllowNoIcons=yes

[Files]
; Bundle the entire PyInstaller --onedir output
Source: "dist\office-mcp-server\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; No Start Menu shortcut needed for a CLI/server tool, but add one for
; easy access to the install directory.
Name: "{group}\Office MCP Server Directory"; Filename: "{app}"

[Registry]
; Add install directory to user PATH so VS Code and terminals can find the exe
Root: HKCU; Subkey: "Environment"; ValueType: expandsz; ValueName: "Path"; \
  ValueData: "{olddata};{app}"; Check: NeedsAddPath(ExpandConstant('{app}'))

[Code]
function NeedsAddPath(Param: string): Boolean;
var
  OrigPath: string;
begin
  if not RegQueryStringValue(HKEY_CURRENT_USER,
    'Environment', 'Path', OrigPath) then
  begin
    Result := True;
    exit;
  end;
  { Look for the path in the existing value (case-insensitive) }
  Result := Pos(';' + Uppercase(Param) + ';',
    ';' + Uppercase(OrigPath) + ';') = 0;
end;

[UninstallDelete]
Type: filesandordirs; Name: "{app}"
