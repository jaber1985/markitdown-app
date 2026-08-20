[Setup]
AppName=MarkItDown
AppVersion=1.0
AppPublisher=MarkItDown
DefaultDirName={autopf}\MarkItDown
DefaultGroupName=MarkItDown
OutputDir=installer
OutputBaseFilename=MarkItDown-Setup
Compression=lzma2
SolidCompression=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Files]
Source: "dist\MarkItDown.exe"; DestDir: "{app}"

[Icons]
Name: "{group}\MarkItDown"; Filename: "{app}\MarkItDown.exe"
Name: "{autodesktop}\MarkItDown"; Filename: "{app}\MarkItDown.exe"

[Run]
Filename: "{app}\MarkItDown.exe"; Description: "Launch MarkItDown"; Flags: postinstall nowait skipifsilent
