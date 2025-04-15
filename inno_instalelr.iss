; Inno Setup Script for PDF Watcher
[Setup]
AppName=PDF Watcher
AppVersion=1.0
DefaultDirName={pf}\PDFWatcher
DefaultGroupName=PDF Watcher
OutputDir=installer
OutputBaseFilename=PDFWatcherSetup
Compression=lzma
SolidCompression=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "dist\main.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "dist\SumatraPDF-3.5.2-64.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "dist\tesseract-ocr-w64-setup-5.5.0.20241111.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "dist\tray_icon.ico"; DestDir: "{app}"; Flags: ignoreversion


[Icons]
Name: "{group}\PDF Watcher"; Filename: "{app}\main.exe"; WorkingDir: "{app}"
Name: "{commondesktop}\PDF Watcher"; Filename: "{app}\main.exe"; WorkingDir: "{app}"

[Run]
Filename: "{app}\main.exe"; Description: "Run PDF Watcher"; Flags: nowait postinstall skipifsilent
