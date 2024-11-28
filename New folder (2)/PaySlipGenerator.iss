[Setup]
AppName=Pay Slip Generator
AppVersion=1.0
DefaultDirName={pf}\Pay Slip Generator
DefaultGroupName=Pay Slip Generator
OutputBaseFilename=PaySlipGeneratorInstaller
Compression=lzma
SolidCompression=yes
LicenseFile=license.txt

[Tasks]
Name: desktopicon; Description: "Create a &desktop icon"; GroupDescription: "Additional icons:"

[Files]
Source: "dist\PaySlipGenerator.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "assets\template.docx"; DestDir: "{app}\assets"; Flags: ignoreversion
Source: "license.txt"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Pay Slip Generator"; Filename: "{app}\PaySlipGenerator.exe"
Name: "{group}\Uninstall Pay Slip Generator"; Filename: "{uninstallexe}"
Name: "{userdesktop}\Pay Slip Generator"; Filename: "{app}\PaySlipGenerator.exe"; Tasks: desktopicon
