; -- Example1.iss --
; Demonstrates copying 3 files and creating an icon.

; SEE THE DOCUMENTATION FOR DETAILS ON CREATING .ISS SCRIPT FILES!

[Setup]
AppName=Sophie BC Made Hendra 417
AppVersion=1
DefaultDirName= c:\sophie
DefaultGroupName=Sophie BC Made Hendra 417
UninstallDisplayIcon={app}\sophie.exe
Compression=lzma2
SolidCompression=yes
;OutputDir=userdocs:Inno Setup Examples Output

[Files]
Source: "sophie.exe"; DestDir: "{app}"

[Icons]
Name: "{group}\Sophie BC Made Hendra 417"; Filename: "{app}\sophie.exe"
