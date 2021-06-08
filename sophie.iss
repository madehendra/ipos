; -- Example1.iss --
; Demonstrates copying 3 files and creating an icon.

; SEE THE DOCUMENTATION FOR DETAILS ON CREATING .ISS SCRIPT FILES!

[Setup]
AppName=Update Program MiPOS
appvername=Null101
DefaultDirName=d:\mipos\
DisableDirPage = yes
SolidCompression=yes
Uninstallable = no
OutputBaseFilename="Update Mipos"

[Files]
;Source: "sophie.exe"; DestDir: "d:\xampp\program sophie";Flags: ignoreversion
Source: "D:\Dropbox\ack center rubah faktur pelunasan\mipos.exe"; DestDir: "{app}"; Flags: ignoreversion
;Source: "myodbc5a.dll"; DestDir: "{sys}"
;Source: "myodbc5a.lib"; DestDir: "{sys}"

