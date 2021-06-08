Attribute VB_Name = "mSendKeys"

Option Explicit
Declare Function ShellAbout Lib "shell32" Alias "ShellAboutA" _
                            (ByVal hWnd As Long, _
                            ByVal szApp As String, _
                            ByVal szOtherStuff As String, _
                            ByVal hIcon As Long) As Long

Public Const VK_TAB = &H9
Declare Sub keybd_event Lib "user32" ( _
ByVal bVk As Byte, ByVal bScan As Byte, _
ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Const KEYEVENTF_KEYUP = &H2
Private Const INPUT_KEYBOARD = 1
Private Type KEYBDINPUT
wVk As Integer
wScan As Integer
dwFlags As Long
time As Long
dwExtraInfo As Long
End Type
Private Type GENERALINPUT
dwType As Long
xi(0 To 23) As Byte
End Type
Private Declare Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As GENERALINPUT, ByVal cbSize As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Public Function SendKeysA(ByVal vKey As Integer, Optional booDown As Boolean = False)
Dim GInput(0) As GENERALINPUT
Dim KInput As KEYBDINPUT
KInput.wVk = vKey
If Not booDown Then
    KInput.dwFlags = KEYEVENTF_KEYUP
End If
GInput(0).dwType = INPUT_KEYBOARD
CopyMemory GInput(0).xi(0), KInput, Len(KInput)
Call SendInput(1, GInput(0), Len(GInput(0)))
End Function
