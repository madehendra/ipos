Attribute VB_Name = "modTray"
''''''''            ''''''''        ''''     ''''
'''    ''           '''    ''         '''   '''
'''    '' '''    '' '''    '' ''''''''  '''''
''''''''   ''    '' ''''''''  '''    ''  '''
'''        ''    '' '''       '''    ''  '''
'''          '''' ' '''       ''''''''   '''
                              '''
                              '''
'''''''''''''''''''''''''''''''''''''''''''''''''''
'This module calls the necessary api functions    '
'to add/modify/remove an icon to the system tray, '
'display a balloon tip on said icon, receive user '
'input on the icon and balloon tip                '
'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
'Programed by Frankie Miklos -aKa- PuPpY          '
'Credit: http://vbnet.mvps.org/                   '
'Last update: Thursday, June 24, 2004 (13:29)     '
'''''''''''''''''''''''''''''''''''''''''''''''''''


Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Any) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const GWL_WNDPROC As Long = (-4)
Public Const GWL_HWNDPARENT As Long = (-8)
Public Const GWL_ID As Long = (-12)
Public Const GWL_STYLE As Long = (-16)
Public Const GWL_EXSTYLE As Long = (-20)
Public Const GWL_USERDATA As Long = (-21)

Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const NIF_STATE = &H8
Public Const NIF_INFO = &H10

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIM_SETFOCUS = &H3
Public Const NIM_SETVERSION = &H4
Public Const NIM_VERSION = &H5

Public Const NIS_HIDDEN = &H1
Public Const NIS_SHAREDICON = &H2

Public Const WM_USER As Long = &H400
Public Const WM_MYHOOK As Long = WM_USER + 1
Public Const WM_NOTIFY As Long = &H4E
Public Const WM_COMMAND As Long = &H111
Public Const WM_CLOSE As Long = &H10
Public Const WM_MOUSEMOVE As Long = &H200
Public Const WM_LBUTTONDOWN As Long = &H201
Public Const WM_LBUTTONUP As Long = &H202
Public Const WM_LBUTTONDBLCLK As Long = &H203
Public Const WM_MBUTTONDOWN As Long = &H207
Public Const WM_MBUTTONUP As Long = &H208
Public Const WM_MBUTTONDBLCLK As Long = &H209
Public Const WM_RBUTTONDOWN As Long = &H204
Public Const WM_RBUTTONUP As Long = &H205
Public Const WM_RBUTTONDBLCLK As Long = &H206

Public Const NIN_BALLOONSHOW = (WM_USER + 2)
Public Const NIN_BALLOONHIDE = (WM_USER + 3)
Public Const NIN_BALLOONTIMEOUT = (WM_USER + 4)
Public Const NIN_BALLOONUSERCLICK = (WM_USER + 5)

Public Enum bFlag
    NIIF_NONE = &H0
    NIIF_INFO = &H1
    NIIF_WARNING = &H2
    NIIF_ERROR = &H3
    NIIF_GUID = &H5
    NIIF_ICON_MASK = &HF
    NIIF_NOSOUND = &H10
End Enum

Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 128
    dwState As Long
    dwStateMask As Long
    szInfo As String * 256
    uTimeoutAndVersion As Long
    szInfoTitle As String * 64
    dwInfoFlags As Long
End Type
   
Global ni As NOTIFYICONDATA
Global lWP As Long


Private Sub UnSubClass(hWnd As Long)

   If lWP <> 0 Then
      SetWindowLong hWnd, GWL_WNDPROC, lWP
      lWP = 0
   End If
   
End Sub

Private Sub SubClass(hWnd As Long)

   On Error Resume Next
   lWP = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
   
End Sub

Public Sub TrayAdd(PB As PictureBox)
   
   With ni
      .cbSize = Len(ni)
      .hWnd = PB.hWnd
      .uId = 1
      .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
      .dwState = NIS_SHAREDICON
      .hIcon = PB.Picture
      .uCallBackMessage = WM_MYHOOK
      
      .szTip = "Tooltip title" & vbNullChar
      .uTimeoutAndVersion = NOTIFYICON_VERSION
   End With
   
   Shell_NotifyIcon NIM_ADD, ni
   
   SubClass PB.hWnd
       
End Sub

Public Sub TrayRemove(PB As PictureBox)
      
   With ni
      .cbSize = Len(ni)
      .hWnd = PB.hWnd
      .uId = 1
   End With
   
   Shell_NotifyIcon NIM_DELETE, ni
   
   UnSubClass PB.hWnd

End Sub

Public Sub TrayBalloon(PB As PictureBox, bTitle As String, bText As String, ByVal bFlag As bFlag)
   
   With ni
      .cbSize = Len(ni)
      .hWnd = PB.hWnd
      .uId = 1
      .uFlags = NIF_INFO
      .dwInfoFlags = bFlag

      .szInfoTitle = bTitle & vbNullChar
      .szInfo = bText & vbNullChar
   End With

   Shell_NotifyIcon NIM_MODIFY, ni

End Sub

Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    On Error Resume Next
    Select Case hWnd
        Case frmBalloon.pbTray.hWnd
            Select Case uMsg
                Case WM_MYHOOK
                    Select Case lParam
                        Case WM_LBUTTONUP
                            'MsgBox "User has left-clicked the system tray icon", vbInformation, "Information"
                        Case WM_RBUTTONUP
                            'MsgBox "User has right-clicked the system tray icon", vbInformation, "Information"
                        Case NIN_BALLOONSHOW
                            'Msgbox "The balloon tip has just been displayed", vbInformation, "Information"
                        Case NIN_BALLOONHIDE
                            'MsgBox "The systray icon was removed when the balloon tip was displayed", vbInformation, "Information"
                        Case NIN_BALLOONUSERCLICK
                            MsgBox "User clicked the balloon tip", vbInformation, "Information"
                        Case NIN_BALLOONTIMEOUT
                            'MsgBox "The balloon tip either timed out or user clicked the close button", vbInformation, "Information"
                        Case WM_MOUSEMOVE
                            'MsgBox "User moved mouse over icon"
                    End Select
                Case Else
                    WindowProc = CallWindowProc(lWP, hWnd, uMsg, wParam, lParam)
                    Exit Function
            End Select
        Case Else
            WindowProc = CallWindowProc(lWP, hWnd, uMsg, wParam, lParam)
    End Select
   
End Function


