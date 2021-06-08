VERSION 5.00
Begin VB.Form frmBalloon 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System Tray Balloon Tip Example"
   ClientHeight    =   2445
   ClientLeft      =   11295
   ClientTop       =   9660
   ClientWidth     =   4500
   Icon            =   "frmBalloon.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   225
      TabIndex        =   6
      Text            =   "Title"
      Top             =   200
      Width           =   4095
   End
   Begin VB.PictureBox pbTray 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   4005
      Picture         =   "frmBalloon.frx":0CCA
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      Top             =   1483
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtMsg 
      Height          =   1215
      Left            =   210
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmBalloon.frx":1254
      Top             =   568
      Width           =   4095
   End
   Begin VB.CommandButton cmdBalloon 
      Caption         =   "Normal"
      Height          =   360
      Index           =   0
      Left            =   180
      TabIndex        =   1
      Top             =   1873
      Width           =   975
   End
   Begin VB.CommandButton cmdBalloon 
      Caption         =   "Info"
      Height          =   360
      Index           =   1
      Left            =   1235
      TabIndex        =   2
      Top             =   1888
      Width           =   975
   End
   Begin VB.CommandButton cmdBalloon 
      Caption         =   "Warning"
      Height          =   360
      Index           =   2
      Left            =   2290
      TabIndex        =   3
      Top             =   1888
      Width           =   975
   End
   Begin VB.CommandButton cmdBalloon 
      Caption         =   "Error"
      Height          =   360
      Index           =   3
      Left            =   3345
      TabIndex        =   4
      Top             =   1888
      Width           =   975
   End
End
Attribute VB_Name = "frmBalloon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''            ''''''''        ''''     ''''
'''    ''           '''    ''         '''   '''
'''    '' '''    '' '''    '' ''''''''  '''''
''''''''   ''    '' ''''''''  '''    ''  '''
'''        ''    '' '''       '''    ''  '''
'''          '''' ' '''       ''''''''   '''
                              '''
                              '''
'''''''''''''''''''''''''''''''''''''''''''''''''''
'This is the main form where you may change the   '
'icon, title, and message of the balloon tip      '
'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
'Programed by Frankie Miklos -aKa- PuPpY          '
'Credit: http://vbnet.mvps.org/                   '
'Last update: Thursday, June 24, 2004 (13:29)     '
'''''''''''''''''''''''''''''''''''''''''''''''''''


Private Sub cmdBalloon_Click(Index As Integer)

    TrayBalloon pbTray, txtTitle.Text, txtMsg.Text, Index

End Sub

Private Sub Form_Load()
Dim tmp
    
    tmp = RegRead(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\", "EnableBalloonTips")
    
    If tmp = 0 Then
        If MsgBox("Balloon tips are currently disabled on your computer. Would you like to enable them?", vbQuestion + vbYesNo, "Enable Balloon Tips?") = vbYes Then
            WriteDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\", "EnableBalloonTips", 1
            If MsgBox("Balloon tips are now enabled, but you must first logoff your computer" & vbNewLine & "and then log back on before the changes will take effect." & vbNewLine & vbNewLine & "Would you like to be logged off now?", vbQuestion + vbYesNo, "Logoff Now?") = vbYes Then
                LogOffNT True
                End
            End If
        Else
            MsgBox "Without balloon tips enabled on your computer, this program will not function properly.", vbExclamation, "Balloon Tips Disabled"
        End If
    End If
                
    TrayAdd pbTray
   
End Sub

Private Sub Form_Unload(Cancel As Integer)

   TrayRemove pbTray
   
End Sub
