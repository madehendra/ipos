VERSION 5.00
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Begin VB.Form cfgLimitPencarian 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Limit Pencarian"
   ClientHeight    =   1155
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3915
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   3915
   Begin BiSANumberBoxProject.BiSANumberBox nLimitPencarian 
      Height          =   360
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   210
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   635
      Appearance      =   0
      Decimals        =   0
      MinValue        =   1
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   " Limit Pencarian"
      CaptionWidth    =   1500
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Records"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2850
      TabIndex        =   1
      Top             =   270
      Width           =   780
   End
End
Attribute VB_Name = "cfgLimitPencarian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
  CenterForm Me
  nLimitPencarian.Value = GetRegistry(reg_LimitPencarian)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SaveRegistry reg_LimitPencarian, IIf(nLimitPencarian.Value < 10, 10, nLimitPencarian.Value)
End Sub

