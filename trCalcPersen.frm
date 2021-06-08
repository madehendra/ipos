VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Begin VB.Form trCalcPersen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kalkulasi Discount"
   ClientHeight    =   930
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   2430
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   930
   ScaleWidth      =   2430
   StartUpPosition =   3  'Windows Default
   Begin BiSAButtonProject.BiSAButton cmdOK 
      Height          =   330
      Left            =   1725
      TabIndex        =   2
      Top             =   240
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   582
      Caption         =   "OK"
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
   End
   Begin BiSANumberBoxProject.BiSANumberBox nPersen 
      Height          =   330
      Left            =   300
      TabIndex        =   0
      Top             =   240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   " "
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
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1455
      TabIndex        =   1
      Top             =   285
      Width           =   255
   End
End
Attribute VB_Name = "trCalcPersen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
  trKasir.nDiscountBayar.Value = trKasir.nSubTotal.Value * nPersen.Value / 100
  trKasir.nTotal.Value = trKasir.nSubTotal.Value - trKasir.nDiscountBayar.Value
  trKasir.nKembali.Value = trKasir.nTunai.Value - trKasir.nTotal.Value
  Unload Me
End Sub

Private Sub Form_Load()
  SetIcon Me.hWnd
  CenterForm Me
End Sub
