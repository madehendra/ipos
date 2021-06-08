VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Begin VB.Form frmKonversi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form Konversi Data"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4035
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   4035
   Begin BiSAButtonProject.BiSAButton BiSAButton1 
      Height          =   495
      Left            =   285
      TabIndex        =   0
      Top             =   270
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   873
      Caption         =   "Konversi Simpanan Wajib"
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
End
Attribute VB_Name = "frmKonversi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BiSAButton1_Click()
  Load frmKonversiSimpananWajib
  frmKonversiSimpananWajib.Show
End Sub

Private Sub Form_Load()
  CenterForm Me
  SetIcon Me.hWnd
End Sub
