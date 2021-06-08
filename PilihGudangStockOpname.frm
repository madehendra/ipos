VERSION 5.00
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Begin VB.Form PilihGudangStockOpname 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pilih Gudang..."
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   3735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OKEY!!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   90
      TabIndex        =   0
      Top             =   630
      Width           =   3525
   End
   Begin BiSATextBoxProject.BiSABrowse cKodePoli 
      Height          =   330
      Left            =   90
      TabIndex        =   1
      Top             =   120
      Width           =   3525
      _ExtentX        =   6218
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
      FontName        =   "Verdana"
      Appearance      =   0
      Button          =   -1  'True
      Caption         =   "Gudang"
      CaptionWidth    =   2000
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
End
Attribute VB_Name = "PilihGudangStockOpname"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data

Private Sub cKodePoli_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "gudang")
  If Not dbData.EOF Then
    cKodePoli.Text = cKodePoli.Browse(dbData)
  End If
End Sub

Private Sub cmdOK_Click()
  trStockOpname.Gudang = cKodePoli.Text
  Unload Me
  Load trStockOpname
  trStockOpname.Show
  trStockOpname.SetFocus
End Sub

Private Sub Form_Load()
Dim n As Single
  
  SetIcon Me.hWnd
  CenterForm Me
  TabIndex cKodePoli, n
  TabIndex cmdOK, n
End Sub


