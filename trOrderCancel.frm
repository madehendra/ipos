VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trOrderCancel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PERHATIAN : Order Cancellation"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   7590
   Begin BiSADateProject.BiSADate dTgl 
      Height          =   330
      Left            =   3135
      TabIndex        =   1
      Top             =   1140
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   582
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
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
   Begin BiSAButtonProject.BiSAButton BiSAButton1 
      Height          =   495
      Left            =   2325
      TabIndex        =   0
      Top             =   1515
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   873
      Caption         =   "ORDER CANCEL"
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
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   45
      TabIndex        =   2
      Top             =   45
      Width           =   7410
   End
End
Attribute VB_Name = "trOrderCancel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub BiSAButton1_Click()
  Set dbData = objData.Browse(GetDSN, "po", , "statuspembelian", sisAssign, 0, " and tgl = '" & Format(dTgl.Value, "yyyy-MM-dd") & "'")
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.EOF
      FrmPB.RunPB
      objData.Update GetDSN, "po", "id = " & GetNull(dbData!ID), Array("tgl", "statusorder"), Array("0000-00-00", 0)
      dbData.MoveNext
    Loop
    FrmPB.EndPB
  Else
    MsgBox "Maaf, tidak ada data untuk tanggal tersebut!!"
  End If
End Sub

Private Sub Form_Load()
Dim cCaption As String
  
  SetIcon Me.hWnd, "SIKD"
  CenterForm Me

  cCaption = "Perhatian: Silahkan pilih tanggal order yg hendak dibatalkan!!" & vbCrLf
  cCaption = cCaption & " Order yg bisa di cancel hanya order yg belum di proses oleh bagian PEMBELIAN " & vbCrLf
  cCaption = cCaption & " Order yg telah sukses di cancel akan ikut menjadi satu dengan order pending (apabila ada)"
  Label1.Caption = cCaption
End Sub
