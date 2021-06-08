VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trNeracaBalance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Balance Neraca"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   5550
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1650
      Left            =   0
      Top             =   15
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2910
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      Begin BiSAButtonProject.BiSAButton cmdCek 
         Height          =   435
         Left            =   1605
         TabIndex        =   2
         Top             =   675
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   767
         Caption         =   "Cek"
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
      Begin BiSADateProject.BiSADate dTgl 
         Height          =   330
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   210
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   582
         Appearance      =   0
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
         Caption         =   "Tanggal"
         CaptionWidth    =   1400
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
      Begin BiSADateProject.BiSADate dTgl 
         Height          =   330
         Index           =   1
         Left            =   3075
         TabIndex        =   1
         Top             =   210
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   582
         Appearance      =   0
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
         Caption         =   "sda"
         CaptionWidth    =   0
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
End
Attribute VB_Name = "trNeracaBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub GetFakturBermasalah()
Dim n As Integer

  vaArray.ReDim 0, -1, 0, 2
  Set dbData = objData.Browse(GetDSN, "bukubesar", "DISTINCT(faktur) as faktur,tgl", "tgl", sisGTEqual, Format(dTgl(0).Value, "yyyy-MM-dd"), " AND tgl <= '" & Format(dTgl(1).Value, "yyyy-MM-dd") & "'")
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.EOF
      FrmPB.RunPB
      If GetSelisih(GetNull(dbData!Faktur)) <> 0 Then
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        n = vaArray.UpperBound(1)
        vaArray(n, 0) = "'" & GetNull(dbData!Faktur)
        vaArray(n, 1) = GetSelisih(GetNull(dbData!Faktur))
        vaArray(n, 2) = Format(GetNull(dbData!tgl), "dd-MM-yyyy")
      End If
      dbData.MoveNext
    Loop
    FrmPB.EndPB
  End If
  
  Dim a As New exportExcel
  a.RecordSource = vaArray
  a.ExportToExcel
End Sub

Private Function GetFaktur()

End Function

Private Function GetSelisih(ByVal cFaktur As String) As Double
Dim db As New ADODB.Recordset
  
  GetSelisih = 0
  Set db = objData.Browse(GetDSN, "bukubesar", "sum(debet-kredit) as saldo", "faktur", sisAssign, cFaktur)
  If Not db.EOF Then
    GetSelisih = GetNull(db!saldo)
  End If
End Function

Private Sub cmdCek_Click()
  GetFakturBermasalah
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  dTgl(0).Value = Date
  dTgl(1).Value = Date
  TabIndex dTgl(0), n
  TabIndex dTgl(1), n
  TabIndex cmdCek, n
End Sub
