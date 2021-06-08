VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trKirimOrderanPromo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kirim Orderan Promo"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4665
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   4665
   Begin BiSAButtonProject.BiSAButton cmdOK 
      Height          =   450
      Left            =   510
      TabIndex        =   2
      Top             =   1110
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   794
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
   Begin BiSADateProject.BiSADate dTgl 
      Height          =   345
      Index           =   0
      Left            =   465
      TabIndex        =   0
      Top             =   345
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   609
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
   Begin BiSADateProject.BiSADate dTgl 
      Height          =   345
      Index           =   1
      Left            =   2220
      TabIndex        =   1
      Top             =   345
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   609
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
End
Attribute VB_Name = "trKirimOrderanPromo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vaArray As New XArrayDB
Dim objData As New CodeSuiteLibrary.Data
Dim dbData As New ADODB.Recordset

Private Sub initvalue()
  dTgl(0).Value = BOM(Date)
  dTgl(1).Value = Date
End Sub

Private Sub cmdOK_Click()
  GetSQL
End Sub

Private Sub Form_Load()
Dim n As Single
  
  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  initvalue
  TabIndex dTgl(0), n
  TabIndex dTgl(1), n
  TabIndex cmdOK, n
End Sub


Private Function GetLastDBUpdate(obj As CodeSuiteLibrary.Data) As String
Dim db As New ADODB.Recordset

    Set db = obj.Browse(GetDSN, "bukubesar", "max(datetime) as lasupdate")
    GetLastDBUpdate = " DB Ver.." & Format(GetNull(db!lasupdate), "dd.MM.yy HH:MM:SS") & " App Ver." & App.Major & "." & App.Minor & "." & App.Revision

End Function

Private Sub GetSQL()
Dim n As Single
Dim cSQL As String
Dim a As New exportExcel

  cSQL = ""
  cSQL = "select s.barcode,sum(m.qty) as qty ,sum(m.qty*m.harga) as total"
  cSQL = cSQL & " from memberorder m"
  cSQL = cSQL & " LEFT JOIN stock s on s.kodestock = m.kodestock"
  
  cSQL = cSQL & " where m.tgl >= '" & Format(dTgl(0).Value, "yyyy-MM-dd") & "' AND m.tgl <= '" & Format(dTgl(1).Value, "yyyy-MM-dd") & "'"
  
  cSQL = cSQL & " GROUP BY s.barcode"
  cSQL = cSQL & " ORDER BY qty"
  cSQL = cSQL & " Desc"

  vaArray.ReDim 0, 1, 0, 2
  Set dbData = objData.Sql(GetDSN, cSQL)
  vaArray(0, 0) = "REKAPAN ORDERAN PROMO TGL INPUT " & Format(dTgl(0).Value, "dd/MM/yyyy") & " sd " & Format(dTgl(1).Value, "dd/MM/yyyy")
  vaArray(1, 0) = aCfg(objData, msNamaPerusahaan) & GetLastDBUpdate(objData)
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.EOF
      FrmPB.RunPB
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!barcode, "")
      vaArray(n, 1) = GetNull(dbData!qty, "")
      vaArray(n, 2) = GetNull(dbData!Total, "")
      dbData.MoveNext
    Loop
    FrmPB.EndPB
    a.RecordSource = vaArray
    a.ExportToExcel
  Else
    MsgBox "Sorry No Data To Display", vbExclamation
  End If
End Sub



