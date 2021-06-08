VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptSaldoStockReguler 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN SALDO STOCK REGULER"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   5520
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1680
      Left            =   0
      Top             =   0
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   2963
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
      Begin VB.CheckBox Check1 
         Caption         =   "Semua Barang"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1305
         TabIndex        =   5
         Top             =   945
         Width           =   1860
      End
      Begin VB.CheckBox chkPembelian 
         Caption         =   "Status Pembelian"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1320
         TabIndex        =   4
         Top             =   720
         Width           =   1860
      End
      Begin BiSADateProject.BiSADate dDate 
         Height          =   330
         Index           =   0
         Left            =   165
         TabIndex        =   0
         Top             =   180
         Visible         =   0   'False
         Width           =   2430
         _ExtentX        =   4286
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
         Caption         =   "Mutasi"
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
      Begin BiSADateProject.BiSADate dDate 
         Height          =   330
         Index           =   1
         Left            =   2670
         TabIndex        =   1
         Top             =   180
         Visible         =   0   'False
         Width           =   1740
         _ExtentX        =   3069
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
         Caption         =   "s.d"
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
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   0
      Top             =   1665
      Width           =   5490
      _ExtentX        =   9684
      _ExtentY        =   1111
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
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   4275
         TabIndex        =   2
         Top             =   105
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   767
         Caption         =   "     &Exit"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         BackColor       =   -2147483633
         Picture         =   "rptSaldoStockReguler.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   3825
         TabIndex        =   3
         Top             =   105
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   767
         Caption         =   ""
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         BackColor       =   -2147483633
         Picture         =   "rptSaldoStockReguler.frx":00A6
      End
   End
End
Attribute VB_Name = "rptSaldoStockReguler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub chkPembelian_Click()
'  MsgBox chkPembelian.Value
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
Dim cSQL As String
Dim n As Single
Dim a As New exportExcel
Dim cKet As String
Dim cStatus As String
Dim dTg As Date
Dim nQt As Single

  cSQL = cSQL & " select s.kodestock,s.barcode,s.nama,sum(k.debet-k.kredit) as qty,s.hargajual as hK,k.`status`,k.tgl,k.keterangan from kartustock k"
  cSQL = cSQL & " left join stock s on s.kodestock = k.kodestock"
  If Check1.Value <> 1 Then
    cSQL = cSQL & " Where s.diskonpenjualan = 30"
  End If
  cSQL = cSQL & " GROUP BY s.kodestock"
  cSQL = cSQL & " Having Sum(k.debet - k.kredit) <> 0"
  
  vaArray.ReDim 0, -1, 0, 8
  Set dbData = objData.Sql(GetDSN, cSQL)
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.EOF
      FrmPB.RunPB
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = "'" & GetNull(dbData!KodeStock)
      vaArray(n, 1) = GetNull(dbData!barcode)
      vaArray(n, 2) = GetNull(dbData!nama)
      vaArray(n, 3) = GetNull(dbData!qty)
      vaArray(n, 4) = Format(GetNull(dbData!hk), "##,###,###,##0.00")
      
      GetStatusKartuStock GetNull(dbData!barcode), cKet, cStatus, dTg, nQt
      
      vaArray(n, 5) = dTg
      vaArray(n, 6) = cStatus
      vaArray(n, 7) = nQt
      vaArray(n, 8) = cKet
      
'      MsgBox chkPembelian.Value

      If chkPembelian.Value = 1 Then
        If vaArray(n, 6) <> 10 Then
          vaArray.DeleteRows n
        End If
      End If
      dbData.MoveNext
    Loop
    FrmPB.EndPB
  End If
  
  vaArray.QuickSort 0, vaArray.UpperBound(1), 5, XORDER_DESCEND, XTYPE_DATE
  a.RecordSource = vaArray
  a.ExportToExcel
End Sub

Private Sub GetStatusKartuStock(ByVal barcode As String, ByRef cKet As String, ByRef cStatus As String, ByRef dTg As Date, ByRef nQt As Single)
Dim db As New ADODB.Recordset
Dim cSQL As String


  cKet = ""
  cSQL = cSQL & " select s.barcode,k.`status`,k.tgl,k.keterangan,k.qty from kartustock k"
  cSQL = cSQL & " left join stock s on s.kodestock = k.kodestock"
  cSQL = cSQL & " Where s.Barcode = '" & barcode & "'"
  cSQL = cSQL & " ORDER BY k.id DESC"
  cSQL = cSQL & " LIMIT 0, 1"
  
  Set db = objData.Sql(GetDSN, cSQL)
  If Not db.EOF Then
    cKet = GetNull(db!keterangan)
    cStatus = GetNull(db!Status)
    dTg = GetNull(db!Tgl)
    nQt = GetNull(db!qty)
  End If
End Sub

Private Sub Form_Load()
  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  chkPembelian.Value = 1
End Sub
