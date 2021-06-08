VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trExportPembelianPenjualan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EXPORT PEMBELIAN PENJUALAN"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   7095
   Begin BiSADateProject.BiSADate dTgl 
      Height          =   330
      Index           =   0
      Left            =   285
      TabIndex        =   2
      Top             =   195
      Width           =   3000
      _ExtentX        =   5292
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
      ForeColor       =   -2147483640
      Caption         =   "Dari Tgl"
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
   Begin BiSATextBoxProject.BiSABrowse cNamaSupplier 
      Height          =   330
      Left            =   3570
      TabIndex        =   3
      Top             =   945
      Width           =   3045
      _ExtentX        =   5371
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
   Begin BiSATextBoxProject.BiSABrowse cSupplier 
      Height          =   330
      Left            =   285
      TabIndex        =   4
      Top             =   945
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   582
      Text            =   "12345678"
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
      Caption         =   "Supplier"
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
   Begin BiSATextBoxProject.BiSABrowse cNamaCustomer 
      Height          =   330
      Left            =   3615
      TabIndex        =   5
      Top             =   2340
      Width           =   3045
      _ExtentX        =   5371
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
   Begin BiSATextBoxProject.BiSABrowse cCustomer 
      Height          =   330
      Left            =   315
      TabIndex        =   6
      Top             =   2340
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   582
      Text            =   "12345678"
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
      Caption         =   "Member"
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
      Left            =   3315
      TabIndex        =   7
      Top             =   195
      Width           =   2145
      _ExtentX        =   3784
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
      ForeColor       =   -2147483640
      Caption         =   "sd."
      CaptionWidth    =   500
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
   Begin BiSAButtonProject.BiSAButton cmdExportPembelian 
      Height          =   435
      Left            =   5595
      TabIndex        =   8
      Top             =   1470
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   767
      Caption         =   "    &Export"
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
      Picture         =   "trExportPembelianPenjualan.frx":0000
   End
   Begin BiSAButtonProject.BiSAButton cmdExportPenjualan 
      Height          =   435
      Left            =   5625
      TabIndex        =   9
      Top             =   2865
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   767
      Caption         =   "    &Export"
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
      Picture         =   "trExportPembelianPenjualan.frx":0286
   End
   Begin VB.Label Label2 
      Caption         =   "EXPORT DETAIL PENJUALAN KE :"
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
      Left            =   375
      TabIndex        =   1
      Top             =   2055
      Width           =   3045
   End
   Begin VB.Label Label1 
      Caption         =   "EXPORT DETAIL PEMBELIAN DARI :"
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
      Left            =   330
      TabIndex        =   0
      Top             =   660
      Width           =   3045
   End
End
Attribute VB_Name = "trExportPembelianPenjualan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lEdit As Boolean
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB
Dim n As Single

Private Sub cCustomer_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "anggota", "kodeanggota,nama,alamat", "kodeanggota", sisContent, cCustomer.Text)
  If Not dbData.EOF Then
    cCustomer.Text = cCustomer.Browse(dbData)
    cNamaCustomer.Text = GetNull(dbData!nama)
  End If
End Sub


Private Sub cmdExportPembelian_Click()
Dim cSQL As String
Dim db As New ADODB.Recordset
Dim a As New exportExcel

  cSQL = "select t.kodesupplier,s.barcode,s.nama,s.hargabeli,s.diskonpenjualan,SUM(p.qty) as qty from pembelian p"
  cSQL = cSQL & " left join totpembelian t on t.nomorpembelian = p.nomorpembelian"
  cSQL = cSQL & " LEFT JOIN stock s on s.kodestock = p.kodestock"
  cSQL = cSQL & " where p.tgl >= '" & Format(dTgl(0).Value, "yyyy-MM-dd") & "' and p.tgl <= '" & Format(dTgl(1).Value, "yyyy-MM-dd") & "' and t.kodesupplier = '" & cSupplier.Text & "'GROUP BY p.kodestock"
  cSQL = cSQL & " ORDER BY s.barcode ASC"
  
  vaArray.ReDim 0, -1, 0, 6
  Set dbData = objData.Sql(GetDSN, cSQL)
  If Not dbData.EOF Then
    Do While Not dbData.EOF '
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!kodesupplier)
      vaArray(n, 1) = GetNull(dbData!barcode)
      vaArray(n, 2) = GetNull(dbData!nama)
      vaArray(n, 3) = GetNull(dbData!hargabeli)
      vaArray(n, 4) = GetNull(dbData!diskonpenjualan)
      vaArray(n, 5) = GetNull(dbData!qty)
      vaArray(n, 6) = vaArray(n, 5) * (vaArray(n, 3) - (vaArray(n, 3) * ((vaArray(n, 4) / 100))))
      dbData.MoveNext
    Loop
    a.RecordSource = vaArray
    a.ExportToExcel
  End If
End Sub

Private Sub cmdExportPenjualan_Click()
Dim cSQL As String
Dim db As New ADODB.Recordset
Dim a As New exportExcel

  cSQL = ""
  cSQL = cSQL & " select t.kodeanggota,s.barcode,s.nama,s.hargabeli,s.diskonpenjualan,SUM(p.qty) as qty from penjualan p"
  cSQL = cSQL & " left join totpenjualan t on t.nomorpenjualan = p.nomorpenjualan"
  cSQL = cSQL & " LEFT JOIN stock s on s.kodestock = p.kodestock"
  cSQL = cSQL & " where p.tgl >= '" & Format(dTgl(0).Value, "yyyy-MM-dd") & "' and p.tgl <= '" & Format(dTgl(1).Value, "yyyy-MM-dd") & "' and t.kodeanggota = '" & cCustomer.Text & "'GROUP BY p.kodestock"
  cSQL = cSQL & " ORDER BY s.barcode ASC"
  
  vaArray.ReDim 0, -1, 0, 6
  Set dbData = objData.Sql(GetDSN, cSQL)
  If Not dbData.EOF Then
    Do While Not dbData.EOF '
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!kodeanggota)
      vaArray(n, 1) = GetNull(dbData!barcode)
      vaArray(n, 2) = GetNull(dbData!nama)
      vaArray(n, 3) = GetNull(dbData!hargabeli)
      vaArray(n, 4) = GetNull(dbData!diskonpenjualan)
      vaArray(n, 5) = GetNull(dbData!qty)
      vaArray(n, 6) = vaArray(n, 5) * (vaArray(n, 3) - (vaArray(n, 3) * ((vaArray(n, 4) / 100))))
      dbData.MoveNext
    Loop
    a.RecordSource = vaArray
    a.ExportToExcel
  End If
End Sub

Private Sub cNamaCustomer_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "anggota", "nama,kodeanggota,alamat", "nama", sisContent, cNamaCustomer.Text)
  If Not dbData.EOF Then
    cNamaCustomer.Text = cNamaCustomer.Browse(dbData)
    cCustomer.Text = GetNull(dbData!kodeanggota)
  End If
End Sub

Private Sub cNamaSupplier_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "supplier", "nama,kodesupplier,alamat", "nama", sisContent, cNamaSupplier.Text)
  If Not dbData.EOF Then
    cNamaSupplier.Text = cNamaSupplier.Browse(dbData)
    cSupplier.Text = GetNull(dbData!kodesupplier)
  End If
End Sub

Private Sub cSupplier_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "supplier", "kodesupplier,nama,alamat", "kodesupplier", sisContent, cSupplier.Text)
  If Not dbData.EOF Then
    cSupplier.Text = cSupplier.Browse(dbData)
    cNamaSupplier.Text = GetNull(dbData!nama)
  End If
End Sub


Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  initvalue
  
  TabIndex dTgl(0), n
  TabIndex dTgl(1), n
  TabIndex cSupplier, n
  TabIndex cNamaSupplier, n
  TabIndex cmdExportPembelian, n
  TabIndex cCustomer, n
  TabIndex cNamaCustomer, n
  TabIndex cmdExportPenjualan, n

End Sub

Private Sub initvalue()
  dTgl(0).Value = BOM(Date)
  dTgl(1).Value = EOM(Date)
  cSupplier.Default
  cNamaSupplier.Default
  cCustomer.Default
  cNamaCustomer.Default
End Sub



