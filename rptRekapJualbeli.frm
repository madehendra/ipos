VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptRekapJualbeli 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rekap Jual Beli"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   7320
   Begin BiSAButtonProject.BiSAButton BiSAButton1 
      Height          =   480
      Left            =   2070
      TabIndex        =   10
      Top             =   3990
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   847
      Caption         =   "Label1"
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
   Begin BiSAButtonProject.BiSAButton cmdRekapPembelian 
      Height          =   390
      Left            =   1320
      TabIndex        =   4
      Top             =   1140
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   688
      Caption         =   "Rekap Pembelian"
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
   Begin BiSADateProject.BiSADate dTglBeli 
      Height          =   330
      Index           =   0
      Left            =   1335
      TabIndex        =   0
      Top             =   330
      Width           =   1545
      _ExtentX        =   2725
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
   Begin BiSADateProject.BiSADate dTglBeli 
      Height          =   330
      Index           =   1
      Left            =   2940
      TabIndex        =   1
      Top             =   345
      Width           =   1545
      _ExtentX        =   2725
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
      Left            =   195
      TabIndex        =   2
      Top             =   735
      Width           =   3165
      _ExtentX        =   5583
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
      Button          =   -1  'True
      Caption         =   "Supplier"
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
      Left            =   3360
      TabIndex        =   3
      Top             =   735
      Width           =   3270
      _ExtentX        =   5768
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
   Begin BiSADateProject.BiSADate dTglJual 
      Height          =   330
      Index           =   1
      Left            =   2865
      TabIndex        =   5
      Top             =   2220
      Width           =   1545
      _ExtentX        =   2725
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
   Begin BiSADateProject.BiSADate dTglJual 
      Height          =   330
      Index           =   0
      Left            =   1290
      TabIndex        =   6
      Top             =   2220
      Width           =   1545
      _ExtentX        =   2725
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
      Left            =   165
      TabIndex        =   7
      Top             =   2610
      Width           =   3165
      _ExtentX        =   5583
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
      Button          =   -1  'True
      Caption         =   "Customer"
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
      Left            =   3330
      TabIndex        =   8
      Top             =   2610
      Width           =   3270
      _ExtentX        =   5768
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
   Begin BiSAButtonProject.BiSAButton cmdRekapPenjualan 
      Height          =   390
      Left            =   1290
      TabIndex        =   9
      Top             =   3015
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   688
      Caption         =   "Rekap Penjualan"
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
   Begin BiSAButtonProject.BiSAButton BiSAButton2 
      Height          =   390
      Left            =   3435
      TabIndex        =   11
      Top             =   1140
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   688
      Caption         =   "Detail Pembelian"
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
   Begin BiSADateProject.BiSADate dTglBeli 
      Height          =   330
      Index           =   2
      Left            =   4515
      TabIndex        =   12
      Top             =   345
      Width           =   1545
      _ExtentX        =   2725
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
Attribute VB_Name = "rptRekapJualbeli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub cmdPreview_Click()
  GetRpt
End Sub

Private Sub BiSAButton1_Click()
Dim cSQL As String
Dim db As New ADODB.Recordset
Dim n As Single


  vaArray.ReDim 0, -1, 0, 4
  
  cSQL = cSQL & " select DISTINCT(p.kodestock) as kodestock,s.barcode,s.nama,p.qty,p.tgl,p.nomorpembelian from pembelian p"
  cSQL = cSQL & " LEFT JOIN stock s on s.kodestock = p.kodestock"
  cSQL = cSQL & " Where p.Tgl >= '2012-07-01'"
  
  Set db = objData.Sql(GetDSN, cSQL)
  If Not db.EOF Then
    FrmPB.InitPB db.RecordCount
    Do While Not db.EOF
      FrmPB.RunPB
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(db!barcode)
      vaArray(n, 1) = GetNull(db!nama)
      vaArray(n, 2) = GetSaldoStock(objData, "", GetNull(db!KodeStock))
      vaArray(n, 3) = "Beli Trakhir " & Format(GetNull(db!Tgl), "dd-MM-yyyy")
      vaArray(n, 4) = GetTransaksiTerakhir(GetNull(db!KodeStock))
      If vaArray(n, 2) = 0 Then vaArray.DeleteRows n
      db.MoveNext
    Loop
    FrmPB.EndPB
    Dim a As New exportExcel
    a.RecordSource = vaArray
    a.ExportToExcel
  End If
End Sub

Private Function GetTransaksiTerakhir(ByVal KodeStock As String) As String
Dim cSQL As String
Dim db As New ADODB.Recordset

  GetTransaksiTerakhir = ""
  
  cSQL = cSQL & " select keterangan,tgl from kartustock"
  cSQL = cSQL & " Where KodeStock = '" & KodeStock & "'"
  cSQL = cSQL & " ORDER BY id desc LIMIT 0,1"
  
  Set db = objData.Sql(GetDSN, cSQL)
  If Not db.EOF Then
    GetTransaksiTerakhir = GetNull(db!keterangan) & " Tgl " & Format(GetNull(db!Tgl), "dd-MM-yyyy")
  End If
End Function


Private Sub BiSAButton2_Click()
Dim cSQL As String
Dim db As New ADODB.Recordset
Dim n As Integer
Dim a As New exportExcel


cSQL = cSQL & " select  s.barcode,s.kodestock,s.hargajual from pembelian p"
cSQL = cSQL & " left join totpembelian t on t.nomorpembelian = p.nomorpembelian"
cSQL = cSQL & " left join stock s on s.kodestock = p.kodestock"
cSQL = cSQL & " where p.tgl >='" & Format(dTglBeli(0).Value, "yyyy-MM-dd") & "' and p.tgl <= '" & Format(dTglBeli(1).Value, "yyyy-MM-dd") & "'"
cSQL = cSQL & " GROUP BY s.barcode"
cSQL = cSQL & " ORDER BY s.barcode ASC"

vaArray.ReDim 0, -1, 0, 2
Set db = objData.Sql(GetDSN, cSQL)
If Not db.EOF Then
  
  FrmPB.InitPB db.RecordCount
  Do While Not db.EOF
    FrmPB.RunPB
    vaArray.InsertRows vaArray.UpperBound(1) + 1
    n = vaArray.UpperBound(1)
    vaArray(n, 0) = GetNull(db!barcode)
    vaArray(n, 1) = GetSaldoStock(objData, "", GetNull(db!KodeStock), Format(dTglBeli(2).Value, ""))
    vaArray(n, 2) = GetNull(db!hargajual)
    If vaArray(n, 1) = 0 Then vaArray.DeleteRows n
    db.MoveNext
  Loop
  FrmPB.EndPB
  
  a.RecordSource = vaArray
  a.ExportToExcel

End If
End Sub

Private Sub cCustomer_ButtonClick()
Dim vaTmp As New XArrayDB

  Set dbData = objData.Browse(GetDSN, "anggota a", "a.kodeanggota,a.nama,a.kodedep,a.alamat,a.telp,a.dd,d.keterangan", "a.kodeanggota", sisContent, cCustomer.Text, , "a.kodeanggota,a.nama", Array("Left join dep d on d.kodedep = a.kodedep"))
  If Not dbData.EOF Then
    cCustomer.Text = cCustomer.Browse(dbData, Array("KODE", "NAMA", "DEP", "ALAMAT"), , Array(10, 20, 6, 10))
    cCustomer.Text = GetNull(dbData!kodeanggota)
    cNamaCustomer.Text = GetNull(dbData!nama, "")
  End If
End Sub

Private Sub cmdRekapPembelian_Click()
Dim cSQL As String
Dim n As Single

cSQL = cSQL & " select tgl,total from totpembelian where tgl >= '" & Format(dTglBeli(0).Value, "yyyy-MM-dd") & "' and tgl <= '" & Format(dTglBeli(1).Value, "yyyy-MM-dd") & "' and kodesupplier = '" & cSupplier.Text & "'"
cSQL = cSQL & " ORDER BY tgl"

  vaArray.ReDim 0, 2, 0, 1
  Set dbData = objData.Sql(GetDSN, cSQL)
  If Not dbData.EOF Then
    vaArray(0, 0) = "Pembelian dari " & cSupplier.Text & " " & cNamaSupplier.Text
    vaArray(1, 0) = "Dari Tgl " & Format(dTglBeli(0).Value, "dd.MM.yyyy") & " sd " & Format(dTglBeli(1).Value, "dd.MM.yyyy")
    vaArray(2, 0) = "Tgl"
    vaArray(2, 1) = "Jumlah"
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = Format(GetNull(dbData!Tgl), "dd.MM.yyyy")
      vaArray(n, 1) = GetNull(dbData!Total)
      dbData.MoveNext
    Loop
    
    Dim a As New exportExcel
    a.RecordSource = vaArray
    a.ExportToExcel
  End If
End Sub

Private Sub cmdRekapPenjualan_Click()
Dim cSQL As String
Dim n As Single

cSQL = ""
cSQL = cSQL & " select tgl,total from totpenjualan where tgl >= '" & Format(dTglJual(0).Value, "yyyy-MM-dd") & "' and tgl <= '" & Format(dTglJual(1).Value, "yyyy-MM-dd") & "' and kodeanggota = '" & cCustomer.Text & "'"
cSQL = cSQL & " ORDER BY tgl"

  vaArray.ReDim 0, 2, 0, 1
  Set dbData = objData.Sql(GetDSN, cSQL)
  If Not dbData.EOF Then
    vaArray(0, 0) = "Penjualan Ke " & cCustomer.Text & " " & cNamaCustomer.Text
    vaArray(1, 0) = "Dari Tgl " & Format(dTglJual(0).Value, "dd.MM.yyyy") & " sd " & Format(dTglJual(1).Value, "dd.MM.yyyy")
    vaArray(2, 0) = "Tgl"
    vaArray(2, 1) = "Jumlah"
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = Format(GetNull(dbData!Tgl), "dd.MM.yyyy")
      vaArray(n, 1) = GetNull(dbData!Total)
      dbData.MoveNext
    Loop
    
    Dim a As New exportExcel
    a.RecordSource = vaArray
    a.ExportToExcel

  End If
End Sub

Private Sub cNamaCustomer_ButtonClick()
Dim vaTmp As New XArrayDB

  Set dbData = objData.Browse(GetDSN, "anggota a", "a.nama,a.kodeanggota,a.kodedep,a.alamat,a.telp,d.keterangan", "a.nama", sisContent, cNamaCustomer.Text, , "a.kodeanggota,a.nama", Array("Left join dep d on d.kodedep = a.kodedep"))
  If Not dbData.EOF Then
    cNamaCustomer.Text = cNamaCustomer.Browse(dbData, Array("Nama", "Kode", "Dep", "Alamat"), , Array(6, 15, 6, 15))
    cCustomer.Text = GetNull(dbData!kodeanggota)
    cNamaCustomer.Text = GetNull(dbData!nama)
  End If
End Sub

Private Sub cNamaSupplier_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "supplier", "kodesupplier,nama,alamat,kota", "nama", sisContent, cNamaSupplier.Text, , "kodesupplier,nama")
  If Not dbData.EOF Then
    cNamaSupplier.Text = cNamaSupplier.Browse(dbData)
    cSupplier.Text = GetNull(dbData!kodesupplier)
    cNamaSupplier.Text = GetNull(dbData!nama, "")
  End If
End Sub

Private Sub cSupplier_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "supplier", "kodesupplier,nama,alamat,kota", "kodesupplier", sisContent, cSupplier.Text, , "kodesupplier,nama")
  If Not dbData.EOF Then
    cSupplier.Text = cSupplier.Browse(dbData)
    cSupplier.Text = GetNull(dbData!kodesupplier)
    cNamaSupplier.Text = GetNull(dbData!nama, "")
  End If
End Sub

Private Sub Form_Load()
Dim n As Single
  
  CenterForm Me
  SetIcon Me.hWnd
'  initvalue
  dTglBeli(0).Value = BOM(Date)
  dTglBeli(1).Value = EOM(Date)
  dTglBeli(2).Value = EOM(Date)
  cSupplier.Default
  cNamaSupplier.Default
  dTglJual(0).Value = BOM(Date)
  dTglJual(1).Value = EOM(Date)
  cCustomer.Default
  cNamaCustomer.Default
  
  TabIndex dTglBeli(0), n
  TabIndex dTglBeli(1), n
  TabIndex cSupplier, n
  TabIndex cNamaSupplier, n
  TabIndex cmdRekapPembelian, n
  TabIndex dTglJual(0), n
  TabIndex dTglJual(1), n
  TabIndex cCustomer, n
  TabIndex cNamaCustomer, n
  TabIndex cmdRekapPenjualan, n
  
End Sub

Private Sub GetRpt()
Dim n As Integer
  vaArray.ReDim 0, -1, 0, 3
  
End Sub

