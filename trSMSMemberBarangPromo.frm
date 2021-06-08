VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trSMSMemberBarangPromo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SMS MEMBER BELUM AMBIL BARANG PROMO"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   6060
   Begin BiSAButtonProject.BiSAButton BiSAButton4 
      Height          =   630
      Left            =   4515
      TabIndex        =   9
      Top             =   2625
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   1111
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
   Begin BiSAButtonProject.BiSAButton BiSAButton3 
      Height          =   585
      Left            =   3585
      TabIndex        =   8
      Top             =   1305
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1032
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
   Begin BiSAButtonProject.BiSAButton BiSAButton1 
      Height          =   345
      Left            =   495
      TabIndex        =   5
      Top             =   2805
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   609
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
   Begin BiSADateProject.BiSADate dTgl 
      Height          =   360
      Left            =   435
      TabIndex        =   0
      Top             =   1200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   635
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
   Begin BiSAButtonProject.BiSAButton cmdOK 
      Height          =   600
      Left            =   450
      TabIndex        =   1
      Top             =   1605
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   1058
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
   Begin BiSAButtonProject.BiSAButton BiSAButton2 
      Height          =   345
      Left            =   525
      TabIndex        =   7
      Top             =   3630
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   609
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
   Begin VB.Label Label4 
      Caption         =   "Rekap member yg TIDAK DP TAPI Dapat barang"
      Height          =   240
      Left            =   525
      TabIndex        =   6
      Top             =   3315
      Width           =   3825
   End
   Begin VB.Label Label3 
      Caption         =   "Rekap member yg sudah DP dan Dapat barang"
      Height          =   240
      Left            =   525
      TabIndex        =   4
      Top             =   2535
      Width           =   3540
   End
   Begin VB.Label Label2 
      Caption         =   "Tgl Transaksi Lebih Dari"
      Height          =   345
      Left            =   435
      TabIndex        =   3
      Top             =   885
      Width           =   1830
   End
   Begin VB.Label Label1 
      Caption         =   "Tampilkan yang belum mengambil barang PROMO : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   450
      TabIndex        =   2
      Top             =   285
      Width           =   3000
   End
End
Attribute VB_Name = "trSMSMemberBarangPromo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nPos As SisPos
Dim lEdit As Boolean
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB
Dim objMenu As New CodeSuiteLibrary.Menu

Private Sub BiSAButton1_Click()
Dim cSQL As String
Dim n As Single
Dim db As New ADODB.Recordset

vaArray.ReDim 0, -1, 0, 6

cSQL = "select t.kodeanggota,a.telp,a.nama,sum(debet-kredit) as total from membertopup t"
cSQL = cSQL & " left join anggota a on a.kodeanggota =  t.kodeanggota"
cSQL = cSQL & " Where t.Tgl >= '2013-01-01'"
cSQL = cSQL & " GROUP BY t.kodeanggota"
cSQL = cSQL & " ORDER BY total desc"

Set db = objData.Sql(GetDSN, cSQL)
If Not db.EOF Then
  FrmPB.InitPB db.RecordCount
  Do While Not db.EOF
    FrmPB.RunPB
    vaArray.InsertRows vaArray.UpperBound(1) + 1
    n = vaArray.UpperBound(1)
    vaArray(n, 0) = "'" & GetNull(db!telp)
    vaArray(n, 1) = GetNull(db!nama)
    vaArray(n, 2) = GetNull(db!Total)
    vaArray(n, 3) = GetDapatBarangpromo(objData, GetNull(db!kodeanggota))
    vaArray(n, 4) = GetKeaktifanBelanja3BulanTerakhir(objData, GetNull(db!kodeanggota))
    vaArray(n, 5) = GetOrderanPromoMember(objData, GetNull(db!kodeanggota))
    vaArray(n, 6) = vaArray(n, 2) - vaArray(n, 3)
    db.MoveNext
  Loop
  FrmPB.EndPB
End If

'vaArray.QuickSort 0, vaArray.UpperBound(1), 6, XORDER_ASCEND, XTYPE_DOUBLE

Dim a As New exportExcel
a.RecordSource = vaArray
a.ExportToExcel

End Sub

Private Function GetDapatBarangpromo(ByVal obj As CodeSuiteLibrary.Data, ByVal kodeanggota As String) As Double
Dim db As New ADODB.Recordset
Dim cSQL As String

cSQL = "select a.nama,sum(p.qty*s.hargajual) as total from promo p"
cSQL = cSQL & " left join stock s on s.barcode = p.barcode"
cSQL = cSQL & " left join anggota a on a.kodeanggota = p.kodeanggota"
cSQL = cSQL & " Where p.Tgl >= '2015-12-17' and a.kodeanggota = '" & kodeanggota & "' group by p.kodeanggota"

  Set db = obj.Sql(GetDSN, cSQL)
  If Not db.EOF Then
    GetDapatBarangpromo = GetNull(db!Total)
  End If
  
End Function

Private Function GetKeaktifanBelanja3BulanTerakhir(ByVal obj As CodeSuiteLibrary.Data, ByVal kodeanggota As String) As Double
Dim db As New ADODB.Recordset
Dim cSQL As String

cSQL = "select sum(qty*harga) as total from penjualan p"
cSQL = cSQL & " left join totpenjualan t on t.nomorpenjualan = p.nomorpenjualan"
cSQL = cSQL & " where t.kodeanggota = '" & kodeanggota & "' and p.discount >= 30 and t.tgl >= DATE_SUB(NOW(),INTERVAL 3 MONTH)"

Set db = obj.Sql(GetDSN, cSQL)
If Not db.EOF Then
  GetKeaktifanBelanja3BulanTerakhir = GetNull(db!Total)
End If
End Function

Private Function GetDPMemberPromo(ByVal obj As CodeSuiteLibrary.Data, ByVal kodeanggota As String) As Double
Dim db As New ADODB.Recordset
Dim cSQL As String

cSQL = " select t.kodeanggota,a.nama,sum(debet-kredit) as total from membertopup t"
cSQL = cSQL & " left join anggota a on a.kodeanggota =  t.kodeanggota"
cSQL = cSQL & " Where t.Tgl >= '2013-01-01' And t.lStatus = 2 and t.kodeanggota = '" & kodeanggota & "'"
cSQL = cSQL & " GROUP BY t.kodeanggota"
cSQL = cSQL & " ORDER BY total desc"

Set db = obj.Sql(GetDSN, cSQL)
If Not db.EOF Then
  GetDPMemberPromo = GetNull(db!Total)
End If
End Function

Private Function GetOrderanPromoMember(ByVal obj As CodeSuiteLibrary.Data, ByVal kodeanggota As String) As Double
Dim db As New ADODB.Recordset
Dim cSQL As String

cSQL = ""
cSQL = cSQL & " select o.kodeanggota,a.nama,sum(o.qty*s.hargajual) as jumlah from orderpromo_karangasem o"
cSQL = cSQL & " LEFT JOIN stock s on s.kodestock  = o.kodestock"
cSQL = cSQL & " LEFT JOIN anggota a on a.kodeanggota = o.kodeanggota"
cSQL = cSQL & " WHERE o.kodeanggota = '" & kodeanggota & "'"
cSQL = cSQL & " GROUP BY o.kodeanggota"

Set db = obj.Sql(GetDSN, cSQL)
If Not db.EOF Then
  GetOrderanPromoMember = GetNull(db!jumlah)
End If

End Function

Private Sub BiSAButton2_Click()
Dim db As New ADODB.Recordset
Dim cSQL As String
Dim n As Single

cSQL = ""
cSQL = cSQL & " select p.kodeanggota,a.telp,a.nama,sum(p.qty*s.hargajual) as total from promo p"
cSQL = cSQL & " left join stock s on s.barcode = p.barcode"
cSQL = cSQL & " left join anggota a on a.kodeanggota = p.kodeanggota"
cSQL = cSQL & " Where p.Tgl >= '2013-01-01'"
cSQL = cSQL & " GROUP BY p.kodeanggota"
cSQL = cSQL & " ORDER BY total desc"

vaArray.ReDim 0, -1, 0, 4

Set db = objData.Sql(GetDSN, cSQL)
If Not db.EOF Then
  FrmPB.InitPB db.RecordCount
  Do While Not db.EOF
    FrmPB.RunPB
    vaArray.InsertRows vaArray.UpperBound(1) + 1
    n = vaArray.UpperBound(1)
    vaArray(n, 0) = "'" & GetNull(db!telp)
    vaArray(n, 1) = GetNull(db!nama)
    vaArray(n, 2) = GetDapatBarangpromo(objData, GetNull(db!kodeanggota))
    vaArray(n, 3) = GetKeaktifanBelanja3BulanTerakhir(objData, GetNull(db!kodeanggota))
    vaArray(n, 4) = GetOrderanPromoMember(objData, GetNull(db!kodeanggota))
    If GetDPMemberPromo(objData, GetNull(db!kodeanggota)) > 0 Then
      vaArray.DeleteRows n
    End If
    db.MoveNext
  Loop
  FrmPB.EndPB
End If

'vaArray.QuickSort 0, vaArray.UpperBound(1), 2, XORDER_DESCEND, XTYPE_DOUBLE

Dim a As New exportExcel
a.RecordSource = vaArray
a.ExportToExcel
    
End Sub

Private Sub BiSAButton3_Click()
Dim db As New ADODB.Recordset
Dim a As New exportExcel
Dim n As Single

vaArray.ReDim 0, -1, 0, 2
Set db = objData.Browse(GetDSN, "promo p", , "p.tgl", sisGTEqual, "2013-01-01", , , Array("LEFT JOIN anggota a ON a.kodeanggota = p.kodeanggota"))
If Not db.EOF Then
  FrmPB.InitPB db.RecordCount
  Do While Not db.EOF
    FrmPB.RunPB

    If GetCariDiOrderan(objData, GetNull(db!barcode), GetNull(db!kodeanggota)) = False Then
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
'      MsgBox "member " & GetNull(db!kodeanggota) & " tidak order barang " & GetNull(db!barcode)
      vaArray(n, 0) = GetNull(db!kodeanggota)
      vaArray(n, 1) = GetNull(db!nama)
      vaArray(n, 2) = GetNull(db!barcode)
    End If
    db.MoveNext
  Loop
  FrmPB.EndPB
     
  a.RecordSource = vaArray
  a.ExportToExcel
  
  MsgBox "selesai"
End If

End Sub

Private Function GetCariDiOrderan(ByVal obj As CodeSuiteLibrary.Data, ByVal kodebarang As String, ByVal kodeanggota As String) As Boolean
Dim db As New ADODB.Recordset

GetCariDiOrderan = False
Set db = objData.Sql(GetDSN, "select * from orderpromo_karangasem where barcode = '" & kodebarang & "' and kodeanggota = '" & kodeanggota & "'")
If Not db.EOF Then
  GetCariDiOrderan = True
End If
End Function

Private Sub BiSAButton4_Click()
Dim db As New ADODB.Recordset
Dim n As Single
Dim a As New exportExcel

vaArray.ReDim 0, -1, 0, 6
Set db = objData.Sql(GetDSN, "select distinct(p.kodeanggota) as kodeanggota,a.telp,a.nama from promo p left join anggota a on a.kodeanggota = p.kodeanggota where p.tgl >='2015-12-17'")
If Not db.EOF Then
  FrmPB.InitPB db.RecordCount
  Do While Not db.EOF
    FrmPB.RunPB
    vaArray.InsertRows vaArray.UpperBound(1) + 1
    n = vaArray.UpperBound(1)
    vaArray(n, 0) = "'" & GetNull(db!telp)
    vaArray(n, 1) = "'" & GetNull(db!kodeanggota)
    vaArray(n, 2) = GetNull(db!nama)
    vaArray(n, 3) = GetDapatBarangpromo(objData, GetNull(db!kodeanggota))
    vaArray(n, 4) = GetSaldoTopUp(objData, GetNull(db!kodeanggota))
    vaArray(n, 5) = vaArray(n, 4) - vaArray(n, 3)
    vaArray(n, 6) = "Bonjour " & vaArray(n, 2) & " barang promo nya sudah bisa diambil senilai " & Format(vaArray(n, 3), "###,###,##") & IIf(vaArray(n, 4) > 0, " DP " & Format(vaArray(n, 4), "###,###,##"), "") & " Saldo " & Format(vaArray(n, 5), "###,###,##") & " Barang yg dapat: " & GetItemBarangPromo(objData, GetNull(db!kodeanggota))
    If vaArray(n, 3) = 0 Then
      vaArray.DeleteRows n
    Else
      'cek barang ini sudah diambil belum
      If IsSudahAmbilPromo(objData, GetNull(db!kodeanggota)) = True Then
        vaArray.DeleteRows n
      End If
    End If
    db.MoveNext
  Loop
  FrmPB.EndPB
End If

vaArray.QuickSort 0, vaArray.UpperBound(1), 5, XORDER_ASCEND, XTYPE_DOUBLE
a.RecordSource = vaArray
a.ExportToExcel

End Sub

Private Function IsSudahAmbilPromo(ByVal obj As CodeSuiteLibrary.Data, ByVal cKodeAnggota As String) As Boolean
Dim db As New ADODB.Recordset

  IsSudahAmbilPromo = False
  Set db = obj.Browse(GetDSN, "totpenjualan", , "jenis", sisAssign, "P", " AND kodeanggota = '" & cKodeAnggota & "' AND flaglunas = 1 and tgl >= '2014-1-20'")
  If Not db.EOF Then
    IsSudahAmbilPromo = True
  End If
  
End Function

Private Function GetItemBarangPromo(ByVal obj As CodeSuiteLibrary.Data, ByVal cKAnggota As String) As String
Dim db As New ADODB.Recordset
Dim cSQL2 As String


Set db = obj.Browse(GetDSN, "promo", "barcode,qty", "kodeanggota", sisAssign, cKAnggota, " and tgl >= '2013-11-01'")


'  cSQL2 = "select s.barcode,p.qty from penjualan p"
'  cSQL2 = cSQL2 & " left join stock s on s.kodestock = p.kodestock"
'  cSQL2 = cSQL2 & " Where p.nomorpenjualan = '" & cNo & "'"
  
'  Set db = obj.Sql(GetDSN, cSQL2)
  If Not db.EOF Then
    Do While Not db.EOF
      GetItemBarangPromo = GetItemBarangPromo & GetNull(db!barcode) & IIf(GetNull(db!qty) > 1, "(" & GetNull(db!qty) & ")", "") & " "
      db.MoveNext
    Loop
  End If
  
End Function

Private Function GetSaldoTopUp(ByVal obj As CodeSuiteLibrary.Data, ByVal cKodeA As String) As Double
Dim db As New ADODB.Recordset
GetSaldoTopUp = 0
Set db = obj.Browse(GetDSN, "membertopup", "sum(debet-kredit) as topup", "tgl", sisGTEqual, "2014-1-20", " and lstatus = '2' and kodeanggota ='" & cKodeA & "'")
If Not db.EOF Then
  GetSaldoTopUp = GetNull(db!topup)
End If
End Function

Private Sub cmdOK_Click()
Dim cSQL As String
Dim n As Single
Dim a As New exportExcel

cSQL = "select t.nomorpenjualan,t.tgl,a.nama,a.telp,t.total,a.kodeanggota,t.jenis FROM totpenjualan t"
cSQL = cSQL & " LEFT JOIN anggota a on a.kodeanggota = t.kodeanggota"
cSQL = cSQL & " Where t.tgl >= '" & Format(dTgl.Value, "yyyy-MM-dd") & "' And flaglunas <> 1 AND jenis = 'P' order by t.tgl "

  vaArray.ReDim 0, -1, 0, 5
  
'  vaArray(0, 0) = aCfg(objData, msNamaPerusahaan) & " " & aCfg(objData, msAlamatPerusahaan)
'
'  vaArray(1, 0) = "NAMA"
'  vaArray(1, 1) = "TELP"
'  vaArray(1, 2) = "TOTAL"
'  vaArray(1, 3) = "DP"
'  vaArray(1, 4) = "SELISIH"
'  vaArray(1, 5) = "ITEM"

  Set dbData = objData.Sql(GetDSN, cSQL)
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.EOF
      FrmPB.RunPB
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = "'" & GetNull(dbData!telp)
      vaArray(n, 1) = UCase(GetNull(dbData!nama))
      vaArray(n, 2) = GetNull(dbData!Total)
      vaArray(n, 3) = GetSaldoTopUpMember2(objData, GetNull(dbData!kodeanggota))
      vaArray(n, 4) = vaArray(n, 2) - vaArray(n, 3)
      vaArray(n, 5) = GetItemBarang(objData, GetNull(dbData!nomorpenjualan))
      If Not isPromo(objData, GetNull(dbData!nomorpenjualan)) Then
        vaArray.DeleteRows n
      End If
      dbData.MoveNext
    Loop
    FrmPB.EndPB
    vaArray.QuickSort 0, vaArray.UpperBound(1), 4, XORDER_DESCEND, XTYPE_DOUBLE
    a.RecordSource = vaArray
    a.ExportToExcel
  End If
End Sub

Private Function isPromo(ByVal obj As CodeSuiteLibrary.Data, ByVal cFkt As String) As Boolean
Dim db As New ADODB.Recordset

  isPromo = True
  Set db = obj.Browse(GetDSN, "penjualan", , "nomorpenjualan", sisAssign, cFkt)
  If Not db.EOF Then
    Do While Not db.EOF
      If GetNull(db!Discount) <> 0 Then
        isPromo = False
        Exit Function
      End If
      db.MoveNext
    Loop
  End If
End Function

Private Function GetItemBarang(ByVal obj As CodeSuiteLibrary.Data, ByVal cNo As String) As String
Dim db As New ADODB.Recordset
Dim cSQL2 As String

  cSQL2 = "select s.barcode,p.qty from penjualan p"
  cSQL2 = cSQL2 & " left join stock s on s.kodestock = p.kodestock"
  cSQL2 = cSQL2 & " Where p.nomorpenjualan = '" & cNo & "'"
  
  Set db = obj.Sql(GetDSN, cSQL2)
  If Not db.EOF Then
    Do While Not db.EOF
      GetItemBarang = GetItemBarang & GetNull(db!barcode) & IIf(GetNull(db!qty) > 1, "(" & GetNull(db!qty) & ")", "") & " "
      db.MoveNext
    Loop
  End If
  
End Function

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  SetIcon Me.hWnd
  TabIndex dTgl, n
  TabIndex cmdOK, n
  dTgl.Value = Now
End Sub
