VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trSMSMemberYgAktif 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SMS MEMBER YG AKTIF"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   4875
   Begin BiSAButtonProject.BiSAButton BiSAButton5 
      Height          =   645
      Left            =   1965
      TabIndex        =   9
      Top             =   4560
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   1138
      Caption         =   "Top Rekrut"
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
   Begin BiSAButtonProject.BiSAButton BiSAButton4 
      Height          =   495
      Left            =   330
      TabIndex        =   8
      Top             =   4575
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   873
      Caption         =   "Stiker"
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
      Height          =   525
      Left            =   1455
      TabIndex        =   7
      Top             =   3375
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   926
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
   Begin BiSAButtonProject.BiSAButton BiSAButton2 
      Height          =   480
      Left            =   945
      TabIndex        =   6
      Top             =   2745
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   847
      Caption         =   "Rebutan Promo"
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
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   105
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1875
      Visible         =   0   'False
      Width           =   975
   End
   Begin BiSAButtonProject.BiSAButton BiSAButton1 
      Height          =   480
      Left            =   1425
      TabIndex        =   4
      Top             =   1935
      Width           =   1755
      _ExtentX        =   3096
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
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      Top             =   1440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin BiSADateProject.BiSADate dTgl 
      Height          =   360
      Left            =   1410
      TabIndex        =   0
      Top             =   645
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   635
      Value           =   "16-02-2013"
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
      Left            =   1410
      TabIndex        =   1
      Top             =   1095
      Visible         =   0   'False
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
   Begin VB.Label Label1 
      Caption         =   "Tampilkan member yg aktif sejak"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1380
      TabIndex        =   2
      Top             =   300
      Width           =   2715
   End
End
Attribute VB_Name = "trSMSMemberYgAktif"
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
Dim nTotalRec As Integer
Dim nSisaBagi As Integer
Dim nRecExcel As Single
Dim a As New exportExcel
Dim n As Single
Dim i As Single


nRecExcel = 100000

cSQL = " select count(nama) as jumlah from anggota"
cSQL = cSQL & " where telp <> '' and lastactivity >='" & Format(dTgl.Value, "yyyy-MM-dd") & "'"
'cSQL = cSQL & " where telp <> ''"

'cSQL = cSQL & " and (alamat like '%rendang%' or alamat like '%menanga%' or  alamat like '%karangasem%')"
cSQL = cSQL & " ORDER BY lastactivity DESC"

Set dbData = objData.Sql(GetDSN, cSQL)
If Not dbData.EOF Then
  nTotalRec = GetNull(dbData!jumlah)
End If
nSisaBagi = nTotalRec - ((nTotalRec \ nRecExcel) * nRecExcel)
nTotalRec = nTotalRec \ nRecExcel

'If nSisaBagi >= 0 Then
'  nTotalRec = nTotalRec + 1
'End If

'If nTotalRec > 0 Then
  cSQL = "select telp as nohp,nama,lastactivity,tgl from anggota"
'  cSQL = cSQL & " where telp <> ''"
  cSQL = cSQL & " where telp <> '' and lastactivity >='" & Format(dTgl.Value, "yyyy-MM-dd") & "'"
'  cSQL = cSQL & " and (alamat like '%rendang%' or alamat like '%menanga%' or  alamat like '%karangasem%')"
  cSQL = cSQL & " ORDER BY lastactivity DESC"
  
  vaArray.ReDim 0, 0, 0, 4
  
  For n = 0 To nTotalRec
    Set dbData = objData.Sql(GetDSN, cSQL & " LIMIT " & nRecExcel * n & "," & nRecExcel)
    If Not dbData.EOF Then
      FrmPB.InitPB dbData.RecordCount
      vaArray(0, 0) = "NOHP"
      vaArray(0, 1) = "NAMA"
      vaArray(0, 2) = "LAST ACTIVITY"
      vaArray(0, 3) = "TGL JOIN"
      vaArray(0, 4) = "OP"
  
      Do While Not dbData.EOF
        FrmPB.RunPB
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        i = vaArray.UpperBound(1)
        vaArray(i, 0) = "'" & GetNull(dbData!nohp)
        vaArray(i, 1) = GetNull(dbData!nama)
        vaArray(i, 2) = GetNull(dbData!lastactivity)
        vaArray(i, 3) = GetNull(dbData!Tgl)
        vaArray(i, 4) = GetSelectOperatorHP(GetNull(dbData!nohp))
        
        If vaArray(i, 4) <> "IM3" And vaArray(i, 4) <> "MENTARI" Then
          vaArray.DeleteRows i
        End If
        
'        If vaArray(i, 4) = "IM3" Or vaArray(i, 4) = "MENTARI" Then
'          vaArray.DeleteRows i
'        End If
      
        dbData.MoveNext
      Loop
      FrmPB.EndPB
    End If
    'vaArray.QuickSort 1, vaArray.UpperBound(1), 4, XORDER_ASCEND, XTYPE_STRING
    MsgBox "Record yg berhasil di proses " & vaArray.UpperBound(1)
    a.RecordSource = vaArray
    a.ExportToExcel
    vaArray.ReDim 0, 0, 0, 4
 
  Next n
  
  MsgBox "SELESAI", vbInformation
'End If

End Sub

Private Function GetPenjualanMember2(ByVal obj As CodeSuiteLibrary.Data, ByVal cMember As String) As Double
Dim db As New ADODB.Recordset
Dim cSQL As String

  GetPenjualanMember2 = 0
  'cSQL = "select sum(p.jumlah) as jumlah from penjualan p"
  cSQL = "select sum(p.harga*p.qty) as jumlah from penjualan p"
  cSQL = cSQL & " left join totpenjualan t on t.nomorpenjualan = p.nomorpenjualan"
  cSQL = cSQL & " Where p.Discount >= 30 And t.kodeanggota = '" & cMember & "' And p.Tgl >= '" & Format(dTgl(0).Value, "yyyy-MM-dd") & "' And p.Tgl <= '" & Format(dTgl(1).Value, "yyyy-MM-dd") & "'"
  Set db = obj.Sql(GetDSN, cSQL)
  If Not db.EOF Then
    GetPenjualanMember2 = GetNull(db!jumlah)
  End If
End Function

Private Sub BiSAButton2_Click()
Dim cSQL As String
Dim db As New ADODB.Recordset
Dim cTot As Double


cSQL = ""
Set dbData = objData.Browse(GetDSN, "promo_rebutan")
If Not dbData.EOF Then
  FrmPB.InitPB dbData.RecordCount
  Do While Not dbData.EOF
    FrmPB.RunPB
    
'    cSQL = "select m.*,p.pgrade,p.cabang from promo_merge m "
'    cSQL = cSQL & " left join promo_pgrade p on p.kodeanggota = m.kodeanggota"
'    cSQL = cSQL & " where barcode = '" & GetNull(dbData!barcode) & "'"
'    cSQL = cSQL & " ORDER BY pgrade DESC LIMIT 0," & GetNull(dbData!qty)
    
    cSQL = "SELECT * from promo_final2"
    cSQL = cSQL & " where barcode = '" & GetNull(dbData!barcode) & "'"
    cSQL = cSQL & " ORDER BY total DESC LIMIT 0," & GetNull(dbData!qty)
    cTot = 0
    Set db = objData.Sql(GetDSN, cSQL)
    If Not db.EOF Then
      Do While Not db.EOF
        'insert ke tabel
        objData.Add GetDSN, "promo_final", Array("barcode", "qty", "harga", "kodeanggota", "nama", "toko"), Array(GetNull(db!barcode), 1, GetNull(dbData!Harga), GetNull(db!kodeanggota), GetNull(db!nama), GetNull(db!toko))
        cTot = cTot + 1
        db.MoveNext
      Loop
    End If
    dbData.MoveNext
  Loop
  FrmPB.EndPB
  '###############################
  'TABEL HASI AKHIR
  '###############################
  'BARCODE
  'QTY
  'HARGA
  'KODEANGGOTA
  'NAMA
  'TOKO
End If
End Sub

Private Sub BiSAButton3_Click()
Dim cSQL As String
Dim nTotalRec As Integer
Dim nSisaBagi As Integer
Dim nRecExcel As Single
Dim a As New exportExcel
Dim n As Single
Dim i As Single


nRecExcel = 10000

cSQL = " select count(nama) as jumlah from anggota"
cSQL = cSQL & " where telp <> '' and lastactivity >='" & Format(dTgl.Value, "yyyy-MM-dd") & "'"
cSQL = cSQL & " ORDER BY lastactivity DESC"

Set dbData = objData.Sql(GetDSN, cSQL)
If Not dbData.EOF Then
  nTotalRec = GetNull(dbData!jumlah)
End If
nSisaBagi = nTotalRec - ((nTotalRec \ nRecExcel) * nRecExcel)
nTotalRec = nTotalRec \ nRecExcel

  cSQL = "select telp as nohp,nama,lastactivity,tgl,kodeanggota from anggota"
  cSQL = cSQL & " where telp <> '' and lastactivity >='" & Format(dTgl.Value, "yyyy-MM-dd") & "'"
  cSQL = cSQL & " ORDER BY lastactivity DESC"
  
  vaArray.ReDim 0, 0, 0, 16
  
  For n = 0 To nTotalRec
    Set dbData = objData.Sql(GetDSN, cSQL & " LIMIT " & nRecExcel * n & "," & nRecExcel)
    If Not dbData.EOF Then
      FrmPB.InitPB dbData.RecordCount
      vaArray(0, 0) = "NOHP"
      vaArray(0, 1) = "NAMA"
      vaArray(0, 2) = "LAST ACTIVITY"
      vaArray(0, 3) = "TGL JOIN"
      vaArray(0, 4) = "OP"
  
      Do While Not dbData.EOF
        FrmPB.RunPB
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        i = vaArray.UpperBound(1)
        vaArray(i, 0) = "'" & GetNull(dbData!nohp)
        vaArray(i, 1) = GetNull(dbData!nama)
        vaArray(i, 2) = GetNull(dbData!lastactivity)
        vaArray(i, 3) = GetNull(dbData!Tgl)
        vaArray(i, 4) = GetSelectOperatorHP(GetNull(dbData!nohp))
        'penjualan bulan 1
        vaArray(i, 5) = getPenjualanTahunBulan(objData, 2014, 1, GetNull(dbData!kodeanggota))
        'penjualan bulan 2
        vaArray(i, 6) = getPenjualanTahunBulan(objData, 2014, 2, GetNull(dbData!kodeanggota))
        'penjualan bulan 3
        vaArray(i, 7) = getPenjualanTahunBulan(objData, 2014, 3, GetNull(dbData!kodeanggota))
        'penjualan bulan 4
        vaArray(i, 8) = getPenjualanTahunBulan(objData, 2014, 4, GetNull(dbData!kodeanggota))
        'penjualan bulan 5
        vaArray(i, 9) = getPenjualanTahunBulan(objData, 2014, 5, GetNull(dbData!kodeanggota))
        'penjualan bulan 6
        vaArray(i, 10) = getPenjualanTahunBulan(objData, 2014, 6, GetNull(dbData!kodeanggota))
        'penjualan bulan 7
        vaArray(i, 11) = getPenjualanTahunBulan(objData, 2014, 7, GetNull(dbData!kodeanggota))
        'penjualan bulan 8
        vaArray(i, 12) = getPenjualanTahunBulan(objData, 2014, 8, GetNull(dbData!kodeanggota))
        'penjualan bulan 9
        vaArray(i, 13) = getPenjualanTahunBulan(objData, 2014, 9, GetNull(dbData!kodeanggota))
        'penjualan bulan 10
        vaArray(i, 14) = getPenjualanTahunBulan(objData, 2014, 10, GetNull(dbData!kodeanggota))
        'penjualan bulan 11
        vaArray(i, 15) = getPenjualanTahunBulan(objData, 2014, 11, GetNull(dbData!kodeanggota))
        'penjualan bulan 12
        vaArray(i, 16) = getPenjualanTahunBulan(objData, 2014, 12, GetNull(dbData!kodeanggota))
        dbData.MoveNext
      Loop
      FrmPB.EndPB
    End If
    a.RecordSource = vaArray
    a.ExportToExcel
    vaArray.ReDim 0, 0, 0, 4
  Next n
  
  MsgBox "SELESAI", vbInformation
End Sub

Private Function getPenjualanTahunBulan(ByVal obj As CodeSuiteLibrary.Data, ByVal nTahun As Integer, nBulan As Integer, cKodeAnggota As String) As Double
Dim cSQL As String
Dim db As New ADODB.Recordset

  'Set db = OBJ.Sql(GetDSN, cSQL)
  cSQL = ""
  cSQL = "select a.telp as nohp,g.op,t.kodeanggota,a.nama,u.nama,sum(qty*harga) as total from penjualan p"
  cSQL = cSQL & " left join totpenjualan t on t.nomorpenjualan = p.nomorpenjualan"
  cSQL = cSQL & " left join anggota a on a.kodeanggota = t.kodeanggota"
  cSQL = cSQL & " left join anggota u on u.kodeanggota = a.kodeupline"
  cSQL = cSQL & " left join gsm g on g.prefix_op = left(a.telp,4)"
  cSQL = cSQL & " where YEAR(p.tgl)= " & nTahun & " and MONTH(p.tgl) = " & nBulan & " and a.telp <> '' and t.kodeanggota = '" & cKodeAnggota & "'"
  cSQL = cSQL & " GROUP BY t.kodeanggota"
  cSQL = cSQL & " ORDER BY total DESC"
  Set db = obj.Sql(GetDSN, cSQL)
  
  If Not db.EOF Then
    getPenjualanTahunBulan = GetNull(db!Total)
  End If
End Function

Private Sub BiSAButton4_Click()
  Load trCetakStikerBarcode
  trCetakStikerBarcode.Show
End Sub

Private Sub BiSAButton5_Click()
Dim db As New ADODB.Recordset
Dim n As Single
  
  vaArray.ReDim 0, -1, 0, 3
  Set db = objData.Browse(GetDSN, "anggota")
  If Not db.EOF Then
    FrmPB.InitPB db.RecordCount
    Do While Not db.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      FrmPB.RunPB
      'cari berapa jumlah rekrutan member ini
      vaArray(n, 0) = GetNull(db!telp)
      vaArray(n, 1) = GetNull(db!nama)
      vaArray(n, 2) = GetNull(db!kodeanggota)
      vaArray(n, 3) = GetRekrutan(objData, vaArray(n, 2), "2014-1-1", "2014-12-31")
    Loop
    FrmPB.EndPB
  End If
End Sub

Private Function GetRekrutan(ByVal obj As CodeSuiteLibrary.Data, cAnggota As String, dTgA As Date, dTgZ As Date) As Double
Dim db As New ADODB.Recordset
  Set db = obj.Sql(GetDSN, "select count(kodeupline) as total where kodeupline = '" & cAnggota & "' and tgl >= '" & Format(dTgA, "yyyy-MM-dd") & "' and tgl <= '" & Format(dTgZ, "yyyy-MM-dd") & "'")
  If Not db.EOF Then
    GetRekrutan = GetNull(db!Total)
  End If
End Function

Private Sub cmdOK_Click()
  GetSMS
'Dim cSQL As String
'Dim n As Single
'Dim a As New exportExcel
'
'  cSQL = ""
'  cSQL = cSQL & " select DISTINCT(t.kodeanggota),a.nama,a.telp,t.tgl from totpenjualan t"
'  cSQL = cSQL & " left join anggota a on a.kodeanggota = t.kodeanggota"
'  cSQL = cSQL & " Where t.Tgl >= '" & Format(dTgl.Value, "yyyy-MM-dd") & "' AND a.telp <> ''"
'
'  vaArray.ReDim 0, 1, 0, 3
'
'  vaArray(0, 0) = aCfg(objData, msNamaPerusahaan) & " " & aCfg(objData, msAlamatPerusahaan)
'
'  vaArray(1, 0) = "NOHP"
'  vaArray(1, 1) = "OPERATOR"
'  vaArray(1, 2) = "NAMA"
'  vaArray(1, 3) = "TGL"
'
'  vaArray.DefaultColumnType(1) = XTYPE_STRING
'  Set dbData = objData.Sql(GetDSN, cSQL)
'  If Not dbData.EOF Then
'    FrmPB.InitPB dbData.RecordCount
'    Do While Not dbData.EOF
'      FrmPB.RunPB
'      vaArray.InsertRows vaArray.UpperBound(1) + 1
'      n = vaArray.UpperBound(1)
'
'      vaArray(n, 0) = "'" & GetNull(dbData!telp)
'      vaArray(n, 1) = GetSelectOperatorHP(GetNull(dbData!telp))
'      vaArray(n, 2) = UCase(GetNull(dbData!nama))
'      vaArray(n, 3) = Format(GetNull(dbData!Tgl), "dd-MM-yyyy")
'      dbData.MoveNext
'    Loop
'    FrmPB.EndPB
'    'export to excel
''    vaArray.QuickSort 2, vaArray.UpperBound(1), 1, XORDER_DESCEND, XTYPE_STRING
'    a.RecordSource = vaArray
'    a.ExportToExcel
'  End If
End Sub

Private Sub GetSMS()
Dim cSQL As String
Dim n As Single
Dim a As New exportExcel
Dim dTglLastAktif As Date


  vaArray.ReDim 0, 1, 0, 3
  
  vaArray(0, 0) = aCfg(objData, msNamaPerusahaan) & " " & aCfg(objData, msAlamatPerusahaan)
  vaArray(1, 0) = "NOHP"
  vaArray(1, 1) = "OPERATOR"
  vaArray(1, 2) = "NAMA"
  vaArray(1, 3) = "TGL"
  
  Set dbData = objData.Browse(GetDSN, "anggota", , "telp", sisDifference, "")
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.EOF
      FrmPB.RunPB
      If GetAktifDariTgl(objData, GetNull(dbData!kodeanggota), dTgl.Value, dTglLastAktif) Then
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        n = vaArray.UpperBound(1)
        vaArray(n, 0) = "'" & GetNull(dbData!telp)
        vaArray(n, 1) = GetSelectOperatorHP(GetNull(dbData!telp))
        vaArray(n, 2) = UCase(GetNull(dbData!nama))
        vaArray(n, 3) = dTglLastAktif
        If vaArray(n, 1) <> "IM3" Or vaArray(n, 1) <> "MENTARI" Then
          vaArray.DeleteRows n
        End If
        objData.Update GetDSN, "anggota", "kodeanggota = '" & dbData!kodeanggota & "'", Array("lastactivity"), Array(vaArray(n, 3))
      End If
      dbData.MoveNext
    Loop
    FrmPB.EndPB
    vaArray.QuickSort 2, vaArray.UpperBound(1), 3, XORDER_DESCEND, XTYPE_DATE, 1, XORDER_DESCEND, XTYPE_STRING
    
'    Dim i As Interior
'    i = 0
'    For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
'      If i = 25 Then i = 0
'      i = i + 1
'      If i = 25 Then
'
'      End If
'    Next n
    
    a.RecordSource = vaArray
    a.ExportToExcel
  End If
End Sub

Private Function GetAktifDariTgl(ByVal obj As CodeSuiteLibrary.Data, ByVal cKodeAnggota As String, ByVal dTgl As Date, ByRef dTglLastAktif As Date) As Boolean
Dim db As New ADODB.Recordset

  GetAktifDariTgl = False
  Set db = obj.Browse(GetDSN, "totpenjualan", , "kodeanggota", sisAssign, cKodeAnggota, " and tgl >= '" & Format(dTgl, "yyyy-MM-dd") & "'", "tgl desc", , 0, 1)
  If Not db.EOF Then
    dTglLastAktif = GetNull(db!Tgl)
    GetAktifDariTgl = True
  End If
  
End Function

Private Sub Command1_Click()
'posting keaktivan member
Dim lSave As Boolean

Set dbData = objData.Browse(GetDSN, "anggota")
If Not dbData.EOF Then
  FrmPB.InitPB dbData.RecordCount
  Do While Not dbData.EOF
    FrmPB.RunPB
    objData.Update GetDSN, "anggota", "kodeanggota = '" & GetNull(dbData!kodeanggota) & "'", Array("lastactivity"), Array((GetLasActivityMember(objData, GetNull(dbData!kodeanggota))))
    dbData.MoveNext
  Loop
  FrmPB.EndPB
End If

End Sub

Private Function GetLasActivityMember(ByVal obj As CodeSuiteLibrary.Data, ByVal cKodeMember As String) As Date
Dim db As New ADODB.Recordset
Dim cSQL As String

  GetLasActivityMember = "2000-01-01"
  cSQL = "select t.tgl FROM penjualan p"
  cSQL = cSQL & " left join totpenjualan t on t.nomorpenjualan = p.nomorpenjualan"
  cSQL = cSQL & " where t.kodeanggota = '" & cKodeMember & "'"
  cSQL = cSQL & " ORDER BY t.tgl DESC LIMIT 0,1"
  
  Set db = obj.Sql(GetDSN, cSQL)
  If Not db.EOF Then
    GetLasActivityMember = GetNull(db!Tgl)
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
