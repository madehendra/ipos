Attribute VB_Name = "Function"
Option Explicit



Dim dbData As New adodb.Recordset
Dim objData As New CodeSuiteLibrary.Data
Public cKasTeller As String
Public cNamaKasTeller As String
Public nRecordsTrial As Double
Public isTrial As Boolean

Public Enum SisCfgTransaksi
  vPembelian = 0
  vPenjualan = 1
  vRtnPembelianDenganFaktur = 2
  vRtnPenjualanDenganFaktur = 3
  vRtnPembelianTanpaFaktur = 4
  vRtnPenjualanTanpaFaktur = 5
End Enum

Public Enum SisModelReturPembelian
  vDefault = 0
  vTunai = 1
  vHutang = 2
  vTitip = 3
End Enum

Public Enum SisNomorFaktur
  fkt_Pembelian = 0
  fkt_ReturPembelian = 1
  fkt_Penjualan = 2
  fkt_PenjualanKasir = 3
  fkt_ReturPenjualan = 4
  fkt_PenyesuaianStock = 5
  fkt_MutasiAntarGudang = 6
  fkt_PelunasanPiutang = 7
  fkt_PelunasanHutang = 8
  fkt_PiutangLainLain = 9
  fkt_HutangLainLain = 10
  fkt_KasBank = 11
  fkt_BiayaLainLain = 12
  fkt_PendapatanLainLain = 13
  fkt_JurnalLain = 14
  fkt_MutasiStock = 15
  
  fkt_OrderPembelian = 16
  fkt_OrderPenjualan = 17
End Enum

Function Level(cRekening As String) As Single
Dim n As Single, nCount As Single
  cRekening = Mid(Trim(cRekening), 3)
  nCount = 1
  For n = 1 To Len(cRekening)
    If Mid(cRekening, n, 1) = "." Then
      nCount = nCount + 1
    End If
  Next
  Level = nCount
End Function
Function GetGroupSales(ByVal cKodeG As String, ByVal obj As CodeSuiteLibrary.Data) As String
Dim dba As New adodb.Recordset

  GetGroupSales = ""
  Set dba = obj.Browse(GetDSN, "groupsales", "kode,keterangan", "status", sisAssign, 1, " and kode = '" & cKodeG & "'")
  If Not dba.EOF Then
    GetGroupSales = cKodeG
  End If
End Function

Public Function GetMemberPiutang(cMember As String) As Double
Dim db As New adodb.Recordset

  GetMemberPiutang = 0
  
  Set db = objData.Browse(GetDSN, "kartupiutang", "sum(debet-kredit) as saldopiutang", "kodeanggota", sisAssign, cMember)
  If Not db.EOF Then
    GetMemberPiutang = GetNull(db!saldopiutang)
  End If
End Function

Public Function GetMemberTopUp(cMember As String) As Double
Dim db As New adodb.Recordset

  GetMemberTopUp = 0
  
  Set db = objData.Browse(GetDSN, "membertopup", "sum(debet-kredit) as saldotopup", "kodeanggota", sisAssign, cMember)
  If Not db.EOF Then
    GetMemberTopUp = GetNull(db!saldotopup)
  End If
End Function

Public Function GetSelectOperatorHP(ByVal cPrefixOperator As String) As String

GetSelectOperatorHP = "LAINNYA"
cPrefixOperator = Left(cPrefixOperator, 4)

'Kartu As : 0852, 0853
If cPrefixOperator = "0852" Or cPrefixOperator = "0853" Then
  GetSelectOperatorHP = "AS"

End If

'Axis : 0831, 0838
If cPrefixOperator = "0831" Or cPrefixOperator = "0838" Then
  GetSelectOperatorHP = "AXIS"

End If

'Im3 : 0856, 0857
If cPrefixOperator = "0856" Or cPrefixOperator = "0857" Then
  GetSelectOperatorHP = "IM3"

End If

'Mentari : 0815, 0816, 0858
If cPrefixOperator = "0815" Or cPrefixOperator = "0816" Or cPrefixOperator = "0858" Then
  GetSelectOperatorHP = "MENTARI"

End If

'Simpati : 0811, 0812, 0813, 0821
If cPrefixOperator = "0811" Or cPrefixOperator = "0812" Or cPrefixOperator = "0813" Or cPrefixOperator = "0821" Then
  GetSelectOperatorHP = "SIMPATI"

End If

'XL : 0817, 0818, 0819, 0859,0877, 0878
If cPrefixOperator = "0817" Or cPrefixOperator = "0818" Or cPrefixOperator = "0819" Or cPrefixOperator = "0859" Or cPrefixOperator = "0877" Or cPrefixOperator = "0878" Then
  GetSelectOperatorHP = "XL"
End If

'Fren : 0885, 0886, 0887, 0888
'Smart : 0881, 0882, 0883, 0884
'Three : 0896, 0897, 0898, 0899

End Function

Function isLunas(ByVal obj As CodeSuiteLibrary.Data, ByVal nomorpenjualan As String, ByRef SisaPiutang As Double) As Boolean
Dim db As New adodb.Recordset
Dim Piutang As Double
Dim Lunas As Double
Dim Retur As Double

  isLunas = True
  Piutang = 0
  Lunas = 0
  SisaPiutang = 0
  
  Set db = obj.Browse(GetDSN, "totpenjualan", "piutang", "nomorpenjualan", sisAssign, nomorpenjualan)
  If Not db.EOF Then
    Piutang = GetNull(db!Piutang)
  End If

  'Retur
  Set db = obj.Browse(GetDSN, "totrtnpenjualan", "sum(total) as total", "nomorpenjualan", sisAssign, nomorpenjualan)
  If Not db.EOF Then
    Retur = GetNull(db!Total)
  End If
  Set db = obj.Browse(GetDSN, "pelunasanpiutang", "sum(discount+pelunasan) as totallunas", "nomorpenjualan", sisAssign, nomorpenjualan)
  If Not db.EOF Then
    Lunas = GetNull(db!totalLunas)
  End If
  SisaPiutang = Piutang - Lunas - Retur
  If SisaPiutang <> 0 Then
    isLunas = False
  End If
End Function

Function isPernahBayar(ByVal obj As CodeSuiteLibrary.Data, ByVal nomorpenjualan As String, ByRef nPernahBayar As Double) As Boolean
Dim db As New adodb.Recordset

  isPernahBayar = False
  nPernahBayar = 0
  
  Set db = obj.Browse(GetDSN, "pelunasanpiutang", "sum(discount+pelunasan) as totallunas", "nomorpenjualan", sisAssign, nomorpenjualan)
  If Not db.EOF Then
    nPernahBayar = GetNull(db!totalLunas)
    If nPernahBayar > 0 Then
      isPernahBayar = True
    End If
  End If
  
End Function

Function GetLevel(ByVal cRekening As String, ByVal nLevel As Single)
Dim nLeft As Single
  Select Case nLevel
    Case 1
      nLeft = 5
    Case 2
      nLeft = 9
    Case 3
      nLeft = 13
    Case 4
      nLeft = 15
  End Select
  GetLevel = Left(cRekening, nLeft)
End Function

Public Function GetValidDataBrowse(ByVal obj As CodeSuiteLibrary.Data, ByVal cTable As String, ByVal cField As String, ByVal cKey As String) As Boolean
Dim db As New adodb.Recordset
GetValidDataBrowse = True

  Set db = obj.Browse(GetDSN, cTable, , cField, sisAssign, cKey)
  If db.EOF Then
    GetValidDataBrowse = False
    Exit Function
  End If
End Function

Function GetLastFaktur(ByVal nPar As SisNomorFaktur, Optional ByVal cFaktur As String = "", Optional ByVal lUpdate As Boolean = False) As String
Dim vaFaktur
Dim db As New adodb.Recordset
Dim obj As New CodeSuiteLibrary.Data
Dim cChar As String
Dim cNomor As String
Dim nCount As Double
Dim cKode As String

  cNomor = 1
  vaFaktur = Array("PB", "RB", "PJ", "CS", "RJ", _
                   "AD", "MT", "PP", "PH", "PL", _
                   "HL", "KB", "CO", "CI", "JR", _
                   "MS", "OB", "OJ")
                   
  
  cChar = vaFaktur(nPar)
  cKode = cChar
  
  If Trim(cFaktur = "") Then
    If lUpdate Then
      
      obj.Add GetDSN, "nomorfaktur", Array("Kode"), Array(cKode)
      Set db = obj.SQL(GetDSN, "Select Last_Insert_id() as Total")
      
      ' Untuk Menghemat Ukuran Table Hapus jika Nomor ID < Nomor yang aktif
      If db.RecordCount > 0 Then
        obj.Delete GetDSN, "NomorFaktur", "Kode", sisAssign, cKode, " and id < " & GetNull(db!Total)
      End If

      
      nCount = 0
    Else
      'Set db = obj.Browse(GetDSN, "NomorFaktur", "Max(ID) as Total", "Kode", sisAssign, cKode)
      Set db = obj.Browse(GetDSN, "NomorFaktur", "ID as Total", "Kode", sisAssign, cKode)
      nCount = 1
    End If
    If db.RecordCount > 0 Then
      cNomor = GetNull(db!Total, 0) + nCount
    End If
  Else
    cNomor = Trim(cFaktur)
  End If

  cNomor = Padl(cNomor, 10, "0")
  cNomor = cChar & Mid(cNomor, 3)
  GetLastFaktur = cNomor
End Function

Function UpdCfgTransaksi(ByVal nPar As SisCfgTransaksi, ByVal nLevel As Double, ByVal nGudang As String, ByVal cSTDGudang As String, _
                         ByVal nDisc1 As String, ByVal nDisc2 As String)
Dim vaField, vaValue
Dim c As String
  c = nGudang & nDisc1 & nDisc2 & cSTDGudang
  vaField = Array("Status", "Level", "Keterangan")
  vaValue = Array(nPar, nLevel, c)
  objData.Update GetDSN, "CfgTransaksi", "Status = '" & nPar & "' and Level = " & nLevel, vaField, vaValue
End Function

Function GetCfgTransaksi(ByVal nPar As SisCfgTransaksi, ByVal nLevel As Single, nGudang As String, cSTDGudang As String, _
                         nDisc1 As String, nDisc2 As String)
Dim cKeterangan As String

  cKeterangan = "111"
  Set dbData = objData.Browse(GetDSN, "CfgTransaksi", , "Status", sisAssign, nPar, " and Level = " & nLevel)
  If dbData.RecordCount > 0 Then
    cKeterangan = GetNull(dbData!keterangan)
  End If
  
  nGudang = Mid(cKeterangan, 1, 1)
  nDisc1 = Mid(cKeterangan, 2, 1)
  nDisc2 = Mid(cKeterangan, 3, 1)
  cSTDGudang = Mid(cKeterangan, 4, 6)
End Function

Sub SetupTransaksi(ByVal nPar As SisCfgTransaksi, cGudang As Object, cNamaGudang As Object, nDiscount As Object, nDiscount2 As Object, nPersDisc As Object, nPersDisc2 As Object)
Dim cSTDGudang As String
Dim cSTDDisc1 As String
Dim cSTDDisc2 As String
Dim cgud As String

  GetCfgTransaksi nPar, GetRegistry(reg_UserLevel), cSTDGudang, cgud, cSTDDisc1, cSTDDisc2
  
  cGudang.Text = cgud
  cGudang.Enabled = cSTDGudang = "1"
  nDiscount.Enabled = cSTDDisc1 <> "1"
  nDiscount2.Enabled = cSTDDisc2 <> "1"
  nPersDisc.Enabled = Not nDiscount.Enabled
  nPersDisc2.Enabled = Not nDiscount2.Enabled
  
  cGudang.BackColor = IIf(cGudang.Enabled, vbWindowBackground, vbButtonFace)
  nDiscount.BackColor = IIf(nDiscount.Enabled, vbWindowBackground, vbButtonFace)
  nDiscount2.BackColor = IIf(nDiscount2.Enabled, vbWindowBackground, vbButtonFace)
  nPersDisc.BackColor = IIf(nPersDisc.Enabled, vbWindowBackground, vbButtonFace)
  nPersDisc2.BackColor = IIf(nPersDisc2.Enabled, vbWindowBackground, vbButtonFace)
  
  
  Set dbData = objData.Browse(GetDSN, "Gudang", , "Kode", sisAssign, cGudang.Text)
  If dbData.RecordCount > 0 Then
    cNamaGudang.Text = GetNull(dbData!keterangan)
  End If
End Sub

'Function GetDSN() As String
'  GetDSN = "Provider=MSDASQL.1;Persist Security Info=True;Data Source=" & GetRegistry(reg_DSN)
'End Function

Function GetSaldoHutangByFaktur(ByVal cFakturPembelian As String) As Double
Dim objData As New CodeSuiteLibrary.Data
Dim totalPembelian As Double
Dim totalRetur As Double
Dim totalLunas As Double
Dim dbSaldoHutang As New adodb.Recordset


'fungsi untuk mengambil saldo hutang berdasarkan no faktur tertentu
'rumus saldo hutang
'saldo piutang = totalpembelian - totalretur - totalpelunasan

  Set dbSaldoHutang = objData.Browse(GetDSN, "totPembelian t", "t.total,t.hutang", "t.Faktur", sisAssign, cFakturPembelian)
  If Not dbSaldoHutang.EOF Then
    totalPembelian = GetNull(dbSaldoHutang!hutang)
  Else
    totalPembelian = 0
  End If
  
  Set dbSaldoHutang = objData.Browse(GetDSN, "totRtnPembelian", "sum(total) as totalRetur", "fktPembelian", sisAssign, cFakturPembelian, " group by fktPembelian")
  If Not dbSaldoHutang.EOF Then
    totalRetur = GetNull(dbSaldoHutang!totalRetur)
  Else
    totalRetur = 0
  End If
  
  Set dbSaldoHutang = objData.Browse(GetDSN, "PelunasanHutang", "sum(totalLunas) as totalLunas, sum(Discount) as Discount", "fkt", sisAssign, cFakturPembelian, " group by fkt")
  If Not dbSaldoHutang.EOF Then
    totalLunas = GetNull(dbSaldoHutang!totalLunas) + GetNull(dbSaldoHutang!Discount)
  Else
    totalLunas = 0
  End If
  GetSaldoHutangByFaktur = totalPembelian - totalRetur - totalLunas
End Function

Function GetSaldoPiutangByFaktur(ByVal cFakturPenjualan As String) As Double
Dim dbSaldoPiutang As adodb.Recordset
Dim cSQL As String
Dim objData As New CodeSuiteLibrary.Data
Dim totalPenjualan As Double
Dim totalRetur As Double
Dim totalLunas As Double

'fungsi untuk mengambil saldo piutang berdasarkan no faktur tertentu
'rumus saldo piutang
'saldo piutang = totalpenjualan - totalretur - totalpelunasan

  Set dbSaldoPiutang = objData.Browse(GetDSN, "totPenjualan t", "t.total,t.piutang", "t.Faktur", sisAssign, cFakturPenjualan)
  If Not dbSaldoPiutang.EOF Then
    totalPenjualan = GetNull(dbSaldoPiutang!Piutang)
  Else
    totalPenjualan = 0
  End If
  
  Set dbSaldoPiutang = objData.Browse(GetDSN, "totRtnPenjualan", "sum(total) as totalRetur", "fktPenjualan", sisAssign, cFakturPenjualan, " group by fktPenjualan")
  If Not dbSaldoPiutang.EOF Then
    totalRetur = GetNull(dbSaldoPiutang!totalRetur)
  Else
    totalRetur = 0
  End If
  
  Set dbSaldoPiutang = objData.Browse(GetDSN, "PelunasanPiutang", "sum(totalLunas) as totalLunas, sum(Discount) as Discount", "fkt", sisAssign, cFakturPenjualan, " group by fkt")
  If Not dbSaldoPiutang.EOF Then
    totalLunas = GetNull(dbSaldoPiutang!totalLunas) + GetNull(dbSaldoPiutang!Discount)
  Else
    totalLunas = 0
  End If
  GetSaldoPiutangByFaktur = CCur(totalPenjualan) - CCur(totalRetur) - CCur(totalLunas)
End Function

Function GetSaldoPiutang(ByVal obj As CodeSuiteLibrary.Data, ByVal Cust As String, Optional ByVal GroupSales As String = "") As Double
Dim db As New adodb.Recordset
Dim cSQL As String

  cSQL = ""
  If Trim(GroupSales) <> "" Then
    cSQL = " and groupsales = '" & GroupSales & "'"
  End If
  GetSaldoPiutang = 0
  Set db = obj.Browse(GetDSN, "kartupiutang ", "sum(debet-kredit) as saldopiutang", "kodeanggota", sisAssign, Cust, cSQL)
  If Not db.EOF Then
    GetSaldoPiutang = GetNull(db!saldopiutang)
  End If
End Function

Function GetSaldoTopUpMember(ByVal obj As CodeSuiteLibrary.Data, ByVal cMember As String)
Dim db As New adodb.Recordset

  'cari saldo top up member
  GetSaldoTopUpMember = 0
  Set db = obj.Browse(GetDSN, "membertopup m", "m.kodeanggota,a.nama,a.alamat,sum(debet) as debet,sum(kredit) as kredit,sum(m.debet-m.kredit) as saldo", "m.kodeanggota", sisAssign, cMember, " GROUP BY m.kodeanggota", , Array("left join anggota a on a.kodeanggota = m.kodeanggota"))
  If Not db.EOF Then
    GetSaldoTopUpMember = GetNull(db!debet) - GetNull(db!kredit)
  End If
End Function

Function GetSaldoTopUpMember2(ByVal obj As CodeSuiteLibrary.Data, ByVal cMember As String)
Dim db As New adodb.Recordset

  'cari saldo top up member
  GetSaldoTopUpMember2 = 0
  Set db = obj.Browse(GetDSN, "membertopup m", "m.kodeanggota,a.nama,a.alamat,sum(debet) as debet,sum(kredit) as kredit,sum(m.debet-m.kredit) as saldo", "m.kodeanggota", sisAssign, cMember, " AND m.tgl >= '2012-06-01' GROUP BY m.kodeanggota", , Array("left join anggota a on a.kodeanggota = m.kodeanggota"))
  If Not db.EOF Then
    GetSaldoTopUpMember2 = GetNull(db!saldo)
  End If
End Function

Function GetSaldoPiutangByCustomer(ByVal cCustomer As String) As Double
Dim dbSaldoPiutang As adodb.Recordset
Dim cSQL As String
Dim objData As New CodeSuiteLibrary.Data
Dim totalPenjualan As Double
Dim totalRetur As Double
Dim totalLunas As Double

'fungsi untuk mengambil saldo piutang berdasarkan no faktur tertentu
'rumus saldo piutang
'saldo piutang = totalpenjualan - totalretur - totalpelunasan

  Set dbSaldoPiutang = objData.Browse(GetDSN, "totPenjualan t", "t.total,sum(t.piutang) as piutang", "t.kodeanggota", sisAssign, cCustomer, " Group by kodeanggota")
  If Not dbSaldoPiutang.EOF Then
    totalPenjualan = GetNull(dbSaldoPiutang!Piutang)
  Else
    totalPenjualan = 0
  End If
  
  Set dbSaldoPiutang = objData.Browse(GetDSN, "totRtnPenjualan", "sum(total) as totalRetur", "kodeanggota", sisAssign, cCustomer, " group by kodeanggota")
  If Not dbSaldoPiutang.EOF Then
    totalRetur = GetNull(dbSaldoPiutang!totalRetur)
  Else
    totalRetur = 0
  End If
  
  Set dbSaldoPiutang = objData.Browse(GetDSN, "totPelunasanPiutang", "sum(total) as totalLunas,sum(Discount) as Discount", "kodeanggota", sisAssign, cCustomer, " group by kodeanggota")
  If Not dbSaldoPiutang.EOF Then
    totalLunas = GetNull(dbSaldoPiutang!totalLunas) + GetNull(dbSaldoPiutang!Discount)
  Else
    totalLunas = 0
  End If
  GetSaldoPiutangByCustomer = totalPenjualan - totalRetur - totalLunas
End Function

Function GetSaldoHutangBySupplier(ByVal cSupplier As String) As Double
Dim dbSaldoHutang As adodb.Recordset
Dim cSQL As String
Dim objData As New CodeSuiteLibrary.Data
Dim totalPenjualan As Double
Dim totalRetur As Double
Dim totalLunas As Double

'fungsi untuk mengambil saldo piutang berdasarkan no faktur tertentu
'rumus saldo piutang
'saldo piutang = totalpenjualan - totalretur - totalpelunasan

  Set dbSaldoHutang = objData.Browse(GetDSN, "totPembelian t", "t.total,sum(t.hutang) as piutang", "t.Supplier", sisAssign, cSupplier, " Group by Supplier")
  If Not dbSaldoHutang.EOF Then
    totalPenjualan = GetNull(dbSaldoHutang!Piutang)
  Else
    totalPenjualan = 0
  End If
  
  Set dbSaldoHutang = objData.Browse(GetDSN, "totRtnPembelian", "sum(total) as totalRetur", "Supplier", sisAssign, cSupplier, " group by Supplier")
  If Not dbSaldoHutang.EOF Then
    totalRetur = GetNull(dbSaldoHutang!totalRetur)
  Else
    totalRetur = 0
  End If
  
  Set dbSaldoHutang = objData.Browse(GetDSN, "totPelunasanHutang", "sum(total) as totalLunas, sum(Discount) as Discount", "Supplier", sisAssign, cSupplier, " group by Supplier")
  If Not dbSaldoHutang.EOF Then
    totalLunas = GetNull(dbSaldoHutang!totalLunas) + GetNull(dbSaldoHutang!Discount)
  Else
    totalLunas = 0
  End If
  GetSaldoHutangBySupplier = totalPenjualan - totalRetur - totalLunas
End Function

Function CheckData(cData, cMsg) As Boolean
  CheckData = True
  If Len(Trim(cData)) = 0 Or cData = 0 Then
    CheckData = False
    MsgBox (cMsg), vbExclamation, App.Title
  End If
End Function

Sub InitGrid(tdb As TDBGrid)
Dim nSplit As Integer
Dim nCol As Integer
Dim nBack As Double
Dim nFore As Double
Dim nBack1 As Double

  nBack = vbDesktop
  nBack1 = vbButtonFace
  nFore = vbWhite

  tdb.BackColor = vbWhite
  tdb.CaptionStyle.BackColor = nBack
  tdb.CaptionStyle.ForeColor = nFore
  tdb.CaptionStyle.Font.Bold = True
  tdb.Appearance = dbg3D
  tdb.BorderStyle = dbgNoBorder
  tdb.HeadLines = 1
  tdb.HighlightRowStyle.BackColor = &H808080
  tdb.HighlightRowStyle.ForeColor = &HFFFFFF
  'Turn RecordSelectors on or off
  tdb.Splits(0).RecordSelectors = False
  'untuk membuat add row and even row
  tdb.Splits(0).AlternatingRowStyle = False
  tdb.RowDividerStyle = dbgRaised
  'Khusus data combo
'  tdb.AnimateWindowTime = 200
'  tdb.AnimateWindowDirection = dbgAnimateCenter
'  tdb.AnimateWindowClose = dbgOppositeDirection
'  tdb.AnimateWindow = dbgSlide
  For nSplit = 0 To tdb.Splits.Count - 1
    tdb.Splits(nSplit).CaptionStyle.BackColor = nBack
    tdb.Splits(nSplit).CaptionStyle.ForeColor = nFore
    tdb.Splits(nSplit).CaptionStyle.Font.Bold = True
    
    tdb.Splits(nSplit).HeadBackColor = nBack
    tdb.Splits(nSplit).HeadForeColor = nFore
    tdb.Splits(nSplit).HeadFont.Bold = True
    
    tdb.Splits(nSplit).SelectedStyle.BackColor = vbHighlight
    tdb.Splits(nSplit).SelectedStyle.ForeColor = &H8000000E
    tdb.Splits(nSplit).MarqueeStyle = dbgHighlightRow
    
    For nCol = 0 To tdb.Splits(nSplit).Columns.Count - 1
      tdb.Splits(nSplit).Columns(nCol).HeadBackColor = nBack
      tdb.Splits(nSplit).Columns(nCol).HeadForeColor = nFore
      tdb.Splits(nSplit).Columns(nCol).HeadFont.Bold = True
      
      tdb.Splits(nSplit).Columns(nCol).FooterBackColor = nBack
      tdb.Splits(nSplit).Columns(nCol).FooterForeColor = nFore
    Next
  Next
  
  For nCol = 0 To tdb.Columns.Count - 1
    tdb.Columns(nCol).HeadBackColor = nBack
    tdb.Columns(nCol).HeadForeColor = nFore
    tdb.Columns(nCol).HeadFont.Bold = False
  Next
End Sub

Function Padl(Optional cCharacter As String = "", Optional nLen As Byte = 0, Optional cChar = " ") As String
Dim n As Byte, x As String
  x = ""
  If Len(cCharacter) < nLen Then
    For n = 1 To nLen - Len(cCharacter)
      x = cChar & x
    Next
    Padl = x & cCharacter
  Else
    Padl = Mid(cCharacter, 1, nLen)
  End If
End Function

Function Padr(Optional cCharacter As String = "", Optional nLen As Byte = 0, Optional cChar = " ") As String
Dim n As Byte, x As String
  cCharacter = Left(cCharacter, nLen)
  x = ""
  If Len(cCharacter) < nLen Then
    For n = 1 To nLen - Len(cCharacter)
      x = cChar & x
    Next
    Padr = cCharacter + x
  Else
    Padr = Mid(cCharacter, 1, nLen)
  End If
End Function

Function Pad(cCharacter As String, Optional nLen As Byte = 0, Optional cChar = " ") As String
  Pad = Padr(cCharacter, nLen, cChar)
End Function

Function Padc(ByVal cCharacter As String, Optional ByVal nLen As Byte = 0, Optional ByVal cChar As String = " ") As String
  Dim nLeft As Byte
  nLeft = Max(nLen - Len(cCharacter), 0)
  nLeft = IIf(nLeft > 0, Int(nLeft / 2), nLeft)
  Padc = Pad(Replicate(cChar, nLeft) & cCharacter, nLen, cChar)
End Function

Function Max(a, n)
  Max = IIf(a > n, a, n)
End Function

Function Proper(cText) As String
Dim n As Single, cRetval As String, lFirst As Boolean
  On Error Resume Next
  lFirst = True
  For n = 1 To Len(cText)
    If lFirst Then
      lFirst = False
      cRetval = cRetval & UCase(Mid(cText, n, 1))
    Else
      cRetval = cRetval & LCase(Mid(cText, n, 1))
    End If
    
    If Mid(cText, n, 1) = " " Then
      lFirst = True
    End If
  Next
  Proper = cRetval
End Function

Function BOM(ByVal dDate As Date) As Date
  BOM = dDate - Day(dDate) + 1
End Function

Function EOM(ByVal dDate As Date) As Date
Dim n As Byte, OldDate As Date
  OldDate = dDate
  Do While Month(dDate) = Month(OldDate)
    dDate = dDate + 1
  Loop
  EOM = dDate - 1
End Function

Function GetMonth(nMonth As Single) As String
Dim vaMonth
  vaMonth = Array("Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember")
  GetMonth = vaMonth(nMonth - 1)
End Function

Function RAT(cChar, cString As String) As Single
Dim n As Single
  For n = Len(cString) To 1 Step -1
    If Mid(cString, n, 1) = cChar Then
      RAT = n
      Exit Function
    End If
  Next
End Function

Function CreateDSN(cDSN As String, ByVal cIPServer As String, Optional cDatabase As String = "Syariah", Optional cUser As String = "root", Optional cPwd As String = "", Optional cPort As String = "", Optional cMYODBCPATH As String = "", Optional cMYODBCFile As String = "")
Dim cKey As String, x As Long, buffer As String * 255

  ' Ambil Posisi Directory System Window
  x = GetSystemDirectory(buffer, 255)
  buffer = Left(buffer, x)

  ' Register DSN
  SetStringValue "HKEY_LOCAL_MACHINE\Software\ODBC\ODBC.INI\ODBC Data Sources", cDSN, "MySQL ODBC Driver"
  
  ' Configurasi DSN
  cKey = "HKEY_LOCAL_MACHINE\Software\ODBC\ODBC.INI\" & cDSN
  
  DeleteKey cKey
  CreateKey cKey
  SetStringValue cKey, "Database", cDatabase
  SetStringValue cKey, "Description", ""
  
  
'  If Left(GetOsVersion, 1) > 5 Then
'    'tidak support
''    MsgBox ("Maaf OS tidak support")
'    SetStringValue cKey, "Driver", "C:\Program Files\MySQL\Connector ODBC 5.2\myodbc5.dll"
'  Else
'    'support
'    SetStringValue cKey, "Driver", Trim(buffer) & "\myodbc3.dll"
'
'  End If
  
'  SetStringValue cKey, "Driver", "C:\Program Files\MySQL\Connector ODBC 5.2\myodbc5a.dll"
  
'  SetStringValue cKey, "Driver", "C:\Program Files\MariaDB\MariaDB ODBC Driver\maodbc.dll"
'  SetStringValue cKey, "Driver", Trim(buffer) & "\myodbc5a.dll"
'  SetStringValue cKey, "Driver", App.Path & "\myodbc5a.dll"
'  SetStringValue cKey, "Driver", cMYODBCPATH & "\myodbc3.dll"
' sebelumnya tolong di set path dari path installasi myodbc nya

  SetStringValue cKey, "Driver", cMYODBCFile
  SetStringValue cKey, "Option", "3"
  SetStringValue cKey, "Password", cPwd
  SetStringValue cKey, "Port", cPort
  SetStringValue cKey, "Server", cIPServer
  SetStringValue cKey, "Stmt", ""
  SetStringValue cKey, "User", cUser
  SetStringValue cKey, "Uid", cUser
  
  'SetStringValue "HKEY_CURRENT_USER\Control Panel\International", "sShortDate", "dd-MM-yyyy"
End Function

Public Function GetOsArchitecture()
    If IsAtLeastVista Then
        GetOsArchitecture = GetVistaOsArchitecture
    Else
        GetOsArchitecture = GetXpOsArchitecture
    End If
End Function

Private Function IsAtLeastVista() As Boolean
    IsAtLeastVista = GetOsVersion >= "6.0"
End Function

Private Function GetOsVersion() As String
    Dim OperatingSystemSet As Object
    Dim OS As Object

    Set OperatingSystemSet = GetObject("winmgmts:{impersonationLevel=impersonate}"). _
                                    InstancesOf("Win32_OperatingSystem")
    For Each OS In OperatingSystemSet
        GetOsVersion = Left$(Trim$(OS.version), 3)
    Next
End Function

Private Function GetVistaOsArchitecture() As String
    Dim OperatingSystemSet As Object
    Dim OS As Object

    Set OperatingSystemSet = GetObject("Winmgmts:"). _
        ExecQuery("SELECT * FROM Win32_OperatingSystem")
    For Each OS In OperatingSystemSet
        GetVistaOsArchitecture = Left$(Trim$(OS.OSArchitecture), 2)
    Next
End Function

Private Function GetXpOsArchitecture() As String
    Dim ComputerSystemSet As Object
    Dim computer As Object
    Dim SystemType As String

    Set ComputerSystemSet = GetObject("Winmgmts:"). _
        ExecQuery("SELECT * FROM Win32_ComputerSystem")
    For Each computer In ComputerSystemSet
        SystemType = UCase$(Left$(Trim$(computer.SystemType), 3))
    Next

    GetXpOsArchitecture = IIf(SystemType = "X86", "32", "64")
End Function

Function Replicate(cString As String, nCount) As String
Dim n, cRetval As String
  For n = 1 To nCount
    cRetval = cRetval & cString
  Next
  Replicate = cRetval
End Function

Function AScan(ByVal vaArray, ByVal cSearch) As Double
Dim n As Double
  AScan = -1
  For n = 0 To UBound(vaArray)
    If vaArray(n) = cSearch Then
      AScan = n
      Exit Function
    End If
  Next
End Function

Function ATextScan(ByVal vaArray, ByVal cSearch) As Double
Dim n As Double
  ATextScan = -1
  For n = 0 To vaArray.UBound
    If vaArray(n).Text = cSearch Then
      ATextScan = n
      Exit Function
    End If
  Next
End Function

Sub GetMinMax(ByVal cTable As String, vaValue, Optional ByVal cField As String = "kodeakun")
  Set dbData = objData.Browse(GetDSN, cTable, "Min(" & cField & ") as Min, Max(" & cField & ") as Max")
  vaValue(0).Text = ""
  vaValue(1).Text = ""
  If Not dbData.EOF Then
    vaValue(0).Text = "" 'GetNull(GetNull(dbData!Min), "")
    vaValue(1).Text = GetNull(GetNull(dbData!Max), "")
  End If
End Sub

Sub PageSetup(TDBReports1 As TDBReports, cFormName As String)
  GetTDBGSetup cFormName, TDBReports1
  With TDBReports1
    .Profiles(0).Active = True
    .Profiles(0).PreviewMaximized = True
    .Profiles(0).PreviewModal = True
    .Profiles(0).PreviewNoMaximize = True
    .Profiles(0).PreviewNoMinimize = True
    .Profiles(0).PreviewNoResize = True
    .Profiles(0).PreviewNoSaveLoad = True
    .PageSetup
  End With
  SaveTDBGSetup cFormName, TDBReports1
End Sub

Private Sub GetTDBGSetup(ByVal cFormName As String, TDBReports1 As TDBReports)
  cFormName = GetRegistry(reg_Username) & cFormName
  With TDBReports1
    If Trim(cFormName) <> "" Then
      Set dbData = objData.Browse(GetDSN, "tdbgrid", , "Name", sisAssign, cFormName)
      If dbData.RecordCount > 0 Then
        .Profiles(0).PrinterMarginLeft = GetNull(dbData!MarginLeft)
        .Profiles(0).PrinterMarginRight = GetNull(dbData!MarginRight)
        .Profiles(0).PrinterMarginTop = GetNull(dbData!MarginTop)
        .Profiles(0).PrinterMarginBottom = GetNull(dbData!MarginBottom)
        .Profiles(0).PrinterLandscape = GetNull(dbData!Orientation) = 1
        .Profiles(0).PrinterPaperSize = GetNull(dbData!PaperSize)
      End If
    End If
  End With
End Sub

Private Sub SaveTDBGSetup(ByVal cFormName As String, TDBReports1 As TDBReports)
Dim vaField, vaValue
  cFormName = GetRegistry(reg_Username) & cFormName
  With TDBReports1
    If Trim(cFormName) <> "" Then
      vaField = Array("Name", "MarginTop", "MarginLeft", "MarginBottom", "MarginRight", "Orientation", "PaperSize")
      vaValue = Array(cFormName, Round(.Profiles(0).PrinterMarginTop, 2), Round(.Profiles(0).PrinterMarginLeft, 2), _
                      Round(.Profiles(0).PrinterMarginBottom, 2), Round(.Profiles(0).PrinterMarginRight, 2), _
                      Abs(.Profiles(0).PrinterLandscape), .Profiles(0).PrinterPaperSize)
      objData.Update GetDSN, "tdbgrid", "Name = '" & cFormName & "'", vaField, vaValue
    End If
  End With
End Sub

Function BetWeen(ByVal cValue, ByVal cLower, ByVal cUpper) As Boolean
  BetWeen = GetNull(cValue) >= cLower And GetNull(cValue) <= cUpper
End Function

Function GetKode(ByVal cTableName, Optional ByVal cKode As String = "Kode") As adodb.Recordset
  Set GetKode = objData.Browse(GetDSN, cTableName, cKode & " as Kode", , , , , cKode)
End Function

Function GetKodeNavigator(cvKode As String, lCounter As Boolean, lKoreksi As Boolean) As String
Dim cCek As String
Dim x, n As Byte
Dim nDigit As Byte
  cCek = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
   For n = 1 To Len(cvKode)
    nDigit = InStr(1, cCek, Mid(cvKode, n, 1), vbTextCompare)
    If nDigit = 0 Then
      nDigit = n - 1
      n = Len(cvKode)
    End If
  Next
  If lCounter = True Then
    If lKoreksi = True Then
      GetKodeNavigator = Right(cvKode, Len(cvKode) - nDigit)
    Else
      GetKodeNavigator = Padl(Val(Right(cvKode, Len(cvKode) - nDigit)) + 1, Len(cvKode) - nDigit, "0")
    End If
  Else
    GetKodeNavigator = Left(cvKode, nDigit)
  End If
End Function

Function Devide(ByVal a As Double, ByVal b As Double) As Double
  If a = 0 Or b = 0 Then
    Devide = 0
  Else
    Devide = a / b
  End If
End Function

Function CheckDemo(ByVal cTableName As String) As Boolean
Dim cMsg
  CheckDemo = True
End Function

Function GetHPP(ByVal obj As CodeSuiteLibrary.Data, ByVal cKode As String, ByVal dTgl As Date) As Double
Dim db As New adodb.Recordset
Dim nHP As Double

  nHP = 0
  Set db = obj.Browse(GetDSN, "stkhp", "hp", "Kode", sisAssign, cKode, " and Tgl <= '" & SisFormat(dTgl, Sis_yyyy_MM_dd) & "'", "Kode,Tgl Desc", , 0, 1)
  If db.RecordCount > 0 Then
    nHP = GetNull(db!hp)
  End If
  GetHPP = nHP
End Function

Function GetPicture(ByVal cPath As String) As String
  On Error GoTo salah
  If Dir(cPath) <> "" Then
    GetPicture = cPath
  Else
    GetPicture = ""
  End If
  Exit Function
salah:
  GetPicture = ""
End Function

Sub InitGauge(pr As ProgressBar, nMax As Double)
  pr.Visible = True
  pr.Min = 0
  pr.Max = Max(nMax, 1)
  pr.Value = 0
End Sub

Sub RunGauge(pr As ProgressBar)
  pr.Value = pr.Value + IIf(pr.Value < pr.Max, 1, 0)
End Sub

Sub EndGauge(pr As ProgressBar)
  pr.Visible = False
End Sub

Function IsSaldoMinus() As Boolean
IsSaldoMinus = True

'  If aCfg(objData, msSaldoMinus) <> "1" Then
'    IsSaldoMinus = False
'  End If
End Function

Function IsInPeriod(ByVal dTgl As Date) As Boolean
Dim db As New adodb.Recordset
Dim obj As New CodeSuiteLibrary.Data
Dim dAwal As Date
Dim dAkhir As Date
Dim lNull As Boolean

  lNull = False
  IsInPeriod = True
  
  If aCfg(objData, msOptAudit) = "Y" Then
'    Set db = obj.Browse(GetDSN, "periode", "min(awal) as awal,max(akhir) as akhir", "status", sisAssign, "0")
'    If db.RecordCount > 0 Then
'      If IsNull(db!Awal) Or IsNull(db!akhir) Then
'        lNull = True
'      End If
'    Else
'      lNull = True
'    End If
'
'    If lNull Then
'      MsgBox "Tanggal Periode Akuntansi Belum di Setup, Transaksi Tidak Bisa Dilanjutkan" & Chr(13) & "Lakukan Setup Periode pada Menu File -> Master Periode", vbExclamation
'      IsInPeriod = False
'      Exit Function
'    Else
'      dAwal = GetNull(db!Awal)
'      dAkhir = GetNull(db!akhir)
'    End If
'
'    If Not (dTgl >= dAwal And dTgl <= dAkhir) Then
'      MsgBox "Periode Transaksi Salah, Transaksi Tidak bisa dilanjutkan", vbExclamation
'      IsInPeriod = False
'    End If
    
    If Int(DateDiff("d", dTgl, Now)) >= Int(GetNull(aCfg(objData, msJumlahHariBlokir))) Then
      IsInPeriod = False
    End If
    
'    If dTgl <= aCfg(objData, msTglAudit) Then
'      MsgBox "Maaf, transaksi untuk tgl tersebut tidak bisa dikoreksi karena sudah melewati proses audit", vbExclamation
'      IsInPeriod = False
'    End If
    
  End If
End Function

Function IsPeriodClosed() As Boolean
'Dim DB As New ADODB.Recordset
'Dim obj As New codesuitelibrary.data
'  Set DB = obj.Browse(GetDSN, "Periode", , "Status", sisAssign, "1", , , , 0, 1)
'  IsPeriodClosed = DB.RecordCount > 0
  IsPeriodClosed = True
End Function

Function DecToRomawi(ByVal nValue As Integer) As String
Dim a As Double
Dim b As Double
Dim c As Double
Dim d As Double
Dim ca As String
Dim cb As String
Dim cc As String
Dim cd As String

  If nValue <= 0 Or nValue > 3000 Then
    MsgBox "Maaf, bilangan harus dalam rentang 1-3000", vbExclamation, "Error Function"
  Else
    a = Devide(nValue, 1000) * 1000
    b = (Devide(nValue, 100) Mod 10) * 100
    c = (Devide(nValue, 10) Mod 10) * 10
    d = (Devide(nValue, 1) Mod 10) * 1

    
    If a = 1000 Then
      ca = "M"
    ElseIf a = 2000 Then
      ca = "MM"
    ElseIf a = 3000 Then
      ca = "MMM"
    End If
    
    If b = 100 Then
      cb = "C"
    ElseIf b = 200 Then
      cb = "CC"
    ElseIf b = 300 Then
      cb = "CCC"
    ElseIf b = 400 Then
      cb = "CD"
    ElseIf b = 500 Then
      cb = "D"
    ElseIf b = 600 Then
      cb = "DC"
    ElseIf b = 700 Then
      cb = "DCC"
    ElseIf b = 800 Then
      cb = "DCCC"
    ElseIf b = 900 Then
      cb = "CM"
    End If
    
    If c = 10 Then
      cc = "X"
    ElseIf c = 20 Then
      cc = "XX"
    ElseIf c = 30 Then
      cc = "XXX"
    ElseIf c = 40 Then
      cc = "XL"
    ElseIf c = 50 Then
      cc = "L"
    ElseIf c = 60 Then
      cc = "LX"
    ElseIf c = 70 Then
      cc = "LXX"
    ElseIf c = 80 Then
      cc = "LXXX"
    ElseIf c = 90 Then
      cc = "XC"
    End If
    
    If d = 1 Then
      cd = "I"
    ElseIf d = 2 Then
      cd = "II"
    ElseIf d = 3 Then
      cd = "III"
    ElseIf d = 4 Then
      cd = "IV"
    ElseIf d = 5 Then
      cd = "V"
    ElseIf d = 6 Then
      cd = "VI"
    ElseIf d = 7 Then
      cd = "VII"
    ElseIf d = 8 Then
      cd = "VIII"
    ElseIf d = 9 Then
      cd = "IX"
    End If

    DecToRomawi = Trim(UCase(ca & cb & cc & cd))
  End If
End Function

Sub SetPage(vaArray As XArrayDB, ByVal nRowPerPage As Double)
Dim nMax As Double
Dim n As Double
Dim nSisa As Integer

  nMax = vaArray.UpperBound(1) + 1
  
  If nMax < nRowPerPage Then
      nSisa = nRowPerPage - nMax
      For n = 1 To nSisa
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        vaArray(vaArray.UpperBound(1), 0) = " "
      Next
  Else
    nMax = nMax Mod nRowPerPage
    If nMax > 0 Then
      For n = nMax + 1 To nRowPerPage
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        vaArray(vaArray.UpperBound(1), 0) = " "
      Next
    End If
  End If
End Sub

'Function isInGrid(ByVal vaFind As XArrayDB, ByVal nCol As Integer, ByVal cKodeBarang, Optional ByVal nNomorValidasi As Integer = 0) As Boolean
'Dim n As Single
'
'  isInGrid = False
'  'verify if the grid is not empty
'  If vaFind.UpperBound(1) >= 0 Then
'    n = vaFind.Find(0, nCol, cKodeBarang)
'    If n >= 0 And nNomorValidasi <> n + 1 Then
'      isInGrid = True
'    End If
'  End If
'End Function

Function isInGrid(ByVal vaFind As XArrayDB, ByVal nCol As Integer, ByVal cKodeBarang, Optional ByVal nNomorValidasi As Integer = 0, Optional ByRef nKe As Integer) As Boolean
Dim n As Single

  isInGrid = False
  nKe = 0
  'verify if the grid is not empty
  If vaFind.UpperBound(1) >= 0 Then
    n = vaFind.Find(0, nCol, cKodeBarang)
    If n >= 0 And nNomorValidasi <> n + 1 Then
      isInGrid = True
      nKe = n
    End If
  End If
End Function


Function Mod50(ByVal nJumlah As Double)
Dim cJumlah As String

  nJumlah = Round(nJumlah)
  cJumlah = Right(str(nJumlah), 2)
  Mod50 = nJumlah - Val(cJumlah)
  If Val(cJumlah) = 50 Or Val(cJumlah) = 0 Then
    Mod50 = nJumlah
  ElseIf Val(cJumlah) < 50 Then
    Mod50 = Mod50 + 50
  ElseIf Val(cJumlah) > 50 Then
    Mod50 = Mod50 + 100
  End If
End Function

Function GetValidNomorUrut(ByVal oNomor As Object, ByVal vaNomor As XArrayDB) As Boolean
GetValidNomorUrut = True
'Desc.
'===============================================================================================
': Function untuk melakukan validasi nomor urut ketika data dimasukkan ke dalam grid
'digunakan untuk menghindari kesalahan seperti "subsbrice out of range" dan error yang lain
'===============================================================================================
'by made hendra, made.hendra@gmail.com

  If oNomor.Value - 2 > vaNomor.UpperBound(1) Or oNomor.Value <= 0 Then
    oNomor.Value = vaNomor.UpperBound(1) + 2
  End If
  If vaNomor.UpperBound(1) < 0 Then
    GetValidNomorUrut = False
  End If
End Function

Function GetNamaBulan(ByVal nBulan As Integer) As String
  Select Case nBulan
    Case 1
      GetNamaBulan = "Januari"
    Case 2
      GetNamaBulan = "Februari"
    Case 3
      GetNamaBulan = "Maret"
    Case 4
      GetNamaBulan = "April"
    Case 5
      GetNamaBulan = "Mei"
    Case 6
      GetNamaBulan = "Juni"
    Case 7
      GetNamaBulan = "Juli"
    Case 8
      GetNamaBulan = "Agustus"
    Case 9
      GetNamaBulan = "September"
    Case 10
      GetNamaBulan = "Oktober"
    Case 11
      GetNamaBulan = "November"
    Case 12
      GetNamaBulan = "Desember"
    Case Else
      GetNamaBulan = ""
  End Select
End Function

Sub OpenDrawer(ByVal cPort As String)
'Desc.
'======================================================
'Function ini digunakan untuk membuka mesin cash drawer
'yang tertancap pada port COM1, edit code dibawah ini se
'perlunya jika ada perubahan yang dirasa perlu.
'Untuk membuka mesin cash drawer pada umumnya menggunakan
'string "CRTL-G" ATAU "0000000000" karakter 0 sebanyak
'10 kali. Jika muncul malfunction silahkan merujuk kembali
'pada referensi manual mesin cash drawer tersebut.
'======================================================
'by. made hendra, made.hendra@gmail.com

Dim n As Double
Dim nBruto As Double
Dim nTotQty As Double


'  Open cPort For Output As #1
'  Print #1, "CTRL-G"
'  Print #1, Chr(27) + Chr(112)
'  Print #1, "0000000000"
'  Close #1
  Shell "cmd Copy /b open.txt lpt1", vbNormalFocus
  
'  With aMainmenu.IO1
'    .Open cPort, ""
'
'    .WriteString "CTRL-G"
'    .Close
'
'    .Open cPort, ""
'    .WriteString "" + Chr(27) + Chr(112)
''     .WriteString "0000000000"
'    .Close
'
'    .Open cPort, ""
''    .WriteString "" + Chr(27) + Chr(112)
'     .WriteString "0000000000"
'    .Close
'
'
'  End With
''  MsgBox "Cash Drawer Open", vbInformation, "OPEN CASH DRAWER"
End Sub

Sub OpenDrawer2()
'Desc.
'======================================================
'Function ini digunakan untuk membuka mesin cash drawer
'yang tertancap pada port COM1, edit code dibawah ini se
'perlunya jika ada perubahan yang dirasa perlu.
'Untuk membuka mesin cash drawer pada umumnya menggunakan
'string "CRTL-G" ATAU "0000000000" karakter 0 sebanyak
'10 kali. Jika muncul malfunction silahkan merujuk kembali
'pada referensi manual mesin cash drawer tersebut.
'======================================================
'by. made hendra, made.hendra@gmail.com

Dim n As Double
Dim nBruto As Double
Dim nTotQty As Double

'  With aMainmenu.IO1
'    .Open "COM1:", ""
'    .WriteString "CTRL-G"
''     .WriteString "0000000000"
'    .Close
'  End With
End Sub

Sub OpenDrawer3()
'Desc.
'======================================================
'Function ini digunakan untuk membuka mesin cash drawer
'yang tertancap pada port COM1, edit code dibawah ini se
'perlunya jika ada perubahan yang dirasa perlu.
'Untuk membuka mesin cash drawer pada umumnya menggunakan
'string "CRTL-G" ATAU "0000000000" karakter 0 sebanyak
'10 kali. Jika muncul malfunction silahkan merujuk kembali
'pada referensi manual mesin cash drawer tersebut.
'======================================================
'by. made hendra, made.hendra@gmail.com

Dim n As Double
Dim nBruto As Double
Dim nTotQty As Double

'  With aMainmenu.IO1
'    .Open "COM1:", ""
'    .WriteString "" + Chr(27) + Chr(112)
''     .WriteString "0000000000"
'    .Close
'  End With
End Sub

Sub OpenNewDrawer(ByVal cKode As String)
'Desc.
'======================================================
'Function ini digunakan untuk membuka mesin cash drawer
'yang tertancap pada port COM1, edit code dibawah ini se
'perlunya jika ada perubahan yang dirasa perlu.
'Untuk membuka mesin cash drawer pada umumnya menggunakan
'string "CRTL-G" ATAU "0000000000" karakter 0 sebanyak
'10 kali. Jika muncul malfunction silahkan merujuk kembali
'pada referensi manual mesin cash drawer tersebut.
'======================================================
'by. made hendra, made.hendra@gmail.com

Dim n As Double
Dim nBruto As Double
Dim nTotQty As Double

'  With aMainmenu.IO1
'    .Open "COM1:", ""
'    .WriteString cKode
'    .Close
'
'    .Open "COM1:", ""
'    .WriteString "" + Chr(27) + Chr(112)
''     .WriteString "0000000000"
'    .Close
'
'    .Open "COM1:", ""
''    .WriteString "" + Chr(27) + Chr(112)
'     .WriteString "0000000000"
'    .Close
'  End With
'  MsgBox "Cash Drawer Open", vbInformation, "OPEN CASH DRAWER"
End Sub


Function Zip(ByVal cPassword As String, Optional lTile As Boolean = True) As String
Dim n As Single
Dim cRetval As String
Dim cZip As String
Dim cPad As String

  If lTile Then
    cPassword = Trim(cPassword)
    cPassword = IIf(cPassword = "", " ", cPassword)
    Do While Len(cPad) <= 20
      cPad = cPad & cPassword
    Loop
    cPassword = Left(cPad, 20)
  End If
  
  For n = 1 To Len(cPassword)
    cRetval = cRetval & Trim(str(Asc(Mid(cPassword, n, 1)) + (n * 59) + 1978))
  Next
  
  For n = 1 To Len(cRetval) Step 2
    cZip = cZip & Chr(Val(Mid(cRetval, n, 2)) + 65)
  Next
  Zip = cZip
End Function

Function GetAccountRegister(ByVal obj As CodeSuiteLibrary.Data, ByVal cUser) As String
Dim db As New adodb.Recordset

  Set db = obj.Browse(GetDSN, "Username", "UserName,FullName,KasTeller", "UserName", sisAssign, cUser)
  If Not db.EOF Then
    GetAccountRegister = GetNull(db!KasTeller)
  End If
End Function

Function GetSaldoKasBank(ByVal obj As CodeSuiteLibrary.Data) As Double
Dim str As String
Dim db As New adodb.Recordset

  GetSaldoKasBank = 0
  str = "select sum(debet-Kredit) as saldokas from saldokasbank where Bank = '" & GetAccountRegister(obj, GetRegistry(reg_Username)) & "' group by Bank"
  Set db = obj.SQL(GetDSN, str)
  If Not db.EOF Then
    GetSaldoKasBank = GetNull(db!saldokas)
  End If
End Function

Function lGetConfig() As Boolean
Dim vaArr As New XArrayDB
Dim vaArrErr As New XArrayDB
Dim db As New adodb.Recordset


  lGetConfig = True
  vaArr.ReDim 0, 19, 0, 0
  vaArrErr.ReDim 0, 19, 0, 0
    
  objData.Delete GetDSN, "config", Trim("keterangan"), sisAssign, "", " OR keterangan IS NULL"
  'Rek Akuntansi
  vaArr(0, 0) = lGetCekValiditasRekening(objData, aCfg(objData, msRekeningDiscountPembelian))
  vaArrErr(0, 0) = "Rekening Discount pembelian"
  vaArr(1, 0) = lGetCekValiditasRekening(objData, aCfg(objData, msRekeningPPnPembelian))
  vaArrErr(1, 0) = "Rekening PPn Pembelian"
  vaArr(2, 0) = lGetCekValiditasRekening(objData, aCfg(objData, msRekeningPenjualan))
  vaArrErr(2, 0) = "Rekening Penjualan"
  vaArr(3, 0) = lGetCekValiditasRekening(objData, aCfg(objData, msRekeningDiscountPenjualan))
  vaArrErr(3, 0) = "Rekening Discount Penjualan"
  vaArr(4, 0) = lGetCekValiditasRekening(objData, aCfg(objData, msRekeningPPnPenjualan))
  vaArrErr(4, 0) = "Rekening PPn Penjualan"
  vaArr(5, 0) = lGetCekValiditasRekening(objData, aCfg(objData, msRekeningCOGS))
  vaArrErr(5, 0) = "Rekening Harga Pokok"
  vaArr(6, 0) = lGetCekValiditasRekening(objData, aCfg(objData, msRekeningPotonganPiutang))
  vaArrErr(6, 0) = "Rekening Potongan Piutang"
  vaArr(7, 0) = lGetCekValiditasRekening(objData, aCfg(objData, msRekeningPotonganHutang))
  vaArrErr(7, 0) = "Rekening Potongan Hutang"
  vaArr(8, 0) = lGetCekValiditasRekening(objData, aCfg(objData, msPiutangDagang))
  vaArrErr(8, 0) = "Rekening Piutang Dagang"
  vaArr(9, 0) = lGetCekValiditasRekening(objData, aCfg(objData, msHutangDagang))
  vaArrErr(9, 0) = "Rekening Hutang Dagang"
  
  'Rek laba
  vaArr(10, 0) = lGetCekValiditasRekening(objData, aCfg(objData, msRekeningLaba))
  vaArrErr(10, 0) = "Rekening Laba"
  
  'Rek persediaan
  vaArr(11, 0) = lGetCekValiditasRekening(objData, aCfg(objData, msRekeningPersediaan))
  vaArrErr(11, 0) = "Rekening Persediaan"
  vaArr(12, 0) = lGetCekValiditasRekening(objData, aCfg(objData, msRekeningPenyesuian))
  vaArrErr(12, 0) = "Rekening Penyesuaian +"
  
  'Cost Center dan Gudang
  'vaArr(13, 0) = lGetCekCostCenter(objData, aCfg(objData, msCostCenterJualBeli))
  vaArrErr(13, 0) = "Cost Center Jual Beli"
  vaArr(14, 0) = lGetCekGudang(objData, aCfg(objData, msGudangPembelian))
  vaArrErr(14, 0) = "Gudang Pembelian"
  vaArr(15, 0) = lGetCekGudang(objData, aCfg(objData, msGudangPenjualan))
  vaArrErr(15, 0) = "Gudang Penjualan"
  
  'rek penyesuaian
  vaArr(16, 0) = lGetCekValiditasRekening(objData, aCfg(objData, msRekeningPenyesuaianKurang))
  vaArrErr(16, 0) = "Rekening Penyesuaian -"
  
  'rek komisi sales
  vaArr(17, 0) = lGetCekValiditasRekening(objData, aCfg(objData, msRekeningHutangSalesman))
  vaArrErr(17, 0) = "Rekening Hutang Salesman"
  vaArr(18, 0) = lGetCekValiditasRekening(objData, aCfg(objData, msRekeningBiayaKomisi))
  vaArrErr(18, 0) = "Rekening Biaya Komisi Penjualan"
  
  vaArr(19, 0) = lGetCekValiditasRekening(objData, aCfg(objData, msRekeningBG))
  vaArrErr(19, 0) = "Rekening BG/Cek"
  
  
  Dim n As Integer
  Dim a As Integer
  Dim cText As String
  
  
  a = vaArr.Find(0, 0, False)
  If a >= 0 Then
    lGetConfig = False
    cText = "Konfigurasi sistem Berikut dibawah ini belum sempurna: " & vbCrLf & vbCrLf
    For n = vaArrErr.LowerBound(1) To vaArrErr.UpperBound(1)
      If vaArr(n, 0) = 0 Then
        cText = cText & vaArrErr(n, 0) & vbCrLf
      End If
    Next n
    MsgBox cText
    Exit Function
  End If
  
End Function

Function lGetTransaction(obj As CodeSuiteLibrary.Data) As Boolean
Dim db As New adodb.Recordset
Dim cStatus As String
Dim cTatus As String

  cStatus = ""
  lGetTransaction = True
  
'  Set db = objData.Browse(GetDSN, "bukubesar", "sum(debet) as debet,sum(kredit) as kredit")
'  If Not db.EOF Then
'    If GetNull(db!debet) <> GetNull(db!kredit) Then
''      cStatus = "- Neraca Tidak Balance " & vbCrLf & "Jika anda mendapatkan pesan tersebut SEGERA hubungi pembuat program ini karena sistem sedang mengalami masalah yg serius (made.hendra@gmail.com, 083 119933 828)" & vbCrLf
'      cStatus = "NERACA TIDAK BALANCE" & vbCrLf
'      lGetTransaction = False
'    End If
'  End If
  
  Set db = objData.Browse(GetDSN, "bukubesar", "kodeakun", "kodeakun", sisAssign, "", " or kodeakun is null")
  If Not db.EOF Then
    cStatus = Trim(cStatus) & "- Field kodeakun pada bukubesar tidak valid (Emty or IS Null)"
    lGetTransaction = False
  End If

  Set db = objData.Browse(GetDSN, "bukubesar", "faktur", "faktur", sisAssign, "", " or faktur is null")
  If Not db.EOF Then
    cStatus = Trim(cStatus) & "- Field faktur pada bukubesar tidak valid (Emty or IS Null)"
    lGetTransaction = False
  End If
  
  If lGetTransaction = False Then
    MsgBox cStatus
  End If
  
End Function

Function lGetTableTransaction(obj As CodeSuiteLibrary.Data, ByVal NamaTabel As String, ByVal cField As String, ByRef cStat As String) As Boolean
Dim dba As adodb.Recordset

  lGetTableTransaction = True
  Set dba = objData.Browse(GetDSN, NamaTabel, cField, cField, sisAssign, "", " or " & cField & " is null")
  If Not dba.EOF Then
    lGetTableTransaction = False
    cStat = "- di table " & NamaTabel & " field " & cField & " has empty rows or is null"
  End If
End Function

Private Function GetModulConfig(ByVal obj As CodeSuiteLibrary.Data, cConfig As SisCfg) As String
Dim db As New adodb.Recordset

  GetModulConfig = ""
  Set db = obj.Browse(GetDSN, "config", , "jenis", sisAssign, cConfig)
  If Not db.EOF Then
    GetModulConfig = GetNull(db!modul) & " > " & GetNull(db!Label)
  End If
End Function

Function lGetCekAkunKas(ByVal obj As CodeSuiteLibrary.Data, ByVal cUser As String) As Boolean
Dim db As New adodb.Recordset
lGetCekAkunKas = False

  Set db = obj.Browse(GetDSN, "akunkas", "kodeakun,username", "username", sisAssign, cUser)
  If Not db.EOF Then
    Set db = obj.Browse(GetDSN, "akun", , "kodeakun", sisAssign, GetNull(db!kodeakun), " and jenis = 'D'")
    If Not db.EOF Then
      lGetCekAkunKas = True
    End If
  End If
End Function

Function lGetCekGudang(ByVal obj As CodeSuiteLibrary.Data, ByVal Gudang As String) As Boolean
Dim db As New adodb.Recordset
lGetCekGudang = False

  Set db = obj.Browse(GetDSN, "gudang", "kodegudang", "kodegudang", sisAssign, Gudang)
  If Not db.EOF Then
    lGetCekGudang = True
  End If
End Function

Function lGetCekCostCenter(ByVal obj As CodeSuiteLibrary.Data, ByVal CostCenter As String) As Boolean
Dim db As New adodb.Recordset
lGetCekCostCenter = False

  Set db = obj.Browse(GetDSN, "costcenter", "kodecostcenter", "kodecostcenter", sisAssign, CostCenter)
  If Not db.EOF Then
    lGetCekCostCenter = True
  End If
End Function

Function lGetCekValiditasRekening(ByVal obj As CodeSuiteLibrary.Data, ByVal Rek As String) As Boolean
Dim db As New adodb.Recordset
lGetCekValiditasRekening = False

  Set db = obj.Browse(GetDSN, "akun", "kodeakun,jenis", "kodeakun", sisAssign, Rek)
  If Not db.EOF Then
    If GetNull(db!jenis) = "D" Then
      lGetCekValiditasRekening = True
      Exit Function
    Else
      lGetCekValiditasRekening = False
      Exit Function
    End If
  End If
End Function

Function GetKonversiMember(ByVal obj As CodeSuiteLibrary.Data, _
ByVal kodeanggota As String, _
ByVal kodedep As String, _
ByVal nama As String, _
ByVal alamat As String, _
ByVal plafond As String, _
ByVal Status As String, _
ByVal tgl As String, _
ByVal telp As String, _
ByVal nopeg As String) As Boolean

Dim vaField As String
Dim db As New adodb.Recordset
Dim vaArr

  vaArr = Array("kodeanggota", "kodedep", "nama", "alamat", "plafond", "status", "tgl", "telp", "nopeg")
  GetKonversiMember = obj.Add(GetDSN, "anggota", vaArr, Array(kodeanggota, kodedep, nama, alamat, plafond, Status, tgl, telp, nopeg))
End Function

Function GetKonversiSupplier(ByVal obj As CodeSuiteLibrary.Data, _
ByVal kodesupplier As String, _
ByVal kodeakun As String, _
ByVal nama As String, _
ByVal alamat As String, _
ByVal telepon As String, _
ByVal fax As String, _
ByVal kota As String) As Boolean

Dim vaField As String
Dim db As New adodb.Recordset
Dim vaArr

  vaArr = Array("kodesupplier", "kodeakun", "nama", "alamat", "telepon", "fax", "kota")
  GetKonversiSupplier = obj.Add(GetDSN, "supplier", vaArr, Array(kodesupplier, kodeakun, nama, alamat, telepon, fax, kota))
End Function

Function GetKonversiGolonganInventory(ByVal obj As CodeSuiteLibrary.Data, _
ByVal kodegolongan As String, _
ByVal keterangan As String) As Boolean

Dim vaField As String
Dim db As New adodb.Recordset
Dim vaArr

  vaArr = Array("kodegolongan", "keterangan")
  GetKonversiGolonganInventory = obj.Add(GetDSN, "golongan", vaArr, Array(kodegolongan, keterangan))
End Function

Function GetKonversiSatuan(ByVal obj As CodeSuiteLibrary.Data, _
ByVal kodesatuan As String, _
ByVal keterangan As String) As Boolean

Dim vaField As String
Dim db As New adodb.Recordset
Dim vaArr

  vaArr = Array("kodesatuan", "keterangan")
  GetKonversiSatuan = obj.Add(GetDSN, "satuan", vaArr, Array(kodesatuan, keterangan))
End Function


Function GetKonversiStock(ByVal obj As CodeSuiteLibrary.Data, _
ByVal kodesatuan As String, _
ByVal nama As String, _
ByVal hargabeli As String, _
ByVal HargaJual As String, _
ByVal jenis As String, _
ByVal kodegolongan As String, _
ByVal cogs As String, _
ByVal kodebarcode As String) As Boolean

Dim vaField As String
Dim db As New adodb.Recordset
Dim vaArr

  vaArr = Array("kodesatuan", "kodegolongan", "barcode", "nama", "hargabeli", "hargajual", "cogs", "jenis")
  GetKonversiStock = obj.Add(GetDSN, "stock", vaArr, Array(kodesatuan, kodegolongan, kodebarcode, nama, hargabeli, HargaJual, cogs, jenis))
End Function

Function lExist(ByVal obj As CodeSuiteLibrary.Data, ByVal cTable As String, ByVal cKolom As String, ByVal cKey As String, Optional ByVal cWhere As String) As Boolean
Dim db As New adodb.Recordset
lExist = False
  
  Set db = obj.Browse(GetDSN, cTable, , cKolom, sisAssign, cKey, cWhere)
  If Not db.EOF Then
    lExist = True
    Exit Function
  End If
End Function

Function GetPoinHadiahMember(ByVal obj As CodeSuiteLibrary.Data, ByVal cMember As String, ByVal dTglX As Date) As Integer
Dim db As New adodb.Recordset
Dim cSQL As String

  GetPoinHadiahMember = 0
  cSQL = "SELECT sum(poinhadiah) as qty FROM poinhadiah WHERE kodeanggota = '" & cMember & "'" & _
  " AND exdate >='" & Format(dTglX, "yyyy-MM-dd") & "'" & _
  " AND status = '1'"
  
  'status 1 artinya poin masih valid/belum ditukar
  
  Set db = obj.SQL(GetDSN, cSQL)
  If Not db.EOF Then
    Do While Not db.EOF
      GetPoinHadiahMember = GetNull(db!qty)
      db.MoveNext
    Loop
  End If
End Function

