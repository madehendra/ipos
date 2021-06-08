Attribute VB_Name = "Trigger"
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data

Public Enum SisKartuStock
  ' Jika Debet Maka Antara 10 - 59
  ' Jika Kredit Maka Antara 60 - 99
  
  ' Debet
  Pembelian = 10            ' D
  ReturPenjualan = 11       ' D
  MutasiKe = 12           ' D
  PenyesuaianStock = 13
  SaldoAwal = 14
  PackingHasil = 15
  Rekanan = 16
  PelunasanRental = 17
  PelunasanRentalPaket = 18
  BuyBack = 19
  
  ' Kredit
  Penjualan = 60            ' K
  ReturPembelian = 61       ' K
  MutasiDari = 62
  PenjualanKasir = 63
  PackingBahan = 64
  Rental = 65
  RentalPaket = 65
  CheckList = 66
  PelunasanRekanan = 67
  Refund = 68
  ReturKonsinyasi = 69
  Komplimen = 70
End Enum

Public Enum SisKartuHutang
  ' HUTANG
  ' Jika Hutang maka antara 10 - 49
  ' Hutang Debet antara 10 - 29
  Sispembelian = 10
  SisSupplierBalance = 11
  
  
  '  Hutang Kredit Antara 30 - 49
  SisReturPembelian = 30
  SisPelunasanHutang = 31
  SisDiscountPelunasanHutang = 32
  
  
  SisMinimumHutang = 10
  SisMaximumHutang = 49
  
  ' PIUTANG
  ' Jika Piutang maka antara 50 - 99
  ' Piutang Debet Antara 50 - 69
  SisPenjualan = 50
  SisMemberBalance = 51
  SisUpdatePiutangDebet = 52
  
  ' Piutang Kredit Antara 70 - 99
  SisReturPenjualan = 70
  SisPelunasanPiutang = 71
  SisDiscountPelunasanPiutang = 72
  SisMemberOrder = 73
  SisUpdatePiutangKredit = 74
  SisPelunasanPiutangSederhana = 75
  
End Enum

Public Enum SisSaldoKasBank
  ' < 50 Debet
  SisSaldo_Penjualan = 10
  SisSaldo_ReturPembelian = 11
  SisSaldo_HutangLain = 12
  SisSaldo_MutasiKe = 13
  SisSaldo_PelunasanPiutang = 14
  SisSaldo_PendapatanLain = 15
  SisSaldo_PenjualanKasir = 16
  SisSaldo_Awal = 17
  SisSaldo_Rental = 18
  SisSaldo_PelunasanRental = 19
  SisSaldo_MusicStudio = 20
  
  ' >= 50 Kredit
  SisSaldo_Pembelian = 50
  SisSaldo_ReturPenjualan = 51
  SisSaldo_PiutangLain = 52
  SisSaldo_MutasiDari = 53
  SisSaldo_PelunasanHutang = 54
  SisSaldo_BiayaLain = 55
  SisSaldo_Rekanan = 56
  SisSaldo_Ambil = 57
  SisSaldo_Pasang = 58
  SisSaldo_PelunasanRekanan = 59
End Enum

Sub UpdSaldoKas(ByVal obj As CodeSuiteLibrary.Data, ByVal par As SisSaldoKasBank, _
                ByVal cFaktur As String, ByVal dTgl As Date, ByVal cBank As String, _
                ByVal cKeterangan As String, ByVal nJumlah As Double, _
                Optional ByVal lDeleteWhenExist As Boolean = True, Optional ByVal Transfer = False)
Dim vaField
Dim vaValue
Dim nDebet As Double
Dim nKredit As Double

  If lDeleteWhenExist Then
    obj.Delete GetDSN, "SaldoKasBank", "Status", sisAssign, par, " And Faktur = '" & cFaktur & "'"
  End If
  
  If nJumlah <> 0 Then
    If par < 50 Then
      nDebet = nJumlah
      nKredit = 0
    Else
      nDebet = 0
      nKredit = nJumlah
    End If
    
    vaField = Array("Status", "Faktur", "Tgl", "Bank", _
                    "Debet", "Kredit", "Keterangan", _
                    "DateTime", "UserName")
    vaValue = Array(par, cFaktur, dTgl, IIf(Transfer = True, cBank, GetAccountRegister(obj, GetRegistry(reg_Username))), _
                    nDebet, nKredit, cKeterangan, _
                    SNow, GetRegistry(reg_Username))
    obj.Add GetDSN, "SaldoKasBank", vaField, vaValue
  End If
End Sub

Function UpdKartuStock(ByVal obj As CodeSuiteLibrary.Data, ByVal par As SisKartuStock, _
                  ByVal cFaktur As String, ByVal dTgl As Date, _
                  ByVal cKode As String, ByVal nQty As Double, _
                  ByVal nHarga As Double, ByVal nDisc As Double, _
                  ByVal cKeterangan As String, ByVal Gudang As String, _
                  ByVal nHP As Double, Optional ByVal lJenis As String = 1) As Boolean
                  
Dim vaField, vaValue
Dim nDebet As Double
Dim nKredit As Double
Dim cUser As String
Dim dbRec As New ADODB.Recordset

  UpdKartuStock = True
  
  cUser = GetRegistry(reg_Username)
  
  If nQty <> 0 Then
    If par < 50 And nQty > 0 Then
      nDebet = nQty
    Else
      nKredit = Abs(nQty)
    End If
    
    vaField = Array("status", "nomor", "tgl", "kodestock", _
                    "qty", "debet", "kredit", "harga", "keterangan", _
                    "datetime", "username", "kodegudang", "hp", "disc", "ljenis")

    vaValue = Array(par, cFaktur, dTgl, cKode, _
                    nQty, nDebet, nKredit, nHarga, cKeterangan, _
                    SNow, cUser, Gudang, nHP, nDisc, lJenis)
                    
    Set dbRec = obj.Browse(GetDSN, "stock", , "kodestock", sisAssign, cKode)
    
    If Not dbRec.EOF Then
'      If dbRec!jenis < 9 Then
'        UpdKartuStock = obj.Add(GetDSN, "kartustock", vaField, vaValue, True)
'      End If
      UpdKartuStock = obj.Add(GetDSN, "kartustock", vaField, vaValue, True)
    End If
  End If
End Function

Function DeleteKartuStock(ByVal obj As CodeSuiteLibrary.Data, ByVal par As SisKartuStock, ByVal cFaktur As String)
  obj.Delete GetDSN, "KartuStock", "Status", sisAssign, par, " and Faktur = '" & cFaktur & "'"
End Function

Sub DeleteTransaksi(ByVal obj As CodeSuiteLibrary.Data, ByVal par As SisKartuStock, ByVal cFaktur As String)
Dim cTotTableName As String
Dim cDetailTableName As String

  Select Case par
    Case Pembelian
      cTotTableName = "TotPembelian"
      cDetailTableName = "Pembelian"
    Case ReturPembelian
      cTotTableName = "TotRtnPembelian"
      cDetailTableName = "RtnPembelian"
    Case Penjualan
      cTotTableName = "TotPenjualan"
      cDetailTableName = "Penjualan"
    Case ReturPenjualan
      cTotTableName = "TotRtnPenjualan"
      cDetailTableName = "RtnPenjualan"
    Case Rekanan
      cTotTableName = "TotRekanan"
      cDetailTableName = "DetRekanan"
  End Select
  
  obj.Delete GetDSN, cTotTableName, "Faktur", sisAssign, cFaktur
  obj.Delete GetDSN, cDetailTableName, "Faktur", sisAssign, cFaktur
  obj.Delete GetDSN, "KartuStock", "Status", sisAssign, par, " and Faktur = '" & cFaktur & "'"
End Sub

Function UpdKartuHutang(ByVal obj As CodeSuiteLibrary.Data, ByVal par As SisKartuHutang, _
                        ByVal cFaktur As String, ByVal dTgl As Date, _
                        Optional ByVal cSupplier As String = "", _
                        Optional ByVal cKeterangan As String = "", Optional ByVal nHutang As Double = 0, Optional ByVal cNow As Variant = Null, _
                        Optional ByVal cUser As String = "", _
                        Optional ByVal lDel As Boolean = True, _
                        Optional ByVal cKodeGroupSales As String = "") As Boolean
Dim vaField
Dim vaValue
Dim nDebet As Double
Dim nKredit As Double
Dim cSC As String
Dim cTableName As String
Dim lSave As Boolean
lSave = True

  If Trim(cKodeGroupSales) = "" Then
    cKodeGroupSales = GetRegistry(reg_KodeGroupPenjualan)
  End If
  
  Select Case par
    Case Is <= 29       ' Hutang Debet
      nDebet = nHutang
    Case Is <= 49       ' Hutang Kredit
      nKredit = nHutang
    Case Is <= 69       ' Piutang Debet
      nDebet = nHutang
    Case Is <= 99       ' Piutang Kredit
      nKredit = nHutang
  End Select
  
  If par >= SisMinimumHutang And par <= SisMaximumHutang Then
    cTableName = "kartuhutang"
    cSC = "kodesupplier"
    vaField = Array("status", "nomorkartuhutang", "tgl", cSC, _
                "keterangan", "debet", "kredit", _
                "username", "datetime")
    If lDel = True Then
      lSave = IIf(lSave, obj.Delete(GetDSN, cTableName, "status", sisAssign, par, " and nomorkartuhutang = '" & cFaktur & "'"), False)
    End If
  Else
    cTableName = "kartupiutang"
    cSC = "kodeanggota"
    vaField = Array("status", "nomorkartupiutang", "tgl", cSC, _
                    "keterangan", "debet", "kredit", _
                    "username", "groupsales", "datetime")
    If lDel = True Then
      lSave = IIf(lSave, obj.Delete(GetDSN, cTableName, "status", sisAssign, par, " and nomorkartupiutang = '" & cFaktur & "'"), False)
    End If
  End If
  
  If nHutang <> 0 Then
    cNow = IIf(IsNull(cNow), SNow, cNow)
    cUser = IIf(Trim(cUser) = "", GetRegistry(reg_Username), cUser)
    If cTableName = "kartupiutang" Then
      vaValue = Array(par, cFaktur, dTgl, cSupplier, _
                      cKeterangan, nDebet, nKredit, _
                      GetRegistry(reg_Username), cKodeGroupSales, cNow)
    Else
      vaValue = Array(par, cFaktur, dTgl, cSupplier, _
                      cKeterangan, nDebet, nKredit, _
                      GetRegistry(reg_Username), cNow)
    End If
    lSave = IIf(lSave, obj.Add(GetDSN, cTableName, vaField, vaValue), False)
  End If
  
  UpdKartuHutang = lSave
End Function

'Function GetSaldoStock(ByVal Obj As CodeSuiteLibrary.Data, ByVal cGudang As String, ByVal cKode As String) As Double
'Dim nSaldo As Double
'
'  Set dbData = Obj.Browse(GetDSN, "stock", , "kodestock", sisAssign, cKode)
'  If dbData.RecordCount > 0 Then
'    cKode = GetNull(dbData!KodeStock)
'
'    If Trim(cGudang) <> "" Then
'      Set dbData = Obj.Browse(GetDSN, "kartustock", "sum(debet-kredit) as qty", , , , "kodegudang = '" & cGudang & "' and kodestock = '" & cKode & "'")
'    Else
'      Set dbData = Obj.Browse(GetDSN, "kartustock", "sum(debet-kredit) as qty", , , , "kodestock = '" & cKode & "'")
'    End If
'
'    If Not dbData.EOF Then
'      nSaldo = nSaldo + GetNull(dbData!qty)
'    End If
'  End If
'
'  GetSaldoStock = nSaldo
'End Function

Function GetSaldoStock(ByVal obj As CodeSuiteLibrary.Data, ByVal cGudang As String, ByVal cKode As String, Optional ByVal dPerTgl = 0) As Double
Dim nSaldo As Double
  
  nSaldo = 0
  Set dbData = obj.Browse(GetDSN, "stock", , "kodestock", sisAssign, cKode)
  If dbData.RecordCount > 0 Then
    If GetNull(dbData!jenis) = 1 Then
      cKode = GetNull(dbData!KodeStock)
      If Trim(cGudang) <> "" Then
        Set dbData = obj.Browse(GetDSN, "kartustock", "sum(debet-kredit) as qty", , , , "kodegudang = '" & cGudang & "' and kodestock = '" & cKode & "' and tgl <= '" & IIf(dPerTgl <> 0, Format(dPerTgl, "yyyy-MM-dd"), Format(Now, "yyyy-MM-dd")) & "' and ljenis = 1")
      Else
        Set dbData = obj.Browse(GetDSN, "kartustock", "sum(debet-kredit) as qty", , , , "kodestock = '" & cKode & "' and tgl <= '" & IIf(dPerTgl <> 0, Format(dPerTgl, "yyyy-MM-dd"), Format(Now, "yyyy-MM-dd")) & "' and ljenis = 1")
      End If
      
      If Not dbData.EOF Then
        nSaldo = nSaldo + GetNull(dbData!qty)
      End If
    End If
  End If
  GetSaldoStock = nSaldo
End Function

Sub GetCetakFakturpenjualan(ByVal obj As CodeSuiteLibrary.Data, ByVal Faktur As String, ByVal lPrint As Boolean)
Dim n As Integer
Dim cTerbilang As String
Dim cField As String
Dim vaJoin
Dim vaGrid As New XArrayDB
Dim cHead As String
  
  cField = "s.nomorpenjualan,s.kodestock,t.barcode,s.qty,s.harga,s.jumlah,s.kodesatuan,s.discount,t.nama as namabarang,t.keterangan as ketbarang"
  vaJoin = Array("LEFT JOIN stock t ON t.kodestock = s.kodestock")
  Set dbData = obj.Browse(GetDSN, "penjualan s", cField, "s.nomorpenjualan", sisAssign, Faktur, , "s.urutfaktur asc", vaJoin)
  If Not dbData.EOF Then
    n = 0
    vaGrid.ReDim 0, dbData.RecordCount - 1, 0, 7
    Do While Not dbData.EOF
       vaGrid(n, 0) = n + 1
       vaGrid(n, 1) = IIf(Trim(GetNull(dbData!ketbarang, "")) <> "", (dbData!ketbarang), (dbData!Namabarang))  '(dbData!Namabarang)
       vaGrid(n, 2) = (dbData!qty)
       vaGrid(n, 3) = (dbData!Harga)
       vaGrid(n, 4) = (dbData!kodesatuan)
       vaGrid(n, 5) = (dbData!Discount)
       vaGrid(n, 6) = (dbData!jumlah)
       vaGrid(n, 7) = (dbData!barcode)
       dbData.MoveNext
      n = n + 1
    Loop
    
    'AMBIL INFORMASI customer
    Set dbData = obj.Browse(GetDSN, "totpenjualan t", "c.kodeanggota,t.fakturasli,t.piutang,t.dp,c.nama, c.alamat, c.kodedep,c.telp,t.subtotal,t.total,t.ppn,t.pajak,t.tgl,t.discount,sa.nama as namasalesman, d.keterangan as namadep", "t.nomorpenjualan", sisAssign, Faktur, , , Array("LEFT JOIN anggota c ON c.kodeanggota = t.kodeanggota", "left join salesman sa on sa.kodesalesman = t.kodesalesman", "left join dep d on d.kodedep = c.kodedep"))
    cTerbilang = "# " & Dec2Text(GetNull(dbData!Total)) & "Rupiah #"
    Dim cTmpSales As String
    
    cTmpSales = GetNull(dbData!namasalesman, "")
    If dbData!Piutang <> 0 Then
      cHead = "INVOICE CREDIT"
    End If
    If dbData!Piutang = 0 Then
      cHead = "INVOICE CASH"
    End If
    
    With frmFaktur.RptFakturPenjualanEdit
      .Parameters("dTgl").ValueExpression = "'" & Format(GetNull(dbData!tgl), "dd-MM-yyyy") & "'"
      .Parameters("cSE").ValueExpression = "'" & Faktur & "'"
      
      .Parameters("cNama").ValueExpression = "'" & GetNull(dbData!nama, "") & "'"
      .Parameters("cAlamat").ValueExpression = "'" & GetNull(dbData!alamat, "") & "'"
      .Parameters("cKota").ValueExpression = "'" & GetNull(dbData!namadep, "") & " Telp " & GetNull(dbData!telp) & "'"
      .Parameters("cKodeAnggota").ValueExpression = "'" & GetNull(dbData!kodeanggota, "") & "'"
      
      .Parameters("cTerbilang").ValueExpression = "'" & cTerbilang & "'"
      .Parameters("cTTD").ValueExpression = "'" & Padc(GetRegistry(reg_FullName), 45) & "'"
      .Parameters("cReceived").ValueExpression = "'" & Padc("", 45) & "'"
      
      .Parameters("nSubtotal").ValueExpression = GetNull(dbData!Subtotal)
      '.Parameters("nPPn").ValueExpression = GetNull(dbData!ppn)
      '.Parameters("nPajak").ValueExpression = GetNull(dbData!PAJAK)
      '.Parameters("nDP").ValueExpression = GetNull(dbData!dp)
      .Parameters("nTotal").ValueExpression = GetNull(dbData!Total)
      '.Parameters("nDiscount").ValueExpression = GetNull(dbData!Discount)
      
      .Parameters("cNamaPerusahaan").ValueExpression = "'" & aCfg(obj, msNamaPerusahaan) & "'"
      .Parameters("cAlamatPerusahaan").ValueExpression = "'Alamat : " & aCfg(obj, msAlamatPerusahaan) & " Telp/Fax " & aCfg(objData, msTelepon) & "/" & aCfg(objData, msFax) & "'"
      .Parameters("cUserName").ValueExpression = "'" & GetRegistry(reg_FullName) & "'"
      .Parameters("cJudul").ValueExpression = "'" & cHead & "'"
      
      .Parameters("cSales").ValueExpression = "'" & cTmpSales & "'"
      .Parameters("cFooter").ValueExpression = "'" & aCfg(objData, msFooterPenjualanNonTunai) & "'"
      .Parameters("cFooter2").ValueExpression = "'" & aCfg(objData, msFooterPenjualanNonTunai2) & "'"
      .Parameters("cFooter3").ValueExpression = "'" & aCfg(objData, msFooterPenjualanNonTunai3) & "'"

      Set .Array = vaGrid
      .Refresh
      If MsgBox("Apakah cetakan mau dalam bentuk kertas A4?!!" & vbCrLf & "Jika tidak maka cetakan akan dalam bentuk 1/2 kertas kuarto", vbYesNo) = vbYes Then
        .Profiles(0).PrinterPaperSize = tdbPPS_A4
      End If
      
      frmFaktur.Visible = False
      If lPrint = False Then
        .PrintPreview
      Else
        .PrintData
      End If
    End With
    Unload frmFaktur
  End If
End Sub

'Sub GetCetakFakturMemberOrder(ByVal obj As CodeSuiteLibrary.Data, ByVal Faktur As String, ByVal lPrint As Boolean)
'Dim n As Integer
'Dim cTerbilang As String
'Dim cField As String
'Dim vaJoin
'Dim vaGrid As New XArrayDB
'Dim cHead As String
'
'  cField = "s.nomormemberorder,s.kodestock,t.barcode,s.qty,s.harga,s.jumlah,s.kodesatuan,s.discount,t.nama as namabarang"
'  vaJoin = Array("LEFT JOIN stock t ON t.kodestock = s.kodestock")
'  Set dbData = obj.Browse(GetDSN, "memberorder s", cField, "s.nomormemberorder", sisAssign, Faktur, , "s.nourut asc", vaJoin)
'  If Not dbData.EOF Then
'    n = 0
'    vaGrid.ReDim 0, dbData.RecordCount - 1, 0, 7
'    Do While Not dbData.EOF
'       vaGrid(n, 0) = n + 1
'       vaGrid(n, 1) = (dbData!Namabarang)
'       vaGrid(n, 2) = (dbData!qty)
'       vaGrid(n, 3) = (dbData!Harga)
'       vaGrid(n, 4) = (dbData!kodesatuan)
'       vaGrid(n, 5) = (dbData!Discount)
'       vaGrid(n, 6) = (dbData!jumlah)
'       vaGrid(n, 7) = (dbData!barcode)
'       dbData.MoveNext
'      n = n + 1
'    Loop
'
'    'AMBIL INFORMASI customer
'    Set dbData = obj.Browse(GetDSN, "totmemberorder t", "c.kodeanggota,c.telp,t.fakturasli,t.piutang,c.nama, c.alamat, c.kodedep,t.subtotal,t.total,t.ppn,t.pajak,t.tgl,t.discount,t.dp,sa.nama as namasalesman, d.keterangan as namadep", "t.nomormemberorder", sisAssign, Faktur, , , Array("LEFT JOIN anggota c ON c.kodeanggota = t.kodeanggota", "left join salesman sa on sa.kodesalesman = t.kodesalesman", "left join dep d on d.kodedep = c.kodedep"))
'    cTerbilang = "# " & Dec2Text(GetNull(dbData!Total)) & "Rupiah #"
'
'    Dim cTmpSales As String
'
'    cTmpSales = GetNull(dbData!namasalesman, "")
'    cHead = "MEMBER ORDER"
'
''    If dbData!Piutang <> 0 Then
''      cHead = "CREDIT"
''    End If
''
''    If dbData!Piutang = 0 Then
''      cHead = "CASH"
''    End If
'
'    With frmFaktur.RptFakturOrder
'      .Parameters("dTgl").ValueExpression = "'" & Format(GetNull(dbData!tgl), "dd-MM-yyyy") & "'"
'      .Parameters("cSE").ValueExpression = "'" & Faktur & "'"
'
'      .Parameters("cKodeAnggota").ValueExpression = "'" & GetNull(dbData!kodeanggota, "") & "'"
'      .Parameters("cNama").ValueExpression = "'" & GetNull(dbData!nama, "") & "'"
'      .Parameters("cAlamat").ValueExpression = "'" & GetNull(dbData!alamat, "") & " " & GetNull(dbData!namadep, "") & "'"
'      .Parameters("cKota").ValueExpression = "'Telp. " & GetNull(dbData!telp) & "'"
'
'
'      .Parameters("cTerbilang").ValueExpression = "'" & cTerbilang & "'"
'      .Parameters("cTTD").ValueExpression = "'" & Padc(GetRegistry(reg_FullName), 45) & "'"
'      .Parameters("cReceived").ValueExpression = "'" & Padc("", 45) & "'"
'
'      .Parameters("nSubtotal").ValueExpression = GetNull(dbData!Subtotal)
'      .Parameters("nPPn").ValueExpression = 0
'      .Parameters("nPajak").ValueExpression = 0
'      .Parameters("nTotal").ValueExpression = GetNull(dbData!Subtotal) - GetNull(dbData!dp)
'      .Parameters("nDiscount").ValueExpression = GetNull(dbData!dp)
'
'      .Parameters("cNamaPerusahaan").ValueExpression = "'" & aCfg(obj, msNamaPerusahaan) & "'"
'      .Parameters("cAlamatPerusahaan").ValueExpression = "'Alamat : " & aCfg(obj, msAlamatPerusahaan) & " Telp/Fax " & aCfg(objData, msTelepon) & "/" & aCfg(objData, msFax) & "'"
'      .Parameters("cUserName").ValueExpression = "'" & GetRegistry(reg_FullName) & "'"
'      .Parameters("cJudul").ValueExpression = "'" & cHead & "'"
'
'      .Parameters("cSales").ValueExpression = "'" & cTmpSales & "'"
'      .Parameters("cFooter").ValueExpression = "'" & aCfg(objData, msFooterPenjualanNonTunai) & "'"
'      .Parameters("cFooter2").ValueExpression = "'" & aCfg(objData, msFooterPenjualanNonTunai2) & "'"
'
'      Set .Array = vaGrid
'      .Refresh
'      If MsgBox("Apakah cetakan mau dalam bentuk kertas A4?!!" & vbCrLf & "Jika tidak maka cetakan akan dalam bentuk 1/2 kertas kuarto", vbYesNo) = vbYes Then
'        .Profiles(0).PrinterPaperSize = tdbPPS_A4
'      End If
'
'      frmFaktur.Visible = False
'      If lPrint = False Then
'        .PrintPreview
'      Else
'        .PrintData
'      End If
'    End With
'    Unload frmFaktur
'  End If
'End Sub
