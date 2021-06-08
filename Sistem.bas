Attribute VB_Name = "Sistem"
Option Explicit

Dim objData As New CodeSuiteLibrary.Data
Dim dbData As New ADODB.Recordset
Public GetDSN As New ADODB.Connection
Public GetID As String
Public GetGroupSalesPenjualan As String

Public Enum sisFlag
  Nul = 0
  Posting = 1
End Enum


Public Enum SisPos
  None = 0
  Add = 1
  Edit = 2
  Delete = 3
End Enum



Public Enum sisModulTransaksi
  kasir = 10
  pembatalankasir = 11
  Penjualan = 12
  pembelian = 13
  returPembelian = 14
  returpenjualan = 15
  pelunasanpiutang = 16
  pelunasanhutang = 17
  StockOpname = 18
  jurnalumum = 19
  SimpananHarian = 20
  simpananwajib = 21
  CicilanBarang = 22
  PelunasanCicilanBarang = 23
  Pinjaman = 24
  PelunasanPinjaman = 25
  MutasiStock = 26
  biaya = 27
  Konsinyasi = 28
  KonsinyasiPulsa = 29
  MutasiKasBank = 30
  Packing = 31
  MemberOrder = 32
  promo = 33
  UpdateKartuPiutang = 34
  MemberTopUp = 35
  Prive = 36
  BuyBack = 37
  PenukaranPoin = 38
  PelunasanPiutangSederhana = 39
  refund = 40
  returKonsinyasi = 41
  Kompliment = 42
End Enum

Public Enum SisCfg
  msKodeKas = 0
  msNama = 1 'Nama Pemilik Usaha
  
  msFakturPrefix = 5
  msTaglinePerusahaan = 6
  msNamaPerusahaan = 7
  msAlamatPerusahaan = 8
  msTelepon = 9
  msFax = 10
  msEmail = 11
  msKota = 12
  msProvinsi = 13
  msJumlahSimpananPokok = 14
  msJumlahSimpananWajib = 15
  msGudangPenyimpanan = 16
  msSaldoMinus = 17
  msDiskonEstimasi = 18
  msKunciRekeningSetoranKas = 19
  
  msKasir1 = 20
  msKasir2 = 21
  msKasir3 = 22
  msOptFakturAsliPembelian = 23
  msDefaultPembelian = 24
  msHapusTransaksiPenjualan = 25
  msCetakPembelian = 26
  msIjinkanHargaBeliDibawahHargajual = 27
  msFooterPenjualanNonTunai = 28
  msFooterPenjualanNonTunai2 = 29
  
  msStikerLeftMargin = 30
  msStikerRightMargin = 31
  msStikerTopMargin = 32
  msStikerBottomMargin = 33
  msStikerOrientation = 34
  msMinimumDeposit = 35
  msPoin = 36
  msKelipatan = 37
  msTerm = 38
  msBulanBlokir = 39

  msPerhitunganKomisi = 40
  msPersenKomisi = 41
  msModelInput = 42
  msOptAudit = 43
  msTglAudit = 44
  msRekeningPrive = 45
  msRekeningSetoranKas = 46
  msRekeningKasTopUp = 47
  msRekeningTopUp = 48
  msRekeningBiayaBarang = 49
  
  msRekeningLaba = 50
  msRekeningSimpananPokok = 51
  msRekeningSimpananWajib = 52
  msRekeningDiscountPembelian = 53
  msRekeningPPnPembelian = 54
  msRekeningPenjualan = 55
  msRekeningDiscountPenjualan = 56
  msRekeningPPnPenjualan = 57
  msRekeningCOGS = 58
  msRekeningPotonganPiutang = 59
  
  msRekeningPotonganHutang = 60
  msRekeningSimpananHarian = 61
  msRekeningPinjaman = 62
  msRekeningPendapatanAdmPinjaman = 63
  msRekeningPendapatanBungaPinjaman = 64
  msRekeningKonsinyasi = 65
  msRekeningPPnkonsinyasi = 66
  msRekeningDiscountkonsinyasi = 67
  msPiutangDagang = 68
  msHutangDagang = 69
  
  msDiscountItemPembelian = 70
  msHargaPenjualanNonTunai = 71
  msDiscountPenjualan = 72
  msCHKdiscountPenjualan = 73
  msOptUp = 74
  msOptKunciAkunKas = 75
  msOptKunciPeriodeAkuntansi = 76
  msNilaiDecimals = 77
  msRekeningHutangBiaya = 78
  msEnableDisableDiscountItemPembelian = 79
  
  msKolomHargaPenjualanNonTunai = 80
  msCetakanPenjualanNonTunai = 81
  msDefaultModelPenjualan = 82
  msRekeningPersediaan = 83
  msRekeningPenyesuian = 84
  msRekeningPenyesuaianKurang = 85
  msRekeningHutangSalesman = 87
  msRekeningBiayaKomisi = 88
  msRekeningBG = 89
  
  msKolomHargaKasir = 90
  msQtyKasir = 91
  msPortPrinter = 92
  msModelInputPembelian = 93
  msCostCenterSimpanPinjam = 94
  msGudangPembelian = 95
  msGudangPenjualan = 96
  msMarkUpHargaJual = 97
  msOtorisasi = 98
  msVersion = 99
  
  msEditTransaksiPenjualan = 100
  msRekeningReturPembelian = 101
  msFooterPenjualanNonTunai3 = 102
  msJumlahHariBlokir = 103
  msBisaEditPembelian = 104
  msOtorisasiPenuh = 105
  msRekeningTitipanKasReturPembelian = 106
  msRekeningFeeKartu = 107
  msMinKartu = 108
  msKunciKasirDelete = 109
  msKasir4 = 110
  msKasir5 = 111
End Enum


Public Enum SisRegistry
  reg_DSN = 0
  reg_UserLevel = 1
  reg_UserID = 2
  reg_Username = 3
  reg_FullName = 4
  reg_Wallpaper = 5
  reg_Copystruk = 6
  reg_TampilkanBarcode = 7
  reg_PrinterAktif = 8
  reg_IP = 9
  reg_ServerUID = 10
  reg_ServerPWD = 11
  reg_Database = 12
  reg_QtyDefaultKasir = 13
  reg_CetakanKasir = 14
  reg_SerialNumber = 15
  reg_PrinterThermal = 16
  reg_AlignmentThermal = 17
  reg_KodeAnggota = 18
  reg_KodeSalesman = 19
  reg_ChkTunaiPenjualan = 20
  reg_F1Key = 21
  reg_F2Key = 22
  reg_LebarKertas = 23
  reg_MarginKiri = 24
  reg_LebarKolom1 = 25
  reg_LebarKolom2 = 26
  reg_LebarKolom3 = 27
  reg_LebarEfektif = 28
  
  reg_LebarKolom1_2 = 29
  reg_LebarKolom2_2 = 30
  reg_LebarKolom3_2 = 31
  reg_CetakLabelCustomer = 32
  reg_MarginBawah = 33
  
  reg_CetakanPenjualanNonTunai = 34
  reg_CetakBerulang = 35
  reg_PortStruk = 36
  reg_KeySecret = 37
  reg_TampilNotifikasi = 38
  reg_OptGroupSales = 39
  reg_KodeGroupPenjualan = 40
  reg_OptModelPelunasanPiutang = 41
  reg_ModePelunasanPiutang = 42
  reg_KodeGroupSalesPembelian = 43
  reg_OpenCashDrawer = 44
  reg_LimitPencarian = 45
End Enum

Public Enum SisFormatType
  Sis_yyyy_MM_dd = 0
  Sis_dd_MM_yyyy = 1
  sis_BilRpPict2 = 2
  Sis_BilRpPict = 3
  Sis_yy_MM_dd = 4
  Sis_dd_MMMM_yyyy = 5
  Sis_dd_MMMM_yy = 6
End Enum


Public Function RumusDiscount(ByVal HargaJual As Double, ByVal Discount1 As Double, ByVal Discount2 As Double, ByVal MarkDown As Double) As Double
Dim nX1 As Double
Dim nX2 As Double

  nX1 = HargaJual - (HargaJual * Discount1 / 100)
  nX2 = nX1 - (nX1 * Discount2 / 100)
  RumusDiscount = nX2 - MarkDown
End Function

Public Function RumusMarkUp(ByVal HargaPokok As Double, ByVal Up1 As Double, ByVal Up2 As Double, ByVal MarkUp As Double) As Double
Dim nX1 As Double
Dim nX2 As Double

  nX1 = HargaPokok + (HargaPokok * Up1 / 100)
  nX2 = nX1 + (nX1 * Up2 / 100)
  RumusMarkUp = nX2 + MarkUp
End Function


Public Function GetRupiah(ByVal nHarga As Double, Optional ByVal nPembulatan As Double = 1) As Double
Dim n As Double
Dim nSelisih As Double

 GetRupiah = nHarga
 If nPembulatan > 0 Then
  n = nHarga \ nPembulatan
  n = n * nPembulatan
  nSelisih = nHarga - n
  If nSelisih > 0 Then
   GetRupiah = ((nHarga \ nPembulatan) * nPembulatan) + nPembulatan
  Else
   GetRupiah = nHarga
  End If
 End If
End Function

Public Sub GetNotifikasiAdd(ByVal cTitle As String, Optional ByVal cText As String = "Mohon Ditunggu..", Optional ByVal cFlag As bFlag = 0)
  TrayAdd aMainmenu.pbTray
  TrayBalloon aMainmenu.pbTray, cTitle, cText, cFlag
End Sub

Public Sub GetNotifikasiRemove()
  TrayRemove aMainmenu.pbTray
End Sub

Function GetNewExpiredApp() As Integer
Dim cSQL As String
Dim db As New ADODB.Recordset
Dim dTglkey As Date
Dim dTglTrs As Date
On Error GoTo Error:


  GetNewExpiredApp = 0
  dTglkey = CryptRC4(FromHexDump(GetNull(db!tokenapp)), GetRegistry(reg_KeySecret))
'  cSQL = " select * from keyapp where keyapp = '" & GetRegistry(reg_KeySecret) & "'"
'  cSQL = cSQL & " ORDER BY id DESC LIMIT 0,1"
'
'  Set db = objData.SQL(GetDSN, cSQL)
'  If Not db.EOF Then
'    'jika punya serial maka bandingkan dengan data
'    dTglkey = CryptRC4(FromHexDump(GetNull(db!tokenapp)), GetRegistry(reg_KeySecret))
'    cSQL = "select * from bukubesar order by tgl desc limit 0,1"
'    Set db = objData.SQL(GetDSN, cSQL)
'    If Not db.EOF Then
'      dTglTrs = GetNull(db!tgl)
'      If DateDiff("d", dTglTrs, dTglkey) > 0 Then
'        GetNewExpiredApp = DateDiff("d", dTglTrs, dTglkey)
'      End If
'    End If
'  End If
  
  GetNewExpiredApp = DateDiff("d", Date, dTglkey)  'CryptRC4(FromHexDump(GetNull(db!tokenapp)), GetRegistry(reg_KeySecret))

Exit Function
Error:
  End
End Function

Function GetHargaBarang(ByVal obj As CodeSuiteLibrary.Data, ByVal kodestk As String, ByVal nHargaBatas) As String

  GetHargaBarang = ""
  Set dbData = obj.Browse(GetDSN, "stock", "hargajual", "kodestock", sisAssign, kodestk)
  If Not dbData.EOF Then
    If GetNull(dbData!HargaJual) > nHargaBatas Then
      GetHargaBarang = "HELLOO. BARANG INI HARGA NYA " & Format(GetNull(dbData!HargaJual), "###,###,###,##00") & vbCrLf & "BENER MAU DISIMPAN.. SERIUS "
    End If
  End If
End Function


Sub GetInfoStockDong(ByVal obj As CodeSuiteLibrary.Data, ByVal KodeStock As String)
Dim db As New ADODB.Recordset
Dim dba As New ADODB.Recordset
Dim cInfo As String

  cInfo = ""
  
'  Set db = obj.Browse(GetDSN, "gudang", "kodegudang")
'  If Not db.EOF Then
'    Do While Not db.EOF
'      Set dba = obj.Browse(GetDSN, "kartustock", "sum(debet-kredit) as saldostock", "kodestock", sisAssign, KodeStock, " and kodegudang = '" & GetNull(db!kodegudang) & "'")
'      If Not dba.EOF Then
'        If GetNull(dba!saldostock) > 0 Then
'          cInfo = "Stock di Gudang " & GetNull(db!kodegudang) & " : " & Format(GetNull(dba!saldostock)) & "  " & vbCrLf & cInfo
'        End If
'      End If
'      db.MoveNext
'    Loop
'  End If
  
  Set dba = obj.Browse(GetDSN, "kartustock", "sum(debet-kredit) as saldostock", "kodestock", sisAssign, KodeStock)
  If Not dba.EOF Then
    
'    If GetNull(dba!saldostock) <> 0 Then
'      cInfo = "Saldo stock : " & Format(GetNull(dba!saldostock)) & "  " & vbCrLf & cInfo
'    End If

    cInfo = "Saldo stock : " & Format(GetNull(dba!saldostock)) & "  " & vbCrLf & cInfo
    'cInfo = cInfo & GetHargaBarang(obj, KodeStock, 500000)
    
  End If
  
  If Trim(cInfo) <> "" Then
    MsgBox cInfo
  End If
  
End Sub

Function GetInfoStockDong2(ByVal obj As CodeSuiteLibrary.Data, ByVal KodeStock As String, ByVal KodeGudangStock) As Double
Dim db As New ADODB.Recordset
Dim dba As New ADODB.Recordset
  
  GetInfoStockDong2 = 0
  
'  Set dba = obj.Browse(GetDSN, "kartustock", "sum(debet-kredit) as saldostock", "kodestock", sisAssign, KodeStock, " and kodegudang = '" & KodeGudangStock & "'")
  Set dba = obj.Browse(GetDSN, "kartustock", "sum(debet-kredit) as saldostock", "kodestock", sisAssign, KodeStock)
  If Not dba.EOF Then
    If GetNull(dba!saldostock) <> 0 Then
      GetInfoStockDong2 = GetNull(dba!saldostock)
    End If
  End If
  
End Function

Function GetKunciAkunKas(ByVal obj As CodeSuiteLibrary.Data) As Boolean
  GetKunciAkunKas = False
  If aCfg(objData, msOptKunciAkunKas) = "Y" Then
    GetKunciAkunKas = True
  End If
End Function


Function GetAvailable(ByVal Faktur As String, ByVal tabel As String, ByVal Kolom As String) As Boolean
Dim db As New ADODB.Recordset
GetAvailable = True

  Set db = objData.Browse(GetDSN, tabel, , Kolom, sisAssign, Faktur)
  If Not db.EOF Then
    GetAvailable = False
  End If
End Function

Function GetNomor(ByVal cTable, ByVal cKolom, ByVal cRandom, ByVal cModulTransaksi As sisModulTransaksi) As String
Dim db As New ADODB.Recordset
Dim nLenFaktur As Byte
Dim cKodeCabang As String
Dim cFormatDepanInvoice As String

'  Set db = objData.Browse(GetDSN, cTable, "max(" & cKolom & ") as id", cKolom, sisPrefix, cModulTransaksi & cRandom & Format(Date, "YYMM"))
'  If Not db.EOF Then
'    GetNomor = cModulTransaksi & cRandom & Format(Date, "YYMM") & Padl(Val(Right(GetNull(db!ID), 3)) + 1, 3, "0")
'  Else
'    GetNomor = cModulTransaksi & cRandom & Format(Date, "YYMM") & Padl(1, 3, "0")
'  End If

  'set format dan panjang faktur
  'nLenFaktur = 3
  
  nLenFaktur = 3
  cKodeCabang = aCfg(objData, msFakturPrefix)
  cFormatDepanInvoice = cKodeCabang & cModulTransaksi & Format(Date, "YYYYMMDD")
  
  
'  cKodeCabang = "ACK"
'  cFormatDepanInvoice = cKodeCabang & cModulTransaksi & Format(Date, "YYYYMM")
  
  Set db = objData.Browse(GetDSN, cTable, "max(" & cKolom & ") as id", cKolom, sisPrefix, cFormatDepanInvoice)
  If Not db.EOF Then
    If GetNull(db!ID) <> 0 Then
      'GetNomor = "GT-12201808111"
      GetNomor = cFormatDepanInvoice & Padl(Val(Right(GetNull(db!ID), Len(GetNull(db!ID)) - Len(cFormatDepanInvoice))) + 1, nLenFaktur, "0") 'Padl(Val(Right(GetNull(db!Id), Len(GetNull(db!Id)) - Len(cModulTransaksi & Format(Date, "YYYYMM")))) + 1, nLenFaktur, "0")     'Padl(Val(Right(GetNull(db!ID), 3)) + 1, 3, "0")
    Else
      GetNomor = cFormatDepanInvoice & Padl(1, nLenFaktur, "0") 'Padl(1, nLenFaktur, "0")
    End If
  End If
  
  If Len(GetNomor) > 20 Then
    MsgBox "Maaf, nomor faktur tidak bisa diciptakan otomatis oleh komputer" & vbCrLf & "Nomor sudah melebihi dari quota yg telah ditentukan" & vbCrLf & "Untuk mendapatkan support, silahkan menghubungi vendor software ini" & vbCrLf & "Maaf atas segala ketidak nyamanan ini"
    GetNomor = ""
  End If
  '
End Function

Function SisFormat(ByVal value, bFormat As SisFormatType, Optional ByVal cNegatifSeparator As String = "") As String
Dim vaFormat
  
  vaFormat = Array("yyyy-MM-dd", "dd-MM-yyyy", "###,###,###,###,###,##0.00", "###,###,###,###,###,###", _
                   "yy-MM-dd", "dd MMMM yyyy", "dd MMMM yy")
  If Len(cNegatifSeparator) > 0 Then
    If value >= 0 Then
      cNegatifSeparator = Space(Len(cNegatifSeparator))
    End If

    SisFormat = Left(cNegatifSeparator, 1) & Format(Abs(value), vaFormat(bFormat)) & Right(cNegatifSeparator, 1)
  Else
    SisFormat = Format(value, vaFormat(bFormat))
  End If
End Function

Function GetRegistry(par As SisRegistry, Optional cDefault = "")
  GetRegistry = GetSetting("madehendra", App.ProductName, "Reg" & par, "")
End Function

Function SaveRegistry(par As SisRegistry, cValue)
  SaveSetting "madehendra", App.ProductName, "Reg" & par, cValue
End Function

Function GetAkunKas(ByVal obj As CodeSuiteLibrary.Data, ByVal UserName As String) As String
Dim db As New ADODB.Recordset

  GetAkunKas = ""
  Set db = obj.Browse(GetDSN, "akunkas", , "username", sisAssign, UserName)
  If Not db.EOF Then
    GetAkunKas = GetNull(db!kodeakun)
  End If
End Function

Function GetCostCenterUser(ByVal obj As CodeSuiteLibrary.Data, ByVal UserName As String) As String
Dim db As New ADODB.Recordset

  GetCostCenterUser = ""
  Set db = obj.Browse(GetDSN, "akunkas", , "username", sisAssign, UserName)
  If Not db.EOF Then
    GetCostCenterUser = GetNull(db!kodecostcenter)
  End If
End Function

Function GetModePenjualanUser(ByVal obj As CodeSuiteLibrary.Data, ByVal UserName As String) As String
Dim db As New ADODB.Recordset

  GetModePenjualanUser = 0
  Set db = obj.Browse(GetDSN, "akunkas", , "username", sisAssign, UserName)
  If Not db.EOF Then
    GetModePenjualanUser = GetNull(db!modepenjualan)
  End If
End Function

Function GetGudangUser(ByVal obj As CodeSuiteLibrary.Data, ByVal UserName As String) As String
Dim db As New ADODB.Recordset

  GetGudangUser = ""
  Set db = obj.Browse(GetDSN, "akunkas", , "username", sisAssign, UserName)
  If Not db.EOF Then
    GetGudangUser = GetNull(db!kodegudang)
  End If
End Function

Function CreateNomorFaktur(ByVal obj As CodeSuiteLibrary.Data, ByVal cModul As sisModulTransaksi, ByVal cTabel As String, ByVal cKolom As String) As String
Dim db As New ADODB.Recordset
  'Buat nomor faktur berdasarkan tabel nomorfaktur
  CreateNomorFaktur = GetNomor("nomorfaktur", "nomorfaktur", 100, cModul)
  
  'Cek ditabel Total yg bersangkutan
  Set db = obj.SQL(GetDSN, "SELECT * from " & cTabel & " WHERE  " & cKolom & " = '" & CreateNomorFaktur & "'")
  If Not db.EOF Then
    'jika sudah pernah digunakan
'    MsgBox "Nomor faktur sudah pernah digunakan", vbCritical, "Error"
    Select Case cModul
      Case sisModulTransaksi.Penjualan
        CreateNomorFaktur = GetNomor("totpenjualan", "nomorpenjualan", 100, cModul)
        obj.Update GetDSN, "nomorfaktur", " modul = '" & cModul & "'", Array("nomorfaktur", "modul"), Array(CreateNomorFaktur, cModul)
      Case sisModulTransaksi.pembelian
        CreateNomorFaktur = GetNomor("totpembelian", "nomorpembelian", 100, cModul)
        obj.Update GetDSN, "nomorfaktur", " modul = '" & cModul & "'", Array("nomorfaktur", "modul"), Array(CreateNomorFaktur, cModul)
    End Select
    'CreateNomorFaktur = "" 'GetNomor(cTabel, cKolom, GetID, cModul)
  Else
    obj.Update GetDSN, "nomorfaktur", " modul = '" & cModul & "'", Array("nomorfaktur", "modul"), Array(CreateNomorFaktur, cModul)
  End If
  
End Function

Function RekSpace(cRekening, cKeterangan) As String
Dim n As Single, nDot As Single
  For n = 1 To Len(cRekening)
    If Mid(cRekening, n, 1) = "." Then
      nDot = nDot + 1
    End If
  Next
  If nDot >= 1 Then
    RekSpace = Space((nDot - 1) * 4) & cKeterangan
  End If
End Function

Sub InitCfg()
Dim cTipe As String
On Error GoTo salah

  Set dbData = objData.Browse(GetDSN, "config")
  If dbData.RecordCount > 0 Then
    Do While Not dbData.EOF
      cTipe = IIf(GetNull(dbData!Tipe) = "D", "[D]", "[C]")
      SaveSetting "madehendra", App.EXEName, "Cfg" & GetNull(dbData!jenis), cTipe & GetNull(dbData!keterangan)
      dbData.MoveNext
    Loop
  End If
  
salah:
If err.Number = 3704 Then
  MsgBox "Invalid database!!", vbExclamation
  End
End If

End Sub

Function aCfg(ByVal obj As CodeSuiteLibrary.Data, ByVal par As SisCfg, Optional cDefault As Variant = "") As Variant
Dim vRetval As Variant
Dim cTipe As String
Dim cValue As String



'  vRetval = GetSetting("madehendra", App.EXEName, "Cfg" & par, "[C]" & cDefault)
'  cTipe = left(vRetval, 3)
'  cValue = Mid(vRetval, 4)
'  Select Case cTipe
'    Case "[D]"
'      aCfg = DateSerial(left(cValue, 4), Mid(cValue, 5, 2), Mid(cValue, 7, 2))
'    Case Else
'      aCfg = cValue
'  End Select

  Set dbData = obj.Browse(GetDSN, "config", "jenis,keterangan,tipe", "jenis", sisAssign, par)
  If Not dbData.EOF Then
    Select Case GetNull(dbData!Tipe)
      Case "C"
        aCfg = GetNull(dbData!keterangan)
      Case "D"
        aCfg = DateSerial(Left(GetNull(dbData!keterangan), 4), Mid(GetNull(dbData!keterangan), 5, 2), Mid(GetNull(dbData!keterangan), 7, 2))
    End Select
  End If

End Function

Function UpdCfg(par As SisCfg, keterangan, Optional ByVal obj As CodeSuiteLibrary.Data, Optional ByVal lLabel As String = "", Optional ByVal cModul As String = "")
Dim cType As String

  cType = "C"
  If VarType(keterangan) = vbDate Then
    keterangan = Format(keterangan, "yyyymmdd")
    cType = "D"
  End If


  If Trim(lLabel) <> "" Then
    obj.Update GetDSN, "config", "jenis = '" & par & "'", Array("jenis", "keterangan", "tipe", "label", "modul"), Array(par, keterangan, cType, lLabel, cModul)
  Else
    obj.Update GetDSN, "config", "jenis = '" & par & "'", Array("jenis", "keterangan", "tipe"), Array(par, keterangan, cType)
  End If
  
End Function

Function CenterForm(bForm As Form, Optional ByVal lZeroTopLeft As Boolean = False, Optional ByVal nUpperMargin = 0)
  If lZeroTopLeft Then
    bForm.Left = 0
    bForm.Top = 0
  Else
'    bForm.Left = (Screen.Width / 2) - (bForm.Width / 2) - 100
'    bForm.Top = (Screen.Height / 2) - (bForm.Height / 2) - 750
    bForm.Left = (aMainmenu.Width / 2) - (bForm.Width / 2) - 129
    bForm.Top = (aMainmenu.Height / 2) - (bForm.Height / 2) - 750 + nUpperMargin
    
  End If
  bForm.Icon = aMainmenu.Icon
End Function

Function GetNull(value, Optional Default As Variant = 0)
  GetNull = IIf(IsNull(value), Default, value)
End Function

Function lCekStatusLunas(ByVal objData As CodeSuiteLibrary.Data, ByVal cNoPenjualan As String) As Boolean
Dim db As New ADODB.Recordset
Dim cSQL As String
Dim nPiutang As Double

  lCekStatusLunas = False
  'cek dulu apakah faktur ini faktur tunai atau bon?
  Set db = objData.Browse(GetDSN, "totpenjualan", , "nomorpenjualan", sisAssign, cNoPenjualan)
  If Not db.EOF Then
    nPiutang = GetNull(db!Piutang)
    If nPiutang <> 0 Then
      cSQL = "select sum(discount) as discount,sum(pelunasan) as pelunasan from pelunasanpiutang"
      cSQL = cSQL & " where nomorpenjualan = '" & cNoPenjualan & "'"
      Set db = objData.SQL(GetDSN, cSQL)
      If Not db.EOF Then
        If (nPiutang - GetNull(db!Discount)) <= GetNull(db!Pelunasan) Then
          lCekStatusLunas = True
          Exit Function
        End If
      End If
    End If
  End If
End Function

Function lCekStatusLunasHutang(ByVal objData As CodeSuiteLibrary.Data, ByVal cNoPembelian As String) As Boolean
Dim db As New ADODB.Recordset
Dim cSQL As String
Dim nPiutang As Double

  lCekStatusLunasHutang = False
  'cek dulu apakah faktur ini faktur tunai atau bon?
  Set db = objData.Browse(GetDSN, "totpembelian", , "nomorpembelian", sisAssign, cNoPembelian)
  If Not db.EOF Then
    nPiutang = GetNull(db!hutang)
    If nPiutang <> 0 Then
      cSQL = "select sum(discount) as discount,sum(pelunasan) as pelunasan from pelunasanhutang"
      cSQL = cSQL & " where nomorpembelian = '" & cNoPembelian & "'"
      Set db = objData.SQL(GetDSN, cSQL)
      If Not db.EOF Then
        If (nPiutang - GetNull(db!Discount)) <= GetNull(db!Pelunasan) Then
          lCekStatusLunasHutang = True
          Exit Function
        End If
      End If
    End If
  End If
End Function

Function GetFormLevel(cFormName, nLevel, Optional cmdAdd As Object, Optional cmdEdit As Object, Optional cmdDelete As Object) As String
Dim cStatus As String

  Set dbData = objData.Browse(GetDSN, "FormLevel", "Status", "Nama", sisAssign, cFormName, " and UserLevel = " & nLevel)
  If dbData.RecordCount > 0 Then
    cStatus = GetNull(dbData!Status)
  End If
  cStatus = IIf(GetRegistry(reg_UserLevel) = 0, "111", cStatus)
  GetFormLevel = cStatus

  On Error Resume Next
  cmdAdd.Enabled = Val(Left(cStatus, 1))
  cmdEdit.Enabled = Val(Mid(cStatus, 2, 1))
  cmdDelete.Enabled = Val(Mid(cStatus, 3, 1))
End Function

Function TabIndex(obj As Object, n As Single)
  obj.TabIndex = n
  n = n + 1
End Function

Function CheckDigit(ByVal cKode As String) As String
Dim n As Double
Dim x As Double
Dim i As Double
Dim a As Double
Dim nBarcode As Double
Dim nSum As Double

  a = 1
  cKode = Replace(cKode, ".", "")
  nBarcode = Val(cKode)
  nSum = 1
  Do While nSum > 0
    nSum = SumBarcode(i, a, nBarcode)
    x = x + nSum
    a = a * 10
  Loop
  x = 10 - Val(Right(Format(x, "################"), 1))
  CheckDigit = Right(Trim(str(x)), 1)
End Function

Private Function SumBarcode(i As Double, ByVal n As Double, ByVal nKode As Double) As Double
  i = i + 1
  If i = 1 Then
    n = Int(nKode / n) * 3
  Else
    i = 0
    n = Int(nKode / n)
  End If
  SumBarcode = n
End Function

Function SNow(Optional ByVal cNow As Variant = Null) As String
  cNow = IIf(IsNull(cNow), Now, cNow)
  SNow = Format(cNow, "yyyy-mm-dd HH:MM:SS")
End Function

Function GetTDBGrid(cName As String, TDBGrid As TDBGrid, Optional lDefault As Boolean = False)
Dim n As Integer
Dim cKeterangan As String
Dim i As Integer
Dim cWidth As String

  If lDefault Then
    cName = cName & "Default"
  End If
  
  Set dbData = objData.Browse(GetDSN, "TDBGRID", , "NAME", sisAssign, cName)
  If dbData.RecordCount > 0 Then
    cKeterangan = GetNull(dbData!keterangan)
    For n = 1 To Len(cKeterangan)
      If Mid(cKeterangan, n, 1) = ";" Then
        If i < TDBGrid.Columns.Count Then
          TDBGrid.Columns(i).Width = Val(cWidth)
        End If
        i = i + 1
        cWidth = ""
      Else
        cWidth = cWidth & Mid(cKeterangan, n, 1)
      End If
    Next
  End If
End Function

Function SaveTDBGrid(cName As String, TDBGrid1 As TDBGrid, Optional lDefault As Boolean = False)
Dim n As Integer
Dim cKeterangan As String
  If lDefault Then
    cName = cName & "Default"
  End If
  
  For n = 0 To TDBGrid1.Columns.Count - 1
    cKeterangan = cKeterangan & TDBGrid1.Columns(n).Width & ";"
  Next
  
  objData.Update GetDSN, "TDBGRID", "NAME = '" & cName & "'", Array("Name", "Keterangan"), Array(cName, cKeterangan)
End Function

Sub SetButton(cmdSimpan As Object, cmdKeluar As Object, cmdAdd As Object, _
              cmdEdit As Object, cmdHapus As Object, nPos, lPar As Boolean, _
              Optional cmdAktivasi As Object)
  On Error Resume Next
  cmdSimpan.Enabled = lPar
  
  cmdAdd.Enabled = Not lPar
  cmdEdit.Enabled = Not lPar
  cmdHapus.Enabled = Not lPar

  'cmdAktivasi.Visible = GetRegistry(reg_UserLevel) = 0
  If lPar Then
    Set cmdKeluar.Picture = aMainmenu.pcCancel.Picture
    cmdKeluar.Caption = "      &Cancel "
  Else
    Set cmdKeluar.Picture = aMainmenu.pcExit.Picture
    cmdKeluar.Caption = "      E&xit"
    
    Select Case nPos
      Case 1
        cmdAdd.SetFocus
      Case 2
        cmdEdit.SetFocus
      Case 3
        cmdHapus.SetFocus
    End Select
    nPos = 0
  End If
  
  If aCfg(objData, msOtorisasi) = "1" Then
    If Not lPar Then
      GetFormLevel cmdSimpan.Parent.Name, GetRegistry(reg_UserLevel), cmdAdd, cmdEdit, cmdHapus
    End If
  End If
End Sub
Function GetInduk(Optional cRekening As String = "", Optional lPad As Boolean = True)
Dim lStop As Boolean, nOldLen As Byte
  lStop = False
  nOldLen = Len(cRekening)
  Do While Not lStop
    cRekening = Trim(cRekening)
    If Right(cRekening, 1) = "." Then
      cRekening = Left(cRekening, Len(cRekening) - 1)
    Else
      lStop = True
    End If
  Loop
  If lPad Then
    GetInduk = Pad(cRekening, nOldLen)
  Else
    GetInduk = cRekening
  End If
End Function

Function GetDetail(cRekening, Optional ByVal cFormat As String = "999.999.99.999", _
                   Optional ByVal nLeft As Single = 3) As String
Dim cPict As String
  cPict = Replace(cFormat, "9", " ")
  cRekening = Mid(cRekening, nLeft)
  cRekening = Trim(cRekening)
  GetDetail = cRekening & Right(cPict, Len(cPict) - Len(cRekening))
End Function

Function GetOpt(opt) As String
Dim n As Single, i As Single, lChar As Boolean
  For n = 0 To opt.Count - 1
    If opt(n).value Then
      With opt(n)
        For i = 1 To Len(.Caption)
          If lChar Then
            GetOpt = UCase(Mid(.Caption, i, 1))
            Exit Function
          End If
          If Mid(.Caption, i, 1) = "&" Then
            lChar = True
          End If
        Next
      End With
    End If
  Next
End Function

Sub SetOpt(opt, cChar As String)
Dim n As Single, i As Single, lChar As Boolean
  opt(0).value = True
  For n = 0 To opt.Count - 1
    With opt(n)
      For i = 1 To Len(.Caption)
        If lChar Then
          lChar = False
          If UCase(Mid(.Caption, i, 1)) = UCase(cChar) Then
            opt(n).value = True
            Exit Sub
          End If
        End If
        
        If Mid(.Caption, i, 1) = "&" Then
          lChar = True
        End If
      Next
    End With
  Next
End Sub

' This sample Visual Basic function calls functions in the DLL for you
Function barcode(ByVal cBarcode As String)
Dim cRetval As String, n As Single
  cRetval = "l"
  For n = 1 To Len(cBarcode)
    cRetval = cRetval & Chr(IIf(n <= 4, Asc("A"), Asc("K")) + Val(Mid(cBarcode, n, 1)))
    If n = 4 Then
      cRetval = cRetval & "k"
    End If
  Next
  cRetval = cRetval & "l"
  barcode = cRetval
End Function

Function Min(a, b)
  Min = IIf(a < b, a, b)
End Function

Function GetAppDescription() As String
  GetAppDescription = App.ProductName & "." & App.Major & "." & App.Minor & "." & App.Revision
End Function

Sub GetIPNumber(ByRef cIPNumber As String, ByRef cDatabase As String, ByRef cDSN As String, ByRef cPort As String, ByRef cKey As String, Optional ByRef cModePelunasanPiutang As String = "")
Dim cFile As String
Dim n As Double
Dim cData As String

  cFile = App.Path & "\config.ini"
  If Dir(cFile) <> "" Then
    Open cFile For Input Shared As #1
    Do While Not EOF(1)
      Line Input #1, cData
      cData = Replace(cData, " ", "")
      
      cIPNumber = GetData(cData, "IP=", cIPNumber)
      cDatabase = GetData(cData, "DATABASE=", cDatabase)
      cPort = GetData(cData, "PORT=", cPort)
      cDSN = GetData(cData, "DSN=", cDSN)
      cKey = GetData(cData, "KEY=", cKey)
      cModePelunasanPiutang = GetData(cData, "LUNASPIUTANG=", cModePelunasanPiutang)
    Loop
    Close #1
  End If
  
  'simpan pada registry
  SaveRegistry reg_DSN, cDSN
  SaveRegistry reg_Database, cDatabase
  SaveRegistry reg_IP, cIPNumber
  SaveRegistry reg_ModePelunasanPiutang, cModePelunasanPiutang
  
  If Trim(cIPNumber) = "" Then
    cIPNumber = "LocalHost"
  End If
  If Trim(cDatabase) = "" Then
    cDatabase = "RENT"
  End If
  If Trim(cDSN) = "" Then
    cDSN = "RENT"
  End If
End Sub


Sub GetPrinterCMD(ByRef cShellPrn As String)
Dim cFile As String
Dim n As Double
Dim cData As String

  cFile = App.Path & "\config.ini"
  If Dir(cFile) <> "" Then
    Open cFile For Input Shared As #1
    Do While Not EOF(1)
      Line Input #1, cData
      cShellPrn = GetData(cData, "PRINTER = ", cShellPrn)
    Loop
    Close #1
  End If
End Sub

Sub GetMyODBCFile(ByRef cMYODBCFile As String)
Dim cFile As String
Dim n As Double
Dim cData As String

  cFile = App.Path & "\config.ini"
  If Dir(cFile) <> "" Then
    Open cFile For Input Shared As #1
    Do While Not EOF(1)
      Line Input #1, cData
      cMYODBCFile = GetData(cData, "MYODBC_FILE = ", cMYODBCFile)
    Loop
    Close #1
  End If
End Sub

Function GetData(ByVal cData As String, ByVal cKey As String, ByVal cDefault As String) As String
Dim n As Double
  cData = LCase(cData)
  cKey = LCase(cKey)
  GetData = cDefault
  n = InStr(1, cData, cKey)
  If n <> 0 Then
    cData = Replace(cData, cKey, "")
    GetData = cData
  End If
End Function

Sub GetRekapanKatalog(ByVal cFileNamePath, ByVal nDiskonNett As Double)
Dim cFile As String
Dim n As Double
Dim cData As String
Dim cKodeBarang As String
Dim nJumlah As Double
Dim nHarga As Double
Dim dba As New ADODB.Recordset
Dim buffer
Dim nTotalTmp As Double
Dim nTmp As Double
Dim vaOrderan As New XArrayDB
Dim vaNotFound As New XArrayDB
Dim ni As Single
Dim nInStock As Double
Dim cTxtKodeInGudang As String
      


On Error GoTo Ero

  vaOrderan.ReDim 0, -1, 0, 2
  vaNotFound.ReDim 0, -1, 0, 1
  cTxtKodeInGudang = ""
  cFile = cFileNamePath
  If Dir(cFile) <> "" Then
    Open cFile For Input Shared As #1
    Do While Not EOF(1)
      Line Input #1, cData
      cData = Replace(cData, " ", "")
      cData = Replace(cData, ")", "")
      
      n = InStr(1, cData, "(")
      If n <> 0 Then
        Dim i As Integer
        i = Len(cData)
        cKodeBarang = Left(cData, n - 1)
        nJumlah = Right(cData, i - n)
      Else
        cKodeBarang = cData
        nJumlah = 1
      End If
      
      vaOrderan.InsertRows vaOrderan.UpperBound(1) + 1
      ni = vaOrderan.UpperBound(1)
      vaOrderan(ni, 0) = UCase(cKodeBarang)
      vaOrderan(ni, 1) = nJumlah 'IIf(nJumlah <> 1, nJumlah, "")
      
      'cari harga dan kalikan dengan jumlah = simpan dalam array
      Set dba = objData.Browse(GetDSN, "stock", "barcode,hargabeli,kodestock", "barcode", sisAssign, cKodeBarang)
      If Not dba.EOF Then
        nTmp = GetNull(dba!hargabeli) * nJumlah
        nTotalTmp = nTotalTmp + nTmp
        
        'cari apakah barang ini ada di gudang stock
        
        nInStock = GetInfoStockDong2(objData, GetNull(dba!KodeStock), aCfg(objData, msGudangPenyimpanan))
        If nInStock > 0 Then
          vaOrderan(ni, 2) = nInStock
          cTxtKodeInGudang = cTxtKodeInGudang & GetNull(dba!barcode) & "(" & nInStock & ")" & vbCrLf
        End If
        
      Else
        buffer = buffer & cKodeBarang & vbCrLf
        vaNotFound.InsertRows vaNotFound.UpperBound(1) + 1
        Dim k As Single
        k = vaNotFound.UpperBound(1)
        vaNotFound(k, 0) = UCase(cKodeBarang)
        vaNotFound(k, 1) = nJumlah
        vaOrderan.DeleteRows ni
      End If
      
    Loop
    Close #1
    
    MsgBox "TPG : " & vbTab & vbTab & Format(nTotalTmp, "###,###,##0.00") & vbCrLf & "Net (TPG - " & aCfg(objData, msDiskonEstimasi) & "%) : " & vbTab & Format(nTotalTmp - (nTotalTmp * nDiskonNett / 100), "###,###,##0.00") & vbCrLf & vbCrLf & "Item Not Found : " & vbCrLf & buffer & vbCrLf & "Barang ini sudah ada di Gudang !!, Jangan Order Lagi" & vbCrLf & cTxtKodeInGudang
    
  End If
  
  Dim a As New exportExcel
  
  If MsgBox("Apakah akan di export ke excel?", vbYesNo + vbInformation) = vbYes Then
    a.RecordSource = vaOrderan
    a.ExportToExcel
  End If
    
  If vaNotFound.UpperBound(1) > -1 Then
    If MsgBox("Apakah kode yg salah akan di export juga?", vbYesNo + vbInformation) = vbYes Then
      aMainmenu.dlg.Filter = "Text File (*.txt)|*.txt"
      aMainmenu.dlg.ShowSave
      Open aMainmenu.dlg.FileName For Output As #1
      For k = 0 To vaNotFound.UpperBound(1)
        Print #1, vaNotFound(k, 0)
      Next k
      Close #1
      ShellEx aMainmenu.dlg.FileName
    End If
  End If
Ero:
End Sub

Function calculateAge(dateOfBird As Date, fromData As Date) As String
       Dim dateNow As Date
       Dim tgl As Date
       Dim tgl1 As Date

       Dim years As Long
       Dim months As Long
       Dim days As Long
       Dim weeks As Long

      Dim yearWord As String
      Dim monthWord As String
      Dim dayWord As String
      Dim weekWord As String

      dateNow = fromData
      tgl = dateOfBird

      ' menghitung tahun
      years = DateDiff("yyyy", tgl, dateNow)
     If Month(tgl) > Month(dateNow) Then
         years = years - 1
      ElseIf Month(tgl) = Month(dateNow) And Day(tgl) > Day(dateNow) Then
          years = years - 1
      ElseIf Month(tgl) = Month(dateNow) And Day(tgl) = Day(dateNow) Then
          GoTo finally ' jika bulan dan tanggal sama maka perhitungan selesai
      End If

      ' menghitung bulan
      tgl = DateAdd("yyyy", years, tgl)
      months = DateDiff("m", tgl, dateNow)
      If Day(tgl) > Day(dateNow) Then
          months = months - 1
      ElseIf Month(tgl) = Month(dateNow) And Day(tgl) >= Day(dateNow) Then
          months = months - 1
      End If

      tgl = DateAdd("m", months, tgl)

      ' menghitung hari
      days = DateDiff("d", tgl, dateNow)
      weeks = days \ 7

      If weeks >= 1 Then
        days = days - weeks * 7
      End If

finally:
      yearWord = IIf(years = 0, "", years & " Tahun ")
      monthWord = IIf(months = 0, "", months & " Bulan ")
      weekWord = IIf(weeks = 0, "", weeks & " Minggu ")
      dayWord = IIf(days = 0, "", days & " Hari ")
      calculateAge = yearWord & monthWord & weekWord & dayWord
      calculateAge = Trim(calculateAge)
  End Function
