Attribute VB_Name = "akun"

Public Enum SisTypeRekening
  SisAktiva = 1
  SisHutang = 2
  SisModal = 3
  SisPendapatan = 4
  SisBiaya = 5
  SisAdministratif = 6
End Enum

Public Enum BisaUpdateBukuBesar
  bbPembelian = 10
  bbRtnPembelian = 11
  bbPenjualan = 12
  bbRtnPenjualan = 13
  bbJurnalUmum = 14
  bbPelunasanPiutang = 15
  bbPelunasanHutang = 16
  bbFOC = 17
  bbKasir = 18
  bbTabunganGuide = 19
  bbKasirKartu = 20
  bbOrderMenu = 21
  bbTravelAgent = 22
  bbPenyesuaian = 23
  bbBiayaTransport = 24
  bbInventaris = 25
End Enum

Public Enum vbTrigger
  msSimpananPokok = 1
  msPembelian = 2
  msReturPembelian = 3
  msPenjualan = 4
  msReturPenjualan = 5
  msPelunasanPiutang = 6
  msPelunasanHutang = 7
  msJurnalUmum = 8
  msSimpananHarian = 9
  msSimpananWajib = 10
  msPelunasanCicilanBarang = 11
  msCicilanBarang = 12
  msPinjaman = 13
  msPelunasanPinjaman = 14
  msBiaya = 15
  msKonsinyasi = 15
  msPenyesuaian = 16
  msMemberBalance = 17
  msSupplierBalance = 18
  msPenjualanKasir = 19
  msMutasiKasBank = 20
  msPencairanBG = 21
  msPacking = 22
  msMemberOrder = 23
  msMemberTopUp = 24
  msPrive = 25
  msBuyBack = 26
  msSaldoAwalStock = 27
  msPelunasanPiutangSederhana = 28
  msRefund = 29
  msKomplimen = 30
End Enum

Function UpdHargaPokok_LAMA_ERR(ByVal obj As CodeSuiteLibrary.Data, ByVal Stock As String, ByVal nQtyIn As Double, ByVal nHargaBeliNet As Double, Optional ByVal lMethod As Integer = 2, Optional ByVal lUpdate As Boolean = True, Optional ByVal nQtyOut As Integer = 0) As Boolean
Dim db As New ADODB.Recordset
Dim HargaPokok As Double
Dim cSQL As String
Dim nPersediaanSebelumnya As Double
Dim nPersediaanMasuk As Double

  'cek apakah saldo stock nya 0
  'jika 0 maka gunakan method 1
  UpdHargaPokok = True
  If GetSaldoStock(obj, "", Stock) = 0 Then
    lMethod = 1
  End If
  If lUpdate = True Then
    Select Case lMethod
      Case 1
        'rumus ini masih relevan jangan dihapus'
        '* REMARK SAJA *'
        cSQL = "select CEILING((sum(debet*hp)-sum(kredit*hp))/sum(debet-kredit) ) as hpp from kartustock where kodestock = '" & Stock & "'"
        Set db = obj.SQL(GetDSN, cSQL)
        If Not db.EOF Then
          If GetNull(db!hpp) <> 0 Then
              UpdHargaPokok = obj.Edit(GetDSN, "stock", "kodestock = '" & Stock & "'", Array("cogs"), Array(GetNull(db!hpp)))
          End If
        End If
        
      Case 2
        'Rumus
        '((saldo stock prev * harga pokok Prev)+(QtyPembelian*hargabelineto))/(QtyPembelian+SaldoStockPrev)
        '
        cSQL = "select * from stock where kodestock = '" & Stock & "'"
        Set db = obj.SQL(GetDSN, cSQL)
        If Not db.EOF Then
          nPersediaanSebelumnya = (GetSaldoStock(obj, "", Stock)) * IIf(GetNull(db!cogs) = 0 And GetNull(db!hargabeli) <> 0, GetNull(db!hargabeli), GetNull(db!cogs))
          nPersediaanMasuk = nQtyIn * nHargaBeliNet
          HargaPokok = (nPersediaanSebelumnya + nPersediaanMasuk) / ((GetSaldoStock(obj, "", Stock) + nQtyIn))
          UpdHargaPokok = obj.Edit(GetDSN, "stock", "kodestock = '" & Stock & "'", Array("cogs"), Array(HargaPokok))
        End If
    End Select
  End If
End Function

Function NewUpdHargaPokok(ByVal obj As CodeSuiteLibrary.Data, cKodeStk As String) As Double
Dim cSQL
Dim db As New ADODB.Recordset
Dim n As Double
  
  cSQL = ""
  cSQL = "SELECT  s.kodegolongan,g.keterangan as namagolongan,s.barcode,s.nama,s.kodesatuan,sum(ks.debet - ks.kredit) as akhir , (sum(ks.debet*ks.hp)-sum(ks.kredit*ks.hp))/(sum(ks.debet-ks.kredit)) as hpp,sum(ks.debet*ks.hp)-sum(ks.kredit*ks.hp) as NilaiPersediaan  "
  cSQL = cSQL & " FROM stock s"
  cSQL = cSQL & " LEFT JOIN golongan g on g.kodegolongan = s.kodegolongan "
  cSQL = cSQL & " LEFT JOIN kartustock ks on ks.kodestock = s.kodestock  "
  cSQL = cSQL & " WHERE s.kodestock = '" & cKodeStk & "'  AND s.jenis = 1"
  cSQL = cSQL & " GROUP BY s.kodegolongan,s.kodestock "
  cSQL = cSQL & " ORDER BY s.kodegolongan,s.kodestock"
  
  Set db = obj.SQL(GetDSN, cSQL)
  If Not db.EOF Then
    'obj.Edit GetDSN, "stock", "kodestock = '" & Stock & "'", Array("cogs"), Array(db!hpp)
    'NewUpdHargaPokok = GetNull(db!hpp)
     n = GetSaldoStockOnly(obj, cKodeStk)
     NewUpdHargaPokok = GetNull(db!nilaipersediaan) / IIf(n = 0, 1, n)
'    MsgBox GetNull(db!hpp)
'    MsgBox GetNull(db!nilaipersediaan)
  End If
  
End Function

Private Function GetSaldoStockOnly(ByVal obj As CodeSuiteLibrary.Data, cKodeStk As String) As Double
' Fungsi ini untuk medapatkan saldo stock dari table kartustock
' Dimana, dalam tabel kartustock ada field ljenis yg membedakan stock
' ljenis = 1 artinya real stock
' ljenis = 0 artinya dummy stock
' dummy stock muncul dari proses refund
' Refund adalah sebuah proses : pengurangan nilai persediaan, hanya nilai persediaan tanpa mengurangi qty stock itu sendiri
' Jadi : karena kita menggunakan tabel kartustock untuk menampung seluruh mutasi stock (apakah untuk menyimpan qty ataukah nilai persediaan)
' maka ketika kita ingin mendapatkan saldo stcok, perlu dipisah.
' Note : persediaan = qty * hp

Dim cSQL As String
Dim db As New ADODB.Recordset

  GetSaldoStockOnly = 1
  cSQL = ""
  cSQL = cSQL & " SELECT SUM(debet-kredit) as stok FROM kartustock ks "
  cSQL = cSQL & " WHERE ks.kodestock = '" & cKodeStk & "' and ljenis = 1"
  
  Set db = obj.SQL(GetDSN, cSQL)
  If Not db.EOF Then
    GetSaldoStockOnly = GetNull(db!stok)
  End If
End Function

Function UpdKodeTr(ByVal obj As CodeSuiteLibrary.Data, ByVal par As vbTrigger, ByVal cFaktur As String, _
              ByVal dTgl As Date, ByVal cRekening As String, ByVal CostCenter As String, _
              Optional ByVal cKeterangan As String = "", Optional ByVal nDebet As Double = 0, _
              Optional ByVal nKredit As Double = 0, Optional cKas As String = "K", _
              Optional ByVal cNow As String = "", _
              Optional ByVal cKodeStk As String = "") As Boolean

' Fungsi untuk menyimpan transaksi ke tabel bukubesar

Dim vaField, vaValue
UpdKodeTr = True
  
  cNow = IIf(cNow = "", SNow, cNow)
  If nDebet <> 0 Or nKredit <> 0 Then
    vaField = Array("faktur", "tgl", "kodeakun", "keterangan", "debet", "kredit", _
                    "kas", "username", "status", "datetime", "kodecostcenter", "kodestock")
    vaValue = Array(cFaktur, dTgl, cRekening, cKeterangan, nDebet, nKredit, _
                    cKas, GetRegistry(reg_Username), par, cNow, CostCenter, cKodeStk)
    UpdKodeTr = obj.Add(GetDSN, "bukubesar", vaField, vaValue)
    'UpdCfg msGudangPembelian, cGudangPembelian.Text, objData, cGudangPembelian.Caption, Me.Caption
  End If
End Function

Function DelKodeTr(ByVal obj As CodeSuiteLibrary.Data, ByVal par As vbTrigger, ByVal Faktur As String) As Boolean
DelKodeTr = True
  
  DelKodeTr = obj.Delete(GetDSN, "bukubesar", "faktur", sisAssign, Faktur, " and status = '" & par & "'")
End Function

Function TypeRekening(ByVal cRekening) As SisTypeRekening
  cRekening = Left(cRekening, 1)
  TypeRekening = Val(cRekening)
End Function


Function GetAkunInventory(ByVal obj As CodeSuiteLibrary.Data, ByVal Kode As String) As String
Dim db As New ADODB.Recordset

  GetAkunInventory = ""
'  Set db = obj.Browse(GetDSN, "stock", , "kodestock", sisAssign, Kode)
'  If Not db.EOF Then
'    GetAkunInventory = GetNull(db!kodeakun)
'  End If
  GetAkunInventory = aCfg(obj, msRekeningPersediaan)
End Function

Function GetAkunSupplier(ByVal obj As CodeSuiteLibrary.Data, ByVal Kode As String) As String

  GetAkunSupplier = ""
  GetAkunSupplier = aCfg(obj, msHutangDagang)
  
End Function

Function GetAkunMember(ByVal obj As CodeSuiteLibrary.Data, ByVal Kode As String) As String

  GetAkunMember = ""
  GetAkunMember = aCfg(obj, msPiutangDagang)
End Function

Function GetHargaBeli(ByVal obj As CodeSuiteLibrary.Data, KodeStock) As Double
Dim db As New ADODB.Recordset
  
  GetHargaBeli = 0
  Set db = obj.Browse(GetDSN, "stock", "hargabeli", "kodestock", sisAssign, KodeStock)
  If Not db.EOF Then
    GetHargaBeli = GetNull(db!hargabeli)
  End If
End Function

Function GetHargaJual(ByVal obj As CodeSuiteLibrary.Data, barcode) As Double
Dim db As New ADODB.Recordset
  
  GetHargaJual = 0
  Set db = obj.Browse(GetDSN, "stock", "hargajual", "barcode", sisAssign, barcode)
  If Not db.EOF Then
    GetHargaJual = GetNull(db!HargaJual)
  End If
End Function

Function GetHargaPokok(ByVal obj As CodeSuiteLibrary.Data, ByVal KodeStock As String, Optional ByVal nHargaPokok As Double = 0) As Double
Dim db As New ADODB.Recordset
Dim cSQL As String
  
  'Berhubung kita harus memasukkan stok non inventory ke dalam kartu stok
  'jadi hpp nya pun harus kita sesuaikan, karena hpp dari stok non inventory diinput
  
  GetHargaPokok = 0
  If nHargaPokok <> 0 Then
    GetHargaPokok = nHargaPokok
  Else
    Set db = obj.Browse(GetDSN, "stock", "cogs,hargabeli,hargajual", "kodestock", sisAssign, KodeStock)
    If Not db.EOF Then
      If GetNull(db!cogs) > 0 Then
        If GetNull(db!cogs) >= GetNull(db!HargaJual) Then
          GetHargaPokok = GetNull(db!hargabeli)
        Else
          GetHargaPokok = GetNull(db!cogs)
        End If
      Else
        GetHargaPokok = GetNull(db!hargabeli)
      End If
    End If
  End If
End Function
