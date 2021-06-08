Attribute VB_Name = "UpdRekening"
Option Explicit

Dim dbData As New ADODB.Recordset

Public Enum SisBukuBesar
  bb_Pembelian = 10
  bb_ReturPembelian = 11
  bb_PelunasanHutang = 12
  bb_penjualan = 13
  
  bb_ReturPenjualan = 20
  bb_PelunasanPiutang = 21
  bb_JurnalLain = 22
  
  bb_PenyesuaianStock = 30
  bb_MutasiKasBank = 31
End Enum

Sub UpdBukuBesar(ByVal obj As SISMyDLL.data, ByVal par As SisBukuBesar, ByVal cFaktur As String, _
                 ByVal dTgl As Date, ByVal cRekening As String, _
                 Optional ByVal cKeterangan As String = "", Optional ByVal nDebet As Double = 0, _
                 Optional ByVal nKredit As Double = 0, Optional cKas As String = "K", _
                 Optional ByVal cNow As String = "")
              
Dim vaField, vaValue

  cNow = IIf(cNow = "", SNow, cNow)
  If (nDebet <> 0 Or nKredit <> 0) And dTgl >= DateSerial(2003, 1, 1) Then
    vaField = Array("Faktur", "Tgl", "Rekening", "Keterangan", "Debet", "Kredit", _
                    "Kas", "UserName", "Status", "DateTime")
    vaValue = Array(cFaktur, dTgl, cRekening, cKeterangan, nDebet, nKredit, _
                    cKas, GetRegistry(reg_UserName), par, cNow)
    obj.Add GetDSN, "BukuBesar", vaField, vaValue
  End If
End Sub

Sub DelBukuBesar(ByVal obj As SISMyDLL.data, ByVal par As SisBukuBesar, Optional ByVal cFaktur As String = "")
Dim cWhere As String
  ' Jika Faktur Kosong maka hapus untuk semua Buku besar jika status = par
  If Trim(cFaktur = "") Then
    cWhere = " and Posting = ' '"
  Else
    cWhere = " And Faktur = '" & cFaktur & "'"
  End If
  
  obj.Delete GetDSN, "BukuBesar", "Status", sisAssign, Trim(Str(par)), cWhere
End Sub

Function UpdRekeningPembelian(ByVal obj As SISMyDLL.data, ByVal cFaktur As String)
Dim db As New ADODB.Recordset

  ' Dr. Persediaan      SubTotal - Discount
  ' Dr. PPn Masukkan    Pajak
  '   Cr. Hutang          Hutang
  '   Cr. Kas             Tunai
    
  DelBukuBesar obj, bb_Pembelian, cFaktur
  Set db = obj.Browse(GetDSN, "TotPembelian t", "t.*,s.Nama as NamaSupplier", "Faktur", sisAssign, cFaktur, , , _
           Array("Left Join Supplier s on t.Supplier = s.Kode"))
  If db.RecordCount > 0 Then
    UpdBukuBesar obj, bb_Pembelian, cFaktur, db!Tgl, aCfg(msRekeningPersediaan), "Pembelian an. " & db!NamaSupplier, db!Subtotal - db!discount
    UpdBukuBesar obj, bb_Pembelian, cFaktur, db!Tgl, aCfg(msRekeningPPnMasukkan), "Pembelian an. " & db!NamaSupplier, db!Pajak
      UpdBukuBesar obj, bb_Pembelian, cFaktur, db!Tgl, aCfg(msRekeningHutang), "Pembelian an. " & db!NamaSupplier, , db!Hutang
      UpdBukuBesar obj, bb_Pembelian, cFaktur, db!Tgl, aCfg(msRekeningKas), "Pembelian an. " & db!NamaSupplier, , db!Tunai
  End If
End Function

Function UpdRekeningMutasiKasBank(ByVal obj As SISMyDLL.data, ByVal cFaktur As String)
Dim db As New ADODB.Recordset
    
  DelBukuBesar obj, bb_MutasiKasBank, cFaktur
  Set db = obj.Browse(GetDSN, "mutasikasbank m", "m.*,b.rekening as RekeningKredit, c.rekening as RekeningDebet,b.Keterangan as DariBank, c.Keterangan as KeBank", , , , , , _
            Array("left join bank b on b.kode = m.dari", _
                  "left join bank c on c.kode = m.ke"))
  If db.RecordCount > 0 Then
    'Dr. Ke
      'Cr. Dari
      
    UpdBukuBesar obj, bb_MutasiKasBank, cFaktur, db!Tgl, GetNull(db!RekeningDebet), db!Keterangan, db!Jumlah, 0
      UpdBukuBesar obj, bb_MutasiKasBank, cFaktur, db!Tgl, GetNull(db!RekeningKredit), db!Keterangan, 0, db!Jumlah
  End If
End Function

Function UpdRekeningReturPembelian(ByVal obj As SISMyDLL.data, ByVal cFaktur As String)
Dim db As New ADODB.Recordset
  
  ' Dr. Hutang          Hutang
  ' Dr. Kas             Tunai
  '   Cr. Persediaan      SubTotal - Discount
  '   Cr. PPn Masukkan    Pajak
    
  DelBukuBesar obj, bb_ReturPembelian, cFaktur
  Set db = obj.Browse(GetDSN, "TotRtnPembelian t", "t.*,s.Nama as NamaSupplier", "Faktur", sisAssign, cFaktur, , , _
           Array("Left Join Supplier s on t.Supplier = s.Kode"))
  If db.RecordCount > 0 Then
    UpdBukuBesar obj, bb_ReturPembelian, cFaktur, db!Tgl, aCfg(msRekeningHutang), "Retur Pembelian an. " & db!NamaSupplier, db!Hutang
    UpdBukuBesar obj, bb_ReturPembelian, cFaktur, db!Tgl, aCfg(msRekeningKas), "Retur Pembelian an. " & db!NamaSupplier, db!Tunai
      UpdBukuBesar obj, bb_ReturPembelian, cFaktur, db!Tgl, aCfg(msRekeningPersediaan), "Retur Pembelian an. " & db!NamaSupplier, , db!Subtotal - db!discount
      UpdBukuBesar obj, bb_ReturPembelian, cFaktur, db!Tgl, aCfg(msRekeningPPnMasukkan), "Retur Pembelian an. " & db!NamaSupplier, , db!Pajak
  End If
End Function

Sub UpdRekeningPenjualan(ByVal obj As SISMyDLL.data, ByVal cFaktur As String)
Dim cStatus As SisBukuBesar
Dim nDiscount As Double
  
  cStatus = bb_penjualan
  Set dbData = obj.Browse(GetDSN, "TotPenjualan t", "t.Faktur,t.Pajak,t.SubTotal,t.Tgl,c.Nama,t.Total,t.Discount,t.Discount2,t.Tunai,t.Piutang,Sum(d.hp) as hp", "t.Faktur", sisAssign, cFaktur, " Group by t.Faktur", , _
               Array("Left Join Customer c on t.Customer = c.Kode", _
                     "Left Join Penjualan d on d.Faktur = t.Faktur"))
               
  DelBukuBesar obj, cStatus, cFaktur
  If dbData.RecordCount > 0 Then
    nDiscount = dbData!discount + dbData!Discount2
    ' Piutang Datang          !Piutang
    ' Kas                     !Tunai
    ' Discount Penjualan      !Discount
    '    Ppn. Masukkan          !Pajak
    '    SubTotal               !SubTotal
    ' Hpp                     !HP
    '    Persediaan             !HP
    
    
    UpdBukuBesar obj, bb_penjualan, dbData!Faktur, dbData!Tgl, aCfg(msRekeningPiutang), "Penjualan an. " & dbData!Nama, dbData!Piutang
    UpdBukuBesar obj, bb_penjualan, dbData!Faktur, dbData!Tgl, aCfg(msRekeningKas), "Penjualan an. " & dbData!Nama, dbData!Tunai
    UpdBukuBesar obj, bb_penjualan, dbData!Faktur, dbData!Tgl, aCfg(msRekeningDiscountPenjualan), "Discount Penjualan an. " & dbData!Nama, nDiscount
      UpdBukuBesar obj, bb_penjualan, dbData!Faktur, dbData!Tgl, aCfg(msRekeningPPnMasukkan), "PPn. Masukkan an. " & dbData!Nama, , dbData!Pajak
      UpdBukuBesar obj, bb_penjualan, dbData!Faktur, dbData!Tgl, aCfg(msRekeningPenjualan), "Penjualan an. " & dbData!Nama, , dbData!Subtotal
      
    UpdBukuBesar obj, bb_penjualan, dbData!Faktur, dbData!Tgl, aCfg(msRekeningHargaPokokPenjualan), "Penjualan an. " & dbData!Nama, dbData!hp
      UpdBukuBesar obj, bb_penjualan, dbData!Faktur, dbData!Tgl, aCfg(msRekeningPersediaan), "Penjualan an. " & dbData!Nama, , dbData!hp
  End If
End Sub

Sub UpdRekeningRtnPenjualan(ByVal obj As SISMyDLL.data, ByVal cFaktur As String)
Dim cStatus As SisBukuBesar
Dim nDiscount As Double
  
  cStatus = bb_ReturPenjualan
  Set dbData = obj.Browse(GetDSN, "TotRtnPenjualan t", "t.Faktur,t.Pajak,t.SubTotal,t.Tgl,c.Nama,t.Total,t.Discount,t.Discount2,t.Tunai,t.Piutang,Sum(d.hp) as hp", "t.Faktur", sisAssign, cFaktur, " Group by t.Faktur", , _
               Array("Left Join Customer c on t.Customer = c.Kode", _
                     "Left Join RtnPenjualan d on d.Faktur = t.Faktur"))
               
  DelBukuBesar obj, cStatus, cFaktur
  If dbData.RecordCount > 0 Then
    nDiscount = dbData!discount + dbData!Discount2
    ' Ppn. Masukkan          !Pajak
    ' SubTotal               !SubTotal
    '   Piutang Datang          !Piutang
    '   Kas                     !Tunai
    '   Discount Penjualan      !Discount
    
    ' Persediaan             !HP
    '   Hpp                     !HP
    
    UpdBukuBesar obj, cStatus, dbData!Faktur, dbData!Tgl, aCfg(msRekeningPPnMasukkan), "PPn. Masukkan an. " & dbData!Nama, dbData!Pajak
    UpdBukuBesar obj, cStatus, dbData!Faktur, dbData!Tgl, aCfg(msRekeningReturPenjualan), "Retur Penjualan an. " & dbData!Nama, dbData!Subtotal
      UpdBukuBesar obj, cStatus, dbData!Faktur, dbData!Tgl, aCfg(msRekeningPiutang), "Retur Penjualan an. " & dbData!Nama, , dbData!Piutang
      UpdBukuBesar obj, cStatus, dbData!Faktur, dbData!Tgl, aCfg(msRekeningKas), "Retur Penjualan an. " & dbData!Nama, , dbData!Tunai
      UpdBukuBesar obj, cStatus, dbData!Faktur, dbData!Tgl, aCfg(msRekeningDiscountPenjualan), "Discount Penjualan an. " & dbData!Nama, , nDiscount
     
    UpdBukuBesar obj, cStatus, dbData!Faktur, dbData!Tgl, aCfg(msRekeningPersediaan), "Retur Penjualan an. " & dbData!Nama, dbData!hp
      UpdBukuBesar obj, cStatus, dbData!Faktur, dbData!Tgl, aCfg(msRekeningHargaPokokPenjualan), "Retur Penjualan an. " & dbData!Nama, , dbData!hp
  End If
End Sub

Sub UpdRekeningPenjualanKasir(ByVal obj As SISMyDLL.data, ByVal cFaktur As String)
Dim cStatus As SisBukuBesar
Dim nDiscount As Double
  
  cStatus = bb_penjualan
  Set dbData = obj.Browse(GetDSN, "TotKasir t", "t.Faktur,t.SubTotal,t.Tgl,t.Discount,Sum(d.hp) as hp", "t.Faktur", sisAssign, cFaktur, " Group by t.Faktur", , _
               Array("Left Join Kasir d on d.Faktur = t.Faktur"))
               
  DelBukuBesar obj, cStatus, cFaktur
  If dbData.RecordCount > 0 Then
    nDiscount = dbData!discount
    ' Kas                     !Tunai
    ' Discount Penjualan      !Discount
    '    SubTotal               !SubTotal
    
    ' Hpp                     !HP
    '    Persediaan             !HP
    
    UpdBukuBesar obj, cStatus, dbData!Faktur, dbData!Tgl, aCfg(msRekeningKas), "Penjualan Kasir", dbData!Subtotal - dbData!discount
    UpdBukuBesar obj, cStatus, dbData!Faktur, dbData!Tgl, aCfg(msRekeningDiscountPenjualan), "Discount Penjualan Kasir", nDiscount
      UpdBukuBesar obj, cStatus, dbData!Faktur, dbData!Tgl, aCfg(msRekeningPenjualan), "Penjualan Kasir", , dbData!Subtotal
    
    UpdBukuBesar obj, cStatus, dbData!Faktur, dbData!Tgl, aCfg(msRekeningHargaPokokPenjualan), "Penjualan Kasir", dbData!hp
      UpdBukuBesar obj, cStatus, dbData!Faktur, dbData!Tgl, aCfg(msRekeningPersediaan), "Penjualan Kasir", , dbData!hp
  End If
End Sub

Sub UpdRekeningPenyesuaian(ByVal obj As SISMyDLL.data, ByVal cFaktur As String)
Dim db As New ADODB.Recordset
Dim cStatus As SisBukuBesar
Dim nDiscount As Double
  
  cStatus = bb_PenyesuaianStock
  Set db = obj.Browse(GetDSN, "adjstock a", "a.*", "Faktur", sisAssign, cFaktur)
  ' Jika Stock Lebih maka
  ' Dr. Persediaan          !HP * !Qty
  '   Cr. Rekening Lebih      !HP * !Qty

  ' Jika Kurang
  ' Dr. Rekening Kurang     !HP * !Qty
  '    Cr. Persediaan         !HP * !Qty
  
  If db.RecordCount > 0 Then
    UpdBukuBesar obj, bb_PenyesuaianStock, cFaktur, db!Tgl, aCfg(msRekeningPersediaan), db!Keterangan, db!hp * db!qty, 0
      UpdBukuBesar obj, bb_PenyesuaianStock, cFaktur, db!Tgl, aCfg(msRekeningPenyesuaianLebih), db!Keterangan, 0, db!hp * db!qty
    
    UpdBukuBesar obj, bb_PenyesuaianStock, cFaktur, db!Tgl, aCfg(msRekeningPenyesuaianKurang), db!Keterangan, db!hp * db!qty, 0
      UpdBukuBesar obj, bb_PenyesuaianStock, cFaktur, db!Tgl, aCfg(msRekeningPersediaan), db!Keterangan, 0, db!hp * db!qty
  End If
End Sub

