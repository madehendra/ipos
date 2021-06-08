Attribute VB_Name = "Report"
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data

Function GetMutasiHutang(ByVal dTglAwal As Date, ByVal dTglAkhir As Date, _
                         ByVal cWilayahAwal As String, ByVal cWilayahAkhir As String, _
                         ByVal cKodeAwal As String, ByVal cKodeAkhir As String) As XArrayDB

Dim vaArray As New XArrayDB
Dim cWhere As String
Dim cField As String
Dim vaJoin
Dim n As Double

  ' Field Result
  ' Wilayah,Namawilayah,Kode,Nama,Alamat,Awal,debet,kredit,Akhir
  
  
  
  cWhere = "c.Wilayah >= '" & cWilayahAwal & "' and c.Wilayah <= '" & cWilayahAkhir & "' "
  cWhere = cWhere & " and c.Kode >= '" & cKodeAwal & "' and c.Kode <= '" & cKodeAkhir & "' "
  cWhere = cWhere & "Group by c.Wilayah,c.Kode"
  cField = "c.Wilayah,w.Keterangan as NamaWilayah,c.Kode,c.Nama,c.Alamat,Sum(a.debet-a.kredit) as Awal,0 as debet,0 as kredit,Sum(a.debet-a.kredit) as Akhir"
  vaJoin = Array("Left Join kartuhutang a on c.Kode = a.Supplier and a.Tgl < '" & Format(dTglAwal, "yyyy-MM-dd") & "'", _
                 "Left Join Wilayah w on w.Kode = c.Wilayah")
                 
  Set dbData = objData.Browse(GetDSN, "Supplier c", cField, , , , cWhere, "c.Wilayah,c.Kode", vaJoin)
  vaArray.LoadRows dbData.GetRows
  
  ' ambil Data Mutasi
  cWhere = "Tgl >= '" & Format(dTglAwal, "yyyy-MM-dd") & "' and Tgl <= '" & Format(dTglAkhir, "yyyy-MM-dd") & "' "
  cWhere = cWhere & " Group by Supplier"
  Set dbData = objData.Browse(GetDSN, "kartuhutang", "Supplier,Sum(debet) as debet,Sum(kredit) as kredit", , , , cWhere)
  Do While Not dbData.EOF
    n = vaArray.Find(0, 2, (dbData!supplier))
    If n >= 0 Then
      vaArray(n, 6) = GetNull(vaArray(n, 6)) + GetNull(dbData!debet)
      vaArray(n, 7) = GetNull(vaArray(n, 7)) + GetNull(dbData!kredit)
      vaArray(n, 8) = GetNull(vaArray(n, 5)) + GetNull(vaArray(n, 6)) - GetNull(vaArray(n, 7))
    End If
    dbData.MoveNext
  Loop
  
  
  
  Set GetMutasiHutang = vaArray
End Function


Function GetMutasiPiutang(ByVal dTglAwal As Date, ByVal dTglAkhir As Date, _
                          ByVal cWilayahAwal As String, ByVal cWilayahAkhir As String, _
                          ByVal cKodeAwal As String, ByVal cKodeAkhir As String) As XArrayDB
Dim vaArray As New XArrayDB
Dim cWhere As String
Dim cField As String
Dim vaJoin
Dim n As Double

  ' Field Result
  ' Wilayah,Namawilayah,Kode,Nama,Alamat,Awal,debet,kredit,Akhir
  
  
  cWhere = "c.Wilayah >= '" & cWilayahAwal & "' and c.Wilayah <= '" & cWilayahAkhir & "' "
  cWhere = cWhere & " and c.Kode >= '" & cKodeAwal & "' and c.Kode <= '" & cKodeAkhir & "' "
  cWhere = cWhere & "Group by c.Wilayah,c.Kode"
  cField = "c.Wilayah,w.Keterangan as NamaWilayah,c.Kode,c.Nama,c.Alamat,Sum(a.debet-a.kredit) as Awal,0 as debet,0 as kredit,Sum(a.debet-a.kredit) as Akhir"
  vaJoin = Array("Left Join KartuPiutang a on c.Kode = a.Customer and a.Tgl < '" & Format(dTglAwal, "yyyy-MM-dd") & "'", _
                 "Left Join Wilayah w on w.Kode = c.Wilayah")
                 
  Set dbData = objData.Browse(GetDSN, "Customer c", cField, , , , cWhere, "c.Wilayah,c.Kode", vaJoin)
  vaArray.LoadRows dbData.GetRows
  
  ' ambil Data Mutasi
  cWhere = "Tgl >= '" & Format(dTglAwal, "yyyy-MM-dd") & "' and Tgl <= '" & Format(dTglAkhir, "yyyy-MM-dd") & "' "
  cWhere = cWhere & " Group by Customer"
  Set dbData = objData.Browse(GetDSN, "KartuPiutang", "Customer,Sum(debet) as debet,Sum(kredit) as kredit", , , , cWhere)
  Do While Not dbData.EOF
    n = vaArray.Find(0, 2, (dbData!Customer))
    If n >= 0 Then
      vaArray(n, 6) = vaArray(n, 6) + (dbData!debet)
      vaArray(n, 7) = vaArray(n, 7) + (dbData!kredit)
      vaArray(n, 8) = vaArray(n, 5) + vaArray(n, 6) - vaArray(n, 7)
    End If
    dbData.MoveNext
  Loop
  
  Set GetMutasiPiutang = vaArray
End Function


Function Getkartuhutang(ByVal cSupplier As String, ByVal dTglAwal As Date, ByVal dTglAkhir As Date) As XArrayDB
Dim nAwal As Double
Dim cSQL As String
Dim n As Double
Dim vaArray As New XArrayDB

  ' Field Result
  ' Tgl, Keterangan, Faktur,debet,kredit,Saldo
  vaArray.ReDim 0, 0, 0, 5
  
  ' Ambil Data Awal
  Set dbData = objData.Browse(GetDSN, "kartuhutang", "Sum(debet) as debet,Sum(kredit) as kredit", "kodesupplier", sisAssign, cSupplier, _
               "AND tgl < '" & Format(dTglAwal, "yyyy-MM-dd") & "'")
  If Not dbData.EOF Then
    nAwal = GetNull(dbData!debet, 0) - GetNull(dbData!kredit, 0)
  End If
  vaArray(0, 1) = "Saldo Awal"
  vaArray(0, 3) = 0
  vaArray(0, 4) = 0
  vaArray(0, 5) = nAwal
  
  ' Ambil Data Mutasi
  Set dbData = objData.Browse(GetDSN, "kartuhutang", "tgl,keterangan,nomorkartuhutang,debet,kredit", "kodesupplier", sisAssign, cSupplier, _
               " AND tgl >= '" & Format(dTglAwal, "yyyy-MM-dd") & "' AND tgl <= '" & Format(dTglAkhir, "yyyy-MM-dd") & "'", _
               "kodesupplier,tgl,id")
  If dbData.RecordCount > 0 Then
    dbData.MoveFirst
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      
      vaArray(n, 0) = Format(GetNull(dbData!tgl), "dd-MM-yyyy")
      vaArray(n, 1) = GetNull(dbData!keterangan)
      vaArray(n, 2) = GetNull(dbData!nomorkartuhutang)
      vaArray(n, 3) = GetNull(dbData!debet)
      vaArray(n, 4) = GetNull(dbData!kredit)
      nAwal = nAwal + GetNull(vaArray(n, 3)) - GetNull(vaArray(n, 4))
      vaArray(n, 5) = nAwal
      
      dbData.MoveNext
    Loop
  End If
  
  Set Getkartuhutang = vaArray
End Function

Function Getkartupiutang(ByVal cCustomer As String, ByVal dTglAwal As Date, ByVal dTglAkhir As Date, ByVal lSaldoAwal As Boolean) As XArrayDB
Dim nAwal As Double
Dim cSQL As String
Dim n As Double
Dim vaArray As New XArrayDB

  ' Field Result
  ' Tgl, Keterangan, nomorkartupiutang,debet,kredit,Saldo
  vaArray.ReDim 0, 0, 0, 6
  
  If lSaldoAwal = True Then
    ' Ambil Data Awal
    Set dbData = objData.Browse(GetDSN, "kartupiutang", "Sum(debet) as debet,Sum(kredit) as kredit", "kodeanggota", sisAssign, cCustomer, _
                 "AND tgl < '" & Format(dTglAwal, "yyyy-MM-dd") & "'")
    If Not dbData.EOF Then
      nAwal = GetNull(dbData!debet, 0) - GetNull(dbData!kredit, 0)
    End If
    vaArray(0, 1) = "Saldo Awal"
    vaArray(0, 3) = 0
    vaArray(0, 4) = 0
    vaArray(0, 5) = nAwal
    vaArray(0, 6) = ""
  End If
  
  ' Ambil Data Mutasi
  Set dbData = objData.Browse(GetDSN, "kartupiutang", "tgl,keterangan,nomorkartupiutang,debet,kredit,groupsales", "kodeanggota", sisAssign, cCustomer, _
               " AND tgl >= '" & Format(dTglAwal, "yyyy-MM-dd") & "' AND Tgl <= '" & Format(dTglAkhir, "yyyy-MM-dd") & "'", _
               "kodeanggota,tgl,id")
  If dbData.RecordCount > 0 Then
    dbData.MoveFirst
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      
      vaArray(n, 0) = Format(GetNull(dbData!tgl), "dd-MM-yyyy")
      vaArray(n, 1) = GetNull(dbData!keterangan)
      vaArray(n, 2) = GetNull(dbData!nomorkartupiutang)
      vaArray(n, 3) = GetNull(dbData!debet)
      vaArray(n, 4) = GetNull(dbData!kredit)
      nAwal = nAwal + GetNull(vaArray(n, 3)) - GetNull(vaArray(n, 4))
      vaArray(n, 5) = nAwal
      vaArray(n, 6) = GetNull(dbData!GroupSales)
      
      dbData.MoveNext
    Loop
  End If
  
  Set Getkartupiutang = vaArray
End Function


Function GetKartuTopUp(ByVal cCustomer As String, ByVal dTglAwal As Date, ByVal dTglAkhir As Date, ByVal lSaldoAwal As Boolean) As XArrayDB
Dim nAwal As Double
Dim cSQL As String
Dim n As Double
Dim vaArray As New XArrayDB

  ' Field Result
  ' Tgl, Keterangan, nomorkartupiutang,debet,kredit,Saldo
  vaArray.ReDim 0, 0, 0, 5
  
  If lSaldoAwal = True Then
    ' Ambil Data Awal
    Set dbData = objData.Browse(GetDSN, "membertopup", "Sum(debet) as debet,Sum(kredit) as kredit", "kodeanggota", sisAssign, cCustomer, _
                 "AND tgl < '" & Format(dTglAwal, "yyyy-MM-dd") & "'")
    If Not dbData.EOF Then
      nAwal = GetNull(dbData!debet, 0) - GetNull(dbData!kredit, 0)
    End If
    vaArray(0, 1) = "Saldo Awal"
    vaArray(0, 3) = 0
    vaArray(0, 4) = 0
    vaArray(0, 5) = nAwal
  End If
  
  ' Ambil Data Mutasi
  Set dbData = objData.Browse(GetDSN, "membertopup", "tgl,keterangan,nomormembertopup,debet,kredit", "kodeanggota", sisAssign, cCustomer, _
               " AND tgl >= '" & Format(dTglAwal, "yyyy-MM-dd") & "' AND Tgl <= '" & Format(dTglAkhir, "yyyy-MM-dd") & "'", _
               "kodeanggota,tgl")
  If dbData.RecordCount > 0 Then
    dbData.MoveFirst
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      
      vaArray(n, 0) = Format(GetNull(dbData!tgl), "dd-MM-yyyy")
      vaArray(n, 1) = GetNull(dbData!keterangan)
      vaArray(n, 2) = GetNull(dbData!nomormembertopup)
      vaArray(n, 3) = GetNull(dbData!debet)
      vaArray(n, 4) = GetNull(dbData!kredit)
      nAwal = nAwal + GetNull(vaArray(n, 3)) - GetNull(vaArray(n, 4))
      vaArray(n, 5) = nAwal
      
      dbData.MoveNext
    Loop
  End If
  
  Set GetKartuTopUp = vaArray
End Function



Function GetRptPPnMasukkan(ByVal dTglAwal As Date, ByVal dTglAkhir As Date) As XArrayDB
Dim vaArray As New XArrayDB
Dim n As Double
Dim cField As String
  ' Field Result
  ' Faktur
  'Tgl
  'Nama Supplier
  'SubTotal
  'Discount1
  'Discount2
  'PPn
  'Total
  
  vaArray.ReDim 0, -1, 0, 7
  cField = "t.Faktur,t.Tgl,s.Nama,t.SubTotal,t.Discount,t.Discount2,t.Pajak,t.Total"
  Set dbData = objData.Browse(GetDSN, "TotPembelian t", cField, "t.Tgl", sisGTEqual, Format(dTglAwal, "yyyy-MM-dd"), _
               " And Tgl <= '" & Format(dTglAkhir, "yyyy-MM-dd") & "'", "t.Tgl,t.Faktur", _
               Array("Left Join Supplier s on t.Supplier = s.Kode"))
  If dbData.RecordCount > 0 Then
    vaArray.LoadRows dbData.GetRows(dbData.RecordCount)
    Do While n <= vaArray.UpperBound(1)
      If vaArray(n, 6) = 0 Then
        vaArray.DeleteRows n
        n = n - 1
      End If
      n = n + 1
    Loop
  End If
  
  Set GetRptPPnMasukkan = vaArray
End Function

Function GetRptPPnKeluaran(ByVal dTglAwal As Date, ByVal dTglAkhir As Date) As XArrayDB
Dim vaArray As New XArrayDB
Dim n As Double
Dim cField As String

' Field Result
' Faktur
'  Tgl
'  Nama Customer
'  SubTotal
'  Discount1
'  Discount2
'  PPn
'  Total
  
  vaArray.ReDim 0, -1, 0, 7
  cField = "t.Faktur,t.Tgl,s.Nama,t.SubTotal,t.Discount,t.Discount2,t.Pajak,t.Total"
  Set dbData = objData.Browse(GetDSN, "TotPenjualan t", cField, "t.Tgl", sisGTEqual, Format(dTglAwal, "yyyy-MM-dd"), _
               " And Tgl <= '" & Format(dTglAkhir, "yyyy-MM-dd") & "'", "t.Tgl,t.Faktur", _
               Array("Left Join Customer s on t.Customer = s.Kode"))
  If dbData.RecordCount > 0 Then
    vaArray.LoadRows dbData.GetRows(dbData.RecordCount)
    Do While n <= vaArray.UpperBound(1)
      If vaArray(n, 6) = 0 Then
        vaArray.DeleteRows n
        n = n - 1
      End If
      n = n + 1
    Loop
  End If
  Set GetRptPPnKeluaran = vaArray
End Function

Function GetRptProdukTerlaris(ByVal dTglAwal As Date, ByVal dTglAkhir As Date) As XArrayDB
Dim vaArray As New XArrayDB
Dim cWhere As String
Dim n As Double

  ' Cari Penjualan
  cWhere = " (k.Status = '60' or k.Status = '63' or k.Status = '11') And Tgl >= '" & Format(dTglAwal, "yyyy-MM-dd") & "' and Tgl <= '" & Format(dTglAkhir, "yyyy-MM-dd") & "' "
  cWhere = cWhere & " Group by k.Status,k.kodestock"
  Set dbData = objData.Browse(GetDSN, "kartustock k", "k.kodestock,s.kodegolongan,k.Status,s.Nama,s.barcode,s.kodesatuan,Sum(k.Qty) as Mutasi,Sum(k.Harga*k.Qty) as Harga", _
               , , , cWhere, "k.Status,k.kodestock", _
               Array("Left Join Stock s on k.kodestock = s.kodestock"))
  vaArray.ReDim 0, -1, 0, 5
  If dbData.RecordCount > 0 Then
    dbData.MoveFirst
    Do While Not dbData.EOF
      If vaArray.UpperBound(1) >= 0 Then
        n = vaArray.Find(0, 0, (dbData!KodeStock))
        If n = -1 Then
          vaArray.InsertRows vaArray.UpperBound(1) + 1
          n = vaArray.UpperBound(1)
        End If
      Else
        vaArray.InsertRows 0
        n = 0
      End If
      
      vaArray(n, 0) = GetNull(dbData!barcode)
      vaArray(n, 1) = GetNull(dbData!nama)
      vaArray(n, 2) = GetNull(dbData!kodegolongan)
      If dbData!Status = "11" Then        ' Jika Retur
        vaArray(n, 4) = GetNull(vaArray(n, 4), 0) + GetNull(dbData!MUTASI)
        vaArray(n, 5) = GetNull(vaArray(n, 5)) - GetNull(dbData!Harga)
      Else
        vaArray(n, 3) = GetNull(vaArray(n, 3)) + GetNull(dbData!MUTASI)
        vaArray(n, 5) = GetNull(vaArray(n, 5)) + GetNull(dbData!Harga)
      End If
      
      dbData.MoveNext
    Loop
  End If
  
  Set GetRptProdukTerlaris = vaArray
End Function

Function GetNilaiPersediaan(ByVal cGolonganAwal As String, ByVal cGolonganAkhir As String, Optional ByVal hargabeli As Boolean = True, Optional ByVal tgl As Date, Optional ByVal lAllGolongan As Boolean = True, Optional ByVal lKodeStock As Boolean = True, Optional ByVal lPilihGudang As Boolean = False, Optional ByVal cKodeGudang As String = "") As XArrayDB
Dim vaArray As New XArrayDB
Dim cField As String
Dim cWhere As String
Dim vaJoin
Dim n As Integer


  'Field array
  'Golongan
  'Nama Golongan
  'Kode
  'Nama
  'Satuan
  'Akhir
  'HP
  'Persediaan
  
  If hargabeli = True Then
    If lKodeStock = True Then
      cField = " s.kodegolongan,g.keterangan as namagolongan,s.kodestock,s.nama,s.kodesatuan,sum(ks.debet - ks.kredit) as akhir ,s.hargabeli, (sum(ks.debet*ks.hp)-sum(ks.kredit*ks.hp))/(sum(ks.debet-ks.kredit)), sum(ks.debet*ks.hp)-sum(ks.kredit*ks.hp) as NilaiPersediaan "
    Else
      cField = " s.kodegolongan,g.keterangan as namagolongan,s.barcode,s.nama,s.kodesatuan,sum(ks.debet - ks.kredit) as akhir ,s.hargabeli, (sum(ks.debet*ks.hp)-sum(ks.kredit*ks.hp))/(sum(ks.debet-ks.kredit)),round(sum(ks.debet*ks.hp),6)-round(sum(ks.kredit*ks.hp),6) as NilaiPersediaan "
    End If
  Else
    If lKodeStock = True Then
      cField = " s.kodegolongan,g.keterangan as namagolongan,s.kodestock,s.nama,s.kodesatuan,sum(ks.debet - ks.kredit) as akhir ,s.hargabeli,(sum(ks.debet*ks.hp)-sum(ks.kredit*ks.hp))/(sum(ks.debet-ks.kredit)),round(sum(ks.debet*ks.hp),6)-round(sum(ks.kredit*ks.hp),6) as NilaiPersediaan "
    Else
      cField = " s.kodegolongan,g.keterangan as namagolongan,s.barcode,s.nama,s.kodesatuan,sum(ks.debet - ks.kredit) as akhir ,s.hargabeli ,(sum(ks.debet*ks.hp)-sum(ks.kredit*ks.hp))/(sum(ks.debet-ks.kredit)),round(sum(ks.debet*ks.hp),6)-round(sum(ks.kredit*ks.hp),6) as NilaiPersediaan "
    End If
  End If
  
  vaJoin = Array("LEFT JOIN golongan g on g.kodegolongan = s.kodegolongan", _
                 "LEFT JOIN kartustock ks on ks.kodestock = s.kodestock")
  If lAllGolongan = False Then
    cWhere = "s.jenis = 1 AND s.kodegolongan >= '" & cGolonganAwal & "' AND  s.kodegolongan <= '" & cGolonganAkhir & "'"
  Else
    cWhere = "s.jenis = 1 "
  End If
  
  
  cWhere = cWhere & " AND ks.tgl <= '" & Format(tgl, "yyyy-MM-dd") & "'"
  
  If lPilihGudang = True Then
    cWhere = cWhere & " AND ks.kodegudang = '" & cKodeGudang & "'"
  End If
  
  cWhere = cWhere & " GROUP BY s.kodegolongan,s.kodestock"
  Set dbData = objData.Browse(GetDSN, "stock s", cField, , , , cWhere, "s.kodegolongan,s.kodestock", vaJoin)
  If dbData.RecordCount > 0 Then
      vaArray.LoadRows dbData.GetRows(dbData.RecordCount)
      n = 0
      Do While n <= vaArray.UpperBound(1)
        'objData.Edit GetDSN, "stock", "kodestock = '" & GetNull(dbData!KodeStock) & "'", Array("tcogs"), Array(vaArray(n, 6))
        If vaArray(n, 5) = 0 Then
          vaArray.DeleteRows n
          n = n - 1
        End If
        n = n + 1
      Loop
  Else
    Exit Function
  End If
  Set GetNilaiPersediaan = vaArray
End Function

Function GetNilaiPersediaan2(ByVal cGolonganAwal As String, ByVal cGolonganAkhir As String, Optional ByVal hargabeli As Boolean = True, Optional ByVal tgl As Date, Optional ByVal lAllGolongan As Boolean = True, Optional ByVal lKodeStock As Boolean = True, Optional ByVal nMinimumStock As Double) As XArrayDB
Dim vaArray As New XArrayDB
Dim cField As String
Dim cWhere As String
Dim vaJoin
Dim n As Integer


  'Field array
  'Golongan
  'Nama Golongan
  'Kode
  'Nama
  'Satuan
  'Akhir
  'HP
  'Persediaan
  
  If hargabeli = True Then
    If lKodeStock = True Then
      cField = " s.kodegolongan,g.keterangan as namagolongan,s.kodestock,s.nama,s.kodesatuan,sum(ks.debet - ks.kredit) as akhir , s.hargabeli,s.hargajual "
    Else
      cField = " s.kodegolongan,g.keterangan as namagolongan,s.barcode,s.nama,s.kodesatuan,sum(ks.debet - ks.kredit) as akhir , s.hargabeli,s.hargajual "
    End If
  Else
    If lKodeStock = True Then
      cField = " s.kodegolongan,g.keterangan as namagolongan,s.kodestock,s.nama,s.kodesatuan,sum(ks.debet - ks.kredit) as akhir , s.hargabeli,s.hargajual "
    Else
      cField = " s.kodegolongan,g.keterangan as namagolongan,s.barcode,s.nama,s.kodesatuan,sum(ks.debet - ks.kredit) as akhir , s.hargabeli,s.hargajual "
    End If
  End If
  
  vaJoin = Array("LEFT JOIN golongan g on g.kodegolongan = s.kodegolongan", _
                 "LEFT JOIN kartustock ks on ks.kodestock = s.kodestock")
  If lAllGolongan = False Then
    cWhere = "s.jenis = 1 AND s.kodegolongan >= '" & cGolonganAwal & "' AND  s.kodegolongan <= '" & cGolonganAkhir & "'"
  Else
    cWhere = "s.jenis = 1 "
  End If
  cWhere = cWhere & " AND ks.tgl <= '" & Format(tgl, "yyyy-MM-dd") & "'"
  cWhere = cWhere & " GROUP BY s.kodegolongan,s.kodestock"
  Set dbData = objData.Browse(GetDSN, "stock s", cField, , , , cWhere, "s.kodegolongan,s.kodestock", vaJoin)
  If dbData.RecordCount > 0 Then
      vaArray.LoadRows dbData.GetRows(dbData.RecordCount)
       n = 0
      Do While n <= vaArray.UpperBound(1)
        If vaArray(n, 5) = 0 Or vaArray(n, 5) > nMinimumStock Then
          vaArray.DeleteRows n
          n = n - 1
        End If
        n = n + 1
      Loop
  Else
    Exit Function
  End If
  Set GetNilaiPersediaan2 = vaArray
End Function

Function GetNilaiPersediaanKonsinyasi(ByVal cGolonganAwal As String, ByVal cGolonganAkhir As String, Optional ByVal hargabeli As Boolean = True, Optional ByVal tgl As Date, Optional ByVal lAllGolongan As Boolean = True, Optional ByVal lKodeStock As Boolean = True) As XArrayDB
Dim vaArray As New XArrayDB
Dim cField As String
Dim cWhere As String
Dim vaJoin
Dim n As Integer


  'Field array
  'Golongan
  'Nama Golongan
  'Kode
  'Nama
  'Satuan
  'Akhir
  'HP
  'Persediaan
  
  If hargabeli = True Then
    If lKodeStock = True Then
      cField = " s.kodegolongan,g.keterangan as namagolongan,s.kodestock,s.nama,s.kodesatuan,sum(ks.debet - ks.kredit) as akhir , s.cogs,sum(ks.debet - ks.kredit)* s.cogs as NilaiPersediaan "
    Else
      cField = " s.kodegolongan,g.keterangan as namagolongan,s.barcode,s.nama,s.kodesatuan,sum(ks.debet - ks.kredit) as akhir , s.cogs,sum(ks.debet - ks.kredit)* s.cogs as NilaiPersediaan "
    End If
  Else
    If lKodeStock = True Then
      cField = " s.kodegolongan,g.keterangan as namagolongan,s.kodestock,s.nama,s.kodesatuan,sum(ks.debet - ks.kredit) as akhir , s.cogs,sum(ks.debet - ks.kredit)* s.cogs as NilaiPersediaan "
    Else
      cField = " s.kodegolongan,g.keterangan as namagolongan,s.barcode,s.nama,s.kodesatuan,sum(ks.debet - ks.kredit) as akhir , s.cogs,sum(ks.debet - ks.kredit)* s.cogs as NilaiPersediaan "
    End If
  End If
  
  vaJoin = Array("LEFT JOIN golongan g on g.kodegolongan = s.kodegolongan", _
                 "LEFT JOIN kartustock ks on ks.kodestock = s.kodestock")
  If lAllGolongan = False Then
    cWhere = "s.kodegolongan >= '" & cGolonganAwal & "' AND  s.kodegolongan <= '" & cGolonganAkhir & "'"
  Else
    cWhere = " 1=1 "
  End If
  cWhere = cWhere & " AND ks.tgl <= '" & Format(tgl, "yyyy-MM-dd") & "'"
  cWhere = cWhere & " GROUP BY s.kodegolongan,s.kodestock"
  Set dbData = objData.Browse(GetDSN, "stock s", cField, , , , cWhere, "s.kodegolongan,s.kodestock", vaJoin)
  If dbData.RecordCount > 0 Then
      vaArray.LoadRows dbData.GetRows(dbData.RecordCount)
       n = 0
      Do While n <= vaArray.UpperBound(1)
        If vaArray(n, 5) = 0 Then
          vaArray.DeleteRows n
          n = n - 1
        End If
        n = n + 1
      Loop
  Else
    Exit Function
  End If
  Set GetNilaiPersediaanKonsinyasi = vaArray
End Function

Function GetHistoryHargaJual(ByVal cKodeStockAwal As String, ByVal cKodeStockAkhir As String, _
                             ByVal dTglAwal As Date, ByVal dTglAkhir As Date) As XArrayDB
Dim vaArray As New XArrayDB
Dim cField As String
Dim cWhere As String
Dim vaJoin

  'Field Array
  'Kode
  'Nama
  'Tgl
  'OldHJ
  'OldDiscount
  'HJ
  'Discount
  'Keterangan
  
  'Group : Kode
  
  cField = "c.Kode,s.Nama,c.Tgl,c.OldHj,c.OldDiscount,c.Hj,c.Discount,c.Keterangan"
  cWhere = "c.Kode >='" & cKodeStockAwal & "' And c.Kode <='" & cKodeStockAkhir & "'"
  cWhere = cWhere & " And c.Tgl >='" & Format(dTglAwal, "yyyy-MM-dd") & "' and c.Tgl <='" & Format(dTglAkhir, "yyyy-MM-dd") & "'"
  vaJoin = Array("Left Join Stock s on c.Kode = s.Kode")
  Set dbData = objData.Browse(GetDSN, "ChangePrice c", cField, , , , cWhere, "c.Kode,c.Tgl", vaJoin)
  If dbData.RecordCount > 0 Then
    vaArray.LoadRows dbData.GetRows(dbData.RecordCount)
    
    Set GetHistoryHargaJual = vaArray
  Else
    MsgBox "Data tidak ada", vbInformation
    Exit Function
  End If

End Function

Function GetHistoryHargaBeli(ByVal cKodeStockAwal As String, ByVal cKodeStockAkhir As String, _
                             ByVal dTglAwal As Date, ByVal dTglAkhir As Date) As XArrayDB
Dim vaArray As New XArrayDB
Dim cField As String
Dim cWhere As String
Dim vaJoin

  'Field Array
  'Kode
  'Nama
  'Tgl
  'OldHB
  'OldDiscount
  'HB
  'Discount
  'Keterangan
  
  'Group : Kode
  
  cField = "c.Kode,s.Nama,c.Tgl,c.OldHB,c.OldDiscount,c.HB,c.Discount,c.Keterangan"
  cWhere = "c.Kode >='" & cKodeStockAwal & "' And c.Kode <='" & cKodeStockAkhir & "'"
  cWhere = cWhere & " And c.Tgl >='" & Format(dTglAwal, "yyyy-MM-dd") & "' and c.Tgl <='" & Format(dTglAkhir, "yyyy-MM-dd") & "'"
  vaJoin = Array("Left Join Stock s on c.Kode = s.Kode")
  Set dbData = objData.Browse(GetDSN, "ChangePrice c", cField, , , , cWhere, "c.Kode,c.Tgl", vaJoin)
  If dbData.RecordCount > 0 Then
    vaArray.LoadRows dbData.GetRows(dbData.RecordCount)
    
    Set GetHistoryHargaBeli = vaArray
  Else
    MsgBox "Data tidak ada", vbInformation
    Exit Function
  End If
End Function

'...........................................
'PENJUALAN
'...........................................

Function GetRptPenjualanKeseluruhan(ByVal dTglAwal As Date, ByVal dTglAkhir As Date) As XArrayDB
Dim vaArray As New XArrayDB
Dim cField As String
Dim cWhere As String
Dim vaJoin

  '   vaArray(n, 0) = Nomor Faktur
  '   vaArray(n, 1) = Tanggal
  '   vaArray(n, 2) = Jatuh Tempo
  '   vaArray(n, 3) = Kode Customer
  '   vaArray(n, 4) = Nama Customer
  '   vaArray(n, 5) = Wilayah
  '   vaArray(n, 6) = Sub Total
  '   vaArray(n, 7) = Discount
  '   vaArray(n, 8) = Pajak
  '   vaArray(n, 9) = Total

  cField = "t.Faktur,t.Tgl,t.jthTmp,t.Customer,c.Nama as NamaCustomer,c.Wilayah,t.SubTotal,Sum(t.Discount + t.discount2) as Discount,t.Pajak,t.Total"
  vaJoin = Array("Left Join Customer c on c.Kode = t.Customer")
  cWhere = " t.Tgl >='" & Format(dTglAwal, "yyyy-MM-dd") & "' And t.Tgl <= '" & Format(dTglAkhir, "yyyy-MM-dd") & "'"
  cWhere = cWhere & " Group By t.Faktur"
  Set dbData = objData.Browse(GetDSN, "TotPenjualan t", cField, , , , cWhere, "t.Faktur", vaJoin)
  If dbData.RecordCount > 0 Then
    vaArray.LoadRows dbData.GetRows(dbData.RecordCount)
  End If
  Set GetRptPenjualanKeseluruhan = vaArray
End Function

Function GetReturPenjualanKeseluruhan(ByVal dTglAwal As Date, ByVal dTglAkhir As Date) As XArrayDB
Dim vaArray As New XArrayDB
Dim cField As String
Dim cWhere As String
Dim vaJoin

  'Field array
  'Faktur
  'Tgl
  'JthTmp
  'FktPenjualan
  'Customer
  'Nama
  'Wilayah
  'SubTotal
  'Discount
  'PPn
  'Total
  
  cField = "t.Faktur,t.Tgl,t.jthTmp,t.FktPenjualan,t.Customer,c.Nama as NamaCustomer,c.Wilayah,t.SubTotal,Sum(t.Discount + t.discount2) as Discount,t.PPn,t.Total"
  vaJoin = Array("Left Join Customer c on c.Kode = t.Customer")
  cWhere = " t.Tgl >='" & Format(dTglAwal, "yyyy-MM-dd") & "' And t.Tgl <= '" & Format(dTglAkhir, "yyyy-MM-dd") & "'"
  cWhere = cWhere & " Group By t.Faktur"
  Set dbData = objData.Browse(GetDSN, "TotRtnPenjualan t", cField, , , , cWhere, "t.Faktur", vaJoin)
  If dbData.RecordCount > 0 Then
    vaArray.LoadRows dbData.GetRows(dbData.RecordCount)
  End If
  Set GetReturPenjualanKeseluruhan = vaArray
End Function

Function GetPenjualanPerProduk(ByVal cKodeAwal As String, ByVal cKodeAkhir As String, ByVal dTglAwal As Date, ByVal dTglAkhir As Date) As XArrayDB
Dim vaArray As New XArrayDB
Dim n As Integer
Dim cField
Dim vaJoin
Dim cWhere
  
  'vaArray(n,0) = Kode
  'vaArray(n,1) = Nama Stock
  'vaarray(n,2) = Faktur
  'vaArray(n,3) = Tgl
  'vaArray(n,4) = QTy
  'vaArray(n,5) = Satuan
  'vaArray(n,6) = Harga
  'vaARray(n,7) = Bruto
  'vaArray(n,8) = Discount
  'vaArray(n,9) = Netto
  
  cField = "d.Kode,s.Nama,d.Faktur,d.Tgl,d.Qty,s.Satuan,"
  cField = cField & "d.Harga,d.Harga*d.Qty as Bruto,(d.Harga*d.Qty)-d.Jumlah as Discount,d.Jumlah"
  cWhere = "d.Kode >= '" & cKodeAwal & "' and d.Kode <= '" & cKodeAkhir & "' And"
  cWhere = cWhere & " d.Tgl >= '" & SisFormat(dTglAwal, Sis_yyyy_MM_dd) & "' and d.Tgl <= '" & SisFormat(dTglAkhir, Sis_yyyy_MM_dd) & "'"
  vaJoin = Array("Left Join Stock s on s.Kode = d.Kode")
  
  Set dbData = objData.Browse(GetDSN, "Penjualan d", cField, , , , cWhere, "d.Kode,d.Tgl,d.ID", vaJoin)
  If dbData.RecordCount > 0 Then
    vaArray.LoadRows dbData.GetRows
  End If
  Set GetPenjualanPerProduk = vaArray
End Function

Function GetPenjualanPerCustomer(ByVal cCustomerAwal As String, ByVal cCustomerAkhir As String, ByVal dTglAwal As Date, ByVal dTglAkhir As Date) As XArrayDB
Dim vaArray As New XArrayDB
Dim cWhere As String
Dim vaJoin
Dim cField As String

'     vaArray(n, 0) = Kode Customer
'     vaArray(n, 1) = Nama Customer
'     vaArray(n, 2) = Alamat
'     vaArray(n, 3) = Wilayah
'     vaArray(n, 4) = Nomor Faktur
'     vaArray(n, 5) = Tgl Penjualan
'     vaArray(n, 6) = Jatuh Tempo
'     vaArray(n, 7) = Sub Total
'     vaArray(n, 8) = Discount
'     vaArray(n, 9) = Pajak
'     vaArray(n, 10) = Total

    cField = "c.kode,c.Nama,c.Alamat,c.Wilayah,t.faktur,t.Tgl,t.JthTmp,t.subtotal,sum(t.discount + t.discount2) as discount,t.Pajak,t.total"
    cWhere = "t.customer >='" & cCustomerAwal & "'and t.Customer <='" & cCustomerAkhir & "'"
    cWhere = cWhere & " And t.tgl >= '" & Format(dTglAwal, "yyyy-MM-dd") & "' and t.Tgl <= '" & Format(dTglAkhir, "yyyy-MM-dd") & "'"
    cWhere = cWhere & " Group By t.Customer,t.Faktur"
    vaJoin = Array("Left join customer c on t.customer = c.kode")
    Set dbData = objData.Browse(GetDSN, "totpenjualan t", cField, , , , cWhere, "t.customer,t.Faktur", vaJoin)
    If dbData.RecordCount > 0 Then
      vaArray.LoadRows dbData.GetRows(dbData.RecordCount)
    End If
    Set GetPenjualanPerCustomer = vaArray
End Function

Function GetLabaKotor(ByVal dTglAwal As Date, ByVal dTglAkhir As Date, cHP As String) As XArrayDB
Dim vaArray As New XArrayDB
Dim cField As String
Dim cWhere As String
Dim vaJoin
Dim n
Dim nDiscount As Double
Dim nDiscount1 As Double
Dim nDiscount2 As Double
Dim nSubTotalHarga As Double

  '   vaArray(n, 0) = Nomor Faktur
  '   vaArray(n, 1) = Tanggal
  '   vaArray(n, 3) = Kode Stock
  '   vaArray(n, 4) = Nama Stock
  '   vaArray(n, 5) = Qty
  '   vaArray(n, 6) = Harga Beli
  '   vaArray(n, 7) = Harga Jual
  '   vaArray(n, 8) = totalHarga Beli
  '   vaArray(n, 9) = totalHarga Jual
  '   vaArray(n, 10) = Laba
  
  
  cField = "p.Faktur,p.Tgl,p.Kode,s.Nama,p.qty,s.satuan,p.harga,p.discount,t.persdisc,t.persdisc2,s.hb,s.hp"
  vaJoin = Array("left join stock s on s.kode = p.kode", _
                 "left join totpenjualan t on t.faktur = p.faktur")
                 
  cWhere = " p.Tgl >='" & Format(dTglAwal, "yyyy-MM-dd") & "' And p.Tgl <= '" & Format(dTglAkhir, "yyyy-MM-dd") & "'"
  Set dbData = objData.Browse(GetDSN, "Penjualan p", cField, , , , cWhere, "p.tgl,p.Faktur", vaJoin)
  If dbData.RecordCount > 0 Then
    vaArray.ReDim 0, dbData.RecordCount - 1, 0, 12
    Do While Not dbData.EOF
        vaArray(n, 0) = GetNull(dbData!Faktur)
        vaArray(n, 1) = Format(GetNull(dbData!tgl), "dd-MM-yyyy")
        vaArray(n, 2) = GetNull(dbData!Kode)
        vaArray(n, 3) = GetNull(dbData!nama)
        vaArray(n, 4) = GetNull(dbData!qty)
        vaArray(n, 5) = GetNull(dbData!Satuan)
        vaArray(n, 6) = GetNull(dbData!Harga)
        nSubTotalHarga = GetNull(dbData!qty) * GetNull(dbData!Harga)
        vaArray(n, 7) = GetNull(dbData!Discount)
        nDiscount = 1 - (GetNull(dbData!Discount) / 100)
        vaArray(n, 8) = GetNull(dbData!PersDisc)
        nDiscount1 = 1 - (GetNull(dbData!PersDisc) / 100)
        vaArray(n, 9) = GetNull(dbData!PersDisc2)
        nDiscount2 = 1 - (GetNull(dbData!PersDisc2) / 100)
        
        vaArray(n, 10) = nSubTotalHarga * nDiscount * nDiscount1 * nDiscount2
        
        If cHP = "H" Then
          vaArray(n, 11) = GetHPP(objData, GetNull(dbData!Kode), dbData!tgl, GetNull(dbData!hp)) * GetNull(dbData!qty)
        Else
          vaArray(n, 11) = GetNull(dbData!hb) * GetNull(dbData!qty)
        End If
        
        vaArray(n, 12) = vaArray(n, 10) - vaArray(n, 11)
        n = n + 1
      
      dbData.MoveNext
    Loop
    Set GetLabaKotor = vaArray
  End If
  
End Function

Private Function GetHPP(obj As CodeSuiteLibrary.Data, ByVal cKode As String, ByVal dTgl As Date, ByVal nSHP As Double) As Double
Dim db As New ADODB.Recordset
Dim cSQL
  GetHPP = nSHP
  
  cSQL = "Select HP From stkhp where Kode = '" & cKode & "' and Tgl <= '" & SisFormat(dTgl, Sis_yyyy_MM_dd) & "' "
  cSQL = cSQL & "Order by Kode,Tgl Desc "
  cSQL = cSQL & "Limit 0,1"
  Set db = objData.SQL(GetDSN, cSQL)
  If db.RecordCount > 0 Then
    GetHPP = GetNull(db!hp, 0)
  End If
End Function

Function GetRekapPenjualanJenisUsaha(ByVal cJenisUsahaAwal As String, ByVal cJenisUsahaAkhir As String, ByVal dTglAwal As Date, ByVal dTglAkhir As Date) As XArrayDB
Dim vaArray As New XArrayDB
Dim cWhere As String
Dim vaJoin
Dim cField As String
Dim n As Integer

'     vaArray(n, 0) = KodeJenisUsaha
'     vaArray(n, 1) = NamaJenisUsaha
'     vaArray(n, 4) = SubTotal
'     vaArray(n, 5) = Discount
'     vaArray(n, 6) = Pajak
'     vaArray(n, 7) = Total

    
    Set dbData = objData.Browse(GetDSN, "JenisUsaha", "Kode,Keterangan", "Kode", sisGTEqual, cJenisUsahaAwal, "And Kode <= '" & cJenisUsahaAkhir & "'", "Kode")
    If dbData.RecordCount > 0 Then
      vaArray.LoadRows dbData.GetRows(dbData.RecordCount)
    End If
        
    vaArray.ReDim 0, vaArray.UpperBound(1), 0, 6
    
    For n = 0 To vaArray.UpperBound(1)
      cWhere = "t.tgl >= '" & Format(dTglAwal, "yyyy-MM-dd") & "' and t.Tgl <= '" & Format(dTglAkhir, "yyyy-MM-dd") & "'"
      cWhere = cWhere & " And c.JenisUsaha >='" & vaArray(n, 0) & "'"
      cWhere = cWhere & " Group By c.JenisUsaha"
      cField = "sum(t.subtotal) as SubTotal,sum(t.discount) as Discount, sum(t.discount2) as discount2,sum(t.Pajak) as Pajak,sum(t.total) as Total"
      vaJoin = Array("Left join customer c on t.customer = c.kode", _
                     "Left Join JenisUsaha j on j.Kode = c.JenisUsaha")
      Set dbData = objData.Browse(GetDSN, "totpenjualan t", cField, , , , cWhere, "c.JenisUsaha", vaJoin)
      If dbData.RecordCount > 0 Then
        vaArray(n, 2) = GetNull(dbData!Subtotal)
        vaArray(n, 3) = GetNull(dbData!Discount) + GetNull(dbData!Discount2)
        vaArray(n, 4) = GetNull(dbData!PAJAK)
        vaArray(n, 5) = GetNull(dbData!Total)
      Else
        vaArray(n, 2) = 0
        vaArray(n, 3) = 0
        vaArray(n, 4) = 0
        vaArray(n, 5) = 0
      End If
    Next
    
    n = 0
    'Hapus baris yang Penjualannya kosong
    Do While n <= vaArray.UpperBound(1)
      If vaArray(n, 2) <= 0 Then
        vaArray.DeleteRows n
        n = n - 1
      End If
      n = n + 1
    Loop
    
    Set GetRekapPenjualanJenisUsaha = vaArray
End Function

'...........................................
'Pembelian
'...........................................

Function GetRptPembelianKeseluruhan(ByVal dTglAwal As Date, ByVal dTglAkhir As Date) As XArrayDB
Dim vaArray As New XArrayDB
Dim cField As String
Dim cWhere As String
Dim vaJoin

  '   vaArray(n, 0) = Nomor Faktur
  '   vaArray(n, 1) = Tanggal
  '   vaArray(n, 2) = Jatuh Tempo
  '   vaArray(n, 3) = Kode Customer
  '   vaArray(n, 4) = Nama Customer
  '   vaArray(n, 5) = Wilayah
  '   vaArray(n, 6) = Sub Total
  '   vaArray(n, 7) = Discount
  '   vaArray(n, 8) = Pajak
  '   vaArray(n, 9) = Total

  cField = "t.Faktur,t.Tgl,t.jthTmp,t.Supplier,c.Nama as NamaSupplier,c.Wilayah,t.SubTotal,Sum(t.Discount + t.discount2) as Discount,t.Pajak,t.Total"
  vaJoin = Array("Left Join Supplier c on c.Kode = t.Supplier")
  cWhere = " t.Tgl >='" & Format(dTglAwal, "yyyy-MM-dd") & "' And t.Tgl <= '" & Format(dTglAkhir, "yyyy-MM-dd") & "'"
  cWhere = cWhere & " Group By t.Faktur"
  Set dbData = objData.Browse(GetDSN, "TotPembelian t", cField, , , , cWhere, "t.Faktur", vaJoin)
  If dbData.RecordCount > 0 Then
    vaArray.LoadRows dbData.GetRows(dbData.RecordCount)
  End If
  Set GetRptPembelianKeseluruhan = vaArray
End Function

Function GetReturPembelianKeseluruhan(ByVal dTglAwal As Date, ByVal dTglAkhir As Date) As XArrayDB
Dim vaArray As New XArrayDB
Dim cField As String
Dim cWhere As String
Dim vaJoin

  'Field array
  'Faktur
  'Tgl
  'JthTmp
  'FktPenjualan
  'supplier
  'Nama
  'Wilayah
  'SubTotal
  'Discount
  'PPn
  'Total
  
  cField = "t.Faktur,t.Tgl,t.jthTmp,t.FktPembelian,t.Supplier,c.Nama as NamaSupplier,c.Wilayah,t.SubTotal,Sum(t.Discount + t.discount2) as Discount,t.PPn,t.Total"
  vaJoin = Array("Left Join Supplier c on c.Kode = t.Supplier")
  cWhere = " t.Tgl >='" & Format(dTglAwal, "yyyy-MM-dd") & "' And t.Tgl <= '" & Format(dTglAkhir, "yyyy-MM-dd") & "'"
  cWhere = cWhere & " Group By t.Faktur"
  Set dbData = objData.Browse(GetDSN, "TotRtnPembelian t", cField, , , , cWhere, "t.Faktur", vaJoin)
  If dbData.RecordCount > 0 Then
    vaArray.LoadRows dbData.GetRows(dbData.RecordCount)
  End If
  Set GetReturPembelianKeseluruhan = vaArray
End Function

Function GetPembelianPerProduk(ByVal cKodeAwal As String, ByVal cKodeAkhir As String, ByVal dTglAwal As Date, ByVal dTglAkhir As Date) As XArrayDB
Dim vaArray As New XArrayDB
Dim n As Integer
Dim cField
Dim vaJoin
Dim cWhere
  
  'vaArray(n,0) = Kode
  'vaArray(n,1) = Nama Stock
  'vaarray(n,2) = Faktur
  'vaArray(n,3) = Tgl
  'vaArray(n,4) = QTy
  'vaArray(n,5) = Satuan
  'vaArray(n,6) = Harga
  'vaARray(n,7) = Bruto
  'vaArray(n,8) = Discount
  'vaArray(n,9) = Netto
  
  cField = "d.Kode,s.Nama,d.Faktur,d.Tgl,d.Qty,s.Satuan,"
  cField = cField & "d.Harga,d.Harga*d.Qty as Bruto,(d.Harga*d.Qty)-d.Jumlah as Discount,d.Jumlah"
  cWhere = "d.Kode >= '" & cKodeAwal & "' and d.Kode <= '" & cKodeAkhir & "' And"
  cWhere = cWhere & " d.Tgl >= '" & SisFormat(dTglAwal, Sis_yyyy_MM_dd) & "' and d.Tgl <= '" & SisFormat(dTglAkhir, Sis_yyyy_MM_dd) & "'"
  vaJoin = Array("Left Join Stock s on s.Kode = d.Kode")
  
  Set dbData = objData.Browse(GetDSN, "Pembelian d", cField, , , , cWhere, "d.Kode,d.Tgl,d.ID", vaJoin)
  If dbData.RecordCount > 0 Then
    vaArray.LoadRows dbData.GetRows
  End If
  Set GetPembelianPerProduk = vaArray
End Function

Function GetPembelianPerSupplier(ByVal cSupplierAwal As String, ByVal cSupplierAkhir As String, ByVal dTglAwal As Date, ByVal dTglAkhir As Date) As XArrayDB
Dim vaArray As New XArrayDB
Dim cWhere As String
Dim vaJoin
Dim cField As String

'     vaArray(n, 0) = Kode Supplier
'     vaArray(n, 1) = Nama Supplier
'     vaArray(n, 2) = Nomor Faktur
'     vaArray(n, 3) = Tgl Penjualan
'     vaArray(n, 4) = Jatuh Tempo
'     vaArray(n, 5) = Sub Total
'     vaArray(n, 6) = Discount
'     vaArray(n, 7) = Pajak
'     vaArray(n, 8) = Total

    cField = "c.kode,c.Nama,t.faktur,t.Tgl,t.JthTmp,t.subtotal,sum(t.discount + t.discount2) as discount,t.Pajak,t.total"
    cWhere = "t.Supplier >='" & cSupplierAwal & "'and t.Supplier <='" & cSupplierAkhir & "' "
    cWhere = cWhere & " And t.tgl >= '" & Format(dTglAwal, "yyyy-MM-dd") & "' and t.Tgl <= '" & Format(dTglAkhir, "yyyy-MM-dd") & "'"
    cWhere = cWhere & " Group By t.Supplier,t.Faktur"
    vaJoin = Array("Left join Supplier c on t.Supplier = c.kode")
    Set dbData = objData.Browse(GetDSN, "totpembelian t", cField, , , , cWhere, "t.Supplier,t.Faktur", vaJoin)
    If dbData.RecordCount > 0 Then
      vaArray.LoadRows dbData.GetRows(dbData.RecordCount)
    End If
    Set GetPembelianPerSupplier = vaArray
End Function

Function GetPembelianPerJenisUsaha(ByVal cJenisUsahaAwal As String, ByVal cJenisUsahaAkhir As String, ByVal dTglAwal As Date, ByVal dTglAkhir As Date) As XArrayDB
Dim vaArray As New XArrayDB
Dim cWhere As String
Dim vaJoin
Dim cField As String

'     vaArray(n, 0) = KodeJenisUsaha
'     vaArray(n, 1) = NamaJenisUsaha
'     vaArray(n, 2) = Faktur
'     vaArray(n, 3) = Tgl
'     vaArray(n, 4) = KodeSupplier
'     vaArray(n, 5) = NamaSupplier
'     vaArray(n, 6) = JthTmp
'     vaArray(n, 7) = SubTotal
'     vaArray(n, 8) = Discount
'     vaArray(n, 9) = Pajak
'     vaArray(n, 10) = Total

    
    cField = "j.Kode as KodeJenisUsaha, j.Keterangan as NamaJenisUsaha,t.faktur,t.Tgl,c.kode as KodeSupplier,c.Nama as NamaSupplier,t.JthTmp,t.subtotal,(t.discount + t.discount2) as discount,t.Pajak,t.total"
    cWhere = "t.tgl >= '" & Format(dTglAwal, "yyyy-MM-dd") & "' and t.Tgl <= '" & Format(dTglAkhir, "yyyy-MM-dd") & "'"
    cWhere = cWhere & " And c.JenisUsaha >='" & cJenisUsahaAwal & "' and c.JenisUsaha <='" & cJenisUsahaAkhir & "'"
    vaJoin = Array("Left join Supplier c on t.Supplier = c.kode", _
                   "Left Join JenisUsaha j on j.Kode = c.JenisUsaha")
    Set dbData = objData.Browse(GetDSN, "totpembelian t", cField, , , , cWhere, "c.JenisUsaha,c.Kode", vaJoin)
    If dbData.RecordCount > 0 Then
      vaArray.LoadRows dbData.GetRows(dbData.RecordCount)
      
      Set GetPembelianPerJenisUsaha = vaArray
    End If
End Function

Function GetRekapPembelianJenisUsaha(ByVal cJenisUsahaAwal As String, ByVal cJenisUsahaAkhir As String, ByVal dTglAwal As Date, ByVal dTglAkhir As Date) As XArrayDB
Dim vaArray As New XArrayDB
Dim cWhere As String
Dim vaJoin
Dim cField As String
Dim n As Integer

'     vaArray(n, 0) = KodeJenisUsaha
'     vaArray(n, 1) = NamaJenisUsaha
'     vaArray(n, 4) = SubTotal
'     vaArray(n, 5) = Discount
'     vaArray(n, 6) = Pajak
'     vaArray(n, 7) = Total

    
    Set dbData = objData.Browse(GetDSN, "JenisUsaha", "Kode,Keterangan", "Kode", sisGTEqual, cJenisUsahaAwal, "And Kode <= '" & cJenisUsahaAkhir & "'", "Kode")
    If dbData.RecordCount > 0 Then
      vaArray.LoadRows dbData.GetRows(dbData.RecordCount)
    End If
        
    vaArray.ReDim 0, vaArray.UpperBound(1), 0, 6
    
    For n = 0 To vaArray.UpperBound(1)
      cWhere = "t.tgl >= '" & Format(dTglAwal, "yyyy-MM-dd") & "' and t.Tgl <= '" & Format(dTglAkhir, "yyyy-MM-dd") & "'"
      cWhere = cWhere & " And c.JenisUsaha >='" & vaArray(n, 0) & "'"
      cWhere = cWhere & " Group By c.JenisUsaha"
      cField = "sum(t.subtotal) as SubTotal,sum(t.discount) as Discount, sum(t.discount2) as discount2,sum(t.Pajak) as Pajak,sum(t.total) as Total"
      vaJoin = Array("Left join Supplier c on t.Supplier = c.kode", _
                     "Left Join JenisUsaha j on j.Kode = c.JenisUsaha")
      Set dbData = objData.Browse(GetDSN, "TotPembelian t", cField, , , , cWhere, "c.JenisUsaha", vaJoin)
      If dbData.RecordCount > 0 Then
        vaArray(n, 2) = GetNull(dbData!Subtotal)
        vaArray(n, 3) = GetNull(dbData!Discount) + GetNull(dbData!Discount2)
        vaArray(n, 4) = GetNull(dbData!PAJAK)
        vaArray(n, 5) = GetNull(dbData!Total)
      Else
        vaArray(n, 2) = 0
        vaArray(n, 3) = 0
        vaArray(n, 4) = 0
        vaArray(n, 5) = 0
      End If
    Next
    
    n = 0
    'Hapus baris yang Penjualannya kosong
    Do While n <= vaArray.UpperBound(1)
      If vaArray(n, 2) <= 0 Then
        vaArray.DeleteRows n
        n = n - 1
      End If
      n = n + 1
    Loop
    
    Set GetRekapPembelianJenisUsaha = vaArray
End Function

'========================================
'SALESMAN
'========================================

Function GetOmzetSalesman(ByVal cKodeAwalSales As String, ByVal cKodeAkhirSales As String, ByVal dTglAwal As Date, ByVal dTglAkhir As Date) As XArrayDB
Dim vaArray As New XArrayDB
Dim cField As String
Dim vaJoin As Variant
Dim cWhere As String
Dim n As Integer

    cWhere = "t.Salesman >='" & cKodeAwalSales & "' and t.Salesman <='" & cKodeAkhirSales & "'"
    cWhere = cWhere & " And t.tgl >= '" & Format(dTglAwal, "yyyy-MM-dd") & "' and t.Tgl <= '" & Format(dTglAkhir, "yyyy-MM-dd") & "'"
    cWhere = cWhere & " Group By t.Salesman,t.Faktur"
    cField = "c.kode,c.nama as NamaSalesman,c.Alamat,t.faktur,t.Customer,t.tgl,t.JthTmp,t.subtotal as SubTotal1,t.Pajak,t.discount as Discount1, t.discount2 as discount2,t.total as Total1,a.Nama as NamaCustomer"
    vaJoin = Array("Left join Salesman c on t.salesman =c.kode", _
                   "Left Join Customer a on a.Kode = t.Customer")
    Set dbData = objData.Browse(GetDSN, "totPenjualan t", cField, , , , cWhere, "t.Salesman,t.Faktur", vaJoin)
    If dbData.RecordCount > 0 Then
      n = 0
      vaArray.ReDim 0, dbData.RecordCount - 1, 0, 12
      Do While Not dbData.EOF
        vaArray(n, 0) = GetNull(dbData!Kode)
        vaArray(n, 1) = GetNull(dbData!namasalesman)
        vaArray(n, 2) = GetNull(dbData!Faktur)
        vaArray(n, 3) = GetNull(dbData!Customer)
        vaArray(n, 4) = GetNull(dbData!namacustomer)
        vaArray(n, 5) = GetNull(dbData!tgl)
        vaArray(n, 6) = GetNull(dbData!jthtmp)
        vaArray(n, 7) = GetNull(dbData!subtotal1)
        vaArray(n, 8) = GetNull(dbData!PAJAK)
        vaArray(n, 9) = GetNull(dbData!Discount1) + GetNull(dbData!Discount2)
        vaArray(n, 10) = GetNull(dbData!total1)
        dbData.MoveNext
        n = n + 1
      Loop
    End If
    Set GetOmzetSalesman = vaArray
End Function

Function GetRekapOmzetSalesman(ByVal cKodeAwalSales As String, ByVal cKodeAkhirSales As String, ByVal dTglAwal As Date, ByVal dTglAkhir As Date) As XArrayDB
Dim dbData1 As New ADODB.Recordset
Dim vaArray As New XArrayDB
Dim cSQL As String
Dim Disc As Double
Dim n As Integer

  cSQL = " Select t.Tgl,t.Faktur,t.Salesman as KodeSales,s.Nama,s.Alamat,s.Kota,Sum(t.SubTotal) as SubTotal1,"
  cSQL = cSQL & " Sum(t.Pajak) as Pajak,Sum(t.Discount)as Discount1, sum(t.Discount2) as Discount2, Sum(t.TOTAL) as TOTAL1"
  cSQL = cSQL & " From TotPenjualan t"
  cSQL = cSQL & " Left Join Salesman s on s.Kode=t.Salesman"
  cSQL = cSQL & " Where t.Salesman >= '" & cKodeAwalSales & "' And t.Salesman <= '" & cKodeAkhirSales & "'"
  cSQL = cSQL & " And t.Tgl >= '" & Format(dTglAwal, "yyyy-mm-dd") & "' And t.Tgl <='" & Format(dTglAkhir, "yyyy-mm-dd") & "'"
  cSQL = cSQL & " Group By t.Salesman"
  cSQL = cSQL & " Order By t.Salesman"
  Set dbData = objData.SQL(GetDSN, cSQL)
  If dbData.RecordCount > 0 Then
    n = 0
    vaArray.ReDim 0, dbData.RecordCount - 1, 0, 8
    Do While Not dbData.EOF
      vaArray(n, 0) = GetNull(dbData!KodeSales)
      vaArray(n, 1) = GetNull(dbData!nama)
      vaArray(n, 2) = GetNull(dbData!alamat)
      vaArray(n, 3) = GetNull(dbData!kota)
      vaArray(n, 4) = GetNull(dbData!subtotal1)
      vaArray(n, 5) = GetNull(dbData!Discount1) + GetNull(dbData!Discount2)
      vaArray(n, 6) = GetNull(dbData!PAJAK)
      vaArray(n, 7) = GetNull(dbData!total1)
      dbData.MoveNext
      n = n + 1
    Loop
  End If
  Set GetRekapOmzetSalesman = vaArray
End Function

'=======================================
'Giro
'=======================================

Function GetGiroMasuk(ByVal dTglAwal As Date, ByVal dTglAkhir As Date) As XArrayDB
Dim vaArray As New XArrayDB
Dim vaJoin
Dim cField As String
  
  'Field Array
  'Faktur
  'Tgl
  'Customer
  'NamaCustomer
  'NoBG
  'JthTm
  'Jumlah
  'bank
  cField = "c.Faktur,c.Tgl,c.Customer,s.nama,c.NoBG,c.JthTmp,c.Jumlah,c.Bank"
  vaJoin = Array("Left Join Customer s on c.Customer = s.Kode")
  Set dbData = objData.Browse(GetDSN, "CekIn c", cField, "c.Tgl", sisGTEqual, Format(dTglAwal, "yyyy-MM-dd"), "And c.Tgl <='" & Format(dTglAkhir, "yyyy-MM-dd") & "'", "c.faktur", vaJoin)
  If dbData.RecordCount > 0 Then
    vaArray.LoadRows dbData.GetRows(dbData.RecordCount)
  Else
    MsgBox "Data tidak ada ", vbInformation
    Exit Function
  End If
  Set GetGiroMasuk = vaArray
End Function

Function GetGiroKeluar(ByVal dTglAwal As Date, ByVal dTglAkhir As Date) As XArrayDB
Dim vaArray As New XArrayDB
Dim vaJoin
Dim cField As String
  
  'Field Array
  'Faktur
  'Tgl
  'Supplier
  'NamaSupplier
  'NoBG
  'JthTm
  'Jumlah
  'bank
  cField = "c.Faktur,c.Tgl,c.Supplier,s.nama,c.NoBG,c.JthTmp,c.Jumlah,c.Bank"
  vaJoin = Array("Left Join Supplier s on c.Supplier = s.Kode")
  Set dbData = objData.Browse(GetDSN, "Cekout c", cField, "c.Tgl", sisGTEqual, Format(dTglAwal, "yyyy-MM-dd"), "And c.Tgl <='" & Format(dTglAkhir, "yyyy-MM-dd") & "'", "c.faktur", vaJoin)
  If dbData.RecordCount > 0 Then
    vaArray.LoadRows dbData.GetRows(dbData.RecordCount)
  Else
    MsgBox "Data tidak ada ", vbInformation
    Exit Function
  End If
  Set GetGiroKeluar = vaArray
End Function

'=========================
'Customer
'=========================

Function GetDaftarCustomer(ByVal cWilayahAwal As String, ByVal cWilayahAkhir As String, ByVal cJenisUsahaAwal As String, ByVal cJenisUsahaAkhir As String, _
                            ByVal cCustomerAwal As String, ByVal cCustomerAkhir As String)

Dim vaArray As New XArrayDB
Dim n As Integer
Dim cField, cWhere As String
Dim vaJoin As Variant

  cField = "s.Kode,Nama,Alamat,Telepon,Wilayah,w.Keterangan as Area,j.kode as KodeJenisUsaha,j.Keterangan,s.Plafond1,s.Plafond2,s.Duedate"
  cWhere = "s.wilayah >='" & cWilayahAwal & "' and s.wilayah <= '" & cWilayahAkhir & "'"
  cWhere = cWhere & " And s.jenisusaha >='" & cJenisUsahaAwal & "' and s.jenisusaha <= '" & cJenisUsahaAkhir & "' "
  cWhere = cWhere & " And s.kode >='" & cCustomerAwal & "' and s.kode <='" & cCustomerAkhir & "'"
  vaJoin = Array("Left Join JenisUsaha j on j.Kode=s.JenisUsaha", " Left Join Wilayah w on w.Kode=s.Wilayah")
  
  Set dbData = objData.Browse(GetDSN, "Customer s", cField, , , , cWhere, "Wilayah,Kode", vaJoin)
  If dbData.RecordCount > 0 Then
    vaArray.ReDim 0, dbData.RecordCount - 1, 0, 10
    Do While Not dbData.EOF
      vaArray(n, 0) = GetNull(dbData!Wilayah)
      vaArray(n, 1) = GetNull(dbData!Area)
      vaArray(n, 2) = GetNull(dbData!Kode)
      vaArray(n, 3) = GetNull(dbData!nama)
      vaArray(n, 4) = GetNull(dbData!alamat)
      vaArray(n, 5) = GetNull(dbData!telepon)
      vaArray(n, 6) = GetNull(dbData!keterangan)
      vaArray(n, 7) = GetNull(dbData!Plafond1)
      vaArray(n, 8) = GetNull(dbData!Plafond2)
      vaArray(n, 9) = GetNull(dbData!DueDate)
     
      dbData.MoveNext
      n = n + 1
    Loop
  End If
  
  Set GetDaftarCustomer = vaArray
End Function

Function GetJatuhTempoPenjualan(ByVal cCustomerAwal As String, ByVal cCustomerAkhir As String, ByVal dJthTmpAwal As Date, ByVal dJthTmpAkhir As Date) As XArrayDB
Dim vaArray As New XArrayDB
Dim cFields As String
Dim cWhere As String
Dim n As Integer
Dim vaJoin As Variant

  'Field array
  'Kode Customer
  'Nama customer
  'Faktur
  'Tanggal
  'JthTmp
  'SubTotal
  'Discount
  'Pajak
  'Total

  cFields = "t.Customer,s.nama,t.faktur,Tgl,t.jthtmp,t.subtotal,(t.discount + t.Discount2) as Discount,t.pajak,t.total"
  cWhere = "t.JthTmp >= '" & Format(dJthTmpAwal, "yyyy-MM-dd") & "' and t.JthTmp <= '" & Format(dJthTmpAkhir, "yyyy-MM-dd") & "'  "
  cWhere = cWhere & " and t.Customer >= '" & cCustomerAwal & "' and t.customer <='" & cCustomerAkhir & "' "
  vaJoin = Array("left join Customer s on s.kode = t.Customer")
  Set dbData = objData.Browse(GetDSN, "totpenjualan t", cFields, , , , cWhere, "s.kode", vaJoin)
  If dbData.RecordCount > 0 Then
    vaArray.LoadRows dbData.GetRows(dbData.RecordCount)
    
    Set GetJatuhTempoPenjualan = vaArray
  Else
    MsgBox "Data tidak ada", vbInformation
    Exit Function
  End If
End Function


Function GetAccountStatement(ByVal cCustomer As String, ByVal dSampaiTanggal As Date) As XArrayDB
Dim vaArray As New XArrayDB
Dim n As Double
Dim nDay As Double
Dim cField As String
Dim cWhere As String
Dim n120 As Double
Dim n90 As Double
Dim n60 As Double
Dim n30 As Double
Dim n01 As Double
Dim nCurr As Double


  cField = "Tgl,Faktur,PO,JthTMP,Piutang,Sisa"
  cWhere = " and Tgl <= '" & Format(dSampaiTanggal, "yyyy-MM-dd") & "'"
  Set dbData = objData.Browse(GetDSN, "TotPenjualan", cField, "Customer", sisAssign, cCustomer, cWhere, "Customer,Tgl,Faktur")
  vaArray.ReDim 0, -1, 0, 6
  If dbData.RecordCount > 0 Then
    Do While Not dbData.EOF
      If GetNull(dbData!Sisa) > 0 Then
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        n = vaArray.UpperBound(1)
        nDay = DateDiff("D", dbData!jthtmp, Format(dSampaiTanggal, "yyyy-MM-dd"))
        Select Case nDay
          Case Is > 120
            n120 = n120 + GetNull(dbData!Sisa)
          Case Is > 90
            n90 = n90 + GetNull(dbData!Sisa)
          Case Is > 60
            n60 = n60 + GetNull(dbData!Sisa)
          Case Is > 30
            n30 = n30 + GetNull(dbData!Sisa)
          Case Is > 1
            n01 = n01 + GetNull(dbData!Sisa)
          Case Else
            nCurr = nCurr + GetNull(dbData!Sisa)
        End Select
        
        vaArray(n, 0) = GetNull(dbData!tgl)
        vaArray(n, 1) = GetNull(dbData!Faktur)
        vaArray(n, 2) = GetNull(dbData!PO)
        vaArray(n, 3) = ""
        vaArray(n, 4) = GetNull(dbData!jthtmp)
        vaArray(n, 5) = GetNull(dbData!Piutang)
        vaArray(n, 6) = GetNull(dbData!Sisa)
      End If
      dbData.MoveNext
    Loop
    Set GetAccountStatement = vaArray
  End If
End Function

Function GetAgingReportCustomer(ByVal cCustomerAwal As String, ByVal cCustomerAkhir As String, ByVal dSampaiTanggal As Date)
Dim vaArray As New XArrayDB
Dim n As Double
Dim nCol As Double
Dim vaJoin
Dim cField As String

  cField = "t.Customer,c.Nama as NamaCustomer,t.Tgl,t.Faktur,t.PO,t.JthTmp,t.Piutang,t.Sisa"
  vaJoin = Array("Left Join Customer c on t.Customer = c.Kode")
  Set dbData = objData.Browse(GetDSN, "TotPenjualan t", cField, "t.Customer", sisGTEqual, cCustomerAwal, " and t.Customer <= '" & cCustomerAkhir & "' and t.Tgl <= '" & Format(dSampaiTanggal, "yyyy-MM-dd") & "'", "t.Customer,t.Faktur,t.Tgl", vaJoin)
  vaArray.ReDim 0, -1, 0, 13
  If dbData.RecordCount > 0 Then
    Do While Not dbData.EOF
      If GetNull(dbData!Sisa) > 0 Then
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        n = vaArray.UpperBound(1)
        vaArray(n, 8) = 0
        vaArray(n, 9) = 0
        vaArray(n, 10) = 0
        vaArray(n, 11) = 0
        vaArray(n, 12) = 0
        vaArray(n, 13) = 0
        
        nCol = GetCol(dbData!jthtmp, dSampaiTanggal) + 8
        vaArray(n, 0) = GetNull(dbData!Customer)
        vaArray(n, 1) = GetNull(dbData!namacustomer)
        vaArray(n, 2) = GetNull(dbData!tgl)
        vaArray(n, 3) = GetNull(dbData!Faktur)
        vaArray(n, 4) = GetNull(dbData!PO)
        vaArray(n, 5) = ""
        vaArray(n, 6) = GetNull(dbData!jthtmp)
        vaArray(n, 7) = GetNull(dbData!Sisa)
        vaArray(n, nCol) = GetNull(vaArray(n, nCol)) + GetNull(dbData!Sisa)
      End If
      dbData.MoveNext
    Loop
    Set GetAgingReportCustomer = vaArray
  End If
End Function

Private Function GetCol(ByVal dJatuhTempo As Date, ByVal dTanggal As Date) As Double
Dim nDay As Double
    nDay = DateDiff("D", dJatuhTempo, dTanggal)
    Select Case nDay
      Case Is > 120
        GetCol = 4
      Case Is > 90
        GetCol = 3
      Case Is > 60
        GetCol = 2
      Case Is > 30
        GetCol = 2
      Case Is >= 1
        GetCol = 1
      Case Else
        GetCol = 0
    End Select
End Function

'=======================
'Supplier
'=======================

Function GetJatuhTempoPembelian(ByVal cSupplierAwal As String, ByVal cSupplierAkhir As String, ByVal dJthTmpAwal As Date, ByVal dJthTmpAkhir As Date) As XArrayDB
Dim vaArray As New XArrayDB
Dim cFields As String
Dim cWhere As String
Dim n As Integer
Dim vaJoin As Variant

  'Field array
  'Kode Supplier
  'Nama Supplier
  'Faktur
  'Tanggal
  'JthTmp
  'SubTotal
  'Discount
  'Pajak
  'Total

  cFields = "t.Supplier,s.nama,t.faktur,Tgl,t.jthtmp,t.subtotal,(t.discount + t.Discount2) as Discount,t.pajak,t.total"
  cWhere = "t.JthTmp >= '" & Format(dJthTmpAwal, "yyyy-MM-dd") & "' and t.JthTmp <= '" & Format(dJthTmpAkhir, "yyyy-MM-dd") & "'  "
  cWhere = cWhere & " and t.Supplier >= '" & cSupplierAwal & "' and t.Supplier <='" & cSupplierAkhir & "' "
  vaJoin = Array("left join Supplier s on s.kode = t.Supplier")
  Set dbData = objData.Browse(GetDSN, "totpembelian t", cFields, , , , cWhere, "s.kode", vaJoin)
  If dbData.RecordCount > 0 Then
    vaArray.LoadRows dbData.GetRows(dbData.RecordCount)
    
    Set GetJatuhTempoPembelian = vaArray
  Else
    MsgBox "Data tidak ada", vbInformation
    Exit Function
  End If
End Function

Function GetAccountStatementSupplier(ByVal cSupplier As String, ByVal dSampaiTgl As Date) As XArrayDB
Dim vaArray As New XArrayDB
Dim n As Double
Dim nDay As Double
Dim cField As String
Dim n120 As Double
Dim n90 As Double
Dim n60 As Double
Dim n30 As Double
Dim n01 As Double
Dim nCurr As Double

  cField = "Supplier,Tgl,Faktur,JthTmp,Hutang,Sisa"
  Set dbData = objData.Browse(GetDSN, "TotPembelian", cField, "Supplier", sisAssign, cSupplier, " and Tgl <= '" & Format(dSampaiTgl, "yyyy-MM-dd") & "'", "Supplier,Tgl,Faktur")
  vaArray.ReDim 0, -1, 0, 6
  If dbData.RecordCount > 0 Then
    Do While Not dbData.EOF
      If GetNull(dbData!Sisa) > 0 Then
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        n = vaArray.UpperBound(1)
        nDay = DateDiff("D", dbData!jthtmp, Format(dSampaiTgl, "yyyy-MM-dd"))
        Select Case nDay
          Case Is > 120
            n120 = n120 + GetNull(dbData!Sisa)
          Case Is > 90
            n90 = n90 + GetNull(dbData!Sisa)
          Case Is > 60
            n60 = n60 + GetNull(dbData!Sisa)
          Case Is > 30
            n30 = n30 + GetNull(dbData!Sisa)
          Case Is > 1
            n01 = n01 + GetNull(dbData!Sisa)
          Case Else
            nCurr = nCurr + GetNull(dbData!Sisa)
        End Select
        
        vaArray(n, 0) = GetNull(dbData!tgl)
        vaArray(n, 1) = GetNull(dbData!Faktur)
        vaArray(n, 2) = ""
        vaArray(n, 3) = GetNull(dbData!jthtmp)
        vaArray(n, 4) = GetNull(dbData!hutang)
        vaArray(n, 5) = GetNull(dbData!Sisa)
      End If
      dbData.MoveNext
    Loop
    Set GetAccountStatementSupplier = vaArray
  End If
End Function

Function GetAgingReportSupplier(ByVal cSupplierAwal As String, ByVal cSupplierAkhir As String, ByVal dSampaiTgl As Date) As XArrayDB
Dim vaArray As New XArrayDB
Dim n As Double
Dim nCol As Double

  Set dbData = objData.Browse(GetDSN, "TotPembelian t", "t.Supplier,c.Nama as NamaSupplier,t.Tgl,t.Faktur,t.JthTmp,t.Hutang,t.Sisa", "t.Supplier", sisGTEqual, cSupplierAwal, " and t.Supplier <= '" & cSupplierAkhir & "' and t.Tgl <= '" & Format(dSampaiTgl, "yyyy-MM-dd") & "'", "t.Supplier,t.Tgl,t.Faktur", _
               Array("Left Join Supplier c on t.Supplier = c.Kode"))
  vaArray.ReDim 0, -1, 0, 13
  If dbData.RecordCount > 0 Then
    Do While Not dbData.EOF
      If GetNull(dbData!Sisa) > 0 Then
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        n = vaArray.UpperBound(1)
        vaArray(n, 7) = 0
        vaArray(n, 8) = 0
        vaArray(n, 9) = 0
        vaArray(n, 10) = 0
        vaArray(n, 11) = 0
        vaArray(n, 12) = 0
        
        nCol = GetCol(dbData!jthtmp, Format(dSampaiTgl, "yyyy-MM-dd")) + 8
        vaArray(n, 0) = GetNull(dbData!supplier)
        vaArray(n, 1) = GetNull(dbData!namasupplier)
        vaArray(n, 2) = GetNull(dbData!tgl)
        vaArray(n, 3) = GetNull(dbData!Faktur)
        vaArray(n, 4) = ""
        vaArray(n, 5) = GetNull(dbData!jthtmp)
        vaArray(n, 6) = GetNull(dbData!Sisa)
        vaArray(n, nCol) = GetNull(vaArray(n, nCol)) + GetNull(dbData!Sisa)
      End If
      dbData.MoveNext
    Loop
   Set GetAgingReportSupplier = vaArray
  End If
End Function

Function GetHutangLain(ByVal dTglAwal As Date, ByVal dTglAkhir As Date) As XArrayDB
Dim vaArray As New XArrayDB
Dim vaJoin
Dim cWhere As String
Dim cField As String

    cField = "h.Faktur,h.Tgl,h.Supplier,s.Nama,h.JthTmp,h.Keterangan,h.bank,h.Total"
    cWhere = "h.Tgl >='" & Format(dTglAwal, "yyyy-MM-dd") & "' And h.Tgl <='" & Format(dTglAkhir, "yyyy-MM-dd") & "'"
    vaJoin = Array("Left Join Supplier s on s.Kode = h.Supplier")
    Set dbData = objData.Browse(GetDSN, "HutangLain h", cField, , , , cWhere, "h.Tgl,h.Faktur,h.Supplier", vaJoin)
    If dbData.RecordCount > 0 Then
      vaArray.LoadRows dbData.GetRows(dbData.RecordCount)
      
      Set GetHutangLain = vaArray
    Else
      MsgBox "Data tidak ada", vbInformation
      Exit Function
    End If
    
End Function

Function GetPiutangLain(ByVal dTglAwal As Date, ByVal dTglAkhir As Date) As XArrayDB
Dim vaArray As New XArrayDB
Dim vaJoin
Dim cWhere As String
Dim cField As String

    cField = "h.Faktur,h.Tgl,h.Customer,s.Nama,h.JthTmp,h.Keterangan,h.bank,h.Total"
    cWhere = "h.Tgl >='" & Format(dTglAwal, "yyyy-MM-dd") & "' And h.Tgl <='" & Format(dTglAkhir, "yyyy-MM-dd") & "'"
    vaJoin = Array("Left Join Customer s on s.Kode = h.Customer")
    Set dbData = objData.Browse(GetDSN, "PiutangLain h", cField, , , , cWhere, "h.Tgl,h.Faktur,h.Customer", vaJoin)
    If dbData.RecordCount > 0 Then
      vaArray.LoadRows dbData.GetRows(dbData.RecordCount)
      
      Set GetPiutangLain = vaArray
    Else
      MsgBox "Data tidak ada", vbInformation
      Exit Function
    End If
    
End Function

Function GetMutasiKasBank(ByVal dTglAwal As Date, ByVal dTglAkhir As Date) As XArrayDB
Dim vaArray As New XArrayDB
Dim cField As String
Dim vaJoin
Dim n As Integer

  cField = "m.Faktur,m.Tgl,m.Dari,a.Keterangan as namadari,m.Ke,b.Keterangan as NamaKe,m.Jumlah,m.Keterangan"
  vaJoin = Array("Left Join Bank a on m.Dari = a.Kode", _
                "Left Join Bank b on m.Ke = b.Kode")
  Set dbData = objData.Browse(GetDSN, "MutasiKasBank m", cField, "m.Tgl", sisGTEqual, Format(dTglAwal, "yyyy-MM-dd"), "And m.Tgl <='" & Format(dTglAkhir, "yyyy-MM-dd") & "'", "m.Tgl,m.Faktur", vaJoin)
  If dbData.RecordCount > 0 Then
    vaArray.ReDim 0, dbData.RecordCount - 1, 0, 7
    n = 0
    Do While Not dbData.EOF
      vaArray(n, 0) = GetNull(dbData!Faktur)
      vaArray(n, 1) = GetNull(dbData!tgl)
      vaArray(n, 2) = GetNull(dbData!DARI)
      vaArray(n, 3) = GetNull(dbData!NamaDari)
      vaArray(n, 4) = GetNull(dbData!Ke)
      vaArray(n, 5) = GetNull(dbData!NamaKe)
      vaArray(n, 6) = GetNull(dbData!jumlah)
      vaArray(n, 7) = GetNull(dbData!keterangan)
      dbData.MoveNext
      n = n + 1
    Loop
    
    Set GetMutasiKasBank = vaArray
  Else
    MsgBox "Data Tidak ada !", vbInformation
    Exit Function
  End If
End Function

Function GetRptKartukredit(ByVal cKartuAwal As String, ByVal cKartuAkhir As String, _
                           ByVal dTglAwal As Date, ByVal dTglAkhir As Date) As XArrayDB
Dim vaArray As New XArrayDB
Dim cField As String
Dim vaJoin
Dim cWhere As String
  
  cField = "t.Kartu,k.Keterangan as NamaKartu,t.Faktur,"
  cField = cField & "t.Tgl,'Penjualan Kasir' as Keterangan,"
  cField = cField & "t.NoKartu,t.NoTrace,t.BayarKartu,t.Administrasi"
  vaJoin = Array("Left Join Kartukredit k on t.Kartu = k.Kode")
  cWhere = "t.Kartu >= '" & cKartuAwal & "' and t.Kartu <= '" & cKartuAkhir & "' and "
  cWhere = cWhere & "t.Tgl >= '" & SisFormat(dTglAwal, Sis_yyyy_MM_dd) & "' and t.Tgl <= '" & SisFormat(dTglAkhir, Sis_yyyy_MM_dd) & "'"
  Set dbData = objData.Browse(GetDSN, "TotKasir t", cField, , , , cWhere, "t.Kartu,t.Tgl,t.Faktur", vaJoin)
  If dbData.RecordCount > 0 Then
    vaArray.LoadRows dbData.GetRows
  End If
  
  Set GetRptKartukredit = vaArray
End Function

Function GetRptSaldoKasBank(ByVal cBank As String, _
                            ByVal dTglAwal As Date, ByVal dTglAkhir As Date) As XArrayDB
Dim vaArray As New XArrayDB
Dim cField
Dim cWhere As String
Dim n As Double
Dim nSaldo As Double

  ' vaArray(n,0) = Tgl
  ' vaarray(n,1) = Keterangan
  ' vaArray(n,2) = Faktur
  ' vaArray(n,3) = debet
  ' vaArray(n,4) = kredit
  ' vaArray(n,5) = Saldo
  
  ' Ambil Data Saldo Awal Kas
  Set dbData = objData.Browse(GetDSN, "SaldoKasBank", "Sum(debet-kredit) as Awal", "Bank", sisAssign, cBank, _
               " and Tgl < '" & SisFormat(dTglAwal, Sis_yyyy_MM_dd) & "'")
  If dbData.RecordCount > 0 Then
    nSaldo = GetNull(dbData!AWAL)
  End If
  
  vaArray.ReDim 0, 0, 0, 5
  vaArray(n, 1) = "Saldo Awal"
  vaArray(n, 5) = nSaldo
  
  cField = "s.Tgl,s.Keterangan,s.Faktur,s.debet,s.kredit,0 as Saldo"
  cWhere = "s.Bank = '" & cBank & "' and "
  cWhere = cWhere & "s.Tgl >= '" & SisFormat(dTglAwal, Sis_yyyy_MM_dd) & "' and s.Tgl <= '" & SisFormat(dTglAkhir, Sis_yyyy_MM_dd) & "'"
  Set dbData = objData.Browse(GetDSN, "SaldoKasBank s", cField, , , , cWhere, "s.Bank,s.Tgl,s.ID")
  If dbData.RecordCount > 0 Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      
      nSaldo = nSaldo + GetNull(dbData!debet) - GetNull(dbData!kredit)
      vaArray(n, 0) = GetNull(dbData!tgl)
      vaArray(n, 1) = GetNull(dbData!keterangan)
      vaArray(n, 2) = GetNull(dbData!Faktur)
      vaArray(n, 3) = GetNull(dbData!debet)
      vaArray(n, 4) = GetNull(dbData!kredit)
      vaArray(n, 5) = nSaldo
      dbData.MoveNext
    Loop
  End If
  Set GetRptSaldoKasBank = vaArray
End Function

Function GetRptHPP(ByVal dTglAwal As Date, ByVal dTglAkhir As Date) As Double
Dim vaSaldo As New XArrayDB
Dim nAwal As Double
Dim nMutasi As Double
Dim n As Double
  
  vaSaldo.ReDim 0, -1, 0, 3
  Set dbData = objData.Browse(GetDSN, "kartustock t", "t.Kode,t.debet,s.hp", "t.status", sisAssign, SisKartuStock.saldoawal, , , _
               Array("Left Join Stock s on t.Kode = s.Kode"))
  If dbData.RecordCount > 0 Then
    Do While Not dbData.EOF
      InsertRow vaSaldo, GetNull(dbData!Kode), GetNull(dbData!debet), , GetNull(dbData!hp)
      dbData.MoveNext
    Loop
  End If
  
  Set dbData = objData.Browse(GetDSN, "KartuStock t", "t.Tgl,t.Kode,t.debet-kredit as Qty,s.HP", "Tgl", sisLTEqual, SisFormat(dTglAkhir, Sis_yyyy_MM_dd), , , _
               Array("Left Join Stock s on t.Kode = s.Kode"))
  If dbData.RecordCount > 0 Then
    Do While Not dbData.EOF
      If dbData!tgl < dTglAwal Then
        nAwal = GetNull(dbData!qty)
        nMutasi = 0
      Else
        nMutasi = GetNull(dbData!qty)
        nAwal = 0
      End If
      InsertRow vaSaldo, GetNull(dbData!Kode), GetNull(nAwal), GetNull(nMutasi), GetNull(dbData!hp)
      dbData.MoveNext
    Loop
  End If
  
  nAwal = 0
  nMutasi = 0
  For n = 0 To vaSaldo.UpperBound(1) Step 1
    nAwal = nAwal + (vaSaldo(n, 1) * vaSaldo(n, 3))
    nMutasi = nMutasi + (vaSaldo(n, 2) * vaSaldo(n, 3))
  Next
  
  GetRptHPP = nMutasi - nAwal
End Function

Private Sub InsertRow(vaSaldo As XArrayDB, ByVal cKode As String, _
            Optional ByVal nAwal As Double = 0, Optional ByVal nMutasi As Double = 0, Optional ByVal nHP As Double = 0)
Dim n As Double
  n = -1
  If vaSaldo.UpperBound(1) >= 0 Then
    n = vaSaldo.Find(0, 0, cKode)
  End If
  If n < 0 Then
    vaSaldo.InsertRows vaSaldo.UpperBound(1) + 1
    n = vaSaldo.UpperBound(1)
  End If
  
  vaSaldo(n, 0) = cKode
  vaSaldo(n, 1) = GetNull(vaSaldo(n, 1)) + nAwal
  vaSaldo(n, 2) = GetNull(vaSaldo(n, 2)) + nAwal + nMutasi
  vaSaldo(n, 3) = nHP
End Sub
