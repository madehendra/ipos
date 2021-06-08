Attribute VB_Name = "LabaRugi"
Option Explicit

Dim vaRpt As New XArrayDB
Dim objData As New CodeSuiteLibrary.Data
Dim dbData As New ADODB.Recordset
Dim vaArray As New XArrayDB

Function GetLabaRugiNetto(ByVal obj As CodeSuiteLibrary.Data, ByVal dAwal As Date, ByVal dAkhir As Date, ByVal lPreview As Boolean) As Double
Dim n As Integer
Dim nCount As Integer
Dim nNext As Integer
Dim nKe As Integer
Dim nTotalBiaya As Double

Dim nPenjualan   As Double
Dim nPotonganPenjualan As Double
Dim nReturPenjualan As Double
Dim nDiscountPelunasan As Double

Dim nPembelian As Double
Dim nPotonganPembelian As Double
Dim nPotonganTambahanPembelian As Double
Dim nReturPembelian As Double
Dim nStockAwal As Double
Dim nStockAkhir As Double
Dim nLabaRugiUsaha As Double
Dim nLabaRugiSimpanPinjam As Double
Dim nLB As Double
Dim nPenjualanCash As Double
Dim nLabaAkhir As Double
Dim nLabaBersih As Double
Dim nPenjualanAngsuran As Double
Dim nPotonganAngsuran As Double

Dim nTotalPenjualan As Double
Dim nTotalPembelian As Double
    
  nCount = 0
  nNext = 0
  nTotalBiaya = 0
  nLB = 0
  
  vaRpt.ReDim 0, 100, 0, 5
  
  '[PENJUALAN CASH/KASIR]
  nPenjualan = 0
  nPotonganPenjualan = 0
  GetPenjualan obj, dAwal, dAkhir, nPenjualan, nPotonganPenjualan
  nPenjualanCash = nPenjualan
  
  GetVarpt 0, "I", "Penjualan Cash", "", , , nPenjualan
  GetVarpt 1, , GetSpasi(1) & "Disc.", "", , , nPotonganPenjualan
  GetVarpt 2, , "Penjualan Bersih", "", , , nPenjualan + GetAbsMin(nPotonganPenjualan)
  GetBarisKosong 3
  
  '[PENJUALAN KREDIT/PIUTANG]
  nPenjualan = 0
  nPotonganPenjualan = 0
  nReturPenjualan = GetReturpenjualan(obj, dAwal, dAkhir)
  nDiscountPelunasan = GetDiscountPelunasan(obj, dAwal, dAkhir)
  GetPenjualanKredit obj, dAwal, dAkhir, nPenjualan, nPotonganPenjualan
  
  GetVarpt 4, "II", "Penjualan Kredit (Before Tax)", "", , , nPenjualan
  GetVarpt 5, , GetSpasi(1) & "Retur Penjualan", "", , , nReturPenjualan
  GetVarpt 6, , GetSpasi(1) & "Disc.", "", , , nPotonganPenjualan
  GetVarpt 7, , GetSpasi(1) & "Disc. Tambahan", "", , , nDiscountPelunasan
  GetVarpt 8, , "Penjualan Bersih", "", , , nPenjualan + GetAbsMin(nReturPenjualan) + GetAbsMin(nPotonganPenjualan) + GetAbsMin(nDiscountPelunasan)
  GetBarisKosong 9
    
  nTotalPenjualan = nPenjualanCash + nPenjualan + GetAbsMin(nPotonganPenjualan) + GetAbsMin(nReturPenjualan) + GetAbsMin(nDiscountPelunasan)
  
  '[PEMBELIAN]
  nPembelian = 0
  nPotonganPembelian = 0
  nStockAwal = 0
  nStockAkhir = 0
  nPotonganTambahanPembelian = 0
  
  Getpembelian obj, dAwal, dAkhir, nPembelian, nPotonganPembelian
  nPotonganTambahanPembelian = GetDiscountPelunasanHutang(obj, dAwal, dAkhir)
  nReturPembelian = GetReturpembelian(obj, dAwal, dAkhir)
  nStockAwal = GetNilaiStockAwal(obj, dAwal)
  nStockAkhir = GetNilaiStockAkhir(obj, dAkhir)
  
  GetVarpt 10, "III", "Stock Awal", , nStockAwal, ""
  GetVarpt 11, "IV", "Stock Akhir", , nStockAkhir, ""
  GetBarisKosong 16
  
  Dim nTotPembelian As Double
  Dim nTotPembelianKonsinyasi As Double
  
  nTotPembelian = nPembelian + GetAbsMin(nPotonganPembelian) + GetAbsMin(nReturPembelian) + GetAbsMin(nPotonganTambahanPembelian)
  
  GetVarpt 12, "V", "Pembelian (Before Tax)", , nPembelian, ""
  GetVarpt 13, , GetSpasi(1) & "Disc.", , nPotonganPembelian, ""
  GetVarpt 14, , GetSpasi(1) & "Retur Pembelian", , nReturPembelian, ""
  GetVarpt 15, , GetSpasi(1) & "Disc. Tambahan", , nPotonganTambahanPembelian, ""
  GetVarpt 16, , " Pembelian Bersih", , nTotPembelian, ""
  GetBarisKosong 17
  
  nTotalPembelian = nStockAwal + GetAbsMin(nStockAkhir) + nTotPembelian
  
  nLabaRugiUsaha = nTotalPenjualan + GetAbsMin(nTotalPembelian)
  
  GetVarpt 18, "VI", "HPP", "", , , nTotalPembelian
  GetVarpt 19, "VII", "Laba/Rugi Usaha", "", , , nLabaRugiUsaha
  GetBarisKosong 20
  
  nLabaAkhir = nLabaRugiUsaha + GetKasMasuk(obj, dAwal, dAkhir) + GetAbsMin(GetKasKeluar(obj, dAwal, dAkhir)) - GetKomisiPenjualan(obj, dAwal, dAkhir)
  
  GetVarpt 21, "VIII", "Kas Masuk", , GetKasMasuk(obj, dAwal, dAkhir), ""
  GetVarpt 22, "VIX", "Kas Keluar", "", , , GetKasKeluar(obj, dAwal, dAkhir)
  GetVarpt 23, "X", "Komisi Penjualan", "", , , GetKomisiPenjualan(obj, dAwal, dAkhir)
  GetVarpt 24, "XI", "Laba/rugi Bersih", "", , , nLabaAkhir
  GetBarisKosong 25
  
  If lPreview = True Then
    GetPreview dAwal, dAkhir
  End If
  GetLabaRugiNetto = nLabaAkhir
End Function

Private Function GetSpasi(Optional ByVal nNumber As Integer = 1) As String
Dim n As Integer

  GetSpasi = ""
  For n = 1 To nNumber
    GetSpasi = GetSpasi & vbTab
  Next
End Function

Private Sub GetVarpt(ByVal nBaris As Integer, _
                     Optional nKol1 As String = "", _
                     Optional nKol2 As String = "", _
                     Optional nKol3 As String = "Rp", _
                     Optional nKol4 As Double = 0, _
                     Optional nKol5 As String = "Rp", _
                     Optional nKol6 As Double = 0)

  vaRpt(nBaris, 0) = nKol1
  vaRpt(nBaris, 1) = nKol2
  vaRpt(nBaris, 2) = nKol3
  vaRpt(nBaris, 3) = nKol4
  vaRpt(nBaris, 4) = nKol5
  vaRpt(nBaris, 5) = nKol6
End Sub

Private Sub GetBarisKosong(ByVal nBaris As Integer)
  vaRpt(nBaris, 0) = " "
  vaRpt(nBaris, 1) = " "
  vaRpt(nBaris, 2) = " "
  vaRpt(nBaris, 3) = " "
  vaRpt(nBaris, 4) = " "
  vaRpt(nBaris, 5) = " "
End Sub

Private Sub GetPreview(ByVal dAwal As Date, ByVal dAkhir As Date)
  With FrmRPT
    .AddPageHeader "LAPORAN LABA/RUGI", tdbHalignCenter, , , , , 12, True
    .AddPageHeader aCfg(objData, msNamaPerusahaan, ""), tdbHalignCenter, , , True, , 14, True
    .AddPageHeader "Periode : " & Format(dAwal, "dd-mm-yyyy") & " s.d " & Format(dAkhir, "dd-mm-yyyy"), tdbHalignCenter, , , True, , 10, True, True
    .AddPageHeader " ", , , , True
    .AddPageHeader " ", , , , True
    
    .AddTableHeader "", , , , 4, , , , , , , , , , , , , , , , , db_None, db_None, db_None, db_None, db_None, db_None
    .AddTableHeader "", , , , , , , , , , , , , , , , , , , , , db_None, db_None, db_None, db_None, db_None, db_None
    .AddTableHeader "", , , , 4, , , , , , , , , , , , , , , , , db_None, db_None, db_None, db_None, db_None, db_None
    .AddTableHeader "", , , , 20, , , , , , , , , , , , , , , , , db_None, db_None, db_None, db_None, db_None, db_None
    .AddTableHeader "", , , , 4, , , , , , , , , , , , , , , , , db_None, db_None, db_None, db_None, db_None, db_None
    .AddTableHeader "", , , , 20, , , , , , , , , , , , , , , , , db_None, db_None, db_None, db_None, db_None, db_None

    .AddTableBody , , , , , , , , , , , , , , , , , db_None, db_None, db_None, db_None, db_None, db_None
    .AddTableBody , , , , , , , , , , , , , , , , , db_None, db_None, db_None, db_None, db_None, db_None
    .AddTableBody , , , , , , , , , , , , , , , , , db_None, db_None, db_None, db_None, db_None, db_None
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight, , , , , , , , , , , , , , , , db_None, db_None, db_None, db_None, db_None, db_None
    .AddTableBody , , , , , , , , , , , , , , , , , db_None, db_None, db_None, db_None, db_None, db_None
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight, , , , , , , , , , , , , , , , db_None, db_None, db_None, db_None, db_None, db_None
    
    .Preview vaRpt, , False
  End With
End Sub

Private Function GetHuruf(ByVal nKe As Integer)
Dim vaHuruf

  vaHuruf = Array("a", "b", "c", "d", "e", "f")
  GetHuruf = vaHuruf(nKe)
End Function

Private Function GetMin(nValue)
Dim cChar As String

    cChar = IIf(nValue < 0, "()", "  ")
    nValue = Format(Abs(GetNull(nValue)), "###,###,###,###,###,##0.00")
    GetMin = Left(cChar, 1) & nValue & Right(cChar, 1)
End Function

Private Function GetAbsMin(ByVal nNumber As Double) As Double
  GetAbsMin = 0 - nNumber
End Function

Private Function GetMin1(nValue)
Dim cChar As String

    cChar = "()"
    nValue = Format(Abs(GetNull(nValue)), "###,###,###,###,###,##0.00")
    If nValue = 0 Then
      GetMin1 = nValue
    Else
      GetMin1 = Left(cChar, 1) & nValue & Right(cChar, 1)
    End If
End Function

Function GetStatusPembelian(ByVal obj As CodeSuiteLibrary.Data, ByVal cFaktur As String) As String
Dim dbData As New ADODB.Recordset
 
 GetStatusPembelian = "0"
 Set dbData = obj.Browse(GetDSN, "TotPembelian", "Faktur,JenisPembelian", "Faktur", sisAssign, cFaktur)
 If Not dbData.EOF Then
  GetStatusPembelian = GetNull(dbData!JenisPembelian, "")
 End If
End Function

'==============================================================================
' NERACA
'==============================================================================

Function GetKasMasuk(ByVal obj As CodeSuiteLibrary.Data, ByVal dAwal As Date, ByVal dAkhir As Date) As Double
  Set dbData = obj.Browse(GetDSN, "pemasukan c", "sum(c.jumlah) as jumlah", "t.tgl", sisGTEqual, Format(dAwal, "yyyy-MM-dd"), " and t.tgl <='" & Format(dAkhir, "yyyy-MM-dd") & "' and m.jenis=2", , _
               Array("Left Join pos m on m.kodepos = c.kodepos", "left join totpemasukan t on t.nomorpemasukan = c.nomorpemasukan"))
  If Not dbData.EOF Then
    GetKasMasuk = GetNull(dbData!Jumlah)
  End If
End Function

Function GetKomisiPenjualan(ByVal obj As CodeSuiteLibrary.Data, ByVal dAwal As Date, ByVal dAkhir As Date) As Double
  Set dbData = obj.Browse(GetDSN, "TotPenjualan", "sum(Komisi) as JumlahKomisi", "tgl", sisGTEqual, Format(dAwal, "yyyy-MM-dd"), " and tgl <='" & Format(dAkhir, "yyyy-MM-dd") & "'")
  If Not dbData.EOF Then
    GetKomisiPenjualan = GetNull(dbData!JumlahKomisi)
  End If
End Function

Function GetKasKeluar(ByVal obj As CodeSuiteLibrary.Data, ByVal dAwal As Date, ByVal dAkhir As Date) As Double
  Set dbData = obj.Browse(GetDSN, "pengeluaran c", "sum(c.jumlah) as jumlah", "t.tgl", sisGTEqual, Format(dAwal, "yyyy-MM-dd"), " and t.tgl <='" & Format(dAkhir, "yyyy-MM-dd") & "' and m.jenis=1", , _
               Array("Left Join pos m on m.kodepos = c.kodepos", "LEFT JOIN totpengeluaran t on t.nomorpengeluaran = c.nomorpengeluaran"))
  If Not dbData.EOF Then
    GetKasKeluar = GetNull(dbData!Jumlah)
  End If
End Function

Sub GetPenjualan(ByVal obj As CodeSuiteLibrary.Data, ByVal dAwal As Date, ByVal dAkhir As Date, ByRef nPenjualan As Double, ByRef nPotongan As Double)
'penjualan kasir
  Set dbData = obj.Browse(GetDSN, "TotKasir", "Sum(Total) as Total,Sum(Discount) as Discount", "Tgl", sisGTEqual, Format(dAwal, "yyyy-mm-dd"), "And Tgl <='" & Format(dAkhir, "yyyy-mm-dd") & "'", "Tgl")
  If Not dbData.EOF Then
    nPenjualan = GetNull(dbData!Total)
    nPotongan = GetNull(dbData!Discount)
  End If
End Sub

Sub GetPenjualanAngsuran(ByVal obj As CodeSuiteLibrary.Data, ByVal dAwal As Date, ByVal dAkhir As Date, ByRef nPenjualanAngsuran As Double, ByRef nPotonganAngsuran As Double)
'penjualan ansuran/konsinyasi
  nPenjualanAngsuran = 0
  nPotonganAngsuran = 0
  Set dbData = obj.Browse(GetDSN, "TotAngsurBarang", "Sum(Total) as Total,Sum(Discount) as Discount", "Tgl", sisGTEqual, Format(dAwal, "yyyy-mm-dd"), "And Tgl <='" & Format(dAkhir, "yyyy-mm-dd") & "'", "Tgl")
  If Not dbData.EOF Then
    nPenjualanAngsuran = GetNull(dbData!Total)
    nPotonganAngsuran = GetNull(dbData!Discount)
  End If
End Sub

Sub GetPenjualanKredit(ByVal obj As CodeSuiteLibrary.Data, ByVal dAwal As Date, ByVal dAkhir As Date, ByRef nPenjualan As Double, ByRef nPotongan As Double)
'penjualan kredit
  
  Set dbData = obj.Browse(GetDSN, "TotPenjualan", "Sum(Total) as Total,Sum(SubTotal) as SubTotal,Sum(Discount) as Discount1,sum(Discount2) as Discount2,Sum(Pajak) as Pajak", "Tgl", sisGTEqual, Format(dAwal, "yyyy-mm-dd"), "And Tgl <='" & Format(dAkhir, "yyyy-mm-dd") & "'", "Tgl")
  If Not dbData.EOF Then
    nPenjualan = GetNull(dbData!Subtotal) + GetNull(dbData!PAJAK)
    nPotongan = GetNull(dbData!Discount1) + GetNull(dbData!Discount2)
  End If
End Sub

Sub Getpembelian(ByVal obj As CodeSuiteLibrary.Data, ByVal dAwal As Date, ByVal dAkhir As Date, ByRef nPembelian As Double, ByRef nPotongan As Double)
'Flag Status pembelian
'1. Pembelian Biasa
'2. Pembelian Konsinyasi

  Set dbData = obj.Browse(GetDSN, "TotPembelian", "Sum(Subtotal) as Total,Sum(Discount) as Discount1,sum(Discount2) as Discount2, sum(pajak) as Pajak", "Tgl", sisGTEqual, Format(dAwal, "yyyy-mm-dd"), " And Tgl <='" & Format(dAkhir, "yyyy-mm-dd") & "'", "Tgl")
  If Not dbData.EOF Then
    nPembelian = GetNull(dbData!Total) + GetNull(dbData!PAJAK)
    nPotongan = GetNull(dbData!Discount1) + GetNull(dbData!Discount2)
  End If
End Sub

Sub GetpembelianKonsinyasi(ByVal obj As CodeSuiteLibrary.Data, ByVal dAwal As Date, ByVal dAkhir As Date, ByRef nPembelianKonsinyasi As Double, ByRef nPotonganPembelianKonsinyasi As Double)
'Flag Status pembelian
'1. Pembelian Biasa
'2. Pembelian Konsinyasi

  nPembelianKonsinyasi = 0
  nPotonganPembelianKonsinyasi = 0
  Set dbData = obj.Browse(GetDSN, "TotPembelian", "Sum(Total) as Total,Sum(Discount) as Discount1,sum(Discount2) as Discount2", "Tgl", sisGTEqual, Format(dAwal, "yyyy-mm-dd"), "And Tgl <='" & Format(dAkhir, "yyyy-mm-dd") & "' and JenisPembelian = '2'", "Tgl")
  If Not dbData.EOF Then
    nPembelianKonsinyasi = GetNull(dbData!Total)
    nPotonganPembelianKonsinyasi = GetNull(dbData!Discount1) + GetNull(dbData!Discount2)
  End If
End Sub

Function GetDiscountPelunasanHutang(ByVal obj As CodeSuiteLibrary.Data, ByVal dAwal As Date, ByVal dAkhir As Date) As Double
  GetDiscountPelunasanHutang = 0
  Set dbData = obj.Browse(GetDSN, "TotPelunasanHutang", "Sum(Discount) as Discount", "Tgl", sisGTEqual, Format(dAwal, "yyyy-mm-dd"), "And Tgl <='" & Format(dAkhir, "yyyy-mm-dd") & "'")
  If Not dbData.EOF Then
    GetDiscountPelunasanHutang = GetNull(dbData!Discount)
  End If
End Function

Function GetReturpembelian(ByVal obj As CodeSuiteLibrary.Data, ByVal dAwal As Date, ByVal dAkhir As Date) As Double
  GetReturpembelian = 0
  Set dbData = obj.Browse(GetDSN, "TotRTnPembelian", "Sum(subTotal) as Total", "Tgl", sisGTEqual, Format(dAwal, "yyyy-mm-dd"), "And Tgl <='" & Format(dAkhir, "yyyy-mm-dd") & "'")
  If Not dbData.EOF Then
    GetReturpembelian = GetNull(dbData!Total)
  End If
End Function

Function GetNilaiStockAwal(ByVal obj As CodeSuiteLibrary.Data, ByVal dAkhir As Date) As Double
Dim cField As String
Dim cWhere As String
Dim vaJoin
Dim n As Integer
Dim nCari As Integer
Dim nTotal As Double

  cField = "s.hargabeli, s.kodestock"
  Set dbData = obj.Browse(GetDSN, "Stock s", cField, , , , cWhere, "s.kodestock")
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    n = 0
    dbData.MoveFirst
    vaArray.ReDim 0, dbData.RecordCount - 1, 0, 3
    Do While Not dbData.EOF
      FrmPB.RunPB
      vaArray(n, 0) = GetNull(dbData!KodeStock, "")
      vaArray(n, 1) = 0
      vaArray(n, 2) = GetNull(dbData!hargabeli)
      vaArray(n, 3) = 0
      n = n + 1
      dbData.MoveNext
    Loop
    FrmPB.EndPB
      
    'Ambil Data STock di KartuStock
    Set dbData = obj.Browse(GetDSN, "kartustock", "kodestock,Sum(debet) as Debet,sum(kredit) as Kredit", "Tgl", sisLT, Format(dAkhir, "yyyy-mm-dd"), "GROUP BY kodestock", "kodestock")
    If Not dbData.EOF Then
      dbData.MoveFirst
      Do While Not dbData.EOF
        nCari = vaArray.Find(0, 0, GetNull(dbData!KodeStock), , , XTYPE_STRING)
        If nCari >= 0 Then
          vaArray(nCari, 1) = vaArray(nCari, 1) + GetNull(dbData!debet) - GetNull(dbData!kredit)
        End If
        dbData.MoveNext
      Loop
    End If
    
    nTotal = 0
    For n = 0 To vaArray.UpperBound(1)
      vaArray(n, 3) = vaArray(n, 1) * vaArray(n, 2)
      nTotal = nTotal + vaArray(n, 3)
    Next
    GetNilaiStockAwal = nTotal
  End If
End Function

Function GetNilaiStockAkhir(ByVal obj As CodeSuiteLibrary.Data, ByVal dAkhir As Date) As Double
Dim cField As String
Dim cWhere As String
Dim vaJoin
Dim n As Integer
Dim nCari As Integer
Dim nTotal As Double
Dim a

  cField = "s.hargabeli,s.kodestock"
  Set dbData = obj.Browse(GetDSN, "stock s", cField, , , , cWhere, "s.kodestock")
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    n = 0
    dbData.MoveFirst
    vaArray.ReDim 0, dbData.RecordCount - 1, 0, 3
    Do While Not dbData.EOF
      FrmPB.RunPB
      vaArray(n, 0) = GetNull(dbData!KodeStock, "")
      vaArray(n, 1) = 0
      vaArray(n, 2) = GetNull(dbData!hargabeli)
      vaArray(n, 3) = 0
      n = n + 1
      dbData.MoveNext
    Loop
    FrmPB.EndPB
      
    'Ambil Data STock di KartuStock
    Set dbData = obj.Browse(GetDSN, "kartustock", "kodestock,Sum(debet) as Debet,sum(kredit) as Kredit", "Tgl", sisLTEqual, Format(dAkhir, "yyyy-mm-dd"), "GROUP BY kodestock", "kodestock")
    If Not dbData.EOF Then
      dbData.MoveFirst
      Do While Not dbData.EOF
        nCari = vaArray.Find(0, 0, GetNull(dbData!KodeStock), , , XTYPE_STRING)
        If nCari >= 0 Then
          vaArray(nCari, 1) = vaArray(nCari, 1) + GetNull(dbData!debet) - GetNull(dbData!kredit)
        End If
        dbData.MoveNext
      Loop
    End If
    
    nTotal = 0
    For n = 0 To vaArray.UpperBound(1)
      vaArray(n, 3) = vaArray(n, 1) * vaArray(n, 2)
      nTotal = nTotal + vaArray(n, 3)
    Next
    GetNilaiStockAkhir = nTotal
  End If
End Function

Function GetReturpenjualan(ByVal obj As CodeSuiteLibrary.Data, ByVal dAwal As Date, ByVal dAkhir As Date) As Double
  GetReturpenjualan = 0
  Set dbData = obj.Browse(GetDSN, "TotRTnPenjualan", "Sum(subTotal) as Total", "Tgl", sisGTEqual, Format(dAwal, "yyyy-mm-dd"), "And Tgl <='" & Format(dAkhir, "yyyy-mm-dd") & "'")
  If Not dbData.EOF Then
    GetReturpenjualan = GetNull(dbData!Total)
  End If
End Function

Function GetDiscountPelunasan(ByVal obj As CodeSuiteLibrary.Data, ByVal dAwal As Date, ByVal dAkhir As Date) As Double
  GetDiscountPelunasan = 0
  Set dbData = obj.Browse(GetDSN, "TotPelunasanPiutang", "Sum(Discount) as Discount", "Tgl", sisGTEqual, Format(dAwal, "yyyy-mm-dd"), "And Tgl <='" & Format(dAkhir, "yyyy-mm-dd") & "'")
  If Not dbData.EOF Then
    GetDiscountPelunasan = GetNull(dbData!Discount)
  End If
End Function
