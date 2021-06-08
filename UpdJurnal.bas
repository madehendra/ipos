Attribute VB_Name = "UpdJurnal"
Option Explicit

Dim dbData As New ADODB.Recordset

Enum BisaUpdateBukuBesar
  bbPembelian = 10
  bbRtnPembelian = 11
  bbPenjualan = 12
  bbRtnPenjualan = 13
  bbKasMasuk = 14
  bbKasKeluar = 15
  bbBiayaLainLain = 16
  bbHutangLainLain = 17
  bbPendapatanLainLain = 18
  bbPiutangLainLain = 19
  bbPelunasanHutang = 20
  bbPelunasanPiutang = 21
  bbJurnalLainLain = 22
  bbPenghapusanPiutang = 23
  bbPenjualanPegawai = 24
  bbRtnPenjualanPegawai = 25
  bbPelunasanPiutangPegawai = 26
  bbPenjualankasir = 27
  
  'enum for simpan pinjam
  bbTabungan = 51
  bbRealisasiKredit = 52
  bbAngsuranKredit = 53
  bbTitipanAngsuran = 54
  bbDeposito = 55
  bbRekeningRRP = 56
  bbPenyusutanAktiva = 67
  bbAmortisasiProvisi = 68
  bbPenambahanPlafond = 69
End Enum

Function UpdBukuBesar(ByVal obj As bisamydll.data, ByVal cStatus As BisaUpdateBukuBesar, _
                      ByVal cFaktur As String, ByVal cRekening As String, ByVal dTgl As Date, _
                      ByVal cKeterangan As String, Optional ByVal nDebet As Double = 0, _
                      Optional ByVal nKredit As Double = 0, Optional ByVal cUser As String = "")
Dim vaField, vaValue
Dim vaJoint

  If nDebet <> 0 Or nKredit <> 0 Then
    vaField = Array("STATUS", "Faktur", "Rekening", "Tgl", "Keterangan", "Debet", "Kredit", "UserName", "DateTime")
    vaValue = Array(cStatus, cFaktur, cRekening, dTgl, cKeterangan, nDebet, nKredit, GetRegistry(reg_UserName), SNow)
    obj.Add GetDSN, "BukuBesar", vaField, vaValue
  End If
End Function

Sub DeleteBukuBesar(ByVal obj As bisamydll.data, ByVal cPar As BisaUpdateBukuBesar, ByVal cFaktur As String)
  obj.Delete GetDSN, "BukuBesar", "Faktur", sisAssign, cFaktur, " and Status = '" & cPar & "' "
End Sub

Sub UpdateBukuBesarRekeningPembelian(ByVal obj As bisamydll.data, Optional ByVal cFakturPembelian As String, Optional ByVal lFromTransaksi As Boolean = False)
Dim cWhere As String

cWhere = ""
If lFromTransaksi And Trim(cFakturPembelian) <> "" Then
  cWhere = " Faktur = '" & cFakturPembelian & "'"
End If

Set dbData = obj.Browse(GetDSN, "TotPembelian t", "t.Faktur,t.tgl,t.Subtotal,t.Discount,t.Discount2,t.Pajak,t.Tunai,t.Hutang,s.Nama as NamaSupplier", , , , cWhere, , Array("Left Join Supplier s on s.kode = t.supplier"))
  If Not dbData.eof Then
    FrmPB.InitPB dbData.RecordCount
    DeleteBukuBesar obj, bbPembelian, GetNull(dbData!Faktur)
    Do While Not dbData.eof
        FrmPB.RunPB
        '-----------------------------
        ' JURNAL
        '-----------------------------
        ' Pembelian
        ' PPn Keluaran
          '   Kas
          '   Hutang
          '   Discount pembelian
        '-----------------------------
          
        UpdBukuBesar obj, bbPembelian, GetNull(dbData!Faktur), aCfg(obj, msRekeningPembelian), GetNull(dbData!tgl, "yyyy-MM-dd"), "Pembelian an. " & GetNull(dbData!NamaSupplier), GetNull(dbData!Subtotal, 0)
        UpdBukuBesar obj, bbPembelian, GetNull(dbData!Faktur), aCfg(obj, msRekeningPPnKeluaran), GetNull(dbData!tgl, "yyyy-MM-dd"), "Pembelian an. " & GetNull(dbData!NamaSupplier), GetNull(dbData!PAJAK, 0)
          UpdBukuBesar obj, bbPembelian, GetNull(dbData!Faktur), aCfg(obj, msRekeningKas), GetNull(dbData!tgl, "yyyy-MM-dd"), "Pembelian an. " & GetNull(dbData!NamaSupplier), 0, GetNull(dbData!tunai, 0)
          UpdBukuBesar obj, bbPembelian, GetNull(dbData!Faktur), aCfg(obj, msRekeningHutang), GetNull(dbData!tgl, "yyyy-MM-dd"), "Pembelian an. " & GetNull(dbData!NamaSupplier), 0, GetNull(dbData!Hutang, 0)
          UpdBukuBesar obj, bbPembelian, GetNull(dbData!Faktur), aCfg(obj, msRekeningDiscountPembelian), GetNull(dbData!tgl, "yyyy-MM-dd"), "Pembelian an. " & GetNull(dbData!NamaSupplier), 0, GetNull(dbData!Discount, 0) + GetNull(dbData!Discount2, 0)
      
        dbData.MoveNext
    Loop
    FrmPB.EndPB
  End If
End Sub


Sub UpdateBukuBesarRekeningReturPembelian(ByVal obj As bisamydll.data, Optional ByVal cFakturRETURPembelian As String, Optional ByVal lFromTransaksi As Boolean = False)
Dim cWhere As String

cWhere = ""
If lFromTransaksi And Trim(cFakturRETURPembelian) <> "" Then
  cWhere = " Faktur = '" & cFakturRETURPembelian & "'"
End If

Set dbData = obj.Browse(GetDSN, "TotRtnPembelian t", "t.Faktur,t.tgl,t.total,t.Subtotal,t.Discount,t.Discount2,t.Pajak,t.Tunai,t.Hutang,s.Nama as NamaSupplier", , , , cWhere, , Array("Left Join Supplier s on s.kode = t.supplier"))
  If Not dbData.eof Then
    FrmPB.InitPB dbData.RecordCount
    DeleteBukuBesar obj, bbRtnPembelian, GetNull(dbData!Faktur)
    Do While Not dbData.eof
        FrmPB.RunPB
        '-----------------------------
        ' JURNAL
        '-----------------------------
        '   Hutang
        '   Discount pembelian
            ' Pembelian
            ' PPn Keluaran
        UpdBukuBesar obj, bbRtnPembelian, GetNull(dbData!Faktur), aCfg(obj, msRekeningHutang), GetNull(dbData!tgl, "yyyy-MM-dd"), "Rtn Pembelian an. " & GetNull(dbData!NamaSupplier), GetNull(dbData!Total)
        UpdBukuBesar obj, bbRtnPembelian, GetNull(dbData!Faktur), aCfg(obj, msRekeningDiscountPembelian), GetNull(dbData!tgl, "yyyy-MM-dd"), "Rtn Pembelian an. " & GetNull(dbData!NamaSupplier), GetNull(dbData!Discount) + GetNull(dbData!Discount2)
          UpdBukuBesar obj, bbRtnPembelian, GetNull(dbData!Faktur), aCfg(obj, msRekeningPembelian), GetNull(dbData!tgl, "yyyy-MM-dd"), "Rtn Pembelian an. " & GetNull(dbData!NamaSupplier), , GetNull(dbData!Subtotal)
          UpdBukuBesar obj, bbRtnPembelian, GetNull(dbData!Faktur), aCfg(obj, msRekeningPPnKeluaran), GetNull(dbData!tgl, "yyyy-MM-dd"), "Rtn Pembelian an. " & GetNull(dbData!NamaSupplier), , GetNull(dbData!PAJAK)
          
        dbData.MoveNext
    Loop
    FrmPB.EndPB
  End If
End Sub

Sub UpdateBukuBesarRekeningPenjualan(ByVal obj As bisamydll.data, Optional ByVal cFakturPenjualan As String, Optional ByVal lFromTransaksi As Boolean = False)
Dim cWhere As String

cWhere = ""
If lFromTransaksi And Trim(cFakturPenjualan) <> "" Then
  cWhere = " Faktur = '" & cFakturPenjualan & "'"
End If

Set dbData = obj.Browse(GetDSN, "TotPenjualan t", "t.Faktur,t.tgl,t.Subtotal,t.Discount,t.Discount2,t.Pajak,t.Tunai,t.Piutang,s.Nama as NamaCustomer", , , , cWhere, , Array("Left Join Customer s on s.kode = t.customer"))
  If Not dbData.eof Then
    FrmPB.InitPB dbData.RecordCount
    DeleteBukuBesar obj, bbPenjualan, GetNull(dbData!Faktur)
    Do While Not dbData.eof
        FrmPB.RunPB
        'AUTO JURNAL
        ' Piutang Dagang          !Piutang    (Aktiva)
        ' Kas                     !Tunai      (Aktiva)
        ' Discount Penjualan      !Discount   (Pendapatan)
        '    Penjualan + PPn Masukan              !SubTotal + Pajak (Pendapatan)
      
        UpdBukuBesar obj, bbPenjualan, GetNull(dbData!Faktur), aCfg(obj, msRekeningPiutang), GetNull(dbData!tgl, ""), "Penjualan an. " & GetNull(dbData!NamaCustomer, ""), GetNull(dbData!Piutang)
        UpdBukuBesar obj, bbPenjualan, GetNull(dbData!Faktur), aCfg(obj, msRekeningKas), GetNull(dbData!tgl, ""), "Penjualan an. " & GetNull(dbData!NamaCustomer, ""), GetNull(dbData!tunai)
        UpdBukuBesar obj, bbPenjualan, GetNull(dbData!Faktur), aCfg(obj, msRekeningDiscountPenjualan), GetNull(dbData!tgl, ""), "Discount Penjualan an. " & GetNull(dbData!NamaCustomer, ""), GetNull(dbData!Discount) + GetNull(dbData!Discount2)
          UpdBukuBesar obj, bbPenjualan, GetNull(dbData!Faktur), aCfg(obj, msRekeningPenjualan), GetNull(dbData!tgl, ""), "Penjualan an. " & GetNull(dbData!NamaCustomer, ""), , GetNull(dbData!Subtotal)
          UpdBukuBesar obj, bbPenjualan, GetNull(dbData!Faktur), aCfg(obj, msRekeningPPnMasukan), GetNull(dbData!tgl, ""), "Penjualan an. " & GetNull(dbData!NamaCustomer, ""), , GetNull(dbData!PAJAK)
        dbData.MoveNext
    Loop
    FrmPB.EndPB
  End If
End Sub


Sub UpdateBukuBesarRekeningReturPenjualan(ByVal obj As bisamydll.data, Optional ByVal cFakturReturPenjualan As String, Optional ByVal lFromTransaksi As Boolean = False)
Dim cWhere As String

cWhere = ""
If lFromTransaksi And Trim(cFakturReturPenjualan) <> "" Then
  cWhere = " Faktur = '" & cFakturReturPenjualan & "'"
End If

Set dbData = obj.Browse(GetDSN, "TotRtnPenjualan t", "t.Faktur,t.tgl,t.Subtotal,t.Discount,t.Discount2,t.Pajak,t.Tunai,t.Piutang,s.Nama as NamaCustomer", , , , cWhere, , Array("Left Join Customer s on s.kode = t.customer"))
  If Not dbData.eof Then
    FrmPB.InitPB dbData.RecordCount
    DeleteBukuBesar obj, bbRtnPenjualan, GetNull(dbData!Faktur)
    Do While Not dbData.eof
        FrmPB.RunPB
        UpdBukuBesar obj, bbRtnPenjualan, GetNull(dbData!Faktur, ""), aCfg(obj, msRekeningReturPenjualan), GetNull(dbData!tgl, ""), "Rtn Penjualan an. " & GetNull(dbData!NamaCustomer), GetNull(dbData!Subtotal)
        UpdBukuBesar obj, bbRtnPenjualan, GetNull(dbData!Faktur, ""), aCfg(obj, msRekeningPPnKeluaran), GetNull(dbData!tgl, ""), "Rtn Penjualan an. " & GetNull(dbData!NamaCustomer), GetNull(dbData!PAJAK)
          UpdBukuBesar obj, bbRtnPenjualan, GetNull(dbData!Faktur, ""), aCfg(obj, msRekeningPiutang), GetNull(dbData!tgl, ""), "Rtn Penjualan an. " & GetNull(dbData!NamaCustomer), , GetNull(dbData!Subtotal)
          UpdBukuBesar obj, bbRtnPenjualan, GetNull(dbData!Faktur, ""), aCfg(obj, msRekeningDiscountPenjualan), GetNull(dbData!tgl, ""), "Rtn Penjualan an. " & GetNull(dbData!NamaCustomer), , GetNull(dbData!Discount) + GetNull(dbData!Discount2)
        dbData.MoveNext
    Loop
    FrmPB.EndPB
  End If
End Sub


Sub UpdateBukuBesarRekeningPenjualanKasir(ByVal obj As bisamydll.data, Optional ByVal cFakturPenjualanKasir As String, Optional ByVal lFromTransaksi As Boolean = False)
Dim cWhere As String

cWhere = ""
If lFromTransaksi And Trim(cFakturPenjualanKasir) <> "" Then
  cWhere = " Faktur = '" & cFakturPenjualanKasir & "'"
End If

Set dbData = obj.Browse(GetDSN, "TotKasir t", "t.*", , , , cWhere)
  If Not dbData.eof Then
    FrmPB.InitPB dbData.RecordCount
    DeleteBukuBesar obj, bbPenjualankasir, GetNull(dbData!Faktur)
    Do While Not dbData.eof
        FrmPB.RunPB
        
        'UPDATE BUKU BESAR
        '-----------------
        'KAS
        'DISCOUNT PENJUALAN
          'PENJUALAN
        UpdBukuBesar obj, bbPenjualankasir, GetNull(dbData!Faktur), aCfg(obj, msRekeningKas), GetNull(dbData!tgl, "yyyy-MM-dd"), "Penjualan kasir", GetNull(dbData!Total)
        UpdBukuBesar obj, bbPenjualankasir, GetNull(dbData!Faktur), aCfg(obj, msRekeningDiscountPenjualan), GetNull(dbData!tgl, "yyyy-MM-dd"), "Penjualan kasir", GetNull(dbData!Discount)
            UpdBukuBesar obj, bbPenjualankasir, GetNull(dbData!Faktur), aCfg(obj, msRekeningPenjualan), GetNull(dbData!tgl, "yyyy-MM-dd"), "Penjualan Kasir", , GetNull(dbData!Total) + GetNull(dbData!Discount)
            
        dbData.MoveNext
    Loop
    FrmPB.EndPB
  End If
End Sub

Sub UpdateBukuBesarRekeningCashIn(ByVal obj As bisamydll.data, Optional ByVal cFakturCashIn As String, Optional ByVal lFromTransaksi As Boolean = False)
Dim cWhere As String

cWhere = ""
If lFromTransaksi And Trim(cFakturCashIn) <> "" Then
  cWhere = " and Faktur = '" & cFakturCashIn & "'"
End If

Set dbData = obj.Browse(GetDSN, "Cost c", "c.*", "m.jenis", sisAssign, "2", cWhere, , Array("Left Join MstCost m on m.kode = c.kode"))
  If Not dbData.eof Then
    FrmPB.InitPB dbData.RecordCount
    DeleteBukuBesar obj, bbKasMasuk, GetNull(dbData!Faktur)
    Do While Not dbData.eof
        FrmPB.RunPB
        '------------
        'AUTO JURNAL
        '------------
        UpdBukuBesar obj, bbKasMasuk, GetNull(dbData!Faktur), aCfg(obj, msRekeningKas), GetNull(dbData!tgl, "yyyy-MM-dd"), "Kas Masuk " & GetNull(dbData!Keterangan), GetNull(dbData!Jumlah), , SNow
          UpdBukuBesar obj, bbKasMasuk, GetNull(dbData!Faktur), aCfg(obj, msRekeningPendapatanKasKecil), GetNull(dbData!tgl, "yyyy-MM-dd"), "Kas Masuk " & GetNull(dbData!Keterangan), , GetNull(dbData!Jumlah), SNow
        dbData.MoveNext
    Loop
    FrmPB.EndPB
  End If
End Sub

Sub UpdateBukuBesarRekeningCashOut(ByVal obj As bisamydll.data, Optional ByVal cFakturCashOut As String, Optional ByVal lFromTransaksi As Boolean = False)
Dim cWhere As String

cWhere = ""
If lFromTransaksi And Trim(cFakturCashOut) <> "" Then
  cWhere = " and Faktur = '" & cFakturCashOut & "'"
End If

Set dbData = obj.Browse(GetDSN, "Cost c", "c.*", "m.jenis", sisAssign, "1", cWhere, , Array("Left Join MstCost m on m.kode = c.kode"))
  If Not dbData.eof Then
    FrmPB.InitPB dbData.RecordCount
    DeleteBukuBesar obj, bbKasKeluar, GetNull(dbData!Faktur)
    Do While Not dbData.eof
        FrmPB.RunPB
        '-----------
        'AUTO JURNAL
        '-----------
        UpdBukuBesar obj, bbKasKeluar, GetNull(dbData!Faktur), aCfg(obj, msRekeningBiayaKasKecil), GetNull(dbData!tgl, "yyyy-MM-dd"), "Kas Keluar " & GetNull(dbData!Keterangan), GetNull(dbData!Jumlah)
            UpdBukuBesar obj, bbKasKeluar, GetNull(dbData!Faktur), aCfg(obj, msRekeningKas), GetNull(dbData!tgl, "yyyy-MM-dd"), "Kas Keluar " & GetNull(dbData!Keterangan), , GetNull(dbData!Jumlah)
        dbData.MoveNext
    Loop
    FrmPB.EndPB
  End If
End Sub
