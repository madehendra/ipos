VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm aMainmenu 
   BackColor       =   &H8000000C&
   Caption         =   "CodeSuite Point Of Sales"
   ClientHeight    =   9300
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   18840
   Icon            =   "aMainmenu.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   18840
      _ExtentX        =   33232
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Close"
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "LogOff"
            Object.ToolTipText     =   "Log Off"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Printer"
            Object.ToolTipText     =   "Setup Printer"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Object.ToolTipText     =   "Help"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Calculator"
            Object.ToolTipText     =   "Calculator"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "About"
            Object.ToolTipText     =   "About"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pbTray 
      Align           =   1  'Align Top
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   0
      Picture         =   "aMainmenu.frx":000C
      ScaleHeight     =   240
      ScaleWidth      =   18780
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   18840
   End
   Begin VB.PictureBox pcCancel 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      Picture         =   "aMainmenu.frx":0596
      ScaleHeight     =   285
      ScaleWidth      =   18780
      TabIndex        =   2
      Top             =   660
      Visible         =   0   'False
      Width           =   18840
   End
   Begin VB.PictureBox pcExit 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      Picture         =   "aMainmenu.frx":0725
      ScaleHeight     =   285
      ScaleWidth      =   18780
      TabIndex        =   3
      Top             =   1005
      Visible         =   0   'False
      Width           =   18840
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   4
      Top             =   8985
      Width           =   18840
      _ExtentX        =   33232
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   9701
            MinWidth        =   9701
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "19:35"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "10/03/2021"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1005
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   1035
      Top             =   2115
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   135
      Top             =   1635
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "aMainmenu.frx":07BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "aMainmenu.frx":0D55
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "aMainmenu.frx":12EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "aMainmenu.frx":1889
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "aMainmenu.frx":1E23
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "aMainmenu.frx":23BD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&FILE"
      Index           =   0
      Begin VB.Menu mnuLogOff 
         Caption         =   "&Log Off"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&INVENTORY"
      Index           =   1
      Begin VB.Menu mnuMasterIsiDataSupplier 
         Caption         =   "Supplier Baru"
      End
      Begin VB.Menu mnuSupplierBalance 
         Caption         =   "Supplier Balance"
      End
      Begin VB.Menu mnuMasterGolongan 
         Caption         =   "GO&LONGAN DAN SATUAN"
         Begin VB.Menu mnuGolonganInventory 
            Caption         =   "Golongan"
         End
         Begin VB.Menu mnuSatuanInventory 
            Caption         =   "Satuan"
         End
         Begin VB.Menu mnuKategori 
            Caption         =   "Kategori"
         End
         Begin VB.Menu mnuKelompokHarga 
            Caption         =   "Kelompok Harga"
         End
      End
      Begin VB.Menu mnuMasterStock 
         Caption         =   "Inventory Baru"
      End
      Begin VB.Menu mnuMutasiStockAntarGudang 
         Caption         =   "Mutasi Stock Antar Gudang"
      End
      Begin VB.Menu sptPacking 
         Caption         =   "-"
      End
      Begin VB.Menu mhuPackingInventory 
         Caption         =   "Pack and UnPack Inventory"
      End
      Begin VB.Menu sptUtilityStock 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUtilityforStock 
         Caption         =   "Utility for Stock/Inventory"
      End
      Begin VB.Menu mnuUtilityForSupplier 
         Caption         =   "Utility for Supplier"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&MEMBER"
      Index           =   2
      Begin VB.Menu mnuMasterIsiDataCustomer 
         Caption         =   "Member Baru"
      End
      Begin VB.Menu mnuStockKontrak 
         Caption         =   "Harga Kontrak"
      End
      Begin VB.Menu mnuDepartment 
         Caption         =   "Department"
      End
      Begin VB.Menu mnuMemberUtility 
         Caption         =   "Utility for Member"
      End
      Begin VB.Menu sptMemberUtility 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMemberBalance 
         Caption         =   "Member Balance"
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&TRANSAKSI"
      Index           =   3
      Begin VB.Menu mnuTrPembelian 
         Caption         =   "PURCHASING"
         Begin VB.Menu mnuTransaksiPembelianTanpaOrder 
            Caption         =   "Pembelian"
         End
         Begin VB.Menu mnuTransaksiPenyesuaianStock 
            Caption         =   "Stock Opname"
         End
         Begin VB.Menu mnuTrReturPembelian 
            Caption         =   "Retur Pembelian"
         End
         Begin VB.Menu mnuRefund 
            Caption         =   "Refund/Cash Back"
         End
         Begin VB.Menu mnuReturKonsinyasi 
            Caption         =   "Retur Konsinyasi"
         End
         Begin VB.Menu mnuTrKomplimen 
            Caption         =   "Komplimen"
         End
      End
      Begin VB.Menu mnuTrPenjualan 
         Caption         =   "KASIR"
         Begin VB.Menu mnuTransaksiPenjualan 
            Caption         =   "1. Penjualan"
         End
         Begin VB.Menu mnuTransaksiPembayaranPiutang 
            Caption         =   "2. Piutang (Terima Uang Penjualan Barang)"
         End
         Begin VB.Menu mnuPengeluaranBiaya 
            Caption         =   "3. Pengeluaran Biaya-Biaya"
         End
         Begin VB.Menu mnuTransaksiPembayaranHutang 
            Caption         =   "4. Bayar Supplier"
         End
         Begin VB.Menu mnuMutasiKasDanBank 
            Caption         =   "5. Mutasi Kas Dan Bank"
         End
         Begin VB.Menu sptPencairanBG 
            Caption         =   "-"
         End
         Begin VB.Menu mnuBuyBack 
            Caption         =   "Beli Kembali Barang Yg Sudah Dijual"
         End
         Begin VB.Menu mnuTransaksiReturPenjualan 
            Caption         =   "Retur Penjualan"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuTrCetakFakturPembayaranHutang 
            Caption         =   "Cetak Faktur Pembayaran Hutang"
            Visible         =   0   'False
         End
         Begin VB.Menu sptTopUp 
            Caption         =   "-"
         End
         Begin VB.Menu mnuMemberTopUp 
            Caption         =   "Member Top Up"
         End
         Begin VB.Menu mnuMemberWithdraw 
            Caption         =   "Member Withdraw"
         End
         Begin VB.Menu mnuTransaksiBG 
            Caption         =   "BG & CEK"
            Visible         =   0   'False
            Begin VB.Menu mnuPencairanBG 
               Caption         =   "Pencairan BG/Cek"
            End
            Begin VB.Menu mnuPembatalanatauPenghapusanBG 
               Caption         =   "Pembatalan atau Penghapusan BG/Cek"
            End
         End
      End
      Begin VB.Menu mnuInputCatatanPelanggan 
         Caption         =   "Input Catatan Pelanggan"
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&ACCOUNTING"
      Index           =   4
      Begin VB.Menu mnuJurnalUmum 
         Caption         =   "Jurnal Umum"
      End
      Begin VB.Menu mnuBuatAkunBaru 
         Caption         =   "Akun atau Pos Accounting Baru"
      End
      Begin VB.Menu mnuPrive 
         Caption         =   "Prive"
      End
      Begin VB.Menu mnuUpdateKartuPiutangMember 
         Caption         =   "Update Kartu Piutang Member"
      End
      Begin VB.Menu mnuTrAkuntansi 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSettingAccounting 
         Caption         =   "SETTINGS"
         Begin VB.Menu mnuGeneralSettingAkuntansi 
            Caption         =   "General Setup"
         End
         Begin VB.Menu mnuRekeningLabaRugiTahunBerjalan 
            Caption         =   "Rekening Laba Rugi Tahun Berjalan"
         End
         Begin VB.Menu mnuRekeningNilaiPersediaan 
            Caption         =   "Rekening Nilai Persediaan"
         End
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&LAPORAN"
      Index           =   5
      Begin VB.Menu mnuInventoryList 
         Caption         =   "Price List"
      End
      Begin VB.Menu mnuLaporanStock 
         Caption         =   "STOCK"
         Begin VB.Menu mnuLaporanDaftarStock 
            Caption         =   "Saldo Stock"
         End
         Begin VB.Menu mnuLaporanKartuStock 
            Caption         =   "Kartu Stock"
         End
         Begin VB.Menu mnuLaporanStp01 
            Caption         =   "-"
         End
         Begin VB.Menu mnuLaporanNilaiPersediaanStock 
            Caption         =   "Nilai Persediaan Stock"
         End
         Begin VB.Menu mnuLaporanMinimumStock 
            Caption         =   "Minimum Stock"
         End
         Begin VB.Menu mnuLaporanStockOpname 
            Caption         =   "Laporan Stock Opname"
         End
         Begin VB.Menu sptKartuStock 
            Caption         =   "-"
         End
         Begin VB.Menu mnuMutasiStock 
            Caption         =   "Mutasi Stock"
         End
         Begin VB.Menu mnuLaporanProdukTerlaris 
            Caption         =   "Laporan Produk Terlaris"
         End
      End
      Begin VB.Menu mnuLaporanSpt01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLaporanSupplier 
         Caption         =   "SUPPLIER"
         Begin VB.Menu mnuDaftarSupplier 
            Caption         =   "Daftar Supplier"
         End
         Begin VB.Menu mnuLaporanSaldoHutang 
            Caption         =   "Laporan Saldo Hutang"
         End
         Begin VB.Menu mnuLaporanKartuHutang 
            Caption         =   "Laporan Kartu Hutang"
         End
         Begin VB.Menu mnuLaporanPelunasanHutang 
            Caption         =   "Laporan Pelunasan Hutang"
         End
      End
      Begin VB.Menu mnuLaporanPembelian 
         Caption         =   "PEMBELIAN"
         Begin VB.Menu mnuLaporanPembelianDetail 
            Caption         =   "Laporan Pembelian"
         End
         Begin VB.Menu mnuLaporanReturPembelian 
            Caption         =   "Laporan Retur Pembelian"
         End
         Begin VB.Menu mnuLaporanPPnMasukan 
            Caption         =   "PPn Masukan"
         End
         Begin VB.Menu mnuLaporanBarangMasuk 
            Caption         =   "Laporan Barang Masuk"
         End
      End
      Begin VB.Menu mnuLaporanJatuhTempoHutang 
         Caption         =   "Laporan Jatuh Tempo Hutang"
      End
      Begin VB.Menu mnuLaporanSpt90 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLaporanCustomer 
         Caption         =   "MEMBER"
         Begin VB.Menu mnuLaporanDaftarCustomer 
            Caption         =   "Daftar Member"
         End
         Begin VB.Menu mnuLaporanSaldoPiutang 
            Caption         =   "Laporan Saldo Piutang"
         End
         Begin VB.Menu mnuLaporanKartuPiutang 
            Caption         =   "Laporan Kartu Piutang"
         End
         Begin VB.Menu mnuLaporanPelunasanPiutang 
            Caption         =   "Laporan Pelunasan Piutang"
         End
         Begin VB.Menu mnuDetailPenjualanItem 
            Caption         =   "Detail Penjualan Item"
         End
         Begin VB.Menu sptMemberTopUp 
            Caption         =   "-"
         End
         Begin VB.Menu mnuLaporanMemberTopUp 
            Caption         =   "Laporan Member Top Up"
         End
         Begin VB.Menu mnuLaporanKartuTopUp 
            Caption         =   "Laporan Kartu Top Up"
         End
         Begin VB.Menu mnuLaporanDetailPenjualanNonInventory 
            Caption         =   "Laporan Detail Penjualan Non Inventory"
         End
      End
      Begin VB.Menu mnuLaporanPenjualan 
         Caption         =   "PENJUALAN"
         Begin VB.Menu mnuAllSalesDetails 
            Caption         =   "All Sales Datails"
         End
         Begin VB.Menu mnuLaporanDetailPenjualan 
            Caption         =   "Laporan Penjualan"
         End
         Begin VB.Menu mnuLaporanReturPenjualan 
            Caption         =   "Laporan Retur Penjualan"
         End
         Begin VB.Menu mnuLaporanReturPenjualanGroupBySalesman 
            Caption         =   "Laporan Retur Penjualan Group by Salesman"
         End
         Begin VB.Menu mnuLaporanPPnKeluaran 
            Caption         =   "Laporan PPn Keluaran"
         End
         Begin VB.Menu mnuLaporanPenjualanHarian 
            Caption         =   "Laporan Penjualan Harian"
         End
         Begin VB.Menu mnuLaporanPenjualanBrutoHarian 
            Caption         =   "Laporan Penjualan Bruto Harian"
         End
         Begin VB.Menu mnuLaporanPenjualanBarangKenaPajak 
            Caption         =   "Laporan Penjualan Barang Kena Pajak"
         End
         Begin VB.Menu mnuLaporanRekapPenjualanBelumLunas 
            Caption         =   "Laporan Rekap Tagihan Penjualan/Pembelian"
         End
         Begin VB.Menu mnuSensusSalesHarian 
            Caption         =   "Sensus Sales Harian"
         End
      End
      Begin VB.Menu mnuLaporanKonsinyasi 
         Caption         =   "KONSINYASI"
         Begin VB.Menu mnuRptStockKonsinyasi 
            Caption         =   "Stock Konsinyasi"
         End
         Begin VB.Menu mnuLaporanPenjualanKonsinyasi 
            Caption         =   "Penjualan Konsinyasi"
         End
      End
      Begin VB.Menu mnuLaporanJatuhTempoPiutang 
         Caption         =   "Laporan Jatuh Tempo Piutang"
      End
      Begin VB.Menu sptSalesmanReport 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLaporanPajakPenjualan 
         Caption         =   "LAPORAN PAJAK PENJUALAN"
         Begin VB.Menu mnuLaporanDetailPajakPenjualan 
            Caption         =   "Laporan Detail Pajak Penjualan"
         End
      End
      Begin VB.Menu mnuSalesmanReport 
         Caption         =   "REPORT SALESMAN"
         Begin VB.Menu mnuSalesmanList 
            Caption         =   "Salesman List"
         End
         Begin VB.Menu mnuOmzetSalesmanReport 
            Caption         =   "Omzet Salesman Report"
         End
         Begin VB.Menu mnuGrossSaleSalesmanReport 
            Caption         =   "Gross Sale Salesman Report"
         End
         Begin VB.Menu mnuRekapitulasiOmzetSales 
            Caption         =   "Rekapitulasi Omzet Sales"
         End
      End
      Begin VB.Menu sptTagihanBulanan 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChartOfAccount 
         Caption         =   "Chart Of Account"
      End
      Begin VB.Menu mnuLaporanGeneralLedger 
         Caption         =   "Laporan General Ledger..."
      End
      Begin VB.Menu mnuLaporanBukuBesar 
         Caption         =   "Laporan Buku Besar"
      End
      Begin VB.Menu mnuLaporanTrialBalance 
         Caption         =   "Trial Balance Report"
      End
      Begin VB.Menu mnuLaporanNeraca 
         Caption         =   "Laporan Neraca"
      End
      Begin VB.Menu mnuLaporanLabaRugiUsaha 
         Caption         =   "Laporan Laba Rugi Usaha"
      End
      Begin VB.Menu mnuGrossSalesReport 
         Caption         =   "Gross Sales Report"
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&SETUP"
      Index           =   6
      Begin VB.Menu MnuPassword 
         Caption         =   "&Create User Name"
      End
      Begin VB.Menu MnuMenuLevel 
         Caption         =   "&Setup Menu Level"
      End
      Begin VB.Menu MnuChangePassword 
         Caption         =   "&Change Password"
      End
      Begin VB.Menu mnuOtorisasi 
         Caption         =   "Otorisasi menambah, mengkoreksi atau menghapus"
      End
      Begin VB.Menu SetupSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSetupSalesman 
         Caption         =   "Master Salesman"
      End
      Begin VB.Menu mstDatabaseGudang 
         Caption         =   "Master Gudang "
      End
      Begin VB.Menu mnuSetupCostCenter 
         Caption         =   "Master Cost Center"
      End
      Begin VB.Menu mnuMasterGroupSales 
         Caption         =   "Master Group Sales"
      End
      Begin VB.Menu mnuMasterKeteranganBayar 
         Caption         =   "Master Keterangan Bayar"
      End
      Begin VB.Menu mnuSettingCostCentre 
         Caption         =   "Setup Cost Centre dan Gudang..."
      End
      Begin VB.Menu mnuSetupRekeningBiaya 
         Caption         =   "Setup Rekening Biaya"
      End
      Begin VB.Menu mnuSetupAkunKas 
         Caption         =   "Setup Akun Kas..."
      End
      Begin VB.Menu mnuSetupPeriodeAkuntansi 
         Caption         =   "Setup Periode Akuntansi"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSetupKartu 
         Caption         =   "Setup Kartu"
      End
      Begin VB.Menu sptKodeKas 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCfgSetPrinter 
         Caption         =   "Setup Printer"
      End
      Begin VB.Menu mnuSetupPortPrinter 
         Caption         =   "Setup Printer Struk"
      End
      Begin VB.Menu mnuGeneralSettings 
         Caption         =   "General Setup Penjualan dan Pembelian..."
      End
      Begin VB.Menu mnuMarkupHarga 
         Caption         =   "Mark Up Harga Jual"
      End
      Begin VB.Menu mnuHargaGrosir 
         Caption         =   "Harga Grosir"
      End
      Begin VB.Menu mnuTarifCOD 
         Caption         =   "Tarif COD"
      End
      Begin VB.Menu MnuSetupInfoPerusahaan 
         Caption         =   "Informasi Perusahaan"
      End
      Begin VB.Menu sptWallpaper 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWallpaper 
         Caption         =   "Wallpaper"
      End
      Begin VB.Menu mnuSetupDatabase 
         Caption         =   "DATABASE"
         Begin VB.Menu mnuTruncateDatabase 
            Caption         =   "Truncate Database"
         End
      End
      Begin VB.Menu mnuRegister 
         Caption         =   "Register ..."
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&WINDOWS"
      Index           =   8
      WindowList      =   -1  'True
      Begin VB.Menu MnuWindows 
         Caption         =   "Tile Horizontally"
         Index           =   0
      End
      Begin VB.Menu MnuWindows 
         Caption         =   "Tile Vertically"
         Index           =   1
      End
      Begin VB.Menu MnuWindows 
         Caption         =   "Cascade"
         Index           =   2
      End
      Begin VB.Menu mnuWindowFullScreen 
         Caption         =   "Full Screen"
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&HELP (WOWSYSTEM - 081 338 414 828)"
      Index           =   9
   End
   Begin VB.Menu mnuMain 
      Caption         =   "MAINTENANCE"
      Index           =   11
      Visible         =   0   'False
      Begin VB.Menu mnuMaintenanceDatabase 
         Caption         =   "Database"
      End
   End
End
Attribute VB_Name = "aMainmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lFirst As Boolean
Dim cKode As String
Dim objMenu As New CodeSuiteLibrary.Menu
Dim objData As New CodeSuiteLibrary.Data
Dim dbData As New ADODB.Recordset
Dim cIPNumber As String
Dim cNamaDatabase As String
Dim cNamaDSN As String
Dim cPort As String
Dim cMYODBCPATH As String
Dim cKeySecret As String
Dim cMYODBCFile As String
Dim cShellPrn As String
Dim cVersionName As String
Dim cAppsName As String
Dim nToExpired As Integer
Dim cModePelunasanPiutang As String
Dim lWIndowStatus As Integer

'Fungsi untuk men disable min max and close button in MDI
Private Const SC_CLOSE As Long = &HF060&
Private Const SC_MAXIMIZE As Long = &HF030&
Private Const SC_MINIMIZE As Long = &HF020&

Private Const xSC_CLOSE As Long = -10&
Private Const xSC_MAXIMIZE As Long = -11&
Private Const xSC_MINIMIZE As Long = -12&

Private Const GWL_STYLE = (-16)
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000

Private Const hWnd_NOTOPMOST = -2
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_FRAMECHANGED = &H20

Private Const MIIM_STATE As Long = &H1&
Private Const MIIM_ID As Long = &H2&
Private Const MFS_GRAYED As Long = &H3&
Private Const WM_NCACTIVATE As Long = &H86

Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    ftype As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Sub GetRandom()
Dim nRound

  Randomize Second(Now)
  nRound = Round(Rnd, 3)
  Do While Len(nRound) < 5
    nRound = Round(Rnd, 3)
  Loop
  GetID = Right(nRound, 3)
End Sub

Private Sub MDIForm_Activate()
  If lFirst Then
  

    
    Me.Caption = App.ProductName & " " '& cVersionName & IIf(authKey(GetRegistry(reg_SerialNumber), MBSerialNumber), " Pro", "Trial")
    lFirst = False
    If Not objMenu.GetPassword(cKode, Me, GetDSN) Then
      End
    Else
      GetNotifikasiAdd "Cek Lisensi"
      GetExpiredApp
      GetNotifikasiRemove
    End If

    mnuLogOff.Caption = "&Log Off " & Trim(objMenu.FullName) & "..."
    Toolbar1.Buttons(2).ToolTipText = mnuLogOff.Caption
    Toolbar1.Buttons(6).ToolTipText = "About " & GetAppDescription
    
    
    SaveRegistry reg_FullName, objMenu.FullName
    SaveRegistry reg_UserLevel, objMenu.UserLevel
    SaveRegistry reg_Username, objMenu.UserName
    SaveRegistry reg_UserID, objMenu.UserID
    
    
    StatusBar1.Panels(2).Width = StatusBar1.Width - StatusBar1.Panels(1).Width - StatusBar1.Panels(3).Width - StatusBar1.Panels(4).Width - StatusBar1.Panels(5).Width - StatusBar1.Panels(6).Width - StatusBar1.Panels(7).Width
    StatusBar1.Panels(1).Text = "USER: " & objMenu.UserName & " LEVEL: " & objMenu.UserLevel & " SERVER: " & UCase(GetRegistry(reg_IP)) & " DATABASE: " & UCase(GetRegistry(reg_Database))
    
    Set dbData = objData.Browse(GetDSN, "bukubesar", "max(datetime) as lasupdate")
    StatusBar1.Panels(2).Text = "DB." & Format(GetNull(dbData!lasupdate), "dd.MM.yy HH:MM:SS") & " Ver." & App.Major & "." & App.Minor & "." & App.Revision & " EXP. " & CryptRC4(FromHexDump(GetToken(objData)), GetRegistry(reg_KeySecret))   'Format(DateAdd("d", GetNewExpiredApp, Date), "dd-MMM-yy")
    'StatusBar1.Panels(8).Text = " exp. " & Format(DateAdd("d", GetNewExpiredApp, Date), "dd-MMM-yy") & " EXP"
    
    cKasTeller = ""
    Set dbData = objData.Browse(GetDSN, "akunkas k", "k.username,k.kodeakun,a.keterangan", "k.username", sisAssign, objMenu.UserName, , , Array("LEFT JOIN akun a ON a.kodeakun = k.kodeakun"))
    If Not dbData.EOF Then
      cKasTeller = GetNull(dbData!kodeakun, "")
      cNamaKasTeller = GetNull(dbData!keterangan)
      If Trim(cKasTeller) = "" Then
        MsgBox "Program tidak bisa dilanjutkan" & vbCrLf & vbCrLf & "Maaf, user ID yang anda gunakan belum memiliki Akun Kas." & vbCrLf & "Silahkan hubungi Administrator untuk mendapatkan Akun Kas"
        End
      End If
    Else
      If objMenu.UserName <> "ROOT" Then
        MsgBox "Program tidak bisa dilanjutkan" & vbCrLf & vbCrLf & "Maaf, user ID yang anda gunakan belum memiliki Akun Kas." & vbCrLf & "Silahkan hubungi Administrator untuk mendapatkan Akun Kas"
        End
      End If
    End If
    
    Me.Picture = LoadPicture(GetPicture(GetRegistry(reg_Wallpaper)))
    
    'Otentikasi
    
'    If authKey(GetRegistry(reg_SerialNumber), GetBIOSName) = False Then
'      If CheckTrial(objData, 100, Trial) = True Then
'        Load frmAbout
'        frmAbout.Show vbModal
'      Else
'        MsgBox "Anda sedang menggunakan versi TRIAL dari program " & App.ProductName & vbCrLf & "Untuk mendapatkan versi penuh (Pro), silahkan mendapatkan serial key sesuai dengan yg dijelaskan pada menu Help > About"
'      End If
'    Else
'      aMainmenu.Caption = App.ProductName & " Activated"
'    End If
    
  End If
  
End Sub

Private Function GetPembanding() As Boolean
Dim db As New ADODB.Recordset
Dim obj As New CodeSuiteLibrary.Data
Dim nSaldoPiutang As Double
Dim nSaldoBukuBesarPiutang As Double

'  Set db = obj.Browse(GetDSN, "kartupiutang", "sum(debet-kredit) as saldo", "tgl", sisGTEqual, "2010-11-1")
'  If Not db.EOF Then
'    nSaldoPiutang = GetNull(db!Saldo)
'  End If
'
'  Set db = obj.Browse(GetDSN, "bukubesar", "sum(debet-kredit) as saldobukubesar", "kodeakun", sisAssign, "1.300.10", " and tgl >= '2010-11-1'")
'  If Not db.EOF Then
'    nSaldoBukuBesarPiutang = GetNull(db!saldobukubesar)
'  End If
'
'  If nSaldoPiutang <> nSaldoBukuBesarPiutang Then
'    MsgBox "Maaf neraca tidak balance " & vbCrLf & "Posisi Saldo Piutang tidak balance dengan pos di Neraca : Saldo Piutang " & Format(nSaldoPiutang, "###,###,##0.00") & " Neraca : " & Format(nSaldoBukuBesarPiutang, "###,###,##0.00") & vbCrLf & _
'    "Untuk keamanan data, segera hub programmer program ini untuk mendapatkan perbaikan..." & _
'    "Terimakasih"
'  End If
  
  'bandingkan kartu piutang dengan neraca
  '1 ambil data mentahan di kartu piutang
  '2 ambil data di neraca/pos bukubesar kartupiutang
  
End Function

Function InitConnection(Optional pAuto As Boolean = False)
  GetNotifikasiAdd "Melakukan Koneksi Ke Server"
  If Not SetAuto() Then
    On Error Resume Next
    Set GetDSN = New ADODB.Connection
    'GetDSN.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=True;Data Source=" & GetRegistry(reg_DSN)
    GetDSN.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=True;Data Source=" & GetRegistry(reg_DSN)
    GetDSN.CursorLocation = adUseClient
'     GetDSN.CursorLocation = adUseServer
'     GetDSN
    GetDSN.Open
  End If
  
  If pAuto Then
    SetAuto pAuto
  End If
  GetNotifikasiRemove
End Function

Function SetAuto(Optional lAuto As Variant = Null) As Boolean
Static l As Boolean
  SetAuto = l
  If Not IsNull(lAuto) Then
    l = lAuto
  End If
End Function

Private Sub MDIForm_Deactivate()
  GetCloseApps cAppsName
End Sub

Private Sub GetCloseApps(ByVal appsName As String)
  Shell "taskkill /f /im " & appsName, vbHide
End Sub

Private Sub MDIForm_Load()
Dim lSave As Boolean
Dim cUID As String
Dim cPwd As String

'  GetScan
'  MsgBox cNamaKomputer

  
  cAppsName = App.EXEName
  If App.PrevInstance = True Then
    GetCloseApps cAppsName
  End If
  
  nRecordsTrial = 0
  
  isTrial = False
  If nRecordsTrial > 0 Then
    isTrial = True
  End If
  
  isTrial = False
  lSave = True
  
  SetIcon Me.hWnd, "SIKD"
  GetRandom
  GetIPNumber cIPNumber, cNamaDatabase, cNamaDSN, cPort, cKeySecret, cModePelunasanPiutang
  
  'inisiasi printer thermal
  GetPrinterCMD cShellPrn
  GetMyODBCFile cMYODBCFile
  'Shell "cmd.exe net use lpt1: /delete /y", vbHide
  Shell "cmd.exe net use * /delete /y", vbHide
  Shell "cmd.exe /c " & cShellPrn, vbHide
  
  Me.Picture = LoadPicture(GetPicture(GetRegistry(reg_Wallpaper)))
'  cUID = "root"
'  cPwd = "www.nusa.id/12012019"

  cUID = "kode"
  cPwd = "FullMoon"
  CreateDSN cNamaDSN, cIPNumber, cNamaDatabase, cUID, cPwd, cPort, cMYODBCPATH, cMYODBCFile
  SaveRegistry reg_ServerUID, cUID
  SaveRegistry reg_ServerPWD, cPwd
  SaveRegistry reg_KeySecret, UCase(cKeySecret)
  lFirst = True
  cKode = ""
  InitConnection
  'InitCfg

  If lGetConfig = False Then
    MsgBox "Maaf configurasi sistem anda belum sempurna" & vbCrLf & "Silahkan sempurnakan terlebih dahulu sebelum program dipakai"
  End If
  
  'memberikan nilai pada konfigurasi default

  If aCfg(objData, msKelipatan) < 100000 Then
    UpdCfg msKelipatan, 200000, objData, "Kelipatan", "Poin Kelipatan"
  End If
  
  If aCfg(objData, msTerm) < 5 Then
      UpdCfg msTerm, 5, objData, "Day Term", "Expire Date Poin"
  End If
  lGetTransaction objData
  On Error Resume Next
End Sub

Public Sub GetExpiredApp()
Dim cSQL As String
Dim db As New ADODB.Recordset
Dim dTglkey As Date
Dim dTglTrs As Date
On Error GoTo Error:
  
  
  cSQL = " select * from keyapp where keyapp = '" & GetRegistry(reg_KeySecret) & "'"
  cSQL = cSQL & " ORDER BY id DESC LIMIT 0,1"
  
  Set db = objData.SQL(GetDSN, cSQL)
  If Not db.EOF Then
    'jika punya serial maka bandingkan dengan data
    dTglkey = CryptRC4(FromHexDump(GetNull(db!tokenapp)), GetRegistry(reg_KeySecret))
    cSQL = "select idbukubesar,tgl from bukubesar order by tgl desc limit 0,1"
    Set db = objData.SQL(GetDSN, cSQL)
    If Not db.EOF Then
      dTglTrs = GetNull(db!tgl)
      If DateDiff("d", dTglTrs, dTglkey) < 0 Or DateDiff("d", getTglServer, dTglkey) < 0 Then
        MsgBox "Serial Key Sudah Expired, Masukkan Serial Key yg Baru", vbCritical, "Error"
        LoadSerial
      Else
        nToExpired = DateDiff("d", dTglTrs, dTglkey)
        If nToExpired <= 5 Then
          MsgBox "Pengguna aplikasi yg terhormat. Masa aktif serial key Anda akan segera habis " & nToExpired & " hari lagi" & vbCrLf & "Silahkan diperbaharui supaya aplikasi tetap bisa digunakan. Terimakasih", vbExclamation + vbCritical
        End If
      End If
    End If
  Else
    'jika tidak punya data serial maka masukkan data serial
    MsgBox "Anda belum memilik serial number untuk aplikasi ini" & vbCrLf & "Masukkan Serial Key yg Baru", vbCritical, "Error"
    LoadSerial
  End If

Exit Sub
Error:
  MsgBox "Serial Key Tidak Valid", vbCritical, "Eror"
  LoadSerial
  End
End Sub

Private Function getTglServer() As Date
Dim db As New ADODB.Recordset
  
  Set db = objData.SQL(GetDSN, "select now() as tgl")
  If Not db.EOF Then
    getTglServer = GetNull(db!tgl)
  End If
End Function

Private Sub LoadSerial()
  Load trRegister
  trRegister.Show vbModal
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  Unload trPenjualan
  GetCloseApps cAppsName
End Sub

Private Sub mhuPackingInventory_Click()
  Load trPacking
  trPacking.Show
End Sub

Private Sub mnuAbout_Click()
'  Load frmAbout
'  frmAbout.Show vbModal
  Load trRegister
  trRegister.Show vbModal
End Sub

Private Sub mnuAllSalesDetails_Click()
  Load rptAllSalesDetail
  rptAllSalesDetail.Show
  rptAllSalesDetail.SetFocus
End Sub

Private Sub mnuBuatAkunBaru_Click()
  Load mstRekening
  mstRekening.Show
  mstRekening.SetFocus
End Sub

Private Sub mnuBuyBack_Click()
  Load trBuyBack
  trBuyBack.Show
End Sub

Private Sub MnuCfgSetPrinter_Click()
  'CommonDialog1.ShowPrinter
End Sub

Private Sub MnuChangePassword_Click()
  objMenu.ChangePassword cKode
End Sub

Private Sub mnuChartOfAccount_Click()
  Load rptCOA
  rptCOA.Show
  rptCOA.SetFocus
End Sub

Private Sub mnuDaftarSupplier_Click()
  Load rptDaftarSupplier
  rptDaftarSupplier.Show
  rptDaftarSupplier.SetFocus
End Sub

Private Sub mnuDepartment_Click()
  Load mstDepartment
  mstDepartment.Show
  mstDepartment.SetFocus
End Sub

Private Sub mnuDetailPenjualanItem_Click()
  Load rptDetailItemPenjualanPerAnggota
  rptDetailItemPenjualanPerAnggota.Show
  rptDetailItemPenjualanPerAnggota.SetFocus
End Sub

Private Sub mnuExit_Click()
  Unload Me
End Sub

Private Sub mnuGeneralSettingAkuntansi_Click()
  Load cfgAutoJurnal
  cfgAutoJurnal.Show
  cfgAutoJurnal.SetFocus
End Sub

Private Sub mnuGeneralSettings_Click()
'  Load trOptions
'  trOptions.Show
'  trOptions.SetFocus
  
  Load trOpsiPenjualanPembelian
  trOpsiPenjualanPembelian.Show
  trOpsiPenjualanPembelian.SetFocus
End Sub

Private Sub mnuGolonganInventory_Click()
  Load mstNewGolongan
  mstNewGolongan.Show
End Sub

Private Sub mnuGrossSaleSalesmanReport_Click()
  Load rptGrossSales
  rptGrossSales.Show
  rptGrossSales.SetFocus
End Sub

Private Sub mnuGrossSalesReport_Click()
  Load rptGrossProfit
  rptGrossProfit.Show
  rptGrossProfit.SetFocus
End Sub

Private Sub mnuHargaGrosir_Click()
  Load mstHargaGrosir
  mstHargaGrosir.Show
End Sub

Private Sub mnuInputCatatanPelanggan_Click()
  Load trCatatanPelanggan
  trCatatanPelanggan.Show
  trCatatanPelanggan.SetFocus
End Sub

Private Sub mnuInventoryList_Click()
  Load rptInventoryList
  rptInventoryList.Show
  rptInventoryList.SetFocus
End Sub

Private Sub mnuJurnalUmum_Click()
  Load trJurnalUmum
  trJurnalUmum.Show
  trJurnalUmum.SetFocus
End Sub

Private Sub mnuKategori_Click()
  Load mstKategori
  mstKategori.Show
  mstKategori.SetFocus
End Sub

Private Sub mnuKelompokHarga_Click()
  Load mstHargaKategori
  mstHargaKategori.Show
  mstHargaKategori.SetFocus
End Sub

Private Sub mnuLaporanBarangMasuk_Click()
  Load rptBarangMasuk
  rptBarangMasuk.Show
  rptBarangMasuk.SetFocus
End Sub

Private Sub mnuLaporanBukuBesar_Click()
  Load rptBukuBesar
  rptBukuBesar.Show
  rptBukuBesar.SetFocus
End Sub

Private Sub mnuLaporanDaftarCustomer_Click()
  Load rptDaftarCustomer
  rptDaftarCustomer.Show
  rptDaftarCustomer.SetFocus
End Sub

Private Sub mnuLaporanDaftarStock_Click()
  'Unload rptStock
  Load PilihGudangRptStock
  PilihGudangRptStock.Show ' vbModal
End Sub

Private Sub mnuLaporanDetailPajakPenjualan_Click()
  Load rptDetailPajakPenjualan
  rptDetailPajakPenjualan.Show
End Sub

Private Sub mnuLaporanDetailPenjualan_Click()
  Load rptPenjualan
  rptPenjualan.Show
  rptPenjualan.SetFocus
End Sub

Private Sub mnuLaporanDetailPenjualanNonInventory_Click()
  Load rptPenjualanNonInventory
  rptPenjualanNonInventory.Show
End Sub


Private Sub mnuLaporanGeneralLedger_Click()
  Load rptJurnalUmum
  rptJurnalUmum.Show
  rptJurnalUmum.SetFocus
End Sub

Private Sub mnuLaporanJatuhTempoHutang_Click()
  Load rptJatuhTempoHutang
  rptJatuhTempoHutang.Show
  rptJatuhTempoHutang.SetFocus
End Sub

Private Sub mnuLaporanJatuhTempoPiutang_Click()
  Load rptJatuhTempoPiutang
  rptJatuhTempoPiutang.Show
  rptJatuhTempoPiutang.SetFocus
End Sub

Private Sub mnuLaporanKartuHutang_Click()
  Load rptKartuHutang
  rptKartuHutang.Show
  rptKartuHutang.SetFocus
End Sub

Private Sub mnuLaporanKartuPiutang_Click()
  Load rptKartuPiutang
  rptKartuPiutang.Show
  rptKartuPiutang.SetFocus
End Sub

Private Sub mnuLaporanKartuStock_Click()
  Load rptKartuStock
  rptKartuStock.Show
  rptKartuStock.SetFocus
End Sub

Private Sub mnuLaporanKartuTopUp_Click()
  Load rptKartuTopUp
  rptKartuTopUp.Show
End Sub

Private Sub mnuLaporanLabaRugiUsaha_Click()
  Load rptLabaRugiUpdate
  rptLabaRugiUpdate.Show
  rptLabaRugiUpdate.SetFocus
End Sub

Private Sub mnuLaporanMemberTopUp_Click()
  Load rptMemberTopUp
  rptMemberTopUp.Show
End Sub

Private Sub mnuLaporanMinimumStock_Click()
  Load rptUnderStock
  rptUnderStock.Show
End Sub

Private Sub mnuLaporanNeraca_Click()
  Load rptNeracaUpdate
  rptNeracaUpdate.Show
  rptNeracaUpdate.SetFocus
End Sub

Private Sub mnuLaporanNilaiPersediaanStock_Click()
  Load rptNilaiPersediaan
  rptNilaiPersediaan.Show
  rptNilaiPersediaan.SetFocus
End Sub

Private Sub mnuLaporanPelunasanHutang_Click()
  Load rptPelunasanHutang
  rptPelunasanHutang.Show
  rptPelunasanHutang.SetFocus
End Sub

Private Sub mnuLaporanPelunasanPiutang_Click()
  Load rptPelunasanPiutang
  rptPelunasanPiutang.Show
  rptPelunasanPiutang.SetFocus
End Sub

Private Sub mnuLaporanPembelianDetail_Click()
  Load rptPembelian
  rptPembelian.Show
  rptPembelian.SetFocus
End Sub

Private Sub mnuLaporanPenjualanBrutoHarian_Click()
  Load rptBrutoHarian
  rptBrutoHarian.Show
  rptBrutoHarian.SetFocus
End Sub

Private Sub mnuLaporanPenjualanHarian_Click()
  Load rptPenjualanHarian
  rptPenjualanHarian.Show
  rptPenjualanHarian.SetFocus
End Sub

Private Sub mnuLaporanPenjualanKonsinyasi_Click()
  Load rptPenjualanKonsinyasi
  rptPenjualanKonsinyasi.Show
  rptPenjualanKonsinyasi.SetFocus
End Sub

Private Sub mnuLaporanPPnKeluaran_Click()
  Load rptPPnKeluaran
  rptPPnKeluaran.Show
  rptPPnKeluaran.SetFocus
End Sub

Private Sub mnuLaporanPPnMasukan_Click()
  Load rptPPnMasukan
  rptPPnMasukan.Show
  rptPPnMasukan.SetFocus
End Sub

Private Sub mnuLaporanProdukTerlaris_Click()
  Load rptProdukTerlaris
  rptProdukTerlaris.Show
End Sub

Private Sub mnuLaporanRekapPenjualanBelumLunas_Click()
  Load rptRekapPenjualanBelumLunas
  rptRekapPenjualanBelumLunas.Show
  rptRekapPenjualanBelumLunas.SetFocus
End Sub

Private Sub mnuLaporanReturPembelian_Click()
  Load rptReturPembelian
  rptReturPembelian.Show
  rptReturPembelian.SetFocus
End Sub

Private Sub mnuLaporanReturPenjualan_Click()
  Load rptReturPenjualan
  rptReturPenjualan.Show
  rptReturPenjualan.SetFocus
End Sub

Private Sub mnuLaporanReturPenjualanGroupBySalesman_Click()
  Load rptReturPenjualanGroupBySalesman
  rptReturPenjualanGroupBySalesman.Show
  rptReturPenjualanGroupBySalesman.SetFocus
End Sub

Private Sub mnuLaporanSaldoHutang_Click()
  Load rptSaldoHutang
  rptSaldoHutang.Show
  rptSaldoHutang.SetFocus
End Sub

Private Sub mnuLaporanSaldoPiutang_Click()
  Load rptSaldoPiutang
  rptSaldoPiutang.Show
  rptSaldoPiutang.SetFocus
End Sub

Private Sub mnuLaporanStockOpname_Click()
  Load rptStockOpname
  rptStockOpname.Show
  rptStockOpname.SetFocus
End Sub

Private Sub mnuLaporanTrialBalance_Click()
  Load rptNeracaPercobaan
  rptNeracaPercobaan.Show
  rptNeracaPercobaan.SetFocus
End Sub

Private Sub mnuLogOff_Click()
  Unload Me
  Me.Show
  GetGroupSalesPenjualan = ""
End Sub

Private Sub mnuMaintenanceDatabase_Click()
  If objMenu.GetPassword(cKode, Me, GetDSN) Then
    If objMenu.UserLevel = 0 Then
      Load frmMaintenance
      frmMaintenance.Show
      frmMaintenance.SetFocus
    Else
      MsgBox "Maaf, anda tidak diperkenankan mengakses menu ini"
    End If
  End If
End Sub

Private Sub mnuMarkupHarga_Click()
  Load cfgMarkUP
  cfgMarkUP.Show
  cfgMarkUP.SetFocus
End Sub

Private Sub mnuMarkUpHargaJual_Click()
  Load cfgMarkUP
  cfgMarkUP.Show
End Sub

Private Sub mnuMasterHadiah_Click()
  Load mstHadiah
  mstHadiah.Show
  mstHadiah.SetFocus
End Sub

Private Sub mnuMasterGroupSales_Click()
  Load mstGroupSales
  mstGroupSales.Show
  mstGroupSales.SetFocus
End Sub

Private Sub mnuMasterIsiDataCustomer_Click()
  Load mstCustomer
  mstCustomer.Show
  mstCustomer.SetFocus
End Sub

Private Sub mnuMasterIsiDataSupplier_Click()
  Load mstSupplier
  mstSupplier.Show
  mstSupplier.SetFocus
End Sub

Private Sub mnuMasterKeteranganBayar_Click()
  Load frmKeteranganBayar
  frmKeteranganBayar.Show
End Sub

Private Sub mnuMasterStock_Click()
  GetNotifikasiAdd "Membuka data stock"
  Load mstStock
  mstStock.Show
  mstStock.SetFocus
  GetNotifikasiRemove
End Sub

Private Sub mnuMemberBalance_Click()
  Load trMemberBalance
  trMemberBalance.Show
  trMemberBalance.SetFocus
End Sub

Private Sub mnuMemberTopUp_Click()
  Load trMemberTopUp
  trMemberTopUp.Show
End Sub

Private Sub mnuMemberUtility_Click()
  Load frmUtility
  frmUtility.Show
  frmUtility.SetFocus
End Sub

Private Sub mnuMemberWithdraw_Click()
  Load trMemberWithdraw
  trMemberWithdraw.Show
End Sub

Private Sub MnuMenuLevel_Click()
  objMenu.SisSetMenu Me, cKode, GetDSN
End Sub

Private Sub mnuMutasiKasDanBank_Click()
  Load trMutasiKasBank
  trMutasiKasBank.Show
  trMutasiKasBank.SetFocus
End Sub

Private Sub mnuMutasiStock_Click()
  Load rptMutasiStock
  rptMutasiStock.Show
  rptMutasiStock.SetFocus
End Sub

Private Sub mnuMutasiStockAntarGudang_Click()
  Load trMutasiStock
  trMutasiStock.Show
  trMutasiStock.SetFocus
End Sub

Private Sub mnuNonInventoryList_Click()
  Load rptInventoryList
  rptInventoryList.Show
  rptInventoryList.SetFocus
End Sub

Private Sub mnuOmzetSalesmanReport_Click()
  Load rptOmzetSalesman
  rptOmzetSalesman.Show
  rptOmzetSalesman.SetFocus
End Sub

Private Sub mnuOtorisasi_Click()
  Load cfgOtorisasi
  cfgOtorisasi.Show
  cfgOtorisasi.SetFocus
End Sub

Private Sub MnuPassword_Click()
  objMenu.AddPassword GetDSN, cKode
End Sub

Private Sub mnuPembatalanatauPenghapusanBG_Click()
  Load trPembatalanBG
  trPembatalanBG.Show
End Sub

Private Sub mnuPencairanBG_Click()
  Load trPencairanBG
  trPencairanBG.Show
End Sub

Private Sub mnuPengeluaranBiaya_Click()
  Load trPengeluaranBiaya
  trPengeluaranBiaya.Show
  trPengeluaranBiaya.SetFocus
End Sub

Private Sub mnuPrive_Click()
  Load trPrive
  trPrive.Show
  trPrive.SetFocus
End Sub

Private Sub mnuRefund_Click()
  Load trRefund
  trRefund.Show
End Sub

Private Sub mnuRegister_Click()
  Load trRegister
  trRegister.Show vbModal
End Sub

Private Sub mnuRekapitulasiOmzetSales_Click()
  Load rptRekapitulasiOmzetSales
  rptRekapitulasiOmzetSales.Show
  rptRekapitulasiOmzetSales.SetFocus
End Sub

Private Sub mnuRekeningLabaRugiTahunBerjalan_Click()
  Load cfgRekeningLabaRugi
  cfgRekeningLabaRugi.Show
  cfgRekeningLabaRugi.SetFocus
End Sub

Private Sub mnuRekeningNilaiPersediaan_Click()
  Load cfgRekeningNilaiPersediaan
  cfgRekeningNilaiPersediaan.Show
  cfgRekeningNilaiPersediaan.SetFocus
End Sub

Private Sub mnuReturKonsinyasi_Click()
  Load trReturKonsinyasi
  trReturKonsinyasi.Show
  trReturKonsinyasi.SetFocus
End Sub

Private Sub mnuRptStockKonsinyasi_Click()
  Load rptStockKonsinyasi
  rptStockKonsinyasi.Show
  rptStockKonsinyasi.SetFocus
End Sub

Private Sub mnuSalesmanList_Click()
Dim n As Integer
Dim cSQL As String
Dim vaArray As New XArrayDB

  cSQL = cSQL & "Select * from salesman order by kodesalesman"
   
  vaArray.ReDim 0, -1, 0, 3
  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!kodesalesman)
      vaArray(n, 1) = GetNull(dbData!nama)
      vaArray(n, 2) = GetNull(dbData!alamat)
      vaArray(n, 3) = GetNull(dbData!telp)
      dbData.MoveNext
    Loop
    With FrmRPT
      .AddPageHeader "SALESMAN LIST", tdbHalignCenter, , , True, dbArial, 12, True, , , False
      .AddPageHeader aCfg(objData, msNamaPerusahaan), tdbHalignCenter, , , True, , 14
      .AddPageHeader "", , , , True
      .AddPageHeader "", , , , True
      
      .AddTableHeader "KODE", , , , 7
      .AddTableHeader "NAMA", , , , 30
      .AddTableHeader "ALAMAT"
      .AddTableHeader "TELEPON", , , , 20
       
      .AddTableBody
      .AddTableBody
      .AddTableBody
      .AddTableBody
                 
      .Preview vaArray, True
    End With
  Else
    MsgBox "Maaf, tidak ada data salesman untuk ditampilkan."
  End If
End Sub

Private Sub mnuSatuanInventory_Click()
  Load mstSatuan
  mstSatuan.Show
  mstSatuan.SetFocus
End Sub

Private Sub mnuSensusSalesHarian_Click()
  Load rptSensusHarian
  rptSensusHarian.Show
  rptSensusHarian.SetFocus
End Sub

Private Sub mnuSettingCostCentre_Click()
  Load cfgCostCenter
  cfgCostCenter.Show
  cfgCostCenter.SetFocus
End Sub

Private Sub mnuSetupAkunKas_Click()
  Load mstAkunKas
  mstAkunKas.Show
  mstAkunKas.SetFocus
End Sub

Private Sub mnuSetupCostCenter_Click()
  Load mstCostCenter
  mstCostCenter.Show
  mstCostCenter.SetFocus
End Sub

Private Sub MnuSetupInfoPerusahaan_Click()
  Load cfgInformasiPerusahaan
  cfgInformasiPerusahaan.Show
  cfgInformasiPerusahaan.SetFocus
End Sub

Private Sub mnuSetupKartu_Click()
  Load trKartuKredit
  trKartuKredit.Show
  trKartuKredit.SetFocus
End Sub

Private Sub mnuSetupPortPrinter_Click()
  Load cfgPortPrinter
  cfgPortPrinter.Show
  cfgPortPrinter.SetFocus
End Sub

Private Sub mnuSetupRekeningBiaya_Click()
  GetNotifikasiAdd "Membuka Setup Rekening Biaya"
  Load mstAkunBiaya
  mstAkunBiaya.Show
  mstAkunBiaya.SetFocus
  GetNotifikasiRemove
End Sub

Private Sub mnuSetupSalesman_Click()
  Load mstSalesman
  mstSalesman.Show
  mstSalesman.SetFocus
End Sub

Private Sub mnuStockKontrak_Click()
  Load trKontrak
  trKontrak.Show
End Sub

Private Sub mnuSupplierBalance_Click()
  Load trSupplierBalance
  trSupplierBalance.Show
  trSupplierBalance.SetFocus
End Sub

Private Sub mnuTarifCOD_Click()
  Load mstCOD
  mstCOD.Show
  mstCOD.SetFocus
End Sub

Private Sub mnuTransaksiPembayaranHutang_Click()
  Load trPelunasanHutang
  trPelunasanHutang.Show
  trPelunasanHutang.SetFocus
End Sub

Private Sub mnuTransaksiPembayaranPiutang_Click()
  
  Unload trPelunasanPiutang
  Unload trPelunasanHutangSederhana
  
  Select Case GetRegistry(reg_ModePelunasanPiutang)
    Case 1
      'ack
      Load trPelunasanPiutang
      trPelunasanPiutang.Show
    Case 2
      'normal
      Load trSelectGroupSales
      trSelectGroupSales.sisModul = enum_OpenPelunasanPiutang
      trSelectGroupSales.Show vbModal
      
  End Select
End Sub

Private Sub OpenPembelian()
  Load trSelectGroupSales
  trSelectGroupSales.cGroupSales.Text = IIf(GetGroupSales(GetRegistry(reg_KodeGroupSalesPembelian), objData) <> "", GetRegistry(reg_KodeGroupSalesPembelian), "")
  trSelectGroupSales.sisModul = enum_OpenPembelian
  trSelectGroupSales.Show vbModal
End Sub

Private Sub mnuTransaksiPembelianTanpaOrder_Click()

  OpenPembelian

'  If GetRegistry(reg_OptGroupSales) = 2 Then
'    GetGroupSalesPenjualan = GetRegistry(reg_KodeGroupPenjualan)
'    If lExist(objData, "groupsales", "kode", GetGroupSalesPenjualan, " and status=1") = True Then
'      SaveRegistry reg_KodeGroupPenjualan, GetGroupSalesPenjualan
'      aMainmenu.MembukaModulPenjualan
'    Else
'      OpenPembelian
'    End If
'  Else
'    OpenPembelian
'  End If
  

  

'  If aCfg(objData, msModelInputPembelian) = "M" Then
'    ' Modul Pembelian Mahir
'    Unload trPelunasanHutang
'    Load trPembelianNonTunai
'
'    trPembelianNonTunai.Show
'    trPembelianNonTunai.SetFocus
'
'  Else
'    ' Modul Pembelian Standar
'    Unload trPelunasanHutang
'    Load trPembelianNonTunaiOri
'
'    trPembelianNonTunaiOri.Show
'    trPembelianNonTunaiOri.SetFocus
'  End If

'  Load trPembelianKonsinyasi
'  trPembelianKonsinyasi.Show 'vbModal
End Sub

Public Sub MembukaModulPelunasan()
  GetNotifikasiAdd "Membuka Modul Pelunasan Piutang "
  Load trPelunasanHutangSederhana
  trPelunasanHutangSederhana.Show
  trPelunasanHutangSederhana.SetFocus
  GetNotifikasiRemove
End Sub

Public Sub MembukaModulPembelian()
  GetNotifikasiAdd "Membuka Modul Pembelian"
  If aCfg(objData, msModelInputPembelian) = "M" Then
    ' Modul Pembelian Mahir
    Unload trPelunasanHutang
    Load trPembelianNonTunai

    trPembelianNonTunai.Show
    trPembelianNonTunai.SetFocus

  Else
    ' Modul Pembelian Standar
    Unload trPelunasanHutang
    Load trPembelianNonTunaiOri

    trPembelianNonTunaiOri.Show
    trPembelianNonTunaiOri.SetFocus
  End If
  GetNotifikasiRemove
End Sub

Private Sub mnuTransaksiPenjualan_Click()
'  If GetRegistry(reg_OptGroupSales) = 1 Then 'ya
'    Load trSelectGroupSales
'    trSelectGroupSales.sisModul = enum_OpenPenjualan
'    trSelectGroupSales.Show vbModal
'  Else
'    If GetRegistry(reg_KodeGroupSales) <> "" Then
'      MembukaModulPenjualan
'    Else
'      MsgBox "Kode Group Sales Belum di setting", vbCritical
'      trSelectGroupSales.Show vbModal
'    End If
'  End If

  'cek dulu status
  If GetRegistry(reg_OptGroupSales) = 2 Then
    GetGroupSalesPenjualan = GetRegistry(reg_KodeGroupPenjualan)
    If lExist(objData, "groupsales", "kode", GetGroupSalesPenjualan, " and status=1") = True Then
      SaveRegistry reg_KodeGroupPenjualan, GetGroupSalesPenjualan
      aMainmenu.MembukaModulPenjualan
    Else
      OpenPenjualan
    End If
  Else
    OpenPenjualan
  End If
  
'  If Trim(GetGroupSalesPenjualan) <> "" Then
'    'cek apakah kode groupsales nya betul
'    'jika betul langsung open modul penjualan
'    If lExist(objData, "groupsales", "kode", GetGroupSalesPenjualan, " and status=1") = True Then
'      SaveRegistry reg_KodeGroupPenjualan, GetGroupSalesPenjualan
'      aMainmenu.MembukaModulPenjualan
'    Else
'      If GetRegistry(reg_OptGroupSales) = 1 Then
'        OpenPenjualan
'      End If
'    End If
'  Else
'    OpenPenjualan
'  End If
End Sub

Private Sub OpenPenjualan()
    Load trSelectGroupSales
    trSelectGroupSales.cGroupSales.Text = GetRegistry(reg_KodeGroupPenjualan)
    trSelectGroupSales.cGroupSales.Text = IIf(GetGroupSales(GetRegistry(reg_KodeGroupPenjualan), objData) <> "", GetRegistry(reg_KodeGroupPenjualan), "")
    trSelectGroupSales.sisModul = enum_OpenPenjualan
    trSelectGroupSales.Show vbModal
End Sub

Public Sub MembukaModulPenjualan()
  GetNotifikasiAdd "Membuka Modul Penjualan "
  Load trPenjualan
  trPenjualan.Show
  trPenjualan.SetFocus
  GetNotifikasiRemove
End Sub

Private Sub mnuTransaksiPenyesuaianStock_Click()
  Load trNewPenyesuaianStock
  trNewPenyesuaianStock.Show
End Sub

Private Sub mnuTransaksiReturPenjualan_Click()
'Dim frm As Form
'Set frm = Me.ActiveForm
'frm.Visible = False 'or frm.hide should work
'Set frm = Nothing

  Load trReturPenjualan
  trReturPenjualan.Show
  trReturPenjualan.SetFocus
End Sub

Private Sub mnuTrKomplimen_Click()
  Load trKomplimen
  trKomplimen.Show
  trKomplimen.SetFocus
End Sub

Private Sub mnuTrReturPembelian_Click()
  Load trReturPembelian
  trReturPembelian.Show
  trReturPembelian.SetFocus
End Sub

Private Sub mnuTruncateDatabase_Click()
  Load frmMaintenance
  frmMaintenance.Show
  frmMaintenance.SetFocus
End Sub

Private Sub mnuUpdateKartuPiutangMember_Click()
  Load trUpdateKartuPiutangMember
  trUpdateKartuPiutangMember.Show
End Sub

Private Sub mnuUtilityforStock_Click()
  Load frmUtilityStock
  frmUtilityStock.Show
  frmUtilityStock.SetFocus
End Sub

Private Sub mnuUtilityForSupplier_Click()
'  Load frmSupplierUtility
'  frmSupplierUtility.Show
'  frmSupplierUtility.SetFocus
Load trStockOpname
trStockOpname.Show
End Sub

Private Sub mnuWallpaper_Click()
  GetWallpaper Me
End Sub

Private Sub mnuWindowFullScreen_Click()
  If mnuWindowFullScreen.Caption = "Full Screen" Then
    EnableMaxButton Me.hWnd, False
    lWIndowStatus = 1
    mnuWindowFullScreen.Caption = "Normal Screen"
  Else
    EnableMaxButton Me.hWnd, True
    lWIndowStatus = 0
    mnuWindowFullScreen.Caption = "Full Screen"
  End If
  Me.WindowState = vbMaximized
End Sub

Private Sub MnuWindows_Click(Index As Integer)
  Select Case Index
    Case 0
      Me.Arrange vbTileHorizontal
    Case 1
      Me.Arrange vbVertical
    Case 2
      Me.Arrange vbCascade
  End Select
End Sub

Private Sub mstDatabaseGudang_Click()
  Load mstGudang
  mstGudang.Show
  mstGudang.SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.key
    Case "Close"
      mnuExit_Click
    Case "Printer"
      'MnuCfgSetPrinter_Click
'      GetPrinterCMD cShellPrn
'
'      MsgBox "Perintah : " & cShellPrn & " Akan dieksekusi!!"
'      Shell "cmd.exe /c " & " net use lpt1: /delete"
'      Shell "cmd.exe /c " & "printer.bat"
'      getDosCMD
'      Shell App.Path & "\printer.bat", vbHide
      GetWallpaper Me
    Case "About"
      mnuAbout_Click
    Case "LogOff"
      mnuLogOff_Click
    Case "Calculator"
      Shell "Calc"
    Case "Help"
      
  End Select
End Sub


Private Sub getDosCMD()
  'CreateObject("WScript.Shell").Run "%COMSPEC% /c " & "D:\Dropbox\ack center rubah faktur pelunasan\pos.bat", 4, False
  Shell Environ$("COMSPEC") & " /C " & "printer.bat", vbNormalFocus
End Sub

Private Sub GetWallpaper(cForm As Form)
  On Error GoTo EmptyPicture:
  CommonDialog1.Filter = "Picture (*.BMP;*.JPG;*.GIF) |*.BMP;*.JPG;*.GIF|"
  CommonDialog1.FileName = GetRegistry(reg_Wallpaper)
  CommonDialog1.Action = 1
  If Trim(CommonDialog1.FileName) <> "" And Dir(CommonDialog1.FileName) <> "" Then
    cForm.Picture = LoadPicture(GetPicture(CommonDialog1.FileName))
    cForm.Hide
    cForm.Show
  End If
  SaveRegistry reg_Wallpaper, CommonDialog1.FileName
  Exit Sub
  
EmptyPicture:

  CommonDialog1.FileName = ""
  cForm.Picture = LoadPicture("")
  Resume Next
End Sub


'Function untuk disable/enable mdi button
Public Function EnableCloseButton(ByVal hWnd As Long, Enable As Boolean) As Integer
    EnableSystemMenuItem hWnd, SC_CLOSE, xSC_CLOSE, Enable, "EnableCloseButton"
End Function

'*******************************************************************************
' Enable / Disable Minimise Button
'-------------------------------------------------------------------------------

Public Sub EnableMinButton(ByVal hWnd As Long, Enable As Boolean)
    Dim lngFormStyle As Long

    ' Enable / Disable System Menu Item
    EnableSystemMenuItem hWnd, SC_MINIMIZE, xSC_MINIMIZE, Enable, "EnableMinButton"

    ' Enable / Disable TitleBar button

    lngFormStyle = GetWindowLong(hWnd, GWL_STYLE)
    If Enable Then
        lngFormStyle = lngFormStyle Or WS_MINIMIZEBOX
    Else
        lngFormStyle = lngFormStyle And Not WS_MINIMIZEBOX
    End If
    SetWindowLong hWnd, GWL_STYLE, lngFormStyle

    ' Dirty, slimy, devious hack to ensure that the changes to the
    ' window's style take immediate effect before the form is shown

    SetParent hWnd, GetParent(hWnd)
    SetWindowPos hWnd, hWnd_NOTOPMOST, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_FRAMECHANGED
End Sub

'*******************************************************************************
' Enable / Disable Maximise Button
'-------------------------------------------------------------------------------

Public Sub EnableMaxButton(ByVal hWnd As Long, Enable As Boolean)
    Dim lngFormStyle As Long

    ' Enable / Disable System Menu Item
    EnableSystemMenuItem hWnd, SC_MAXIMIZE, xSC_MAXIMIZE, Enable, "EnableMaxButton"

    ' Enable / Disable TitleBar button
    lngFormStyle = GetWindowLong(hWnd, GWL_STYLE)
    If Enable Then
        lngFormStyle = lngFormStyle Or WS_MAXIMIZEBOX
    Else
        lngFormStyle = lngFormStyle And Not WS_MAXIMIZEBOX
    End If
    SetWindowLong hWnd, GWL_STYLE, lngFormStyle

    ' Dirty, slimy, devious hack to ensure that the changes to the
    ' window's style take immediate effect before the form is shown

    SetParent hWnd, GetParent(hWnd)
    SetWindowPos hWnd, hWnd_NOTOPMOST, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_FRAMECHANGED
End Sub

Private Sub EnableSystemMenuItem(hWnd As Long, Item As Long, Dummy As Long, Enable As Boolean, FuncName As String)
On Error Resume Next

    Dim hMenu       As Long
    Dim MII         As MENUITEMINFO
    Dim lngMenuID   As Long

    If IsWindow(hWnd) = 0 Then
        err.Raise vbObjectError, "modCloseBtn::" & FuncName, _
            "modCloseBtn::" & FuncName & "() - Invalid Window Handle"
        Exit Sub
    End If

    ' Retrieve a handle to the window's system menu
    hMenu = GetSystemMenu(hWnd, 0)

    ' Retrieve the menu item information for the Max menu item/button
    MII.cbSize = Len(MII)
    MII.dwTypeData = String$(80, 0)
    MII.cch = Len(MII.dwTypeData)
    MII.fMask = MIIM_STATE

    If Enable Then
        MII.wID = Dummy
    Else
        MII.wID = Item
    End If

    If GetMenuItemInfo(hMenu, MII.wID, False, MII) = 0 Then
        err.Raise vbObjectError, "modCloseBtn::" & FuncName, _
            "modCloseBtn::" & FuncName & "() - Menu Item Not Found"
        Exit Sub
    End If

    ' Switch the ID of the menu item so that VB can not undo the action itself
    lngMenuID = MII.wID

    If Enable Then
        MII.wID = Item
    Else
        MII.wID = Dummy
    End If

    MII.fMask = MIIM_ID
    If SetMenuItemInfo(hMenu, lngMenuID, False, MII) = 0 Then
        err.Raise vbObjectError, "modCloseBtn::" & FuncName, _
            "modCloseBtn::" & FuncName & "() - Error encountered " & _
            "changing ID"
        Exit Sub
    End If

    ' Set the enabled / disabled state of the menu item
    If Enable Then
        MII.fState = MII.fState And Not MFS_GRAYED
    Else
        MII.fState = MII.fState Or MFS_GRAYED
    End If

    MII.fMask = MIIM_STATE
    If SetMenuItemInfo(hMenu, MII.wID, False, MII) = 0 Then
         err.Raise vbObjectError, "modCloseBtn::" & FuncName, _
            "modCloseBtn::" & FuncName & "() - Error encountered " & _
            "changing state"
        Exit Sub
    End If

    ' Activate the non-client area of the window to update the titlebar, and
    ' draw the Max button in its new state.
    SendMessage hWnd, WM_NCACTIVATE, True, 0
End Sub
