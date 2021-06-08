VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptLabaRugi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LABA RUGI"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   6990
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   915
      Left            =   0
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   1614
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   4
      BackColor       =   -2147483633
      Begin BiSADateProject.BiSADate dDate 
         Height          =   330
         Index           =   0
         Left            =   390
         TabIndex        =   0
         Top             =   300
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   582
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
         Caption         =   "ANTARA  TANGGAL"
         CaptionWidth    =   2000
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
      Begin BiSADateProject.BiSADate dDate 
         Height          =   330
         Index           =   1
         Left            =   4005
         TabIndex        =   1
         Top             =   315
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   582
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
         Caption         =   "S.D"
         CaptionWidth    =   500
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
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   645
      Left            =   0
      Top             =   915
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   1138
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   4
      BackColor       =   -2147483633
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   5775
         TabIndex        =   2
         Top             =   120
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   767
         Caption         =   "     &Exit"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         BackColor       =   -2147483633
         Picture         =   "rptLabaRugi.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   5340
         TabIndex        =   3
         Top             =   120
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   767
         Caption         =   ""
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         BackColor       =   -2147483633
         Picture         =   "rptLabaRugi.frx":00A6
      End
   End
End
Attribute VB_Name = "rptLabaRugi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
  
  nLabaAkhir = nLabaRugiUsaha + GetKasMasuk(obj, dAwal, dAkhir) + GetAbsMin(GetKasKeluar(obj, dAwal, dAkhir))
  
  GetVarpt 21, "VIII", "Kas Masuk", "", , , GetKasMasuk(obj, dAwal, dAkhir)
  GetVarpt 22, "VIX", "Kas Keluar", "", , , GetKasKeluar(obj, dAwal, dAkhir)
  GetVarpt 23, "X", "Laba/rugi Bersih", "", , , nLabaAkhir
  GetBarisKosong 24
  
  If lPreview = True Then
    GetPreview dAwal, dAkhir
  End If
  GetLabaRugiNetto = nLabaAkhir
End Function

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

Private Function GetSpasi(Optional ByVal nNumber As Integer = 1) As String
Dim n As Integer

  GetSpasi = ""
  For n = 1 To nNumber
    GetSpasi = GetSpasi & vbTab
  Next
End Function

Private Function GetAbsMin(ByVal nNumber As Double) As Double
  GetAbsMin = 0 - nNumber
End Function

Private Sub GetBarisKosong(ByVal nBaris As Integer)
  vaRpt(nBaris, 0) = " "
  vaRpt(nBaris, 1) = " "
  vaRpt(nBaris, 2) = " "
  vaRpt(nBaris, 3) = " "
  vaRpt(nBaris, 4) = " "
  vaRpt(nBaris, 5) = " "
End Sub

Private Sub cmdKeluar_Click()
Unload Me
End Sub

Private Sub cmdPreview_Click()
  GetLabaRugiNetto objData, dDate(0).Value, dDate(1).Value, True
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd
  CenterForm Me
  dDate(0).Value = BOM(Date)
  dDate(1).Value = EOM(Date)
  TabIndex dDate(0), n
  TabIndex dDate(1), n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub
