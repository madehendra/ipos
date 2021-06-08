VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trMutasiSimpanan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mutasi Simpanan Harian"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   10935
   Begin BiSAFramProject.BiSAFrame FrameMutasiTabungan 
      Height          =   4545
      Left            =   0
      Top             =   0
      Width           =   10950
      _ExtentX        =   19315
      _ExtentY        =   8017
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      Begin BiSAFramProject.BiSAFrame BiSAFrame2 
         Height          =   4455
         Left            =   6330
         Top             =   60
         Width           =   4485
         _ExtentX        =   7911
         _ExtentY        =   7858
         Caption         =   "MUTASI"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         BackColor       =   -2147483633
         Begin BiSANumberBoxProject.BiSANumberBox nAwal 
            Height          =   420
            Left            =   135
            TabIndex        =   0
            Top             =   390
            Width           =   4305
            _ExtentX        =   7594
            _ExtentY        =   741
            Appearance      =   0
            Enabled         =   0   'False
            BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   12632256
            Caption         =   "SALDO AWAL"
            CaptionWidth    =   1600
            BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin BiSANumberBoxProject.BiSANumberBox nMutasi 
            Height          =   420
            Left            =   135
            TabIndex        =   1
            Top             =   855
            Width           =   4305
            _ExtentX        =   7594
            _ExtentY        =   741
            BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "JUMLAH"
            CaptionWidth    =   1600
            BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin BiSANumberBoxProject.BiSANumberBox nAkhir 
            Height          =   420
            Left            =   135
            TabIndex        =   2
            Top             =   1335
            Width           =   4305
            _ExtentX        =   7594
            _ExtentY        =   741
            Appearance      =   0
            Enabled         =   0   'False
            BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   12632256
            Caption         =   "SALDO AKHIR"
            CaptionWidth    =   1600
            BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin BiSATextBoxProject.BiSATextBox cKeteranganTabungan 
         Height          =   330
         Left            =   120
         TabIndex        =   3
         Top             =   2850
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   582
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Verdana"
         Caption         =   "Keterangan"
         CaptionWidth    =   1500
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekeningJurnal 
         Height          =   330
         Left            =   3390
         TabIndex        =   4
         Top             =   2490
         Width           =   2610
         _ExtentX        =   4604
         _ExtentY        =   582
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Verdana"
         BackColor       =   12632256
         Enabled         =   0   'False
         Appearance      =   0
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
      Begin BiSATextBoxProject.BiSATextBox cDK 
         Height          =   330
         Left            =   135
         TabIndex        =   5
         Top             =   2145
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   582
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Verdana"
         BackColor       =   12632256
         Enabled         =   0   'False
         Appearance      =   0
         Caption         =   "D/K"
         CaptionWidth    =   1500
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
      Begin BiSATextBoxProject.BiSATextBox cNamaKodeTransaksi 
         Height          =   330
         Left            =   2565
         TabIndex        =   6
         Top             =   1785
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   582
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Verdana"
         BackColor       =   12632256
         Enabled         =   0   'False
         Appearance      =   0
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
      Begin BiSATextBoxProject.BiSABrowse cKodeTransaksi 
         Height          =   330
         Left            =   135
         TabIndex        =   7
         Top             =   1785
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Verdana"
         Button          =   -1  'True
         Caption         =   "Kode Transaksi"
         CaptionWidth    =   1500
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
      Begin BiSATextBoxProject.BiSABrowse cRekeningJurnal 
         Height          =   330
         Left            =   135
         TabIndex        =   8
         Top             =   2490
         Width           =   3240
         _ExtentX        =   5715
         _ExtentY        =   582
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Verdana"
         BackColor       =   12632256
         Enabled         =   0   'False
         Appearance      =   0
         Caption         =   "Rekening"
         CaptionWidth    =   1500
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
      Begin BiSAFramProject.BiSAFrame BiSAFrame1 
         Height          =   1740
         Left            =   45
         Top             =   30
         Width           =   6240
         _ExtentX        =   11007
         _ExtentY        =   3069
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         Begin BiSATextBoxProject.BiSATextBox cNama 
            Height          =   330
            Left            =   120
            TabIndex        =   19
            Top             =   840
            Width           =   5955
            _ExtentX        =   10504
            _ExtentY        =   582
            BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "Verdana"
            BackColor       =   12632256
            Enabled         =   0   'False
            Appearance      =   0
            Caption         =   "Nama"
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
         Begin BiSATextBoxProject.BiSATextBox cFaktur 
            Height          =   300
            Left            =   2610
            TabIndex        =   20
            Top             =   120
            Width           =   3435
            _ExtentX        =   6059
            _ExtentY        =   529
            Text            =   "12345678901234567890"
            BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "Verdana"
            BackColor       =   12632256
            Enabled         =   0   'False
            MaxLength       =   20
            Appearance      =   0
            Caption         =   "No Transaksi"
            CaptionWidth    =   1200
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
         Begin BiSADateProject.BiSADate dTgl 
            Height          =   330
            Left            =   135
            TabIndex        =   21
            Top             =   120
            Width           =   2430
            _ExtentX        =   4286
            _ExtentY        =   582
            Value           =   "13-10-2005"
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
            Caption         =   "Tanggal"
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
         Begin BiSATextBoxProject.BiSATextBox cAlamat 
            Height          =   330
            Left            =   120
            TabIndex        =   22
            Top             =   1200
            Width           =   5955
            _ExtentX        =   10504
            _ExtentY        =   582
            BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "Verdana"
            BackColor       =   12632256
            Enabled         =   0   'False
            Appearance      =   0
            Caption         =   "Alamat"
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
         Begin BiSATextBoxProject.BiSABrowse cKodeAnggota 
            Height          =   330
            Left            =   120
            TabIndex        =   23
            Top             =   480
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   582
            Text            =   "12345678"
            BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "Verdana"
            Appearance      =   0
            Button          =   -1  'True
            Caption         =   "Anggota"
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
      Begin VB.Label Label2 
         Caption         =   "[K] = Setoran    [D] = Penarikan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2325
         TabIndex        =   14
         Top             =   2190
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Saldo Minimum Mengendap"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   465
         TabIndex        =   13
         Top             =   3675
         Width           =   2850
      End
      Begin VB.Label Label4 
         Caption         =   "Saldo Min. Dpt Bunga"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   465
         TabIndex        =   12
         Top             =   3945
         Width           =   2850
      End
      Begin VB.Label lbSaldoMinimum 
         Alignment       =   1  'Right Justify
         Caption         =   "Not Available"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3480
         TabIndex        =   11
         Top             =   3675
         Width           =   2295
      End
      Begin VB.Label lbSaldoDapatBunga 
         Alignment       =   1  'Right Justify
         Caption         =   "Not Available"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3480
         TabIndex        =   10
         Top             =   3945
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "Catatan:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   165
         TabIndex        =   9
         Top             =   3435
         Width           =   720
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame5 
      Height          =   585
      Left            =   0
      Top             =   4530
      Width           =   10950
      _ExtentX        =   19315
      _ExtentY        =   1032
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   9780
         TabIndex        =   15
         Top             =   75
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   767
         Caption         =   "     E&xit"
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
         Picture         =   "trMutasiSimpanan.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdBatal 
         Height          =   435
         Left            =   7200
         TabIndex        =   16
         Top             =   75
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   767
         Caption         =   "    &Clear/Batal"
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
         Picture         =   "trMutasiSimpanan.frx":00A6
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   8700
         TabIndex        =   17
         Top             =   75
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   767
         Caption         =   "    &Save"
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
         Picture         =   "trMutasiSimpanan.frx":0330
      End
      Begin VB.Label Label6 
         Caption         =   "Esc = Keluar/Exit"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   135
         TabIndex        =   18
         Top             =   150
         Width           =   2130
      End
   End
End
Attribute VB_Name = "trMutasiSimpanan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub InitValue(Optional ByVal lRekening As Boolean = True)
  cFaktur.Default
  cNama.Default
  cAlamat.Default
  cKodeTransaksi.Default
  cNamaKodeTransaksi.Default
  cDK.Default
  cRekeningJurnal.Default
  cNamaRekeningJurnal.Default
  cKeteranganTabungan.Default
  nAwal.Default
  nMutasi.Default
  nAkhir.Default
  cKodeAnggota.Default
  
  'get info tabungan
'  lbSaldoMinimum.Caption = Format(aCfg(objData, msSaldoMinimum), "###,###,###,###,##0.00")
'  lbSetoranMinimum.Caption = Format(aCfg(objData, msSetoranMinimum), "###,###,###,###,##0.00")
'  lbSaldoDapatBunga.Caption = Format(aCfg(objData, msSaldoDapatBunga), "###,###,###,###,##0.00")
End Sub

Private Sub InitTabIndex()
Dim n As Single
  
  TabIndex dTgl, n
  TabIndex cFaktur, n
  TabIndex cKodeAnggota, n
  TabIndex cNama, n
  TabIndex cAlamat, n
  TabIndex cKodeTransaksi, n
  TabIndex cNamaKodeTransaksi, n
  TabIndex cDK, n
  TabIndex cRekeningJurnal, n
  TabIndex cNamaRekeningJurnal, n
  TabIndex cKeteranganTabungan, n
  TabIndex nAwal, n
  TabIndex nMutasi, n
  TabIndex nAkhir, n
  TabIndex cmdSimpan, n
  TabIndex cmdBatal, n
  TabIndex cmdKeluar, n
End Sub

Private Sub cKodeAnggota_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "anggota", "kodeanggota,nama,alamat", "kodeanggota", sisContent, cKodeAnggota.Text)
  If Not dbData.EOF Then
    cKodeAnggota.Text = cKodeAnggota.Browse(dbData, Array("Kode", "Nama", "Alamat"), , Array(8, 10, 20))
    cNama.Text = GetNull(dbData!nama)
    cAlamat.Text = GetNull(dbData!alamat)
    nAwal.Value = GetSaldoTabungan(objData, GetNull(dbData!kodeanggota), "01-01-1900", Date)
  End If
End Sub

Private Sub cKodeTransaksi_ButtonClick()
  Set dbData = objData.PICK(GetDSN, "KodeTransaksi k", "k.Kodetransaksi", cKodeTransaksi, "k.Kodetransaksi,k.Keterangan,k.DK,k.Kas,k.kodeakun")
  If Not dbData.EOF Then
    cNamaKodeTransaksi.Text = GetNull(dbData!Keterangan)
    cDK.Text = GetNull(dbData!DK)
    cRekeningJurnal.Default
    cNamaRekeningJurnal.Default
    If GetNull(dbData!kas) = "K" Then
      cRekeningJurnal.Text = cKasTeller
    Else
      cRekeningJurnal.Text = GetNull(dbData!kodeakun)
    End If
    Set dbData = objData.Browse(GetDSN, "akun", , "kodeakun", sisAssign, cRekeningJurnal.Text)
    If Not dbData.EOF Then
      cNamaRekeningJurnal.Text = GetNull(dbData!Keterangan, "")
    End If
    cKeteranganTabungan.Text = cNamaKodeTransaksi.Text & " a.n " & cNama.Text
    nMutasi_Change
  End If
End Sub

Private Sub cmdBatal_Click()
  InitValue
  dTgl.SetFocus
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdSimpan_Click()
Dim lSave As Boolean
Dim cRekeningDebet As String
Dim cRekeningKredit As String

  lSave = True
  If ValidSaving Then
     objData.Start GetDSN
     lSave = IIf(lSave, objData.Add(GetDSN, "mutasitabungan", Array("nomormutasitabungan", "kodetransaksi", "kodeakun", "username", "kodeanggota", "tgl", "jumlah", "datetime", "dk", "keterangan"), Array(cFaktur.Text, cKodeTransaksi.Text, cRekeningJurnal.Text, GetRegistry(reg_UserName), cKodeAnggota.Text, Format(dTgl.Value, "yyyy-MM-dd"), nMutasi.Value, SNow, cDK.Text, cKeteranganTabungan.Text)), False)
     If cDK.Text = "D" Then
      cRekeningDebet = aCfg(objData, msRekeningSimpananHarian)
      cRekeningKredit = cRekeningJurnal.Text
     Else
      cRekeningDebet = cRekeningJurnal.Text
      cRekeningKredit = aCfg(objData, msRekeningSimpananHarian)
     End If
     lSave = IIf(lSave, DelKodeTr(objData, msSimpananHarian, cFaktur.Text), False)
     lSave = IIf(lSave, UpdKodeTr(objData, msSimpananHarian, cFaktur.Text, Format(dTgl.Value, "yyyy-MM-dd"), cRekeningDebet, aCfg(objData, msCostCenterSimpanPinjam), cKeteranganTabungan.Text, nMutasi.Value, 0, cDK.Text, SNow), False)
        lSave = IIf(lSave, UpdKodeTr(objData, msSimpananHarian, cFaktur.Text, Format(dTgl.Value, "yyyy-MM-dd"), cRekeningKredit, aCfg(objData, msCostCenterSimpanPinjam), cKeteranganTabungan.Text, 0, nMutasi.Value, cDK.Text, SNow), False)
     If lSave Then
      objData.Save GetDSN
     Else
      objData.Cancel GetDSN
     End If
     InitValue
     cKodeAnggota.SetFocus
     cFaktur.Text = GetNomor("mutasitabungan", "nomormutasitabungan", GetID, SisModulTransaksi.SimpananHarian)
  End If
End Sub

Private Function ValidSaving() As Boolean
ValidSaving = True
  'cek faktur
  If Trim(cFaktur.Text) = "" Then
    MsgBox "Faktur transaksi tidak valid/kosong", vbExclamation, Me.Caption
    ValidSaving = False
    Exit Function
  End If
  'cek no rek
  If Trim(cKodeAnggota.Text) = "" Then
    MsgBox "No Rekening tidak valid", vbExclamation, Me.Caption
    ValidSaving = False
    Exit Function
  End If
  'cek kode transaksi
  If Trim(cKodeTransaksi.Text) = "" Then
    MsgBox "Kode Transaksi tidak valid", vbExclamation, Me.Caption
    ValidSaving = False
    Exit Function
  End If
  'cek DK
  If Trim(cDK.Text) = "" Or (cDK.Text <> "D" And cDK.Text <> "K") Then
    MsgBox "Kode D/K Transaksi Tidak valid", vbExclamation, Me.Caption
    ValidSaving = False
    Exit Function
  End If
  'cek rekening jurnal
  If Trim(cRekeningJurnal.Text) = "" Then
    MsgBox "Kode Rekening Jurnal Tidak Valid", vbExclamation, Me.Caption
    ValidSaving = False
    Exit Function
  End If
  'cek jumlah mutasi
  If nMutasi.Value <= 0 Then
    MsgBox "Jumlah Mutasi Tidak Valid, transaksi tidak akan disimpan", vbExclamation, Me.Caption
    ValidSaving = False
    Exit Function
  End If

  'cek saldo akhir
  If nAkhir.Value < 0 Then
    MsgBox "Saldo akhir simpanan tidak valid, transaksi tidak akan disimpan", vbExclamation, Me.Caption
    ValidSaving = False
    Exit Function
  End If
End Function


Private Sub Form_Load()
  SetIcon Me.hWnd
  CenterForm Me, True
  dTgl.Value = Date
  InitValue
  InitTabIndex
  cFaktur.Text = GetNomor("mutasitabungan", "nomormutasitabungan", GetID, SisModulTransaksi.SimpananHarian)
End Sub

Private Sub nMutasi_Change()
  nAkhir.Value = nAwal.Value + IIf(cDK.Text = "K", nMutasi.Value, -nMutasi.Value)
End Sub


