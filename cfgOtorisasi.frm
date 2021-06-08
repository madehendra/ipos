VERSION 5.00
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Begin VB.Form cfgOtorisasi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Otorisasi"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   10920
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   5190
      Left            =   0
      Top             =   45
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   9155
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      Begin BiSAFramProject.BiSAFrame BiSAFrame5 
         Height          =   585
         Left            =   3465
         Top             =   2520
         Width           =   2445
         _ExtentX        =   4313
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
         Begin VB.OptionButton optKunciSetoranKas 
            Caption         =   "&Tidak"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   930
            TabIndex        =   8
            Top             =   150
            Width           =   810
         End
         Begin VB.OptionButton optKunciSetoranKas 
            Caption         =   "&Ya"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   180
            TabIndex        =   7
            Top             =   150
            Width           =   540
         End
      End
      Begin BiSAFramProject.BiSAFrame BiSAFrame3 
         Height          =   675
         Left            =   255
         Top             =   705
         Width           =   10320
         _ExtentX        =   18203
         _ExtentY        =   1191
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
         Begin VB.OptionButton optOtorisasi 
            Caption         =   "&1 Ya terapkan saja"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   150
            TabIndex        =   3
            Top             =   210
            Width           =   1800
         End
         Begin VB.OptionButton optOtorisasi 
            Caption         =   "&2 Tidak, tidak usah diterapkan. Setiap user boleh menambah, mengkoreksi atau menghapus data."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Index           =   1
            Left            =   2055
            TabIndex        =   2
            Top             =   90
            Width           =   7800
         End
      End
      Begin BiSAFramProject.BiSAFrame BiSAFrame2 
         Height          =   570
         Left            =   270
         Top             =   2535
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   1005
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
         Begin VB.OptionButton optKunciAkunKas 
            Caption         =   "&Tidak"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   1
            Left            =   1005
            TabIndex        =   5
            Top             =   150
            Width           =   765
         End
         Begin VB.OptionButton optKunciAkunKas 
            Caption         =   "&Ya"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   195
            TabIndex        =   4
            Top             =   150
            Width           =   600
         End
      End
      Begin BiSAFramProject.BiSAFrame BiSAFrame6 
         Height          =   600
         Left            =   270
         Top             =   3435
         Width           =   3930
         _ExtentX        =   6932
         _ExtentY        =   1058
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
         Begin VB.OptionButton optProsesAudit 
            Caption         =   "&Tidak"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   1035
            TabIndex        =   11
            Top             =   195
            Width           =   780
         End
         Begin VB.OptionButton optProsesAudit 
            Caption         =   "&Ya"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   180
            TabIndex        =   10
            Top             =   195
            Width           =   570
         End
         Begin BiSANumberBoxProject.BiSANumberBox nHariBlokir 
            Height          =   360
            Left            =   2085
            TabIndex        =   12
            Top             =   135
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   635
            Appearance      =   0
            Decimals        =   0
            BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   " "
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
      Begin BiSAFramProject.BiSAFrame BiSAFrame4 
         Height          =   615
         Left            =   270
         Top             =   4410
         Width           =   3990
         _ExtentX        =   7038
         _ExtentY        =   1085
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
         Begin VB.OptionButton optEditPembelian 
            Caption         =   "&Tidak"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   180
            TabIndex        =   14
            Top             =   195
            Width           =   780
         End
         Begin VB.OptionButton optEditPembelian 
            Caption         =   "&Ya Bisa"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   1035
            TabIndex        =   13
            Top             =   195
            Width           =   1050
         End
      End
      Begin BiSAFramProject.BiSAFrame BiSAFrame7 
         Height          =   570
         Left            =   270
         Top             =   1665
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   1005
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
         Begin VB.OptionButton optOtorisasiPenuh 
            Caption         =   "&Ya"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   180
            TabIndex        =   18
            Top             =   150
            Width           =   540
         End
         Begin VB.OptionButton optOtorisasiPenuh 
            Caption         =   "&Tidak"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   915
            TabIndex        =   17
            Top             =   150
            Width           =   810
         End
      End
      Begin BiSAFramProject.BiSAFrame BiSAFrame8 
         Height          =   570
         Left            =   4890
         Top             =   3450
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   1005
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
         Begin VB.OptionButton optOtorisasiKasir 
            Caption         =   "&Tidak"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   915
            TabIndex        =   20
            Top             =   150
            Width           =   810
         End
         Begin VB.OptionButton optOtorisasiKasir 
            Caption         =   "&Ya"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   180
            TabIndex        =   19
            Top             =   150
            Width           =   540
         End
      End
      Begin VB.Label Label7 
         Caption         =   "Ijinkan Kasir Menghapus Item di Keranjang Penjualan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4920
         TabIndex        =   21
         Top             =   3210
         Width           =   6450
      End
      Begin VB.Label Label6 
         Caption         =   "Otorisasi Penuh?? Artinya tetap bisa melakukan koreksi tapi harus mendapatkan otorisasi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   285
         TabIndex        =   16
         Top             =   1455
         Width           =   6450
      End
      Begin VB.Label Label3 
         Caption         =   "Pembelian Bisa Di Edit?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   315
         TabIndex        =   15
         Top             =   4200
         Width           =   2715
      End
      Begin VB.Label Label5 
         Caption         =   "Lakukan Proses Audit ? (max 5 Hari)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   345
         TabIndex        =   9
         Top             =   3210
         Width           =   2715
      End
      Begin VB.Label Label4 
         Caption         =   "Kunci Setoran Kas Bank?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3495
         TabIndex        =   6
         Top             =   2265
         Width           =   1965
      End
      Begin VB.Label Label2 
         Caption         =   "Kunci Akun Kas Pada Setiap Transaksi?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   315
         TabIndex        =   1
         Top             =   2280
         Width           =   3000
      End
      Begin VB.Label Label1 
         Caption         =   "Apakah otorisasi/wewenang untuk melakukan proses menambah, mengkoreksi atau menghapus data  akan diterapkan dalam seluruh user?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   300
         TabIndex        =   0
         Top             =   225
         Width           =   6705
      End
   End
End
Attribute VB_Name = "cfgOtorisasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objData As New CodeSuiteLibrary.Data

Private Sub Form_Load()
Dim n As Single
  
  CenterForm Me
  SetIcon Me.hWnd
  
  SetOpt optOtorisasiPenuh, aCfg(objData, msOtorisasiPenuh)
  SetOpt optOtorisasi, aCfg(objData, msOtorisasi)
  SetOpt optKunciAkunKas, aCfg(objData, msOptKunciAkunKas)
  
'  SetOpt optAktifkanPeriodeAkuntansi, aCfg(objData, msOptKunciPeriodeAkuntansi)
  
  SetOpt optKunciSetoranKas, aCfg(objData, msKunciRekeningSetoranKas)
  SetOpt optProsesAudit, aCfg(objData, msOptAudit)
  SetOpt optEditPembelian, aCfg(objData, msBisaEditPembelian)
  SetOpt optOtorisasiKasir, aCfg(objData, msKunciKasirDelete)
  
  nHariBlokir.value = aCfg(objData, msJumlahHariBlokir)

  TabIndex optOtorisasi(0), n
  TabIndex optOtorisasi(1), n
  TabIndex optOtorisasiPenuh(0), n
  TabIndex optOtorisasiPenuh(1), n
  TabIndex optKunciAkunKas(0), n
  TabIndex optKunciAkunKas(1), n
  TabIndex optKunciSetoranKas(0), n
  TabIndex optKunciSetoranKas(1), n
  TabIndex optProsesAudit(0), n
  TabIndex optProsesAudit(1), n
  TabIndex nHariBlokir, n
  TabIndex optEditPembelian(0), n
  TabIndex optEditPembelian(1), n

End Sub

Private Sub Form_Unload(Cancel As Integer)

  UpdCfg msOtorisasiPenuh, GetOpt(optOtorisasiPenuh), objData, Label6.Caption, Me.Caption
  
  UpdCfg msOtorisasi, GetOpt(optOtorisasi), objData, Label1.Caption, Me.Caption
  UpdCfg msOptKunciAkunKas, GetOpt(optKunciAkunKas), objData, Label2.Caption, Me.Caption
  
'  UpdCfg msOptKunciPeriodeAkuntansi, GetOpt(optAktifkanPeriodeAkuntansi), objData, Label3.Caption, Me.Caption
  
  UpdCfg msKunciRekeningSetoranKas, GetOpt(optKunciSetoranKas), objData, Label4.Caption, Me.Caption
  UpdCfg msOptAudit, GetOpt(optProsesAudit), objData, Label5.Caption, Me.Caption
  UpdCfg msBisaEditPembelian, GetOpt(optEditPembelian), objData, Label3.Caption, Me.Caption

  
  UpdCfg msJumlahHariBlokir, nHariBlokir.value, objData, "Jumlah Hari Transaksi di Blokir", Me.Caption
  
  UpdCfg msKunciKasirDelete, GetOpt(optOtorisasiKasir), objData, Label1.Caption, Me.Caption
  
End Sub

