VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Begin VB.Form trOpsiPenjualanPembelian 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setting"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   17010
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7830
      Left            =   105
      TabIndex        =   0
      Top             =   135
      Width           =   16785
      _cx             =   29607
      _cy             =   13811
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   2
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   -2147483633
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   "PENJUALAN|PEMBELIAN|CETAKAN"
      Align           =   0
      CurrTab         =   0
      FirstTab        =   0
      Style           =   3
      Position        =   0
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   0
      BorderWidth     =   0
      BoldCurrent     =   -1  'True
      DogEars         =   -1  'True
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      TabCaptionPos   =   4
      TabPicturePos   =   0
      CaptionEmpty    =   ""
      Separators      =   0   'False
      Begin BiSAFramProject.BiSAFrame BiSAFrame4 
         Height          =   7455
         Left            =   45
         Top             =   330
         Width           =   16695
         _ExtentX        =   29448
         _ExtentY        =   13150
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
         Begin VB.CheckBox chkDiscountPenjualan 
            Caption         =   "Check1"
            Height          =   195
            Left            =   2430
            TabIndex        =   8
            Top             =   3150
            Width           =   255
         End
         Begin BiSAFramProject.BiSAFrame BiSAFrame16 
            Height          =   1296
            Left            =   5568
            Top             =   4764
            Width           =   2952
            _ExtentX        =   5212
            _ExtentY        =   2275
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
            Begin VB.OptionButton optPoin 
               Caption         =   "&1 Ya"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   0
               Left            =   210
               TabIndex        =   3
               Top             =   165
               Width           =   945
            End
            Begin VB.OptionButton optPoin 
               Caption         =   "&2 Tidak"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   1
               Left            =   1365
               TabIndex        =   2
               Top             =   165
               Width           =   945
            End
            Begin BiSANumberBoxProject.BiSANumberBox nKelipatan 
               Height          =   405
               Left            =   120
               TabIndex        =   1
               Top             =   795
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   714
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
            Begin BiSANumberBoxProject.BiSANumberBox nDay 
               Height          =   405
               Left            =   1695
               TabIndex        =   4
               Top             =   795
               Width           =   795
               _ExtentX        =   1402
               _ExtentY        =   714
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
            Begin VB.Label Label26 
               Caption         =   "Kelipatan"
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
               Left            =   150
               TabIndex        =   6
               Top             =   585
               Width           =   735
            End
            Begin VB.Label Label27 
               Caption         =   "Term(day)"
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
               Left            =   1740
               TabIndex        =   5
               Top             =   570
               Width           =   765
            End
         End
         Begin BiSANumberBoxProject.BiSANumberBox nMinimumDeposit 
            Height          =   330
            Left            =   2430
            TabIndex        =   7
            Top             =   4605
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   582
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
         Begin BiSAFramProject.BiSAFrame BiSAFrame10 
            Height          =   480
            Left            =   2415
            Top             =   2550
            Width           =   6105
            _ExtentX        =   10769
            _ExtentY        =   847
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
            Begin VB.OptionButton optIjin 
               Caption         =   "&1 BELI"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   0
               Left            =   210
               TabIndex        =   11
               Top             =   150
               Width           =   990
            End
            Begin VB.OptionButton optIjin 
               Caption         =   "&2 POKOK"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   1
               Left            =   1380
               TabIndex        =   10
               Top             =   150
               Width           =   960
            End
            Begin VB.OptionButton optIjin 
               Caption         =   "&3 ABAIKAN SAJA!!"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   2
               Left            =   2796
               TabIndex        =   9
               Top             =   150
               Width           =   1770
            End
         End
         Begin BiSAFramProject.BiSAFrame BiSAFrame7 
            Height          =   465
            Left            =   2430
            Top             =   1755
            Width           =   6090
            _ExtentX        =   10742
            _ExtentY        =   820
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
            Begin VB.OptionButton optPerhitunganKomisi 
               Caption         =   "&1 Manual"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   0
               Left            =   195
               TabIndex        =   13
               Top             =   105
               Width           =   1050
            End
            Begin VB.OptionButton optPerhitunganKomisi 
               Caption         =   "&2 Otomatis"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   1
               Left            =   1335
               TabIndex        =   12
               Top             =   105
               Width           =   1170
            End
            Begin BiSANumberBoxProject.BiSANumberBox nKomisi 
               Height          =   315
               Left            =   2565
               TabIndex        =   14
               Top             =   75
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   556
               Appearance      =   0
               Decimals        =   0
               MaxValue        =   100
               MinValue        =   1
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
            Begin VB.Label Label12 
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   3315
               TabIndex        =   15
               Top             =   135
               Width           =   285
            End
         End
         Begin BiSAFramProject.BiSAFrame BiSAFrame6 
            Height          =   750
            Left            =   2430
            Top             =   975
            Width           =   6090
            _ExtentX        =   10742
            _ExtentY        =   1323
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
            Begin VB.OptionButton optHargaPenjualan 
               Caption         =   "&1 Harga pada data master"
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
               Left            =   195
               TabIndex        =   18
               Top             =   120
               Width           =   2310
            End
            Begin VB.OptionButton optHargaPenjualan 
               Caption         =   "&2 Harga terakhir yg diperoleh tiap tiap customer"
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
               Index           =   1
               Left            =   195
               TabIndex        =   17
               Top             =   420
               Width           =   3792
            End
            Begin VB.OptionButton optHargaPenjualan 
               Caption         =   "&3 Harga Kontrak"
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
               Index           =   2
               Left            =   4140
               TabIndex        =   16
               Top             =   135
               Width           =   1710
            End
         End
         Begin BiSAFramProject.BiSAFrame BiSAFrame1 
            Height          =   450
            Left            =   2430
            Top             =   300
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   794
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
            Begin VB.OptionButton optKolomHargaPenjualan 
               Caption         =   "&1 Dapat diedit"
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
               Index           =   0
               Left            =   180
               TabIndex        =   20
               Top             =   90
               Width           =   1425
            End
            Begin VB.OptionButton optKolomHargaPenjualan 
               Caption         =   "&2 Tidak dapat diedit"
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
               Index           =   1
               Left            =   1680
               TabIndex        =   19
               Top             =   90
               Width           =   1755
            End
         End
         Begin BiSANumberBoxProject.BiSANumberBox nDiscountPenjualan 
            Height          =   330
            Left            =   2715
            TabIndex        =   21
            Top             =   3075
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   582
            Appearance      =   0
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
         Begin BiSAFramProject.BiSAFrame BiSAFrame12 
            Height          =   480
            Left            =   2412
            Top             =   3516
            Width           =   2916
            _ExtentX        =   5133
            _ExtentY        =   847
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
            Begin VB.OptionButton optSaldoMinus 
               Caption         =   "&2 Tidak"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   1
               Left            =   1425
               TabIndex        =   23
               Top             =   165
               Width           =   990
            End
            Begin VB.OptionButton optSaldoMinus 
               Caption         =   "&1 Ya"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   0
               Left            =   210
               TabIndex        =   22
               Top             =   150
               Width           =   750
            End
         End
         Begin BiSAFramProject.BiSAFrame BiSAFrame14 
            Height          =   540
            Left            =   2415
            Top             =   3990
            Width           =   2925
            _ExtentX        =   5159
            _ExtentY        =   953
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
            Begin BiSANumberBoxProject.BiSANumberBox nQtyDecimals 
               Height          =   330
               Left            =   330
               TabIndex        =   24
               Top             =   90
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   582
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
         Begin BiSAFramProject.BiSAFrame BiSAFrame8 
            Height          =   540
            Left            =   2415
            Top             =   4995
            Width           =   2925
            _ExtentX        =   5159
            _ExtentY        =   953
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
            Begin VB.OptionButton optDefaultPenjualan 
               Caption         =   "&Tunai"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   0
               Left            =   210
               TabIndex        =   26
               Top             =   195
               Width           =   972
            End
            Begin VB.OptionButton optDefaultPenjualan 
               Caption         =   "&Bon"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   1
               Left            =   1212
               TabIndex        =   25
               Top             =   204
               Width           =   870
            End
         End
         Begin BiSAFramProject.BiSAFrame BiSAFrame17 
            Height          =   480
            Left            =   9705
            Top             =   405
            Width           =   2910
            _ExtentX        =   5133
            _ExtentY        =   847
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
            Begin VB.OptionButton optHapusTransaksiPenjualan 
               Caption         =   "&1 Ya"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   0
               Left            =   210
               TabIndex        =   28
               Top             =   150
               Width           =   750
            End
            Begin VB.OptionButton optHapusTransaksiPenjualan 
               Caption         =   "&2 Tidak"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   1
               Left            =   1035
               TabIndex        =   27
               Top             =   150
               Width           =   990
            End
         End
         Begin BiSAFramProject.BiSAFrame BiSAFrame20 
            Height          =   480
            Left            =   9705
            Top             =   1305
            Width           =   2910
            _ExtentX        =   5133
            _ExtentY        =   847
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
            Begin VB.OptionButton optNotifikasiStock 
               Caption         =   "&2 Tidak"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   1
               Left            =   1035
               TabIndex        =   30
               Top             =   150
               Width           =   990
            End
            Begin VB.OptionButton optNotifikasiStock 
               Caption         =   "&1 Ya"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   0
               Left            =   210
               TabIndex        =   29
               Top             =   150
               Width           =   750
            End
         End
         Begin BiSAFramProject.BiSAFrame BiSAFrame15 
            Height          =   480
            Left            =   13560
            Top             =   405
            Width           =   2910
            _ExtentX        =   5133
            _ExtentY        =   847
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
            Begin VB.OptionButton optEditTransaksiPenjualan 
               Caption         =   "&2 Tidak"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   1
               Left            =   1035
               TabIndex        =   70
               Top             =   150
               Width           =   990
            End
            Begin VB.OptionButton optEditTransaksiPenjualan 
               Caption         =   "&1 Ya"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   0
               Left            =   210
               TabIndex        =   69
               Top             =   150
               Width           =   750
            End
         End
         Begin BiSAFramProject.BiSAFrame BiSAFrame23 
            Height          =   480
            Left            =   13560
            Top             =   1290
            Width           =   2910
            _ExtentX        =   5133
            _ExtentY        =   847
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
            Begin VB.OptionButton optGroupSales 
               Caption         =   "&1 Ya"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   0
               Left            =   210
               TabIndex        =   80
               Top             =   150
               Width           =   750
            End
            Begin VB.OptionButton optGroupSales 
               Caption         =   "&2 Tidak"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   1
               Left            =   1035
               TabIndex        =   79
               Top             =   150
               Width           =   990
            End
         End
         Begin BiSATextBoxProject.BiSABrowse cKodeGroupSalesDefault 
            Height          =   330
            Left            =   9450
            TabIndex        =   83
            Top             =   2625
            Width           =   4500
            _ExtentX        =   7938
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
            Appearance      =   0
            Button          =   -1  'True
            Caption         =   "*Group Sales Default"
            CaptionWidth    =   2500
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
         Begin BiSATextBoxProject.BiSATextBox cNamaGroupSalesDefault 
            Height          =   330
            Left            =   13980
            TabIndex        =   84
            Top             =   2625
            Width           =   2490
            _ExtentX        =   4392
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
            BackColor       =   -2147483633
            Enabled         =   0   'False
            Appearance      =   0
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
         Begin BiSAFramProject.BiSAFrame BiSAFrame24 
            Height          =   480
            Left            =   9705
            Top             =   2070
            Width           =   2910
            _ExtentX        =   5133
            _ExtentY        =   847
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
            Begin VB.OptionButton optModelPelunasanPiutang 
               Caption         =   "&1 Per Faktur"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   0
               Left            =   195
               TabIndex        =   86
               Top             =   150
               Width           =   1230
            End
            Begin VB.OptionButton optModelPelunasanPiutang 
               Caption         =   "&2 Bebas"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   1
               Left            =   1530
               TabIndex        =   85
               Top             =   150
               Width           =   1035
            End
         End
         Begin BiSAFramProject.BiSAFrame BiSAFrame25 
            Height          =   540
            Left            =   2415
            Top             =   5550
            Width           =   2925
            _ExtentX        =   5159
            _ExtentY        =   953
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
            Begin BiSANumberBoxProject.BiSANumberBox nBulanBlokir 
               Height          =   330
               Left            =   135
               TabIndex        =   89
               Top             =   105
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   582
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
            Begin VB.Label Label32 
               Caption         =   ".. x Bulan dari Nota dibuat"
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
               Left            =   825
               TabIndex        =   91
               Top             =   150
               Width           =   1935
            End
         End
         Begin BiSANumberBoxProject.BiSANumberBox nMinKartu 
            Height          =   330
            Left            =   2430
            TabIndex        =   92
            TabStop         =   0   'False
            Top             =   6120
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   582
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
         Begin VB.Label Label33 
            Caption         =   "Min Bayar u/ Pakai Kartu"
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
            Left            =   180
            TabIndex        =   93
            Top             =   6165
            Width           =   1935
         End
         Begin VB.Label Label31 
            Caption         =   "Block Member yg blm lunas"
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
            Left            =   195
            TabIndex        =   90
            Top             =   5715
            Width           =   1935
         End
         Begin VB.Label Label29 
            Caption         =   "Model Pelunasan Piutang"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   9435
            TabIndex        =   87
            Top             =   1845
            Width           =   3090
         End
         Begin VB.Label Label28 
            Caption         =   "*Group Sales?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   13080
            TabIndex        =   78
            Top             =   1065
            Width           =   1335
         End
         Begin VB.Label Label16 
            Caption         =   "Transaksi Tunai Bisa di Edit?"
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
            Left            =   13020
            TabIndex        =   71
            Top             =   180
            Width           =   2715
         End
         Begin VB.Label Label4 
            Caption         =   "Kolom Harga"
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
            Left            =   195
            TabIndex        =   43
            Top             =   390
            Width           =   1065
         End
         Begin VB.Label Label1 
            Caption         =   "Harga yg tercantum dalam penjualan non tunai adalah:"
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
            Left            =   195
            TabIndex        =   42
            Top             =   765
            Width           =   4170
         End
         Begin VB.Label Label11 
            Caption         =   "Perhitungan Komisi"
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
            Left            =   195
            TabIndex        =   41
            Top             =   1860
            Width           =   1530
         End
         Begin VB.Label Label14 
            Caption         =   "JANGAN Ijinkan HARGA JUAL di bawah HARGA : ?!!!"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   195
            TabIndex        =   40
            Top             =   2280
            Width           =   3870
         End
         Begin VB.Label Label17 
            Caption         =   "Default Diskon Penj % / Item"
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
            Left            =   195
            TabIndex        =   39
            Top             =   3135
            Width           =   1815
         End
         Begin VB.Label Label2 
            Caption         =   "Apabila dicentang, maka diskon yg ada pada master akan diabaikan dan diganti dengan diskon ini."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   432
            Left            =   3840
            TabIndex        =   38
            Top             =   3060
            Width           =   5184
         End
         Begin VB.Label Label19 
            Caption         =   "Saldo Minus Diijinkan??"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   276
            Left            =   192
            TabIndex        =   37
            Top             =   3636
            Width           =   1932
         End
         Begin VB.Label Label23 
            Caption         =   "Qty Decimals ???"
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
            Left            =   195
            TabIndex        =   36
            Top             =   4125
            Width           =   1935
         End
         Begin VB.Label Label22 
            Caption         =   "Minimum Deposit %"
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
            Left            =   195
            TabIndex        =   35
            Top             =   4665
            Width           =   1935
         End
         Begin VB.Label Label25 
            Caption         =   "Sistem Poin ??"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   276
            Left            =   5556
            TabIndex        =   34
            Top             =   4524
            Width           =   1128
         End
         Begin VB.Label Label3 
            Caption         =   "Default Penjualan"
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
            Left            =   195
            TabIndex        =   33
            Top             =   5175
            Width           =   1935
         End
         Begin VB.Label Label8 
            Caption         =   "Transaksi Penjualan Bisa dihapus?"
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
            Left            =   9420
            TabIndex        =   32
            Top             =   195
            Width           =   2715
         End
         Begin VB.Label Label15 
            Caption         =   "Tampilkan Notifikasi Jumlah Stok?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   9435
            TabIndex        =   31
            Top             =   1080
            Width           =   3090
         End
      End
      Begin BiSAFramProject.BiSAFrame BiSAFrame5 
         Height          =   7455
         Left            =   17430
         Top             =   330
         Width           =   16695
         _ExtentX        =   29448
         _ExtentY        =   13150
         Caption         =   "Default"
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
         Begin BiSAFramProject.BiSAFrame BiSAFrame18 
            Height          =   525
            Left            =   2085
            Top             =   630
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   926
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
            Begin VB.OptionButton optItemDiskonPembelian 
               Caption         =   "&Enable"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   0
               Left            =   165
               TabIndex        =   45
               Top             =   165
               Width           =   780
            End
            Begin VB.OptionButton optItemDiskonPembelian 
               Caption         =   "&Disable"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   1
               Left            =   1245
               TabIndex        =   44
               Top             =   165
               Width           =   870
            End
         End
         Begin BiSANumberBoxProject.BiSANumberBox nDiscountItemPembelian 
            Height          =   330
            Left            =   2115
            TabIndex        =   46
            Top             =   255
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   582
            Appearance      =   0
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
         Begin BiSAFramProject.BiSAFrame BiSAFrame19 
            Height          =   525
            Left            =   2070
            Top             =   1170
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   926
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
            Begin VB.OptionButton optDefaultPembelian 
               Caption         =   "&Bon"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   1
               Left            =   1245
               TabIndex        =   48
               Top             =   165
               Width           =   870
            End
            Begin VB.OptionButton optDefaultPembelian 
               Caption         =   "&Tunai"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   0
               Left            =   165
               TabIndex        =   47
               Top             =   165
               Width           =   780
            End
         End
         Begin BiSAFramProject.BiSAFrame BiSAFrame13 
            Height          =   855
            Left            =   90
            Top             =   2565
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   1508
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
            Begin BiSANumberBoxProject.BiSANumberBox nDiskonEstimasi 
               Height          =   330
               Left            =   2130
               TabIndex        =   52
               Top             =   390
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   582
               Appearance      =   0
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
            Begin VB.Label Label20 
               Caption         =   "Menentukan diskon u/ harga netto dari estimasi"
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
               Left            =   195
               TabIndex        =   53
               Top             =   105
               Width           =   3570
            End
         End
         Begin BiSAFramProject.BiSAFrame BiSAFrame9 
            Height          =   630
            Left            =   90
            Top             =   3420
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   1111
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
            Begin VB.OptionButton optCetakPembelian 
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
               Height          =   210
               Index           =   1
               Left            =   3075
               TabIndex        =   55
               Top             =   180
               Width           =   870
            End
            Begin VB.OptionButton optCetakPembelian 
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
               Height          =   210
               Index           =   0
               Left            =   2130
               TabIndex        =   54
               Top             =   180
               Width           =   780
            End
            Begin VB.Label Label6 
               Caption         =   "Cetak Faktur?"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   195
               TabIndex        =   56
               Top             =   195
               Width           =   1155
            End
         End
         Begin BiSAFramProject.BiSAFrame BiSAFrame21 
            Height          =   525
            Left            =   2070
            Top             =   1935
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   926
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
            Begin VB.OptionButton optInvoiceAsliPembelian 
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
               Height          =   210
               Index           =   0
               Left            =   165
               TabIndex        =   74
               Top             =   165
               Width           =   780
            End
            Begin VB.OptionButton optInvoiceAsliPembelian 
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
               Height          =   210
               Index           =   1
               Left            =   1245
               TabIndex        =   73
               Top             =   165
               Width           =   870
            End
         End
         Begin BiSAFramProject.BiSAFrame BiSAFrame22 
            Height          =   630
            Left            =   90
            Top             =   4065
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   1111
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
            Begin VB.OptionButton optModelInputPembelian 
               Caption         =   "&Sederhana"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   0
               Left            =   2130
               TabIndex        =   76
               Top             =   180
               Width           =   1155
            End
            Begin VB.OptionButton optModelInputPembelian 
               Caption         =   "&Mahir"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   1
               Left            =   3360
               TabIndex        =   75
               Top             =   180
               Width           =   870
            End
            Begin VB.Label Label24 
               Caption         =   "Model Input"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   195
               TabIndex        =   77
               Top             =   195
               Width           =   1155
            End
         End
         Begin VB.Label Label21 
            Caption         =   "Wajib Menyertakan Nomor Invoice Asli?"
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
            Left            =   165
            TabIndex        =   72
            Top             =   1755
            Width           =   2985
         End
         Begin VB.Label Label9 
            Caption         =   "Default Diskon % / Item"
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
            Left            =   165
            TabIndex        =   51
            Top             =   300
            Width           =   1815
         End
         Begin VB.Label Label10 
            Caption         =   "Diskon Item Pembelian"
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
            Left            =   165
            TabIndex        =   50
            Top             =   690
            Width           =   1815
         End
         Begin VB.Label Label13 
            Caption         =   "Default Pembelian"
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
            Left            =   165
            TabIndex        =   49
            Top             =   1140
            Width           =   1545
         End
      End
      Begin BiSAFramProject.BiSAFrame BiSAFrame2 
         Height          =   7455
         Left            =   17730
         Top             =   330
         Width           =   16695
         _ExtentX        =   29448
         _ExtentY        =   13150
         Caption         =   "Cetakan"
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
         Begin BiSAFramProject.BiSAFrame BiSAFrame3 
            Height          =   465
            Left            =   2355
            Top             =   255
            Width           =   4440
            _ExtentX        =   7832
            _ExtentY        =   820
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
            Begin VB.OptionButton optCetakanPenjualan 
               Caption         =   "&1 Nota NCR"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   195
               TabIndex        =   59
               Top             =   150
               Width           =   1395
            End
            Begin VB.OptionButton optCetakanPenjualan 
               Caption         =   "&2 Nota Wartel"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   1665
               TabIndex        =   58
               Top             =   150
               Width           =   1335
            End
            Begin VB.OptionButton optCetakanPenjualan 
               Caption         =   "&3 Struk"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   2
               Left            =   3210
               TabIndex        =   57
               Top             =   150
               Width           =   1080
            End
         End
         Begin BiSAFramProject.BiSAFrame BiSAFrame11 
            Height          =   480
            Left            =   2385
            Top             =   2310
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   847
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
            Begin VB.OptionButton optUp 
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
               Left            =   825
               TabIndex        =   61
               Top             =   150
               Width           =   840
            End
            Begin VB.OptionButton optUp 
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
               Left            =   135
               TabIndex        =   60
               Top             =   150
               Width           =   585
            End
         End
         Begin BiSATextBoxProject.BiSATextBox cFooterPenjualanNonTunai 
            Height          =   330
            Left            =   2385
            TabIndex        =   62
            Top             =   1065
            Width           =   7935
            _ExtentX        =   13996
            _ExtentY        =   582
            Text            =   "1234567890123456789012345678901234567890"
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
            Alignment       =   2
            MaxLength       =   255
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
         Begin BiSATextBoxProject.BiSATextBox cFooterPenjualanNonTunai2 
            Height          =   330
            Left            =   2385
            TabIndex        =   63
            Top             =   1410
            Width           =   7920
            _ExtentX        =   13970
            _ExtentY        =   582
            Text            =   "1234567890123456789012345678901234567890"
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
            Alignment       =   2
            MaxLength       =   255
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
         Begin BiSATextBoxProject.BiSATextBox cFooterPenjualanNonTunai3 
            Height          =   330
            Left            =   2385
            TabIndex        =   88
            Top             =   1755
            Width           =   7935
            _ExtentX        =   13996
            _ExtentY        =   582
            Text            =   "1234567890123456789012345678901234567890"
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
            Alignment       =   2
            MaxLength       =   255
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
         Begin VB.Label Label30 
            Caption         =   "Footer Nota"
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
            Left            =   195
            TabIndex        =   67
            Top             =   1110
            Width           =   1185
         End
         Begin VB.Label Label18 
            Caption         =   "Print Up (Kepada) pd Nota"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   180
            TabIndex        =   66
            Top             =   2430
            Width           =   2115
         End
         Begin VB.Label Label7 
            Caption         =   "*) Khusus untuk pencetakan dengan nota"
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
            Left            =   2385
            TabIndex        =   65
            Top             =   810
            Width           =   3075
         End
         Begin VB.Label Label5 
            Caption         =   "Model Cetakan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   180
            TabIndex        =   64
            Top             =   405
            Width           =   1560
         End
      End
   End
   Begin SizerOneLibCtl.ElasticOne ElasticOne3 
      Height          =   690
      Left            =   75
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   7980
      Width           =   16815
      _cx             =   29660
      _cy             =   1217
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   4
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   15645
         TabIndex        =   81
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
         Picture         =   "trOpsiPenjualanPembelian.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   14535
         TabIndex        =   82
         Top             =   120
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
         Picture         =   "trOpsiPenjualanPembelian.frx":00A6
      End
   End
End
Attribute VB_Name = "trOpsiPenjualanPembelian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data


Private Sub cKodeGroupSalesDefault_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "groupsales", , , , , "status = 1")
  If Not dbData.EOF Then
    cKodeGroupSalesDefault.Text = cKodeGroupSalesDefault.Browse(dbData)
    cNamaGroupSalesDefault.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdSimpan_Click()
  GetNotifikasiAdd "Menyimpan konfigurasi"
  SaveTab_PenjualanNonTunai
  SaveTab_Pembelian
  GetNotifikasiRemove
End Sub

Private Sub SaveTab_PenjualanNonTunai()
  UpdCfg msOptUp, GetOpt(optUp), objData, Label18.Caption, Me.Caption
  UpdCfg msKolomHargaPenjualanNonTunai, GetOpt(optKolomHargaPenjualan), objData, Label4.Caption, Me.Caption
  'Simpan model cetakan dalam registry
  SaveRegistry reg_CetakanPenjualanNonTunai, GetOpt(optCetakanPenjualan)
  SaveRegistry reg_TampilNotifikasi, GetOpt(optNotifikasiStock)
  SaveRegistry reg_OptGroupSales, GetOpt(optGroupSales)
  SaveRegistry reg_OptModelPelunasanPiutang, GetOpt(optGroupSales)
  SaveRegistry reg_OptModelPelunasanPiutang, GetOpt(optModelPelunasanPiutang)
  
  
  'UpdCfg msCetakanPenjualanNonTunai, GetOpt(optCetakanPenjualan), objData, Label5.Caption, Me.Caption
 ' UpdCfg msJumlahCetakanPenjualanNonTunai, nJumlahCetakan.Value, objData, Label6.Caption, Me.Caption
  UpdCfg msHargaPenjualanNonTunai, GetOpt(optHargaPenjualan), objData, Label1.Caption, Me.Caption
  UpdCfg msPerhitunganKomisi, GetOpt(optPerhitunganKomisi), objData, Label11.Caption, Me.Caption
  UpdCfg msPersenKomisi, nKomisi.Value, objData, "Persentase Komisi", Me.Caption
  UpdCfg msIjinkanHargaBeliDibawahHargajual, GetOpt(optIjin), objData, "Ijinkan Harga Beli Di Bawah Harga Jual", Me.Caption
  
  UpdCfg msPersenKomisi, nKomisi.Value, objData, "Persentase Komisi", Me.Caption
  UpdCfg msFooterPenjualanNonTunai, cFooterPenjualanNonTunai.Text, objData, "Footer Invoince Penjualan", Me.Caption
  UpdCfg msFooterPenjualanNonTunai2, cFooterPenjualanNonTunai2.Text, objData, "Footer Invoince Penjualan2", Me.Caption
  UpdCfg msFooterPenjualanNonTunai3, cFooterPenjualanNonTunai3.Text, objData, "Footer Invoince Penjualan3", Me.Caption

  UpdCfg msDiscountPenjualan, nDiscountPenjualan.Value, objData, Label9.Caption, Me.Caption
  UpdCfg msCHKdiscountPenjualan, chkDiscountPenjualan.Value, objData, "Cek Discount Penjualan", Me.Caption
  UpdCfg msSaldoMinus, GetOpt(optSaldoMinus), objData, "Ijinkan Saldo Minus", Me.Caption
  UpdCfg msNilaiDecimals, nQtyDecimals.Value, objData, "Qty Decimals", Me.Caption
  UpdCfg msMinimumDeposit, nMinimumDeposit.Value, objData, "Minimum Deposit", Me.Caption
  UpdCfg msPoin, GetOpt(optPoin), objData, "Sistem Poin", Me.Caption
  UpdCfg msHapusTransaksiPenjualan, GetOpt(optHapusTransaksiPenjualan), objData, "Hapus Transaksi Penjualan", Me.Caption
  UpdCfg msEditTransaksiPenjualan, GetOpt(optEditTransaksiPenjualan), objData, "Edit Transaksi Penjualan", Me.Caption

  UpdCfg msKelipatan, nKelipatan.Value, objData, "Kelipatan", Me.Caption
  UpdCfg msTerm, nDay.Value, objData, "Day Term", Me.Caption
  UpdCfg msDefaultModelPenjualan, GetOpt(optDefaultPenjualan), objData, "Default Penjualan", Me.Caption
  
  UpdCfg msBulanBlokir, nBulanBlokir.Value, objData, "Bulan Blokir", Me.Caption
  UpdCfg msMinKartu, nMinKartu.Value, objData, "Minimal Pembayaran Pakai Kartu", Me.Caption
  
  
End Sub

Private Sub SaveTab_Pembelian()
  UpdCfg msDiscountItemPembelian, nDiscountItemPembelian.Value, objData, Label9.Caption, Me.Caption
  UpdCfg msDiskonEstimasi, nDiskonEstimasi.Value, objData, Label20.Caption, Me.Caption
  UpdCfg msCetakPembelian, GetOpt(optCetakPembelian), objData, "Cetak Pembelian", Me.Caption
  UpdCfg msOptFakturAsliPembelian, GetOpt(optInvoiceAsliPembelian), objData, "Invoice Asli Pembelian", Me.Caption
  UpdCfg msEnableDisableDiscountItemPembelian, GetOpt(optItemDiskonPembelian), objData, "Enable Diskon Pembelian", Me.Caption
  UpdCfg msDefaultPembelian, GetOpt(optDefaultPembelian), objData, "Default Transaksi Pembelian", Me.Caption
  UpdCfg msModelInputPembelian, GetOpt(optModelInputPembelian), objData, "Model Input Pembelian", Me.Caption

End Sub

Private Sub LoadTab_Pembelian()
  nDiscountItemPembelian.Value = aCfg(objData, msDiscountItemPembelian)
  nDiskonEstimasi.Value = aCfg(objData, msDiskonEstimasi)
  SetOpt optCetakPembelian, aCfg(objData, msCetakPembelian)
  SetOpt optItemDiskonPembelian, aCfg(objData, msEnableDisableDiscountItemPembelian)
  SetOpt optDefaultPembelian, aCfg(objData, msDefaultPembelian)
  SetOpt optInvoiceAsliPembelian, aCfg(objData, msOptFakturAsliPembelian)
  SetOpt optModelInputPembelian, aCfg(objData, msModelInputPembelian)
End Sub

Private Sub LoadTab_PenjualanNonTunai()
  
  SetOpt optKolomHargaPenjualan, aCfg(objData, msKolomHargaPenjualanNonTunai, "1")
  'SetOpt optCetakanPenjualan, aCfg(objData, msCetakanPenjualanNonTunai, "1")
  SetOpt optCetakanPenjualan, GetRegistry(reg_CetakanPenjualanNonTunai)
  SetOpt optNotifikasiStock, GetRegistry(reg_TampilNotifikasi)
  SetOpt optGroupSales, GetRegistry(reg_OptGroupSales)
  SetOpt optHargaPenjualan, aCfg(objData, msHargaPenjualanNonTunai, "1")
  SetOpt optModelPelunasanPiutang, GetRegistry(reg_OptModelPelunasanPiutang)

  SetOpt optPerhitunganKomisi, aCfg(objData, msPerhitunganKomisi)
  nKomisi.Value = aCfg(objData, msPersenKomisi)
  SetOpt optIjin, aCfg(objData, msIjinkanHargaBeliDibawahHargajual)
  cFooterPenjualanNonTunai.Text = aCfg(objData, msFooterPenjualanNonTunai)
  cFooterPenjualanNonTunai2.Text = aCfg(objData, msFooterPenjualanNonTunai2)
  cFooterPenjualanNonTunai3.Text = aCfg(objData, msFooterPenjualanNonTunai3)
  nDiscountPenjualan.Value = aCfg(objData, msDiscountPenjualan)
  chkDiscountPenjualan.Value = aCfg(objData, msCHKdiscountPenjualan)
  SetOpt optUp, aCfg(objData, msOptUp)
  SetOpt optDefaultPenjualan, aCfg(objData, msDefaultModelPenjualan)
  SetOpt optSaldoMinus, aCfg(objData, msSaldoMinus)
  nQtyDecimals.Value = aCfg(objData, msNilaiDecimals)
  nMinimumDeposit.Value = aCfg(objData, msMinimumDeposit)
  SetOpt optHapusTransaksiPenjualan, aCfg(objData, msHapusTransaksiPenjualan)
  SetOpt optEditTransaksiPenjualan, aCfg(objData, msEditTransaksiPenjualan)

  SetOpt optPoin, aCfg(objData, msPoin)
  nKelipatan.Value = aCfg(objData, msKelipatan)
  nDay.Value = aCfg(objData, msTerm)
  cKodeGroupSalesDefault.Text = GetRegistry(reg_KodeGroupPenjualan)
  nBulanBlokir.Value = aCfg(objData, msBulanBlokir)
  nMinKartu.Value = aCfg(objData, msMinKartu)
  
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  SetIcon Me.hWnd

  TabIndex optCetakanPenjualan(0), n
  TabIndex optCetakanPenjualan(1), n

  TabIndex optUp(0), n
  TabIndex optUp(1), n
  TabIndex optKolomHargaPenjualan(0), n
  TabIndex optKolomHargaPenjualan(1), n
  TabIndex optHargaPenjualan(0), n
  TabIndex optHargaPenjualan(1), n
  TabIndex optPerhitunganKomisi(0), n
  TabIndex optPerhitunganKomisi(1), n
  TabIndex nKomisi, n
  TabIndex optIjin(0), n
  TabIndex optIjin(1), n
  TabIndex optSaldoMinus(0), n
  TabIndex optSaldoMinus(1), n
  TabIndex nQtyDecimals, n
  TabIndex nMinimumDeposit, n
  TabIndex optPoin(0), n
  TabIndex optPoin(1), n
  TabIndex nKelipatan, n
  TabIndex nDay, n
  TabIndex optModelPelunasanPiutang(0), n
  TabIndex optModelPelunasanPiutang(1), n
  TabIndex cKodeGroupSalesDefault, n
  TabIndex cNamaGroupSalesDefault, n
  
  
  TabIndex nDiscountItemPembelian, n
  TabIndex nDiskonEstimasi, n
'  TabIndex optDiscountPembelian(0), n
'  TabIndex optDiscountPembelian(1), n
  
  
'  LoadTab_PenjualanKasir
  LoadTab_PenjualanNonTunai
  LoadTab_Pembelian
  
'  TabOne1.TabVisible(0) = False
  
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
End Sub

Private Sub optDiscountPembelian_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    keybd_event VK_TAB, 0, 0, 0
  End If
End Sub

Private Sub optIjin_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    keybd_event VK_TAB, 0, 0, 0
  End If
End Sub

Private Sub optKolomHargaKasir_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    keybd_event VK_TAB, 0, 0, 0
  End If
End Sub

Private Sub optKolomHargaPenjualan_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    keybd_event VK_TAB, 0, 0, 0
  End If
End Sub

Private Sub optModelInput_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    keybd_event VK_TAB, 0, 0, 0
  End If
End Sub
