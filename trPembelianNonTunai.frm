VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trPembelianNonTunaiOri 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PEMBELIAN"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   17190
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   17190
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3000
      Left            =   180
      TabIndex        =   39
      Top             =   15
      Width           =   9420
      Begin BiSAButtonProject.BiSAButton cmdGetOrder 
         Height          =   330
         Left            =   3300
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   450
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         Caption         =   "Get Order"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
      End
      Begin BiSATextBoxProject.BiSABrowse cNamaSupplier 
         Height          =   330
         Left            =   3540
         TabIndex        =   41
         Top             =   810
         Width           =   3075
         _ExtentX        =   5424
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
      Begin BiSATextBoxProject.BiSATextBox cKota 
         Height          =   330
         Left            =   4335
         TabIndex        =   42
         Top             =   1170
         Width           =   2280
         _ExtentX        =   4022
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
         CaptionWidth    =   0
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
         Left            =   270
         TabIndex        =   43
         Top             =   1170
         Width           =   4035
         _ExtentX        =   7117
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
      Begin BiSATextBoxProject.BiSABrowse cSupplier 
         Height          =   330
         Left            =   270
         TabIndex        =   44
         Top             =   810
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
         BackColor       =   -2147483633
         Enabled         =   0   'False
         Appearance      =   0
         Caption         =   "Supplier"
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
         Left            =   285
         TabIndex        =   45
         Top             =   450
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   582
         Value           =   "16-01-2016"
         Appearance      =   0
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
      Begin BiSATextBoxProject.BiSABrowse cFaktur 
         Height          =   330
         Left            =   285
         TabIndex        =   46
         Top             =   1545
         Width           =   3750
         _ExtentX        =   6615
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
         BackColor       =   -2147483633
         Appearance      =   0
         Caption         =   "Nomor"
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
      Begin BiSAButtonProject.BiSAButton cmdImportWizard 
         Height          =   360
         Left            =   6960
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   300
         Visible         =   0   'False
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   635
         Caption         =   "Import Wizard"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
      End
      Begin BiSANumberBoxProject.BiSANumberBox nPersDisc 
         Height          =   330
         Left            =   285
         TabIndex        =   48
         Top             =   1905
         Width           =   2070
         _ExtentX        =   3651
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
         Caption         =   "Discount"
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
      Begin BiSATextBoxProject.BiSATextBox cFakturAsli 
         Height          =   330
         Left            =   285
         TabIndex        =   49
         Top             =   2280
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   582
         Text            =   "12345678901234567890"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Verdana"
         BackColor       =   16777215
         MaxLength       =   20
         Appearance      =   0
         Caption         =   "Fak. Asli"
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
   Begin VB.Frame Frame4 
      Height          =   3000
      Left            =   9645
      TabIndex        =   32
      Top             =   15
      Width           =   7320
      Begin BiSATextBoxProject.BiSABrowse cAkunKas 
         Height          =   330
         Left            =   375
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1335
         Width           =   3255
         _ExtentX        =   5741
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
         Enabled         =   0   'False
         Appearance      =   0
         Caption         =   "Akun Kas"
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
      Begin BiSADateProject.BiSADate dJthTmp 
         Height          =   330
         Left            =   375
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   645
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   582
         Value           =   "16-01-2016"
         Appearance      =   0
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
         Caption         =   "Due Date"
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
      Begin BiSANumberBoxProject.BiSANumberBox nPPn 
         Height          =   330
         Left            =   375
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   990
         Width           =   2175
         _ExtentX        =   3836
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
         Caption         =   "PPn(%)"
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
      Begin BiSATextBoxProject.BiSABrowse cGudang 
         Height          =   330
         Left            =   390
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1695
         Width           =   2550
         _ExtentX        =   4498
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
         Caption         =   "Gudang"
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
      Begin BiSATextBoxProject.BiSATextBox cNamaGudang 
         Height          =   330
         Left            =   2970
         TabIndex        =   37
         Top             =   1695
         Width           =   2445
         _ExtentX        =   4313
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
         CaptionWidth    =   1400
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
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   4350
         Top             =   1035
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin BiSATextBoxProject.BiSABrowse cBuyer 
         Height          =   330
         Left            =   390
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   2055
         Width           =   2550
         _ExtentX        =   4498
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
         Caption         =   "Buyer"
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
   Begin BiSAFramProject.BiSAFrame BisaFrame2 
      Height          =   4725
      Left            =   105
      Top             =   2985
      Width           =   17010
      _ExtentX        =   30004
      _ExtentY        =   8334
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      BackColor       =   -2147483633
      Begin VB.Frame Frame2 
         Height          =   960
         Left            =   9405
         TabIndex        =   17
         Top             =   3645
         Width           =   7440
         Begin VB.CheckBox chkTunai 
            Caption         =   "Check1"
            Height          =   195
            Left            =   3750
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   555
            Width           =   195
         End
         Begin BiSANumberBoxProject.BiSANumberBox nTunai 
            Height          =   330
            Left            =   3990
            TabIndex        =   19
            Top             =   480
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   582
            Appearance      =   0
            Enabled         =   0   'False
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
         Begin BiSANumberBoxProject.BiSANumberBox nHutang 
            Height          =   330
            Left            =   5565
            TabIndex        =   20
            Top             =   480
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   582
            Appearance      =   0
            Enabled         =   0   'False
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
         Begin VB.Label Label5 
            Caption         =   "Hutang"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   5595
            TabIndex        =   22
            Top             =   240
            Width           =   780
         End
         Begin VB.Label Label4 
            Caption         =   "Tunai"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   3960
            TabIndex        =   21
            Top             =   240
            Width           =   570
         End
      End
      Begin VB.Frame Frame1 
         Height          =   960
         Left            =   0
         TabIndex        =   16
         Top             =   3645
         Width           =   9375
         Begin BiSAButtonProject.BiSAButton BiSAButton1 
            Height          =   330
            Left            =   6090
            TabIndex        =   23
            Top             =   420
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   582
            Caption         =   "Print Barcode"
            BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483633
         End
         Begin BiSANumberBoxProject.BiSANumberBox nTotal 
            Height          =   315
            Left            =   4560
            TabIndex        =   24
            Top             =   435
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            Appearance      =   0
            Enabled         =   0   'False
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
         Begin BiSANumberBoxProject.BiSANumberBox nPajak 
            Height          =   315
            Left            =   3210
            TabIndex        =   25
            Top             =   435
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            Appearance      =   0
            Enabled         =   0   'False
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
         Begin BiSANumberBoxProject.BiSANumberBox nDiscount 
            Height          =   315
            Left            =   1755
            TabIndex        =   26
            Top             =   435
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            Appearance      =   0
            Enabled         =   0   'False
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
         Begin BiSANumberBoxProject.BiSANumberBox nSubTotal 
            Height          =   315
            Left            =   210
            TabIndex        =   27
            Top             =   435
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            Appearance      =   0
            Enabled         =   0   'False
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
         Begin VB.Label Label6 
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   4560
            TabIndex        =   31
            Top             =   195
            Width           =   720
         End
         Begin VB.Label Label3 
            Caption         =   "Pajak"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   3210
            TabIndex        =   30
            Top             =   210
            Width           =   720
         End
         Begin VB.Label Label2 
            Caption         =   "Discount"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   1755
            TabIndex        =   29
            Top             =   210
            Width           =   720
         End
         Begin VB.Label Label1 
            Caption         =   "Subtotal"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   210
            TabIndex        =   28
            Top             =   210
            Width           =   720
         End
      End
      Begin BiSANumberBoxProject.BiSANumberBox nQty 
         Height          =   330
         Left            =   9300
         TabIndex        =   0
         Top             =   60
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   582
         Appearance      =   0
         Decimals        =   0
         MinValue        =   0
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
      Begin BiSANumberBoxProject.BiSANumberBox nHarga 
         Height          =   330
         Left            =   11730
         TabIndex        =   1
         Top             =   60
         Width           =   1620
         _ExtentX        =   2858
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
      Begin BiSATextBoxProject.BiSATextBox cSatuan 
         Height          =   330
         Left            =   10515
         TabIndex        =   2
         Top             =   60
         Width           =   1200
         _ExtentX        =   2117
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
      Begin BiSATextBoxProject.BiSABrowse cNama 
         Height          =   330
         Left            =   2835
         TabIndex        =   3
         Top             =   60
         Width           =   6435
         _ExtentX        =   11351
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
      Begin BiSATextBoxProject.BiSABrowse cBarcode 
         Height          =   330
         Left            =   645
         TabIndex        =   4
         Top             =   60
         Width           =   2175
         _ExtentX        =   3836
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
         GetPicture      =   1
         Button          =   -1  'True
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
      Begin BiSANumberBoxProject.BiSANumberBox nNomor 
         Height          =   330
         Left            =   90
         TabIndex        =   5
         Top             =   60
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   582
         Appearance      =   0
         Decimals        =   0
         DecimalPoint    =   ""
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   3075
         Left            =   90
         TabIndex        =   6
         Top             =   420
         Width           =   16740
         _ExtentX        =   29528
         _ExtentY        =   5424
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "NO."
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "BARCODE"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "NAMA BARANG"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "QTY"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "SATUAN"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "HARGA"
         Columns(5).DataField=   ""
         Columns(5).NumberFormat=   "###,###,###,###"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "DISC"
         Columns(6).DataField=   ""
         Columns(6).NumberFormat=   "###,###,###,###,##0.00"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "JUMLAH"
         Columns(7).DataField=   ""
         Columns(7).NumberFormat=   "###,###,###,###"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "KODESTOCK"
         Columns(8).DataField=   ""
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "ID"
         Columns(9).DataField=   ""
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   10
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=10"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=926"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=847"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=514"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=3863"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3784"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=512"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=11456"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=11377"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=512"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=2117"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2037"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=197122"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2170"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2090"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=513"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=2884"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2805"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=514"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=1296"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=1217"
         Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=514"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(36)=   "Column(7).Width=4154"
         Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=4075"
         Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=514"
         Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(41)=   "Column(8).Width=1508"
         Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=1429"
         Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=516"
         Splits(0)._ColumnProps(45)=   "Column(8).Visible=0"
         Splits(0)._ColumnProps(46)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(47)=   "Column(9).Width=2725"
         Splits(0)._ColumnProps(48)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(49)=   "Column(9)._WidthInPix=2646"
         Splits(0)._ColumnProps(50)=   "Column(9)._ColStyle=516"
         Splits(0)._ColumnProps(51)=   "Column(9).Visible=0"
         Splits(0)._ColumnProps(52)=   "Column(9).Order=10"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         BorderStyle     =   0
         ColumnFooters   =   -1  'True
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1,5
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   16777215
         RowDividerColor =   12632256
         RowSubDividerColor=   12632256
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=2,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000005&,.bold=0,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bgcolor=&H8000000F&"
         _StyleDefs(11)  =   ":id=2,.fgcolor=&H80000007&,.bold=0,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
         _StyleDefs(13)  =   ":id=2,.fontname=Tahoma"
         _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(15)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(16)  =   ":id=3,.fontname=Tahoma"
         _StyleDefs(17)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(18)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(28)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(29)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(31)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(32)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(33)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(34)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(35)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(36)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=1"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=0"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=66,.parent=13,.alignment=1"
         _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=14"
         _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=15,.alignment=1"
         _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=17"
         _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=2"
         _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
         _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
         _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
         _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=58,.parent=13,.alignment=1"
         _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
         _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
         _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
         _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=62,.parent=13,.alignment=1"
         _StyleDefs(66)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
         _StyleDefs(67)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
         _StyleDefs(68)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
         _StyleDefs(69)  =   "Splits(0).Columns(8).Style:id=74,.parent=13"
         _StyleDefs(70)  =   "Splits(0).Columns(8).HeadingStyle:id=71,.parent=14"
         _StyleDefs(71)  =   "Splits(0).Columns(8).FooterStyle:id=72,.parent=15"
         _StyleDefs(72)  =   "Splits(0).Columns(8).EditorStyle:id=73,.parent=17"
         _StyleDefs(73)  =   "Splits(0).Columns(9).Style:id=70,.parent=13"
         _StyleDefs(74)  =   "Splits(0).Columns(9).HeadingStyle:id=67,.parent=14"
         _StyleDefs(75)  =   "Splits(0).Columns(9).FooterStyle:id=68,.parent=15"
         _StyleDefs(76)  =   "Splits(0).Columns(9).EditorStyle:id=69,.parent=17"
         _StyleDefs(77)  =   "Named:id=33:Normal"
         _StyleDefs(78)  =   ":id=33,.parent=0"
         _StyleDefs(79)  =   "Named:id=34:Heading"
         _StyleDefs(80)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(81)  =   ":id=34,.wraptext=-1"
         _StyleDefs(82)  =   "Named:id=35:Footing"
         _StyleDefs(83)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(84)  =   "Named:id=36:Selected"
         _StyleDefs(85)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(86)  =   "Named:id=37:Caption"
         _StyleDefs(87)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(88)  =   "Named:id=38:HighlightRow"
         _StyleDefs(89)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(90)  =   "Named:id=39:EvenRow"
         _StyleDefs(91)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(92)  =   "Named:id=40:OddRow"
         _StyleDefs(93)  =   ":id=40,.parent=33"
         _StyleDefs(94)  =   "Named:id=41:RecordSelector"
         _StyleDefs(95)  =   ":id=41,.parent=34"
         _StyleDefs(96)  =   "Named:id=42:FilterBar"
         _StyleDefs(97)  =   ":id=42,.parent=33"
      End
      Begin BiSANumberBoxProject.BiSANumberBox nJumlah 
         Height          =   330
         Left            =   14115
         TabIndex        =   7
         Top             =   60
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   582
         Appearance      =   0
         Decimals        =   0
         Enabled         =   0   'False
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
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
      Begin BiSAButtonProject.BiSAButton cmdOK 
         Height          =   336
         Left            =   16464
         TabIndex        =   8
         Top             =   60
         Width           =   396
         _ExtentX        =   688
         _ExtentY        =   582
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
         Picture         =   "trPembelianNonTunai.frx":0000
      End
      Begin BiSANumberBoxProject.BiSANumberBox nDisc1 
         Height          =   330
         Left            =   13380
         TabIndex        =   9
         Top             =   60
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   582
         Appearance      =   0
         Enabled         =   0   'False
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
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
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   660
      Left            =   105
      Top             =   7695
      Width           =   17025
      _ExtentX        =   30030
      _ExtentY        =   1164
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
      Begin BiSAButtonProject.BiSAButton cmdHapus 
         Height          =   435
         Left            =   2235
         TabIndex        =   10
         Top             =   105
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   767
         Caption         =   "    &Delete"
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
         Picture         =   "trPembelianNonTunai.frx":059A
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   14265
         TabIndex        =   11
         Top             =   105
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
         Picture         =   "trPembelianNonTunai.frx":0824
      End
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   1170
         TabIndex        =   12
         Top             =   105
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   767
         Caption         =   "  &Edit"
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
         Picture         =   "trPembelianNonTunai.frx":09C3
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   435
         Left            =   105
         TabIndex        =   13
         Top             =   105
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   767
         Caption         =   "  &Add"
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
         Picture         =   "trPembelianNonTunai.frx":0AEF
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   15795
         TabIndex        =   14
         Top             =   105
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
         Picture         =   "trPembelianNonTunai.frx":0C9A
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   14700
         TabIndex        =   15
         Top             =   105
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
         Picture         =   "trPembelianNonTunai.frx":0D40
      End
   End
End
Attribute VB_Name = "trPembelianNonTunaiOri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nPos As SisPos
Dim lEdit As Boolean

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim objMenu As New CodeSuiteLibrary.Menu

Dim vaArray As New XArrayDB
Dim vaDelete As New XArrayDB
Dim vaExport As New XArrayDB

Dim cKode As String
Dim cID
Dim nSaldoStock As Double
Dim cStatusPembelian As String
Dim nDiskonExcel As Double
Dim nQtyTmp As Single

Dim nHargaJual As Double

'Dim Excel As Excel.Application
'Dim ExcelWBk As Excel.Workbook
'Dim ExcelWS As Excel.Worksheet

Private Sub StartExcel()
'  On Error GoTo err:
'  Set Excel = GetObject(, "Excel.Application")
'  Exit Sub
'err:
'  Set Excel = CreateObject("Excel.Application")
End Sub

Private Sub CloseWorkSheet()
'  ExcelWBk.Close
'  Excel.Quit
End Sub

Private Sub FinishExcel()
  'Jangan lupa, selalu bersihkan memory saat mengakhiri
'  If Not ExcelWS Is Nothing Then Set ExcelWS = Nothing
'  If Not ExcelWBk Is Nothing Then Set ExcelWBk = Nothing
'  If Not Excel Is Nothing Then Set Excel = Nothing
End Sub

'Private Sub GetLoadExcel()
'Dim lSave As Boolean
'Dim vaField, vaValue
'Dim i, j, n As Integer
'Dim dbData As New ADODB.Recordset
'Dim Wb As Excel.Workbook
'
'  On Error GoTo err:
'  StartExcel
'  lSave = True
'
'  Excel.Workbooks.Close
'  Set ExcelWBk = Excel.Workbooks.Open(CommonDialog1.FileName)
'  Set ExcelWS = ExcelWBk.Worksheets(1)
'
'
'  FrmPB.InitPB ExcelWS.Cells.SpecialCells(xlCellTypeLastCell).Row
'  Dim cBarcode
'  Dim cQty
'
'  For i = 1 To ExcelWS.Cells.SpecialCells(xlCellTypeLastCell).Row
'    FrmPB.RunPB
'    With ExcelWS
'      Set dbData = objData.Browse(GetDSN, "stock", "kodestock,nama,hargabeli,diskonpenjualan,kodesatuan,barcode", "barcode", sisAssign, .Cells(i, 1).Value)
'      If Not dbData.EOF Then
'        vaArray.InsertRows vaArray.UpperBound(1) + 1
'        n = vaArray.UpperBound(1)
'        vaArray(n, 0) = n + 1
'        vaArray(n, 1) = .Cells(i, 1).Value
'        vaArray(n, 2) = GetNull(dbData!nama)
'        vaArray(n, 3) = .Cells(i, 2).Value
'        vaArray(n, 4) = GetNull(dbData!kodesatuan)
'
''        If Trim(.Cells(i, 3)) = "" Then
''          MsgBox "empty"
''        Else
''          MsgBox Trim(.Cells(i, 3))
''        End If
'
''        vaArray(n, 5) = .Cells(i, 3).Value 'IIf(.Cells(i, 3) <> 0, .Cells(i, 3), GetNull(dbData!hargabeli))
''        vaArray(n, 6) = .Cells(i, 4).Value 'IIf(.Cells(i, 4) <> 0, .Cells(i, 4), IIf(Trim(nDiskonExcel) = "", 0, nDiskonExcel))
'
'        vaArray(n, 5) = IIf(Trim(.Cells(i, 3)) = "", GetNull(dbData!hargabeli), GetNull(.Cells(i, 3).Value))
'        vaArray(n, 6) = IIf(Trim(.Cells(i, 4)) = "", IIf(GetNull(dbData!diskonpenjualan) = 0, 0, GetNull(dbData!diskonpenjualan) + 3), GetNull(.Cells(i, 4).Value))
'
'        vaArray(n, 7) = (vaArray(n, 5) - (vaArray(n, 5) * vaArray(n, 6) / 100)) * vaArray(n, 3)
'
'        vaArray(n, 8) = GetNull(dbData!KodeStock)
'        vaArray(n, 9) = cID
'      Else
'        'jika data yg di import tidak ada dalam database simpan
'      End If
'    End With
'  Next i
'  nNomor.Value = vaArray.UpperBound(1) + 2
'  Set TDBGrid1.Array = vaArray
'  TDBGrid1.ReBind
'  TDBGrid1.Refresh
'  SumTotal
'  FrmPB.EndPB
'  CloseWorkSheet
'  FinishExcel
'
'err:
'End Sub

Private Sub BiSAButton1_Click()
  If MsgBox("Cetak Barcode?", vbYesNo) = vbYes Then
    Dim a As New exportExcel
    Dim na As Integer
    Dim ni As Single
    
'        vaExport.ReDim 0, 0, 0, 1
        vaExport.ReDim 0, nQtyTmp - 1, 0, 3
'        vaExport(0, 0) = "Balasan Order member " & cNamaCustomer.Text & " No: " & cFaktur.Text & " Tg. " & dTgl.Value
        Dim i As Single
        i = 0
        For na = vaArray.LowerBound(1) To vaArray.UpperBound(1)
'          vaExport.InsertRows na
          For ni = 1 To vaArray(na, 3)
           
            vaExport(i, 0) = "Rp. " & Format(vaArray(na, 5), "###,###,###")
            vaExport(i, 1) = vaArray(na, 1)
            vaExport(i, 2) = vaArray(na, 1) 'export harga jual
            vaExport(i, 3) = vaArray(na, 2)
            'vaExport(na, 3) = vaArray(na, 3) 'export diskon jual
            i = i + 1
          Next ni
          
        Next na
        
'    vaArray(n, 0) = nNomor.Value
'    vaArray(n, 1) = cBarcode.Text
'    vaArray(n, 2) = cNama.Text
'    vaArray(n, 3) = nQty.Value
'    vaArray(n, 4) = cSatuan.Text
'    vaArray(n, 5) = nHarga.Value
'    vaArray(n, 6) = nDisc1.Value
'    vaArray(n, 7) = nJumlah.Value
 '    vaArray(n, 8) = cKode
'    vaArray(n, 9) = cID
        
'      cfgStikerBarcode.PrintBarcode vaExport, 1  'GetOpt(opt)
        
      'a.RecordSource = vaExport
      'a.ExportToExcel
  End If
  
'  Dim n As Double
'  vaArray.ReDim 0, -1, 0, 2
'  For n = 0 To 100 - 1
'    vaArray.InsertRows n
'    vaArray(n, 0) = "A"
'    vaArray(n, 1) = "B"
'    vaArray(n, 2) = "C"
'  Next
  cfgStikerBarcode.PrintBarcode vaExport, 1
End Sub

Private Sub cAkunKas_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", "kodeakun,keterangan", "jenis", sisAssign, "D", , "kodeakun")
  If Not dbData.EOF Then
    cAkunKas.Text = cAkunKas.Browse(dbData)
  End If
End Sub

Private Sub cBarcode_ButtonClick()
  If Len(cBarcode.Text) >= 3 Then
    Set dbData = objData.Browse(GetDSN, "stock s", "s.barcode,s.nama,s.hargabeli,s.kodesatuan,s.hargajual,s.kodestock", "s.barcode", sisContent, cBarcode.Text, " AND s.jenis < 9 and s.statusnonaktif <> 1")
    If Not dbData.EOF Then
      cBarcode.Text = cBarcode.Browse(dbData, Array("BARCODE", "NAMA", "BELI", "SATUAN"), , Array(10, 35, 10, 8))
      GetDataStock
      SumJumlah
    Else
      cBarcode.Default
    End If
  Else
    MsgBox "Ketikkan 3 karakter atau lebih pencarian", vbCritical
  End If
End Sub

Private Sub SumBayar()
  nHutang.value = nTotal.value - IIf(nTunai.value > nTotal.value, nTotal.value, nTunai.value)
End Sub

Private Sub cBuyer_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "salesman", "kodesalesman,nama")
  If Not dbData.EOF Then
    cBuyer.Text = cBuyer.Browse(dbData)
  End If
End Sub

Private Sub cFaktur_ButtonClick()
Dim db As New ADODB.Recordset
Dim n As Integer
Dim lSave As Boolean

  If aCfg(objData, msOtorisasiPenuh) = "Y" Then
    If objMenu.GetPassword("", Me, GetDSN) Then
      If objMenu.UserLevel <> 0 Then
          MsgBox "Maaf, Anda tidak diberikan wewenang untuk melakukan PENGEDITAN." & vbCrLf & _
                 "Hanya user dengan LEVEL 0 (SUPERVISOR) yg diijinkan", vbInformation, "OTORISASI not ALLOWED"
          Exit Sub
'      Else
'        MsgBox "OTORISASI DIBATALKAN", vbCritical
'        Exit Sub
      End If
    Else
      Exit Sub
    End If
  End If

  lSave = True
  
  Set db = objData.Browse(GetDSN, "totpembelian", "nomorpembelian,tgl,subtotal,total,hutang", "nomorpembelian", sisContent, cFaktur.Text, " and tgl = '" & Format(dTgl.value, "yyyy-MM-dd") & "' and kodesupplier = '" & cSupplier.Text & "'")
  If Not db.EOF Then
    cFaktur.Text = cFaktur.Browse(db)
    'ambil nilai total
    Set db = objData.Browse(GetDSN, "totpembelian t", "t.*,g.keterangan as namagudang", "t.nomorpembelian", sisAssign, cFaktur.Text, , , Array("left join gudang g on g.kodegudang = t.kodegudang"))
    If Not db.EOF Then
      cStatusPembelian = GetNull(db!statuspembelian)
      cFakturAsli.Text = GetNull(db!fakturasli, "")
      dJthTmp.value = GetNull(db!jthtmp)
      nPersDisc.value = GetNull(db!PersDisc, 0)
      nPPn.value = GetNull(db!ppn, 0)
      nSubTotal.value = GetNull(db!Subtotal, 0)
      nDiscount.value = GetNull(db!Discount, 0)
      nPajak.value = GetNull(db!PAJAK, 0)
      nTotal.value = GetNull(db!Total, 0)
      nTunai.value = GetNull(db!Tunai, 0)
      nHutang.value = GetNull(db!hutang, "")
      cAkunKas.Text = GetNull(db!kodeakun)
      cBuyer.Text = GetNull(db!kodesalesman, "")
      cGudang.Text = GetNull(db!kodegudang, "")
      cNamaGudang.Text = GetNull(db!namagudang, "")
      If GetNull(db!hutang) = 0 Then
        chkTunai.value = 1
      Else
        chkTunai.value = 0
      End If
    End If
    'ambil nilai detail
    Dim nQtyTmp As Single
    nQtyTmp = 0
    Set db = objData.Browse(GetDSN, "pembelian p", "s.barcode,p.kodestock,s.Nama,p.qty,p.kodesatuan,p.harga,p.discount,p.jumlah", "nomorpembelian", sisAssign, cFaktur.Text, , , Array("Left join stock s on s.kodestock = p.kodestock"))
    If Not db.EOF Then
      vaArray.ReDim 0, -1, 0, 10
      Do While Not db.EOF
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        n = vaArray.UpperBound(1)
        vaArray(n, 0) = n + 1
        vaArray(n, 1) = GetNull(db!barcode)
        vaArray(n, 2) = GetNull(db!nama)
        vaArray(n, 3) = GetNull(db!qty)
        vaArray(n, 4) = GetNull(db!kodesatuan)
        vaArray(n, 5) = GetNull(db!Harga)
        vaArray(n, 6) = GetNull(db!Discount)
        vaArray(n, 7) = GetNull(db!jumlah)
        vaArray(n, 8) = GetNull(db!KodeStock)
        nQtyTmp = nQtyTmp + vaArray(n, 3)
        db.MoveNext
      Loop
      Set TDBGrid1.Array = vaArray
      TDBGrid1.ReBind
      TDBGrid1.Refresh
      TDBGrid1.Columns(3).FooterText = Format(nQtyTmp, "###,###,##0")
      Me.Refresh
      nNomor.value = vaArray.UpperBound(1) + 2
    End If
    
    If nPos = Delete Then
      If MsgBox("Data akan dihapus?", vbYesNo) = vbYes Then
        objData.Start GetDSN
        
        
        Dim cSQL As String
        cSQL = ""
        cSQL = " select distinct(nomorpelunasanhutang) as nomorpelunasanhutang from pelunasanhutang where nomorpembelian = '" & cFaktur.Text & "'"
        Set db = objData.SQL(GetDSN, cSQL)
        If Not db.EOF Then
          
          If MsgBox("Transaksi ini sudah pernah dilunasi sebelumnya!" & vbCrLf & "Dengan menghapus berarti seluruh data pelunasan yg berkenaan dengan transaksi ini akan ikut terhapus juga" & vbCrLf & "Apakah anda yakin akan menghapus?", vbYesNo) = vbYes Then
            lSave = IIf(lSave, objData.Delete(GetDSN, "totpelunasanhutang", "nomorpelunasanhutang", sisAssign, GetNull(db!nomorpelunasanhutang)), False)
            lSave = IIf(lSave, objData.Delete(GetDSN, "pelunasanhutang", "nomorpelunasanhutang", sisAssign, GetNull(db!nomorpelunasanhutang)), False)
            lSave = IIf(lSave, objData.Delete(GetDSN, "kartuHutang", "nomorkartuHutang", sisAssign, GetNull(db!nomorpelunasanhutang)), False)
            lSave = IIf(lSave, objData.Delete(GetDSN, "bukubesar", "faktur", sisAssign, GetNull(db!nomorpelunasanhutang)), False)
          Else
            MsgBox "Penghapusan dibatalkan"
            GetEdit False
            initvalue
            Exit Sub
          End If
        End If
        
        lSave = IIf(lSave, DelKodeTr(objData, msPembelian, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "pembelian", "nomorpembelian", sisAssign, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "kartuhutang", "nomorkartuhutang", sisAssign, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "kartustock", "nomor", sisAssign, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "totpembelian", "nomorpembelian", sisAssign, cFaktur.Text), False)
        If lSave Then
          objData.Save GetDSN
          
          lSave = True
          objData.Start GetDSN
          
          'LAKUKAN UPDATE HARGA POKOK UNTUK MASING MASING PRODUK
          For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
            lSave = IIf(lSave, NewUpdHargaPokok(objData, vaArray(n, 8)), False)
          Next n
          
          If lSave Then
            objData.Save GetDSN
          Else
            objData.Cancel GetDSN
          End If
        
        Else
          objData.Cancel GetDSN
        End If
      End If
      GetEdit False
      initvalue
    End If
    If nPos = Edit Then
      SendKeysA vbKeyReturn, True
    End If
  End If
End Sub

Private Sub cFaktur_Validate(Cancel As Boolean)
  If nPos = Edit Then
      cFaktur.Enabled = False
      cFaktur.BackColor = vbButtonFace
  End If
End Sub

Private Sub cGudang_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "gudang", "kodegudang,keterangan", "lstatus", sisAssign, "A")
  If Not dbData.EOF Then
    cGudang.Text = cGudang.Browse(dbData)
    cGudang.Text = GetNull(dbData!kodegudang)
    cNamaGudang.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub chkTunai_Click()
  If chkTunai.value = 1 Then
    nTunai.value = nTotal.value
    nHutang.value = 0
  Else
    nTunai.value = 0
    nHutang.value = nTotal.value
  End If
End Sub

Private Sub GetDataStock()
  cBarcode.Text = GetNull(dbData!barcode, "")
  cKode = GetNull(dbData!KodeStock, "")
  cNama.Text = GetNull(dbData!nama, "")
  cSatuan.Text = GetNull(dbData!kodesatuan, "")
  nHarga.value = GetNull(dbData!hargabeli, 0)
  
End Sub

Private Function GetReplaceDataMySQL(cData) As Double
  GetReplaceDataMySQL = Replace(cData, ",", "")
  GetReplaceDataMySQL = Replace(cData, ".", "")
End Function

Private Sub cmdAdd_Click()
  nPos = Add
  GetEdit True
  GetFakturBrowse False
  cFaktur.Text = CreateNomorFaktur(objData, sisModulTransaksi.pembelian, "totpembelian", "nomorpembelian")
  chkTunai.Enabled = True
'  cmdGetOrder.Visible = True
End Sub

Private Sub cmdAktivasi_Click()
  frmAktivasi.Action Me
End Sub

Private Sub cmdEdit_Click()
  If aCfg(objData, msBisaEditPembelian) = "T" Then
    If GetRegistry(reg_UserLevel) <> 0 Then
      If objMenu.GetPassword("", Me, GetDSN) Then
        If objMenu.UserLevel <> 0 Then
          MsgBox "Maaf, Anda tidak diberikan wewenang untuk melakukan PENGEDITAN." & vbCrLf & _
                 "Hanya user dengan LEVEL 0 (SUPERVISOR) yg diijinkan", vbInformation, "OTORISASI not ALLOWED"
          Exit Sub
        End If
      Else
        Unload Me
        GetEdit False
        Exit Sub
      End If
    End If
  End If
  
  nPos = Edit
  GetEdit True
  
  GetFakturBrowse True
'  chkTunai.Enabled = False
  cmdGetOrder.Visible = False
End Sub

Private Sub GetFakturBrowse(ByVal lStat As Boolean)
  cFaktur.Button = lStat
  If lStat = True Then
    cFaktur.BackColor = vbWindowBackground
  Else
    cFaktur.BackColor = vbButtonFace
  End If
End Sub


Private Sub GetEdit(lPar As Boolean)
  'BiSAFrame1.Enabled = lPar
  Frame3.Enabled = lPar
  Frame4.Enabled = lPar
  BiSAFrame2.Enabled = lPar
  lEdit = lPar
  initvalue

  SetButton cmdSimpan, cmdKeluar, cmdAdd, cmdEdit, cmdHapus, nPos, lPar, cmdAktivasi
  GetFakturBrowse False
  If lPar Then
    dTgl.SetFocus
    If nPos = Add Then
      cFaktur.Enabled = False
      cFaktur.BackColor = vbButtonFace
    Else
      cFaktur.Enabled = True
      cFaktur.BackColor = vbWindowBackground
      cFaktur.CaptionBackColor = vbButtonFace
    End If
  End If
End Sub

Private Sub cmdGetOrder_Click()
Dim n As Integer

  If MsgBox("Apakah anda yakin akan mengambil data dari purchase order yang outstanding?", vbYesNo) = vbYes Then
    vaArray.ReDim 0, -1, 0, 10
    Set dbData = objData.Browse(GetDSN, "po p", "p.id,s.kodestock,s.barcode,s.nama,p.qty,s.kodesatuan,p.harga,p.diskonpenjualan", "p.statuspembelian", sisAssign, 0, " and p.statusorder = 1", , Array("left join stock s on s.kodestock = p.kodestock"))
    If Not dbData.EOF Then
      cStatusPembelian = 1
      Do While Not dbData.EOF
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        n = vaArray.UpperBound(1)
        vaArray(n, 0) = n + 1
        vaArray(n, 1) = GetNull(dbData!barcode)
        vaArray(n, 2) = GetNull(dbData!nama)
        vaArray(n, 3) = GetNull(dbData!qty)
        vaArray(n, 4) = GetNull(dbData!kodesatuan)
        vaArray(n, 5) = GetNull(dbData!Harga)
        vaArray(n, 6) = GetNull(dbData!diskonpenjualan)
        vaArray(n, 7) = (vaArray(n, 5) - (vaArray(n, 5) * vaArray(n, 6) / 100)) * vaArray(n, 3)
        vaArray(n, 8) = GetNull(dbData!KodeStock)
        vaArray(n, 9) = GetNull(dbData!ID)
        dbData.MoveNext
      Loop
    End If
    SumTotal
    Set TDBGrid1.Array = vaArray
    TDBGrid1.ReBind
    TDBGrid1.Refresh
  End If
End Sub

Private Sub cmdHapus_Click()
  If aCfg(objData, msBisaEditPembelian) = "T" Then
    If GetRegistry(reg_UserLevel) <> 0 Then
      If objMenu.GetPassword("", Me, GetDSN) Then
        If objMenu.UserLevel <> 0 Then
          MsgBox "Maaf, Anda tidak diberikan wewenang untuk melakukan PENGHAPUSAN." & vbCrLf & _
                 "Hanya user dengan LEVEL 0 (SUPERVISOR) yg diijinkan", vbInformation, "OTORISASI not ALLOWED"
          Exit Sub
        End If
      Else
        Unload Me
        GetEdit False
        Exit Sub
      End If
    End If
  End If
  nPos = Delete
  GetEdit True
  GetFakturBrowse True
  cmdGetOrder.Visible = False
End Sub

Private Sub cmdImportWizard_Click()
  CommonDialog1.Filter = "Excel File (*.xls)|*.xls"
  CommonDialog1.ShowOpen
  If CommonDialog1.FileName <> "" Then
'    GetLoadExcel
  End If
End Sub

Private Sub cmdKeluar_Click()
  If lEdit Then
    GetEdit False
    initvalue
  Else
    Unload Me
  End If
End Sub

Private Function validOK() As Boolean
  validOK = True
  If Not GetValidDataBrowse(objData, "stock", "kodestock", cKode) Then
    MsgBox "Barang tersebut tidak ada dalam database"
    cBarcode.SetFocus
    validOK = False
    Exit Function
  End If
End Function

Private Sub cmdOK_Click()
Dim n As Integer
Dim nJumlah1 As Double


  If validOK() Then
   
    'jika baris <= Nomor
    If vaArray.UpperBound(1) + 2 <= nNomor.value Then
      vaArray.ReDim 0, nNomor.value - 1, 0, 10
      n = vaArray.UpperBound(1)
    'jika baris=0
    ElseIf vaArray.UpperBound(1) = -1 Then
      nNomor.value = 1
      vaArray.ReDim 0, nNomor.value - 1, 0, 10
      n = vaArray.UpperBound(1)
    Else
      n = nNomor.value - 1
    End If
        
    vaArray(n, 0) = nNomor.value
    vaArray(n, 1) = cBarcode.Text
    vaArray(n, 2) = cNama.Text
    vaArray(n, 3) = nQty.value
    vaArray(n, 4) = cSatuan.Text
    vaArray(n, 5) = nHarga.value
    vaArray(n, 6) = nDisc1.value
    vaArray(n, 7) = nJumlah.value
    vaArray(n, 8) = cKode
    vaArray(n, 9) = cID
    vaArray(n, 10) = 0 'untuk menampung selisih harga pokok
    
    TDBGrid1.Array = vaArray
    TDBGrid1.ReBind
    TDBGrid1.MoveNext
    
    nJumlah1 = 0
    nQtyTmp = 0
    For n = 0 To vaArray.UpperBound(1)
      nJumlah1 = nJumlah1 + vaArray(n, 7)
      nQtyTmp = nQtyTmp + vaArray(n, 3)
    Next
    nSubTotal.value = nJumlah1
    TDBGrid1.Columns(3).FooterText = Format(nQtyTmp, "###,###,##0")
    
    SumTotal
    
    InitValue1
    
    nNomor.value = vaArray.UpperBound(1) + 2
    cBarcode.SetFocus
  End If
End Sub

Private Sub SumTotal()
Dim n As Double
  
  nSubTotal.value = 0
  For n = 0 To vaArray.UpperBound(1)
    nSubTotal.value = nSubTotal.value + vaArray(n, 7)
  Next
  
  If nPersDisc.Enabled = True Then
    nDiscount.value = nPersDisc.value / 100 * (nSubTotal.value)
  End If
  
  nPajak.value = (nPPn.value / 100) * (nSubTotal.value - (nDiscount.value + nDiscount.value))
  nTotal.value = nSubTotal.value + nPajak.value - nDiscount.value
  If chkTunai.value = 1 Then
    nTunai.value = nTotal.value
    nHutang.value = 0
  Else
    nHutang.value = nTotal.value
    nTunai.value = 0
  End If
End Sub

Private Function ValidSaving() As Boolean
Dim n As Integer

  ValidSaving = True
  
  If vaArray.UpperBound(1) < 0 Then
    MsgBox "Nota kosong, data tidak disimpan"
    ValidSaving = False
    Exit Function
  End If
  
  If Trim(cFaktur.Text) = "" Then
     MsgBox "Maaf Nomor Faktur Kosong/Tidak Valid" & vbCrLf & "Data tidak bisa disimpan"
     ValidSaving = False
     Exit Function
  End If
  
  If cSupplier.Text = "" Then
    MsgBox "Kode Supplier tidak terisi" & vbCrLf & "Data tidak bisa disimpan", vbExclamation
    ValidSaving = False
    Exit Function
  End If
  
  If cAkunKas.Text = "" Then
    MsgBox "Akun Kas tidak terisi" & vbCrLf & "Data tidak bisa disimpan", vbExclamation
    ValidSaving = False
    Exit Function
  End If
  
  If Not GetValidDataBrowse(objData, "supplier", "kodesupplier", cSupplier.Text) Then
    MsgBox "Maaf, data supplier tidak tertera dengan benar" & vbCrLf & "Data tidak bisa disimpan", vbExclamation
    ValidSaving = False
    Exit Function
  End If
  
  If Not GetValidDataBrowse(objData, "gudang", "kodegudang", cGudang.Text) Then
    MsgBox "Kode gudang tidak tertera dengan benar" & vbCrLf & "Data tidak bisa disimpan", vbExclamation
    cGudang.SetFocus
    ValidSaving = False
    Exit Function
  End If
  
  If Not GetValidDataBrowse(objData, "akun", "kodeakun", cAkunKas.Text) Then
    MsgBox "Kode akun tidak tertera dengan benar" & vbCrLf & "Data tidak bisa disimpan", vbExclamation
    ValidSaving = False
    Exit Function
  End If
  
'  If Not GetValidDataBrowse(objData, "salesman", "kodesalesman", cBuyer.Text) Then
'    MsgBox "Kode buyer tidak tertera dengan benar" & vbCrLf & "Data tidak bisa disimpan", vbExclamation
'    ValidSaving = False
'    cBuyer.SetFocus
'    Exit Function
'  End If
  
  'Jika kode gudang tidak valid, maka penyimpanan data tidak diijinkan
  Set dbData = objData.Browse(GetDSN, "gudang", "lstatus", "kodegudang", sisAssign, cGudang.Text)
  If Not dbData.EOF Then
    If GetNull(dbData!lStatus) <> "A" Then
      MsgBox "Kode Gudang tidak valid, atau tidak aktif, Data tidak bisa disimpan", vbExclamation
      ValidSaving = False
      Exit Function
    End If
  End If
End Function

Private Sub cmdSimpan_Click()
Dim lSave As Boolean
Dim Faktur As String
Dim n As Integer

lSave = True
  'simpan pada tabel totpembelian
  'simpan pada tabel pembelian
  'simpan pada tabel kartustock
  'simpan pada tabel kartuhutang
  
'    vaArray(n, 0) = nNomor.Value
'    vaArray(n, 1) = cBarcode.Text
'    vaArray(n, 2) = cNama.Text
'    vaArray(n, 3) = nQty.Value
'    vaArray(n, 4) = cSatuan.Text
'    vaArray(n, 5) = nHarga.Value
'    vaArray(n, 6) = nDisc1.Value
'    vaArray(n, 7) = nJumlah.Value
'    vaArray(n, 8) = cKode
'    vaArray(n, 9) = cID
  
  If ValidSaving Then
    'MsgBox "Hello"
    GetNotifikasiAdd "Menyimpan Pembelian"
    objData.Start GetDSN
    Faktur = cFaktur.Text
        
    lSave = IIf(lSave, objData.Update(GetDSN, "totpembelian", "nomorpembelian = '" & Faktur & "'", Array("nomorpembelian", "fakturasli", "tgl", "jthtmp", "kodesupplier", "ppn", "persdisc", "persdisc2", "subtotal", "pajak", "discount", "discount2", "total", "tunai", "hutang", "datetime", "username", "kodeakun", "kodecostcenter", "kodesalesman", "statuspembelian", "kodegudang", "kodegroupsales"), Array(Faktur, cFakturAsli.Text, Format(dTgl.value, "yyyy-MM-dd"), Format(dJthTmp.value, "yyyy-MM-dd"), cSupplier.Text, nPPn.value, nPersDisc.value, 0, nSubTotal.value, nPajak.value, nDiscount.value, 0, nTotal.value, nTunai.value, nHutang.value, SNow, GetRegistry(reg_Username), cAkunKas.Text, GetCostCenterUser(objData, GetRegistry(reg_Username)), cBuyer.Text, cStatusPembelian, cGudang.Text, GetRegistry(reg_KodeGroupSalesPembelian))), False)
    lSave = IIf(lSave, objData.Delete(GetDSN, "pembelian", "nomorpembelian", sisAssign, Faktur), False)
    lSave = IIf(lSave, objData.Delete(GetDSN, "kartustock", "nomor", sisAssign, Faktur), False)
    
    For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
      lSave = IIf(lSave, objData.Add(GetDSN, "pembelian", Array("nomorpembelian", "kodegudang", "tgl", "kodestock", "qty", "harga", "kodesatuan", "discount", "jumlah"), Array(Faktur, cGudang.Text, Format(dTgl.value, "yyyy-MM-dd"), vaArray(n, 8), vaArray(n, 3), vaArray(n, 5), vaArray(n, 4), vaArray(n, 6), vaArray(n, 7))), False)
      
      'Update harga beli dengan harga beli terakhir
      lSave = IIf(lSave, objData.Edit(GetDSN, "stock", "kodestock = '" & vaArray(n, 8) & "'", Array("hargabeli"), Array(vaArray(n, 5) - (vaArray(n, 5) * vaArray(n, 6) / 100))), False)
      
      '***PENTING***
      'UPDATE KARTU STOCK
      'Cek nilai persediaan terlebih dahulu
      'Jika nilai persediaan minus, gunakan HPP baru dan jika tidak gunakan Harga beli untuk menambah nilai persediaan
      '------------------------------------------------------------------------
      If GetSaldoStock(objData, "", vaArray(n, 8)) < 0 Then
        'vaArray(n, 5) = NewUpdHargaPokok(objData, vaArray(n, 8))
        vaArray(n, 10) = vaArray(n, 7)
        '***PENTING***
        vaArray(n, 7) = NewUpdHargaPokok(objData, vaArray(n, 8)) * vaArray(n, 3)
        vaArray(n, 10) = vaArray(n, 10) - vaArray(n, 7) 'temukan selisih cogs akibat stok mines
        
        lSave = IIf(lSave, UpdKartuStock(objData, SisKartuStock.pembelian, Faktur, Format(dTgl.value, "yyyy-MM-dd"), vaArray(n, 8), vaArray(n, 3), vaArray(n, 5), vaArray(n, 6), "Pembelian Non Tunai an. " & cNamaSupplier.Text & " Gudang " & cGudang.Text, cGudang.Text, NewUpdHargaPokok(objData, vaArray(n, 8))), False)
        'update harga cogs dengan yg terakhirm
        lSave = IIf(lSave, objData.Edit(GetDSN, "stock", "kodestock = '" & vaArray(n, 8) & "'", Array("cogs"), Array(NewUpdHargaPokok(objData, vaArray(n, 8)))), False)
      Else
        '***PENTING***
        lSave = IIf(lSave, UpdKartuStock(objData, SisKartuStock.pembelian, Faktur, Format(dTgl.value, "yyyy-MM-dd"), vaArray(n, 8), vaArray(n, 3), vaArray(n, 5), vaArray(n, 6), "Pembelian Non Tunai an. " & cNamaSupplier.Text & " Gudang " & cGudang.Text, cGudang.Text, vaArray(n, 5) - (vaArray(n, 5) * vaArray(n, 6) / 100)), False)
        'update harga cogs dengan yg terakhirm
        lSave = IIf(lSave, objData.Edit(GetDSN, "stock", "kodestock = '" & vaArray(n, 8) & "'", Array("cogs"), Array(vaArray(n, 5) - (vaArray(n, 5) * vaArray(n, 6) / 100))), False)
      End If
      
      

    Next n
    If cStatusPembelian = 1 Then
      'jika statuspembelian = 1 maka update tabel po juga
      For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
        lSave = IIf(lSave, objData.Update(GetDSN, "po", "id = " & vaArray(n, 9), Array("statuspembelian", "fakturpembelian"), Array(1, Faktur)), False)
      Next n
      
      'update status cancel jika ada
      For n = vaDelete.LowerBound(1) To vaDelete.UpperBound(1)
        lSave = IIf(lSave, objData.Update(GetDSN, "po", "id = " & vaDelete(n, 1), Array("statuscancel", "fakturpembelian"), Array(1, Faktur)), False)
        lSave = IIf(lSave, objData.Update(GetDSN, "po", "id = " & vaDelete(n, 1), Array("statuspembelian", "fakturpembelian"), Array(1, Faktur)), False)
      Next n
    End If
    
    'isi field flaglunas
    lSave = IIf(lSave, objData.Edit(GetDSN, "totpembelian", "nomorpembelian = '" & Faktur & "'", Array("flaglunas"), Array(0)), False)
    If chkTunai.value = 1 Then
      lSave = IIf(lSave, objData.Edit(GetDSN, "totpembelian", "nomorpembelian = '" & Faktur & "'", Array("flaglunas"), Array(1)), False)
    Else
      If lCekStatusLunasHutang(objData, Faktur) = True Then
        lSave = IIf(lSave, objData.Edit(GetDSN, "totpembelian", "nomorpembelian = '" & Faktur & "'", Array("flaglunas"), Array(1)), False)
      End If
    End If
  
    lSave = IIf(lSave, UpdKartuHutang(objData, Sispembelian, Faktur, Format(dTgl.value, "yyyy-MM-dd"), cSupplier.Text, "Pembelian Non Tunai an. " & cNamaSupplier.Text, nHutang.value, SNow, GetRegistry(reg_Username)), False)
    
    ' Inventory (1)
    ' Purchase Tax (2)
    ' Non Inventory Expenses (5)
    '    Acc Payable (2)
    '    Cash Bank (1)
    
    'Posting inventory
    'Hapus dulu di bukubesar
    
    lSave = IIf(lSave, DelKodeTr(objData, msPembelian, Faktur), False)
    'Debet
    For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
      Dim db As New ADODB.Recordset
      
      Set db = objData.Browse(GetDSN, "stock", "kodestock,asbiaya", "kodestock", sisAssign, vaArray(n, 8))
      If Not db.EOF Then
        If GetNull(db!asbiaya) = "1" Then
          
'------------------------------------------
'Remark akuntansi sebelumnya
'==========================================
'          'Konfig 1
'          If aCfg(objData, msJenisDiscountPembelian) = "Y" Then
'              lSave = IIf(lSave, UpdKodeTr(objData, msPembelian, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningBiayaBarang), aCfg(objData, msCostCenterJualBeli), "Pembelian Inventory an " & cNamaSupplier.Text, vaArray(n, 3) * vaArray(n, 5), 0, "", SNow), False)
'              'Discount Pembelian per item
'              lSave = IIf(lSave, UpdKodeTr(objData, msPembelian, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningDiscountPembelian), aCfg(objData, msCostCenterJualBeli), "Dsc Itm Pembelian an " & cNamaSupplier.Text, 0, vaArray(n, 3) * vaArray(n, 5) - vaArray(n, 7), "", SNow), False)
'
'          'Konfig 2 No Diskon
'          Else
'              lSave = IIf(lSave, UpdKodeTr(objData, msPembelian, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningBiayaBarang), aCfg(objData, msCostCenterJualBeli), "Pembelian Inventory an " & cNamaSupplier.Text, (vaArray(n, 3) * vaArray(n, 5)) - (vaArray(n, 3) * vaArray(n, 5) * vaArray(n, 6) / 100), 0, "", SNow), False)
'          End If
'------------------------------------------

'          lSave = IIf(lSave, UpdKodeTr(objData, msPembelian, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningBiayaBarang), aCfg(objData, msCostCenterJualBeli), "Pembelian Inventory an " & vaArray(n, 2), vaArray(n, 3) * vaArray(n, 5) , 0, "", SNow, vaArray(n, 8)), False)
          lSave = IIf(lSave, UpdKodeTr(objData, msPembelian, Faktur, Format(dTgl.value, "yyyy-MM-dd"), aCfg(objData, msRekeningBiayaBarang), GetCostCenterUser(objData, GetRegistry(reg_Username)), "Pembelian Inventory an " & vaArray(n, 2), vaArray(n, 7), 0, "", SNow, vaArray(n, 8)), False)
'         Discount Pembelian per item
'          lSave = IIf(lSave, UpdKodeTr(objData, msPembelian, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningDiscountPembelian), aCfg(objData, msCostCenterJualBeli), "Dsc Itm Pembelian an " & cNamaSupplier.Text, 0, vaArray(n, 3) * vaArray(n, 5) - vaArray(n, 7), "", SNow, vaArray(n, 8)), False)

          
        Else
'------------------------------------------
'Remark akuntansi sebelumnya
'==========================================
'          'Konfig 1
'           If aCfg(objData, msJenisDiscountPembelian) = "Y" Then
'              lSave = IIf(lSave, UpdKodeTr(objData, msPembelian, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), GetAkunInventory(objData, vaArray(n, 8)), aCfg(objData, msCostCenterJualBeli), "Pembelian Inventory an " & cNamaSupplier.Text, vaArray(n, 3) * vaArray(n, 5), 0, "", SNow), False)
'              'Discount Pembelian per item
'              lSave = IIf(lSave, UpdKodeTr(objData, msPembelian, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningDiscountPembelian), aCfg(objData, msCostCenterJualBeli), "Dsc Itm Pembelian an " & cNamaSupplier.Text, 0, vaArray(n, 3) * vaArray(n, 5) - vaArray(n, 7), "", SNow), False)
'
'          'Konfig 2 No Diskon
'          Else
'              lSave = IIf(lSave, UpdKodeTr(objData, msPembelian, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), GetAkunInventory(objData, vaArray(n, 8)), aCfg(objData, msCostCenterJualBeli), "Pembelian Inventory an " & cNamaSupplier.Text, (vaArray(n, 3) * vaArray(n, 5)) - (vaArray(n, 3) * vaArray(n, 5) * vaArray(n, 6) / 100), 0, "", SNow), False)
'          End If
'-------------------------------------------
          
'          lSave = IIf(lSave, UpdKodeTr(objData, msPembelian, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), GetAkunInventory(objData, vaArray(n, 8)), aCfg(objData, msCostCenterJualBeli), "Pembelian Inventory an " & vaArray(n, 2), vaArray(n, 3) * vaArray(n, 5), 0, "", SNow, vaArray(n, 8)), False)
          lSave = IIf(lSave, UpdKodeTr(objData, msPembelian, Faktur, Format(dTgl.value, "yyyy-MM-dd"), GetAkunInventory(objData, vaArray(n, 8)), GetCostCenterUser(objData, GetRegistry(reg_Username)), "Pembelian Inventory an " & vaArray(n, 2), vaArray(n, 7), 0, "", SNow, vaArray(n, 8)), False)
          If vaArray(n, 10) <> 0 Then
            lSave = IIf(lSave, UpdKodeTr(objData, msPembelian, Faktur, Format(dTgl.value, "yyyy-MM-dd"), aCfg(objData, msRekeningCOGS), GetCostCenterUser(objData, GetRegistry(reg_Username)), "COGS pembelian an " & vaArray(n, 2), vaArray(n, 10), 0, "", SNow, vaArray(n, 8)), False)
          End If
          'Discount Pembelian per item
'          lSave = IIf(lSave, UpdKodeTr(objData, msPembelian, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningDiscountPembelian), aCfg(objData, msCostCenterJualBeli), "Dsc Itm Pembelian an " & cNamaSupplier.Text, 0, vaArray(n, 3) * vaArray(n, 5) - vaArray(n, 7), "", SNow, vaArray(n, 8)), False)

          
        End If
      End If
      
      'MsgBox lSave
      'Update COGS pada tabel stock
      'lSave = UpdHargaPokok(objData, vaArray(n, 8), vaArray(n, 3), vaArray(n, 5))
      
    Next n
    
    'PPn
    lSave = IIf(lSave, UpdKodeTr(objData, msPembelian, Faktur, Format(dTgl.value, "yyyy-MM-dd"), aCfg(objData, msRekeningPPnPembelian), GetCostCenterUser(objData, GetRegistry(reg_Username)), "PPn Pembelian an " & cNamaSupplier.Text, nPajak.value, 0, "", SNow), False)
    'Discount seluruhnya
    lSave = IIf(lSave, UpdKodeTr(objData, msPembelian, Faktur, Format(dTgl.value, "yyyy-MM-dd"), aCfg(objData, msRekeningDiscountPembelian), GetCostCenterUser(objData, GetRegistry(reg_Username)), "Dsc Tot Pembelian an " & cNamaSupplier.Text, 0, nDiscount.value, "", SNow), False)
    
    'Kredit
    'Hutang
    lSave = IIf(lSave, UpdKodeTr(objData, msPembelian, Faktur, Format(dTgl.value, "yyyy-MM-dd"), GetAkunSupplier(objData, cSupplier.Text), GetCostCenterUser(objData, GetRegistry(reg_Username)), "Hutang Pembelian an " & cNamaSupplier.Text, 0, nHutang.value, "", SNow), False)
    'kas
    lSave = IIf(lSave, UpdKodeTr(objData, msPembelian, Faktur, Format(dTgl.value, "yyyy-MM-dd"), cAkunKas.Text, GetCostCenterUser(objData, GetRegistry(reg_Username)), "Kas untuk pembelian an " & cNamaSupplier.Text, 0, nTunai.value, "", SNow), False)
    
    If lSave Then
      objData.Save GetDSN
      UpdateHargaPokokStockTrPembelian vaArray
      GetNotifikasiRemove
    Else
      MsgBox "Data tidak berhasil disimpan", vbCritical
      objData.Cancel GetDSN
    End If
    
    If aCfg(objData, msCetakPembelian) = "Y" Then
      GetCetak Faktur
    End If
    initvalue
    GetEdit False
    
  End If
End Sub

Private Sub UpdateHargaPokokStockTrPembelian(ByVal vaArrayHP As XArrayDB)
Dim n As Single

'    vaArray(n, 0) = nNomor.Value
'    vaArray(n, 1) = cBarcode.Text
'    vaArray(n, 2) = cNama.Text
'    vaArray(n, 3) = nQty.Value
'    vaArray(n, 4) = cSatuan.Text
'    vaArray(n, 5) = nHarga.Value
'    vaArray(n, 6) = nDisc1.Value
'    vaArray(n, 7) = nJumlah.Value
'    vaArray(n, 8) = cKode
'    vaArray(n, 9) = cID
  
  
  'update harga pokok pada tabel stock untuk masing masing barang
  For n = vaArrayHP.LowerBound(1) To vaArrayHP.UpperBound(1)
'      MsgBox vaArrayHP(n, 2)
      'UpdHargaPokok objData, vaArrayHP(n, 8), vaArrayHP(n, 3), vaArrayHP(n, 5)
      objData.Edit GetDSN, "stock", "kodestock = '" & vaArray(n, 9) & "'", Array("cogs"), Array(NewUpdHargaPokok(objData, vaArray(n, 9)))
  Next n
End Sub

Private Sub GetCetak(ByVal cFak As String)
  trPrintPembelian.noOrder = cFak
  Set dbData = objData.Browse(GetDSN, "totpembelian t", "t.*,a.*", "t.nomorpembelian", sisAssign, cFak, , , Array("left join supplier a on a.kodesupplier = t.kodesupplier"))
  If Not dbData.EOF Then
    trPrintPembelian.nSubTotal = GetNull(dbData!Subtotal)
    trPrintPembelian.nDiscount = 0
    trPrintPembelian.nCash = GetNull(dbData!Tunai)
    trPrintPembelian.nChange = GetNull(dbData!hutang)
    trPrintPembelian.cKodeMember = GetNull(dbData!kodesupplier)
    trPrintPembelian.cMember = GetNull(dbData!nama)
    trPrintPembelian.cTeleponMember = 0
    trPrintPembelian.Ups = 0
    Load trPrintPembelian
    trPrintPembelian.Show vbModal
  End If
End Sub

Private Sub cSupplier_ButtonClick()
'  Set dbData = objData.Browse(GetDSN, "supplier", "kodesupplier,nama,alamat,kota", "kodesupplier", sisContent, cSupplier.Text, , "kodesupplier,nama")
'  If Not dbData.EOF Then
'    cSupplier.Text = cSupplier.Browse(dbData)
'    cSupplier.Text = GetNull(dbData!kodesupplier)
'    cNamaSupplier.Text = GetNull(dbData!nama, "")
'    cAlamat.Text = GetNull(dbData!alamat, "")
'    cKota.Text = GetNull(dbData!kota, "")
'  End If
End Sub

Private Sub cNama_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "stock s", "s.Barcode,s.nama,s.hargabeli,s.kodesatuan,s.hargajual,s.kodestock", "s.nama", sisContent, cNama.Text, " AND s.jenis < 9 and s.statusnonaktif <> 1")
  If Not dbData.EOF Then
    cNama.Text = cNama.Browse(dbData, Array("BARCODE", "NAMA", "BELI", "SATUAN"), , Array(10, 35, 10, 8))
    GetDataStock
    SumJumlah
  Else
    cNama.Default
  End If
End Sub

Private Sub cNamaSupplier_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "supplier", "kodesupplier,nama,alamat,kota", "nama", sisContent, cNamaSupplier.Text, " or kodesupplier like '%" & cNamaSupplier.Text & "%'", "kodesupplier,nama")
  If Not dbData.EOF Then
    cNamaSupplier.Text = cNamaSupplier.Browse(dbData)
    cSupplier.Text = GetNull(dbData!kodesupplier)
    cNamaSupplier.Text = GetNull(dbData!nama, "")
    cAlamat.Text = GetNull(dbData!alamat, "")
    cKota.Text = GetNull(dbData!kota, "")
  End If
End Sub

Private Sub dTgl_Validate(Cancel As Boolean)
  If Not IsInPeriod(dTgl.value) Or (dTgl.value > Date) Then
    Cancel = True
    dTgl.SetFocus
    'GetEdit False
  End If
End Sub

Private Sub Form_Activate()
  Me.Caption = "PEMBELIAN * " & GetRegistry(reg_KodeGroupSalesPembelian)
End Sub

Private Sub Form_Load()
Dim n As Single

'  If CheckTrial(nRecordsTrial, TrialPembelian) = True Then
'    End
'  End If

  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  GetEdit False
  initvalue
  
  
  Set dbData = objData.Browse(GetDSN, "costcenter", , "kodecostcenter", sisAssign, GetCostCenterUser(objData, GetRegistry(reg_Username)))
  If Not dbData.EOF Then
    Frame3.Caption = "Cost Centre : " & GetNull(dbData!keterangan)
  End If
  
  TabIndex dTgl, n
  TabIndex cSupplier, n
  TabIndex cNamaSupplier, n
  TabIndex cFaktur, n
  TabIndex cAlamat, n
  TabIndex nPersDisc, n
  
  TabIndex cFakturAsli, n
  'TabIndex dJthTmp, n
  'TabIndex nPPn, n
  'TabIndex cAkunKas, n
  'TabIndex cGudang, n
  'TabIndex cNamaGudang, n
  'TabIndex cBuyer, n
  
  
  TabIndex nNomor, n
  TabIndex cBarcode, n
  TabIndex cNama, n
  TabIndex nQty, n
  TabIndex cSatuan, n
  TabIndex nHarga, n
  TabIndex nDisc1, n
  TabIndex nJumlah, n
  TabIndex cmdOK, n
  TabIndex nTunai, n
  
  TabIndex cmdAdd, n
  TabIndex cmdEdit, n
  TabIndex cmdHapus, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  TabIndex cmdAktivasi, n
End Sub

Private Sub initvalue()
Dim dbgudang As New ADODB.Recordset

  cmdGetOrder.Visible = False
  cStatusPembelian = 0
  cFaktur.Default
  dTgl.value = Date
  dJthTmp.value = Date
  nPersDisc.value = 0
  nPPn.value = 0
  cFakturAsli.Default
  cSupplier.Default
  cNamaSupplier.Default
  cBuyer.Default
  cAlamat.Default
  cKota.Default
  nSubTotal.value = 0
  nPajak.value = 0
  nDiscount.value = 0
  nTotal.value = 0
  nTunai.value = 0
  nHutang.value = 0
  chkTunai.value = 0
  chkTunai.Enabled = True
  cAkunKas.Text = cKasTeller
  cGudang.Text = aCfg(objData, msGudangPembelian)
  cNamaGudang.Default
  Set dbgudang = objData.Browse(GetDSN, "gudang", "kodegudang,keterangan", "kodegudang", sisAssign, cGudang.Text)
  If Not dbgudang.EOF Then
    cNamaGudang.Text = GetNull(dbgudang!keterangan)
  End If

  
  vaArray.ReDim 0, -1, 0, 10
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
  InitValue1
  vaDelete.ReDim 0, -1, 0, 1
  
  TDBGrid1.Columns(3).FooterText = ""
  nQty.Decimals = aCfg(objData, msNilaiDecimals)
  nDisc1.Enabled = True
  nDisc1.BackColor = vbWhite
  If aCfg(objData, msEnableDisableDiscountItemPembelian) = "D" Then
    nDisc1.Enabled = False
  End If
  chkTunai.value = 0
  If nPos = Add Then
    If aCfg(objData, msDefaultPembelian) = "T" Then
      chkTunai.value = 1
    End If
  End If
  
  cGudang.Enabled = True
  cGudang.BackColor = vbWhite
  If GetRegistry(reg_UserLevel) <> 0 Then
  
    cGudang.Enabled = False
    cGudang.BackColor = vbButtonFace
  End If

  cGudang.Text = GetGudangUser(objData, GetRegistry(reg_Username))
  Set dbData = objData.Browse(GetDSN, "gudang", "kodegudang,keterangan", "kodegudang", sisAssign, cGudang.Text)
  If Not dbData.EOF Then
    cNamaGudang.Text = GetNull(dbData!keterangan)
  Else
    cNamaGudang.Default
  End If
End Sub

Private Sub nDisc1_Validate(Cancel As Boolean)
  SumJumlah
End Sub

Private Sub nDisc2_Validate(Cancel As Boolean)
  SumJumlah
End Sub

Private Sub nDiscount_Validate(Cancel As Boolean)
  SumTotal
End Sub

Private Sub nDiscount2_Validate(Cancel As Boolean)
  SumTotal
End Sub

Private Sub InitValue1()
  nNomor.value = 1
  cBarcode.Default
  cNama.Default
  nQty.value = 1
  cSatuan.Default
  nHarga.value = 0
  nDisc1.value = aCfg(objData, msDiscountItemPembelian, 0)
  nJumlah.value = 0
  cKode = ""
End Sub

Private Sub nBiaya_Validate(Cancel As Boolean)
  SumTotal
End Sub

Private Sub SumJumlah()
Dim nSubJumlah As Double

  nSubJumlah = nHarga.value * nQty.value
  nSubJumlah = nSubJumlah - (nSubJumlah * (nDisc1.value / 100))
  nJumlah.value = nSubJumlah
End Sub

Private Sub nHarga_Validate(Cancel As Boolean)
  SumJumlah
End Sub

Private Sub nNomor_Validate(Cancel As Boolean)
Dim n As Single
  
  If GetValidNomorUrut(nNomor, vaArray) Then
    n = nNomor.value - 1
    If n <= vaArray.UpperBound(1) Then
      cBarcode.Text = vaArray(n, 1)
      cNama.Text = vaArray(n, 2)
      nQty.value = vaArray(n, 3)
      cSatuan.Text = vaArray(n, 4)
      nHarga.value = vaArray(n, 5)
      nDisc1.value = vaArray(n, 6)
      nJumlah.value = vaArray(n, 7)
      cKode = vaArray(n, 8)
      cID = vaArray(n, 9)
    End If
  End If
End Sub

Private Sub nPersDisc_Validate(Cancel As Boolean)
  SumTotal
End Sub

Private Sub nPersDisc2_Validate(Cancel As Boolean)
  SumTotal
End Sub

Private Sub nPPn_Validate(Cancel As Boolean)
  SumTotal
End Sub

Private Sub nQty_Validate(Cancel As Boolean)
  SumJumlah
End Sub

Private Sub nTunai_Validate(Cancel As Boolean)
  SumBayar
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim n As Integer
Dim nQtyTmp As Integer

  If KeyCode = vbKeyDelete Then
    If vaArray.UpperBound(1) >= 0 Then
      vaDelete.InsertRows vaDelete.UpperBound(1) + 1
      n = vaDelete.UpperBound(1)
      vaDelete(n, 0) = TDBGrid1.Columns(1).Text
      vaDelete(n, 1) = TDBGrid1.Columns(9).Text
      
      TDBGrid1.Delete
      SumTotal
      For n = 0 To vaArray.UpperBound(1)
        vaArray(n, 0) = n + 1
        nQtyTmp = nQtyTmp + vaArray(n, 3)
      Next
      TDBGrid1.Columns(3).FooterText = Format(nQtyTmp, "###,###,##0")
      nNomor.value = vaArray.UpperBound(1) + 2
      TDBGrid1.ReBind
    End If
  End If
End Sub
