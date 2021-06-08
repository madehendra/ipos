VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form mstCustomer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Member Account..."
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   14670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   14670
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   2790
      Left            =   90
      Top             =   105
      Width           =   14400
      _ExtentX        =   25400
      _ExtentY        =   4921
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
      Begin BiSADateProject.BiSADate dTgl 
         Height          =   330
         Left            =   7350
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   270
         Width           =   3060
         _ExtentX        =   5398
         _ExtentY        =   582
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
         Caption         =   "Tgl Registrasi"
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
      Begin VB.OptionButton optAnggota 
         Caption         =   "&Non Member"
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
         Left            =   2820
         TabIndex        =   13
         Top             =   1230
         Width           =   1440
      End
      Begin VB.OptionButton optAnggota 
         Caption         =   "&A Member"
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
         Left            =   1770
         TabIndex        =   12
         Top             =   1230
         Width           =   1110
      End
      Begin BiSANumberBoxProject.BiSANumberBox nPlafond 
         Height          =   330
         Left            =   7395
         TabIndex        =   9
         Top             =   1785
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   582
         Appearance      =   0
         Decimals        =   0
         xxxx            =   10000000
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Max Plafond"
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
      Begin BiSATextBoxProject.BiSABrowse cKode 
         Height          =   330
         Left            =   135
         TabIndex        =   0
         Top             =   75
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   582
         Text            =   "123456789012345"
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
         MaxLength       =   15
         Appearance      =   0
         GetPicture      =   1
         Caption         =   "Kode"
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
      Begin BiSATextBoxProject.BiSATextBox cNama 
         Height          =   330
         Left            =   135
         TabIndex        =   1
         Top             =   435
         Width           =   5475
         _ExtentX        =   9657
         _ExtentY        =   582
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
         MaxLength       =   40
         Appearance      =   0
         Caption         =   "Nama"
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
      Begin BiSATextBoxProject.BiSATextBox cAlamat 
         Height          =   330
         Left            =   135
         TabIndex        =   2
         Top             =   795
         Width           =   5475
         _ExtentX        =   9657
         _ExtentY        =   582
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
         MaxLength       =   50
         Appearance      =   0
         Caption         =   "Alamat"
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
      Begin BiSATextBoxProject.BiSABrowse cKodeDep 
         Height          =   330
         Left            =   7395
         TabIndex        =   15
         Top             =   1080
         Width           =   3105
         _ExtentX        =   5477
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
         Caption         =   "Department"
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
      Begin BiSATextBoxProject.BiSATextBox cNamaDepartment 
         Height          =   330
         Left            =   10515
         TabIndex        =   16
         Top             =   1080
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   582
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
      Begin BiSATextBoxProject.BiSATextBox cTelepon 
         Height          =   330
         Left            =   7395
         TabIndex        =   17
         Top             =   1440
         Width           =   5025
         _ExtentX        =   8864
         _ExtentY        =   582
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
         MaxLength       =   20
         Appearance      =   0
         Caption         =   "Telepon"
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
      Begin BiSATextBoxProject.BiSABrowse cKodeUpline 
         Height          =   330
         Left            =   135
         TabIndex        =   18
         Top             =   1530
         Width           =   3120
         _ExtentX        =   5503
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
         Caption         =   "Sponsor"
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
      Begin BiSATextBoxProject.BiSATextBox cNamaUpline 
         Height          =   330
         Left            =   3285
         TabIndex        =   19
         Top             =   1530
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   582
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
      Begin BiSANumberBoxProject.BiSANumberBox nLevel 
         Height          =   330
         Left            =   7395
         TabIndex        =   20
         Top             =   735
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   582
         Appearance      =   0
         Decimals        =   0
         MaxValue        =   10
         MinValue        =   1
         xxxx            =   1
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Level"
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
      Begin BiSANumberBoxProject.BiSANumberBox nJatuhTempoPembayaran 
         Height          =   330
         Left            =   7395
         TabIndex        =   22
         Top             =   2160
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   582
         Appearance      =   0
         Decimals        =   0
         xxxx            =   30
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Jatuh Tempo"
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
      Begin BiSANumberBoxProject.BiSANumberBox nDiskonMember 
         Height          =   330
         Left            =   10395
         TabIndex        =   24
         Top             =   2160
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   582
         Appearance      =   0
         Decimals        =   0
         xxxx            =   30
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Diskon Item"
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
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   6300
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin BiSADateProject.BiSADate BiSADate1 
         Height          =   330
         Left            =   150
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1905
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   582
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
         Caption         =   "Tgl Lahir"
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
      Begin VB.Label Label3 
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
         Left            =   12780
         TabIndex        =   25
         Top             =   2175
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Hari"
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
         Left            =   9810
         TabIndex        =   23
         Top             =   2205
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Status"
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
         Left            =   195
         TabIndex        =   11
         Top             =   1215
         Width           =   735
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   90
      Top             =   6735
      Width           =   14400
      _ExtentX        =   25400
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
      Begin BiSAButtonProject.BiSAButton cmdImportFromExcel 
         Height          =   435
         Left            =   7005
         TabIndex        =   27
         Top             =   105
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   767
         Caption         =   "Import From Excel"
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
      Begin BiSAButtonProject.BiSAButton BiSAButton1 
         Height          =   435
         Left            =   3825
         TabIndex        =   26
         Top             =   105
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   767
         Caption         =   "Perbaiki Text Nama dan Alamat "
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
      Begin BiSAButtonProject.BiSAButton cmdExport 
         Height          =   435
         Left            =   8865
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   105
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   767
         Caption         =   "Export To Excel"
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
      Begin BiSAButtonProject.BiSAButton cmdHapus 
         Height          =   435
         Left            =   2220
         TabIndex        =   3
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
         Picture         =   "mstCustomer.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   11655
         TabIndex        =   4
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
         Picture         =   "mstCustomer.frx":028A
      End
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   1170
         TabIndex        =   5
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
         Picture         =   "mstCustomer.frx":0429
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   435
         Left            =   105
         TabIndex        =   6
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
         Picture         =   "mstCustomer.frx":0555
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   13215
         TabIndex        =   7
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
         Picture         =   "mstCustomer.frx":0700
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   12120
         TabIndex        =   8
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
         Picture         =   "mstCustomer.frx":07A6
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   3840
      Left            =   105
      TabIndex        =   10
      Top             =   2895
      Width           =   14370
      _ExtentX        =   25347
      _ExtentY        =   6773
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "NO"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "KODE"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "DEPT"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "NAMA"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "ALAMAT"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "UPLINE"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "NAMA UPLINE"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "LEVEL"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Telp"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   9
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   873
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   15790320
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=9"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1032"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=953"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=2672"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2593"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=512"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=2752"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2672"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=512"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=8149"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=8070"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=512"
      Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(25)=   "Column(4).Width=6853"
      Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=6773"
      Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=516"
      Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(31)=   "Column(5).Width=1270"
      Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=1191"
      Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=516"
      Splits(0)._ColumnProps(36)=   "Column(5).Visible=0"
      Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(38)=   "Column(6).Width=1984"
      Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=1905"
      Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=516"
      Splits(0)._ColumnProps(43)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(44)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(45)=   "Column(7).Width=1746"
      Splits(0)._ColumnProps(46)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(47)=   "Column(7)._WidthInPix=1667"
      Splits(0)._ColumnProps(48)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(49)=   "Column(7)._ColStyle=516"
      Splits(0)._ColumnProps(50)=   "Column(7).Visible=0"
      Splits(0)._ColumnProps(51)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(52)=   "Column(8).Width=2725"
      Splits(0)._ColumnProps(53)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(54)=   "Column(8)._WidthInPix=2646"
      Splits(0)._ColumnProps(55)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(56)=   "Column(8)._ColStyle=516"
      Splits(0)._ColumnProps(57)=   "Column(8).Order=9"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      RowDividerStyle =   0
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   15790320
      RowDividerColor =   15790320
      RowSubDividerColor=   15790320
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2"
      _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=58,.parent=13"
      _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=14"
      _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=15"
      _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=17"
      _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.alignment=0"
      _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=54,.parent=13,.alignment=0"
      _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=32,.parent=13,.alignment=0"
      _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=46,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(5).Style:id=50,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(6).Style:id=66,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(6).HeadingStyle:id=63,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(6).FooterStyle:id=64,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(6).EditorStyle:id=65,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(7).Style:id=62,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
      _StyleDefs(68)  =   "Named:id=33:Normal"
      _StyleDefs(69)  =   ":id=33,.parent=0"
      _StyleDefs(70)  =   "Named:id=34:Heading"
      _StyleDefs(71)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(72)  =   ":id=34,.wraptext=-1"
      _StyleDefs(73)  =   "Named:id=35:Footing"
      _StyleDefs(74)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(75)  =   "Named:id=36:Selected"
      _StyleDefs(76)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(77)  =   "Named:id=37:Caption"
      _StyleDefs(78)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(79)  =   "Named:id=38:HighlightRow"
      _StyleDefs(80)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(81)  =   "Named:id=39:EvenRow"
      _StyleDefs(82)  =   ":id=39,.parent=33,.bgcolor=&HE8E8E8&"
      _StyleDefs(83)  =   "Named:id=40:OddRow"
      _StyleDefs(84)  =   ":id=40,.parent=33"
      _StyleDefs(85)  =   "Named:id=41:RecordSelector"
      _StyleDefs(86)  =   ":id=41,.parent=34"
      _StyleDefs(87)  =   "Named:id=42:FilterBar"
      _StyleDefs(88)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "mstCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lClick As Boolean
Dim dbData As New ADODB.Recordset
Dim dbSupplier As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim lEdit As Boolean
Dim nPos As SisPos
Dim Validasi As Variant
Dim vaArray As New XArrayDB

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

Private Sub HapusData()
Dim cInfo As String
Dim lSave As Boolean
  cInfo = "Kode: " & cKode.Text & vbCrLf
  cInfo = cInfo & "Nama: " & cNama.Text & vbCrLf
  cInfo = cInfo & "Alamat: " & cAlamat.Text & vbCrLf

  lSave = True
  If MsgBox("Data Benar-benar dihapus ?" & vbCrLf & vbCrLf & cInfo, vbQuestion + vbYesNo) = vbYes Then
    objData.Start GetDSN
    If lExist(objData, "totpenjualan", "kodeanggota", GetKode) Then
      MsgBox "Maaf, data ini masih digunakan oleh sistem" & vbCrLf & "Tidak bisa dihapus"
      InitDel
      Exit Sub
    End If
    
    If lExist(objData, "totrtnpenjualan", "kodeanggota", GetKode) Then
      MsgBox "Maaf, data ini masih digunakan oleh sistem" & vbCrLf & "Tidak bisa dihapus"
      InitDel
      Exit Sub
    End If

    lSave = IIf(lSave, objData.Delete(GetDSN, "anggota", "kodeanggota", sisAssign, GetKode), False)
    If lSave Then
      objData.Save GetDSN
    Else
      objData.Cancel GetDSN
    End If
  End If
  InitDel
End Sub

Private Sub InitDel()
  initvalue
  GetLoadRows
  GetEdit False
End Sub

Private Function GetKode() As String
  GetKode = cKode.Text
End Function

Private Sub BiSAButton1_Click()
  GetPerbaikiNama
End Sub

Private Sub cKode_Validate(Cancel As Boolean)
  Set dbData = objData.Browse(GetDSN, "anggota", , "kodeanggota", sisAssign, GetKode)
  If dbData.RecordCount > 0 Then
    GetMemory
    If nPos = Delete Then
      HapusData
    End If
  End If
End Sub

Private Sub cKodeDep_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "dep", "kodedep,keterangan", , , , , "kodedep")
  If Not dbData.EOF Then
    cKodeDep.Text = cKodeDep.Browse(dbData, Array("Kode Dept", "Nama"), , Array(15, 25))
    cNamaDepartment.Text = GetNull(dbData!keterangan, "")
  End If
End Sub

Private Sub cKodeUpline_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "anggota", "kodeanggota,nama", "nama", sisContent, cKodeUpline.Text)
  If Not dbData.EOF Then
    cKodeUpline.Text = cKodeUpline.Browse(dbData)
    cKodeUpline.Text = GetNull(dbData!kodeanggota)
    cNamaUpline.Text = GetNull(dbData!nama)
  End If
End Sub

Private Sub cmdAktivasi_Click()
  frmAktivasi.Action Me
End Sub

Private Sub cmdAdd_Click()
  GetEdit True
  initvalue
  cKode.SetFocus
  nPos = Add
  cKode.Text = GetNomorAnggota
End Sub

Private Function GetNomorAnggota() As String
Dim cSQL As String
  
  cSQL = ""
  cSQL = cSQL & " select max(right(kodeanggota,4))+1 as nomoranggota from anggota"
  cSQL = cSQL & " Where length(kodeanggota) > 4 And Right(kodeanggota, 4) * 1 >= 1"

GetNomorAnggota = ""

  Set dbData = objData.Browse(GetDSN, "anggota", "max(kodeanggota)+1 as nomoranggota")
  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.EOF Then
    GetNomorAnggota = Padl(GetNull(dbData!nomoranggota), 6, "0")
  End If
End Function

Private Sub GetEdit(lPar As Boolean)
  BiSAFrame1.Enabled = lPar
  lEdit = lPar
  SetButton cmdSimpan, cmdKeluar, cmdAdd, cmdEdit, cmdHapus, nPos, lPar
End Sub

Private Sub cmdEdit_Click()
  GetEdit True
  cKode.SetFocus
  nPos = Edit
End Sub

Private Sub cmdExport_Click()
Dim a As New exportExcel
    
    CommonDialog1.Filter = "Excel File (*.xls)|*.xls"
    CommonDialog1.ShowSave
    If Trim(CommonDialog1.FileName) <> "" Then
      a.RecordSource = vaArray
      a.ExportToExcel , , , , CommonDialog1.FileName
      Set a = Nothing
      MsgBox "Export to Excel Berhasil", vbInformation
    End If
End Sub

Private Sub cmdHapus_Click()
  nPos = Delete
  GetEdit True
  cKode.SetFocus
  HapusData
End Sub

Private Sub cmdImportFromExcel_Click()
Dim lSave As Boolean
Dim vaField, vaValue
Dim i, j As Integer


  CommonDialog1.Filter = "Excel File (*.xls)|*.xls"
  CommonDialog1.ShowOpen
  
  
  If Trim(CommonDialog1.FileName) <> "" Then
    StartExcel
    lSave = True
    
    objData.Start GetDSN
      
    Excel.Workbooks.Close
    Set ExcelWBk = Excel.Workbooks.Open(CommonDialog1.FileName)
    Set ExcelWS = ExcelWBk.Worksheets(1)
    
    MsgBox "Mohon bersabar, tunggu sampai indikator menunjukkan selesai"
    
    'MsgBox "Yakin akan menambahkan " & ExcelWS.Cells.SpecialCells(xlCellTypeLastCell).Row & " Item stock baru??"
    'Loop berikut untuk membaca nilai setiap baris
    'mulai dari baris pertama sampai ketiga
    'dan setiap baris terdiri dari 2 kolom
    
    FrmPB.InitPB 2000
    Dim cIKode, cINama, cIAlamat, cIStatus, cIKodeDep, cIKodeUpline, dITgl, cIHP
  
    For i = 2 To ExcelWS.Cells.SpecialCells(xlCellTypeLastCell).Row
      FrmPB.RunPB
     For j = 1 To 8
      With ExcelWS
        cIKode = .Cells(i, 1).value
        cINama = .Cells(i, 2).value
        cIAlamat = .Cells(i, 3).value
        cIStatus = .Cells(i, 4).value
        cIKodeDep = .Cells(i, 5).value
        cIKodeUpline = .Cells(i, 6).value
        dITgl = Format(.Cells(i, 7).value, "yyyy-MM-dd")
        cIHP = .Cells(i, 8).value
      End With
     Next j
     
      If Trim(cINama) = "" Then
        Exit For
      End If
      vaField = Array("kodeanggota", "nama", "alamat", "status", "kodedep", "kodeupline", "tgl", "telp")
      vaValue = Array(cIKode, StrConv(cINama, vbProperCase), cIAlamat, cIStatus, cIKodeDep, cIKodeUpline, dITgl, cIHP)
      lSave = IIf(lSave, objData.Update(GetDSN, "anggota", "kodeanggota = '" & cIKode & "'", vaField, vaValue), False)
  
    Next i
    If lSave Then
      objData.Save GetDSN
    Else
      objData.Cancel GetDSN
    End If
    FrmPB.EndPB
     
    CloseWorkSheet
    FinishExcel
  End If
  GetLoadRows
End Sub

Private Sub cmdKeluar_Click()
  If Not lEdit Then
    Unload Me
  Else
    initvalue
    GetEdit False
  End If
End Sub

Private Sub cmdSimpan_Click()
Dim vaField
Dim vaValue
Dim lSave As Boolean
lSave = True

  If ValidSaving() Then
    objData.Start GetDSN
    vaField = Array("kodeanggota", "nama", "alamat", "plafond", "status", "kodedep", "tgl", "telp", "kodeupline", "nlevel", "dd", "diskon")
    vaValue = Array(Trim(GetKode), StrConv(cNama.Text, vbProperCase), StrConv(cAlamat.Text, vbProperCase), nPlafond.value, GetOpt(optAnggota), cKodeDep.Text, Format(dTgl.value, "yyyy-MM-dd"), Trim(cTelepon.Text), cKodeUpline.Text, nLevel.value, nJatuhTempoPembayaran.value, nDiskonMember.value)
    lSave = IIf(lSave, objData.Update(GetDSN, "anggota", "kodeanggota = '" & GetKode & "'", vaField, vaValue), False)
    If lSave Then
      objData.Save GetDSN
    Else
      objData.Cancel GetDSN
    End If
    initvalue
    GetEdit False
    GetLoadRows
    TDBGrid1.Refresh
  End If
End Sub

Private Sub GetPerbaikiNama()
Dim lSave As Boolean

  lSave = True
  objData.Start GetDSN
  Set dbData = objData.Browse(GetDSN, "anggota", "kodeanggota,nama,alamat")
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.EOF
      FrmPB.RunPB
      lSave = IIf(lSave, objData.Edit(GetDSN, "anggota", "kodeanggota = '" & GetNull(dbData!kodeanggota) & "'", Array("nama", "alamat"), Array(StrConv(GetNull(dbData!nama), vbProperCase), StrConv(GetNull(dbData!alamat), vbProperCase))), False)
      dbData.MoveNext
    Loop
    FrmPB.EndPB
  End If
    
  If lSave Then
    objData.Save GetDSN
    MsgBox "OK, data sudah selesai diperbaiki", vbInformation
  Else
    objData.Cancel GetDSN
  End If
  
End Sub

Static Function ValidSaving() As Boolean
Dim db As New ADODB.Recordset

  ValidSaving = True
  
  If Not CheckData(cKode.Text, "Kode Harus Diisi, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    cKode.SetFocus
    Exit Function
  End If
  
'  If Not CheckData(cTelepon.Text, "Nomer Tel Harus diisi") Then
'    ValidSaving = False
'    cTelepon.SetFocus
'    Exit Function
'  End If
  
  If Not CheckData(cNama.Text, "Nama Harus Diisi, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    cNama.SetFocus
    Exit Function
  End If
  
  Set db = objData.Browse(GetDSN, "dep", , "kodedep", sisAssign, cKodeDep.Text)
  If db.EOF Then
    MsgBox "Kode Department tidak tertera dengan benar" & vbCrLf & "Data tidak bisa disimpan"
    ValidSaving = False
    Exit Function
  End If
  
  If Trim(Len(cNama.Text)) < 3 Then
    MsgBox "Maaf Nama Pelanggan harus lebih panjang minimal 3 karakter", vbInformation, "Tidak bisa disimpan"
    ValidSaving = False
    Exit Function
  End If
End Function

Private Sub initvalue()
  cKode.Default
  cNama.Default
  cAlamat.Default
  cNamaUpline.Default
  nPlafond.Default
  cKodeUpline.Default
  cKodeDep.Default
  cNamaDepartment.Default
  cTelepon.Default
  optAnggota(0).value = True
  dTgl.value = Date
  nLevel.value = 1
  nJatuhTempoPembayaran.value = 0
  nDiskonMember.value = 0
End Sub

Private Sub GetMemory()
  Set dbData = objData.Browse(GetDSN, "anggota a", "a.*", "a.kodeanggota", sisAssign, GetKode)
  If dbData.RecordCount > 0 Then
    cNama.Text = GetNull(dbData!nama)
    cAlamat.Text = GetNull(dbData!alamat)
    nPlafond.value = GetNull(dbData!plafond)
    cKodeDep.Text = GetNull(dbData!kodedep)
    cKodeUpline.Text = GetNull(dbData!kodeupline)
    SetOpt optAnggota, GetNull(dbData!Status)
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  initvalue
  GetEdit False
  TabIndex cKode, n
  TabIndex cNama, n
  TabIndex cAlamat, n
  TabIndex optAnggota(0), n
  TabIndex optAnggota(1), n
  TabIndex cKodeUpline, n
  TabIndex nLevel, n
  TabIndex cKodeDep, n
  TabIndex cTelepon, n
  TabIndex nPlafond, n
  TabIndex nJatuhTempoPembayaran, n
  TabIndex nDiskonMember, n
  
  TabIndex cmdAdd, n
  TabIndex cmdEdit, n
  TabIndex cmdHapus, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  TabIndex cmdAktivasi, n
  
  GetLoadRows
End Sub

Private Sub GetLoadRows()
Dim n As Integer
Dim db As New ADODB.Recordset

  vaArray.ReDim 0, -1, 0, 8
  Set dbData = objData.Browse(GetDSN, "anggota", "kodeanggota,nama,alamat,kodedep,kodeupline,nlevel,telp", "kodeanggota", sisContent, TDBGrid1.Columns(1).FilterText, " and nama LIKE '%" & TDBGrid1.Columns(3).FilterText & "%' AND alamat LIKE '%" & TDBGrid1.Columns(4).FilterText & "%' AND kodedep LIKE '%" & TDBGrid1.Columns(2).FilterText & "%' LIMIT 0,25")
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = n + 1
      vaArray(n, 1) = GetNull(dbData!kodeanggota)
      vaArray(n, 2) = GetNull(dbData!kodedep)
      vaArray(n, 3) = GetNull(dbData!nama)
      vaArray(n, 4) = GetNull(dbData!alamat)
      vaArray(n, 5) = GetNull(dbData!kodeupline)
      vaArray(n, 7) = GetNull(dbData!nLevel)
      vaArray(n, 8) = GetNull(dbData!telp)
      vaArray(n, 6) = ""
      Set db = objData.Browse(GetDSN, "anggota", , "kodeanggota", sisAssign, vaArray(n, 5))
      If Not db.EOF Then
        vaArray(n, 6) = GetNull(db!nama)
      End If
      
      dbData.MoveNext
    Loop
  End If
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
  TDBGrid1.Refresh
End Sub

Private Sub optAnggota_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    keybd_event VK_TAB, 0, 0, 0
  End If
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    GetLoadRows
  End If
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim db As New ADODB.Recordset
Dim dba As New ADODB.Recordset

  Set db = objData.Browse(GetDSN, "anggota a", "a.*,d.keterangan,up.nama as namaupline", "a.kodeanggota", sisAssign, TDBGrid1.Columns(1).Text, , , Array("left join dep d on d.kodedep = a.kodedep", "left join anggota up on up.kodeupline = up.kodeanggota"))
  If Not db.EOF Then
    cKode.Text = GetNull(db!kodeanggota)
    cNama.Text = GetNull(db!nama)
    cAlamat.Text = GetNull(db!alamat)
    nPlafond.value = GetNull(db!plafond)
    cKodeDep.Text = GetNull(db!kodedep)
    dTgl.value = GetNull(db!tgl)
    SetOpt optAnggota, GetNull(db!Status)
    cNamaDepartment.Text = GetNull(db!keterangan)
    cTelepon.Text = GetNull(db!telp)
    nJatuhTempoPembayaran.value = GetNull(db!dd)
    nDiskonMember.value = GetNull(db!diskon)
    cKodeUpline.Text = GetNull(db!kodeupline)
    nLevel.value = 1
    cNamaUpline.Default
    Set dba = objData.Browse(GetDSN, "anggota", "nama", "kodeanggota", sisAssign, GetNull(db!kodeupline))
    If Not dba.EOF Then
      cNamaUpline.Text = GetNull(dba!nama)
    End If
  End If
End Sub
