VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Begin VB.Form cfgPortPrinter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Port Printer Settings...."
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   9690
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   2445
      Left            =   15
      Top             =   3735
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   4313
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
      Begin BiSATextBoxProject.BiSATextBox cFooterKasir 
         Height          =   330
         Left            =   1530
         TabIndex        =   12
         Top             =   135
         Width           =   4425
         _ExtentX        =   7805
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
         MaxLength       =   40
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
      Begin BiSATextBoxProject.BiSATextBox cFooterKasir2 
         Height          =   330
         Left            =   1515
         TabIndex        =   13
         Top             =   540
         Width           =   4425
         _ExtentX        =   7805
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
         MaxLength       =   40
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
      Begin BiSATextBoxProject.BiSATextBox cFooterKasir3 
         Height          =   330
         Left            =   1500
         TabIndex        =   29
         Top             =   930
         Width           =   4425
         _ExtentX        =   7805
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
         MaxLength       =   40
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
      Begin BiSATextBoxProject.BiSATextBox cFooterKasir4 
         Height          =   330
         Left            =   1485
         TabIndex        =   30
         Top             =   1305
         Width           =   4425
         _ExtentX        =   7805
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
         MaxLength       =   40
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
      Begin BiSATextBoxProject.BiSATextBox cFooterKasir5 
         Height          =   330
         Left            =   1485
         TabIndex        =   31
         Top             =   1665
         Width           =   4425
         _ExtentX        =   7805
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
         MaxLength       =   40
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
      Begin VB.Label Label2 
         Caption         =   "Footer Note : "
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
         Left            =   255
         TabIndex        =   25
         Top             =   975
         Width           =   1005
      End
      Begin VB.Label Label7 
         Caption         =   "Footer Struk : "
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
         Left            =   255
         TabIndex        =   15
         Top             =   510
         Width           =   1005
      End
      Begin VB.Label Label5 
         Caption         =   "Header Struk :"
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
         Left            =   255
         TabIndex        =   14
         Top             =   165
         Width           =   1110
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   3675
      Left            =   15
      Top             =   45
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   6482
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
      Begin BiSAFramProject.BiSAFrame BiSAFrame4 
         Height          =   645
         Left            =   165
         Top             =   1515
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   1138
         Caption         =   "Cetak Label Customer"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   4
         BackColor       =   -2147483633
         Begin VB.OptionButton optCetakLabelCustomer 
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
            Height          =   195
            Index           =   0
            Left            =   225
            TabIndex        =   17
            Top             =   270
            Width           =   615
         End
         Begin VB.OptionButton optCetakLabelCustomer 
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
            Height          =   195
            Index           =   1
            Left            =   885
            TabIndex        =   16
            Top             =   270
            Width           =   855
         End
      End
      Begin BiSAButtonProject.BiSAButton cmdAuto 
         Height          =   345
         Left            =   7545
         TabIndex        =   8
         Top             =   375
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   609
         Caption         =   "Auto"
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
      Begin VB.OptionButton optPrintRataKiri 
         Caption         =   "&2 Rata Kanan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   195
         TabIndex        =   1
         Top             =   750
         Width           =   1515
      End
      Begin VB.OptionButton optPrintRataKiri 
         Caption         =   "&1 Rata Kiri"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   195
         TabIndex        =   0
         Top             =   405
         Width           =   1185
      End
      Begin BiSANumberBoxProject.BiSANumberBox nLebarKertas 
         Height          =   330
         Left            =   2685
         TabIndex        =   3
         Top             =   375
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   582
         Appearance      =   0
         Decimals        =   0
         MaxValue        =   80
         MinValue        =   0
         xxxx            =   80
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Lebar Kertas"
         CaptionWidth    =   1300
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin BiSANumberBoxProject.BiSANumberBox nMarginKiri 
         Height          =   330
         Left            =   2685
         TabIndex        =   4
         Top             =   750
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   582
         Appearance      =   0
         Decimals        =   0
         MaxValue        =   80
         MinValue        =   0
         xxxx            =   80
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Margin Kiri"
         CaptionWidth    =   1300
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin BiSANumberBoxProject.BiSANumberBox nLebarKolom1 
         Height          =   330
         Left            =   4920
         TabIndex        =   5
         Top             =   375
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   582
         Appearance      =   0
         Decimals        =   0
         MaxValue        =   80
         MinValue        =   0
         xxxx            =   80
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Kolom 1"
         CaptionWidth    =   1300
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin BiSANumberBoxProject.BiSANumberBox nLebarKolom2 
         Height          =   330
         Left            =   4920
         TabIndex        =   6
         Top             =   750
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   582
         Appearance      =   0
         Decimals        =   0
         MaxValue        =   80
         MinValue        =   0
         xxxx            =   80
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Kolom 2"
         CaptionWidth    =   1300
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin BiSANumberBoxProject.BiSANumberBox nLebarKolom3 
         Height          =   330
         Left            =   4920
         TabIndex        =   7
         Top             =   1125
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   582
         Appearance      =   0
         Decimals        =   0
         MaxValue        =   80
         MinValue        =   0
         xxxx            =   80
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Kolom 3"
         CaptionWidth    =   1300
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin BiSANumberBoxProject.BiSANumberBox nKolom1_2 
         Height          =   330
         Left            =   6945
         TabIndex        =   9
         Top             =   375
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   582
         Appearance      =   0
         Decimals        =   0
         MaxValue        =   80
         MinValue        =   0
         xxxx            =   80
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionWidth    =   1300
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
      Begin BiSANumberBoxProject.BiSANumberBox nKolom2_2 
         Height          =   330
         Left            =   6960
         TabIndex        =   10
         Top             =   750
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   582
         Appearance      =   0
         Decimals        =   0
         MaxValue        =   80
         MinValue        =   0
         xxxx            =   80
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionWidth    =   1300
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
      Begin BiSANumberBoxProject.BiSANumberBox nKolom3_2 
         Height          =   330
         Left            =   6960
         TabIndex        =   11
         Top             =   1125
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   582
         Appearance      =   0
         Decimals        =   0
         MaxValue        =   80
         MinValue        =   0
         xxxx            =   80
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionWidth    =   1300
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
      Begin BiSANumberBoxProject.BiSANumberBox nMarginBawah 
         Height          =   330
         Left            =   2685
         TabIndex        =   18
         Top             =   1125
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   582
         Appearance      =   0
         Decimals        =   0
         MaxValue        =   80
         MinValue        =   0
         xxxx            =   80
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Margin Bawah"
         CaptionWidth    =   1300
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin BiSAFramProject.BiSAFrame BiSAFrame5 
         Height          =   645
         Left            =   2325
         Top             =   1515
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   1138
         Caption         =   "Cetak Berulang"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   4
         BackColor       =   -2147483633
         Begin VB.OptionButton optCetakBerulang 
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
            Height          =   195
            Index           =   1
            Left            =   750
            TabIndex        =   20
            Top             =   270
            Width           =   870
         End
         Begin VB.OptionButton optCetakBerulang 
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
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   19
            Top             =   270
            Width           =   615
         End
      End
      Begin BiSAFramProject.BiSAFrame BiSAFrame1 
         Height          =   645
         Left            =   4500
         Top             =   1500
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   1138
         Caption         =   "Tampilkan Barcode"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   4
         BackColor       =   -2147483633
         Begin VB.OptionButton optTampilBarcode 
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
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   22
            Top             =   270
            Width           =   615
         End
         Begin VB.OptionButton optTampilBarcode 
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
            Height          =   195
            Index           =   1
            Left            =   750
            TabIndex        =   21
            Top             =   270
            Width           =   870
         End
      End
      Begin BiSATextBoxProject.BiSATextBox cPortStruk 
         Height          =   330
         Left            =   180
         TabIndex        =   23
         Top             =   2445
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   582
         Text            =   "1234567890"
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
         MaxLength       =   10
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
      Begin BiSAFramProject.BiSAFrame BiSAFrame6 
         Height          =   645
         Left            =   165
         Top             =   2865
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   1138
         Caption         =   "Open Cashdrawer"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   4
         BackColor       =   -2147483633
         Begin VB.OptionButton optCashDrawerOpen 
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
            Height          =   195
            Index           =   1
            Left            =   885
            TabIndex        =   28
            Top             =   270
            Width           =   855
         End
         Begin VB.OptionButton optCashDrawerOpen 
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
            Height          =   195
            Index           =   0
            Left            =   225
            TabIndex        =   27
            Top             =   270
            Width           =   615
         End
      End
      Begin VB.Label Label3 
         Caption         =   "*Jika dikosongi sudah defaullt di set LPT1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1920
         TabIndex        =   26
         Top             =   2490
         Width           =   3585
      End
      Begin VB.Label Label1 
         Caption         =   "PORT YG DIGUNAKAN"
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
         Left            =   195
         TabIndex        =   24
         Top             =   2205
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Khusus cetakan menggunakan Printer Thermal.  Silahkan dipilih, Rata Kiri/Rata Kanan dan tentukan lebar kolom cetakan"
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
         Left            =   180
         TabIndex        =   2
         Top             =   105
         Width           =   9195
      End
   End
End
Attribute VB_Name = "cfgPortPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objData As New CodeSuiteLibrary.Data

Private Sub cmdAuto_Click()
  nLebarKolom1.value = Round((nLebarKertas.value - nMarginKiri.value) * 19 / 100)
  nLebarKolom2.value = Round((nLebarKertas.value - nMarginKiri.value) * 37 / 100)
  nLebarKolom3.value = Round((nLebarKertas.value - nMarginKiri.value) * 44 / 100)
  nKolom1_2.value = Round((nLebarKertas.value - nMarginKiri.value) * 15 / 100)
  nKolom2_2.value = Round((nLebarKertas.value - nMarginKiri.value) * 26 / 100)
  nKolom3_2.value = Round((nLebarKertas.value - nMarginKiri.value) * 41 / 100)
End Sub

Private Sub Form_Load()
  CenterForm Me
  SetIcon Me.hWnd
  SetOpt optTampilBarcode, GetRegistry(reg_TampilkanBarcode)
  SetOpt optPrintRataKiri, GetRegistry(reg_AlignmentThermal)
  SetOpt optCetakLabelCustomer, GetRegistry(reg_CetakLabelCustomer)
  SetOpt optCetakBerulang, GetRegistry(reg_CetakBerulang)
  SetOpt optCashDrawerOpen, GetRegistry(reg_OpenCashDrawer)
  
  cPortStruk.Text = GetRegistry(reg_PortStruk)
  
  nLebarKertas.value = GetRegistry(reg_LebarKertas)
  nMarginKiri.value = GetRegistry(reg_MarginKiri)
  nMarginBawah.value = GetRegistry(reg_MarginBawah)
  nLebarKolom1.value = GetRegistry(reg_LebarKolom1)
  nLebarKolom2.value = GetRegistry(reg_LebarKolom2)
  nLebarKolom3.value = GetRegistry(reg_LebarKolom3)
  
  nKolom1_2.value = GetRegistry(reg_LebarKolom1_2)
  nKolom2_2.value = GetRegistry(reg_LebarKolom2_2)
  nKolom3_2.value = GetRegistry(reg_LebarKolom3_2)
  
  cFooterKasir.Text = aCfg(objData, msKasir1)
  cFooterKasir2.Text = aCfg(objData, msKasir2)
  cFooterKasir3.Text = aCfg(objData, msKasir3)
  cFooterKasir4.Text = aCfg(objData, msKasir4)
  cFooterKasir5.Text = aCfg(objData, msKasir5)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SaveRegistry reg_AlignmentThermal, GetOpt(optPrintRataKiri)
  SaveRegistry reg_CetakLabelCustomer, GetOpt(optCetakLabelCustomer)
  SaveRegistry reg_TampilkanBarcode, GetOpt(optTampilBarcode)
  SaveRegistry reg_CetakBerulang, GetOpt(optCetakBerulang)
  SaveRegistry reg_PortStruk, cPortStruk.Text
  SaveRegistry reg_OpenCashDrawer, GetOpt(optCashDrawerOpen)
   
  SaveRegistry reg_LebarKertas, nLebarKertas.value
  SaveRegistry reg_MarginKiri, nMarginKiri.value
  SaveRegistry reg_MarginBawah, nMarginBawah.value
  SaveRegistry reg_LebarKolom1, nLebarKolom1.value
  SaveRegistry reg_LebarKolom2, nLebarKolom2.value
  SaveRegistry reg_LebarKolom3, nLebarKolom3.value
  
  SaveRegistry reg_LebarKolom1_2, nKolom1_2.value
  SaveRegistry reg_LebarKolom2_2, nKolom2_2.value
  SaveRegistry reg_LebarKolom3_2, nKolom3_2.value
  
  SaveRegistry reg_LebarEfektif, nLebarKertas.value - nMarginKiri.value
  
  UpdCfg msKasir1, cFooterKasir.Text, objData, Label5, Me.Caption
  UpdCfg msKasir2, cFooterKasir2.Text, objData, Label5.Caption, Me.Caption
  UpdCfg msKasir3, cFooterKasir3.Text, objData, Label5.Caption, Me.Caption
  
  UpdCfg msKasir4, cFooterKasir4.Text, objData, Label5.Caption, Me.Caption
  UpdCfg msKasir5, cFooterKasir5.Text, objData, Label5.Caption, Me.Caption
  
  MsgBox "Data Berhasil Disimpan", vbInformation, "Sukses"
End Sub

