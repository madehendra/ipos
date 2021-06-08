VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trLunasPiutang 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Metode Pembayaran"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   9315
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton opt 
      Caption         =   "WITHDRAW/PENARIKAN"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   330
      TabIndex        =   22
      Top             =   4740
      Width           =   2895
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   585
      Left            =   15
      Top             =   7350
      Width           =   9285
      _ExtentX        =   16378
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
         Left            =   8145
         TabIndex        =   18
         Top             =   75
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   767
         Caption         =   "     &Cancel"
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
         Picture         =   "trLunasPiutang.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   7065
         TabIndex        =   19
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
         Picture         =   "trLunasPiutang.frx":00A6
      End
   End
   Begin BiSATextBoxProject.BiSATextBox cNamaAkunKas 
      Height          =   330
      Left            =   5760
      TabIndex        =   17
      Top             =   1950
      Width           =   2640
      _ExtentX        =   4657
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
   Begin VB.OptionButton opt 
      Caption         =   "BG/CEK/SURAT PENGAKUAN HUTANG"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   285
      TabIndex        =   10
      Top             =   2520
      Width           =   4035
   End
   Begin VB.OptionButton opt 
      Caption         =   "TUNAI/TRANSFER"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   285
      TabIndex        =   9
      Top             =   2025
      Width           =   2205
   End
   Begin BiSANumberBoxProject.BiSANumberBox nBG 
      Height          =   330
      Index           =   0
      Left            =   4455
      TabIndex        =   2
      Top             =   3090
      Width           =   2415
      _ExtentX        =   4260
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
   Begin BiSATextBoxProject.BiSATextBox cNoBG 
      Height          =   330
      Index           =   0
      Left            =   990
      TabIndex        =   1
      Top             =   3090
      Width           =   3450
      _ExtentX        =   6085
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
      Caption         =   "Reff No"
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
   Begin BiSATextBoxProject.BiSABrowse cAkunKas 
      Height          =   330
      Left            =   2520
      TabIndex        =   0
      Top             =   1950
      Width           =   3225
      _ExtentX        =   5689
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
      Caption         =   "Akun Kas"
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
   Begin BiSATextBoxProject.BiSATextBox cNoBG 
      Height          =   330
      Index           =   1
      Left            =   990
      TabIndex        =   3
      Top             =   3435
      Width           =   3450
      _ExtentX        =   6085
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
      Caption         =   "Reff No"
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
   Begin BiSANumberBoxProject.BiSANumberBox nBG 
      Height          =   330
      Index           =   1
      Left            =   4455
      TabIndex        =   4
      Top             =   3435
      Width           =   2415
      _ExtentX        =   4260
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
   Begin BiSATextBoxProject.BiSATextBox cNoBG 
      Height          =   330
      Index           =   2
      Left            =   990
      TabIndex        =   5
      Top             =   3780
      Width           =   3450
      _ExtentX        =   6085
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
      Caption         =   "Reff No"
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
   Begin BiSANumberBoxProject.BiSANumberBox nBG 
      Height          =   330
      Index           =   2
      Left            =   4455
      TabIndex        =   6
      Top             =   3780
      Width           =   2415
      _ExtentX        =   4260
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
   Begin BiSANumberBoxProject.BiSANumberBox nTotalBG 
      Height          =   330
      Left            =   4455
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4155
      Width           =   2415
      _ExtentX        =   4260
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
   Begin BiSANumberBoxProject.BiSANumberBox nTotalYangHarusDibayar 
      Height          =   330
      Left            =   270
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   405
      Width           =   5445
      _ExtentX        =   9604
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
      Caption         =   "Yg harus dibayar: "
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
   Begin BiSADateProject.BiSADate dJthTempo 
      Height          =   330
      Index           =   0
      Left            =   6885
      TabIndex        =   12
      Top             =   3090
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   582
      Value           =   "13-10-2005"
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
   Begin BiSADateProject.BiSADate dJthTempo 
      Height          =   330
      Index           =   1
      Left            =   6885
      TabIndex        =   13
      Top             =   3435
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   582
      Value           =   "13-10-2005"
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
   Begin BiSADateProject.BiSADate dJthTempo 
      Height          =   330
      Index           =   2
      Left            =   6885
      TabIndex        =   14
      Top             =   3780
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   582
      Value           =   "13-10-2005"
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
   Begin BiSANumberBoxProject.BiSANumberBox nTunai 
      Height          =   330
      Left            =   510
      TabIndex        =   20
      Top             =   825
      Width           =   5190
      _ExtentX        =   9155
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
      Caption         =   "Tunai"
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
   Begin BiSANumberBoxProject.BiSANumberBox nKembalian 
      Height          =   600
      Left            =   510
      TabIndex        =   21
      Top             =   1215
      Width           =   7905
      _ExtentX        =   13944
      _ExtentY        =   1058
      Appearance      =   0
      Enabled         =   0   'False
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Kembali"
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
   Begin BiSANumberBoxProject.BiSANumberBox nWithDraw 
      Height          =   330
      Left            =   1095
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1845
      _ExtentX        =   3254
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
   Begin BiSATextBoxProject.BiSATextBox cKodeAnggota 
      Height          =   330
      Left            =   3045
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   5040
      Width           =   2640
      _ExtentX        =   4657
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
   Begin BiSATextBoxProject.BiSATextBox cNamaAnggota 
      Height          =   330
      Left            =   5715
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   5040
      Width           =   3405
      _ExtentX        =   6006
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
   Begin BiSANumberBoxProject.BiSANumberBox nSaldoTopUp 
      Height          =   330
      Left            =   1605
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   5415
      Width           =   4095
      _ExtentX        =   7223
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
      Caption         =   "Saldo Top Up"
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
   Begin BiSANumberBoxProject.BiSANumberBox nTarikTunai 
      Height          =   345
      Left            =   1605
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   6165
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   609
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
      Caption         =   "Tarik Tunai"
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
   Begin BiSANumberBoxProject.BiSANumberBox nSisaKurang 
      Height          =   330
      Left            =   1605
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   6555
      Width           =   4095
      _ExtentX        =   7223
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
      Caption         =   "Sisa Kurang"
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
   Begin BiSANumberBoxProject.BiSANumberBox nTunaiSisa 
      Height          =   330
      Left            =   1605
      TabIndex        =   29
      Top             =   6945
      Width           =   4095
      _ExtentX        =   7223
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
      Caption         =   "Tunai"
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
   Begin BiSANumberBoxProject.BiSANumberBox nKembaliSisa 
      Height          =   330
      Left            =   5715
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   6945
      Width           =   3375
      _ExtentX        =   5953
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
   Begin VB.Frame Frame1 
      Caption         =   "Ya/tidak (tarik tunai)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   5745
      TabIndex        =   32
      Top             =   5505
      Visible         =   0   'False
      Width           =   2850
      Begin VB.OptionButton optTarikTunai 
         Caption         =   "Tidak"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   735
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   270
         Width           =   795
      End
      Begin VB.OptionButton optTarikTunai 
         Caption         =   "Ya"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   270
         Width           =   600
      End
   End
   Begin BiSANumberBoxProject.BiSANumberBox nJaminan 
      Height          =   330
      Left            =   1605
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   5790
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   582
      Appearance      =   0
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
      Caption         =   "Jaminan"
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
   Begin VB.Label Label4 
      Caption         =   "Kembali"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5775
      TabIndex        =   31
      Top             =   6600
      Width           =   840
   End
   Begin VB.Label Label2 
      Caption         =   "JATUH TEMPO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   16
      Top             =   2820
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "JUMLAH"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4500
      TabIndex        =   15
      Top             =   2820
      Width           =   915
   End
   Begin VB.Label Label3 
      Caption         =   "NO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   105
      Width           =   2250
   End
End
Attribute VB_Name = "trLunasPiutang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub cAkunKas_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", "kodeakun,keterangan", "kodeakun", sisContent, cAkunKas.Text, " and jenis = 'D'")
  If Not dbData.EOF Then
    cAkunKas.Text = cAkunKas.Browse(dbData)
    cNamaAkunKas.Text = GetNull(dbData!keterangan)
    trPelunasanPiutang.cPubAkun = cAkunKas.Text
  End If
End Sub

Private Sub initvalue()
Dim n As Single

'  opt(0).Value = True
'  opt(1).Value = True
  
  Label3.Caption = ""
  cAkunKas.Default
  cNamaAkunKas.Default
  optTarikTunai(1).Value = True
  
  If GetRegistry(reg_UserLevel = 0) Then
    cAkunKas.Enabled = True
  Else
    If GetKunciAkunKas(objData) Then
      cAkunKas.Enabled = False
    End If
  End If
  
  trPelunasanPiutang.cPubAkun = GetAkunKas(objData, GetRegistry(reg_UserName))
  cAkunKas.Text = GetAkunKas(objData, GetRegistry(reg_UserName))
  
  nTotalYangHarusDibayar.Value = 0
  nTunai.Value = 0
  nKembalian.Value = 0
  For n = cNoBG.LBound To cNoBG.UBound
    cNoBG(n).Default
  Next n
  For n = nBG.LBound To nBG.UBound
    nBG(n).Default
  Next n
  nTotalBG.Default
End Sub

Private Function isValidSaving() As Boolean
isValidSaving = True

  If opt(1).Value = True Then
    If nTotalYangHarusDibayar.Value > nTotalBG.Value Then
      MsgBox "Nominal yang dibayarkan kurang, transaksi tidak bisa dilanjutkan"
      isValidSaving = False
      Exit Function
    End If
  End If
  
  If opt(0).Value = True Then
  
    If trPelunasanPiutang.cPubAkun = "" Then
      MsgBox "Rekening/akun kas belum terisi, transaksi tidak bisa dilanjutkan"
      isValidSaving = False
      Exit Function
    End If
    If nTunai.Value < nTotalYangHarusDibayar.Value Then
      MsgBox "Pembayaran Kurang"
      isValidSaving = False
      Exit Function
    End If
    
  If nJaminan.Value > GetSaldoTopUpMember(objData, cKodeAnggota.Text) Then
    MsgBox "Maaf, nilai jaminan tidak boleh lebih dari Saldo Top Up"
    nJaminan.Value = 0
  End If
  End If
  
  If opt(2).Value = True Then
    If nSisaKurang.Value > 0 Then
      If nTunaiSisa.Value < nSisaKurang.Value Then
        MsgBox "Maaf Pembayaran Kurang... Penyimpanan tidak bisa dilanjutkan"
        isValidSaving = False
        Exit Function
      End If
    End If
  End If
  
End Function

Private Sub cmdKeluar_Click()
  trPelunasanPiutang.lClose = True
  Me.Hide
End Sub

Private Sub cmdSimpan_Click()
Dim n As Integer

  If isValidSaving Then
    vaArray.ReDim 0, 2, 0, 2
    
'    'kirim seluruh parameter
'    For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
'      vaArray(n, 0) = cNoBG(n).Text
'      vaArray(n, 1) = nBG(n).Value
'      vaArray(n, 2) = Format(dJthTempo(n).Value, "yyyy-MM-dd")
'    Next n
'
'    Set trPelunasanPiutang.vaPubReff = vaArray
'    trPelunasanPiutang.nPubTotal = nTotalBG.Value
'    trPelunasanPiutang.lPubStatus = opt(0).Value
'    trPelunasanPiutang.cPubAkun = cAkunKas.Text
'    trPelunasanPiutang.nWithDraw = nSaldoTopUp.Value  'nWithDraw.Value
'    trPelunasanPiutang.nTarikTunai = nTarikTunai.Value
'    trPelunasanPiutang.nSisaKurangTopUp = nSisaKurang.Value
'    trPelunasanPiutang.nSaldoTopUp = nSaldoTopUp.Value
'    trPelunasanPiutang.nTunai = nTunaiSisa.Value
'    trPelunasanPiutang.nJaminan = nJaminan.Value
'    trPelunasanPiutang.nKembalian = nKembaliSisa.Value
'    trPelunasanPiutang.nTotYgHarusDibayar = nTotalYangHarusDibayar.Value
'
'    trPelunasanPiutang.lTarikTunai = False
'
''    If optTarikTunai(0).Value = True Then
''      trPelunasanPiutang.lTarikTunai = True
''    ElseIf optTarikTunai(1).Value = True Then
''      trPelunasanPiutang.lTarikTunai = False
''    End If
'
'    If opt(0).Value = True Then
'      trPelunasanPiutang.nMetodePembayaran = 0
'    ElseIf opt(1).Value = True Then
'      trPelunasanPiutang.nMetodePembayaran = 1
'    ElseIf opt(2).Value = True Then
'      trPelunasanPiutang.nMetodePembayaran = 2
'    End If
    
    trPelunasanPiutang.cPubAkun = cAkunKas.Text
    trPelunasanPiutang.lClose = False
    Me.Hide
    
  End If
End Sub

Private Sub Form_Activate()
Dim nSisaTopUp As Double

  nTunai_Change
  
  nWithDraw.Value = 0
  nSaldoTopUp.Value = 0
  nTarikTunai.Value = 0
  nSisaTopUp = 0
  nSisaKurang.Value = nTotalYangHarusDibayar.Value - nSaldoTopUp.Value
'  Frame1.Visible = False
  
  nJaminan.Value = 0
  
  Set dbData = objData.Browse(GetDSN, "membertopup m", "m.kodeanggota,a.nama,a.alamat,sum(m.debet-m.kredit) as saldo", "m.tgl", sisLTEqual, Format(Date, "yyyy-MM-dd"), "  and m.kodeanggota = '" & cKodeAnggota.Text & "' GROUP BY m.kodeanggota", , Array("left join anggota a on a.kodeanggota = m.kodeanggota"))
  If Not dbData.EOF Then
    If GetNull(dbData!saldo) >= nTotalYangHarusDibayar.Value Then
      nWithDraw.Value = nTotalYangHarusDibayar.Value
      nSaldoTopUp.Value = GetNull(dbData!saldo)
      
      nSisaTopUp = GetSaldoTopUpMember(objData, cKodeAnggota.Text) - GetSaldoPiutang(objData, cKodeAnggota.Text)  'nSaldoTopUp.Value - nWithDraw.Value
      
'      If nSisaTopUp > 0 Then
'        nTarikTunai.Value = nSisaTopUp
'      Else
'        If optTarikTunai(0).Value = True Then
'          MsgBox "Maaf, sisa topup tidak bisa ditarik karena masih ada nota yg belum lunas"
'          nTarikTunai.Value = 0
'          optTarikTunai(1).Value = True
'          optTarikTunai(0).Enabled = False
'        End If
'      End If
      
      nSisaKurang.Value = 0
'      Frame1.Visible = True
    Else
      'nWithDraw.Value = GetNull(dbData!saldo)
      nWithDraw.Value = nTotalYangHarusDibayar.Value
      nSaldoTopUp.Value = GetNull(dbData!saldo)
      nTarikTunai.Value = 0
      nSisaKurang.Value = nTotalYangHarusDibayar.Value - nSaldoTopUp.Value
'      Frame1.Visible = False
      Frame1.Enabled = False
      optTarikTunai(1).Value = True
    End If
  End If
  
  If nSaldoTopUp.Value > 0 Then
    opt(2).Value = True
  End If
  
End Sub

Private Sub Form_Load()
Dim n As Single
  
  CenterForm Me
  SetIcon Me.hwnd
'  opt(0).Value = True
  optTarikTunai(1).Value = True
  

  
  TabIndex nTunai, n
  TabIndex opt(0), n
  TabIndex cAkunKas, n
  TabIndex opt(1), n
  TabIndex cNoBG(0), n
  TabIndex nBG(0), n
  TabIndex dJthTempo(0), n
  TabIndex cNoBG(1), n
  TabIndex nBG(1), n
  TabIndex dJthTempo(1), n
  TabIndex cNoBG(2), n
  TabIndex nBG(2), n
  TabIndex dJthTempo(2), n
  
  TabIndex opt(2), n
  TabIndex nTunaiSisa, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  
  dJthTempo(0).Value = DateAdd("d", 7, Now)
  dJthTempo(1).Value = DateAdd("d", 7, Now)
  dJthTempo(2).Value = DateAdd("d", 7, Now)
   
   
  If GetRegistry(reg_UserLevel) = 0 Then
    cAkunKas.Enabled = True
  Else
    If GetKunciAkunKas(objData) Then
      cAkunKas.Enabled = False
    End If
  End If
  Dim db As New ADODB.Recordset
  
  cAkunKas.Text = GetAkunKas(objData, GetRegistry(reg_UserName))
  trPelunasanPiutang.cPubAkun = GetAkunKas(objData, GetRegistry(reg_UserName))
  Set db = objData.Browse(GetDSN, "akun", , "kodeakun", sisAssign, cAkunKas.Text)
  If Not db.EOF Then
    cNamaAkunKas.Text = GetNull(db!keterangan)
  End If
  
End Sub

Private Sub nBG_Change(Index As Integer)
Dim n As Single
Dim nTemp As Double
  
  For n = nBG.LBound To nBG.UBound
    nTemp = nTemp + nBG(n).Value
  Next n
  
  nTotalBG.Value = nTemp
End Sub

Private Sub nJaminan_Validate(Cancel As Boolean)
  
'  If optTarikTunai(1).Value = True Then
'    If nSaldoTopUp.Value - nTotalYangHarusDibayar.Value > 0 Then
'      nJaminan.Value = 0
'      nTarikTunai.Value = nSaldoTopUp.Value - nTotalYangHarusDibayar.Value
'    End If
'  End If
  
  If nJaminan.Value > GetSaldoTopUpMember(objData, cKodeAnggota.Text) Then
    MsgBox "Maaf, nilai jaminan tidak boleh lebih dari Saldo Top Up"
    nJaminan.Value = 0
  End If
  nSaldoTopUp.Value = GetSaldoTopUpMember(objData, cKodeAnggota.Text) '- nJaminan.Value
  
'  If nSaldoTopUp.Value - nJaminan.Value - nWithDraw.Value >= 0 Then
'    Frame1.Enabled = True
'    nSisaKurang.Value = 0
'    nTarikTunai.Value = nSaldoTopUp.Value - nJaminan.Value - nWithDraw.Value
'  Else
'    Frame1.Enabled = False
'    optTarikTunai(1).Value = True
'    nTarikTunai.Value = 0
'    nSisaKurang.Value = nTotalYangHarusDibayar.Value - (nSaldoTopUp.Value - nJaminan.Value)
'  End If
  PenjumlahanTK
End Sub

Private Sub PenjumlahanTK()
  'nSisaKurang.Value = nSaldoTopUp.Value - nJaminan.Value - nTotalYangHarusDibayar.Value - nTarikTunai.Value
   nSisaKurang.Value = nTotalYangHarusDibayar.Value - (nSaldoTopUp.Value - nJaminan.Value)
End Sub

Private Sub nTarikTunai_Click()
'  PenjumlahanTK
End Sub

Private Sub nTunai_Change()
  nKembalian.Value = nTunai.Value - nTotalYangHarusDibayar.Value
End Sub

Private Sub nTunaiSisa_Change()
  nKembaliSisa.Value = nTunaiSisa.Value - nSisaKurang.Value
End Sub

Private Sub opt_Click(Index As Integer)
  If Index = 2 And nSaldoTopUp.Value <= 0 Then
'    MsgBox "Tidak ada saldo top up", vbExclamation
    opt(0).Value = True
  End If
End Sub

Private Sub opt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeysA vbKeyTab, True
  End If
End Sub

Private Sub opt_Validate(Index As Integer, Cancel As Boolean)
  If Index = 2 And nSaldoTopUp.Value <= 0 Then
'    MsgBox "Tidak ada saldo top up", vbExclamation
    opt(0).Value = True
  End If
End Sub

Private Sub optTarikTunai_Click(Index As Integer)
'  If optTarikTunai(1).Value = True Then
'    If nSaldoTopUp.Value - nTotalYangHarusDibayar.Value > 0 Then
'
'      nJaminan.Value = 0
'      nTarikTunai.Value = nSaldoTopUp.Value - nTotalYangHarusDibayar.Value
'
'    End If
'  End If
End Sub


