VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Begin VB.Form cfgInformasiPerusahaan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INFORMASI PERUSAHAAN"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   8850
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   4020
      Left            =   75
      Top             =   75
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   7091
      Caption         =   "Info Perusahaan"
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      Begin BiSATextBoxProject.BiSATextBox cNama 
         Height          =   330
         Left            =   630
         TabIndex        =   0
         Top             =   360
         Width           =   5865
         _ExtentX        =   10345
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
         MaxLength       =   50
         Appearance      =   0
         GetPicture      =   1
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
      Begin BiSATextBoxProject.BiSATextBox cAlamat 
         Height          =   336
         Left            =   636
         TabIndex        =   1
         Top             =   696
         Width           =   5868
         _ExtentX        =   10372
         _ExtentY        =   609
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
         Height          =   336
         Left            =   636
         TabIndex        =   2
         Top             =   1032
         Width           =   5868
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
         MaxLength       =   50
         Appearance      =   0
         Caption         =   "Telepon"
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
      Begin BiSATextBoxProject.BiSATextBox cFax 
         Height          =   336
         Left            =   636
         TabIndex        =   3
         Top             =   1368
         Width           =   5868
         _ExtentX        =   10372
         _ExtentY        =   609
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
         Caption         =   "Fax"
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
      Begin BiSATextBoxProject.BiSATextBox cEmail 
         Height          =   336
         Left            =   636
         TabIndex        =   4
         Top             =   1704
         Width           =   4392
         _ExtentX        =   7752
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
         MaxLength       =   30
         Appearance      =   0
         Caption         =   "E-Mail"
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
         Height          =   336
         Left            =   636
         TabIndex        =   5
         Top             =   2052
         Width           =   4392
         _ExtentX        =   7752
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
         MaxLength       =   30
         Appearance      =   0
         Caption         =   "Kota"
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
      Begin BiSATextBoxProject.BiSATextBox cProvinsi 
         Height          =   336
         Left            =   636
         TabIndex        =   6
         Top             =   2400
         Width           =   4392
         _ExtentX        =   7752
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
         MaxLength       =   30
         Appearance      =   0
         Caption         =   "Provinsi"
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
      Begin BiSATextBoxProject.BiSATextBox cOwner 
         Height          =   336
         Left            =   636
         TabIndex        =   7
         Top             =   2736
         Width           =   4392
         _ExtentX        =   7752
         _ExtentY        =   609
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
         MaxLength       =   30
         Appearance      =   0
         Caption         =   "Owner"
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
      Begin BiSATextBoxProject.BiSATextBox cTagline 
         Height          =   336
         Left            =   636
         TabIndex        =   10
         Top             =   3072
         Width           =   6876
         _ExtentX        =   12144
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
         MaxLength       =   30
         Appearance      =   0
         Caption         =   "Tagline"
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
      Begin BiSATextBoxProject.BiSATextBox cFakturPrefix 
         Height          =   336
         Left            =   636
         TabIndex        =   11
         Top             =   3408
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   609
         Text            =   "222"
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
         MaxLength       =   3
         Appearance      =   0
         Caption         =   "Faktur Pre"
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
      Height          =   630
      Left            =   75
      Top             =   4125
      Width           =   8685
      _ExtentX        =   15319
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
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   7515
         TabIndex        =   8
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
         Picture         =   "cfgInformasiPerusahaan.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   6435
         TabIndex        =   9
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
         Picture         =   "cfgInformasiPerusahaan.frx":00A6
      End
   End
End
Attribute VB_Name = "cfgInformasiPerusahaan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objData As New CodeSuiteLibrary.Data

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdReg_Click()
'  Dim tmp
'  tmp = EncDec(txbEncryptedKeyCode.Text, txbKeyCode.Text)
'
'  If tmp <> txbKeyCode.Text Then
'      MsgBox "Encrypted file has been tampered with; not pass"
'  Else
'      If txbTestPassword.Text <> txbKeyCode.Text Then
'          MsgBox "Incorrect password, not pass"
'          Exit Sub
'      ElseIf Len(txbTestPassword.Text) <> keyCodeLen Then
'          MsgBox "Incorrect password, not pass"
'          Exit Sub
'      End If
'  End If
'  MsgBox "Terimakasih sudah melakukan registrasi" & "Selamat menggunakan"
'  UpdCfg msReg, tmp
End Sub

Private Sub cmdSimpan_Click()
  UpdCfg msNamaPerusahaan, cNama.Text, objData, cNama.Caption, Me.Caption
  UpdCfg msAlamatPerusahaan, cAlamat.Text, objData, cAlamat.Caption, Me.Caption
  UpdCfg msTelepon, cTelepon.Text, objData, cTelepon.Caption, Me.Caption
  UpdCfg msFax, cFax.Text, objData, cFax.Caption, Me.Caption
  UpdCfg msEmail, cEmail.Text, objData, cEmail.Caption, Me.Caption
  UpdCfg msKota, cKota.Text, objData, cKota.Caption, Me.Caption
  UpdCfg msProvinsi, cProvinsi.Text, objData, cProvinsi.Caption, Me.Caption
  UpdCfg msNama, cOwner.Text, objData, cNama.Caption, Me.Caption
  UpdCfg msTaglinePerusahaan, cTagline.Text, objData, cTagline.Text, Me.Caption
  UpdCfg msFakturPrefix, cFakturPrefix.Text, objData, cFakturPrefix.Text, Me.Caption
  Unload Me
End Sub

Private Sub Form_Load()
Dim n As Single
  
  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  TabIndex cNama, n
  TabIndex cAlamat, n
  TabIndex cTelepon, n
  TabIndex cFax, n
  TabIndex cEmail, n
  TabIndex cKota, n
  TabIndex cProvinsi, n
  TabIndex cOwner, n
  TabIndex cTagline, n
  TabIndex cFakturPrefix, n

  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  
  cNama.Text = aCfg(objData, msNamaPerusahaan, "")
  cAlamat.Text = aCfg(objData, msAlamatPerusahaan, "")
  cTelepon.Text = aCfg(objData, msTelepon, "")
  cFax.Text = aCfg(objData, msFax, "")
  cEmail.Text = aCfg(objData, msEmail, "")
  cKota.Text = aCfg(objData, msKota, "")
  cProvinsi.Text = aCfg(objData, msProvinsi, "")
  cOwner.Text = aCfg(objData, msNama)
  cTagline.Text = aCfg(objData, msTaglinePerusahaan)
  cFakturPrefix.Text = aCfg(objData, msFakturPrefix)
'  Text1.Text = EncDec(cNamaKomputer, Text1.Text)
End Sub
