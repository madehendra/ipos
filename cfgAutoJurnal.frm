VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Begin VB.Form cfgAutoJurnal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Konfigurasi Auto Jurnal..."
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   10980
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   8445
      Left            =   0
      Top             =   15
      Width           =   10950
      _ExtentX        =   19315
      _ExtentY        =   14896
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
      Begin BiSATextBoxProject.BiSABrowse cDiscountPembelian 
         Height          =   330
         Left            =   135
         TabIndex        =   2
         Top             =   690
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
         Caption         =   "Discount Pembelian"
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekeningDiscountPembelian 
         Height          =   330
         Left            =   4650
         TabIndex        =   3
         Top             =   690
         Width           =   3915
         _ExtentX        =   6906
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
      Begin BiSATextBoxProject.BiSABrowse cPPnPembelian 
         Height          =   330
         Left            =   135
         TabIndex        =   4
         Top             =   1035
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
         Caption         =   "PPn Pembelian"
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekeningPPnPembelian 
         Height          =   330
         Left            =   4650
         TabIndex        =   5
         Top             =   1035
         Width           =   3915
         _ExtentX        =   6906
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
      Begin BiSATextBoxProject.BiSABrowse cRekeningPenjualan 
         Height          =   330
         Left            =   960
         TabIndex        =   6
         Top             =   1425
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
         Caption         =   "Penjualan"
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekeningPenjualan 
         Height          =   330
         Left            =   5475
         TabIndex        =   7
         Top             =   1410
         Width           =   3915
         _ExtentX        =   6906
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
      Begin BiSATextBoxProject.BiSABrowse cRekeningDiscountPenjualan 
         Height          =   330
         Left            =   960
         TabIndex        =   8
         Top             =   1770
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
         Caption         =   "Discount Penjualan"
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekeningDiscountPenjualan 
         Height          =   330
         Left            =   5475
         TabIndex        =   9
         Top             =   1770
         Width           =   3915
         _ExtentX        =   6906
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
      Begin BiSATextBoxProject.BiSABrowse cRekeningPPnPenjualan 
         Height          =   330
         Left            =   960
         TabIndex        =   10
         Top             =   2115
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
         Caption         =   "PPn Penjualan"
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekeningPPnPenjualan 
         Height          =   330
         Left            =   5475
         TabIndex        =   11
         Top             =   2115
         Width           =   3915
         _ExtentX        =   6906
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
      Begin BiSATextBoxProject.BiSABrowse cRekeningCOGS 
         Height          =   330
         Left            =   960
         TabIndex        =   12
         Top             =   2460
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
         Caption         =   "Harga Pokok Penjualan"
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekeningCOGS 
         Height          =   330
         Left            =   5475
         TabIndex        =   13
         Top             =   2460
         Width           =   3915
         _ExtentX        =   6906
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
      Begin BiSATextBoxProject.BiSABrowse cRekeningPotonganPiutang 
         Height          =   330
         Left            =   135
         TabIndex        =   14
         Top             =   2850
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
         Caption         =   "Potongan Pelunasan Piutang"
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekeningPotonganPiutang 
         Height          =   330
         Left            =   4650
         TabIndex        =   15
         Top             =   2850
         Width           =   3915
         _ExtentX        =   6906
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
      Begin BiSATextBoxProject.BiSABrowse cRekeningPotonganHutang 
         Height          =   330
         Left            =   135
         TabIndex        =   16
         Top             =   3195
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
         Caption         =   "Potongan Pelunasan Hutang"
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekeningPotonganHutang 
         Height          =   330
         Left            =   4650
         TabIndex        =   17
         Top             =   3195
         Width           =   3915
         _ExtentX        =   6906
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
      Begin BiSATextBoxProject.BiSABrowse cRekeningPiutangDagang 
         Height          =   330
         Left            =   960
         TabIndex        =   19
         Top             =   3585
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
         Caption         =   "Piutang Dagang"
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekeningPiutangDagang 
         Height          =   330
         Left            =   5475
         TabIndex        =   20
         Top             =   3585
         Width           =   3915
         _ExtentX        =   6906
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
      Begin BiSATextBoxProject.BiSABrowse cRekeningHutangDagang 
         Height          =   330
         Left            =   960
         TabIndex        =   21
         Top             =   3945
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
         Caption         =   "Hutang Dagang"
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekeningHutangDagang 
         Height          =   330
         Left            =   5475
         TabIndex        =   22
         Top             =   3945
         Width           =   3915
         _ExtentX        =   6906
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
      Begin BiSATextBoxProject.BiSABrowse cRekeningHutangSalesman 
         Height          =   330
         Left            =   165
         TabIndex        =   23
         Top             =   4680
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
         Caption         =   "Hutang Salesman"
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekeningHutangSalesman 
         Height          =   330
         Left            =   4680
         TabIndex        =   24
         Top             =   4680
         Width           =   3915
         _ExtentX        =   6906
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
      Begin BiSATextBoxProject.BiSABrowse cRekeningBiayaKomisi 
         Height          =   330
         Left            =   165
         TabIndex        =   25
         Top             =   5025
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
         Caption         =   "Biaya Komisi Salesman"
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekeningBiayaKomisi 
         Height          =   330
         Left            =   4680
         TabIndex        =   26
         Top             =   5025
         Width           =   3915
         _ExtentX        =   6906
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
      Begin BiSATextBoxProject.BiSABrowse cRekeningBG 
         Height          =   330
         Left            =   165
         TabIndex        =   27
         Top             =   5505
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
         Caption         =   "Rekening BG"
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekeningBG 
         Height          =   330
         Left            =   4680
         TabIndex        =   28
         Top             =   5505
         Width           =   3915
         _ExtentX        =   6906
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
      Begin BiSATextBoxProject.BiSABrowse cRekeningToUp 
         Height          =   330
         Left            =   165
         TabIndex        =   29
         Top             =   5880
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
         Caption         =   "Rekening Top Up"
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekeningToUp 
         Height          =   330
         Left            =   4680
         TabIndex        =   30
         Top             =   5880
         Width           =   3915
         _ExtentX        =   6906
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
      Begin BiSATextBoxProject.BiSABrowse cRekeningAktivaTopUp 
         Height          =   330
         Left            =   165
         TabIndex        =   31
         Top             =   6255
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
         Caption         =   "Rekening Aktiva Top Up"
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekeningAktivaTopUp 
         Height          =   330
         Left            =   4680
         TabIndex        =   32
         Top             =   6255
         Width           =   3915
         _ExtentX        =   6906
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
      Begin BiSATextBoxProject.BiSABrowse cRekDefaultSetoranKas 
         Height          =   330
         Left            =   165
         TabIndex        =   33
         Top             =   6675
         Width           =   5400
         _ExtentX        =   9525
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
         Caption         =   "Default Rekening Setoran Kas"
         CaptionWidth    =   2800
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekDefaultSetoranKas 
         Height          =   330
         Left            =   5580
         TabIndex        =   34
         Top             =   6675
         Width           =   3915
         _ExtentX        =   6906
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
      Begin BiSATextBoxProject.BiSABrowse cRekeningPrive 
         Height          =   330
         Left            =   150
         TabIndex        =   35
         Top             =   7065
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
         Caption         =   "Rekening Prive"
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekeningPrive 
         Height          =   330
         Left            =   4680
         TabIndex        =   36
         Top             =   7065
         Width           =   3915
         _ExtentX        =   6906
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
      Begin BiSATextBoxProject.BiSABrowse cRekeningReturPembelian 
         Height          =   330
         Left            =   960
         TabIndex        =   37
         Top             =   4305
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
         Caption         =   "Retur Pembelian"
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekeningReturPembelian 
         Height          =   330
         Left            =   5475
         TabIndex        =   38
         Top             =   4305
         Width           =   3915
         _ExtentX        =   6906
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
      Begin BiSATextBoxProject.BiSABrowse cRekeningTitipanKasReturPembelian 
         Height          =   336
         Left            =   156
         TabIndex        =   39
         Top             =   7416
         Width           =   4500
         _ExtentX        =   7938
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
         Appearance      =   0
         Button          =   -1  'True
         Caption         =   "Titipan Kas Retur Beli"
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekeningTitipanKasReturPembelian 
         Height          =   336
         Left            =   4680
         TabIndex        =   40
         Top             =   7416
         Width           =   3912
         _ExtentX        =   6906
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
      Begin BiSATextBoxProject.BiSABrowse cRekeningPendapatanFeeKartu 
         Height          =   345
         Left            =   150
         TabIndex        =   41
         Top             =   7770
         Width           =   4500
         _ExtentX        =   7938
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
         Appearance      =   0
         Button          =   -1  'True
         Caption         =   "Pendapatan Fee Kartu"
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekeningPendapatanFeeKartu 
         Height          =   330
         Left            =   4680
         TabIndex        =   42
         Top             =   7770
         Width           =   3915
         _ExtentX        =   6906
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
      Begin VB.Label Label2 
         Caption         =   "Daftar Rekening Akuntansi untuk Pos sbb:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   192
         TabIndex        =   18
         Top             =   132
         Width           =   5460
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   645
      Left            =   15
      Top             =   8460
      Width           =   10950
      _ExtentX        =   19315
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
      BackColor       =   -2147483633
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   9765
         TabIndex        =   0
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
         Picture         =   "cfgAutoJurnal.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   8700
         TabIndex        =   1
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
         Picture         =   "cfgAutoJurnal.frx":00A6
      End
   End
End
Attribute VB_Name = "cfgAutoJurnal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data

Private Sub cDiscountPembelian_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", , "kodeakun", sisPrefix, "4", " AND jenis = 'D'")
  If Not dbData.EOF Then
    cDiscountPembelian.Text = cDiscountPembelian.Browse(dbData)
    cNamaRekeningDiscountPembelian.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdSimpan_Click()
  objData.Start GetDSN
  UpdCfg msRekeningDiscountPembelian, cDiscountPembelian.Text, objData, cDiscountPembelian.Caption, Me.Caption
  UpdCfg msRekeningPPnPembelian, cPPnPembelian.Text, objData, cPPnPembelian.Caption, Me.Caption
  UpdCfg msRekeningPenjualan, cRekeningPenjualan.Text, objData, cRekeningPenjualan.Caption, Me.Caption
  UpdCfg msRekeningDiscountPenjualan, cRekeningDiscountPenjualan.Text, objData, cRekeningDiscountPenjualan.Caption, Me.Caption
  UpdCfg msRekeningPPnPenjualan, cRekeningPPnPenjualan.Text, objData, cRekeningPPnPenjualan.Caption, Me.Caption
  UpdCfg msRekeningCOGS, cRekeningCOGS.Text, objData, cRekeningCOGS.Caption, Me.Caption
  UpdCfg msRekeningPotonganPiutang, cRekeningPotonganPiutang.Text, objData, cRekeningPotonganPiutang.Caption, Me.Caption
  UpdCfg msRekeningPotonganHutang, cRekeningPotonganHutang.Text, objData, cRekeningPotonganHutang.Caption, Me.Caption
  UpdCfg msPiutangDagang, cRekeningPiutangDagang.Text, objData, cRekeningPiutangDagang.Caption, Me.Caption
  UpdCfg msHutangDagang, cRekeningHutangDagang.Text, objData, cRekeningHutangDagang.Caption, Me.Caption
  
  UpdCfg msRekeningReturPembelian, cRekeningReturPembelian.Text, objData, cRekeningReturPembelian.Caption, Me.Caption
  
  UpdCfg msRekeningHutangSalesman, cRekeningHutangSalesman.Text, objData, cRekeningHutangSalesman.Caption, Me.Caption
  UpdCfg msRekeningBiayaKomisi, cRekeningBiayaKomisi.Text, objData, cRekeningBiayaKomisi.Caption, Me.Caption
  UpdCfg msRekeningBG, cRekeningBG.Text, objData, cRekeningBG.Caption, Me.Caption
  UpdCfg msRekeningTopUp, cRekeningToUp.Text, objData, cRekeningToUp.Caption, Me.Caption
  UpdCfg msRekeningKasTopUp, cRekeningAktivaTopUp.Text, objData, cRekeningAktivaTopUp.Caption, Me.Caption
  UpdCfg msRekeningSetoranKas, cRekDefaultSetoranKas.Text, objData, cRekDefaultSetoranKas.Caption, Me.Caption
  UpdCfg msRekeningPrive, cRekeningPrive.Text, objData, cRekeningPrive.Caption, Me.Caption
  UpdCfg msRekeningTitipanKasReturPembelian, cRekeningTitipanKasReturPembelian.Text, objData, cRekeningTitipanKasReturPembelian.Caption, Me.Caption
  
  UpdCfg msRekeningFeeKartu, cRekeningPendapatanFeeKartu.Text, objData, cRekeningPendapatanFeeKartu.Caption, Me.Caption
  
  objData.Save GetDSN
  MsgBox "Data telah tersimpan", vbInformation
End Sub

Private Sub cPPnPembelian_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", , "kodeakun", sisPrefix, "2", " AND jenis = 'D'")
  If Not dbData.EOF Then
    cPPnPembelian.Text = cPPnPembelian.Browse(dbData)
    cNamaRekeningPPnPembelian.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub cRekDefaultSetoranKas_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", , "kodeakun", sisPrefix, "1", " AND jenis = 'D'")
  If Not dbData.EOF Then
    cRekDefaultSetoranKas.Text = cRekDefaultSetoranKas.Browse(dbData)
    cNamaRekDefaultSetoranKas.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub cRekeningAktivaTopUp_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", , "kodeakun", sisPrefix, "1", " AND jenis = 'D'")
  If Not dbData.EOF Then
    cRekeningAktivaTopUp.Text = cRekeningAktivaTopUp.Browse(dbData)
    cNamaRekeningAktivaTopUp.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub cRekeningBG_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", , "kodeakun", sisPrefix, "1", " AND jenis = 'D'")
  If Not dbData.EOF Then
    cRekeningBG.Text = cRekeningBG.Browse(dbData)
    cNamaRekeningBG.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub cRekeningBiayaKomisi_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", , "kodeakun", sisPrefix, "5", " AND jenis = 'D'")
  If Not dbData.EOF Then
    cRekeningBiayaKomisi.Text = cRekeningBiayaKomisi.Browse(dbData)
    cNamaRekeningBiayaKomisi.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub cRekeningCOGS_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", , "kodeakun", sisPrefix, "5", " AND jenis = 'D'")
  If Not dbData.EOF Then
    cRekeningCOGS.Text = cRekeningCOGS.Browse(dbData)
    cNamaRekeningCOGS.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub cRekeningDiscountPenjualan_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", , "kodeakun", sisPrefix, "4", " AND jenis = 'D'")
  If Not dbData.EOF Then
    cRekeningDiscountPenjualan.Text = cRekeningDiscountPenjualan.Browse(dbData)
    cNamaRekeningDiscountPenjualan.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub cRekeningHutangDagang_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", , "kodeakun", sisPrefix, "2", " AND jenis = 'D'")
  If Not dbData.EOF Then
    cRekeningHutangDagang.Text = cRekeningHutangDagang.Browse(dbData)
    cNamaRekeningHutangDagang.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub cRekeningHutangSalesman_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", , "kodeakun", sisPrefix, "2", " AND jenis = 'D'")
  If Not dbData.EOF Then
    cRekeningHutangSalesman.Text = cRekeningHutangSalesman.Browse(dbData)
    cNamaRekeningHutangSalesman.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub cRekeningPendapatanFeeKartu_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", , "kodeakun", sisPrefix, "4", " AND jenis = 'D'")
  If Not dbData.EOF Then
    cRekeningPendapatanFeeKartu.Text = cRekeningPendapatanFeeKartu.Browse(dbData)
    cNamaRekeningPendapatanFeeKartu.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub cRekeningPenjualan_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", , "kodeakun", sisPrefix, "4", " AND jenis = 'D'")
  If Not dbData.EOF Then
    cRekeningPenjualan.Text = cRekeningPenjualan.Browse(dbData)
    cNamaRekeningPenjualan.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub cRekeningPiutangDagang_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", , "kodeakun", sisPrefix, "1", " AND jenis = 'D'")
  If Not dbData.EOF Then
    cRekeningPiutangDagang.Text = cRekeningPiutangDagang.Browse(dbData)
    cNamaRekeningPiutangDagang.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub cRekeningPotonganHutang_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", , "kodeakun", sisPrefix, "4", " AND jenis = 'D'")
  If Not dbData.EOF Then
    cRekeningPotonganHutang.Text = cRekeningPotonganHutang.Browse(dbData)
    cNamaRekeningPotonganHutang.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub cRekeningPotonganPiutang_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", , "kodeakun", sisPrefix, "5", " AND jenis = 'D'")
  If Not dbData.EOF Then
    cRekeningPotonganPiutang.Text = cRekeningPotonganPiutang.Browse(dbData)
    cNamaRekeningPotonganPiutang.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub cRekeningPPnPenjualan_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", , "kodeakun", sisPrefix, "2", " AND jenis = 'D'")
  If Not dbData.EOF Then
    cRekeningPPnPenjualan.Text = cRekeningPPnPenjualan.Browse(dbData)
    cNamaRekeningPPnPenjualan.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub cRekeningPrive_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", , "kodeakun", sisPrefix, "3", " AND jenis = 'D'")
  If Not dbData.EOF Then
    cRekeningPrive.Text = cRekeningPrive.Browse(dbData)
    cNamaRekeningPrive.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub cRekeningReturPembelian_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", , "kodeakun", sisPrefix, "2", " AND jenis = 'D' and keterangan like '%" & cRekeningReturPembelian.Text & "%'")
  If Not dbData.EOF Then
    cRekeningReturPembelian.Text = cRekeningReturPembelian.Browse(dbData)
    cNamaRekeningReturPembelian.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub cRekeningTitipanKasReturPembelian_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", , "kodeakun", sisPrefix, "1", " AND jenis = 'D'")
  If Not dbData.EOF Then
    cRekeningTitipanKasReturPembelian.Text = cRekeningTitipanKasReturPembelian.Browse(dbData)
    cNamaRekeningTitipanKasReturPembelian.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub cRekeningToUp_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", , "kodeakun", sisPrefix, "2", " AND jenis = 'D'")
  If Not dbData.EOF Then
    cRekeningToUp.Text = cRekeningBG.Browse(dbData)
    cNamaRekeningToUp.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  SetIcon Me.hWnd
  TabIndex cDiscountPembelian, n
  TabIndex cPPnPembelian, n
  TabIndex cRekeningPenjualan, n
  TabIndex cRekeningDiscountPenjualan, n
  TabIndex cRekeningPPnPenjualan, n
  TabIndex cRekeningCOGS, n
  TabIndex cRekeningPotonganPiutang, n
  TabIndex cRekeningPotonganHutang, n
  TabIndex cRekeningPiutangDagang, n
  TabIndex cRekeningHutangDagang, n
  TabIndex cRekeningReturPembelian, n
  TabIndex cRekeningHutangSalesman, n
  TabIndex cRekeningBiayaKomisi, n
  TabIndex cRekeningBG, n
  TabIndex cRekeningToUp, n
  TabIndex cRekeningAktivaTopUp, n
  TabIndex cRekDefaultSetoranKas, n
  TabIndex cRekeningPrive, n
  TabIndex cRekeningTitipanKasReturPembelian, n
  TabIndex cRekeningPendapatanFeeKartu, n
  
  
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  
  cDiscountPembelian.Text = aCfg(objData, msRekeningDiscountPembelian)
  cNamaRekeningDiscountPembelian.Text = GetNamaRekening(cDiscountPembelian.Text)
  cPPnPembelian.Text = aCfg(objData, msRekeningPPnPembelian)
  cNamaRekeningPPnPembelian.Text = GetNamaRekening(cPPnPembelian.Text)
  cRekeningPenjualan.Text = aCfg(objData, msRekeningPenjualan)
  cNamaRekeningPenjualan.Text = GetNamaRekening(cRekeningPenjualan.Text)
  cRekeningDiscountPenjualan.Text = aCfg(objData, msRekeningDiscountPenjualan)
  cNamaRekeningDiscountPenjualan.Text = GetNamaRekening(cRekeningDiscountPenjualan.Text)
  cRekeningPPnPenjualan.Text = aCfg(objData, msRekeningPPnPenjualan)
  cNamaRekeningPPnPenjualan.Text = GetNamaRekening(cRekeningPPnPenjualan.Text)
  cRekeningCOGS.Text = aCfg(objData, msRekeningCOGS)
  cNamaRekeningCOGS.Text = GetNamaRekening(cRekeningCOGS.Text)
  cRekeningPotonganPiutang.Text = aCfg(objData, msRekeningPotonganPiutang)
  cNamaRekeningPotonganPiutang.Text = GetNamaRekening(cRekeningPotonganPiutang.Text)
  cRekeningPotonganHutang.Text = aCfg(objData, msRekeningPotonganHutang)
  cNamaRekeningPotonganHutang.Text = GetNamaRekening(cRekeningPotonganHutang.Text)
  cRekeningPiutangDagang.Text = aCfg(objData, msPiutangDagang)
  cNamaRekeningPiutangDagang.Text = GetNamaRekening(cRekeningPiutangDagang.Text)
  cRekeningHutangDagang.Text = aCfg(objData, msHutangDagang)
  cNamaRekeningHutangDagang.Text = GetNamaRekening(cRekeningHutangDagang.Text)
  cRekeningReturPembelian.Text = aCfg(objData, msRekeningReturPembelian)
  cNamaRekeningReturPembelian.Text = GetNamaRekening(cRekeningReturPembelian.Text)
  cRekeningHutangSalesman.Text = aCfg(objData, msRekeningHutangSalesman)
  cNamaRekeningHutangSalesman.Text = GetNamaRekening(cRekeningHutangSalesman.Text)
  cRekeningBiayaKomisi.Text = aCfg(objData, msRekeningBiayaKomisi)
  cNamaRekeningBiayaKomisi.Text = GetNamaRekening(cRekeningBiayaKomisi.Text)
  cRekeningBG.Text = aCfg(objData, msRekeningBG)
  cNamaRekeningBG.Text = GetNamaRekening(cRekeningBG.Text)
  
  cRekeningToUp.Text = aCfg(objData, msRekeningTopUp)
  cNamaRekeningToUp.Text = GetNamaRekening(cRekeningToUp.Text)
  
  cRekeningAktivaTopUp.Text = aCfg(objData, msRekeningKasTopUp)
  cNamaRekeningAktivaTopUp.Text = GetNamaRekening(cRekeningAktivaTopUp.Text)
  
  cRekDefaultSetoranKas.Text = aCfg(objData, msRekeningSetoranKas)
  cNamaRekDefaultSetoranKas.Text = GetNamaRekening(cRekDefaultSetoranKas.Text)
  
  cRekeningPrive.Text = aCfg(objData, msRekeningPrive)
  cNamaRekeningPrive.Text = GetNamaRekening(cRekeningPrive.Text)
  cRekeningTitipanKasReturPembelian.Text = aCfg(objData, msRekeningTitipanKasReturPembelian)
  cNamaRekeningTitipanKasReturPembelian.Text = GetNamaRekening(cRekeningTitipanKasReturPembelian.Text)
  
  cRekeningPendapatanFeeKartu.Text = aCfg(objData, msRekeningFeeKartu)
  cNamaRekeningPendapatanFeeKartu.Text = GetNamaRekening(cRekeningPendapatanFeeKartu.Text)
  
  
End Sub

Private Function GetNamaRekening(cAkun As String) As String
  GetNamaRekening = ""
  Set dbData = objData.Browse(GetDSN, "akun", "keterangan", "kodeakun", sisAssign, cAkun)
  If Not dbData.EOF Then
    GetNamaRekening = GetNull(dbData!keterangan, "")
  End If
End Function

