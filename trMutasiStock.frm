VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trMutasiStock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mutasi Stock..."
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   12210
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   12210
   Begin SizerOneLibCtl.ElasticOne ElasticOne2 
      Height          =   585
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6705
      Width           =   12210
      _cx             =   21537
      _cy             =   1032
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
      Align           =   2
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
      Begin SizerOneLibCtl.ElasticOne ElasticOne5 
         Height          =   570
         Left            =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   15
         Width           =   12210
         _cx             =   21537
         _cy             =   1005
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
         Appearance      =   0
         MousePointer    =   0
         _ConvInfo       =   1
         Version         =   700
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   2
         AutoSizeChildren=   8
         BorderWidth     =   1
         ChildSpacing    =   1
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
         GridRows        =   7
         GridCols        =   8
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"trMutasiStock.frx":0000
         Begin BiSAButtonProject.BiSAButton cmdHapus 
            Height          =   480
            Left            =   2025
            TabIndex        =   2
            Top             =   15
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   847
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
            Picture         =   "trMutasiStock.frx":00B6
         End
         Begin BiSAButtonProject.BiSAButton cmdAktivasi 
            Height          =   480
            Left            =   3270
            TabIndex        =   3
            Top             =   15
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   847
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
            Picture         =   "trMutasiStock.frx":0340
         End
         Begin BiSAButtonProject.BiSAButton cmdEdit 
            Height          =   480
            Left            =   1035
            TabIndex        =   4
            Top             =   15
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   847
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
            Picture         =   "trMutasiStock.frx":04DF
         End
         Begin BiSAButtonProject.BiSAButton cmdAdd 
            Height          =   480
            Left            =   15
            TabIndex        =   5
            Top             =   15
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   847
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
            Picture         =   "trMutasiStock.frx":060B
         End
         Begin BiSAButtonProject.BiSAButton cmdKeluar 
            Cancel          =   -1  'True
            Height          =   480
            Left            =   11040
            TabIndex        =   6
            Top             =   15
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   847
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
            Picture         =   "trMutasiStock.frx":07B6
         End
         Begin BiSAButtonProject.BiSAButton cmdSimpan 
            Height          =   480
            Left            =   9840
            TabIndex        =   7
            Top             =   15
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   847
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
            Picture         =   "trMutasiStock.frx":085C
         End
      End
   End
   Begin SizerOneLibCtl.ElasticOne ElasticOne1 
      Height          =   1890
      Left            =   0
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   12210
      _cx             =   21537
      _cy             =   3334
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
      Appearance      =   0
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   1
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
      Begin BiSATextBoxProject.BiSATextBox cKeterangan 
         Height          =   360
         Left            =   105
         TabIndex        =   9
         Top             =   1155
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   635
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
         Caption         =   "Ket."
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
      Begin BiSATextBoxProject.BiSABrowse cNomor 
         Height          =   345
         Left            =   105
         TabIndex        =   10
         Top             =   795
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   609
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
      Begin BiSADateProject.BiSADate dTgl 
         Height          =   360
         Left            =   105
         TabIndex        =   11
         Top             =   420
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   635
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
   End
   Begin SizerOneLibCtl.ElasticOne ElasticOne4 
      Height          =   4815
      Left            =   0
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1890
      Width           =   12210
      _cx             =   21537
      _cy             =   8493
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
      Align           =   5
      AutoSizeChildren=   8
      BorderWidth     =   4
      ChildSpacing    =   1
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
      GridRows        =   2
      GridCols        =   11
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"trMutasiStock.frx":0AE2
      Begin BiSATextBoxProject.BiSATextBox cSatuan 
         Height          =   345
         Left            =   10335
         TabIndex        =   13
         Top             =   60
         Width           =   1290
         _ExtentX        =   2275
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
      Begin BiSANumberBoxProject.BiSANumberBox nQty 
         Height          =   345
         Left            =   9015
         TabIndex        =   14
         Top             =   60
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   609
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
      Begin BiSATextBoxProject.BiSABrowse cKeGudang 
         Height          =   345
         Left            =   7440
         TabIndex        =   15
         Top             =   60
         Width           =   1560
         _ExtentX        =   2752
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
      Begin BiSATextBoxProject.BiSABrowse cDariGudang 
         Height          =   345
         Left            =   5865
         TabIndex        =   16
         Top             =   60
         Width           =   1560
         _ExtentX        =   2752
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
         Height          =   345
         Left            =   11640
         TabIndex        =   17
         Top             =   60
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   609
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
         BackColor       =   -2147483633
         Picture         =   "trMutasiStock.frx":0B94
      End
      Begin BiSATextBoxProject.BiSABrowse cNama 
         Height          =   345
         Left            =   2250
         TabIndex        =   18
         Top             =   60
         Width           =   3600
         _ExtentX        =   6350
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
         Height          =   345
         Left            =   765
         TabIndex        =   19
         Top             =   60
         Width           =   1470
         _ExtentX        =   2593
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
      Begin BiSANumberBoxProject.BiSANumberBox nNo 
         Height          =   345
         Left            =   60
         TabIndex        =   20
         Top             =   60
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   609
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
      Begin TrueOleDBGrid70.TDBGrid GridMutasi 
         Height          =   4335
         Left            =   60
         TabIndex        =   21
         Top             =   420
         Width           =   12090
         _ExtentX        =   21325
         _ExtentY        =   7646
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "No"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Kode"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Nama"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Dari"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Ke"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Qty"
         Columns(5).DataField=   ""
         Columns(5).NumberFormat=   "###,###,###,##0.00"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Satuan"
         Columns(6).DataField=   ""
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   7
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   873
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   15790320
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1244"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1164"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=2619"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2540"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=512"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=6403"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=6324"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=512"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=2752"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2672"
         Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(25)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=516"
         Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(31)=   "Column(5).Width=2328"
         Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=2249"
         Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=514"
         Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(37)=   "Column(6).Width=3069"
         Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=2990"
         Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=512"
         Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         Appearance      =   0
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
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bold=0,.fontsize=1125"
         _StyleDefs(11)  =   ":id=2,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=Tahoma"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(14)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(15)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(16)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(17)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(18)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(19)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(20)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(21)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(22)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(23)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(24)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(25)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(26)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(27)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(28)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(29)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(30)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(31)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(32)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(33)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=58,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.alignment=0"
         _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=54,.parent=13,.alignment=0"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=46,.parent=13"
         _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.alignment=0"
         _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(62)  =   "Named:id=33:Normal"
         _StyleDefs(63)  =   ":id=33,.parent=0"
         _StyleDefs(64)  =   "Named:id=34:Heading"
         _StyleDefs(65)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(66)  =   ":id=34,.wraptext=-1"
         _StyleDefs(67)  =   "Named:id=35:Footing"
         _StyleDefs(68)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(69)  =   "Named:id=36:Selected"
         _StyleDefs(70)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(71)  =   "Named:id=37:Caption"
         _StyleDefs(72)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(73)  =   "Named:id=38:HighlightRow"
         _StyleDefs(74)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(75)  =   "Named:id=39:EvenRow"
         _StyleDefs(76)  =   ":id=39,.parent=33,.bgcolor=&HE6E6E6&"
         _StyleDefs(77)  =   "Named:id=40:OddRow"
         _StyleDefs(78)  =   ":id=40,.parent=33"
         _StyleDefs(79)  =   "Named:id=41:RecordSelector"
         _StyleDefs(80)  =   ":id=41,.parent=34"
         _StyleDefs(81)  =   "Named:id=42:FilterBar"
         _StyleDefs(82)  =   ":id=42,.parent=33"
      End
   End
End
Attribute VB_Name = "trMutasiStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nPos As SisPos
Dim lEdit As Boolean
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB
Dim cID As String


Private Sub GetFakturBrowse(ByVal lStat As Boolean)
  cNomor.Button = lStat
End Sub

Private Sub cDariGudang_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "gudang", "kodegudang,keterangan", "kodegudang", sisDifference, cKeGudang.Text, " and lstatus = 'A'", "kodegudang")
  If Not dbData.EOF Then
    cDariGudang.Text = cDariGudang.Browse(dbData)
  End If
End Sub

Private Sub cKeGudang_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "gudang", "kodegudang,keterangan", "kodegudang", sisDifference, cDariGudang.Text, " and lstatus = 'A'", "kodegudang")
  If Not dbData.EOF Then
    cKeGudang.Text = cKeGudang.Browse(dbData)
  End If
End Sub

Private Sub cKode_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "stock", "barcode,nama,kodesatuan,hargajual,kodestock", "barcode", sisContent, cKode.Text, " AND jenis < 9 AND statusnonaktif <> 1")
  If Not dbData.EOF Then
    cKode.Text = cKode.Browse(dbData, Array("BARCODE", "NAMA", "SATUAN", "JUAL", "SKU"), , Array(13, 35, 10, 8, 8))
    'cKode.Text = cKode.Browse(dbData)
    cKode.Text = GetNull(dbData!barcode)
    cNama.Text = GetNull(dbData!nama)
    cSatuan.Text = GetNull(dbData!kodesatuan, "")
    cID = GetNull(dbData!KodeStock)
  End If
End Sub

Private Sub cmdAktivasi_Click()
  frmAktivasi.Action Me
End Sub

Private Sub cmdAdd_Click()
  nPos = Add
  GetEdit True
  GetFakturBrowse False
  cNomor.Text = GetNomor("totmutasistock", "nomormutasistock", GetID, sisModulTransaksi.MutasiStock)
End Sub

Private Sub cmdEdit_Click()
  nPos = Edit
  GetEdit True
  GetFakturBrowse True
End Sub

Private Sub cmdHapus_Click()
  nPos = Delete
  GetEdit True
  GetFakturBrowse True
End Sub

Private Sub cmdKeluar_Click()
  If lEdit Then
    GetEdit False
    InitValueGrid
    initvalue
    GetFakturBrowse False
  Else
    Unload Me
  End If
End Sub

Private Function ValidSaving() As Boolean
  ValidSaving = True
End Function

Private Sub cmdOK_Click()
Dim n As Integer

If isValidOk Then
  'jika baris <= Nomor
    If vaArray.UpperBound(1) + 2 <= nNo.Value Then
      vaArray.ReDim 0, nNo.Value - 1, 0, 7
      n = vaArray.UpperBound(1)
    'jika baris=0
    ElseIf vaArray.UpperBound(1) = -1 Then
      nNo.Value = 1
      vaArray.ReDim 0, nNo.Value - 1, 0, 7
      n = vaArray.UpperBound(1)
    Else
      n = nNo.Value - 1
    End If
        
    vaArray(n, 0) = nNo.Value
    vaArray(n, 1) = cKode.Text
    vaArray(n, 2) = cNama.Text
    vaArray(n, 3) = cDariGudang.Text
    vaArray(n, 4) = cKeGudang.Text
    vaArray(n, 5) = nQty.Value
    vaArray(n, 6) = cSatuan.Text
    vaArray(n, 7) = cID
    
    GridMutasi.Array = vaArray
    GridMutasi.ReBind
    nNo.Value = vaArray.UpperBound(1) + 2
    nNo.SetFocus
    InitValueGrid
  End If
End Sub

Private Function isValidOk()
isValidOk = True
  
  If Trim(cKeGudang.Text) = "" Then
    isValidOk = False
    Exit Function
  End If
  
  If Trim(cDariGudang.Text) = "" Then
    isValidOk = False
    Exit Function
  End If
  
  'jika barang di gudang tidak cukup maka munculkan pesan error
  If GetInfoStockDong2(objData, cID, cDariGudang.Text) - nQty.Value < 0 Then
    MsgBox "Maaf barang di gudang tidak cukup jumlahnya untuk dimutasikan.. " & vbCrLf & "Jumlah stock di gudang : " & GetInfoStockDong2(objData, cID, cDariGudang.Text), vbExclamation
    isValidOk = False
    Exit Function
  End If
  
  'Jika kode gudang tidak valid, maka penyimpanan data tidak diijinkan
  Set dbData = objData.Browse(GetDSN, "gudang", "lstatus", "kodegudang", sisAssign, cDariGudang.Text)
  If Not dbData.EOF Then
    If GetNull(dbData!lStatus) <> "A" Then
      MsgBox "Kode Gudang tidak valid, atau tidak aktif, Data tidak bisa disimpan", vbExclamation
      isValidOk = False
      Exit Function
    End If
  End If
  
  Set dbData = objData.Browse(GetDSN, "gudang", "lstatus", "kodegudang", sisAssign, cKeGudang.Text)
  If Not dbData.EOF Then
    If GetNull(dbData!lStatus) <> "A" Then
      MsgBox "Kode Gudang tidak valid, atau tidak aktif, Data tidak bisa disimpan", vbExclamation
      isValidOk = False
      Exit Function
    End If
  End If
  
End Function

Private Sub cmdSimpan_Click()
Dim lSave As Boolean
Dim Faktur As String
Dim n As Integer

  lSave = True
  objData.Start GetDSN
  Faktur = cNomor.Text
  If Faktur = "" Then
    MsgBox "Maaf data tidak bisa disimpan. Nomor transaksi tidak disertakan"
    cmdKeluar_Click
    Exit Sub
  End If
  
  'Simpan dulu di table induk
  'Hapus terlebih dahulu tabel child yg berelasi
  
  lSave = IIf(lSave, objData.Delete(GetDSN, "totmutasistock", "nomormutasistock", sisAssign, Faktur), False)
  lSave = IIf(lSave, objData.Delete(GetDSN, "mutasistock", "nomormutasistock", sisAssign, Faktur), False)
  lSave = IIf(lSave, objData.Delete(GetDSN, "kartustock", "nomor", sisAssign, Faktur), False)
  'update table totmutasistock
  lSave = IIf(lSave, objData.Update(GetDSN, "totmutasistock", "nomormutasistock = '" & Faktur & "'", Array("nomormutasistock", "username", "tgl", "datetime", "keterangan"), Array(Faktur, GetRegistry(reg_Username), Format(dTgl.Value, "yyyy-MM-dd"), SNow, cKeterangan.Text)), False)
  
  For n = 0 To vaArray.UpperBound(1)
    'Simpan di table mutasistock dari
    lSave = IIf(lSave, objData.Add(GetDSN, "mutasistock", Array("nomormutasistock", "kodestock", "gudangdari", "gudangke", "qty"), Array(cNomor.Text, vaArray(n, 7), vaArray(n, 3), vaArray(n, 4), vaArray(n, 5))), False)
    
    'Update table ke kartustock
'    lSave = IIf(lSave, UpdKartuStock(objData, SisKartuStock.MutasiDari, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), vaArray(n, 7), vaArray(n, 5), 0, 0, "Mutasi Stock dari " & vaArray(n, 3) & " ke " & vaArray(n, 4), vaArray(n, 3)), False)
'    lSave = IIf(lSave, UpdKartuStock(objData, SisKartuStock.MutasiKe, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), vaArray(n, 7), vaArray(n, 5), 0, 0, "Mutasi Stock Ke " & vaArray(n, 4) & " dari " & vaArray(n, 3), vaArray(n, 4)), False)
    lSave = IIf(lSave, UpdKartuStock(objData, SisKartuStock.MutasiDari, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), vaArray(n, 7), vaArray(n, 5), GetHargaPokok(objData, vaArray(n, 7)), 0, "Mutasi Stock ke " & vaArray(n, 4), vaArray(n, 3), GetHargaPokok(objData, vaArray(n, 7))), False)
    lSave = IIf(lSave, UpdKartuStock(objData, SisKartuStock.MutasiKe, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), vaArray(n, 7), vaArray(n, 5), GetHargaPokok(objData, vaArray(n, 7)), 0, "Mutasi Stock dari " & vaArray(n, 3), vaArray(n, 4), GetHargaPokok(objData, vaArray(n, 7))), False)
  
  Next
  
  If lSave Then
    objData.Save GetDSN
  Else
    objData.Cancel GetDSN
  End If
  GetEdit False
End Sub

Private Sub cNama_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "stock", "nama,barcode,kodestock,kodesatuan", "nama", sisContent, cNama.Text, " AND jenis < 9 AND statusnonaktif <> 1")
  If Not dbData.EOF Then
    cNama.Text = cNama.Browse(dbData)
    cKode.Text = GetNull(dbData!barcode, "")
    cSatuan.Text = GetNull(dbData!kodesatuan, "")
    cID = GetNull(dbData!KodeStock)
  End If
End Sub

Private Sub cNomor_ButtonClick()
Dim lSave As Boolean

  lSave = True
  Set dbData = objData.Browse(GetDSN, "totmutasistock", "nomormutasistock,keterangan,username", "tgl", sisAssign, Format(dTgl.Value, "yyyy-MM-dd"))
  If Not dbData.EOF Then
    cNomor.Text = cNomor.Browse(dbData)
    cKeterangan.Text = GetNull(dbData!keterangan)
    GetLoadRows
    Me.Refresh
    If nPos = Delete Then
      If MsgBox("Yakin data akan dihapus?", vbYesNo) = vbYes Then
        'Hapus terlebih dahulu table child
        'stockopname, kartustock
        lSave = IIf(lSave, objData.Delete(GetDSN, "kartustock", "nomor", sisAssign, cNomor.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "mutasistock", "nomormutasistock", sisAssign, cNomor.Text), False)
        'Hapus table master
        lSave = IIf(lSave, objData.Delete(GetDSN, "totmutasistock", "nomormutasistock", sisAssign, cNomor.Text), False)
      End If
      GetEdit False
      initvalue
    End If
  End If
End Sub

Private Sub GetLoadRows()
Dim n As Integer

  vaArray.ReDim 0, -1, 0, 7
  Set dbData = objData.Browse(GetDSN, "totmutasistock t", "t.nomormutasistock,t.keterangan,m.kodestock,s.nama,s.barcode,s.kodesatuan,s.kodestock,m.gudangdari,m.gudangke,m.qty", "t.nomormutasistock", sisAssign, cNomor.Text, , , Array("left join mutasistock m on m.nomormutasistock = t.nomormutasistock", "left join stock s on s.kodestock = m.kodestock"))
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = n + 1
      vaArray(n, 1) = GetNull(dbData!barcode)
      vaArray(n, 2) = GetNull(dbData!nama)
      vaArray(n, 3) = GetNull(dbData!gudangdari)
      vaArray(n, 4) = GetNull(dbData!gudangke)
      vaArray(n, 5) = GetNull(dbData!qty)
      vaArray(n, 6) = GetNull(dbData!kodesatuan)
      vaArray(n, 7) = GetNull(dbData!KodeStock)
      dbData.MoveNext
    Loop
  End If
  Set GridMutasi.Array = vaArray
  GridMutasi.ReBind
  GridMutasi.Refresh
End Sub

Private Sub Form_Activate()
  Me.Refresh
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  initvalue
  GetEdit False
  TabIndex dTgl, n
  TabIndex cNomor, n
  TabIndex cKeterangan, n
  'Tabindex untuk gridmutasi
  TabIndex nNo, n
  TabIndex cKode, n
  TabIndex cNama, n
  TabIndex cDariGudang, n
  TabIndex cKeGudang, n
  TabIndex nQty, n
  TabIndex cSatuan, n
  TabIndex cmdOK, n
  
  TabIndex cmdAdd, n
  TabIndex cmdEdit, n
  TabIndex cmdHapus, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  TabIndex cmdAktivasi, n
End Sub

Private Sub initvalue()
  dTgl.Value = Date
  cNomor.Default
  cKeterangan.Default
  nNo.Value = 1
  vaArray.ReDim 0, -1, 0, 7
  Set GridMutasi.Array = vaArray
  GridMutasi.ReBind
  GridMutasi.Refresh
End Sub

Private Sub InitValueGrid()
  cKode.Default
  cNama.Default
  cDariGudang.Default
  cKeGudang.Default
  nQty.Value = 1
  cSatuan.Default
  cID = ""
End Sub

Private Sub GetEdit(lPar As Boolean)
  ElasticOne1.Enabled = lPar
  ElasticOne4.Enabled = lPar
  lEdit = lPar
  initvalue
  SetButton cmdSimpan, cmdKeluar, cmdAdd, cmdEdit, cmdHapus, nPos, lPar, cmdAktivasi
  
  If lPar Then
    dTgl.SetFocus
    If nPos = Add Then
      cNomor.Enabled = False
      cNomor.BackColor = vbButtonFace
    Else
      cNomor.Enabled = True
      cNomor.BackColor = vbWindowBackground
      cNomor.CaptionBackColor = vbButtonFace
    End If
  End If
End Sub

Private Sub GridMutasi_KeyDown(KeyCode As Integer, Shift As Integer)
Dim n As Integer

  If KeyCode = vbKeyDelete Then
    If vaArray.UpperBound(1) >= 0 Then
      GridMutasi.Delete
      For n = 0 To vaArray.UpperBound(1)
        vaArray(n, 0) = n + 1
      Next
      nNo.Value = vaArray.UpperBound(1) + 2
      GridMutasi.ReBind
    End If
  End If
End Sub

Private Sub nNo_Validate(Cancel As Boolean)
Dim n As Single
  
  If GetValidNomorUrut(nNo, vaArray) Then
    n = nNo.Value - 1
    If n <= vaArray.UpperBound(1) Then
      cKode.Text = vaArray(n, 1)
      cNama.Text = vaArray(n, 2)
      cDariGudang.Text = vaArray(n, 3)
      cKeGudang.Text = vaArray(n, 4)
      nQty.Value = vaArray(n, 5)
      cSatuan.Text = vaArray(n, 6)
      cID = vaArray(n, 7)
    End If
  End If
End Sub

