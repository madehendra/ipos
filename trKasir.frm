VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form trKasir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "... ::: KASIR :::..."
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11745
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   11745
   WindowState     =   2  'Maximized
   Begin SizerOneLibCtl.ElasticOne ElasticOne2 
      Height          =   435
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1545
      Width           =   11745
      _cx             =   20717
      _cy             =   767
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
      Align           =   1
      AutoSizeChildren=   8
      BorderWidth     =   2
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
      GridRows        =   1
      GridCols        =   8
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"trKasir.frx":0000
      Begin BiSANumberBoxProject.BiSANumberBox nNumber 
         Height          =   375
         Left            =   30
         TabIndex        =   7
         Top             =   30
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   661
         Appearance      =   0
         Decimals        =   0
         DecimalPoint    =   ""
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
      Begin BiSANumberBoxProject.BiSANumberBox nQty 
         Height          =   375
         Left            =   5625
         TabIndex        =   8
         Top             =   30
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   661
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
      Begin BiSATextBoxProject.BiSATextBox cSatuan 
         Height          =   375
         Left            =   6420
         TabIndex        =   9
         Top             =   30
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   661
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
         BackColor       =   12640511
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
      Begin BiSANumberBoxProject.BiSANumberBox nHarga 
         Height          =   375
         Left            =   7275
         TabIndex        =   10
         Top             =   30
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   661
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
         BackColor       =   12640511
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
      Begin BiSANumberBoxProject.BiSANumberBox nJumlah 
         Height          =   375
         Left            =   8970
         TabIndex        =   11
         Top             =   30
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   661
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
         BackColor       =   12640511
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
         Height          =   375
         Left            =   11130
         TabIndex        =   12
         Top             =   30
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   661
         Caption         =   "+"
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
      End
      Begin BiSATextBoxProject.BiSABrowse cNama 
         Height          =   375
         Left            =   2310
         TabIndex        =   13
         Top             =   30
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   661
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
         Height          =   375
         Left            =   495
         TabIndex        =   14
         Top             =   30
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   661
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
   End
   Begin SizerOneLibCtl.ElasticOne ElasticOne1 
      Height          =   1545
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   11745
      _cx             =   20717
      _cy             =   2725
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
      Align           =   1
      AutoSizeChildren=   8
      BorderWidth     =   2
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
      GridRows        =   1
      GridCols        =   1
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"trKasir.frx":008A
      Begin BiSANumberBoxProject.BiSANumberBox nPenjualan 
         Height          =   1485
         Left            =   30
         TabIndex        =   15
         Top             =   30
         Width           =   11685
         _ExtentX        =   20611
         _ExtentY        =   2619
         Appearance      =   0
         Enabled         =   0   'False
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   72
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
   Begin SizerOneLibCtl.ElasticOne ElasticOne3 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   7320
      Width           =   11745
      _cx             =   20717
      _cy             =   1085
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
      GridRows        =   2
      GridCols        =   6
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"trKasir.frx":00BE
      Begin BiSAButtonProject.BiSAButton cmdBayar 
         Height          =   465
         Left            =   9240
         TabIndex        =   1
         Top             =   15
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   820
         Caption         =   "    &Bayar"
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
         Picture         =   "trKasir.frx":0133
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Height          =   465
         Left            =   10500
         TabIndex        =   2
         Top             =   15
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   820
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
         Picture         =   "trKasir.frx":03B9
      End
      Begin BiSATextBoxProject.BiSABrowse cAkunKas 
         Height          =   465
         Left            =   15
         TabIndex        =   26
         Top             =   15
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   820
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
      Begin VB.Label Label1 
         Caption         =   "F6 = Mode Otomatis F7 = Manual"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3465
         TabIndex        =   27
         Top             =   15
         Width           =   5760
      End
   End
   Begin SizerOneLibCtl.ElasticOne ElasticOne4 
      Height          =   5340
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1980
      Width           =   11745
      _cx             =   20717
      _cy             =   9419
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
      AutoSizeChildren=   7
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
      Begin BiSAFramProject.BiSAFrame BisaBayar 
         Height          =   2985
         Left            =   2070
         Top             =   795
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   5265
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   2
         BackColor       =   -2147483633
         Begin SizerOneLibCtl.ElasticOne ElasticOne5 
            Height          =   2985
            Left            =   0
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   0
            Width           =   7695
            _cx             =   13573
            _cy             =   5265
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
            AutoSizeChildren=   7
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
            Begin SizerOneLibCtl.ElasticOne ElasticOne6 
               Height          =   2955
               Left            =   15
               TabIndex        =   17
               TabStop         =   0   'False
               Top             =   15
               Width           =   7665
               _cx             =   13520
               _cy             =   5212
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
               Appearance      =   1
               MousePointer    =   0
               _ConvInfo       =   1
               Version         =   700
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   ""
               Align           =   0
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
               GridRows        =   9
               GridCols        =   10
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"trKasir.frx":045F
               Begin BiSAButtonProject.BiSAButton cmdPersen 
                  Height          =   360
                  Left            =   3315
                  TabIndex        =   18
                  TabStop         =   0   'False
                  Top             =   510
                  Width           =   510
                  _ExtentX        =   900
                  _ExtentY        =   635
                  Caption         =   "% [F3]"
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
               Begin BiSANumberBoxProject.BiSANumberBox nSubTotal 
                  Height          =   345
                  Left            =   315
                  TabIndex        =   19
                  Top             =   150
                  Width           =   2910
                  _ExtentX        =   5133
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
                  BackColor       =   12640511
                  Caption         =   "Sub Total"
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
               Begin BiSANumberBoxProject.BiSANumberBox nDiscountBayar 
                  Height          =   360
                  Left            =   315
                  TabIndex        =   20
                  Top             =   510
                  Width           =   2910
                  _ExtentX        =   5133
                  _ExtentY        =   635
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
                  Caption         =   "Disc (Rp)"
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
               Begin BiSANumberBoxProject.BiSANumberBox nTotal 
                  Height          =   345
                  Left            =   4365
                  TabIndex        =   21
                  Top             =   150
                  Width           =   2985
                  _ExtentX        =   5265
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
                  BackColor       =   12640511
                  Caption         =   "Total"
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
                  Height          =   360
                  Left            =   4365
                  TabIndex        =   22
                  Top             =   510
                  Width           =   2985
                  _ExtentX        =   5265
                  _ExtentY        =   635
                  Appearance      =   0
                  BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "Tunai"
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
               Begin BiSAButtonProject.BiSAButton cmdSimpan 
                  Height          =   465
                  Left            =   6525
                  TabIndex        =   23
                  Top             =   2355
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   820
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
                  Picture         =   "trKasir.frx":0550
               End
               Begin BiSANumberBoxProject.BiSANumberBox nKembali 
                  Height          =   885
                  Left            =   315
                  TabIndex        =   24
                  Top             =   1365
                  Width           =   7035
                  _ExtentX        =   12409
                  _ExtentY        =   1561
                  Appearance      =   0
                  Enabled         =   0   'False
                  BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Times New Roman"
                     Size            =   48
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BackColor       =   -2147483634
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
               Begin VB.Label Label2 
                  Caption         =   "Kembali"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   315
                  TabIndex        =   25
                  Top             =   1125
                  Width           =   1440
               End
            End
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   5295
         Left            =   0
         TabIndex        =   6
         Top             =   60
         Width           =   11700
         _ExtentX        =   20638
         _ExtentY        =   9340
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
         Columns(3).NumberFormat=   "###,###,##0.00"
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
         Columns(5).NumberFormat=   "###,###,##0.00"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "JUMLAH"
         Columns(6).DataField=   ""
         Columns(6).NumberFormat=   "###,###,##0.00"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=767"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=688"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=514"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=3201"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3122"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=512"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=5927"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=5847"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=512"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=1402"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1323"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=197122"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=1535"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=1455"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=513"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=2990"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2910"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=514"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=3836"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=3757"
         Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=514"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         Appearance      =   0
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   0
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   12632256
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
         _StyleDefs(13)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bgcolor=&H8000000F&,.bold=0"
         _StyleDefs(15)  =   ":id=3,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(16)  =   ":id=3,.fontname=MS Sans Serif"
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
         _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.alignment=2,.appearance=0,.bold=0"
         _StyleDefs(28)  =   ":id=14,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(29)  =   ":id=14,.fontname=Verdana"
         _StyleDefs(30)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(31)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(32)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(33)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=1"
         _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
         _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=0"
         _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=66,.parent=13,.alignment=1"
         _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=14"
         _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=15,.alignment=1"
         _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=17"
         _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=2"
         _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(59)  =   "Splits(0).Columns(5).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(60)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
         _StyleDefs(61)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
         _StyleDefs(62)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
         _StyleDefs(63)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.alignment=1,.bgcolor=&HFFFFFF&"
         _StyleDefs(64)  =   ":id=62,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(65)  =   ":id=62,.fontname=Tahoma"
         _StyleDefs(66)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(67)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(68)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(69)  =   "Named:id=33:Normal"
         _StyleDefs(70)  =   ":id=33,.parent=0"
         _StyleDefs(71)  =   "Named:id=34:Heading"
         _StyleDefs(72)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(73)  =   ":id=34,.wraptext=-1"
         _StyleDefs(74)  =   "Named:id=35:Footing"
         _StyleDefs(75)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(76)  =   "Named:id=36:Selected"
         _StyleDefs(77)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(78)  =   "Named:id=37:Caption"
         _StyleDefs(79)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(80)  =   "Named:id=38:HighlightRow"
         _StyleDefs(81)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(82)  =   "Named:id=39:EvenRow"
         _StyleDefs(83)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(84)  =   "Named:id=40:OddRow"
         _StyleDefs(85)  =   ":id=40,.parent=33"
         _StyleDefs(86)  =   "Named:id=41:RecordSelector"
         _StyleDefs(87)  =   ":id=41,.parent=34"
         _StyleDefs(88)  =   "Named:id=42:FilterBar"
         _StyleDefs(89)  =   ":id=42,.parent=33"
      End
   End
End
Attribute VB_Name = "trKasir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB
Dim vaTmp As New XArrayDB
Dim cFaktur As String

Private Sub PrintStrukUSB()
  trPrintKasir.noOrder = cFaktur
  Set dbData = objData.Browse(GetDSN, "totkasir t", "t.*", "t.nomorkasir", sisAssign, cFaktur)
  If Not dbData.EOF Then
    trPrintKasir.nSubTotal = GetNull(dbData!Subtotal)
    trPrintKasir.nDiscount = GetNull(dbData!Discount)
    trPrintKasir.nTotal = GetNull(dbData!Total)
    trPrintKasir.nCash = GetNull(dbData!Tunai)
    trPrintKasir.nChange = GetNull(dbData!Tunai) - GetNull(dbData!Total)
    
    Load trPrintKasir
    trPrintKasir.Show vbModal
  End If
End Sub



Private Sub PrintThermal()
Dim n As Double
Dim nBruto As Double
Dim nTotQty As Double
Dim nHargaArray As Double


  Open "lpt1" For Output As #1
        Print #1, Chr(27); Chr(33); Chr(4);

    Print #1, Chr(27) & Chr(97) & Chr(1)
    Print #1, "STRUK KASIR"
    Print #1, aCfg(objData, msNamaPerusahaan)
    Print #1, aCfg(objData, msAlamatPerusahaan)
    Select Case GetRegistry(reg_AlignmentThermal)
      Case 1 ' rata kiri
                Print #1, Chr(27) & Chr(97) & Chr(0)
      Case 2 ' rata kanan
          Print #1, Chr(27) & Chr(97) & Chr(2)
    End Select

    Print #1, ""
    Print #1, "No. " & cFaktur
    Print #1, Format(Now, "dd-MM-yyyy HH:MM:SS")
    Print #1, ""
    
    Print #1, Replicate("-", 27)
    Print #1, Padl("Qty", 6); Padl("Hrg Net", 11); Padl("Jml", 10)
    Print #1, Replicate("-", 27)
    For n = 0 To vaArray.UpperBound(1)
      If vaArray(n, 3) <> 0 Then
        nBruto = nBruto + (vaArray(n, 3) * vaArray(n, 5))
        nTotQty = nTotQty + vaArray(n, 3)
        Print #1, vaArray(n, 2) ' vaArray(n, 1) kolom REF barang ditiadakan karena terlalu panjang
        Print #1, Padl(Format(vaArray(n, 3), "#,##0"), 3) & " x " & Padl(Format(vaArray(n, 5), "#,###,##0"), 8) & " = " & Padl(Format(vaArray(n, 3) * vaArray(n, 5), "#,###,##0"), 10)
      End If
    Next
    
    Print #1, Replicate("-", 27)
    Print #1, Format(nTotQty, "###,###,##0") & " Items"
    Print #1, Padl("Bruto  : ", 9); Padl(Format(nBruto, "###,###,##0"), 10)
    Print #1, Padl("Disc   : ", 9); Padl(Format(nDiscountBayar.Value, "###,###,##0"), 10)
    Print #1, Padl("Total  : ", 9); Padl(Format(nTotal.Value, "###,###,##0"), 10)
    Print #1, Padl("Tunai  : ", 9); Padl(Format(nTunai.Value, "###,###,##0"), 10)
    Print #1, Padl("Kembali: ", 9); Padl(Format(nKembali.Value, "###,###,##0"), 10)
    Print #1, Chr(27) & Chr(97) & Chr(1) ' rata tengah
    If Trim(aCfg(objData, msKasir1, "")) <> "" Then
      Print #1, aCfg(objData, msKasir1, "")
    End If
    If Trim(aCfg(objData, msKasir2, "")) <> "" Then
      Print #1, aCfg(objData, msKasir2, "")
    End If
    If Trim(aCfg(objData, msKasir3, "")) <> "" Then
      Print #1, aCfg(objData, msKasir3, "")
    End If
    If Trim(aCfg(objData, msKasir4, "")) <> "" Then
      Print #1, aCfg(objData, msKasir4, "")
    End If
    If Trim(aCfg(objData, msKasir5, "")) <> "" Then
      Print #1, aCfg(objData, msKasir5, "")
    End If
    If Trim(aCfg(objData, msKasir6, "")) <> "" Then
      Print #1, aCfg(objData, msKasir6, "")
    End If
    If Trim(aCfg(objData, msKasir7, "")) <> "" Then
      Print #1, aCfg(objData, msKasir7, "")
    End If
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, Chr(27); Chr(33); Chr(0);
   Close #1
End Sub


Private Sub SumJumlah()
  nJumlah.Value = nHarga.Value * nQty.Value
  vaTmp(0, 6) = nJumlah.Value
End Sub

Private Sub cKode_LostFocus()
Dim cID As String

  If cKode.Text <> "" Then
    Set dbData = objData.Browse(GetDSN, "stock", "barcode,nama,kodesatuan,hargajual,kodestock", "barcode", sisAssign, cKode.Text, , "nama")
    If Not dbData.EOF Then
      cKode.Text = GetNull(dbData!barcode, "")
      cNama.Text = GetNull(dbData!nama, "")
      cSatuan.Text = GetNull(dbData!kodesatuan, "")
      nHarga.Value = GetNull(dbData!hargajual) 'IIf(GetNull(dbData!hargajual, 0) <= 0, GetNull(dbData!hargabeli) + GetNull(dbData!hargabeli) * 15, GetNull(dbData!hargajual))
      vaTmp(0, 1) = cKode.Text
      vaTmp(0, 2) = cNama.Text
      vaTmp(0, 3) = nQty.Value
      vaTmp(0, 4) = cSatuan.Text
      vaTmp(0, 5) = nHarga.Value
      vaTmp(0, 6) = nJumlah.Value
      vaTmp(0, 7) = GetNull(dbData!KodeStock)
      SumJumlah
      If aCfg(objData, msModelInput) = "2" Then 'otomatis
        If nNumber.Value > vaArray.UpperBound(1) Then
          SaveOk
          cKode.SetFocus
        End If
      End If
    End If
  End If
End Sub

Private Sub cmdBayar_Click()
  LockKasir True
  InitBisaBayar
  BisaBayar.Visible = True
  nSubTotal.Value = nPenjualan.Value
  nDiscountBayar.SetFocus
  HitungKembalian
End Sub

Private Sub HitungKembalian()
  nKembali.Value = nTunai.Value - (nSubTotal.Value - nDiscountBayar.Value)
End Sub

Private Sub cmdKeluar_Click()
  If MsgBox("KELUAR DARI MENU KASIR?", vbYesNo + vbInformation) = vbYes Then
    Unload Me
  Else
    cKode.SetFocus
  End If
End Sub

Private Sub cmdOK_Click()
  SaveOk
End Sub

Private Sub SaveOk()
Dim n As Integer
On Error Resume Next
    
  vaTmp(0, 0) = nNumber.Value
  vaTmp(0, 1) = cKode.Text
  vaTmp(0, 2) = cNama.Text
  vaTmp(0, 3) = nQty.Value
  vaTmp(0, 4) = cSatuan.Text
  vaTmp(0, 5) = nHarga.Value
  vaTmp(0, 6) = nJumlah.Value
  vaTmp(0, 7) = GetNull(dbData!KodeStock)
  
  If isValidOk Then
    If nNumber.Value - 1 > vaArray.UpperBound(1) Then
      nNumber.Value = vaArray.UpperBound(1) + 1
      vaArray.InsertRows nNumber.Value
    Else
      nNumber.Value = nNumber.Value - 1
    End If
    vaArray(nNumber.Value, 0) = nNumber.Value + 1
    vaArray(nNumber.Value, 1) = vaTmp(0, 1)
    vaArray(nNumber.Value, 2) = vaTmp(0, 2)
    vaArray(nNumber.Value, 3) = vaTmp(0, 3)
    vaArray(nNumber.Value, 4) = vaTmp(0, 4)
    vaArray(nNumber.Value, 5) = vaTmp(0, 5)
    vaArray(nNumber.Value, 6) = vaTmp(0, 6)
    vaArray(nNumber.Value, 7) = vaTmp(0, 7)
    Set TDBGrid1.Array = vaArray
    TDBGrid1.ReBind
    TDBGrid1.Refresh
    Initdetail
    SumTotal
  End If
  cKode.SetFocus
End Sub

Private Function isValidOk() As Boolean
isValidOk = True
  
  If vaTmp(0, 6) = 0 Then
    isValidOk = False
    Exit Function
  End If
  
  Set dbData = objData.Browse(GetDSN, "stock", , "kodestock", sisAssign, vaTmp(0, 7))
  If Not dbData.EOF Then
    If GetHargaPokok(objData, vaTmp(0, 7)) > vaTmp(0, 5) Then
      MsgBox "Stop" & vbCrLf & "Maaf. tidak bisa dilanjutkan." & vbCrLf & "Harga jual tidak sesuai, silahkan hubungi supervisor untuk penjelasan lebih lanjut." & vbCrLf & "Terimaksih"
      isValidOk = False
      Exit Function
    End If
  End If
  
  If Not GetValidDataBrowse(objData, "stock", "kodestock", vaTmp(0, 7)) Then
    MsgBox "Maaf data barang tersebut tidak ada dalam database" & vbCrLf & "Data tidak bisa disimpan"
    isValidOk = False
    Exit Function
  End If
  
  If aCfg(objData, msSaldoMinus) = 2 Then
    If GetSaldoStock(objData, "", vaTmp(0, 8)) < nQty.Value + SearchQtyInGrid(vaTmp(0, 8)) Then
      MsgBox "Maaf, stok tidak mencukupi" & vbCrLf & _
      "stok untuk barang " & cNama.Text & " hanya tersedia " & GetSaldoStock(objData, "", vaTmp(0, 8))
      isValidOk = False
    End If
  End If
  
End Function

Private Function SearchQtyInGrid(ByVal cOde As String) As Integer
Dim n As Integer
Dim nTmpQty As Integer

  nTmpQty = 0
  For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
    If vaArray(n, 8) = cOde Then
      nTmpQty = nTmpQty + vaArray(n, 3)
    End If
  Next n
  SearchQtyInGrid = nTmpQty
End Function


Private Sub cmdPersen_Click()
  Load trCalcPersen
  trCalcPersen.Show vbModal
End Sub

Private Sub cmdSimpan_Click()
Dim lSave As Boolean
Dim vaField
Dim vaValue
Dim n As Integer
Dim db As New ADODB.Recordset

lSave = True

  cFaktur = CreateNomorFaktur(objData, sisModulTransaksi.kasir, "totkasir", "nomorkasir")
  If isValidSaving Then
    
    If Not GetAvailable(cFaktur, "totkasir", "nomorkasir") Then
      cFaktur = GetNomor("totkasir", "nomorkasir", GetID, kasir)
    End If
    
    objData.Start GetDSN
    
    vaField = Array("nomorkasir", "subtotal", "discount", "total", "tunai", "username", "tgl", "datetime")
    vaValue = Array(cFaktur, nPenjualan.Value, nDiscountBayar.Value, nTotal.Value, nTunai.Value, GetRegistry(reg_UserName), Format(Date, "yyyy-MM-dd"), SNow)
    
    lSave = IIf(lSave, objData.Update(GetDSN, "totkasir", "nomorkasir = '" & cFaktur & "'", vaField, vaValue), False)

    vaField = Array("nomorkasir", "kodestock", "qty", "harga", "jumlah", "hargabeli")
    For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
      vaValue = Array(cFaktur, vaArray(n, 7), vaArray(n, 3), vaArray(n, 5), vaArray(n, 6), GetHargaBeli(objData, vaArray(n, 7)))
      lSave = IIf(lSave, objData.Add(GetDSN, "kasir", vaField, vaValue), False)
      lSave = IIf(lSave, UpdKartuStock(objData, PenjualanKasir, cFaktur, Format(Date, "yyyy-MM-dd"), vaArray(n, 7), vaArray(n, 3), vaArray(n, 5), 0, "Penjualan Kasir No. " & cFaktur, aCfg(objData, msGudangPenjualan), GetHargaPokok(objData, vaArray(n, 7))), False)
    
      Set db = objData.Browse(GetDSN, "stock", "kodestock,asbiaya", "kodestock", sisAssign, vaArray(n, 7))
      If Not db.EOF Then
        If GetNull(db!asbiaya) <> "1" Then
        
        'HP (5)
          'persediaan (1)
          
        lSave = IIf(lSave, UpdKodeTr(objData, msPenjualanKasir, cFaktur, Format(Date, "yyyy-MM-dd"), aCfg(objData, msRekeningCOGS), aCfg(objData, msCostCenterJualBeli), "Harga Pokok Penjualan Kasir No " & cFaktur, GetHargaBeli(objData, vaArray(n, 7)) * vaArray(n, 3), 0, "N", SNow), False)
            lSave = IIf(lSave, UpdKodeTr(objData, msPenjualanKasir, cFaktur, Format(Date, "yyyy-MM-dd"), GetAkunInventory(objData, vaArray(n, 7)), aCfg(objData, msCostCenterJualBeli), "Harga Pokok Penjualan Kasir No " & cFaktur, 0, GetHargaBeli(objData, vaArray(n, 7)) * vaArray(n, 3), "N"), False)
            
'        lSave = IIf(lSave, UpdKodeTr(objData, msPenjualanKasir, cFaktur, Format(Date, "yyyy-MM-dd"), aCfg(objData, msRekeningCOGS), aCfg(objData, msCostCenterJualBeli), "Harga Pokok Penjualan Kasir No " & cFaktur, GetHargaBeli(objData, vaArray(n, 7)), 0, "N", SNow), False)
'            lSave = IIf(lSave, UpdKodeTr(objData, msPenjualanKasir, cFaktur, Format(Date, "yyyy-MM-dd"), GetAkunInventory(objData, vaArray(n, 7)), aCfg(objData, msCostCenterJualBeli), "Harga Pokok Penjualan Kasir No " & cFaktur, 0, GetHargaBeli(objData, vaArray(n, 7)), "N"), False)
            
        End If
      End If
    Next n
        
    'Kas
      'Penjualan
    lSave = IIf(lSave, UpdKodeTr(objData, msPenjualanKasir, cFaktur, Format(Date, "yyyy-MM-dd"), cAkunKas.Text, aCfg(objData, msCostCenterJualBeli), "Penjualan Kasir no " & cFaktur, nTotal.Value, 0, "K", SNow), False)
        lSave = IIf(lSave, UpdKodeTr(objData, msPenjualanKasir, cFaktur, Format(Date, "yyyy-MM-dd"), aCfg(objData, msRekeningPenjualan), aCfg(objData, msCostCenterJualBeli), "Penjualan Kasir No " & cFaktur, 0, nTotal.Value, "N"), False)
    
    'discount penjualan (5)
      'penjualan
    lSave = IIf(lSave, UpdKodeTr(objData, msPenjualanKasir, cFaktur, Format(Date, "yyyy-MM-dd"), aCfg(objData, msRekeningDiscountPenjualan), aCfg(objData, msCostCenterJualBeli), "Discount Kasir no " & cFaktur, nDiscountBayar.Value, 0, "N", SNow), False)
        lSave = IIf(lSave, UpdKodeTr(objData, msPenjualanKasir, cFaktur, Format(Date, "yyyy-MM-dd"), aCfg(objData, msRekeningPenjualan), aCfg(objData, msCostCenterJualBeli), "Discount Kasir no " & cFaktur, 0, nDiscountBayar.Value, "N"), False)
      
        
    If lSave Then
      objData.Save GetDSN
      If MsgBox("Lakukan pencetakan ke Printer?", vbYesNo + vbInformation) = vbYes Then
        PrintThermal
      End If
    Else
      objData.Cancel GetDSN
    End If
    
    BisaBayar.Visible = False
    
    LockKasir False
    initvalue
    cKode.SetFocus
  End If
  
End Sub
Private Function isValidSaving() As Boolean
isValidSaving = True

  If nKembali.Value < 0 Then
    MsgBox "Maaf.Transaksi tidak bisa dilanjutkan/disimpan" & vbCrLf & "Pembayaran kurang dari total yang harus dibayar" & vbCrLf & "Tekan tombol ESC untuk kembali ke menu sebelumnya"
    nTunai.SetFocus
    isValidSaving = False
    Exit Function
  End If
  
  If nSubTotal.Value = 0 Then
    MsgBox "Maaf.Transaksi tidak bisa dilanjutkan/disimpan" & vbCrLf & "Jumlah transaksi tidak ada"
    isValidSaving = False
    Exit Function
  End If
  
  If Trim(cAkunKas.Text) = "" Then
    MsgBox "Maaf transaksi tidak bisa dilanjutkan" & vbCrLf & "Akun Kas belum di setting"
    isValidSaving = False
    Exit Function
  End If
End Function

Private Sub cNama_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "stock", "barcode,nama,kodesatuan,hargajual,kodestock", "nama", sisContent, cNama.Text, , "nama")
  If Not dbData.EOF Then
    cNama.Text = cNama.Browse(dbData)
    cNama.Text = GetNull(dbData!nama, "")
    cKode.Text = GetNull(dbData!barcode, "")
    cSatuan.Text = GetNull(dbData!kodesatuan, "")
    nHarga.Value = GetNull(dbData!hargajual, 0)
  End If
SumJumlah
End Sub

Private Sub cNama_Validate(Cancel As Boolean)
  SumJumlah
End Sub

Private Sub SumTotal()
Dim n As Single
Dim nTmpTotal As Double

  nTmpTotal = 0
  For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
    nTmpTotal = nTmpTotal + vaArray(n, 6)
  Next
  nPenjualan.Value = nTmpTotal
End Sub

Private Sub Form_Activate()
  Me.Refresh
  Me.WindowState = 2
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyF5
      cmdBayar_Click
    Case vbKeyF3
      cmdPersen_Click
    Case vbKeyEscape
      BisaBayar.Visible = False
      LockKasir False
      cKode.SetFocus
    Case vbKeyF6
      Select Case aCfg(objData, msModelInput)
        Case "1"
          UpdCfg msModelInput, "2", objData, "Mode Input", "Kasir"
          Label1.Caption = "F5 = BAYAR, F6 = MODE, MANUAL OFF"
          Exit Sub
        Case "2"
          UpdCfg msModelInput, "1", objData, "Mode Input", "Kasir"
          Label1.Caption = "F5 = BAYAR, F6 = MODE, MANUAL ON"
          Exit Sub
      End Select
  End Select
End Sub

Private Sub Form_Load()
'  If CheckTrial(nRecordsTrial, TrialKasir) = True Then
'    End
'  End If

  SetIcon Me.hwnd, "SIKD"
  Me.KeyPreview = True
  InitTabIndex
  initvalue
  If aCfg(objData, msModelInput) = "1" Then
    Label1.Caption = "F5 = BAYAR, F6 = MODE, MANUAL ON"
  Else
    Label1.Caption = "F5 = BAYAR, F6 = MODE, MANUAL OFF"
  End If
End Sub

Private Sub Form_Resize()
  Me.Refresh
  TDBGrid1.Columns(0).Width = nNumber.Width
  TDBGrid1.Columns(1).Width = cKode.Width
  TDBGrid1.Columns(2).Width = cNama.Width
  TDBGrid1.Columns(3).Width = nQty.Width + 15
  TDBGrid1.Columns(4).Width = cSatuan.Width + 15
  TDBGrid1.Columns(5).Width = nHarga.Width + 15
  TDBGrid1.Columns(6).Width = nJumlah.Width + 15
  TDBGrid1.Refresh
End Sub

Private Sub initvalue()
  nQty.Value = aCfg(objData, msQtyKasir, 1)
  BisaBayar.Visible = False
  nPenjualan.Default
  InitGrid
  Initdetail
  cAkunKas.Text = cKasTeller
  If aCfg(objData, msKolomHargaKasir) = 1 Then
    nHarga.Enabled = True
  Else
    nHarga.Enabled = False
  End If
End Sub

Private Sub Initdetail()
  nNumber.Value = vaArray.UpperBound(1) + 2
  cKode.Default
  cNama.Default
  nQty.Value = aCfg(objData, msQtyKasir, 1)
  cSatuan.Default
  nHarga.Default
  nJumlah.Default
  vaTmp.ReDim 0, 0, 0, 7
  vaTmp(0, 6) = 0
End Sub

Private Sub InitGrid()
  vaTmp.ReDim 0, -1, 0, 7
  vaArray.ReDim 0, -1, 0, 7
  TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
  TDBGrid1.Refresh
End Sub

Private Sub InitTabIndex()
Dim n As Single

  TabIndex nNumber, n
  TabIndex cKode, n
  TabIndex cNama, n
  TabIndex nQty, n
  TabIndex cSatuan, n
  TabIndex nHarga, n
  TabIndex nJumlah, n
  TabIndex cmdOK, n
  TabIndex cmdBayar, n
  TabIndex cmdKeluar, n
  
  TabIndex nDiscountBayar, n
  TabIndex nTunai, n
  TabIndex cmdSimpan, n
End Sub

Private Sub nDiscountBayar_Validate(Cancel As Boolean)
  If nDiscountBayar.Value <= nSubTotal.Value Then
    HitungKembalian
    nTotal.Value = nSubTotal.Value - nDiscountBayar.Value
  Else
    nDiscountBayar.Value = 0
    MsgBox "Jumlah diskon tidak valid"
    nDiscountBayar.SetFocus
  End If
End Sub

Private Sub nHarga_Validate(Cancel As Boolean)
  SumJumlah
End Sub

Private Sub nNumber_Validate(Cancel As Boolean)
Dim n As Single

  If GetValidNomorUrut(nNumber, vaArray) Then
    n = nNumber.Value - 1
    If n <= vaArray.UpperBound(1) Then
      
      If MsgBox("Item Mau DIKOREKSI atau DIHAPUS?" & vbCrLf & "Tekan YES untuk KOREKSI atau tekan NO untuk HAPUS", vbInformation + vbYesNo) = vbYes Then
        cKode.Text = vaArray(n, 1)
        cNama.Text = vaArray(n, 2)
        nQty.Value = vaArray(n, 3)
        cSatuan.Text = vaArray(n, 4)
        nHarga.Value = vaArray(n, 5)
        nJumlah.Value = vaArray(n, 6)
        
        vaTmp(0, 0) = nQty.Value
        vaTmp(0, 1) = vaArray(n, 1)
        vaTmp(0, 2) = vaArray(n, 2)
        vaTmp(0, 3) = vaArray(n, 3)
        vaTmp(0, 4) = vaArray(n, 4)
        vaTmp(0, 5) = vaArray(n, 5)
        vaTmp(0, 6) = vaArray(n, 6)
        vaTmp(0, 7) = vaArray(n, 7)

      Else
        vaArray.DeleteRows n
        For n = 0 To vaArray.UpperBound(1)
          vaArray(n, 0) = n + 1
        Next
        nNumber.Value = vaArray.UpperBound(1) + 2
        Set TDBGrid1.Array = vaArray
        TDBGrid1.ReBind
        TDBGrid1.Refresh
        SumTotal
      End If
    End If
  End If
End Sub

Private Sub nQty_Validate(Cancel As Boolean)
  SumJumlah
End Sub

Private Sub nTunai_Validate(Cancel As Boolean)
  HitungKembalian
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim n As Integer

  If KeyCode = vbKeyDelete Then
    If vaArray.UpperBound(1) >= 0 Then
      TDBGrid1.Delete
      For n = 0 To vaArray.UpperBound(1)
        vaArray(n, 0) = n + 1
      Next
      nNumber.Value = vaArray.UpperBound(1) + 2
      TDBGrid1.ReBind
    End If
    SumTotal
  End If
End Sub

Private Sub InitBisaBayar()
Dim n As Integer
  
  nSubTotal.Default
  nDiscountBayar.Default
  nTotal.Default
  nTunai.Default
  nKembali.Default
End Sub

Private Sub LockKasir(ByVal lStat As Boolean)
  nNumber.Enabled = Not lStat
  cKode.Enabled = Not lStat
  cNama.Enabled = Not lStat
  nQty.Enabled = Not lStat
  cmdOK.Enabled = Not lStat
  TDBGrid1.Enabled = Not lStat
  cmdBayar.Enabled = Not lStat
  cmdKeluar.Enabled = Not lStat
End Sub
