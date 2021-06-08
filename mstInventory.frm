VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form mstInventory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entry Data"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   9555
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   8115
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9555
      _cx             =   16854
      _cy             =   14314
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
      Appearance      =   2
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   -2147483633
      TabOutlineColor =   0
      FrontTabForeColor=   -2147483630
      Caption         =   "&Entry Data...|&List Data"
      Align           =   5
      CurrTab         =   1
      FirstTab        =   0
      Style           =   4
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
      Begin SizerOneLibCtl.ElasticOne ElasticOne2 
         Height          =   7740
         Left            =   45
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   330
         Width           =   9465
         _cx             =   16695
         _cy             =   13653
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
         Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
            Height          =   7635
            Left            =   45
            TabIndex        =   18
            Top             =   45
            Width           =   10860
            _ExtentX        =   19156
            _ExtentY        =   13467
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "KODE"
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "BARCODE"
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "INVENTORY"
            Columns(2).DataField=   ""
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "GOLONGAN"
            Columns(3).DataField=   ""
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "SATUAN"
            Columns(4).DataField=   ""
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   5
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   873
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   15790320
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=5"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2566"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2487"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=2990"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2910"
            Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=5345"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=5265"
            Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(16)=   "Column(3).Width=3307"
            Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=3228"
            Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(21)=   "Column(4).Width=2328"
            Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2249"
            Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   0
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            Appearance      =   0
            DataMode        =   4
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=255,.bold=0,.fontsize=825,.italic=0"
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
            _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(52)  =   "Named:id=33:Normal"
            _StyleDefs(53)  =   ":id=33,.parent=0"
            _StyleDefs(54)  =   "Named:id=34:Heading"
            _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(56)  =   ":id=34,.wraptext=-1"
            _StyleDefs(57)  =   "Named:id=35:Footing"
            _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(59)  =   "Named:id=36:Selected"
            _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(61)  =   "Named:id=37:Caption"
            _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(63)  =   "Named:id=38:HighlightRow"
            _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(65)  =   "Named:id=39:EvenRow"
            _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(67)  =   "Named:id=40:OddRow"
            _StyleDefs(68)  =   ":id=40,.parent=33"
            _StyleDefs(69)  =   "Named:id=41:RecordSelector"
            _StyleDefs(70)  =   ":id=41,.parent=34"
            _StyleDefs(71)  =   "Named:id=42:FilterBar"
            _StyleDefs(72)  =   ":id=42,.parent=33"
         End
      End
      Begin SizerOneLibCtl.ElasticOne ElasticOne1 
         Height          =   7740
         Left            =   -10110
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   330
         Width           =   9465
         _cx             =   16695
         _cy             =   13653
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
         Begin BiSAFramProject.BiSAFrame BiSAFrame1 
            Height          =   7110
            Left            =   15
            Top             =   0
            Width           =   9465
            _ExtentX        =   16695
            _ExtentY        =   12541
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
            Begin BiSANumberBoxProject.BiSANumberBox nHargaBeli 
               Height          =   330
               Left            =   105
               TabIndex        =   2
               Top             =   1905
               Width           =   2940
               _ExtentX        =   5186
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
               Caption         =   "Harga Beli"
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
               Left            =   105
               TabIndex        =   3
               Top             =   840
               Width           =   5670
               _ExtentX        =   10001
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
            Begin BiSATextBoxProject.BiSABrowse cSatuan 
               Height          =   330
               Left            =   105
               TabIndex        =   4
               Top             =   1155
               Width           =   2235
               _ExtentX        =   3942
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
               Caption         =   "Satuan"
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
            Begin BiSATextBoxProject.BiSATextBox cNamaSatuan 
               Height          =   330
               Left            =   2355
               TabIndex        =   5
               Top             =   1155
               Width           =   2775
               _ExtentX        =   4895
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
            Begin BiSATextBoxProject.BiSABrowse cGolongan 
               Height          =   330
               Left            =   105
               TabIndex        =   6
               Top             =   1470
               Width           =   2850
               _ExtentX        =   5027
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
               Caption         =   "Golongan"
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
            Begin BiSATextBoxProject.BiSATextBox cNamaGolongan 
               Height          =   330
               Left            =   2955
               TabIndex        =   7
               Top             =   1470
               Width           =   2775
               _ExtentX        =   4895
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
            Begin BiSANumberBoxProject.BiSANumberBox nHargaJual 
               Height          =   330
               Left            =   105
               TabIndex        =   8
               Top             =   2220
               Width           =   2940
               _ExtentX        =   5186
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
               Caption         =   "Harga Jual"
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
            Begin BiSATextBoxProject.BiSATextBox cKode 
               Height          =   330
               Left            =   105
               TabIndex        =   9
               Top             =   210
               Width           =   3735
               _ExtentX        =   6588
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
               Caption         =   "Kode"
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
               Left            =   105
               TabIndex        =   10
               Top             =   525
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
               GetPicture      =   1
               Button          =   -1  'True
               Caption         =   "Barcode"
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
            Left            =   15
            Top             =   7095
            Width           =   9465
            _ExtentX        =   16695
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
            Begin BiSAButtonProject.BiSAButton cmdHapus 
               Height          =   435
               Left            =   2220
               TabIndex        =   11
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
               Picture         =   "mstInventory.frx":0000
            End
            Begin BiSAButtonProject.BiSAButton cmdAktivasi 
               Height          =   435
               Left            =   3390
               TabIndex        =   12
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
               Picture         =   "mstInventory.frx":028A
            End
            Begin BiSAButtonProject.BiSAButton cmdEdit 
               Height          =   435
               Left            =   1170
               TabIndex        =   13
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
               Picture         =   "mstInventory.frx":0429
            End
            Begin BiSAButtonProject.BiSAButton cmdAdd 
               Height          =   435
               Left            =   105
               TabIndex        =   14
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
               Picture         =   "mstInventory.frx":0555
            End
            Begin BiSAButtonProject.BiSAButton cmdKeluar 
               Cancel          =   -1  'True
               Height          =   435
               Left            =   8160
               TabIndex        =   15
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
               Picture         =   "mstInventory.frx":0700
            End
            Begin BiSAButtonProject.BiSAButton cmdSimpan 
               Height          =   435
               Left            =   6945
               TabIndex        =   16
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
               Picture         =   "mstInventory.frx":07A6
            End
         End
      End
   End
End
Attribute VB_Name = "mstInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB
Dim lEdit As Boolean
Dim nPos As SisPos

Private Sub GetEdit(lPar As Boolean)
  lEdit = lPar
  BiSAFrame1.Enabled = lPar
  SetButton cmdSimpan, cmdKeluar, cmdAdd, cmdEdit, cmdHapus, nPos, lPar, cmdAktivasi
End Sub

Private Sub GetLoadRows()
Dim n As Integer

  vaArray.ReDim 0, -1, 0, 4
  Set dbData = objData.Browse(GetDSN, "stock s", "s.kodestock,s.barcode,s.nama,g.keterangan,s.kodesatuan", "s.kodestock", sisContent, tdbgrid1.Columns(0).FilterText, " AND barcode LIKE '%" & tdbgrid1.Columns(1).FilterText & "%' AND nama LIKE '%" & tdbgrid1.Columns(2).FilterText & "%' AND g.keterangan LIKE '%" & tdbgrid1.Columns(3).FilterText & "%' AND s.kodesatuan LIKE '%" & tdbgrid1.Columns(4).FilterText & "%'", "s.kodestock desc", Array("LEFT JOIN golongan g on g.kodegolongan = s.kodegolongan"), 0, 100)
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!KodeStock)
      vaArray(n, 1) = GetNull(dbData!Barcode)
      vaArray(n, 2) = GetNull(dbData!nama)
      vaArray(n, 3) = GetNull(dbData!Keterangan)
      vaArray(n, 4) = GetNull(dbData!kodesatuan)
      dbData.MoveNext
    Loop
  End If
  Set tdbgrid1.Array = vaArray
  tdbgrid1.ReBind
  tdbgrid1.Refresh
End Sub

Private Sub cBarcode_ButtonClick()
Set dbData = objData.Browse(GetDSN, "stock", "barcode,kodestock,nama", "barcode", sisContent, cBarcode.Text, , "nama")
If Not dbData.EOF Then
  cBarcode.Text = cBarcode.Browse(dbData)
  cKode.Text = GetNull(dbData!KodeStock, "")
  GetMemory
End If
End Sub

Private Sub cBarcode_Validate(Cancel As Boolean)
  If nPos = Add Then
    If Trim(cBarcode.Text) <> "" Then
      Set dbData = objData.Browse(GetDSN, "stock", "barcode", "barcode", sisAssign, cBarcode.Text)
      If Not dbData.EOF Then
        If MsgBox("Peringatan!!" & vbCrLf & "Barcode yg ingin anda masukkan sudah terpakai sebelumnya, silahkan gunakan Kode Barcode yg lain." & vbCrLf & "Apakah akan dilanjutkan?" & vbCrLf & "Jika Ya, maka akan ada Kode Barcode yg sama dalam komputer.", vbYesNo + vbInformation) = vbNo Then
          cBarcode.Text = ""
        End If
      End If
    End If
  End If
End Sub

Private Sub cGolongan_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "golongan", "kodegolongan,keterangan", "kodegolongan", sisContent, cGolongan.Text, , "kodegolongan,keterangan")
  If Not dbData.EOF Then
    If dbData.RecordCount > 1 Then
      cGolongan.Text = cGolongan.Browse(dbData)
    End If
    cGolongan.Text = GetNull(dbData!kodegolongan)
    cNamaGolongan.Text = GetNull(dbData!Keterangan, "")
  Else
    cGolongan.Default
    cNamaGolongan.Default
  End If
End Sub

Private Function ValidKode() As Boolean
  ValidKode = True
End Function

Private Sub GetMemory()
Dim n As Single
  
  Set dbData = objData.Browse(GetDSN, "stock s", "s.nama,s.kodesatuan,sa.keterangan as keterangansatuan,s.hargabeli,s.hargajual,s.kodegolongan,ga.keterangan as keterangangolongan", "s.kodestock", sisAssign, cKode.Text, , , Array("left join satuan sa on sa.kodesatuan = s.kodesatuan", "left join golongan ga on ga.kodegolongan = s.kodegolongan"))
  If Not dbData.EOF Then
    cNama.Text = GetNull(dbData!nama, "")
    cSatuan.Text = GetNull(dbData!kodesatuan, "")
    cNamaSatuan.Text = GetNull(dbData!keterangansatuan, "")
    nHargaBeli.Value = GetNull(dbData!hargabeli)
    nHargaJual.Value = GetNull(dbData!hargajual)
    cGolongan.Text = GetNull(dbData!kodegolongan, "")
    cNamaGolongan.Text = GetNull(dbData!KeteranganGolongan, "")
    If nPos = Delete Then
      DeleteStock
    End If
  End If
End Sub

Private Sub cmdAdd_Click()
  GetEdit True
  InitValue
  cBarcode.SetFocus
  nPos = Add
End Sub

Private Sub cmdAktivasi_Click()
  frmAktivasi.Action Me
End Sub

Private Sub cmdEdit_Click()
  GetEdit True
  cBarcode.SetFocus
  nPos = Edit
End Sub

Private Sub cmdHapus_Click()
  GetEdit True
  nPos = Delete
  cBarcode.SetFocus
  DeleteStock
  GetLoadRows
End Sub

Private Sub DeleteStock()
Dim cInfo As String
  
  cInfo = "Kode: " & cKode.Text & vbCrLf
  cInfo = cInfo & "Nama: " & cNama.Text & vbCrLf
  cInfo = cInfo & "Satuan: " & cSatuan.Text & vbCrLf
  cInfo = cInfo & "Golongan: " & cNamaGolongan.Text & vbCrLf
  cInfo = cInfo & "Harga Beli: " & Format(nHargaBeli.Value, "###,###,###,##0.00") & vbCrLf
  cInfo = cInfo & "Harga Jual: " & Format(nHargaJual.Value, "###,###,###,##0.00") & vbCrLf
  
  If MsgBox("Data Benar-benar Dihapus ?" & vbCrLf & vbCrLf & cInfo, vbQuestion + vbYesNo) = vbYes Then
    objData.Start GetDSN
    objData.Delete GetDSN, "stock", "kodestock", sisAssign, cKode.Text
    objData.Save GetDSN
  End If
  InitValue
  GetEdit False
End Sub

Private Sub cmdKeluar_Click()
  If lEdit Then
    GetEdit False
    InitValue
  Else
    Unload Me
  End If
End Sub

Private Sub cmdSimpan_Click()
Dim vaField, vaValue
Dim nMin As Double
Dim nMax As Double
Dim n As Single
Dim cFaktur As String
Dim dTgl As Date
Dim lSave As Boolean
lSave = True

  If ValidSaving Then
    objData.Start GetDSN
    
    ' Simpan table Stock
    vaField = Array("nama", "barcode", "kodegolongan", _
                    "kodesatuan", _
                    "hargabeli", "hargajual")
    vaValue = Array(cNama.Text, cBarcode.Text, cGolongan.Text, _
                    cSatuan.Text, _
                    nHargaBeli.Value, nHargaJual.Value)
    
    If Trim(cKode.Text) <> "" Then
      lSave = IIf(lSave, objData.Update(GetDSN, "stock", "kodestock = '" & cKode.Text & "'", vaField, vaValue), False)
    Else
      lSave = IIf(lSave, objData.Add(GetDSN, "stock", vaField, vaValue), False)
    End If
    
    If Not lSave Then
      objData.Cancel GetDSN
    Else
      objData.Save GetDSN
    End If
    
    GetEdit False
    InitValue
    GetLoadRows
  End If
End Sub

Private Function ValidSaving() As Boolean
  ValidSaving = True
End Function

Private Sub cSatuan_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "satuan", "kodesatuan,keterangan", "kodesatuan", sisContent, cSatuan.Text, , "kodesatuan,keterangan")
  If Not dbData.EOF Then
    If dbData.RecordCount > 1 Then
      cSatuan.Text = cSatuan.Browse(dbData)
    End If
    cSatuan.Text = GetNull(dbData!kodesatuan)
    cNamaSatuan.Text = GetNull(dbData!Keterangan, "")
  Else
    cSatuan.Default
    cNamaSatuan.Default
  End If
End Sub

Private Sub Form_Load()
Dim n As Single
  
  SetIcon Me.hWnd, "SIKD"
  GetEdit False
  CenterForm Me, True
  InitValue
  
  TabIndex cKode, n
  TabIndex cBarcode, n
  TabIndex cNama, n
  TabIndex cSatuan, n
  TabIndex cNamaSatuan, n
  TabIndex cGolongan, n
  TabIndex cNamaGolongan, n
  TabIndex nHargaBeli, n
  TabIndex nHargaJual, n
  
  'optional
  TabIndex cmdAdd, n
  TabIndex cmdEdit, n
  TabIndex cmdHapus, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  TabIndex cmdAktivasi, n
  GetLoadRows
End Sub

Private Sub InitValue()
  cKode.Default
  cBarcode.Default
  cNama.Default
  cSatuan.Default
  cNamaSatuan.Default
  cGolongan.Default
  cNamaGolongan.Default
  nHargaBeli.Default
  nHargaJual.Default
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    GetLoadRows
  End If
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  Set dbData = objData.Browse(GetDSN, "stock s", "s.kodestock,s.barcode,s.nama,s.kodesatuan,sa.keterangan as keterangansatuan,s.hargabeli,s.hargajual,s.kodegolongan,ga.keterangan as keterangangolongan", "s.kodestock", sisAssign, tdbgrid1.Columns(0).Text, , , Array("left join satuan sa on sa.kodesatuan = s.kodesatuan", "left join golongan ga on ga.kodegolongan = s.kodegolongan"))
  If Not dbData.EOF Then
    cKode.Text = GetNull(dbData!KodeStock)
    cBarcode.Text = GetNull(dbData!Barcode)
    cNama.Text = GetNull(dbData!nama, "")
    cSatuan.Text = GetNull(dbData!kodesatuan, "")
    cNamaSatuan.Text = GetNull(dbData!keterangansatuan, "")
    nHargaBeli.Value = GetNull(dbData!hargabeli)
    nHargaJual.Value = GetNull(dbData!hargajual)
    cGolongan.Text = GetNull(dbData!kodegolongan, "")
    cNamaGolongan.Text = GetNull(dbData!KeteranganGolongan, "")
  End If
End Sub


