VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptBarangMasuk 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laporan Barang Masuk"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   12045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   12045
   Begin SizerOneLibCtl.ElasticOne ElasticOne3 
      Height          =   570
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   7050
      Width           =   12045
      _cx             =   21246
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
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   10860
         TabIndex        =   1
         Top             =   60
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
         Picture         =   "rptBarangMasuk.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   10410
         TabIndex        =   2
         Top             =   60
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
         Picture         =   "rptBarangMasuk.frx":00A6
      End
   End
   Begin SizerOneLibCtl.ElasticOne ElasticOne2 
      Height          =   5460
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1590
      Width           =   12045
      _cx             =   21246
      _cy             =   9631
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
      Begin TrueOleDBGrid70.TDBGrid datagrid1 
         Height          =   5400
         Left            =   30
         TabIndex        =   4
         Top             =   15
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   9525
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Tgl"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Masuk Gudang"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Kode"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Item"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Qty"
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "###,###,###,##0"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Satuan"
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Status"
         Columns(6).DataField=   ""
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Faktur"
         Columns(7).DataField=   ""
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   8
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   873
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   15790320
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2090"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2011"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=512"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=2884"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2805"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=512"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=2223"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2143"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=512"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=4419"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=4339"
         Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(25)=   "Column(4).Width=1826"
         Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1746"
         Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=514"
         Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(31)=   "Column(5).Width=1349"
         Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=1270"
         Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=512"
         Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(37)=   "Column(6).Width=2858"
         Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=2778"
         Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=512"
         Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(43)=   "Column(7).Width=3069"
         Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=2990"
         Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=512"
         Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
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
         ColumnFooters   =   -1  'True
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
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.alignment=0"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=46,.parent=13,.alignment=0"
         _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=50,.parent=13,.alignment=0"
         _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=58,.parent=13,.alignment=1"
         _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(5).Style:id=62,.parent=13,.alignment=0"
         _StyleDefs(53)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(6).Style:id=66,.parent=13,.alignment=0"
         _StyleDefs(57)  =   "Splits(0).Columns(6).HeadingStyle:id=63,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(6).FooterStyle:id=64,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(6).EditorStyle:id=65,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(7).Style:id=54,.parent=13,.alignment=0"
         _StyleDefs(61)  =   "Splits(0).Columns(7).HeadingStyle:id=51,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(7).FooterStyle:id=52,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(7).EditorStyle:id=53,.parent=17"
         _StyleDefs(64)  =   "Named:id=33:Normal"
         _StyleDefs(65)  =   ":id=33,.parent=0"
         _StyleDefs(66)  =   "Named:id=34:Heading"
         _StyleDefs(67)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(68)  =   ":id=34,.wraptext=-1"
         _StyleDefs(69)  =   "Named:id=35:Footing"
         _StyleDefs(70)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(71)  =   "Named:id=36:Selected"
         _StyleDefs(72)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(73)  =   "Named:id=37:Caption"
         _StyleDefs(74)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(75)  =   "Named:id=38:HighlightRow"
         _StyleDefs(76)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(77)  =   "Named:id=39:EvenRow"
         _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&HD3D3D3&"
         _StyleDefs(79)  =   "Named:id=40:OddRow"
         _StyleDefs(80)  =   ":id=40,.parent=33"
         _StyleDefs(81)  =   "Named:id=41:RecordSelector"
         _StyleDefs(82)  =   ":id=41,.parent=34"
         _StyleDefs(83)  =   "Named:id=42:FilterBar"
         _StyleDefs(84)  =   ":id=42,.parent=33"
      End
   End
   Begin SizerOneLibCtl.ElasticOne ElasticOne1 
      Height          =   1590
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   12045
      _cx             =   21246
      _cy             =   2805
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
         Height          =   1545
         Left            =   60
         Top             =   0
         Width           =   11910
         _ExtentX        =   21008
         _ExtentY        =   2725
         Caption         =   "K r i t e r i a         "
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
         Begin VB.CheckBox chkStock 
            Height          =   225
            Left            =   2880
            TabIndex        =   17
            Top             =   1080
            Width           =   195
         End
         Begin VB.CheckBox chkGudang 
            Height          =   225
            Left            =   1530
            TabIndex        =   15
            Top             =   1080
            Width           =   195
         End
         Begin BiSATextBoxProject.BiSABrowse cGudang 
            Height          =   330
            Left            =   1755
            TabIndex        =   14
            Top             =   1035
            Width           =   1035
            _ExtentX        =   1826
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
         Begin VB.OptionButton optStatus 
            Caption         =   "Mutasi Gudang"
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
            Index           =   3
            Left            =   4935
            TabIndex        =   12
            Top             =   735
            Width           =   1560
         End
         Begin VB.OptionButton optStatus 
            Caption         =   "Retur Penjualan"
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
            Left            =   3420
            TabIndex        =   11
            Top             =   735
            Width           =   1560
         End
         Begin VB.OptionButton optStatus 
            Caption         =   "Pembelian"
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
            Left            =   2340
            TabIndex        =   10
            Top             =   735
            Width           =   1095
         End
         Begin VB.OptionButton optStatus 
            Caption         =   "All Status"
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
            Left            =   1290
            TabIndex        =   9
            Top             =   735
            Width           =   1095
         End
         Begin BiSADateProject.BiSADate dTgl 
            Height          =   330
            Index           =   0
            Left            =   195
            TabIndex        =   6
            Top             =   270
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   582
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
            Caption         =   "Tgl"
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
            Index           =   1
            Left            =   2850
            TabIndex        =   7
            Top             =   270
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   582
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
            Caption         =   "sd"
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
         Begin BiSATextBoxProject.BiSABrowse cKodeStock 
            Height          =   330
            Left            =   3090
            TabIndex        =   16
            Top             =   1035
            Width           =   3375
            _ExtentX        =   5953
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
            Button          =   -1  'True
            Caption         =   "Stock"
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
         Begin VB.Label Label2 
            Caption         =   "Masuk Gudang"
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
            Left            =   240
            TabIndex        =   13
            Top             =   1050
            Width           =   1170
         End
         Begin VB.Label Label1 
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
            Height          =   165
            Left            =   255
            TabIndex        =   8
            Top             =   720
            Width           =   570
         End
      End
   End
End
Attribute VB_Name = "rptBarangMasuk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub cGudang_ButtonClick()
Dim db As New ADODB.Recordset

  Set db = objData.Browse(GetDSN, "gudang", "kodegudang,keterangan")
  If Not db.EOF Then
    cGudang.Text = cGudang.Browse(db)
  End If
End Sub

Private Sub chkGudang_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeysA vbKeyTab, True
  End If
End Sub

Private Sub chkStock_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeysA vbKeyTab, True
  End If
End Sub

Private Sub cKodeStock_ButtonClick()
Dim db As New ADODB.Recordset
  Set db = objData.Browse(GetDSN, "stock", "kodestock,nama,kodesatuan", "nama", sisContent, cKodeStock.Text)
  If Not db.EOF Then
    cKodeStock.Text = cKodeStock.Browse(db)
    cKodeStock.Text = GetNull(db!KodeStock, "")
  End If
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
Dim cStat As SisKartuStock

  If optStatus(1).Value = True Then
    cStat = SisKartuStock.pembelian
  ElseIf optStatus(2).Value = True Then
    cStat = SisKartuStock.returpenjualan
  ElseIf optStatus(3).Value = True Then
    cStat = SisKartuStock.MutasiKe
  End If
  GetData cStat, cGudang.Text
End Sub

Private Sub GetData(ByVal cStatus As String, ByVal cKodeGudang As String)
Dim cSQl As String
Dim n As Integer
Dim nTotal As Double

  nTotal = 0
  vaArray.ReDim 0, -1, 0, 7
  cSQl = ""
  cSQl = cSQl & " select k.tgl,k.kodegudang,k.kodestock,s.nama,k.qty,s.kodesatuan,k.status,k.nomor from kartustock k"
  cSQl = cSQl & " left join stock s on s.kodestock = k.kodestock"
  If optStatus(0).Value <> True Then
    cSQl = cSQl & " where k.status = '" & cStatus & "'"
  Else
    cSQl = cSQl & " where 1=1 "
  End If
  If chkGudang.Value <> 0 Then
    cSQl = cSQl & " and k.kodegudang = '" & cKodeGudang & "'"
  End If
  If chkStock.Value <> 0 Then
    cSQl = cSQl & " and k.kodestock = '" & cKodeStock.Text & "'"
  End If
  Set dbData = objData.Sql(GetDSN, cSQl)
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = Format(GetNull(dbData!tgl), "dd/MM/yy")
      vaArray(n, 1) = GetNull(dbData!kodegudang)
      vaArray(n, 2) = GetNull(dbData!KodeStock)
      vaArray(n, 3) = GetNull(dbData!nama)
      vaArray(n, 4) = GetNull(dbData!qty)
      vaArray(n, 5) = GetNull(dbData!kodesatuan)
      vaArray(n, 6) = GetNull(dbData!status)
      vaArray(n, 7) = GetNull(dbData!nomor)
      nTotal = nTotal + vaArray(n, 4)
      dbData.MoveNext
    Loop
  End If
  datagrid1.Columns(4).FooterText = Format(nTotal, "###,###,##0")
  Set datagrid1.Array = vaArray
  datagrid1.ReBind
  datagrid1.Refresh
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  SetIcon Me.hWnd
  TabIndex dTgl(0), n
  TabIndex dTgl(1), n
  TabIndex optStatus(0), n
  TabIndex optStatus(1), n
  TabIndex optStatus(2), n
  TabIndex optStatus(3), n
  TabIndex chkGudang, n
  TabIndex cGudang, n
  TabIndex chkStock, n
  TabIndex cKodeStock, n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
  optStatus(0).Value = True
  cGudang.Text = aCfg(objData, msGudangPembelian)
End Sub

Private Sub optStatus_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeysA vbKeyTab, True
  End If
End Sub
