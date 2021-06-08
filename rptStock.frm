VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form rptStock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daftar Inventory"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7455
   ScaleWidth      =   11625
   Begin SizerOneLibCtl.ElasticOne ElasticOne3 
      Height          =   525
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6930
      Width           =   11625
      _cx             =   20505
      _cy             =   926
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
      GridCols        =   3
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"rptStock.frx":0000
      Begin BiSAButtonProject.BiSAButton BiSAButton1 
         Height          =   495
         Left            =   15
         TabIndex        =   5
         Top             =   15
         Width           =   10365
         _ExtentX        =   18283
         _ExtentY        =   873
         Caption         =   "Excel..."
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
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   495
         Left            =   10440
         TabIndex        =   4
         Top             =   15
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   873
         Caption         =   "     &Preview"
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
         Picture         =   "rptStock.frx":004B
      End
   End
   Begin SizerOneLibCtl.ElasticOne ElasticOne2 
      Height          =   6570
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Width           =   11625
      _cx             =   20505
      _cy             =   11589
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
      Begin TrueOleDBGrid70.TDBGrid tdbgrid1 
         Height          =   6510
         Left            =   30
         TabIndex        =   3
         Top             =   15
         Width           =   11550
         _ExtentX        =   20373
         _ExtentY        =   11483
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
         Columns(2).Caption=   "BARCODE"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "INVENTORY"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "GOLONGAN"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "UNIT"
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "ON HAND"
         Columns(6).DataField=   ""
         Columns(6).NumberFormat=   "###,###,###,###"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "HARGA JUAL"
         Columns(7).DataField=   ""
         Columns(7).NumberFormat=   "###,###,###"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   8
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).AllowSizing=   -1  'True
         Splits(0).Size  =   414
         Splits(0).Size.vt=   2
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   873
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   15790320
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1164"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1058"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=2434"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2328"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=2408"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2302"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=513"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=4577"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=4471"
         Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(25)=   "Column(4).Width=3096"
         Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2990"
         Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=516"
         Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(31)=   "Column(5).Width=1614"
         Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=1508"
         Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=516"
         Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(37)=   "Column(6).Width=1984"
         Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=1879"
         Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=514"
         Splits(0)._ColumnProps(42)=   "Column(6).FetchStyle=1"
         Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(44)=   "Column(7).Width=2831"
         Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=2725"
         Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=514"
         Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=255,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=975,.italic=0"
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
         _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=70,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=67,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=68,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=69,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=2"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=50,.parent=13"
         _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=54,.parent=13"
         _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=66,.parent=13,.alignment=1"
         _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=63,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=64,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=65,.parent=17"
         _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=62,.parent=13,.alignment=1"
         _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
         _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
         _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
         _StyleDefs(66)  =   "Named:id=33:Normal"
         _StyleDefs(67)  =   ":id=33,.parent=0"
         _StyleDefs(68)  =   "Named:id=34:Heading"
         _StyleDefs(69)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(70)  =   ":id=34,.wraptext=-1"
         _StyleDefs(71)  =   "Named:id=35:Footing"
         _StyleDefs(72)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(73)  =   "Named:id=36:Selected"
         _StyleDefs(74)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(75)  =   "Named:id=37:Caption"
         _StyleDefs(76)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(77)  =   "Named:id=38:HighlightRow"
         _StyleDefs(78)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(79)  =   "Named:id=39:EvenRow"
         _StyleDefs(80)  =   ":id=39,.parent=33,.bgcolor=&HE9E9E9&"
         _StyleDefs(81)  =   "Named:id=40:OddRow"
         _StyleDefs(82)  =   ":id=40,.parent=33"
         _StyleDefs(83)  =   "Named:id=41:RecordSelector"
         _StyleDefs(84)  =   ":id=41,.parent=34"
         _StyleDefs(85)  =   "Named:id=42:FilterBar"
         _StyleDefs(86)  =   ":id=42,.parent=33"
      End
   End
   Begin SizerOneLibCtl.ElasticOne ElasticOne1 
      Height          =   360
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11625
      _cx             =   20505
      _cy             =   635
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
   End
End
Attribute VB_Name = "rptStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Gudang As String
'Public xlApp As Excel.Application ' Excel Application Object
'Public xlBook As Excel.Workbook ' Excel Workbook Object


Dim nPos As SisPos
Dim lEdit As Boolean
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB
Dim nLebarAwalKolomInventory  As Double

Public Sub SetExcel(xlFileName As String, _
    xlWorksheet As String, _
    xlCellName As String, _
    xlCellContents As String)

'    On Error GoTo SetExcel_Err 'eror handler
'
'    ' inisiasi Excel App Object
'    Set xlApp = CreateObject("Excel.Application")
'
'    ' inisiasi Excel Workbook Object.
'    Set xlBook = xlApp.Workbooks.Open(xlFileName)
'
'    ' mengisi nilai tertentu ke cell tujuan
'    xlBook.Worksheets(xlWorksheet).Range(xlCellName).Value = xlCellContents
'
'    ' menyimpan file dan menutup spreadsheet
'    xlBook.Save
'    xlBook.Close savechanges:=False
'    xlApp.Quit
'    Set xlApp = Nothing
'    Set xlBook = Nothing
'    Exit Sub
'
'SetExcel_Err:
'    MsgBox "SetExcel Error: " & err.Number & "-" & err.Description
'    Resume Next
End Sub

Private Sub BiSAButton1_Click()
    Dim a As New exportExcel
    a.RecordSource = vaArray
    a.ExportToExcel
End Sub

Private Sub cmdPreview_Click()
   With FrmRPT
      .AddPageHeader "Daftar Inventory - " & Gudang, tdbHalignCenter, , , True, dbArial, 10, , , , False
      .AddPageHeader aCfg(objData, msNamaPerusahaan), tdbHalignCenter, , , True, , 14
      .AddPageHeader "", , , , True
      .AddPageHeader "", , , , True
      
      .AddTableHeader , , , , , , , , , , , , , , , , , , , False
      .AddTableHeader "KODE", , , , 8
      .AddTableHeader "BARCODE", , , , 9
      .AddTableHeader "NAMA"
      .AddTableHeader "KET", , , , 16
      .AddTableHeader "SATUAN", , , , 8
      .AddTableHeader "SALDO STOCK", , , , 12
      .AddTableHeader "HARGA JUAL", , , , 12
      .AddTableHeader "JUMLAH", , , , 14
      
      .AddTableGroupHeader True, "[]", , , , 10 ',True, , , , 10, , , , , , , , , , , , , , , , True
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
            
            
      .AddTableBody , , , , , , , , , , , , , False
      .AddTableBody
      .AddTableBody
      .AddTableBody
      .AddTableBody
      .AddTableBody
      .AddTableBody Sis_Rpt_Number2, tdbHalignRight
      .AddTableBody Sis_Rpt_Number2, tdbHalignRight
      .AddTableBody Sis_Rpt_Number2, tdbHalignRight
      
      .AddTableFooter , , , , , , , , , , , , , , , , , , , False
      .AddTableFooter "Total", , tdbHalignCenter, , , , , , , , , , , , 7
      .AddTableFooter
      .AddTableFooter
      .AddTableFooter
      .AddTableFooter
      .AddTableFooter
      .AddTableFooter
      .AddTableFooter "&Sum", Sis_Rpt_Number2
      
'      .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
'      .AddTableGroupFooter
'      .AddTableGroupFooter
'      .AddTableGroupFooter
'      .AddTableGroupFooter
'      .AddTableGroupFooter
'      .AddTableGroupFooter
'      .AddTableGroupFooter
'      .AddTableGroupFooter

      
      .Preview vaArray, True
   End With
End Sub

Private Sub Form_Activate()
  Me.Refresh
  Me.WindowState = 2
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  GetLoadRows
  nLebarAwalKolomInventory = TDBGrid1.Columns(3).Width
End Sub

Private Sub GetLoadRows()
Dim n As Integer


  vaArray.ReDim 0, -1, 0, 8

  Set dbData = objData.Browse(GetDSN, "stock s", "s.kodegolongan,s.kodestock,s.barcode,s.nama,g.keterangan,s.kodesatuan,sum(k.debet-k.kredit) as saldostock,s.hargajual", "k.kodegudang", sisContent, Trim(Gudang), " GROUP BY s.kodestock", "s.kodegolongan,s.kodestock desc", Array("LEFT JOIN golongan g on g.kodegolongan = s.kodegolongan", "LEFT JOIN kartustock k on k.kodestock = s.kodestock"))
    If Not dbData.EOF Then
      FrmPB.InitPB dbData.RecordCount
      Do While Not dbData.EOF
        FrmPB.RunPB
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        n = vaArray.UpperBound(1)
        vaArray(n, 0) = GetNull(dbData!kodegolongan)
        vaArray(n, 1) = GetNull(dbData!KodeStock)
        vaArray(n, 2) = GetNull(dbData!barcode)
        vaArray(n, 3) = GetNull(dbData!nama)
        vaArray(n, 4) = GetNull(dbData!keterangan)
        vaArray(n, 5) = GetNull(dbData!kodesatuan)
        vaArray(n, 6) = GetNull(dbData!saldostock)
        vaArray(n, 7) = GetNull(dbData!HargaJual)
        vaArray(n, 8) = vaArray(n, 6) * vaArray(n, 7)
        If vaArray(n, 6) = 0 Then
          vaArray.DeleteRows n
        End If
        dbData.MoveNext
      Loop
      FrmPB.EndPB
    End If
  
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
  TDBGrid1.Refresh
End Sub

Private Sub Form_Resize()
Dim nSisaLebar As Double

    Me.Refresh
    nSisaLebar = TDBGrid1.Width - TDBGrid1.Columns(0).Width - TDBGrid1.Columns(1).Width - TDBGrid1.Columns(2).Width - TDBGrid1.Columns(4).Width - TDBGrid1.Columns(5).Width - TDBGrid1.Columns(6).Width - TDBGrid1.Columns(7).Width
    If Me.WindowState = 2 Then
      TDBGrid1.Columns(3).Width = nSisaLebar - 1000
    Else
      TDBGrid1.Columns(3).Width = nLebarAwalKolomInventory
    End If
End Sub

Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
  TDBGrid1.Update
End Sub

Private Sub TDBGrid1_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueOleDBGrid70.StyleDisp)
On Error Resume Next

  Select Case Col
      Case 6
          Dim Col2 As Long
          Col2 = CLng(TDBGrid1.Columns(6).CellText(Bookmark))
          If Col2 < 0 Then CellStyle.ForeColor = vbRed
  End Select
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    GetLoadRows
  End If
End Sub


