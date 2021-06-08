VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptPenjualanGroupByCustomer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laporan Penjualan Group By Customer"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11115
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   11115
   Begin SizerOneLibCtl.ElasticOne ElasticOne3 
      Height          =   570
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6735
      Width           =   11115
      _cx             =   19606
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
      Begin BiSAButtonProject.BiSAButton BiSAButton5 
         Height          =   390
         Left            =   7200
         TabIndex        =   17
         Top             =   135
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   688
         Caption         =   "Label1"
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
      Begin BiSAButtonProject.BiSAButton BiSAButton3 
         Height          =   330
         Left            =   5940
         TabIndex        =   15
         Top             =   120
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   582
         Caption         =   "F1"
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
      Begin BiSAButtonProject.BiSAButton BiSAButton2 
         Height          =   330
         Left            =   4860
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   120
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   582
         Caption         =   "Lookup"
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
      Begin BiSAButtonProject.BiSAButton cmdCetak 
         Height          =   435
         Left            =   9090
         TabIndex        =   12
         Top             =   75
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   767
         Caption         =   "Cetak"
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
         Height          =   405
         Left            =   45
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   105
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   714
         Caption         =   "Rekapitulasi"
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
      Begin BiSAButtonProject.BiSAButton cmdExportToExcel 
         Height          =   435
         Left            =   8445
         TabIndex        =   10
         Top             =   75
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   767
         Caption         =   "Exl"
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
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   9975
         TabIndex        =   6
         Top             =   75
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
         Picture         =   "rptPenjualanGroupByCustomer.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   7950
         TabIndex        =   7
         Top             =   75
         Width           =   465
         _ExtentX        =   820
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
         Picture         =   "rptPenjualanGroupByCustomer.frx":00A6
      End
      Begin BiSATextBoxProject.BiSABrowse cCustomer 
         Height          =   330
         Left            =   2175
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   120
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   582
         Text            =   "12345678"
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
         Caption         =   "Member"
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
      Begin BiSAButtonProject.BiSAButton BiSAButton4 
         Height          =   330
         Left            =   6510
         TabIndex        =   16
         Top             =   120
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   582
         Caption         =   "F2"
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
   End
   Begin SizerOneLibCtl.ElasticOne ElasticOne2 
      Height          =   5670
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1065
      Width           =   11115
      _cx             =   19606
      _cy             =   10001
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
         Height          =   5625
         Left            =   0
         TabIndex        =   5
         Top             =   15
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   9922
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
         Columns(1).Caption=   "Nama"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Bln H-1"
         Columns(2).DataField=   ""
         Columns(2).NumberFormat=   "###,###,###,##0.00"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Bln H"
         Columns(3).DataField=   ""
         Columns(3).NumberFormat=   "###,###,###,##0.00"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Total H + (H-1)"
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "###,###,###,##0.00"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Bonus H-1"
         Columns(6).DataField=   ""
         Columns(6).NumberFormat=   "###,###,##0.00"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   873
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   15790320
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1270"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1191"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=514"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=6350"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=6271"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=3175"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=3096"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=514"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=3228"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=3149"
         Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=514"
         Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(25)=   "Column(4).Width=2699"
         Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2619"
         Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=514"
         Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(31)=   "Column(5).Width=1693"
         Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=1614"
         Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=514"
         Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(37)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=514"
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
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=1"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=54,.parent=13,.alignment=3"
         _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=1"
         _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=1"
         _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=1"
         _StyleDefs(53)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.alignment=1"
         _StyleDefs(57)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(60)  =   "Named:id=33:Normal"
         _StyleDefs(61)  =   ":id=33,.parent=0"
         _StyleDefs(62)  =   "Named:id=34:Heading"
         _StyleDefs(63)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(64)  =   ":id=34,.wraptext=-1"
         _StyleDefs(65)  =   "Named:id=35:Footing"
         _StyleDefs(66)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(67)  =   "Named:id=36:Selected"
         _StyleDefs(68)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(69)  =   "Named:id=37:Caption"
         _StyleDefs(70)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(71)  =   "Named:id=38:HighlightRow"
         _StyleDefs(72)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(73)  =   "Named:id=39:EvenRow"
         _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&HD3D3D3&"
         _StyleDefs(75)  =   "Named:id=40:OddRow"
         _StyleDefs(76)  =   ":id=40,.parent=33"
         _StyleDefs(77)  =   "Named:id=41:RecordSelector"
         _StyleDefs(78)  =   ":id=41,.parent=34"
         _StyleDefs(79)  =   "Named:id=42:FilterBar"
         _StyleDefs(80)  =   ":id=42,.parent=33"
      End
   End
   Begin SizerOneLibCtl.ElasticOne ElasticOne1 
      Height          =   1065
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11115
      _cx             =   19606
      _cy             =   1879
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
      Begin VB.OptionButton Option1 
         Caption         =   "Harga Member"
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
         Left            =   2835
         TabIndex        =   9
         Top             =   690
         Width           =   1560
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Harga Katalog"
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
         Left            =   1275
         TabIndex        =   8
         Top             =   675
         Value           =   -1  'True
         Width           =   1560
      End
      Begin BiSADateProject.BiSADate dTgl 
         Height          =   330
         Index           =   0
         Left            =   180
         TabIndex        =   3
         Top             =   225
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
         Left            =   2835
         TabIndex        =   4
         Top             =   225
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
   End
End
Attribute VB_Name = "rptPenjualanGroupByCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB
Dim lClick As Boolean
Dim nBatasTutupPoin As Double
Dim cTex As String

Private Sub BiSAButton1_Click()
Dim cSQL As String
Dim db As New ADODB.Recordset
Dim n As Single


  cSQL = cSQL & " select a.telp,t.kodeanggota,a.nama,a.kodeupline,sum(jumlah) as jumlahtotal from penjualan p"
  cSQL = cSQL & " left join stock s on s.kodestock = p.kodestock"
  cSQL = cSQL & " left join totpenjualan t on t.nomorpenjualan = p.nomorpenjualan"
  cSQL = cSQL & " left join anggota a on a.kodeanggota = t.kodeanggota"
  
  cSQL = cSQL & " Where"
  cSQL = cSQL & " p.tgl >= '" & Format(dTgl(0).Value, "yyyy-MM-dd") & "' and p.tgl <= '" & Format(dTgl(1).Value, "yyyy-MM-dd") & "'"
  cSQL = cSQL & " and s.diskonpenjualan >= 30"
  
  cSQL = cSQL & " GROUP BY t.kodeanggota"
  cSQL = cSQL & " ORDER BY jumlahtotal DESC"
  
  vaArray.ReDim 0, 1, 0, 10
  Set db = objData.SQL(GetDSN, cSQL)
  If Not db.EOF Then
    FrmPB.InitPB db.RecordCount
    
    vaArray(0, 0) = "Rekap Penjualan Leader/Member Tgl " & Format(dTgl(0).Value, "dd MMMM yyyy") & " sd " & Format(dTgl(1).Value, "dd MMMM yyyy")
    vaArray(0, 1) = ""
    vaArray(0, 2) = ""
    vaArray(0, 3) = ""
    vaArray(0, 4) = ""
    vaArray(0, 5) = ""
    vaArray(0, 6) = ""
    vaArray(0, 7) = ""
    vaArray(0, 8) = ""
    vaArray(0, 9) = ""
    vaArray(0, 10) = ""
    
    vaArray(1, 0) = "NOHP"
    vaArray(1, 1) = "ID"
    vaArray(1, 2) = "Nama"
    vaArray(1, 3) = "Netto Reguler"
    vaArray(1, 4) = "TPG Reguler"
    vaArray(1, 5) = "Netto Promo"
    vaArray(1, 6) = "Piutang"
    vaArray(1, 7) = "Top Up"
    vaArray(1, 8) = "TPG Penjualan sd L4"
    vaArray(1, 9) = "TPG Penjualan L1"
    vaArray(1, 10) = "Downline"
    
  
    Do While Not db.EOF
      FrmPB.RunPB
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = "'" & GetNull(db!telp)
      vaArray(n, 1) = "'" & GetNull(db!kodeanggota)
      vaArray(n, 2) = GetNull(db!nama) & " < " & GetNamaUpline(objData, GetNull(db!kodeupline))
      vaArray(n, 3) = Format(GetNull(db!jumlahtotal), "###,###,##0")
      vaArray(n, 4) = Format(GetNull(db!jumlahtotal) / 0.7, "###,###,##0")
      vaArray(n, 5) = Format(GetPembelanjaanPromo(GetNull(db!kodeanggota)), "###,###,##0")
      vaArray(n, 6) = Format(GetSaldoPiutangDanTopUp(GetNull(db!kodeanggota)), "###,###,##0")
      vaArray(n, 7) = Format(GetSaldoPiutangDanTopUp2(GetNull(db!kodeanggota)), "###,###,##0")
      vaArray(n, 8) = Format(GetPenjualanLevel4(objData, GetNull(db!kodeanggota)), "###,###,##0")
      vaArray(n, 9) = Format(GetPenjualanLevel1(objData, GetNull(db!kodeanggota)), "###,###,##0")
      vaArray(n, 10) = Format(GetPenjualanLevel12(objData, GetNull(db!kodeanggota)), "###,###,##0")
      db.MoveNext
    Loop
    FrmPB.EndPB
  End If
  
    Dim a As New exportExcel
    a.RecordSource = vaArray
    a.ExportToExcel
  
End Sub

Private Function GetNamaUpline(ByVal obj As CodeSuiteLibrary.Data, ByVal cUpline As String) As String
Dim db As New ADODB.Recordset

  GetNamaUpline = ""
  Set db = obj.Browse(GetDSN, "anggota", , "kodeanggota", sisAssign, cUpline)
  If Not db.EOF Then
    GetNamaUpline = GetNull(db!nama) & "(" & Format(GetPenjualanMember(obj, cUpline), "###,###,##0") & ")"
  End If
End Function


Private Function GetPenjualanLevel1(ByVal obj As CodeSuiteLibrary.Data, ByVal cKodeUpline As String) As Double
Dim cSQL As String
Dim db As New ADODB.Recordset
Dim nTotalPenjualan As Double

  nTotalPenjualan = 0
  cSQL = "select * from anggota"
  cSQL = cSQL & " Where kodeupline = '" & cKodeUpline & "'"
  Set db = obj.SQL(GetDSN, cSQL)
  If Not db.EOF Then
    Do While Not db.EOF
      nTotalPenjualan = nTotalPenjualan + GetPenjualanMemberLevel1(obj, GetNull(db!kodeanggota))
      db.MoveNext
    Loop
  End If
  GetPenjualanLevel1 = nTotalPenjualan
End Function

Private Function GetPenjualanLevel4(ByVal obj As CodeSuiteLibrary.Data, ByVal cKodeUpline As String) As Double
Dim cSQL As String
Dim db, db2, db3, db4 As New ADODB.Recordset
Dim nTotalPenjualan As Double

  nTotalPenjualan = 0
  cSQL = "select * from anggota"
  cSQL = cSQL & " Where kodeupline = '" & cKodeUpline & "'"
  Set db = obj.SQL(GetDSN, cSQL)
  If Not db.EOF Then
    Do While Not db.EOF
      If GetPenjualanMember(obj, GetNull(db!kodeanggota)) > nBatasTutupPoin Then
        Set db2 = obj.SQL(GetDSN, cSQLevel(GetNull(db!kodeanggota)))
        If Not db2.EOF Then
          Do While Not db2.EOF
            If GetPenjualanMember(obj, GetNull(db2!kodeanggota)) > nBatasTutupPoin Then
              Set db3 = obj.SQL(GetDSN, cSQLevel(GetNull(db2!kodeanggota)))
              If Not db3.EOF Then
                Do While Not db3.EOF
                    If GetPenjualanMember(obj, GetNull(db3!kodeanggota)) > nBatasTutupPoin Then
                    Set db4 = obj.SQL(GetDSN, cSQLevel(GetNull(db3!kodeanggota)))
                    If Not db4.EOF Then
                      Do While Not db4.EOF
                        nTotalPenjualan = nTotalPenjualan + GetPenjualanMember(obj, GetNull(db4!kodeanggota))
                        db4.MoveNext
                      Loop
                    End If
                  End If
                  nTotalPenjualan = nTotalPenjualan + GetPenjualanMember(obj, GetNull(db3!kodeanggota))
                  db3.MoveNext
                Loop
              End If
            End If
            nTotalPenjualan = nTotalPenjualan + GetPenjualanMember(obj, GetNull(db2!kodeanggota))
            db2.MoveNext
          Loop
        End If
      End If
      nTotalPenjualan = nTotalPenjualan + GetPenjualanMember(obj, GetNull(db!kodeanggota))
      db.MoveNext
    Loop
  End If
  GetPenjualanLevel4 = nTotalPenjualan
End Function

Private Function GetPenjualanL4Detail(ByVal obj As CodeSuiteLibrary.Data, ByVal cKodeUpline As String) As String
Dim cSQL As String
Dim db, db2, db3, db4 As New ADODB.Recordset
Dim nTotalPenjualan As Double
Dim cTextL1, cTextL2, cTextL3, cTextL4 As String

  nTotalPenjualan = 0
  cSQL = "select * from anggota"
  cSQL = cSQL & " Where kodeupline = '" & cKodeUpline & "'"
  Set db = obj.SQL(GetDSN, cSQL)
  If Not db.EOF Then
    Do While Not db.EOF
      If GetPenjualanMember(obj, GetNull(db!kodeanggota)) > nBatasTutupPoin Then
        cTextL1 = cTextL1 & vbTab & "(L1) " & GetNull(db!nama) & " (" & Format(GetPenjualanMember(obj, GetNull(db!kodeanggota)), "###,###,##0") & ")" & vbCrLf
        Set db2 = obj.SQL(GetDSN, cSQLevel(GetNull(db!kodeanggota)))
        If Not db2.EOF Then
          Do While Not db2.EOF
            If GetPenjualanMember(obj, GetNull(db2!kodeanggota)) > nBatasTutupPoin Then
              cTextL1 = cTextL1 & vbTab & vbTab & "(L2) " & GetNull(db2!nama) & " (" & Format(GetPenjualanMember(obj, GetNull(db2!kodeanggota)), "###,###,##0") & ")" & vbCrLf
              Set db3 = obj.SQL(GetDSN, cSQLevel(GetNull(db2!kodeanggota)))
              If Not db3.EOF Then
                Do While Not db3.EOF
                    If GetPenjualanMember(obj, GetNull(db3!kodeanggota)) > nBatasTutupPoin Then
                    cTextL1 = cTextL1 & vbTab & vbTab & vbTab & "(L3) " & GetNull(db3!nama) & " (" & Format(GetPenjualanMember(obj, GetNull(db3!kodeanggota)), "###,###,##0") & ")" & vbCrLf
                    Set db4 = obj.SQL(GetDSN, cSQLevel(GetNull(db3!kodeanggota)))
                    If Not db4.EOF Then
                      Do While Not db4.EOF
                        cTextL1 = cTextL1 & vbTab & vbTab & vbTab & vbTab & "(L4) " & GetNull(db4!nama) & " (" & Format(GetPenjualanMember(obj, GetNull(db4!kodeanggota)), "###,###,##0") & ")" & vbCrLf
                        db4.MoveNext
                      Loop
                    End If
                  End If
                  db3.MoveNext
                Loop
              End If
            End If
            db2.MoveNext
          Loop
        End If
      End If
      db.MoveNext
    Loop
  End If
  GetPenjualanL4Detail = cTextL1
End Function

Private Function GetPenjualanLevel42(ByVal obj As CodeSuiteLibrary.Data, ByVal cKodeUpline As String) As Double
Dim cSQL As String
Dim db, db2, db3, db4 As New ADODB.Recordset
Dim nTotalPenjualan As Double

  nTotalPenjualan = 0
  cSQL = "select * from anggota"
  cSQL = cSQL & " Where kodeupline = '" & cKodeUpline & "'"
  Set db = obj.SQL(GetDSN, cSQL)
  If Not db.EOF Then
    Do While Not db.EOF
      
      Set db2 = obj.SQL(GetDSN, cSQLevel(GetNull(db!kodeanggota)))
      If Not db2.EOF Then
        Do While Not db2.EOF
          Set db3 = obj.SQL(GetDSN, cSQLevel(GetNull(db2!kodeanggota)))
          If Not db3.EOF Then
            Do While Not db3.EOF
              Set db4 = obj.SQL(GetDSN, cSQLevel(GetNull(db3!kodeanggota)))
              If Not db4.EOF Then
                Do While Not db4.EOF
                  nTotalPenjualan = nTotalPenjualan + GetPenjualanMember(obj, GetNull(db4!kodeanggota))
                  db4.MoveNext
                Loop
              End If
              nTotalPenjualan = nTotalPenjualan + GetPenjualanMember(obj, GetNull(db3!kodeanggota))
              db3.MoveNext
            Loop
          End If
          nTotalPenjualan = nTotalPenjualan + GetPenjualanMember(obj, GetNull(db2!kodeanggota))
          db2.MoveNext
        Loop
      End If
      
      nTotalPenjualan = nTotalPenjualan + GetPenjualanMember(obj, GetNull(db!kodeanggota))
      db.MoveNext
    Loop
  End If
  GetPenjualanLevel42 = nTotalPenjualan
End Function

Private Function cSQLevel(ByVal cKodeSQLanggota As String) As String
  cSQLevel = "SELECT * FROM anggota WHERE kodeupline = '" & cKodeSQLanggota & "'"
End Function

Private Function GetPenjualanMember(ByVal obj As CodeSuiteLibrary.Data, ByVal cMember As String) As Double
Dim db As New ADODB.Recordset
Dim cSQL As String

  GetPenjualanMember = 0
  'cSQL = "select sum(p.jumlah) as jumlah from penjualan p"
  cSQL = "select sum(p.harga*p.qty) as jumlah from penjualan p"
  cSQL = cSQL & " left join totpenjualan t on t.nomorpenjualan = p.nomorpenjualan"
  cSQL = cSQL & " Where  t.kodeanggota = '" & cMember & "' And p.Tgl >= '" & Format(dTgl(0).Value, "yyyy-MM-dd") & "' And p.Tgl <= '" & Format(dTgl(1).Value, "yyyy-MM-dd") & "'"
  Set db = obj.SQL(GetDSN, cSQL)
  If Not db.EOF Then
    GetPenjualanMember = GetNull(db!jumlah)
  End If
  
End Function

Private Function GetPenjualanLevel12(ByVal obj As CodeSuiteLibrary.Data, ByVal cKodeUpline As String) As String
Dim cSQL As String
Dim db As New ADODB.Recordset
Dim nTotalPenjualan As String

  nTotalPenjualan = ""
  cSQL = "select * from anggota"
  cSQL = cSQL & " Where kodeupline = '" & cKodeUpline & "'"
  Set db = obj.SQL(GetDSN, cSQL)
  If Not db.EOF Then
    Do While Not db.EOF
      If GetPenjualanMemberLevel1(obj, GetNull(db!kodeanggota)) <> 0 Then
        nTotalPenjualan = IIf(nTotalPenjualan = "", "", nTotalPenjualan & ", ") & (db!nama) & "(" & Format(GetPenjualanMemberLevel1(obj, GetNull(db!kodeanggota)), "###,###,##0") & "/" & Format(GetPenjualanLevel4(obj, GetNull(db!kodeanggota)), "###,###,##0") & ")"
      End If
      db.MoveNext
    Loop
  End If
  GetPenjualanLevel12 = nTotalPenjualan
End Function

Private Function GetPenjualanMemberLevel1(ByVal obj As CodeSuiteLibrary.Data, ByVal cKodeDownline As String) As Double
Dim db As New ADODB.Recordset
Dim cSQL As String

  GetPenjualanMemberLevel1 = 0
'  cSQL = "select sum(p.jumlah) as jumlah from penjualan p"
  cSQL = "select sum(p.qty*p.harga) as jumlah from penjualan p"
  cSQL = cSQL & " left join totpenjualan t on t.nomorpenjualan = p.nomorpenjualan"
  cSQL = cSQL & " Where p.Discount >= 20 And t.kodeanggota = '" & cKodeDownline & "' And p.Tgl >= '" & Format(dTgl(0).Value, "yyyy-MM-dd") & "' And p.Tgl <= '" & Format(dTgl(1).Value, "yyyy-MM-dd") & "'"
  Set db = obj.SQL(GetDSN, cSQL)
  If Not db.EOF Then
    If GetNull(db!jumlah) > nBatasTutupPoin Then
      GetPenjualanMemberLevel1 = GetNull(db!jumlah)
    Else
      GetPenjualanMemberLevel1 = 0
    End If
  End If
End Function

Private Function GetSaldoPiutangDanTopUp(ByVal anggota As String) As Double
Dim TempPiutang As Double

  'cari jumlah/saldo piutang terakhir setelah dipotong retur
  
  Set dbData = objData.Browse(GetDSN, "kartupiutang", "sum(debet-kredit) as saldopiutang", "kodeanggota", sisAssign, anggota, " and tgl <='" & Format(dTgl(1).Value, "yyyy-MM-dd") & "'")
  If Not dbData.EOF Then
    GetSaldoPiutangDanTopUp = GetNull(dbData!saldopiutang)
  End If
  
End Function

Private Function GetSaldoPiutangDanTopUp2(ByVal anggota As String) As Double
Dim TempTopUp As Double

  'cari saldo top up member
  Set dbData = objData.Browse(GetDSN, "membertopup m", "m.kodeanggota,a.nama,a.alamat,sum(debet) as debet,sum(kredit) as kredit,sum(m.debet-m.kredit) as saldo", "m.kodeanggota", sisAssign, anggota, " and m.tgl <= '" & Format(dTgl(1).Value, "yyyy-MM-dd") & "' GROUP BY m.kodeanggota", , Array("left join anggota a on a.kodeanggota = m.kodeanggota"))
  If Not dbData.EOF Then
    GetSaldoPiutangDanTopUp2 = GetNull(dbData!saldo)
  End If
End Function


Private Function GetPembelanjaanPromo(ByVal anggota As String) As Double
Dim cSQL As String
Dim db As New ADODB.Recordset


  cSQL = cSQL & " select t.kodeanggota,a.nama,sum(jumlah) as jumlahtotal from penjualan p"
  cSQL = cSQL & " left join stock s on s.kodestock = p.kodestock"
  cSQL = cSQL & " left join totpenjualan t on t.nomorpenjualan = p.nomorpenjualan"
  cSQL = cSQL & " left join anggota a on a.kodeanggota = t.kodeanggota"
  
  cSQL = cSQL & " Where"
  cSQL = cSQL & " p.tgl >= '" & Format(dTgl(0).Value, "yyyy-MM-dd") & "' and p.tgl <= '" & Format(dTgl(1).Value, "yyyy-MM-dd") & "'"
  cSQL = cSQL & " and s.diskonpenjualan = 0"
  cSQL = cSQL & " and a.kodeanggota = '" & anggota & "'"
  'cSQL = cSQL & " and s.barcode <> 'CTSM' and s.barcode <> 'BSK' and s.barcode <> 'CTSM0911' and s.barcode <> 'CTSMPRO0911' "
  cSQL = cSQL & " and s.barcode <> 'CTSM' and s.barcode <> 'BSK' and s.barcode <> 'CTSM0911' and s.barcode <> 'CTSMPRO0911' "
  
  cSQL = cSQL & " GROUP BY t.kodeanggota"
  cSQL = cSQL & " ORDER BY jumlahtotal DESC"
  
  Set db = objData.SQL(GetDSN, cSQL)
  If Not db.EOF Then
    GetPembelanjaanPromo = GetNull(db!jumlahtotal)
  End If
  
End Function

Private Sub BiSAButton2_Click()
Dim cTex As String

  cTex = GetDetailJaringan(objData, cCustomer.Text)
  MsgBox cTex
  
  'create file systemobject
  Dim fileSystem As FileSystemObject
  'create a textstream
  Dim exportFile As TextStream
  Set fileSystem = New FileSystemObject
  'open the file for writing
  Set exportFile = fileSystem.CreateTextFile("C:\exportedDetail.txt", True)
  'use GetString method of the recordset to put the data in the textfile
'    exportFile.Write (rsExport.GetString(adClipString, , "|", vbCrLf, ""))
  exportFile.Write cTex
  'close the stream
  exportFile.Close
  Set fileSystem = Nothing
  
  MsgBox "Report Exported on c:\exportedDetail.txt"
'  Set exportFile = fileSystem.OpenTextFile("c:\exportedDetail.txt", ForWriting)
'  exportFile.Read
End Sub

Private Function GetDetailJaringan(ByVal obj As CodeSuiteLibrary.Data, ByVal cKUpline As String) As String
Dim db As New ADODB.Recordset
Dim cTxt As String

  cTxt = ""
  Set db = obj.Browse(GetDSN, "anggota", "kodeanggota,nama", "kodeanggota", sisAssign, cKUpline)
  If Not db.EOF Then
    cTxt = "Nama: " & GetNull(db!nama) & vbCrLf
'    cTxt = cTxt & "Peringkat Franchis TUTUP TPS OK TUTUP NBF OK" & vbCrLf
    cTxt = cTxt & "Belanja Sendiri : " & Format(GetPenjualanMember(objData, GetNull(db!kodeanggota)), "###,###,##0") & vbCrLf
    cTxt = cTxt & "Total Penjualan Group (L1-4) : " & Format(GetPenjualanLevel4(objData, cKUpline), "###,###,##0") & vbCrLf & _
    "Total L1 Saja : " & Format(GetPenjualanLevel1(objData, GetNull(db!kodeanggota)), "###,###,##0") & _
    vbCrLf & GetPenjualanL4Detail(objData, cKUpline) & vbCrLf
  End If
  GetDetailJaringan = cTxt
End Function

Private Sub BiSAButton3_Click()
Dim db As New ADODB.Recordset
Dim cTex As String

  cTex = ""
  Set db = objData.Browse(GetDSN, "anggota", , "kodeupline", sisAssign, "123412341234", " or kodeupline is null or kodeupline = '' or kodeupline = '6000324639' or kodeupline = '7000000000'")
  If Not db.EOF Then
    Do While Not db.EOF
      'cek jika member ini tidak belanja, lewat
'      If GetPenjualanMember(objData, GetNull(db!kodeanggota)) > nBatasTutupPoin Then
      If GetPenjualanLevel1(objData, GetNull(db!kodeanggota)) > nBatasTutupPoin Then
        cTex = cTex & GetDetailJaringan(objData, GetNull(db!kodeanggota))
      End If
'      cTex = cTex & GetDetailJaringan(objData, GetNull(db!kodeanggota))
      db.MoveNext
    Loop
    
    'create file systemobject
    Dim fileSystem As FileSystemObject
    'create a textstream
    Dim exportFile As TextStream
    Set fileSystem = New FileSystemObject
    'open the file for writing
    Set exportFile = fileSystem.CreateTextFile("C:\exportedData.txt", True)
    'use GetString method of the recordset to put the data in the textfile
'    exportFile.Write (rsExport.GetString(adClipString, , "|", vbCrLf, ""))
    exportFile.Write cTex
    'close the stream
    exportFile.Close
    Set fileSystem = Nothing
    
    MsgBox "Report Exported on c:\exportedData.txt"
  End If
End Sub

Private Sub BiSAButton4_Click()
Dim cSQL As String
Dim db As New ADODB.Recordset
Dim n As Single


  cSQL = cSQL & " select a.telp,t.kodeanggota,a.nama,a.kodeupline,sum(jumlah) as jumlahtotal from penjualan p"
  cSQL = cSQL & " left join stock s on s.kodestock = p.kodestock"
  cSQL = cSQL & " left join totpenjualan t on t.nomorpenjualan = p.nomorpenjualan"
  cSQL = cSQL & " left join anggota a on a.kodeanggota = t.kodeanggota"
  
  cSQL = cSQL & " Where"
  cSQL = cSQL & " p.tgl >= '" & Format(dTgl(0).Value, "yyyy-MM-dd") & "' and p.tgl <= '" & Format(dTgl(1).Value, "yyyy-MM-dd") & "'"
  cSQL = cSQL & " and s.diskonpenjualan >= 30"
  
  cSQL = cSQL & " GROUP BY t.kodeanggota"
  cSQL = cSQL & " ORDER BY jumlahtotal DESC"
  
  vaArray.ReDim 0, 1, 0, 10
  Set db = objData.SQL(GetDSN, cSQL)
  If Not db.EOF Then
    FrmPB.InitPB db.RecordCount
    
    vaArray(0, 0) = "Rekap Penjualan Leader/Member Tgl " & Format(dTgl(0).Value, "dd MMMM yyyy") & " sd " & Format(dTgl(1).Value, "dd MMMM yyyy")
    vaArray(0, 1) = ""
    vaArray(0, 2) = ""
    vaArray(0, 3) = ""
    vaArray(0, 4) = ""
    vaArray(0, 5) = ""
    vaArray(0, 6) = ""
    vaArray(0, 7) = ""
    vaArray(0, 8) = ""
    vaArray(0, 9) = ""
    vaArray(0, 10) = ""
    
    vaArray(1, 0) = "NOHP"
    vaArray(1, 1) = "ID"
    vaArray(1, 2) = "Nama"
    vaArray(1, 3) = "Netto Reguler"
    vaArray(1, 4) = "TPG Reguler"
    vaArray(1, 5) = "Netto Promo"
    vaArray(1, 6) = "Piutang"
    vaArray(1, 7) = "Top Up"
    vaArray(1, 8) = "TPG Penjualan sd L4"
    vaArray(1, 9) = "TPG Penjualan L1"
    vaArray(1, 10) = "Downline"
    
  
    Do While Not db.EOF
      FrmPB.RunPB
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = "'" & GetNull(db!telp)
      vaArray(n, 1) = "'" & GetNull(db!kodeanggota)
      vaArray(n, 2) = GetNull(db!nama) & " < " & GetNamaUpline(objData, GetNull(db!kodeupline))
      vaArray(n, 3) = Format(GetNull(db!jumlahtotal), "###,###,##0")
      vaArray(n, 4) = Format(GetNull(db!jumlahtotal) / 0.7, "###,###,##0")
      vaArray(n, 5) = Format(GetPembelanjaanPromo(GetNull(db!kodeanggota)), "###,###,##0")
      vaArray(n, 6) = Format(GetSaldoPiutangDanTopUp(GetNull(db!kodeanggota)), "###,###,##0")
      vaArray(n, 7) = Format(GetSaldoPiutangDanTopUp2(GetNull(db!kodeanggota)), "###,###,##0")
      vaArray(n, 8) = Format(GetPenjualanLevel4(objData, GetNull(db!kodeanggota)), "###,###,##0")
      vaArray(n, 9) = Format(GetPenjualanLevel1(objData, GetNull(db!kodeanggota)), "###,###,##0")
      vaArray(n, 10) = Format(GetPenjualanLevel12(objData, GetNull(db!kodeanggota)), "###,###,##0")
      If GetPenjualanMember(objData, GetNull(db!kodeupline)) > 0 Then
        vaArray.DeleteRows n
      End If
      db.MoveNext
    Loop
    FrmPB.EndPB
  End If
  
    Dim a As New exportExcel
    a.RecordSource = vaArray
    a.ExportToExcel
End Sub

Private Sub BiSAButton5_Click()
cTex = ""
GetRekursif cCustomer.Text
MsgBox cTex
End Sub

Private Sub cCustomer_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "anggota", "kodeanggota,nama", "nama", sisContent, cCustomer.Text)
  If Not dbData.EOF Then
    cCustomer.Text = cCustomer.Browse(dbData, Array("KODE", "NAMA"), , Array(10, 20))
    cCustomer.Text = GetNull(dbData!kodeanggota)
  End If
End Sub

Private Sub cmdCetak_Click()
  GetRpt
End Sub

Private Sub cmdExportToExcel_Click()
  Dim a As New exportExcel
  a.RecordSource = vaArray
  a.ExportToExcel
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
  GetData
End Sub

Private Sub datagrid1_HeadClick(ByVal ColIndex As Integer)
Dim n As Integer
  If vaArray.UpperBound(1) >= 0 Then
    If lClick Then
      Select Case ColIndex
        Case 1
          vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 1, XORDER_ASCEND, XTYPE_STRING
          lClick = Not lClick
        Case 2
          vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 2, XORDER_ASCEND, XTYPE_DOUBLE
          lClick = Not lClick
        Case 3
          vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 3, XORDER_ASCEND, XTYPE_DOUBLE
          lClick = Not lClick
        Case 4
          vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 4, XORDER_ASCEND, XTYPE_DOUBLE
          lClick = Not lClick
      End Select
    Else
      Select Case ColIndex
        Case 1
          vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 1, XORDER_DESCEND, XTYPE_STRING
          lClick = Not lClick
        Case 2
          vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 2, XORDER_DESCEND, XTYPE_DOUBLE
          lClick = Not lClick
        Case 3
          vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 3, XORDER_DESCEND, XTYPE_DOUBLE
          lClick = Not lClick
        Case 4
          vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 4, XORDER_DESCEND, XTYPE_DOUBLE
          lClick = Not lClick
      End Select
    End If
    For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
      vaArray(n, 0) = n + 1
    Next n
    DataGrid1.ReBind
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  lClick = True
  CenterForm Me
  SetIcon Me.hWnd
  
  nBatasTutupPoin = -1
  
  cCustomer.Default
  TabIndex dTgl(0), n
  TabIndex dTgl(1), n
  TabIndex cmdPreview, n
  TabIndex cmdExportToExcel, n
  TabIndex cmdCetak, n
  TabIndex cmdKeluar, n
  
  dTgl(0).Value = BOM(Date)
  dTgl(1).Value = EOM(Date)
End Sub

Private Sub GetData()
Dim n As Integer
Dim nAwal As Double
Dim nMutasi As Double
Dim nTotal As Double
Dim nHB As Double

  vaArray.ReDim 0, -1, 0, 6
  nAwal = 0
  nMutasi = 0
  nTotal = 0
  Set dbData = objData.Browse(GetDSN, "anggota a", "a.kodeanggota,a.nama,g.nama as namaupline", , , , , "a.nama", Array("left join anggota g on g.kodeanggota = a.kodeupline"))
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = n + 1
      vaArray(n, 1) = GetNull(dbData!nama) & " <- " & GetNull(dbData!namaupline)
      nHB = 0
      nHB = GetDataPenjualananggota(GetNull(dbData!kodeanggota), dTgl(0).Value, dTgl(1).Value, False) / IIf(Option1(0).Value = True, 0.7, 1)
      vaArray(n, 2) = nHB 'BulatkanAngka((60 / 100 * nHB) * (3 / 100))
      vaArray(n, 3) = GetPenjualanMember2(objData, GetNull(dbData!kodeanggota)) 'GetDataPenjualananggota(GetNull(dbData!kodeanggota), dTgl(0).Value, dTgl(1).Value, True) / IIf(Option1(0).Value = True, 0.7, 1)
      vaArray(n, 4) = vaArray(n, 2) + vaArray(n, 3)
      vaArray(n, 5) = Format(Devide(vaArray(n, 3) - vaArray(n, 2), vaArray(n, 4)) * 100, "###,###,###0")
      vaArray(n, 6) = BulatkanAngka((60 / 100 * nHB) * (3 / 100))
      
      nAwal = nAwal + vaArray(n, 2)
      nMutasi = nMutasi + vaArray(n, 3)
      nTotal = nTotal + vaArray(n, 4)
      dbData.MoveNext
      If vaArray(n, 4) = 0 Then
        vaArray.DeleteRows n
      End If
    Loop
  End If
  DataGrid1.Columns(2).FooterText = Format(nAwal, "###,###,###,##0")
  DataGrid1.Columns(3).FooterText = Format(nMutasi, "###,###,###,##0")
  DataGrid1.Columns(4).FooterText = Format(nTotal, "###,###,###,##0")
  DataGrid1.Columns(5).FooterText = Format(Devide(nMutasi - nAwal, nTotal) * 100, "###,###,###,##0")
  Set DataGrid1.Array = vaArray
  DataGrid1.ReBind
  DataGrid1.Refresh
End Sub

Private Function GetPenjualanMember2(ByVal obj As CodeSuiteLibrary.Data, ByVal cMember As String) As Double
Dim db As New ADODB.Recordset
Dim cSQL As String

  GetPenjualanMember2 = 0
  'cSQL = "select sum(p.jumlah) as jumlah from penjualan p"
  cSQL = "select sum(p.harga*p.qty) as jumlah from penjualan p"
  cSQL = cSQL & " left join totpenjualan t on t.nomorpenjualan = p.nomorpenjualan"
  cSQL = cSQL & " Where p.Discount >= 30 And t.kodeanggota = '" & cMember & "' And p.Tgl >= '" & Format(dTgl(0).Value, "yyyy-MM-dd") & "' And p.Tgl <= '" & Format(dTgl(1).Value, "yyyy-MM-dd") & "'"
  Set db = obj.SQL(GetDSN, cSQL)
  If Not db.EOF Then
    GetPenjualanMember2 = GetNull(db!jumlah)
  End If
End Function

Private Sub GetRpt()

   With FrmRPT
    .AddPageHeader "Rekapitulasi Omzet", tdbHalignCenter, , , True, , 12, True, , , False, tdbPageHeaderSect
    .AddPageHeader GetNamaBulan(Month(dTgl(0).Value)) & " " & Year(dTgl(0).Value), tdbHalignCenter, , , True, , 12, True, , , False, tdbPageHeaderSect
    .AddPageHeader aCfg(objData, msNamaPerusahaan), tdbHalignCenter, , , True
    .AddPageHeader "", , , , True, , 12, True, , , False, tdbPageHeaderSect
      
    .AddTableHeader "No", , , , 4
    .AddTableHeader "Nama"
    .AddTableHeader "Bln H-1", , , , 13
    .AddTableHeader "Bln H", , , , 13
    .AddTableHeader "Total H+(H-1)", , , , 14
    .AddTableHeader "%", , , , 8
    .AddTableHeader "Bonus Bln H-1", , , , 11
    
    
    .AddTableBody
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2, , , 8
    .AddTableBody Sis_Rpt_Number2, , , 15
    .AddTableBody Sis_Rpt_Number2, , , 15
    .AddTableBody Sis_Rpt_Number2, , , 15
    .AddTableBody Sis_Rpt_Number2, , , 15
    
    
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    
    .Refresh
    .Preview vaArray
  End With
End Sub
Public Function BulatkanAngka(lngAngka As Double) As Long
  If lngAngka Mod 100 > 0 Then
    Dim lngHasil As Long
    lngHasil = lngAngka \ 100
    BulatkanAngka = (lngHasil * 100) + 100
  Else
    BulatkanAngka = lngAngka
  End If
End Function


Private Function GetDataPenjualananggota(ByVal anggota As String, ByVal TglAwal As Date, ByVal TglAkhir As Date, ByVal lMutasi As Boolean) As Double
Dim db As New ADODB.Recordset

  If lMutasi Then
    Set db = objData.Browse(GetDSN, "totpenjualan", "sum(total) as total", "kodeanggota", sisAssign, anggota, " and tgl >= '" & Format(TglAwal, "yyyy-MM-dd") & "' and tgl <= '" & Format(TglAkhir, "yyyy-MM-dd") & "'")
  Else
'    Set db = objData.Browse(GetDSN, "totpenjualan", "sum(total) as total", "kodeanggota", sisAssign, anggota, " and tgl < '" & Format(TglAwal, "yyyy-MM-dd") & "'")
    Set db = objData.Browse(GetDSN, "totpenjualan", "sum(total) as total", "kodeanggota", sisAssign, anggota, " and tgl <= '" & Format(EOM(DateAdd("m", -1, TglAwal)), "yyyy-MM-dd") & "' and tgl >= '" & Format(BOM(DateAdd("m", -1, TglAwal)), "yyyy-MM-dd") & "'")
  End If
  If Not db.EOF Then
    GetDataPenjualananggota = GetNull(db!Total)
  End If
End Function

Private Sub GetRekursif(ByVal cParentID As String)
Dim db As New ADODB.Recordset
Dim c, i As Integer

  Set db = objData.Browse(GetDSN, "anggota", "kodeanggota,kodeupline,nama", "kodeupline", sisAssign, cParentID)
  If Not db.EOF Then
    Do While Not db.EOF
      If GetNull(db!kodeupline) = cParentID Then
        GetRekursif (db!kodeanggota)
      End If
      db.MoveNext
    Loop
  End If
End Sub

Private Function nGetTab(ByVal nTab As Integer) As String
Dim i As Integer

  For i = 0 To nTab
    nGetTab = nGetTab & vbTab & " "
  Next i
  
End Function
