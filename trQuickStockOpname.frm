VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trQuickStockOpname 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quick Stock Opname"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   11790
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   0
      Top             =   6495
      Width           =   11760
      _ExtentX        =   20743
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
         Left            =   1170
         TabIndex        =   0
         Top             =   75
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
         Picture         =   "trQuickStockOpname.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   2325
         TabIndex        =   1
         Top             =   75
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
         Picture         =   "trQuickStockOpname.frx":028A
      End
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   2775
         TabIndex        =   2
         Top             =   75
         Visible         =   0   'False
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
         Picture         =   "trQuickStockOpname.frx":0429
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   435
         Left            =   105
         TabIndex        =   3
         Top             =   75
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
         Picture         =   "trQuickStockOpname.frx":0555
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   10545
         TabIndex        =   4
         Top             =   105
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
         Picture         =   "trQuickStockOpname.frx":0700
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   9450
         TabIndex        =   5
         Top             =   105
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
         Picture         =   "trQuickStockOpname.frx":07A6
      End
   End
   Begin BiSAFramProject.BiSAFrame BisaFrame2 
      Height          =   4455
      Left            =   0
      Top             =   2055
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   7858
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   3900
         Left            =   60
         TabIndex        =   6
         Top             =   510
         Width           =   11640
         _ExtentX        =   20532
         _ExtentY        =   6879
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
         Columns(1).Caption=   "Barcode"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Nama"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Beli"
         Columns(3).DataField=   ""
         Columns(3).NumberFormat=   "###,###,###"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Qty"
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "###,###,###,###,##0.00"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   5
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=5"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=926"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=847"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=514"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=3016"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2937"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=512"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=9657"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=9578"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=512"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=3334"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=3254"
         Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=514"
         Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(25)=   "Column(4).Width=2355"
         Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2275"
         Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=197122"
         Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
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
         HeadLines       =   1.5
         FootLines       =   0
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   16777215
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
         _StyleDefs(13)  =   ":id=2,.fontname=Tahoma"
         _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(15)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
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
         _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(28)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(29)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(31)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(32)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(33)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(34)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(35)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(36)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=1"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=0"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=70,.parent=13,.alignment=1"
         _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=67,.parent=14"
         _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=68,.parent=15"
         _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=69,.parent=17"
         _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=66,.parent=13,.alignment=1"
         _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=14"
         _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=15,.alignment=1"
         _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=17"
         _StyleDefs(57)  =   "Named:id=33:Normal"
         _StyleDefs(58)  =   ":id=33,.parent=0"
         _StyleDefs(59)  =   "Named:id=34:Heading"
         _StyleDefs(60)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(61)  =   ":id=34,.wraptext=-1"
         _StyleDefs(62)  =   "Named:id=35:Footing"
         _StyleDefs(63)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(64)  =   "Named:id=36:Selected"
         _StyleDefs(65)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(66)  =   "Named:id=37:Caption"
         _StyleDefs(67)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(68)  =   "Named:id=38:HighlightRow"
         _StyleDefs(69)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(70)  =   "Named:id=39:EvenRow"
         _StyleDefs(71)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(72)  =   "Named:id=40:OddRow"
         _StyleDefs(73)  =   ":id=40,.parent=33"
         _StyleDefs(74)  =   "Named:id=41:RecordSelector"
         _StyleDefs(75)  =   ":id=41,.parent=34"
         _StyleDefs(76)  =   "Named:id=42:FilterBar"
         _StyleDefs(77)  =   ":id=42,.parent=33"
      End
      Begin BiSATextBoxProject.BiSABrowse cNama 
         Height          =   330
         Left            =   2340
         TabIndex        =   16
         Top             =   150
         Width           =   5460
         _ExtentX        =   9631
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
         Left            =   615
         TabIndex        =   17
         Top             =   150
         Width           =   1695
         _ExtentX        =   2990
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
      Begin BiSANumberBoxProject.BiSANumberBox nNomor 
         Height          =   330
         Left            =   60
         TabIndex        =   18
         Top             =   150
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   582
         Appearance      =   0
         Decimals        =   0
         DecimalPoint    =   ""
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
      Begin BiSANumberBoxProject.BiSANumberBox nQty 
         Height          =   330
         Left            =   9660
         TabIndex        =   19
         Top             =   150
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   582
         Appearance      =   0
         Decimals        =   0
         MinValue        =   -9999999
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
      Begin BiSANumberBoxProject.BiSANumberBox nHarga 
         Height          =   330
         Left            =   7815
         TabIndex        =   20
         Top             =   150
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   582
         Appearance      =   0
         Decimals        =   0
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
      Begin BiSAButtonProject.BiSAButton cmdOK 
         Height          =   330
         Left            =   11070
         TabIndex        =   21
         Top             =   150
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   582
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
         Picture         =   "trQuickStockOpname.frx":0A2C
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   2055
      Left            =   0
      Top             =   0
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   3625
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
      Begin BiSANumberBoxProject.BiSANumberBox nColKode 
         Height          =   360
         Left            =   6405
         TabIndex        =   13
         Top             =   1050
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   635
         Appearance      =   0
         Decimals        =   0
         MinValue        =   1
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
         MinimumError    =   "Minimal angka yg valid adalah 1"
      End
      Begin BiSAButtonProject.BiSAButton BiSAButton1 
         Height          =   465
         Left            =   1605
         TabIndex        =   12
         Top             =   1485
         Width           =   4410
         _ExtentX        =   7779
         _ExtentY        =   820
         Caption         =   "IMPORT EXCEL"
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
      Begin BiSATextBoxProject.BiSATextBox cKeterangan 
         Height          =   330
         Left            =   1590
         TabIndex        =   7
         Top             =   1125
         Width           =   4470
         _ExtentX        =   7885
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
      Begin BiSATextBoxProject.BiSABrowse cNamaGudang 
         Height          =   330
         Left            =   3555
         TabIndex        =   8
         Top             =   765
         Width           =   2505
         _ExtentX        =   4419
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
      Begin BiSATextBoxProject.BiSABrowse cGudang 
         Height          =   330
         Left            =   75
         TabIndex        =   9
         Top             =   765
         Width           =   3405
         _ExtentX        =   6006
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
         Caption         =   "Gudang"
         CaptionWidth    =   1400
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
         Left            =   75
         TabIndex        =   10
         Top             =   75
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   582
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
         CaptionWidth    =   1400
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
         Height          =   330
         Left            =   75
         TabIndex        =   11
         Top             =   420
         Width           =   3405
         _ExtentX        =   6006
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
         Caption         =   "Nomor"
         CaptionWidth    =   1400
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
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   9120
         Top             =   615
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin BiSANumberBoxProject.BiSANumberBox nColQty 
         Height          =   360
         Left            =   7155
         TabIndex        =   14
         Top             =   1050
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   635
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
      Begin VB.Label Label3 
         Caption         =   "KOLOM"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6630
         TabIndex        =   23
         Top             =   540
         Width           =   870
      End
      Begin VB.Label Label2 
         Caption         =   "Qty"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7215
         TabIndex        =   22
         Top             =   750
         Width           =   435
      End
      Begin VB.Label Label1 
         Caption         =   " Barcode"
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
         Left            =   6240
         TabIndex        =   15
         Top             =   765
         Width           =   885
      End
   End
End
Attribute VB_Name = "trQuickStockOpname"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Gudang As String

Dim nPos As SisPos
Dim lEdit As Boolean
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB
Dim cKode As String

Dim Excel As Excel.Application
Dim ExcelWBk As Excel.Workbook
Dim ExcelWS As Excel.Worksheet

Private Sub StartExcel()
  On Error GoTo err:
  Set Excel = GetObject(, "Excel.Application")
  Exit Sub
err:
  Set Excel = CreateObject("Excel.Application")
End Sub

Private Sub CloseWorkSheet()
  ExcelWBk.Close
  Excel.Quit
End Sub

Private Sub FinishExcel()
  'Jangan lupa, selalu bersihkan memory saat mengakhiri
  If Not ExcelWS Is Nothing Then Set ExcelWS = Nothing
  If Not ExcelWBk Is Nothing Then Set ExcelWBk = Nothing
  If Not Excel Is Nothing Then Set Excel = Nothing
End Sub
Private Sub GetFakturBrowse(ByVal lStat As Boolean)
  cNomor.Button = lStat
End Sub

Private Sub BiSAButton1_Click()
  If nColKode.Value >= 1 And nColQty.Value >= 1 Then
    CommonDialog1.Filter = "Excel File (*.xls)|*.xls"
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
      GetLoadExcel
    End If
  Else
    MsgBox "Value dari Kol Kode dan Kol Qty belum benar", vbExclamation
  End If
End Sub

Private Sub GetLoadExcel()
Dim lSave As Boolean
Dim vaField, vaValue
Dim i, j, n As Integer
Dim dbData As New ADODB.Recordset
Dim Wb As Excel.Workbook

'  On Error GoTo err:
  StartExcel
  lSave = True
  
  Excel.Workbooks.Close
  Set ExcelWBk = Excel.Workbooks.Open(CommonDialog1.FileName)
  Set ExcelWS = ExcelWBk.Worksheets(1)
  
  
  FrmPB.InitPB ExcelWS.Cells.SpecialCells(xlCellTypeLastCell).Row
  Dim cBarcode
  Dim cQty
  
  GetRefreshGrid
  
  For i = 1 To ExcelWS.Cells.SpecialCells(xlCellTypeLastCell).Row
    FrmPB.RunPB
    With ExcelWS
      Set dbData = objData.Browse(GetDSN, "stock", "kodestock,nama,hargabeli,diskonpenjualan,kodesatuan,barcode", "barcode", sisAssign, .Cells(i, nColKode.Value).Value)
      If Not dbData.EOF Then
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        n = vaArray.UpperBound(1)
        vaArray(n, 0) = n + 1
        vaArray(n, 1) = .Cells(i, nColKode.Value).Value
        vaArray(n, 2) = GetNull(dbData!nama)
        vaArray(n, 3) = GetNull(dbData!hargabeli)
        vaArray(n, 4) = .Cells(i, nColQty.Value).Value
        vaArray(n, 5) = GetNull(dbData!KodeStock)
        If vaArray(n, 4) = 0 Then vaArray.DeleteRows n
      Else
        'jika data yg di import tidak ada dalam database simpan
      End If
    End With
  Next i
  
  Set tdbgrid1.Array = vaArray
  tdbgrid1.ReBind
  tdbgrid1.Refresh
  
  FrmPB.EndPB
  CloseWorkSheet
  FinishExcel
  
'err:
'MsgBox "Err"
End Sub

Private Sub cBarcode_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "stock", "barcode,nama,hargabeli,kodestock", "barcode", sisContent, cBarcode.Text)
  If Not dbData.EOF Then
    cBarcode.Text = cBarcode.Browse(dbData)
    cNama.Text = GetNull(dbData!nama)
    nHarga.Value = GetNull(dbData!hargabeli)
    cKode = GetNull(dbData!KodeStock)
  End If
End Sub

Private Sub cGudang_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "gudang", "kodegudang,keterangan")
  If Not dbData.EOF Then
    cGudang.Text = cGudang.Browse(dbData)
    cNamaGudang.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub cmdAktivasi_Click()
  frmAktivasi.Action Me
End Sub

Private Sub cmdAdd_Click()
  nPos = Add
  GetEdit True
  GetFakturBrowse False
  cNomor.Text = GetNomor("totstockopname", "nomorstockopname", GetID, sisModulTransaksi.StockOpname)
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
    initvalue
  Else
    Unload Me
  End If
End Sub

Private Function ValidSaving() As Boolean
  ValidSaving = True
End Function

Private Sub cmdOK_Click()
Dim n As Integer
Dim nJumlah1 As Double
Dim nQtyTmp As Single

  
  If validOK() Then
    'jika baris <= Nomor
    If vaArray.UpperBound(1) + 2 <= nNomor.Value Then
      vaArray.ReDim 0, nNomor.Value - 1, 0, 5
      n = vaArray.UpperBound(1)
    'jika baris=0
    ElseIf vaArray.UpperBound(1) = -1 Then
      nNomor.Value = 1
      vaArray.ReDim 0, nNomor.Value - 1, 0, 5
      n = vaArray.UpperBound(1)
    Else
      n = nNomor.Value - 1
    End If
        
    vaArray(n, 0) = nNomor.Value
    vaArray(n, 1) = cBarcode.Text
    vaArray(n, 2) = cNama.Text
    vaArray(n, 3) = nHarga.Value
    vaArray(n, 4) = nQty.Value
    vaArray(n, 5) = cKode
    
    tdbgrid1.Array = vaArray
    tdbgrid1.ReBind
    tdbgrid1.MoveNext
    InitValue1
    nNomor.Value = vaArray.UpperBound(1) + 2
    cBarcode.SetFocus
  End If
End Sub

Private Sub InitValue1()
  nNomor.Value = 1
  cBarcode.Default
  cNama.Default
  nHarga.Value = 0
  nQty.Value = 0
  cKode = ""
End Sub

Private Function validOK() As Boolean

  validOK = True
  If Not GetValidDataBrowse(objData, "stock", "kodestock", cKode) Then
    MsgBox "Barang tersebut tidak ada dalam database "
    cBarcode.SetFocus
    validOK = False
    Exit Function
  End If
  
  If Not GetValidDataBrowse(objData, "stock", "barcode", cBarcode.Text) Then
    MsgBox "Barang tersebut tidak ada dalam database "
    cBarcode.SetFocus
    validOK = False
    Exit Function
  End If
End Function

Private Sub cmdSimpan_Click()
Dim lSave As Boolean
Dim Faktur As String
Dim n As Integer

  lSave = True
  objData.Start GetDSN
  Faktur = cNomor.Text
  If nPos = Add Then
    If Not GetAvailable(cNomor.Text, "totstockopname", "nomorstockopname") Then
      Faktur = GetNomor("totstockopname", "nomorstockopname", GetID, sisModulTransaksi.StockOpname)
    End If
  End If
  
  lSave = IIf(lSave, objData.Delete(GetDSN, "totstockopname", "nomorstockopname", sisAssign, Faktur), False)
  lSave = IIf(lSave, objData.Delete(GetDSN, "stockopname", "nomorstockopname", sisAssign, Faktur), False)
  lSave = IIf(lSave, objData.Delete(GetDSN, "kartustock", "nomor", sisAssign, Faktur), False)
  lSave = IIf(lSave, objData.Delete(GetDSN, "bukubesar", "faktur", sisAssign, Faktur), False)
  
  lSave = IIf(lSave, objData.Update(GetDSN, "totstockopname", "nomorstockopname = '" & Faktur & "'", Array("nomorstockopname", "tgl", "keterangan", "username", "datetime", "kodegudang"), Array(Faktur, Format(dTgl.Value, "yyyy-MM-dd"), cKeterangan.Text, GetRegistry(reg_UserName), SNow, cGudang.Text)), False)
  
  For n = 0 To vaArray.UpperBound(1)
    lSave = IIf(lSave, objData.Add(GetDSN, "stockopname", Array("nomorstockopname", "kodestock", "adjust", "kodegudang"), Array(Faktur, vaArray(n, 5), vaArray(n, 4), cGudang.Text)), False)
    lSave = IIf(lSave, UpdKartuStock(objData, SisKartuStock.PenyesuaianStock, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), vaArray(n, 5), vaArray(n, 4), GetHargaPokok(objData, vaArray(n, 5)), 0, "Penyesuaian " & cKeterangan.Text, cGudang.Text, GetHargaPokok(objData, vaArray(n, 5))), False)
    
    If vaArray(n, 4) < 0 Then
      'jurnal
      ' biaya
      '  persediaan
      lSave = IIf(lSave, UpdKodeTr(objData, msPenyesuaian, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningPenyesuaianKurang), aCfg(objData, msCostCenterJualBeli), "Penyesuaian " & vaArray(n, 2), Abs(vaArray(n, 4)) * GetHargaBeli(objData, vaArray(n, 5)), 0, "", SNow), False)
        lSave = IIf(lSave, UpdKodeTr(objData, msPenyesuaian, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), GetAkunInventory(objData, vaArray(n, 5)), aCfg(objData, msCostCenterJualBeli), "Penyesuaian " & vaArray(n, 2), 0, Abs(vaArray(n, 4)) * GetHargaBeli(objData, vaArray(n, 5)), "", SNow), False)
    Else
      'jurnal
      ' persediaan
      '  pendapatan/modal
      lSave = IIf(lSave, UpdKodeTr(objData, msPenyesuaian, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), GetAkunInventory(objData, vaArray(n, 5)), aCfg(objData, msCostCenterJualBeli), "Penyesuaian " & vaArray(n, 2), Abs(vaArray(n, 4)) * GetHargaBeli(objData, vaArray(n, 5)), 0, "", SNow), False)
        lSave = IIf(lSave, UpdKodeTr(objData, msPenyesuaian, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningPenyesuian), aCfg(objData, msCostCenterJualBeli), "Penyesuaian " & vaArray(n, 2), 0, Abs(vaArray(n, 4)) * GetHargaBeli(objData, vaArray(n, 5)), "", SNow), False)
    End If
  Next
    
  If lSave Then
    objData.Save GetDSN
  Else
    objData.Cancel GetDSN
  End If
  GetEdit False
End Sub

Private Sub cNama_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "stock", "nama,barcode,hargabeli,kodestock", "nama", sisContent, cNama.Text)
  If Not dbData.EOF Then
    cNama.Text = cNama.Browse(dbData)
    cBarcode.Text = GetNull(dbData!barcode)
    nHarga.Value = GetNull(dbData!hargabeli)
    cKode = GetNull(dbData!KodeStock)
  End If
End Sub


Private Sub cNomor_ButtonClick()
Dim lSave As Boolean
Dim n As Single

lSave = True

  Set dbData = objData.Browse(GetDSN, "totstockopname t", "t.nomorstockopname,t.keterangan,t.username,t.kodegudang,t.keterangan,g.keterangan as namagudang", "t.tgl", sisAssign, Format(dTgl.Value, "yyyy-MM-dd"), , , Array("LEFT JOIN gudang g ON g.kodegudang = t.kodegudang"))
  If Not dbData.EOF Then
    cNomor.Text = cNomor.Browse(dbData)
    cGudang.Text = GetNull(dbData!kodegudang)
    cNamaGudang.Text = GetNull(dbData!namagudang)
    cKeterangan.Text = GetNull(dbData!keterangan)
    objData.Start GetDSN
    If nPos = Delete Then
      'munculkan konten yg mau di hapus
      Set dbData = objData.Browse(GetDSN, "stockopname o", "s.barcode,s.kodestock,s.nama,s.hargabeli,o.adjust", "o.nomorstockopname", sisAssign, cNomor.Text, , , Array("left join stock s on s.kodestock = o.kodestock"))
      If Not dbData.EOF Then
        Do While Not dbData.EOF
          vaArray.InsertRows vaArray.UpperBound(1) + 1
          n = vaArray.UpperBound(1)
          vaArray(n, 0) = n + 1
          vaArray(n, 1) = GetNull(dbData!barcode)
          vaArray(n, 2) = GetNull(dbData!nama)
          vaArray(n, 3) = GetNull(dbData!hargabeli)
          vaArray(n, 4) = GetNull(dbData!adjust)
          vaArray(n, 5) = GetNull(dbData!KodeStock)
          dbData.MoveNext
        Loop
        Set tdbgrid1.Array = vaArray
        tdbgrid1.ReBind
        tdbgrid1.Refresh
        Me.Refresh
      End If
      If MsgBox("Yakin data akan dihapus?", vbYesNo) = vbYes Then
        lSave = IIf(lSave, objData.Delete(GetDSN, "totstockopname", "nomorstockopname", sisAssign, cNomor.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "stockopname", "nomorstockopname", sisAssign, cNomor.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "kartustock", "nomor", sisAssign, cNomor.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "bukubesar", "faktur", sisAssign, cNomor.Text), False)
      End If
      If lSave Then
        objData.Save GetDSN
      Else
        objData.Cancel GetDSN
      End If
      GetEdit False
      initvalue
    End If
  End If
End Sub

Private Sub dTgl_Validate(Cancel As Boolean)
  If Not IsInPeriod(dTgl.Value) Or (dTgl.Value > Date) Then
    Cancel = True
    dTgl.SetFocus
    GetEdit False
  End If
End Sub

Private Sub Form_Activate()
  Me.Refresh
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  initvalue
  Me.Caption = "STOCK OPNAME GUDANG " & UCase(Gudang)
  GetEdit False
  TabIndex dTgl, n
  TabIndex cNomor, n
  TabIndex cGudang, n
  TabIndex cKeterangan, n

  TabIndex nNomor, n
  TabIndex cBarcode, n
  TabIndex cNama, n
  TabIndex nHarga, n
  TabIndex nQty, n
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
  cGudang.Default
  cNamaGudang.Default
  cKeterangan.Default
  nQty.Default
  nColKode.Value = 3
  nColQty.Value = 6
  GetRefreshGrid
  InitValue1
End Sub

Private Sub GetRefreshGrid()
  vaArray.ReDim 0, -1, 0, 5
  Set tdbgrid1.Array = vaArray
  tdbgrid1.ReBind
  tdbgrid1.Refresh
End Sub

Private Sub GetEdit(lPar As Boolean)
  BiSAFrame1.Enabled = lPar
  BiSAFrame2.Enabled = lPar
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

Private Sub Form_Resize()
Dim nSisaLebar As Double

  If Me.WindowState = 2 Then
    Me.Refresh
    nSisaLebar = tdbgrid1.Width - tdbgrid1.Columns(0).Width - tdbgrid1.Columns(1).Width - tdbgrid1.Columns(3).Width - tdbgrid1.Columns(4).Width - tdbgrid1.Columns(5).Width - tdbgrid1.Columns(6).Width - tdbgrid1.Columns(7).Width
    tdbgrid1.Columns(2).Width = nSisaLebar - 1000
  End If
End Sub

Private Sub nNomor_Validate(Cancel As Boolean)
Dim n As Single
  
  If GetValidNomorUrut(nNomor, vaArray) Then
    n = nNomor.Value - 1
    If n <= vaArray.UpperBound(1) Then
      cBarcode.Text = vaArray(n, 1)
      cNama.Text = vaArray(n, 2)
      nHarga.Value = vaArray(n, 3)
      nQty.Value = vaArray(n, 4)
      cKode = vaArray(n, 5)
    End If
  End If
End Sub

Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
  tdbgrid1.Update
End Sub

Private Sub tdbgrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
  If Not IsNumeric(tdbgrid1.Columns(6).Value) Then
    Cancel = True
    Exit Sub
  End If
  If Not IsNumeric(tdbgrid1.Columns(7).Value) Then
    Cancel = True
    Exit Sub
  End If
  If ColIndex < 6 Then
    Cancel = True
    Exit Sub
  End If
  tdbgrid1.Columns(7).Value = Val(tdbgrid1.Columns(5).Value) + Val(tdbgrid1.Columns(6).Value)
End Sub

Private Sub TDBGrid1_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueOleDBGrid70.StyleDisp)
On Error Resume Next

  Select Case Col
      Case 5
          Dim Col1 As Long
          Col1 = CLng(tdbgrid1.Columns(5).CellText(Bookmark))
          If Col1 < 0 Then CellStyle.ForeColor = vbRed
      Case 6
          Dim Col2 As Long
          Col2 = CLng(tdbgrid1.Columns(6).CellText(Bookmark))
          If Col2 < 0 Then CellStyle.ForeColor = vbRed
      Case 7
          Dim Col3 As Long
          Col3 = CLng(tdbgrid1.Columns(7).CellText(Bookmark))
          If Col3 < 0 Then CellStyle.ForeColor = vbRed
          
  End Select
End Sub
