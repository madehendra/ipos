VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trWL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BACKORDER/ WAITING LIST/ WL"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12765
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   12765
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   675
      Left            =   0
      Top             =   7440
      Width           =   12765
      _ExtentX        =   22516
      _ExtentY        =   1191
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
         Left            =   11535
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
         Picture         =   "trWL.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   10455
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
         Picture         =   "trWL.frx":00A6
      End
   End
   Begin BiSAFramProject.BiSAFrame BisaFrame2 
      Height          =   5730
      Left            =   0
      Top             =   1725
      Width           =   12780
      _ExtentX        =   22543
      _ExtentY        =   10107
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
      Begin BiSANumberBoxProject.BiSANumberBox nQty 
         Height          =   330
         Left            =   7035
         TabIndex        =   2
         Top             =   75
         Width           =   1275
         _ExtentX        =   2249
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
      Begin BiSANumberBoxProject.BiSANumberBox nNomor 
         Height          =   330
         Left            =   105
         TabIndex        =   3
         Top             =   75
         Width           =   555
         _ExtentX        =   979
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   5190
         Left            =   105
         TabIndex        =   4
         Top             =   420
         Width           =   12630
         _ExtentX        =   22278
         _ExtentY        =   9155
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
         Columns(2).Caption=   "Barang"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Kode"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Qty"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Harga"
         Columns(5).DataField=   ""
         Columns(5).NumberFormat=   "###,###,###"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Jumlah"
         Columns(6).DataField=   ""
         Columns(6).NumberFormat=   "###,###,###"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=926"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=847"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=514"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=2884"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2805"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=512"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=5953"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=5874"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=512"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=2381"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2302"
         Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=197120"
         Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(25)=   "Column(4).Width=2275"
         Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2196"
         Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=514"
         Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(31)=   "Column(5).Width=3043"
         Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=2963"
         Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=514"
         Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(37)=   "Column(6).Width=3969"
         Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=3889"
         Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=514"
         Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
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
         ColumnFooters   =   -1  'True
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1.5
         FootLines       =   1
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
         _StyleDefs(16)  =   ":id=3,.fontname=Tahoma"
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
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=66,.parent=13,.alignment=0"
         _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=14"
         _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=15,.alignment=1"
         _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=17"
         _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
         _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
         _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=1"
         _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=54,.parent=13,.alignment=1"
         _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=51,.parent=14"
         _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=52,.parent=15"
         _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=53,.parent=17"
         _StyleDefs(65)  =   "Named:id=33:Normal"
         _StyleDefs(66)  =   ":id=33,.parent=0"
         _StyleDefs(67)  =   "Named:id=34:Heading"
         _StyleDefs(68)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(69)  =   ":id=34,.wraptext=-1"
         _StyleDefs(70)  =   "Named:id=35:Footing"
         _StyleDefs(71)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(72)  =   "Named:id=36:Selected"
         _StyleDefs(73)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(74)  =   "Named:id=37:Caption"
         _StyleDefs(75)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(76)  =   "Named:id=38:HighlightRow"
         _StyleDefs(77)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(78)  =   "Named:id=39:EvenRow"
         _StyleDefs(79)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(80)  =   "Named:id=40:OddRow"
         _StyleDefs(81)  =   ":id=40,.parent=33"
         _StyleDefs(82)  =   "Named:id=41:RecordSelector"
         _StyleDefs(83)  =   ":id=41,.parent=34"
         _StyleDefs(84)  =   "Named:id=42:FilterBar"
         _StyleDefs(85)  =   ":id=42,.parent=33"
      End
      Begin BiSAButtonProject.BiSAButton cmdOK 
         Height          =   330
         Left            =   12300
         TabIndex        =   5
         Top             =   75
         Width           =   390
         _ExtentX        =   688
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
         Picture         =   "trWL.frx":032C
      End
      Begin BiSATextBoxProject.BiSABrowse cBarcode 
         Height          =   330
         Left            =   675
         TabIndex        =   6
         Top             =   75
         Width           =   1590
         _ExtentX        =   2805
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
      Begin BiSATextBoxProject.BiSABrowse cNama 
         Height          =   330
         Left            =   2280
         TabIndex        =   7
         Top             =   75
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
      Begin BiSATextBoxProject.BiSATextBox cKode 
         Height          =   330
         Left            =   5670
         TabIndex        =   8
         Top             =   75
         Width           =   1350
         _ExtentX        =   2381
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
      Begin BiSANumberBoxProject.BiSANumberBox nHarga 
         Height          =   330
         Left            =   8325
         TabIndex        =   9
         Top             =   75
         Width           =   1695
         _ExtentX        =   2990
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
      Begin BiSANumberBoxProject.BiSANumberBox nJumlah 
         Height          =   330
         Left            =   10050
         TabIndex        =   10
         Top             =   75
         Width           =   2220
         _ExtentX        =   3916
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
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1740
      Left            =   0
      Top             =   0
      Width           =   12765
      _ExtentX        =   22516
      _ExtentY        =   3069
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
      Begin BiSATextBoxProject.BiSATextBox cKodeMember 
         Height          =   330
         Left            =   5310
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   750
         Width           =   2325
         _ExtentX        =   4101
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
      Begin BiSADateProject.BiSADate dTgl 
         Height          =   330
         Left            =   75
         TabIndex        =   12
         Top             =   405
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
      Begin BiSATextBoxProject.BiSABrowse cNamaMember 
         Height          =   330
         Left            =   75
         TabIndex        =   13
         Top             =   750
         Width           =   5205
         _ExtentX        =   9181
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
         Caption         =   "Member"
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
      Begin BiSATextBoxProject.BiSATextBox cAlamatMember 
         Height          =   345
         Left            =   75
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1095
         Width           =   5205
         _ExtentX        =   9181
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
         Caption         =   "Alamat"
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
      Begin BiSATextBoxProject.BiSATextBox cTelpMember 
         Height          =   345
         Left            =   5310
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1095
         Width           =   2325
         _ExtentX        =   4101
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
Attribute VB_Name = "trWL"
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
Dim vaExport As New XArrayDB

Private Sub cBarcode_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "stock s", "s.kodestock,s.barcode,s.nama,s.kodesatuan,s.hargabeli,s.hargajual", "s.barcode", sisContent, cBarcode.Text, " AND s.jenis < 9 and s.statusnonaktif <> 1")
  If Not dbData.EOF Then
    cBarcode.Text = cBarcode.Browse(dbData)
    GetDataStock
  Else
    cBarcode.Default
  End If
End Sub


Private Sub GetDataStock()
  cBarcode.Text = GetNull(dbData!Barcode, "")
  cKode.Text = GetNull(dbData!KodeStock, "")
  cNama.Text = GetNull(dbData!nama, "")
  nHarga.Value = GetNull(dbData!hargajual, 0)
  nJumlah.Value = nQty.Value * nHarga.Value
End Sub

Private Function validOK() As Boolean
  validOK = True
  If Not GetValidDataBrowse(objData, "stock", "kodestock", cKode.Text) Then
    MsgBox "Barang tersebut tidak ada dalam database"
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

Private Sub cmdExportOrderan_Click()
Dim cSQL As String
Dim n As Single
Dim db As New ADODB.Recordset


  cSQL = " "
  cSQL = cSQL & " select a.nama,o.barcode,sum(o.debet-o.kredit) as jumlah from orderan  o"
  cSQL = cSQL & " left join anggota a on a.kodeanggota = o.kodeanggota"
  cSQL = cSQL & " GROUP BY o.kodeanggota"
  cSQL = cSQL & " Having Sum(o.debet - o.kredit) > 0"
  cSQL = cSQL & " ORDER BY o.barcode,o.kodeanggota"
  
  Set db = objData.Sql(GetDSN, cSQL)
  vaArray.ReDim 0, -1, 0, 2
  If Not db.EOF Then
    Do While Not db.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(db!nama)
      vaArray(n, 1) = GetNull(db!Barcode)
      vaArray(n, 2) = GetNull(db!jumlah)
      
      db.MoveNext
    Loop
    Dim a As New exportExcel
    a.RecordSource = vaArray
    a.ExportToExcel
  Else
    MsgBox "Maaf, tidak ada data untuk ditampilkan"
  End If
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
Dim n As Integer
Dim nJumlah1 As Double
Dim nQtyTmp As Single

  If validOK() Then
   
    'jika baris <= Nomor
    If vaArray.UpperBound(1) + 2 <= nNomor.Value Then
      vaArray.ReDim 0, nNomor.Value - 1, 0, 6
      n = vaArray.UpperBound(1)
    'jika baris=0
    ElseIf vaArray.UpperBound(1) = -1 Then
      nNomor.Value = 1
      vaArray.ReDim 0, nNomor.Value - 1, 0, 6
      n = vaArray.UpperBound(1)
    Else
      n = nNomor.Value - 1
    End If
        
    vaArray(n, 0) = nNomor.Value
    vaArray(n, 1) = cBarcode.Text
    vaArray(n, 2) = cNama.Text
    vaArray(n, 3) = cKode.Text
    vaArray(n, 4) = nQty.Value
    vaArray(n, 5) = nHarga.Value
    vaArray(n, 6) = nJumlah.Value
          
    TDBGrid1.Array = vaArray
    TDBGrid1.ReBind
    TDBGrid1.MoveNext
    
    nJumlah1 = 0
    For n = 0 To vaArray.UpperBound(1)
      nJumlah1 = nJumlah1 + vaArray(n, 6)
      nQtyTmp = nQtyTmp + vaArray(n, 4)
    Next
    TDBGrid1.Columns(4).FooterText = Format(nQtyTmp, "###,###,##0")
    TDBGrid1.Columns(6).FooterText = Format(nJumlah1, "###,###,##0")

    
    InitValue1
    
    nNomor.Value = vaArray.UpperBound(1) + 2
    cBarcode.SetFocus
  End If
End Sub

Private Function ValidSaving() As Boolean
Dim n As Integer

  ValidSaving = True
  If Not GetValidDataBrowse(objData, "anggota", "kodeanggota", cKodeMember.Text) Then
    MsgBox "Member tersebut tidak ada dalam database", vbExclamation
    cNamaMember.SetFocus
    ValidSaving = False
    Exit Function
  End If
End Function

Private Sub cmdSimpan_Click()
Dim lSave As Boolean
Dim Faktur As String
Dim n As Integer
Dim vaField
Dim vaValue

lSave = True
  
  If ValidSaving Then
    objData.Start GetDSN
        
    vaField = Array("status", "tgl", "kodeanggota", "kodestock", "barcode", "kredit", "username", "datetime")
    
    lSave = IIf(lSave, objData.Delete(GetDSN, "orderan", "tgl", sisAssign, Format(dTgl.Value, "yyyy-MM-dd"), " and kodeanggota = '" & cKodeMember.Text & "' AND status = 6"), False)
    For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
      vaValue = Array("6", Format(dTgl.Value, "yyyy-MM-dd"), cKodeMember.Text, vaArray(n, 3), vaArray(n, 1), vaArray(n, 4), GetRegistry(reg_UserName), SNow)
      lSave = IIf(lSave, objData.Add(GetDSN, "orderan", vaField, vaValue), False)
    Next n
    
    If lSave Then
      objData.Save GetDSN
    Else
      objData.Cancel GetDSN
    End If
    
    initvalue
    InitValue1
  End If
End Sub

Private Sub cNama_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "stock s", "s.kodestock,s.Barcode,s.nama,s.kodesatuan,s.hargabeli,s.hargajual", "s.nama", sisContent, cNama.Text, " AND s.jenis < 9 and s.statusnonaktif <> 1")
  If Not dbData.EOF Then
    cNama.Text = cNama.Browse(dbData)
    GetDataStock
  Else
    cNama.Default
  End If
End Sub

Private Sub cNamaMember_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "anggota a", "a.kodeanggota,a.nama,a.kodedep,a.alamat,a.telp,a.dd,d.keterangan", "a.nama", sisContent, cNamaMember.Text, , "a.kodeanggota,a.nama", Array("Left join dep d on d.kodedep = a.kodedep"))
  If Not dbData.EOF Then
    cNamaMember.Text = cNamaMember.Browse(dbData, Array("KODE", "NAMA", "DEP", "ALAMAT"), , Array(10, 20, 6, 10))
    cNamaMember.Text = GetNull(dbData!kodeanggota)
    cNamaMember.Text = GetNull(dbData!nama, "")
    cKodeMember.Text = GetNull(dbData!kodeanggota)
    cAlamatMember.Text = GetNull(dbData!alamat, "")
    cTelpMember.Text = GetNull(dbData!telp)
    GetLoadOrderan dTgl.Value, cKodeMember.Text, 1
  End If
End Sub

Private Sub GetLoadOrderan(ByVal Tgl As Date, ByVal cMember As String, ByVal Status As String)
Dim n As Single
Dim nJumlah1, nQtyTmp As Double

    'ambil data orderan member
    vaArray.ReDim 0, -1, 0, 6
    Set dbData = objData.Browse(GetDSN, "orderan o", "o.barcode,s.nama,o.kodestock,o.debet,s.hargajual", "o.kodeanggota", sisAssign, cMember, "and o.tgl = '" & Format(Tgl, "yyyy-MM-dd") & "' AND o.status ='" & Status & "'", , Array("LEFT JOIN stock s on s.kodestock = o.kodestock"))
    If Not dbData.EOF Then
      Do While Not dbData.EOF
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        n = vaArray.UpperBound(1)
        vaArray(n, 0) = n + 1
        vaArray(n, 1) = GetNull(dbData!Barcode)
        vaArray(n, 2) = GetNull(dbData!nama)
        vaArray(n, 3) = GetNull(dbData!KodeStock)
        vaArray(n, 4) = GetNull(dbData!debet)
        vaArray(n, 5) = GetNull(dbData!hargajual)
        vaArray(n, 6) = vaArray(n, 4) * vaArray(n, 5)
        dbData.MoveNext
      Loop
    Else
      MsgBox "Maaf member an " & cNamaMember.Text & vbCrLf & _
      "tidak ada WAITING LIST di tgl " & Format(dTgl.Value, "dd-MM-yyyy"), vbExclamation
    End If
    
    nJumlah1 = 0
    For n = 0 To vaArray.UpperBound(1)
      nJumlah1 = nJumlah1 + vaArray(n, 6)
      nQtyTmp = nQtyTmp + vaArray(n, 4)
    Next
    TDBGrid1.Columns(4).FooterText = Format(nQtyTmp, "###,###,##0")
    TDBGrid1.Columns(6).FooterText = Format(nJumlah1, "###,###,##0")
    Set TDBGrid1.Array = vaArray
    TDBGrid1.ReBind
    TDBGrid1.Refresh
End Sub

Private Sub dTgl_Validate(Cancel As Boolean)
  If cKodeMember.Text <> Trim("") Then
    GetLoadOrderan dTgl.Value, cKodeMember.Text, 1
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  
  initvalue
    
  
  TabIndex dTgl, n
  TabIndex cNamaMember, n
  
  TabIndex nNomor, n
  TabIndex cBarcode, n
  TabIndex cNama, n
  TabIndex nQty, n
  TabIndex nHarga, n
  TabIndex nJumlah, n
  TabIndex cmdOK, n
  
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
End Sub

Private Sub initvalue()

  dTgl.Value = Date
  cNamaMember.Default
  cKodeMember.Default
  cAlamatMember.Default
  cTelpMember.Default
  
  vaArray.ReDim 0, -1, 0, 6
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
  
  InitValue1
  
  TDBGrid1.Columns(4).FooterText = ""
  TDBGrid1.Columns(6).FooterText = ""

End Sub

Private Sub InitValue1()
  nNomor.Value = 1
  cBarcode.Default
  cNama.Default
  nQty.Value = 1
  nHarga.Value = 0
  nJumlah.Value = 0
  cKode.Default
End Sub

Private Sub nNomor_Validate(Cancel As Boolean)
Dim n As Single
  
  If GetValidNomorUrut(nNomor, vaArray) Then
    n = nNomor.Value - 1
    If n <= vaArray.UpperBound(1) Then
      cBarcode.Text = vaArray(n, 1)
      cNama.Text = vaArray(n, 2)
      cKode.Text = vaArray(n, 3)
      nQty.Value = vaArray(n, 4)
      nHarga.Value = vaArray(n, 5)
      nJumlah.Value = vaArray(n, 6)
    End If
  End If
End Sub

Private Sub nQty_Validate(Cancel As Boolean)
  nJumlah.Value = nQty.Value * nHarga.Value
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim n As Integer
Dim nQtyTmp As Integer
Dim nTempJumlah As Double


  If KeyCode = vbKeyDelete Then
    If vaArray.UpperBound(1) >= 0 Then
      TDBGrid1.Delete
      For n = 0 To vaArray.UpperBound(1)
        vaArray(n, 0) = n + 1
        nQtyTmp = nQtyTmp + vaArray(n, 4)
        nTempJumlah = nTempJumlah + vaArray(n, 6)
      Next
      
      TDBGrid1.Columns(4).FooterText = Format(nQtyTmp, "###,###,##0")
      TDBGrid1.Columns(6).FooterText = Format(nTempJumlah, "###,###,##0")
      nNomor.Value = vaArray.UpperBound(1) + 2
      TDBGrid1.ReBind
    End If
  End If
End Sub



