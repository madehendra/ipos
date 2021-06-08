VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trTukarSize 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tukar Size Form"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13980
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   13980
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   735
      Left            =   0
      Top             =   6660
      Width           =   13980
      _ExtentX        =   24659
      _ExtentY        =   1296
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
         TabIndex        =   0
         Top             =   150
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
         Picture         =   "trTukarSize.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   3390
         TabIndex        =   1
         Top             =   150
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
         Picture         =   "trTukarSize.frx":028A
      End
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   1170
         TabIndex        =   2
         Top             =   150
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
         Picture         =   "trTukarSize.frx":0429
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   435
         Left            =   105
         TabIndex        =   3
         Top             =   150
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
         Picture         =   "trTukarSize.frx":0555
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   12810
         TabIndex        =   4
         Top             =   150
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
         Picture         =   "trTukarSize.frx":0700
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   11730
         TabIndex        =   5
         Top             =   150
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
         Picture         =   "trTukarSize.frx":07A6
      End
   End
   Begin BiSAFramProject.BiSAFrame BisaFrame2 
      Height          =   4950
      Left            =   30
      Top             =   1725
      Width           =   13980
      _ExtentX        =   24659
      _ExtentY        =   8731
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
         Left            =   5025
         TabIndex        =   6
         Top             =   75
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   582
         Appearance      =   0
         Decimals        =   0
         MinValue        =   0
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
      Begin BiSATextBoxProject.BiSABrowse cNama 
         Height          =   330
         Index           =   0
         Left            =   2220
         TabIndex        =   7
         Top             =   75
         Width           =   2790
         _ExtentX        =   4921
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
         Index           =   0
         Left            =   645
         TabIndex        =   8
         Top             =   75
         Width           =   1560
         _ExtentX        =   2752
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
         Left            =   105
         TabIndex        =   9
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
         Height          =   4440
         Left            =   105
         TabIndex        =   10
         Top             =   420
         Width           =   13380
         _ExtentX        =   23601
         _ExtentY        =   7832
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "NO."
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "BARCODE"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "NAMA"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "QTY"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "BARCODE"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "NAMA"
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "BIAYA"
         Columns(6).DataField=   ""
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "TOTAL BIAYA"
         Columns(7).DataField=   ""
         Columns(7).NumberFormat=   "###,###,###"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   8
         Splits(0)._UserFlags=   0
         Splits(0).PartialRightColumn=   0   'False
         Splits(0).MarqueeStyle=   3
         Splits(0).AllowRowSizing=   0   'False
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AllowColSelect=   0   'False
         Splits(0).AllowRowSelect=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=926"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=847"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=514"
         Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=2778"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2699"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=512"
         Splits(0)._ColumnProps(13)=   "Column(1).WrapText=1"
         Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(15)=   "Column(2).Width=4921"
         Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=4842"
         Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(19)=   "Column(2)._ColStyle=512"
         Splits(0)._ColumnProps(20)=   "Column(2).WrapText=1"
         Splits(0)._ColumnProps(21)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(22)=   "Column(3).Width=1508"
         Splits(0)._ColumnProps(23)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(3)._WidthInPix=1429"
         Splits(0)._ColumnProps(25)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(26)=   "Column(3)._ColStyle=197122"
         Splits(0)._ColumnProps(27)=   "Column(3).WrapText=1"
         Splits(0)._ColumnProps(28)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(29)=   "Column(4).Width=2805"
         Splits(0)._ColumnProps(30)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(4)._WidthInPix=2725"
         Splits(0)._ColumnProps(32)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(33)=   "Column(4)._ColStyle=512"
         Splits(0)._ColumnProps(34)=   "Column(4).WrapText=1"
         Splits(0)._ColumnProps(35)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(36)=   "Column(5).Width=5133"
         Splits(0)._ColumnProps(37)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(5)._WidthInPix=5054"
         Splits(0)._ColumnProps(39)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(40)=   "Column(5)._ColStyle=512"
         Splits(0)._ColumnProps(41)=   "Column(5).WrapText=1"
         Splits(0)._ColumnProps(42)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(43)=   "Column(6).Width=2381"
         Splits(0)._ColumnProps(44)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(45)=   "Column(6)._WidthInPix=2302"
         Splits(0)._ColumnProps(46)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(47)=   "Column(6)._ColStyle=516"
         Splits(0)._ColumnProps(48)=   "Column(6).WrapText=1"
         Splits(0)._ColumnProps(49)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(50)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(51)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(52)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(53)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(54)=   "Column(7)._ColStyle=516"
         Splits(0)._ColumnProps(55)=   "Column(7).WrapText=1"
         Splits(0)._ColumnProps(56)=   "Column(7).Order=8"
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
         InsertMode      =   0   'False
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
         _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1,.wraptext=-1"
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
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=66,.parent=13,.alignment=1"
         _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=14"
         _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=15,.alignment=1"
         _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=17"
         _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=0"
         _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
         _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
         _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=0"
         _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=54,.parent=13"
         _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=51,.parent=14"
         _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=52,.parent=15"
         _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=53,.parent=17"
         _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=74,.parent=13"
         _StyleDefs(66)  =   "Splits(0).Columns(7).HeadingStyle:id=71,.parent=14"
         _StyleDefs(67)  =   "Splits(0).Columns(7).FooterStyle:id=72,.parent=15"
         _StyleDefs(68)  =   "Splits(0).Columns(7).EditorStyle:id=73,.parent=17"
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
      Begin BiSAButtonProject.BiSAButton cmdOK 
         Height          =   330
         Left            =   13500
         TabIndex        =   11
         Top             =   75
         Width           =   420
         _ExtentX        =   741
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
         Picture         =   "trTukarSize.frx":0A2C
      End
      Begin BiSATextBoxProject.BiSABrowse cBarcode 
         Height          =   330
         Index           =   1
         Left            =   5880
         TabIndex        =   17
         Top             =   75
         Width           =   1575
         _ExtentX        =   2778
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
      Begin BiSATextBoxProject.BiSABrowse cNama 
         Height          =   330
         Index           =   1
         Left            =   7455
         TabIndex        =   18
         Top             =   75
         Width           =   2925
         _ExtentX        =   5159
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
      Begin BiSANumberBoxProject.BiSANumberBox nBiaya 
         Height          =   330
         Left            =   10395
         TabIndex        =   19
         Top             =   75
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
         Appearance      =   0
         Decimals        =   0
         MinValue        =   0
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
      Begin BiSANumberBoxProject.BiSANumberBox nTotalBiaya 
         Height          =   330
         Left            =   11730
         TabIndex        =   20
         Top             =   75
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   582
         Appearance      =   0
         Decimals        =   0
         MinValue        =   0
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
      Height          =   1770
      Left            =   0
      Top             =   0
      Width           =   13980
      _ExtentX        =   24659
      _ExtentY        =   3122
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
      Begin BiSATextBoxProject.BiSABrowse cNamaCustomer 
         Height          =   330
         Left            =   3675
         TabIndex        =   12
         Top             =   420
         Width           =   3750
         _ExtentX        =   6615
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
      Begin BiSATextBoxProject.BiSATextBox cAlamat 
         Height          =   330
         Left            =   75
         TabIndex        =   13
         Top             =   735
         Width           =   4035
         _ExtentX        =   7117
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
      Begin BiSATextBoxProject.BiSABrowse cCustomer 
         Height          =   330
         Left            =   75
         TabIndex        =   14
         Top             =   420
         Width           =   3600
         _ExtentX        =   6350
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
         TabIndex        =   15
         Top             =   105
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   582
         Value           =   "03-04-2011"
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
      Begin BiSATextBoxProject.BiSABrowse cFaktur 
         Height          =   330
         Left            =   75
         TabIndex        =   16
         Top             =   1050
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
   End
End
Attribute VB_Name = "trTukarSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public cNoOrder As String

Dim nPos As SisPos
Dim lEdit As Boolean
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub GetFakturBrowse(ByVal lStat As Boolean)
  cFaktur.Button = lStat
End Sub


Private Sub cFaktur_ButtonClick()
Dim db As New ADODB.Recordset
Dim n As Integer
Dim lSave As Boolean
Dim nQtyTmp As Single

lSave = True
nQtyTmp = 0

  Set db = objData.Browse(GetDSN, "totpenjualan", "nomorpenjualan,tgl,subtotal,total,piutang", "nomorpenjualan", sisContent, cFaktur.Text, " and tgl = '" & Format(dTgl.Value, "yyyy-MM-dd") & "' and kodeanggota = '" & cCustomer.Text & "'")
  If Not db.EOF Then
    cFaktur.Text = cFaktur.Browse(db)
    'ambil nilai total
    Set db = objData.Browse(GetDSN, "totpenjualan t", "t.*,g.keterangan as namagudang", "t.nomorpenjualan", sisAssign, cFaktur.Text, , , Array("left join gudang g on g.kodegudang = t.kodegudang"))
    If Not db.EOF Then

    End If
    
    'ambil nilai detail
    Set db = objData.Browse(GetDSN, "penjualan p", "s.barcode,p.kodestock,s.Nama,p.qty,p.kodesatuan,p.harga,p.discount,p.jumlah", "nomorpenjualan", sisAssign, cFaktur.Text, , "p.urutfaktur asc", Array("Left join stock s on s.kodestock = p.kodestock"))
    If Not db.EOF Then
      vaArray.ReDim 0, -1, 0, 9
      Do While Not db.EOF
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        n = vaArray.UpperBound(1)
        vaArray(n, 0) = n + 1
        vaArray(n, 1) = GetNull(db!Barcode)
        vaArray(n, 2) = GetNull(db!nama)
        vaArray(n, 3) = GetNull(db!qty)
        vaArray(n, 4) = GetNull(db!kodesatuan)
        vaArray(n, 5) = GetNull(db!Harga)
        vaArray(n, 6) = GetNull(db!Discount)
        vaArray(n, 7) = GetNull(db!jumlah)
        vaArray(n, 8) = GetNull(db!KodeStock)
        nQtyTmp = nQtyTmp + vaArray(n, 3)
        db.MoveNext
      Loop
      Set TDBGrid1.Array = vaArray
      TDBGrid1.ReBind
      TDBGrid1.Refresh
      TDBGrid1.Columns(3).FooterText = Format(nQtyTmp, "###,###,##0")
      Me.Refresh
      nNomor.Value = vaArray.UpperBound(1) + 2
    End If
    
    If nPos = Delete Then
      If MsgBox("Data akan dihapus?", vbYesNo) = vbYes Then
        objData.Start GetDSN
        
        'Patch
        Dim cSQL As String
        cSQL = ""
        cSQL = " select distinct(nomorpelunasanpiutang) as nomorpelunasanpiutang from pelunasanpiutang where nomorpenjualan = '" & cFaktur.Text & "'"
        Set db = objData.Sql(GetDSN, cSQL)
        If Not db.EOF Then
          
          If MsgBox("Transaksi ini sudah pernah dilunasi sebelumnya!" & vbCrLf & "Dengan menghapus berarti seluruh data pelunasan yg berkenaan dengan transaksi ini akan ikut terhapus juga" & vbCrLf & "Apakah anda yakin akan menghapus?", vbYesNo) = vbYes Then
            'rutin menghapus pada modul pelunasan piutang
            lSave = IIf(lSave, DelKodeTr(objData, msPelunasanPiutang, GetNull(db!nomorpelunasanpiutang)), False)
            lSave = IIf(lSave, objData.Delete(GetDSN, "pelunasanpiutang", "nomorpelunasanpiutang", sisAssign, GetNull(db!nomorpelunasanpiutang)), False)
            lSave = IIf(lSave, objData.Delete(GetDSN, "kartupiutang", "nomorkartupiutang", sisAssign, GetNull(db!nomorpelunasanpiutang)), False)
            lSave = IIf(lSave, objData.Delete(GetDSN, "totpelunasanpiutang", "nomorpelunasanpiutang", sisAssign, GetNull(db!nomorpelunasanpiutang)), False)
            lSave = IIf(lSave, objData.Delete(GetDSN, "bukubesar", "faktur", sisAssign, GetNull(db!nomorpelunasanpiutang)), False)
          Else
            MsgBox "Penghapusan Dibatalkan"
            GetEdit False
            initvalue
            objData.Cancel GetDSN
            Exit Sub
          End If
          
          
        End If
       
        cSQL = ""
        cSQL = " select * from totrtnpenjualan where nomorpenjualan = '" & cFaktur.Text & "'"
        Set db = objData.Sql(GetDSN, cSQL)
        If Not db.EOF Then
          If MsgBox("Transaksi ini masih dirujuk oleh retur penjualan!" & vbCrLf & "Dengan menghapus berarti seluruh rujukan pada retur penjualan akan ikut dihapus pula" & vbCrLf & "Apakah anda yakin akan menghapus?", vbYesNo) = vbYes Then
            Do While Not db.EOF
              lSave = IIf(lSave, objData.Edit(GetDSN, "totrtnpenjualan", "nomorreturpenjualan = '" & GetNull(db!nomorreturpenjualan) & "'", Array("nomorpenjualan"), Array("")), False)
              db.MoveNext
            Loop
          Else
            MsgBox "Penghapusan dibatalkan"
            GetEdit False
            initvalue
            objData.Cancel GetDSN
            Exit Sub
          End If
        End If
        
        'end patch
        
        'Update dulu ke table order
        lSave = IIf(lSave, objData.Edit(GetDSN, "totmemberorder", "nomorpenjualan = '" & cFaktur.Text & "'", Array("nomorpenjualan", "status"), Array("", 0)), False)

        'Rutin menghapus transaksi penjualan
        lSave = IIf(lSave, DelKodeTr(objData, msPenjualan, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "penjualan", "nomorpenjualan", sisAssign, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "kartupiutang", "nomorkartupiutang", sisAssign, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "kartustock", "nomor", sisAssign, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "totpenjualan", "nomorpenjualan", sisAssign, cFaktur.Text), False)
        
        If lSave Then
          objData.Save GetDSN
        Else
          objData.Cancel GetDSN
        End If
        
      End If
      initvalue
      GetEdit False
    End If
    If nPos = Edit Then
      SendKeysA vbKeyReturn, True
    End If
  End If
End Sub

Private Sub cFaktur_Validate(Cancel As Boolean)
  If nPos = Edit Then
      cFaktur.Enabled = False
      cFaktur.BackColor = vbButtonFace
  End If
End Sub

Private Sub cmdAddOrder_Click()
Dim db As New ADODB.Recordset
Dim nJumlah1 As Double
Dim n As Integer

  frmLoadOrder.cKodeMember = cCustomer.Text
  frmLoadOrder.Show vbModal
  
  If cNoOrder <> "" Then
    vaArray.ReDim 0, -1, 0, 9
    Set db = objData.Browse(GetDSN, "memberorder m", "m.*,s.nama as namastock,s.barcode,s.kodesatuan,s.jenis", "t.nomormemberorder", sisAssign, cNoOrder, , "m.nourut asc", Array("LEFT JOIN totmemberorder t on t.nomormemberorder = m.nomormemberorder", "LEFT JOIN stock s on s.kodestock = m.kodestock"))
    If Not db.EOF Then
      Do While Not db.EOF
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        n = vaArray.UpperBound(1)
        vaArray(n, 0) = n + 1
        vaArray(n, 1) = GetNull(db!Barcode)
        vaArray(n, 2) = GetNull(db!namastock)
        vaArray(n, 3) = GetNull(db!qty)
        vaArray(n, 4) = GetNull(db!kodesatuan)
        vaArray(n, 5) = GetNull(db!Harga)
        vaArray(n, 6) = GetNull(db!Discount)
        vaArray(n, 7) = GetNull(db!jumlah)
        vaArray(n, 8) = GetNull(db!KodeStock)
        vaArray(n, 9) = GetNull(db!jenis)
        db.MoveNext
      Loop
    End If
    
    TDBGrid1.Array = vaArray
    TDBGrid1.ReBind
    
    nJumlah1 = 0
    For n = 0 To vaArray.UpperBound(1)
      nJumlah1 = nJumlah1 + vaArray(n, 7)
    Next
    
    
    SumTotal
    
    InitValue1
    
    nNomor.Value = vaArray.UpperBound(1) + 2
  End If
End Sub

Private Sub cmdAktivasi_Click()
  frmAktivasi.Action Me
End Sub

Private Sub GetDataStock()
Dim db As New ADODB.Recordset
  
End Sub

Private Sub cmdAdd_Click()
Dim i As Integer

  nPos = Add
  GetEdit True
  GetFakturBrowse False
  cFaktur.Text = GetNomor("totpenjualan", "nomorpenjualan", GetID, SisModulTransaksi.penjualan)
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


Private Function validOK() As Boolean
Dim nKe As Integer

  validOK = True
  SumTotal
  InitValue1
  
End Function

Private Sub cmdOK_Click()
Dim n As Integer
Dim nJumlah1 As Double
Dim nQtyTmp As Single

  
  If validOK() Then
    'jika baris <= Nomor
    If vaArray.UpperBound(1) + 2 <= nNomor.Value Then
      vaArray.ReDim 0, nNomor.Value - 1, 0, 9
      n = vaArray.UpperBound(1)
    'jika baris=0
    ElseIf vaArray.UpperBound(1) = -1 Then
      nNomor.Value = 1
      vaArray.ReDim 0, nNomor.Value - 1, 0, 9
      n = vaArray.UpperBound(1)
    Else
      n = nNomor.Value - 1
    End If
    
    TDBGrid1.Array = vaArray
    TDBGrid1.ReBind
    TDBGrid1.MoveNext
    
    
    nJumlah1 = 0
    nQtyTmp = 0
    For n = 0 To vaArray.UpperBound(1)
      nJumlah1 = nJumlah1 + vaArray(n, 7)
      nQtyTmp = nQtyTmp + vaArray(n, 3)
    Next
  
    TDBGrid1.Columns(3).FooterText = Format(nQtyTmp, "###,###,##0")
    SumTotal
    InitValue1
    nNomor.Value = vaArray.UpperBound(1) + 2
    
  End If
End Sub

Private Sub SumTotal()
Dim n As Double
    

End Sub

Private Function ValidSaving() As Boolean
  ValidSaving = True
End Function

Private Sub cmdSimpan_Click()
Dim lSave As Boolean
Dim Faktur As String
Dim n As Integer
Dim i As Integer
Dim nValueTunai As Double
Dim nValueKredit As Double

lSave = True
'
'  If isValidSaving Then
'    objData.Start GetDSN
'    Faktur = cFaktur.Text
'    If nPos = Add Then
'      If Not GetAvailable(cFaktur.Text, "totpenjualan", "nomorpenjualan") Then
'        Faktur = GetNomor("totpenjualan", "nomorpenjualan", GetID, SisModulTransaksi.penjualan)
'      End If
'    End If
'
'    lSave = IIf(lSave, objData.Update(GetDSN, "totpenjualan", "nomorpenjualan = '" & Faktur & "'", Array("nomorpenjualan", "fakturasli", "tgl", "jthtmp", "kodeanggota", "ppn", "persdisc", "persdisc2", "subtotal", "pajak", "discount", "discount2", "total", "tunai", "piutang", "datetime", "username", "kodeakun", "kodecostcenter", "flaglunas", "kodesalesman", "komisi", "dp", "kodegudang", "upkepada"), Array(Faktur, cFakturAsli.Text, Format(dTgl.Value, "yyyy-MM-dd"), Format(dJthTmp.Value, "yyyy-MM-dd"), cCustomer.Text, nPPn.Value, nPersDisc.Value, 0, nSubTotal.Value, nPajak.Value, nDiscount.Value, 0, nTotal.Value, nTunai.Value, nPiutang.Value, SNow, GetRegistry(reg_UserName), cAkunKas.Text, aCfg(objData, msCostCenterJualBeli), 0, cSalesman.Text, nKomisi.Value, nDP.Value, cGudang.Text, cUp.Text)), False)
'    lSave = IIf(lSave, objData.Delete(GetDSN, "penjualan", "nomorpenjualan", sisAssign, Faktur), False)
'    lSave = IIf(lSave, objData.Delete(GetDSN, "kartustock", "nomor", sisAssign, Faktur), False)
'
'    'Update status order menjadi 1
'    lSave = IIf(lSave, objData.Edit(GetDSN, "totmemberorder", "nomormemberorder = '" & cNoOrder & "'", Array("status", "nomorpenjualan"), Array(1, Faktur)), False)
'    For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
'      nValueTunai = 0
'      nValueKredit = 0
'      If chkTunai.Value = 1 Then
'        nValueTunai = vaArray(n, 5) - (vaArray(n, 5) * vaArray(n, 6) / 100)
'        nValueKredit = 0
'      Else
'        nValueKredit = vaArray(n, 5) - (vaArray(n, 5) * vaArray(n, 6) / 100)
'        nValueTunai = 0
'      End If
'
'      lSave = IIf(lSave, objData.Add(GetDSN, "penjualan", Array("nomorpenjualan", "kodegudang", "tgl", "kodestock", "qty", "harga", "kodesatuan", "discount", "jumlah", "hb", "tunai", "piutang", "urutfaktur"), Array(Faktur, cGudang.Text, Format(dTgl.Value, "yyyy-MM-dd"), vaArray(n, 8), vaArray(n, 3), vaArray(n, 5), vaArray(n, 4), vaArray(n, 6), vaArray(n, 7), GetHargaBeli(objData, vaArray(n, 8)), nValueTunai, nValueKredit, vaArray(n, 0))), False)
'      lSave = IIf(lSave, UpdKartuStock(objData, SisKartuStock.penjualan, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), vaArray(n, 8), vaArray(n, 3), vaArray(n, 5), vaArray(n, 6), "Penjualan Non Tunai an. " & cNamaCustomer.Text, cGudang.Text), False)
'
'      'Update status lunas
'      If lCekStatusLunas(objData, Faktur) = True Then
'        lSave = IIf(lSave, objData.Edit(GetDSN, "penjualan", "nomorpenjualan = '" & Faktur & "'", Array("statuslunas"), Array(1)), False)
'      Else
'        lSave = IIf(lSave, objData.Edit(GetDSN, "penjualan", "nomorpenjualan = '" & Faktur & "'", Array("statuslunas"), Array(0)), False)
'      End If
'
'    Next n
'
'    'isi field flaglunas
'    'cek apakah dp yg dibayarkan lebih dari/sama dengan yg diminta
'
'    If nDP.Value >= nTotal.Value Then
'      lSave = IIf(lSave, objData.Edit(GetDSN, "totpenjualan", "nomorpenjualan = '" & Faktur & "'", Array("flaglunas"), Array(1)), False)
'    End If
'
'    If chkTunai.Value = 1 Then
'      lSave = IIf(lSave, objData.Edit(GetDSN, "totpenjualan", "nomorpenjualan = '" & Faktur & "'", Array("flaglunas"), Array(1)), False)
'    Else
'      If lCekStatusLunas(objData, Faktur) = True Then
'        lSave = IIf(lSave, objData.Edit(GetDSN, "totpenjualan", "nomorpenjualan = '" & Faktur & "'", Array("flaglunas"), Array(1)), False)
'      End If
'    End If
'
'    lSave = IIf(lSave, UpdKartuHutang(objData, SisPenjualan, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), cCustomer.Text, "Penjualan Non Tunai an. " & cNamaCustomer.Text, nPiutang.Value, SNow, GetRegistry(reg_UserName)), False)
'
'    'jika dibayar tunai dan ada dp maka posting ke kartupiutang
'
'    If chkTunai.Value = 1 And nDP.Value <> 0 Then
'      lSave = IIf(lSave, UpdKartuHutang(objData, SisPenjualan, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), cCustomer.Text, "Pengembalian DP dengan barang an. " & cNamaCustomer.Text, nDP.Value, SNow, GetRegistry(reg_UserName)), False)
'    End If
'
'    If chkTunai.Value <> 1 And nDP.Value <> 0 Then
'      lSave = IIf(lSave, UpdKartuHutang(objData, SisPenjualan, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), cCustomer.Text, "Pengembalian DP dengan barang an. " & cNamaCustomer.Text, nTotal.Value, SNow, GetRegistry(reg_UserName)), False)
'    End If
'
'    lSave = IIf(lSave, DelKodeTr(objData, vbTrigger.msPenjualan, Faktur), False)
'
'    'Piutang, Kas
'    'Kas, piutang
'    '   Penjualan
'
'    'Diskon Penjualan
'    '   Penjualan
'
'    'PPn Penjualan
'    '   Penjualan
'
'    'Inventory
'
'
'    lSave = IIf(lSave, UpdKodeTr(objData, msPenjualan, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), GetAkunMember(objData, cCustomer.Text), aCfg(objData, msCostCenterJualBeli), "Penjualan an " & cNamaCustomer.Text, nPiutang.Value + nDP.Value), False)
'    lSave = IIf(lSave, UpdKodeTr(objData, msPenjualan, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), cAkunKas.Text, aCfg(objData, msCostCenterJualBeli), "Penjualan an " & cNamaCustomer.Text, nTunai.Value), False)
'        lSave = IIf(lSave, UpdKodeTr(objData, msPenjualan, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningPenjualan), aCfg(objData, msCostCenterJualBeli), "Penjualan an " & cNamaCustomer.Text, 0, nTotal.Value), False)
'
'    'Posting balik dp yg sudah pernah dilakukan:
'
'    'Debet
'    Dim nTmp As Double
'    Dim nSaldoTmp As Double
'    Dim nTmpCOGS As Double
'    Dim nTmpSaldoCOGS As Double
'    Dim db As New ADODB.Recordset
'
'    For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
'      'Discount Pembelian per item
'      nTmp = vaArray(n, 3) * vaArray(n, 5) - vaArray(n, 7)
'      nSaldoTmp = nSaldoTmp + nTmp
'      lSave = IIf(lSave, UpdKodeTr(objData, msPenjualan, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningDiscountPenjualan), aCfg(objData, msCostCenterJualBeli), "Dsc Item Penjualan an " & cNamaCustomer.Text, nTmp, 0, "", SNow), False)
'
'      Set db = objData.Browse(GetDSN, "stock", "kodestock,asbiaya", "kodestock", sisAssign, vaArray(n, 8))
'      If Not db.EOF Then
'        If GetNull(db!asbiaya) <> "1" Then
'          'posting cogs
'          nTmpCOGS = vaArray(n, 3) * GetHargaPokok(objData, vaArray(n, 8))
'          lSave = IIf(lSave, UpdKodeTr(objData, msPenjualan, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningCOGS), aCfg(objData, msCostCenterJualBeli), "COGS Penjualan an " & vaArray(n, 2), nTmpCOGS, 0, "", SNow), False)
'            lSave = IIf(lSave, UpdKodeTr(objData, msPenjualan, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), GetAkunInventory(objData, vaArray(n, 8)), aCfg(objData, msCostCenterJualBeli), "COGS Penjualan an " & vaArray(n, 2), 0, nTmpCOGS, "", SNow), False)
'        End If
'      End If
'    Next n
'
'    'Kredit
'    lSave = IIf(lSave, UpdKodeTr(objData, msPenjualan, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningPenjualan), aCfg(objData, msCostCenterJualBeli), "Dsc Item Penjualan an  " & cNamaCustomer.Text, 0, nSaldoTmp), False)
'
'    'PPn
'    lSave = IIf(lSave, UpdKodeTr(objData, msPenjualan, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, SisCfg.msRekeningPPnPenjualan), aCfg(objData, msCostCenterJualBeli), "PPn Penjualan an " & cNamaCustomer.Text, 0, nPajak.Value, "", SNow), False)
'        lSave = IIf(lSave, UpdKodeTr(objData, msPenjualan, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningPenjualan), aCfg(objData, msCostCenterJualBeli), "PPn Penjualan an " & cNamaCustomer.Text, nPajak.Value, 0), False)
'
'    'Discount seluruhnya
'    lSave = IIf(lSave, UpdKodeTr(objData, msPenjualan, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningDiscountPenjualan), aCfg(objData, msCostCenterJualBeli), "Dsc Total Penjualan an " & cNamaCustomer.Text, nDiscount.Value, 0, "", SNow), False)
'        lSave = IIf(lSave, UpdKodeTr(objData, msPenjualan, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningPenjualan), aCfg(objData, msCostCenterJualBeli), "Dsc Total Penjualan an " & cNamaCustomer.Text, 0, nDiscount.Value, "", SNow), False)
'
'    'Komisi salesman
'    ' Hutang komisi
'    lSave = IIf(lSave, UpdKodeTr(objData, msPenjualan, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningBiayaKomisi), aCfg(objData, msCostCenterJualBeli), "Komisi Penjualan Sales " & cSalesman.Text, nKomisi.Value, 0, "", SNow), False)
'        lSave = IIf(lSave, UpdKodeTr(objData, msPenjualan, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningHutangSalesman), aCfg(objData, msCostCenterJualBeli), "Komisi Penjualan Sales " & cSalesman.Text, 0, nKomisi.Value, "", SNow), False)
'
'
'    If lSave Then
'      objData.Save GetDSN
'    Else
'      objData.Cancel GetDSN
'    End If
'
'    If lSave = True Then
'      If MsgBox("Apakah akan mencetak transaksi ke printer?", vbYesNo + vbInformation) = vbYes Then
'        If aCfg(objData, msCetakanPenjualanNonTunai) = 1 Then
'          'Penjualan Nota
'          GetCetakFakturpenjualan objData, Faktur, False
'          Unload frmFaktur
'        Else
'          'print struk
'          If GetRegistry(reg_PortPrinterKasir) = "USB" Then
'            For i = 1 To aCfg(objData, msJumlahCetakanPenjualanNonTunai)
'              trPrint2.noOrder = Faktur
'              Set dbData = objData.Browse(GetDSN, "totpenjualan t", "t.*,a.*", "t.nomorpenjualan", sisAssign, Faktur, , , Array("left join anggota a on a.kodeanggota = t.kodeanggota"))
'              If Not dbData.EOF Then
'                trPrint2.nSubTotal = GetNull(dbData!Subtotal)
'                trPrint2.nDiscount = GetNull(dbData!dp)
'                trPrint2.nCash = GetNull(dbData!Tunai)
'                trPrint2.nChange = GetNull(dbData!Piutang)
'                trPrint2.cKodeMember = GetNull(dbData!kodeanggota)
'                trPrint2.cMember = GetNull(dbData!nama)
'                trPrint2.cTeleponMember = GetNull(dbData!telp)
'                trPrint2.Ups = GetNull(dbData!upkepada)
'                Load trPrint2
'                trPrint2.Show vbModal
'              End If
'            Next i
'          Else
'            For n = 1 To aCfg(objData, msJumlahCetakanPenjualanNonTunai)
'              If MsgBox("Tekan enter untuk melanjutkan pencetakan ke-" & n, vbYesNo + vbInformation) = vbYes Then
'                PrintStruk Faktur
'              End If
'            Next n
'          End If
'        End If
'      End If
'    Else
'      MsgBox "Maaf, terjadi masalah dalam proses penyimpanan" & vbCrLf & "Data tidak bisa disimpan"
'    End If
'    initvalue
'    GetEdit False
'  End If

End Sub

Private Function isValidSaving() As Boolean
Dim dba As New ADODB.Recordset
Dim nPernahBayar As Double

isValidSaving = True
  
End Function

Private Sub cNamaCustomer_ButtonClick()
Dim vaTmp As New XArrayDB

  Set dbData = objData.Browse(GetDSN, "anggota a", "a.kodeanggota,a.nama,a.kodedep,a.alamat,a.telp,d.keterangan", "a.nama", sisContent, cNamaCustomer.Text, , "a.kodeanggota,a.nama", Array("Left join dep d on d.kodedep = a.kodedep"))
  If Not dbData.EOF Then
    cNamaCustomer.Text = cNamaCustomer.Browse(dbData, Array("Kode", "Nama", "Dep", "Alamat"), , Array(6, 15, 6, 15))
    cCustomer.Text = GetNull(dbData!kodeanggota)
    
    cNamaCustomer.Text = GetNull(dbData!nama, "")
    cAlamat.Text = GetNull(dbData!alamat, "")

  End If
End Sub

Private Sub Form_Load()
Dim n As Single

'  If CheckTrial(nRecordsTrial, TrialPenjualan) = True Then
'    End
'  End If
  
  
  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  initvalue
  GetEdit False
  
  Set dbData = objData.Browse(GetDSN, "costcenter", , "kodecostcenter", sisAssign, aCfg(objData, msCostCenterJualBeli))
  If Not dbData.EOF Then
  End If
  
  TabIndex dTgl, n
  TabIndex cCustomer, n
  TabIndex cNamaCustomer, n
  TabIndex cAlamat, n
  TabIndex cFaktur, n
  TabIndex nNomor, n
  TabIndex cBarcode, n
  TabIndex cNama, n
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
Dim dbgudang As New ADODB.Recordset
  
  
  cFaktur.Default
  dTgl.Value = Date
  cCustomer.Default
  cNamaCustomer.Default
  cAlamat.Default
  cNoOrder = ""
  vaArray.ReDim 0, -1, 0, 9
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
  InitValue1
  TDBGrid1.Columns(3).FooterText = ""
End Sub

Private Sub InitValue1()
  nNomor.Value = 1
  nQty.Value = 1
End Sub

Private Sub GetEdit(lPar As Boolean)
  BiSAFrame1.Enabled = lPar
  BisaFrame2.Enabled = lPar
  lEdit = lPar
  initvalue
  SetButton cmdSimpan, cmdKeluar, cmdAdd, cmdEdit, cmdHapus, nPos, lPar, cmdAktivasi
  
  If lPar Then
    dTgl.SetFocus
    If nPos = Add Then
      cFaktur.Enabled = False
      cFaktur.BackColor = vbButtonFace
    Else
      cFaktur.Enabled = True
      cFaktur.BackColor = vbWindowBackground
      cFaktur.CaptionBackColor = vbButtonFace
    End If
  End If
End Sub

Private Sub SumJumlah()

End Sub

Private Sub nNomor_Validate(Cancel As Boolean)
Dim n As Single
  
  If GetValidNomorUrut(nNomor, vaArray) Then
    n = nNomor.Value - 1
    If n <= vaArray.UpperBound(1) Then
    End If
  End If
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim n As Integer
Dim nQtyTmp As Integer

  If KeyCode = vbKeyDelete Then
    If vaArray.UpperBound(1) >= 0 Then
      TDBGrid1.Delete
      SumTotal
      For n = 0 To vaArray.UpperBound(1)
        vaArray(n, 0) = n + 1
        nQtyTmp = nQtyTmp + vaArray(n, 3)
      Next
      TDBGrid1.Columns(3).FooterText = Format(nQtyTmp, "###,###,##0")
      nNomor.Value = vaArray.UpperBound(1) + 2
      TDBGrid1.ReBind
    End If
  End If
End Sub


