VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trNewPenyesuaianStock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Penyesuaian Stock"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   11940
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   675
      Left            =   60
      Top             =   5985
      Width           =   11760
      _ExtentX        =   20743
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
      Begin BiSAButtonProject.BiSAButton BiSAButton1 
         Height          =   435
         Left            =   90
         TabIndex        =   19
         Top             =   90
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   767
         Caption         =   "Init Stok Awal"
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
         Picture         =   "trNewPenyesuaianStock.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdHapus 
         Height          =   435
         Left            =   7815
         TabIndex        =   0
         Top             =   120
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
         Picture         =   "trNewPenyesuaianStock.frx":059A
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   8985
         TabIndex        =   1
         Top             =   120
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
         Picture         =   "trNewPenyesuaianStock.frx":0824
      End
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   2355
         TabIndex        =   2
         Top             =   105
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
         Picture         =   "trNewPenyesuaianStock.frx":09C3
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   435
         Left            =   6750
         TabIndex        =   3
         Top             =   120
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
         Picture         =   "trNewPenyesuaianStock.frx":0AEF
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   10560
         TabIndex        =   4
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
         Picture         =   "trNewPenyesuaianStock.frx":0C9A
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   9450
         TabIndex        =   5
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
         Picture         =   "trNewPenyesuaianStock.frx":0D40
      End
   End
   Begin BiSAFramProject.BiSAFrame BisaFrame2 
      Height          =   4035
      Left            =   -15
      Top             =   1950
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   7117
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      BackColor       =   -2147483633
      Begin BiSANumberBoxProject.BiSANumberBox nQty 
         Height          =   330
         Left            =   9255
         TabIndex        =   6
         Top             =   45
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   582
         Appearance      =   0
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
         Left            =   2445
         TabIndex        =   7
         Top             =   45
         Width           =   3060
         _ExtentX        =   5398
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
         Left            =   690
         TabIndex        =   8
         Top             =   45
         Width           =   1710
         _ExtentX        =   3016
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
         Top             =   45
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
         Height          =   3510
         Left            =   105
         TabIndex        =   10
         Top             =   420
         Width           =   11745
         _ExtentX        =   20717
         _ExtentY        =   6191
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
         Columns(4).Caption=   "On Hand (Comp)"
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "###,###,###,###,##0.00"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "On Stock (Real)"
         Columns(5).DataField=   ""
         Columns(5).NumberFormat=   "###,###,###,###,##0.00"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   6
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=6"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=926"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=847"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=514"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=3122"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3043"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=512"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=5503"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=5424"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=512"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=2963"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2884"
         Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=514"
         Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(25)=   "Column(4).Width=3572"
         Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=3493"
         Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=197122"
         Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(31)=   "Column(5).Width=4048"
         Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=3969"
         Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=514"
         Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
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
         BorderStyle     =   0
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1,5
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
         _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=1"
         _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
         _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
         _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
         _StyleDefs(61)  =   "Named:id=33:Normal"
         _StyleDefs(62)  =   ":id=33,.parent=0"
         _StyleDefs(63)  =   "Named:id=34:Heading"
         _StyleDefs(64)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(65)  =   ":id=34,.wraptext=-1"
         _StyleDefs(66)  =   "Named:id=35:Footing"
         _StyleDefs(67)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(68)  =   "Named:id=36:Selected"
         _StyleDefs(69)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(70)  =   "Named:id=37:Caption"
         _StyleDefs(71)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(72)  =   "Named:id=38:HighlightRow"
         _StyleDefs(73)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(74)  =   "Named:id=39:EvenRow"
         _StyleDefs(75)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(76)  =   "Named:id=40:OddRow"
         _StyleDefs(77)  =   ":id=40,.parent=33"
         _StyleDefs(78)  =   "Named:id=41:RecordSelector"
         _StyleDefs(79)  =   ":id=41,.parent=34"
         _StyleDefs(80)  =   "Named:id=42:FilterBar"
         _StyleDefs(81)  =   ":id=42,.parent=33"
      End
      Begin BiSAButtonProject.BiSAButton cmdOK 
         Height          =   330
         Left            =   11475
         TabIndex        =   11
         Top             =   45
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
         Picture         =   "trNewPenyesuaianStock.frx":0FC6
      End
      Begin BiSANumberBoxProject.BiSANumberBox nOnHand 
         Height          =   330
         Left            =   7185
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   45
         Width           =   2040
         _ExtentX        =   3598
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
      Begin BiSANumberBoxProject.BiSANumberBox nHargaBeli 
         Height          =   330
         Left            =   5535
         TabIndex        =   18
         Top             =   45
         Width           =   1620
         _ExtentX        =   2858
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
      Height          =   1890
      Left            =   90
      Top             =   60
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   3334
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      BackColor       =   -2147483633
      Begin BiSATextBoxProject.BiSATextBox cKeterangan 
         Height          =   330
         Left            =   1575
         TabIndex        =   17
         Top             =   1260
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
         Left            =   3345
         TabIndex        =   13
         Top             =   900
         Width           =   2700
         _ExtentX        =   4763
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
         TabIndex        =   14
         Top             =   900
         Width           =   3255
         _ExtentX        =   5741
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
         TabIndex        =   15
         Top             =   210
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
         TabIndex        =   16
         Top             =   555
         Width           =   3300
         _ExtentX        =   5821
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
         Left            =   4155
         Top             =   180
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "trNewPenyesuaianStock"
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


'Dim Excel As Excel.Application
'Dim ExcelWBk As Excel.Workbook
'Dim ExcelWS As Excel.Worksheet

'Private Sub StartExcel()
'  On Error GoTo err:
'  Set Excel = GetObject(, "Excel.Application")
'  Exit Sub
'err:
'  Set Excel = CreateObject("Excel.Application")
'End Sub

'Private Sub CloseWorkSheet()
'  ExcelWBk.Close
'  Excel.Quit
'End Sub

'Private Sub FinishExcel()
'  'Jangan lupa, selalu bersihkan memory saat mengakhiri
'  If Not ExcelWS Is Nothing Then Set ExcelWS = Nothing
'  If Not ExcelWBk Is Nothing Then Set ExcelWBk = Nothing
'  If Not Excel Is Nothing Then Set Excel = Nothing
'End Sub


Private Sub GetFakturBrowse(ByVal lStat As Boolean)
  cNomor.Button = lStat
End Sub

Private Sub BiSAButton1_Click()
'init stock
Dim cMsg As String
cMsg = "Akan memasukkan stok awal?, pastikan data excel sdh di format dengan benar (Kode,Qty,HargaBeli/HPP Awal)"
If MsgBox(cMsg, vbYesNo) = vbYes Then
  CommonDialog1.Filter = "Excel File (*.xls)|*.xls"
  CommonDialog1.ShowOpen
  If CommonDialog1.FileName <> "" Then
'    GetLoadExcel
  End If
End If
End Sub

'Private Sub GetLoadExcel()
'Dim lSave As Boolean
'Dim vaField, vaValue
'Dim i, j, n As Integer
'Dim dbData As New ADODB.Recordset
''Dim Wb As Excel.Workbook
'
'  On Error GoTo err:
''  StartExcel
'  lSave = True
'
''  Excel.Workbooks.Close
''  Set ExcelWBk = Excel.Workbooks.Open(CommonDialog1.FileName)
''  Set ExcelWS = ExcelWBk.Worksheets(1)
'
'
''  FrmPB.InitPB ExcelWS.Cells.SpecialCells(xlCellTypeLastCell).Row
'  Dim cBarcode
'  Dim cQty
'
'  For i = 1 To ExcelWS.Cells.SpecialCells(xlCellTypeLastCell).Row
'    FrmPB.RunPB
'    With ExcelWS
'      Set dbData = objData.Browse(GetDSN, "stock", "kodestock,nama,hargabeli,diskonpenjualan,kodesatuan,barcode", "kodestock", sisAssign, .Cells(i, 1).Value)
'      If Not dbData.EOF Then
'        vaArray.InsertRows vaArray.UpperBound(1) + 1
'        n = vaArray.UpperBound(1)
'        vaArray(n, 0) = n + 1
'        vaArray(n, 1) = .Cells(i, 1).Value  'kodestock
'        vaArray(n, 2) = .Cells(i, 2).Value  'qty
'        vaArray(n, 3) = .Cells(i, 3).Value  'hargabeli
'
'        lSave = IIf(lSave, UpdKartuStock(objData, SisKartuStock.PenyesuaianStock, cNomor.Text, Format(dTgl.Value, "yyyy-MM-dd"), vaArray(n, 1), vaArray(n, 2), vaArray(n, 3), 0, "Init STOCK ", cGudang.Text, vaArray(n, 3)), False)
'        '*********
'        ' JURNAL
'        '--------*
'        ' persediaan (D)
'        '  pendapatan/modal (K)
'
'        lSave = IIf(lSave, UpdKodeTr(objData, msPenyesuaian, cNomor.Text, Format(dTgl.Value, "yyyy-MM-dd"), GetAkunInventory(objData, vaArray(n, 6)), GetCostCenterUser(objData, GetRegistry(reg_username)), "Init Stock " & vaArray(n, 2), vaArray(n, 2) * vaArray(n, 3), 0, "", SNow), False)
'          lSave = IIf(lSave, UpdKodeTr(objData, msPenyesuaian, cNomor.Text, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningPenyesuian), GetCostCenterUser(objData, GetRegistry(reg_username)), "Init Stock " & vaArray(n, 2), 0, vaArray(n, 2) * vaArray(n, 3), "", SNow), False)
'
'      Else
'        'jika data yg di import tidak ada dalam database simpan
'      End If
'    End With
'  Next i
'  If lSave Then
'    objData.Save GetDSN
'  Else
'    objData.Cancel GetDSN
'  End If
'  FrmPB.EndPB
'  CloseWorkSheet
'  FinishExcel
'
'err:
'End Sub

Private Sub cBarcode_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "stock s", "s.kodestock,s.barcode,s.hargabeli,s.nama,s.kodesatuan,s.hargajual,s.hargabeli,s.jenis", "s.barcode", sisContent, cBarcode.Text, " and s.statusnonaktif <> 1")
  If Not dbData.EOF Then
    cBarcode.Text = cBarcode.Browse(dbData)
    cNama.Text = GetNull(dbData!nama)
    cBarcode.Text = GetNull(dbData!barcode)
    nHargaBeli.value = GetNull(dbData!hargabeli)
    nOnHand.value = GetSaldoStock(objData, cGudang.Text, GetNull(dbData!KodeStock), dTgl.value)
    nQty.value = nOnHand.value
    cKode = GetNull(dbData!KodeStock)
  Else
    cBarcode.Default
  End If
End Sub

Private Sub cGudang_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "gudang", "kodegudang,keterangan", "lstatus", sisAssign, "A")
  If Not dbData.EOF Then
    cGudang.Text = cGudang.Browse(dbData) 'GetNull(dbData!kodegudang)
    cNamaGudang.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub cGudang_Validate(Cancel As Boolean)
  cGudang.Enabled = False
End Sub

Private Sub cmdAktivasi_Click()
  frmAktivasi.Action Me
End Sub

Private Sub cmdAdd_Click()
  nPos = Add
  GetEdit True
  GetFakturBrowse False
  cNomor.Text = CreateNomorFaktur(objData, sisModulTransaksi.StockOpname, "totstockopname", "nomorstockopname")
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

Private Function validOK() As Boolean
validOK = True

End Function

Private Sub cmdOK_Click()
Dim n As Integer

  If validOK() Then
   
    'jika baris <= Nomor
    If vaArray.UpperBound(1) + 2 <= nNomor.value Then
      vaArray.ReDim 0, nNomor.value - 1, 0, 6
      n = vaArray.UpperBound(1)
    'jika baris=0
    ElseIf vaArray.UpperBound(1) = -1 Then
      nNomor.value = 1
      vaArray.ReDim 0, nNomor.value - 1, 0, 6
      n = vaArray.UpperBound(1)
    Else
      n = nNomor.value - 1
    End If
        
    vaArray(n, 0) = nNomor.value
    vaArray(n, 1) = cBarcode.Text
    vaArray(n, 2) = cNama.Text
    vaArray(n, 3) = nHargaBeli.value 'kodeharga
    vaArray(n, 4) = nOnHand.value 'harga beli
    vaArray(n, 5) = nQty.value '3
    vaArray(n, 6) = cKode ' 5
    
    TDBGrid1.Array = vaArray
    TDBGrid1.ReBind
    
    InitValue1
    
    nNomor.value = vaArray.UpperBound(1) + 2
    cBarcode.SetFocus
  End If
End Sub

Private Sub InitValue1()
  cBarcode.Default
  cNama.Default
  nHargaBeli.Default
  nOnHand.Default
  nNomor.value = 1
  nQty.Default
End Sub

Private Sub cmdSimpan_Click()
Dim lSave As Boolean
Dim Faktur As String
Dim n As Integer

  lSave = True
  objData.Start GetDSN
  Faktur = cNomor.Text
  lSave = IIf(lSave, objData.Delete(GetDSN, "totstockopname", "nomorstockopname", sisAssign, Faktur), False)
  lSave = IIf(lSave, objData.Delete(GetDSN, "stockopname", "nomorstockopname", sisAssign, Faktur), False)
  lSave = IIf(lSave, objData.Delete(GetDSN, "kartustock", "nomor", sisAssign, Faktur), False)
  lSave = IIf(lSave, objData.Delete(GetDSN, "bukubesar", "faktur", sisAssign, Faktur), False)
  lSave = IIf(lSave, objData.Update(GetDSN, "totstockopname", "nomorstockopname = '" & Faktur & "'", Array("nomorstockopname", "tgl", "keterangan", "username", "datetime", "kodegudang"), Array(Faktur, Format(dTgl.value, "yyyy-MM-dd"), "", GetRegistry(reg_Username), SNow, Gudang)), False)
  
  For n = 0 To vaArray.UpperBound(1)
    lSave = IIf(lSave, objData.Add(GetDSN, "stockopname", Array("nomorstockopname", "kodestock", "adjust", "kodegudang"), Array(Faktur, vaArray(n, 6), vaArray(n, 5) - vaArray(n, 4), cGudang.Text)), False)
    lSave = IIf(lSave, UpdKartuStock(objData, SisKartuStock.PenyesuaianStock, Faktur, Format(dTgl.value, "yyyy-MM-dd"), vaArray(n, 6), vaArray(n, 5) - vaArray(n, 4), GetHargaPokok(objData, vaArray(n, 6)), 0, "Penyesuaian " & cKeterangan.Text, cGudang.Text, GetHargaPokok(objData, vaArray(n, 6))), False)
    
    If vaArray(n, 5) - vaArray(n, 4) < 0 Then
      'AutoJurnal :
      ' Biaya
      '  Persediaan
      lSave = IIf(lSave, UpdKodeTr(objData, msPenyesuaian, Faktur, Format(dTgl.value, "yyyy-MM-dd"), aCfg(objData, msRekeningPenyesuaianKurang), GetCostCenterUser(objData, GetRegistry(reg_Username)), "Penyesuaian " & vaArray(n, 2), Abs(vaArray(n, 5) - vaArray(n, 4)) * GetHargaPokok(objData, vaArray(n, 6)), 0, "", SNow, vaArray(n, 6)), False)
        lSave = IIf(lSave, UpdKodeTr(objData, msPenyesuaian, Faktur, Format(dTgl.value, "yyyy-MM-dd"), GetAkunInventory(objData, vaArray(n, 6)), GetCostCenterUser(objData, GetRegistry(reg_Username)), "Penyesuaian " & vaArray(n, 2), 0, Abs(vaArray(n, 5) - vaArray(n, 4)) * GetHargaPokok(objData, vaArray(n, 6)), "", SNow, vaArray(n, 6)), False)
    Else
      'AutoJurnal :
      ' Persediaan
      '  Pendapatan/modal
      lSave = IIf(lSave, UpdKodeTr(objData, msPenyesuaian, Faktur, Format(dTgl.value, "yyyy-MM-dd"), GetAkunInventory(objData, vaArray(n, 6)), GetCostCenterUser(objData, GetRegistry(reg_Username)), "Penyesuaian " & vaArray(n, 2), Abs(vaArray(n, 5) - vaArray(n, 4)) * GetHargaPokok(objData, vaArray(n, 6)), 0, "", SNow, vaArray(n, 6)), False)
        lSave = IIf(lSave, UpdKodeTr(objData, msPenyesuaian, Faktur, Format(dTgl.value, "yyyy-MM-dd"), aCfg(objData, msRekeningPenyesuian), GetCostCenterUser(objData, GetRegistry(reg_Username)), "Penyesuaian " & vaArray(n, 2), 0, Abs(vaArray(n, 5) - vaArray(n, 4)) * GetHargaPokok(objData, vaArray(n, 6)), "", SNow, vaArray(n, 6)), False)
    End If
  Next
    
  If lSave Then
    objData.Save GetDSN
  Else
    MsgBox "Maaf data tidak bisa disimpan", vbCritical
    objData.Cancel GetDSN
  End If
  GetEdit False
End Sub

Private Sub cNama_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "stock s", "s.nama,s.kodestock,s.Barcode,s.hargabeli,s.kodesatuan,s.hargajual,s.hargabeli,s.jenis", "s.nama", sisContent, cNama.Text, " and s.statusnonaktif <> 1")
  If Not dbData.EOF Then
    cNama.Text = cNama.Browse(dbData)
    cBarcode.Text = GetNull(dbData!barcode)
    nOnHand.value = GetSaldoStock(objData, cGudang.Text, GetNull(dbData!KodeStock), dTgl.value)
    nQty.value = nOnHand.value
    cKode = GetNull(dbData!KodeStock)
    nHargaBeli.value = GetNull(dbData!hargabeli)
  Else
    cNama.Default
  End If
End Sub

Private Sub cNomor_ButtonClick()
Dim objMenu As New CodeSuiteLibrary.Menu
Dim lSave As Boolean

  If aCfg(objData, msOtorisasiPenuh) = "Y" Then
    If objMenu.GetPassword("", Me, GetDSN) Then
      If objMenu.UserLevel <> 0 Then
          MsgBox "Maaf, Anda tidak diberikan wewenang untuk melakukan PENGEDITAN." & vbCrLf & _
                 "Hanya user dengan LEVEL 0 (SUPERVISOR) yg diijinkan", vbInformation, "OTORISASI not ALLOWED"
          Exit Sub
'      Else
'        MsgBox "OTORISASI DIBATALKAN", vbCritical
'        Exit Sub
      End If
    Else
      Exit Sub
    End If
  End If
  
  lSave = True

  Set dbData = objData.Browse(GetDSN, "totstockopname", "nomorstockopname,keterangan,username", "tgl", sisAssign, Format(dTgl.value, "yyyy-MM-dd"))
  If Not dbData.EOF Then
    cNomor.Text = cNomor.Browse(dbData)
    Me.Refresh
    objData.Start GetDSN
    If nPos = Delete Then
      'munculkan konten yg mau di hapus
      
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
  If Not IsInPeriod(dTgl.value) Or (dTgl.value > Date) Then
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
  InitValue1
  Me.Caption = "STOCK OPNAME GUDANG " & UCase(Gudang)
  GetEdit False
  TabIndex dTgl, n
  TabIndex cNomor, n
  TabIndex cGudang, n
  TabIndex cNamaGudang, n
  TabIndex cKeterangan, n
  TabIndex nNomor, n
  TabIndex cBarcode, n
  TabIndex cNama, n
  TabIndex nOnHand, n
  TabIndex nQty, n
  TabIndex cmdOK, n
  TabIndex cmdAdd, n
  TabIndex cmdEdit, n
  TabIndex cmdHapus, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  TabIndex cmdAktivasi, n
End Sub
Private Sub GetRows2()
Dim n As Integer
Dim cSQL As String
  
  cSQL = ""
  If Trim(Gudang) <> "" Then
    cSQL = " AND k.kodegudang = '" & Gudang & "'"
  End If
  vaArray.ReDim 0, -1, 0, 7
  Set dbData = objData.Browse(GetDSN, "stock s", "s.kodestock,s.barcode,s.nama,g.keterangan,s.kodesatuan,p.adjust,sum(k.debet-k.kredit) as saldostock", "s.kodestock", sisContent, TDBGrid1.Columns(0).FilterText, " AND barcode LIKE '%" & TDBGrid1.Columns(1).FilterText & "%' AND nama LIKE '%" & TDBGrid1.Columns(2).FilterText & "%' AND g.keterangan LIKE '%" & TDBGrid1.Columns(3).FilterText & "%' AND s.kodesatuan LIKE '%" & TDBGrid1.Columns(4).FilterText & "%'  AND p.nomorstockopname = '" & cNomor.Text & "' GROUP BY s.kodestock", "s.kodestock desc", Array("LEFT JOIN golongan g on g.kodegolongan = s.kodegolongan", "LEFT JOIN stockopname p on p.kodestock = s.kodestock", "LEFT JOIN kartustock k on k.kodestock = s.kodestock " & cSQL))
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!KodeStock)
      vaArray(n, 1) = GetNull(dbData!barcode)
      vaArray(n, 2) = GetNull(dbData!nama)
      vaArray(n, 3) = GetNull(dbData!keterangan)
      vaArray(n, 4) = GetNull(dbData!kodesatuan)
      vaArray(n, 6) = GetNull(dbData!adjust)
      vaArray(n, 7) = GetNull(dbData!saldostock)
      vaArray(n, 5) = vaArray(n, 7) - vaArray(n, 6)
      dbData.MoveNext
    Loop
  End If
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
  TDBGrid1.Refresh
End Sub

Private Sub initvalue()
  dTgl.value = Date
  cNomor.Default
  cGudang.Default
  cNamaGudang.Default
  cKeterangan.Default
  cGudang.Enabled = True
  vaArray.ReDim 0, -1, 0, 7
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
  TDBGrid1.Refresh
  InitValue1
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
    nSisaLebar = TDBGrid1.Width - TDBGrid1.Columns(0).Width - TDBGrid1.Columns(1).Width - TDBGrid1.Columns(3).Width - TDBGrid1.Columns(4).Width - TDBGrid1.Columns(5).Width - TDBGrid1.Columns(6).Width - TDBGrid1.Columns(7).Width
    TDBGrid1.Columns(2).Width = nSisaLebar - 1000
  End If
End Sub

Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
  TDBGrid1.Update
End Sub

Private Sub tdbgrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
  If Not IsNumeric(TDBGrid1.Columns(6).value) Then
    Cancel = True
    Exit Sub
  End If
  If Not IsNumeric(TDBGrid1.Columns(7).value) Then
    Cancel = True
    Exit Sub
  End If
  If ColIndex < 6 Then
    Cancel = True
    Exit Sub
  End If
  TDBGrid1.Columns(7).value = Val(TDBGrid1.Columns(5).value) + Val(TDBGrid1.Columns(6).value)
End Sub

Private Sub TDBGrid1_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal col As Integer, ByVal CellStyle As TrueOleDBGrid70.StyleDisp)
On Error Resume Next

  Select Case col
      Case 5
          Dim col1 As Long
          col1 = CLng(TDBGrid1.Columns(5).CellText(Bookmark))
          If col1 < 0 Then CellStyle.ForeColor = vbRed
      Case 6
          Dim col2 As Long
          col2 = CLng(TDBGrid1.Columns(6).CellText(Bookmark))
          If col2 < 0 Then CellStyle.ForeColor = vbRed
      Case 7
          Dim Col3 As Long
          Col3 = CLng(TDBGrid1.Columns(7).CellText(Bookmark))
          If Col3 < 0 Then CellStyle.ForeColor = vbRed
          
  End Select
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim n As Integer
Dim nQtyTmp As Integer

On Error Resume Next

  If KeyCode = vbKeyDelete Then
    If vaArray.UpperBound(1) >= 0 Then
      TDBGrid1.Delete
      TDBGrid1.Update
      For n = 0 To vaArray.UpperBound(1)
        vaArray(n, 0) = n + 1
      Next
      nNomor.value = vaArray.UpperBound(1) + 2
      TDBGrid1.ReBind
    End If
  End If
End Sub
