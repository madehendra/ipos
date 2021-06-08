VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trPelunasanCicilanBarang 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pelunasan Cicilan Barang..."
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   7830
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   585
      Left            =   0
      Top             =   6075
      Width           =   7830
      _ExtentX        =   13811
      _ExtentY        =   1032
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
         Picture         =   "trPelunasanCicilanBarang.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   3390
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
         Picture         =   "trPelunasanCicilanBarang.frx":028A
      End
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   1170
         TabIndex        =   2
         Top             =   75
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
         Picture         =   "trPelunasanCicilanBarang.frx":0429
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
         Picture         =   "trPelunasanCicilanBarang.frx":0555
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   6690
         TabIndex        =   4
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
         Picture         =   "trPelunasanCicilanBarang.frx":0700
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   5610
         TabIndex        =   5
         Top             =   75
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
         Picture         =   "trPelunasanCicilanBarang.frx":07A6
      End
   End
   Begin BiSAFramProject.BiSAFrame BisaFrame2 
      Height          =   4275
      Left            =   0
      Top             =   1830
      Width           =   7830
      _ExtentX        =   13811
      _ExtentY        =   7541
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
      Begin BiSANumberBoxProject.BiSANumberBox nTotalPelunasan 
         Height          =   330
         Left            =   5910
         TabIndex        =   12
         Top             =   3765
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   582
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
         Height          =   3660
         Left            =   75
         TabIndex        =   6
         Top             =   60
         Width           =   7665
         _ExtentX        =   13520
         _ExtentY        =   6456
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   4
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "No Cicilan"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Bulan"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Tahun"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Barang"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Jumlah"
         Columns(5).DataField=   ""
         Columns(5).NumberFormat=   "###,###,###,##0.00"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   6
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=6"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=503"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=423"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2170"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2090"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=1455"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1376"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=513"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=1773"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1693"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=513"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=4524"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=4445"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=516"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=1455"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=1376"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=197122"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         Appearance      =   0
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1.5
         FootLines       =   0
         Caption         =   "Lembar Cicilan"
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
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=50,.parent=13"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=54,.parent=13"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=2"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=2"
         _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
         _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
         _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
         _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=28,.parent=13"
         _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=14"
         _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=17"
         _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=66,.parent=13,.alignment=1"
         _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=63,.parent=14"
         _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=64,.parent=15,.alignment=1"
         _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=65,.parent=17"
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
      Begin BiSATextBoxProject.BiSABrowse cAkunKas 
         Height          =   330
         Left            =   75
         TabIndex        =   13
         Top             =   3780
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
         Enabled         =   0   'False
         Appearance      =   0
         Caption         =   "Akun Kas"
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
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1860
      Left            =   0
      Top             =   0
      Width           =   7830
      _ExtentX        =   13811
      _ExtentY        =   3281
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
         Left            =   3330
         TabIndex        =   7
         Top             =   765
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
         TabIndex        =   8
         Top             =   1080
         Width           =   5955
         _ExtentX        =   10504
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
         TabIndex        =   9
         Top             =   765
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
         Caption         =   "Anggota"
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
         Top             =   450
         Width           =   2910
         _ExtentX        =   5133
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
      Begin BiSATextBoxProject.BiSABrowse cFaktur 
         Height          =   330
         Left            =   75
         TabIndex        =   11
         Top             =   1395
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
      Begin VB.Label lbCostCenter 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   135
         TabIndex        =   14
         Top             =   150
         Width           =   6030
      End
   End
End
Attribute VB_Name = "trPelunasanCicilanBarang"
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

Private Sub GetFakturBrowse(ByVal lStat As Boolean)
  cFaktur.Button = lStat
End Sub

Private Sub cAkunKas_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", "kodeakun,keterangan", "jenis", sisAssign, "D", , "kodeakun")
  If Not dbData.EOF Then
    cAkunKas.Text = cAkunKas.Browse(dbData)
  End If
End Sub

Private Sub cFaktur_ButtonClick()
Dim db As New ADODB.Recordset
Dim n As Integer
Dim lSave As Boolean

  lSave = True
  vaArray.ReDim 0, -1, 0, 5
  Set dbData = objData.Browse(GetDSN, "totpelunasancicilan", "nomorpelunasancicilan,kodeakun,jumlah", "kodeanggota", sisAssign, cCustomer.Text)
  If Not dbData.EOF Then
    cFaktur.Text = cFaktur.Browse(dbData)
    cAkunKas.Text = GetNull(dbData!kodeakun)
    nTotalPelunasan.Value = GetNull(dbData!Jumlah)
    Set dbData = objData.Browse(GetDSN, "cicilan c", "c.lunas,c.nomorcicilan,c.nomorpelunasancicilan,c.bulan,c.tahun,c.jumlah,s.nama", "c.nomorpelunasancicilan", sisAssign, cFaktur.Text, , , Array("LEFT JOIN totcicilan t on t.nomorcicilan = c.nomorcicilan", "LEFT JOIN stock s on s.kodestock = t.kodestock"))
    If Not dbData.EOF Then
      Do While Not dbData.EOF
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        n = vaArray.UpperBound(1)
        vaArray(n, 0) = -1
        vaArray(n, 1) = GetNull(dbData!nomorcicilan)
        vaArray(n, 2) = GetNull(dbData!Bulan)
        vaArray(n, 3) = GetNull(dbData!Tahun)
        vaArray(n, 4) = GetNull(dbData!nama)
        vaArray(n, 5) = GetNull(dbData!Jumlah)
        dbData.MoveNext
      Loop
      Set tdbgrid1.Array = vaArray
      tdbgrid1.ReBind
    End If
    Me.Refresh
    If nPos = Delete Then
      If MsgBox("Apakah anda yakin data akan dihapus", vbYesNo) = vbYes Then
        objData.Start GetDSN
        
        lSave = IIf(lSave, DelKodeTr(objData, msPelunasanCicilanBarang, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "totpelunasancicilan", "nomorpelunasancicilan", sisAssign, cFaktur.Text), False)
        'Update status
        'kembalikan status menjadi 0 dan clear kan no pelunasan
        
        For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
          lSave = IIf(lSave, objData.Edit(GetDSN, "cicilan", " nomorcicilan = '" & vaArray(n, 1) & "' AND bulan = '" & vaArray(n, 2) & "' AND tahun = '" & vaArray(n, 3) & "'", Array("nomorpelunasancicilan", "lunas"), Array("", "0")), False)
        Next n

        If lSave Then
          objData.Save GetDSN
        Else
          objData.Cancel GetDSN
        End If
      End If
      InitValue
      vaArray.ReDim 0, -1, 0, 5
      Set tdbgrid1.Array = vaArray
      tdbgrid1.ReBind
      tdbgrid1.Refresh
      GetEdit False
    End If
  End If
  
  Set tdbgrid1.Array = vaArray
  tdbgrid1.ReBind
  tdbgrid1.Refresh
End Sub

Private Sub cmdAktivasi_Click()
  frmAktivasi.Action Me
End Sub

Private Sub cmdAdd_Click()
  nPos = Add
  GetEdit True
  GetFakturBrowse False
  cFaktur.Text = GetNomor("totpelunasancicilan", "nomorpelunasancicilan", GetID, SisModulTransaksi.PelunasanCicilanBarang)
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
    InitValue
  Else
    Unload Me
  End If
End Sub

Private Function ValidSaving() As Boolean
  ValidSaving = True
End Function

Private Sub ccustomer_ButtonClick()
Dim n As Integer

  vaArray.ReDim 0, -1, 0, 5
  Set dbData = objData.Browse(GetDSN, "anggota", "kodeanggota,nama,alamat", "kodeanggota", sisContent, cCustomer.Text, , "kodeanggota,nama")
  If Not dbData.EOF Then
    cCustomer.Text = cCustomer.Browse(dbData)
    cCustomer.Text = GetNull(dbData!kodeanggota)
    cNamaCustomer.Text = GetNull(dbData!nama, "")
    cAlamat.Text = GetNull(dbData!alamat, "")
    If nPos = Add Then
    Set dbData = objData.Browse(GetDSN, "cicilan c", "c.nomorcicilan,c.bulan,c.tahun,s.nama,c.jumlah", "t.kodeanggota", sisAssign, cCustomer.Text, " and c.nomorpelunasancicilan = ''", "c.tahun,c.bulan asc", Array("LEFT JOIN totcicilan t on t.nomorcicilan = c.nomorcicilan", "LEFT JOIN stock s on s.kodestock = t.kodestock"))
    If Not dbData.EOF Then
      Do While Not dbData.EOF
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        n = vaArray.UpperBound(1)
        vaArray(n, 0) = 0
        vaArray(n, 1) = GetNull(dbData!nomorcicilan)
        vaArray(n, 2) = GetNull(dbData!Bulan)
        vaArray(n, 3) = GetNull(dbData!Tahun)
        vaArray(n, 4) = GetNull(dbData!nama)
        vaArray(n, 5) = GetNull(dbData!Jumlah)
        dbData.MoveNext
      Loop
      Set tdbgrid1.Array = vaArray
      tdbgrid1.ReBind
      tdbgrid1.Refresh
    End If
    End If
  End If
End Sub

Private Sub cmdSimpan_Click()
Dim lSave As Boolean
Dim Faktur As String
Dim n As Integer
Dim i As Integer
lSave = True
  
  If isValidSaving Then
    objData.Start GetDSN
    'simpan di tabel totpelunasancicilan
    lSave = IIf(lSave, objData.Delete(GetDSN, "totpelunasancicilan", "nomorpelunasancicilan", sisAssign, cFaktur.Text), False)
    lSave = IIf(lSave, objData.Add(GetDSN, "totpelunasancicilan", Array("nomorpelunasancicilan", "username", "kodeanggota", "kodeakun", "tgl", "jumlah"), Array(cFaktur.Text, GetRegistry(reg_UserName), cCustomer.Text, cAkunKas.Text, Format(dTgl.Value, "yyyy-MM-dd"), nTotalPelunasan.Value)), False)
    'update tabel cicilan
    For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
      If vaArray(n, 0) = -1 Then
        lSave = IIf(lSave, objData.Edit(GetDSN, "cicilan", " nomorcicilan = '" & vaArray(n, 1) & "' AND bulan = '" & vaArray(n, 2) & "' AND tahun = '" & vaArray(n, 3) & "'", Array("nomorpelunasancicilan", "lunas"), Array(cFaktur.Text, "1")), False)
      Else
        lSave = IIf(lSave, objData.Edit(GetDSN, "cicilan", " nomorcicilan = '" & vaArray(n, 1) & "' AND bulan = '" & vaArray(n, 2) & "' AND tahun = '" & vaArray(n, 3) & "'", Array("nomorpelunasancicilan", "lunas"), Array("", "0")), False)
      End If
    Next n
    
    'Update bukubesar
    'kas
    ' Piutang Customer
    lSave = IIf(lSave, DelKodeTr(objData, msPelunasanCicilanBarang, cFaktur.Text), False)
    lSave = IIf(lSave, UpdKodeTr(objData, msPelunasanCicilanBarang, cFaktur.Text, Format(dTgl.Value, "yyyy-MM-dd"), cAkunKas.Text, aCfg(objData, msCostCenterJualBeli), "Pelunasan Cicilan", nTotalPelunasan.Value, 0, "K", SNow), False)
      lSave = IIf(lSave, UpdKodeTr(objData, msPelunasanCicilanBarang, cFaktur.Text, Format(dTgl.Value, "yyyy-MM-dd"), GetAkunMember(objData, cCustomer.Text), aCfg(objData, msCostCenterJualBeli), "Pelunasan Cicilan", 0, nTotalPelunasan.Value, "K", SNow), False)
    
    If lSave Then
      objData.Save GetDSN
    Else
      objData.Cancel GetDSN
    End If
    InitValue
    GetEdit False
  End If
End Sub

Private Function isValidSaving() As Boolean
Dim dba As New ADODB.Recordset
isValidSaving = True

End Function

Private Sub cNamaCustomer_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "anggota", "nama,kodeanggota,alamat", "nama", sisContent, cNamaCustomer.Text, , "nama,kodeanggota")
  If Not dbData.EOF Then
    cNamaCustomer.Text = cNamaCustomer.Browse(dbData)
    cCustomer.Text = GetNull(dbData!kodeanggota)
    cNamaCustomer.Text = GetNull(dbData!nama, "")
    cAlamat.Text = GetNull(dbData!alamat, "")
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  InitValue
  GetEdit False
  
  Set dbData = objData.Browse(GetDSN, "costcenter", , "kodecostcenter", sisAssign, aCfg(objData, msCostCenterJualBeli))
  If Not dbData.EOF Then
    lbCostCenter.Caption = "Cost Centre : " & GetNull(dbData!Keterangan)
  End If
  
  TabIndex dTgl, n
  TabIndex cCustomer, n
  TabIndex cNamaCustomer, n
  TabIndex cAlamat, n
  TabIndex cFaktur, n
  TabIndex cmdAdd, n
  TabIndex cmdEdit, n
  TabIndex cmdHapus, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  TabIndex cmdAktivasi, n
End Sub

Private Sub InitValue()
Dim n As Integer

  cCustomer.Default
  cNamaCustomer.Default
  cAlamat.Default
  dTgl.Value = Date
  cFaktur.Default
  cAkunKas.Text = cKasTeller
  nTotalPelunasan.Default
  vaArray.ReDim 0, -1, 0, 5
  Set tdbgrid1.Array = vaArray
  tdbgrid1.ReBind
End Sub

Private Sub GetEdit(lPar As Boolean)
  BiSAFrame1.Enabled = lPar
  BiSAFrame2.Enabled = lPar
  lEdit = lPar
  InitValue
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

Private Sub tdbgrid1_AfterColUpdate(ByVal ColIndex As Integer)
Dim n As Integer
Dim nTmpTotal As Double

  tdbgrid1.Update
  nTmpTotal = 0
  For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
    If vaArray(n, 0) = -1 Then
      nTmpTotal = nTmpTotal + vaArray(n, 5)
    End If
  Next n
  nTotalPelunasan.Value = nTmpTotal
End Sub
