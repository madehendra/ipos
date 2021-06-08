VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form mstPeriodeAkuntansi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setup Periode Akuntansi"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   8805
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1920
      Left            =   15
      Top             =   0
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   3387
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   4
      BackColor       =   -2147483633
      Begin VB.OptionButton optAktif 
         Caption         =   "&1 Non Aktif"
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
         Index           =   1
         Left            =   2445
         TabIndex        =   12
         Top             =   1560
         Width           =   1095
      End
      Begin VB.OptionButton optAktif 
         Caption         =   "&0 Aktif"
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
         Index           =   0
         Left            =   1590
         TabIndex        =   11
         Top             =   1560
         Width           =   825
      End
      Begin BiSADateProject.BiSADate dAwal 
         Height          =   330
         Left            =   165
         TabIndex        =   0
         Top             =   825
         Width           =   2775
         _ExtentX        =   4895
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
         Caption         =   "AWAL"
         CaptionWidth    =   1300
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
         Left            =   165
         TabIndex        =   1
         Top             =   90
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   582
         Text            =   "1234"
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
         MaxLength       =   4
         Appearance      =   0
         Caption         =   "KODE"
         CaptionWidth    =   1300
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
      Begin BiSATextBoxProject.BiSATextBox cKeterangan 
         Height          =   330
         Left            =   165
         TabIndex        =   2
         Top             =   465
         Width           =   6240
         _ExtentX        =   11007
         _ExtentY        =   582
         Text            =   "1234567890123456789012345678901234567890"
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
         MaxLength       =   40
         Appearance      =   0
         Caption         =   "KETERANGAN"
         CaptionWidth    =   1300
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
      Begin BiSADateProject.BiSADate dAkhir 
         Height          =   330
         Left            =   165
         TabIndex        =   3
         Top             =   1185
         Width           =   2775
         _ExtentX        =   4895
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
         Caption         =   "AKHIR"
         CaptionWidth    =   1300
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
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   3840
      Left            =   0
      Top             =   1890
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   6773
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   4
      BackColor       =   -2147483633
      Begin TrueOleDBGrid70.TDBGrid DataGrid1 
         Height          =   3720
         Left            =   105
         TabIndex        =   4
         Top             =   75
         Width           =   8610
         _ExtentX        =   15187
         _ExtentY        =   6562
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "KODE"
         Columns(0).DataField=   "Kode"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "AWAL"
         Columns(1).DataField=   "Awal"
         Columns(1).NumberFormat=   "dd/MM/yyyy"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "AKHIR"
         Columns(2).DataField=   "Akhir"
         Columns(2).NumberFormat=   "dd/MM/yyyy"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "STATUS"
         Columns(3).DataField=   "Status"
         Columns(3).NumberFormat=   "FormatText Event"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "KETERANGAN"
         Columns(4).DataField=   "Keterangan"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   5
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=5"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1455"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1376"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=512"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2275"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2196"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2117"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2037"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=2064"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1984"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=513"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=6720"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=6641"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=516"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         Appearance      =   2
         DefColWidth     =   0
         HeadLines       =   1.5
         FootLines       =   1
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFCFCED&,.bold=0,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bgcolor=&HEBDACB&"
         _StyleDefs(11)  =   ":id=2,.fgcolor=&H8000000D&,.bold=0,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
         _StyleDefs(13)  =   ":id=2,.fontname=MS Sans Serif"
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
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=0"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=54,.parent=13"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=2"
         _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
         _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
         _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
         _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=32,.parent=13"
         _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14"
         _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
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
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   0
      Top             =   5715
      Width           =   8775
      _ExtentX        =   15478
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
      BorderStyle     =   4
      BackColor       =   -2147483633
      Begin BiSAButtonProject.BiSAButton cmdHapus 
         Height          =   435
         Left            =   2205
         TabIndex        =   5
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
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
         Picture         =   "mstPeriodeAkuntansi.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   3375
         TabIndex        =   6
         Top             =   105
         Width           =   405
         _ExtentX        =   714
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
         Picture         =   "mstPeriodeAkuntansi.frx":028A
      End
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   1155
         TabIndex        =   7
         Top             =   105
         Width           =   1020
         _ExtentX        =   1799
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
         Picture         =   "mstPeriodeAkuntansi.frx":0429
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   435
         Left            =   90
         TabIndex        =   8
         Top             =   105
         Width           =   1035
         _ExtentX        =   1826
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
         Picture         =   "mstPeriodeAkuntansi.frx":0555
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   7605
         TabIndex        =   9
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
         Picture         =   "mstPeriodeAkuntansi.frx":0700
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   6525
         TabIndex        =   10
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
         Picture         =   "mstPeriodeAkuntansi.frx":07A6
      End
   End
End
Attribute VB_Name = "mstPeriodeAkuntansi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lClick As Boolean
Dim nPos As SisPos
Dim lEdit As Boolean
Dim dbData As New ADODB.Recordset
Dim db As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data

Private Sub GetEdit(lPar As Boolean)
  lEdit = lPar
  BiSAFrame1.Enabled = lPar
  SetButton cmdSimpan, cmdKeluar, cmdAdd, cmdEdit, cmdHapus, nPos, lPar
End Sub

Private Sub cKode_Validate(Cancel As Boolean)
  With dbData
    If Trim(cKode.Text) = "" Then
      cKode.Text = 1
      If .RecordCount > 0 Then
        .MoveLast
        cKode.Text = Val(!Kode) + 1
      End If
    End If
    cKode.Text = Padl(Trim(cKode.Text), cKode.MaxLength, "0")
    
    If Not .BOF Then .MoveFirst
      .Find "Kode='" & GetInduk(cKode.Text, False) & "'"
      If Not .EOF Then
        If nPos = Add Then
          MsgBox "Data tersebut sudah ada...!", vbInformation
          Cancel = True
          initvalue
          cKode.SetFocus
          Exit Sub
'        Else
'          GetMemory
'          If !status = "1" Then
'            MsgBox "Periode Sudah Ditutup, Tidak Bisa di Edit / Dihapus .. !", vbExclamation
'            Cancel = True
'            cKode.SetFocus
'            Exit Sub
'          End If
'          If nPos = Delete Then HapusData
        End If
      Else
        If nPos <> Add Then
          MsgBox "Data tersebut belum ada...!", vbInformation
          Cancel = True
          initvalue
          cKode.SetFocus
        End If
      End If
  End With
End Sub

Private Sub cmdAktivasi_Click()
  frmAktivasi.Action Me
End Sub

Private Sub cmdAdd_Click()
  nPos = Add
  GetEdit True
  cKode.Enabled = True
  cKode.SetFocus
  initvalue
  
  Set db = objData.Browse(GetDSN, "Periode", "Max(Akhir) as Akhir")
  If db.RecordCount > 0 Then
    dAwal.Value = DateAdd("d", 1, GetNull(db!akhir, Date - 1))
    dAkhir.Value = EOM(dAwal.Value)
  End If
End Sub

Private Sub initvalue()
  cKode.Default
  cKeterangan.Default
  dAwal.Value = Date
  dAkhir.Value = Date
  lClick = False
End Sub

Private Sub HapusData()
  If Trim(cKode.Text) = "" Then
    MsgBox "Kode tidak boleh kosong"
    cKode.SetFocus
    GetEdit False
  End If
  
  If MsgBox("Data benar-benar dihapus...?", vbQuestion + vbYesNo) = vbYes Then
    objData.Delete GetDSN, "Periode", "Kode", sisAssign, cKode.Text
  End If
  initvalue
  GetSQL
  GetEdit False
End Sub

Private Sub GetMemory()
  cKode.Text = (dbData!Kode)
  cKeterangan.Text = Trim(dbData!keterangan)
  dAwal.Value = Format(dbData!Awal, "dd/MM/yyyy")
  dAkhir.Value = Format(dbData!akhir, "dd/MM/yyyy")
  SetOpt optAktif, GetNull(dbData!status)
End Sub

Private Sub cmdEdit_Click()
  nPos = Edit
  GetEdit True
  cKode.SetFocus
End Sub

Private Sub cmdHapus_Click()
  nPos = Delete
  GetEdit True
  cKode.SetFocus
End Sub

Private Sub cmdKeluar_Click()
  If lEdit Then
    GetEdit False
    initvalue
  Else
    Unload Me
  End If
End Sub

Private Sub cmdSimpan_Click()
  Dim vaField, vaValue
  If ValidSaving() Then
    vaField = Array("kode", "keterangan", "awal", "akhir", "status")
    vaValue = Array(cKode.Text, cKeterangan.Text, dAwal.Value, dAkhir.Value, GetOpt(optAktif))
    objData.Update GetDSN, "Periode", "Kode='" & cKode.Text & "'", vaField, vaValue
    GetSQL
    initvalue
    GetEdit False
   Else
    cKode.SetFocus
    Exit Sub
  End If
End Sub

Private Function ValidSaving() As Boolean
  ValidSaving = True
  If Not CheckData(cKode.Text, "Kode harus diisi...?") Then
    ValidSaving = False
    cKode.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cKeterangan.Text, "Keterangan harus diisi...?") Then
    ValidSaving = False
    cKeterangan.SetFocus
    Exit Function
  End If
End Function

Private Sub DataGrid1_Click()
  lClick = True
End Sub

Private Sub DataGrid1_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
  If Value = 0 Then
    Value = "Aktif"
  Else
    Value = "Tutup"
  End If
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 38 Or KeyCode = 40 Then
    lClick = True
  End If
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  If lClick Then
    GetMemory
  End If
  lClick = False
End Sub

Private Sub Form_Load()
Dim n As Single
  
  CenterForm Me
  initvalue
  SetIcon Me.hWnd
  
  GetSQL
  GetEdit False
  
  TabIndex cKode, n
  TabIndex cKeterangan, n
  TabIndex dAwal, n
  TabIndex dAkhir, n
  TabIndex optAktif(0), n
  TabIndex optAktif(1), n
  
  TabIndex cmdAdd, n
  TabIndex cmdEdit, n
  TabIndex cmdHapus, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  
End Sub

Private Sub GetSQL()
  Set dbData = objData.Browse(GetDSN, "Periode", , , , , , "Kode")
  Set DataGrid1.DataSource = dbData
  DataGrid1.Refresh
End Sub


