VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trPenukaranHadiah 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FORM PENUKARAN HADIAH"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   10215
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1395
      Left            =   15
      Top             =   45
      Width           =   10170
      _ExtentX        =   17939
      _ExtentY        =   2461
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
      Begin BiSADateProject.BiSADate dTanggal 
         Height          =   330
         Left            =   120
         TabIndex        =   0
         Top             =   150
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   582
         Value           =   "19-11-2003"
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
      Begin BiSATextBoxProject.BiSABrowse cNama 
         Height          =   330
         Left            =   6270
         TabIndex        =   1
         Top             =   150
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   582
         Text            =   "123456"
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
         MaxLength       =   50
         Appearance      =   0
         Button          =   -1  'True
         CaptionWidth    =   1500
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
      Begin BiSATextBoxProject.BiSABrowse cAlamat 
         Height          =   330
         Left            =   120
         TabIndex        =   2
         Top             =   525
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   582
         Text            =   "123456"
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
         MaxLength       =   50
         Appearance      =   0
         Caption         =   "Alamat"
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
      Begin BiSATextBoxProject.BiSABrowse cFaktur 
         Height          =   330
         Left            =   120
         TabIndex        =   3
         Top             =   915
         Width           =   3225
         _ExtentX        =   5689
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
         Caption         =   "Faktur"
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
      Begin BiSATextBoxProject.BiSABrowse cCustomer 
         Height          =   330
         Left            =   3015
         TabIndex        =   21
         Top             =   150
         Width           =   3225
         _ExtentX        =   5689
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
      Begin BiSATextBoxProject.BiSABrowse cNamaDepartmen 
         Height          =   330
         Left            =   4470
         TabIndex        =   22
         Top             =   525
         Width           =   3240
         _ExtentX        =   5715
         _ExtentY        =   582
         Text            =   "123456"
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
         MaxLength       =   50
         Appearance      =   0
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
      Height          =   4815
      Left            =   15
      Top             =   1470
      Width           =   10170
      _ExtentX        =   17939
      _ExtentY        =   8493
      Caption         =   "AKUMULASI POIN HADIAH"
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
      Begin BiSAButtonProject.BiSAButton BiSAButton2 
         Height          =   435
         Left            =   1080
         TabIndex        =   19
         Top             =   3975
         Visible         =   0   'False
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   767
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
      Begin BiSAButtonProject.BiSAButton BiSAButton1 
         Height          =   450
         Left            =   165
         TabIndex        =   18
         Top             =   3960
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   794
         Caption         =   "PROSES"
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   2850
         Left            =   120
         TabIndex        =   4
         Top             =   1020
         Width           =   9900
         _ExtentX        =   17463
         _ExtentY        =   5027
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   4
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "FAKTUR"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "TGL"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "TGL EX"
         Columns(3).DataField=   ""
         Columns(3).NumberFormat=   "dd-MM-yyyy"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "POIN"
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "###,###,###,###,##0.00"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "TUKAR"
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   6
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=6"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=609"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=529"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=197124"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=4048"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3969"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=197124"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=197124"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=4207"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=4128"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=197121"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2461"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2381"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=197122"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=197124"
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
         ColumnFooters   =   -1  'True
         DataMode        =   4
         DefColWidth     =   0
         Enabled         =   0   'False
         HeadLines       =   1
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.alignment=1,.bold=0,.fontsize=825"
         _StyleDefs(15)  =   ":id=3,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=70,.parent=13,.alignment=2"
         _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=67,.parent=14"
         _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=68,.parent=15"
         _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=69,.parent=17"
         _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=58,.parent=13,.alignment=1"
         _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
         _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
         _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=28,.parent=13"
         _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=14"
         _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=15"
         _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=17"
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
         _StyleDefs(71)  =   ":id=37,.parent=34,.alignment=2,.bold=0,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(72)  =   ":id=37,.strikethrough=0,.charset=0"
         _StyleDefs(73)  =   ":id=37,.fontname=Tahoma"
         _StyleDefs(74)  =   "Named:id=38:HighlightRow"
         _StyleDefs(75)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(76)  =   "Named:id=39:EvenRow"
         _StyleDefs(77)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(78)  =   "Named:id=40:OddRow"
         _StyleDefs(79)  =   ":id=40,.parent=33"
         _StyleDefs(80)  =   "Named:id=41:RecordSelector"
         _StyleDefs(81)  =   ":id=41,.parent=34"
         _StyleDefs(82)  =   "Named:id=42:FilterBar"
         _StyleDefs(83)  =   ":id=42,.parent=33"
      End
      Begin BiSATextBoxProject.BiSABrowse cKodeHadiah 
         Height          =   330
         Left            =   120
         TabIndex        =   11
         Top             =   255
         Width           =   3225
         _ExtentX        =   5689
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
         Caption         =   "Hadiah"
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
      Begin BiSATextBoxProject.BiSABrowse cNamaHadiah 
         Height          =   330
         Left            =   3390
         TabIndex        =   12
         Top             =   255
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   582
         Text            =   "123456"
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
         MaxLength       =   50
         Appearance      =   0
         Button          =   -1  'True
         CaptionWidth    =   1500
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
      Begin BiSANumberBoxProject.BiSANumberBox nPoinHadiah 
         Height          =   330
         Left            =   6765
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   255
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   582
         BorderStyle     =   0
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
         BackColor       =   -2147483633
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
      Begin BiSANumberBoxProject.BiSANumberBox nSisaPoin 
         Height          =   330
         Left            =   7185
         TabIndex        =   14
         Top             =   4350
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   582
         BorderStyle     =   0
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
         Caption         =   "Sisa Poin"
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
      Begin BiSANumberBoxProject.BiSANumberBox nPoinDitukar 
         Height          =   330
         Left            =   8295
         TabIndex        =   15
         Top             =   3945
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   582
         BorderStyle     =   0
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
         CaptionWidth    =   2000
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
         Left            =   1545
         TabIndex        =   16
         Top             =   630
         Width           =   1185
         _ExtentX        =   2090
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
      Begin BiSANumberBoxProject.BiSANumberBox nJumlah 
         Height          =   330
         Left            =   2760
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   630
         Width           =   1185
         _ExtentX        =   2090
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
      Begin VB.Label Label1 
         Caption         =   "POIN"
         Height          =   300
         Left            =   7815
         TabIndex        =   20
         Top             =   330
         Width           =   495
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame4 
      Height          =   630
      Left            =   15
      Top             =   6300
      Width           =   10170
      _ExtentX        =   17939
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
         Left            =   2700
         TabIndex        =   5
         Top             =   120
         Visible         =   0   'False
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
         Picture         =   "trPenukaranHadiah.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   90
         TabIndex        =   6
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
         Picture         =   "trPenukaranHadiah.frx":028A
      End
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   1635
         TabIndex        =   7
         Top             =   120
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
         Picture         =   "trPenukaranHadiah.frx":0429
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   435
         Left            =   555
         TabIndex        =   8
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
         Picture         =   "trPenukaranHadiah.frx":0555
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   9015
         TabIndex        =   9
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
         Picture         =   "trPenukaranHadiah.frx":0700
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   7935
         TabIndex        =   10
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
         Picture         =   "trPenukaranHadiah.frx":07A6
      End
   End
End
Attribute VB_Name = "trPenukaranHadiah"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nPos As SisPos
Dim lClick As Boolean
Dim lStart As Boolean
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim objMenu As New CodeSuiteLibrary.Menu
Dim vaArray As New XArrayDB
Dim lEdit As Boolean

Private Sub BiSAButton1_Click()
Dim n As Integer
Dim nHadiah As Integer
Dim nPoinReg As Integer
Dim nPoinSisa As Integer

  GetData
  
  nHadiah = nJumlah.Value
  For n = 0 To vaArray.UpperBound(1)
      If nHadiah - vaArray(n, 4) >= 0 Then
        nHadiah = nHadiah - vaArray(n, 4)
        vaArray(n, 5) = vaArray(n, 4)
        vaArray(n, 4) = 0
      Else
        If vaArray(n, 4) > nHadiah Then
          vaArray(n, 4) = vaArray(n, 4) - nHadiah
          vaArray(n, 5) = nHadiah
          nHadiah = 0
        End If
      End If
    
  Next n
  
  For n = 0 To vaArray.UpperBound(1)
    nPoinReg = nPoinReg + vaArray(n, 5)
    nPoinSisa = nPoinSisa + vaArray(n, 4)
  Next n

  
  If nHadiah > 0 Then
    MsgBox "Maaf Poin Tidak Mencukupi"
    GetData
  End If
  
  nPoinDitukar.Value = nPoinReg
  nSisaPoin.Value = nPoinSisa

  TDBGrid1.ReBind
  TDBGrid1.Refresh
End Sub

Private Sub BiSAButton2_Click()
Dim cSQL As String
Dim dba As New ADODB.Recordset
Dim vaField
Dim vaValue
Dim lSave As Boolean

  lSave = True
    
  cSQL = cSQL & " select DISTINCT(a.kodeanggota) as kodeanggota,a.nama from pelunasanpiutang p"
  cSQL = cSQL & " LEFT JOIN totpelunasanpiutang t  on t.nomorpelunasanpiutang = p.nomorpelunasanpiutang"
  cSQL = cSQL & " LEFT JOIN anggota a on a.kodeanggota = t.kodeanggota"
  cSQL = cSQL & " where t.tgl >= '2015-10-1'"
  cSQL = cSQL & " ORDER BY a.kodeanggota"
  
  vaField = Array("faktur", "tgl", "kodeanggota", "poinhadiah", "exdate", "status")

  Set dba = objData.Sql(GetDSN, cSQL)
  If Not dba.EOF Then
    Do While Not dba.EOF
      'cari satu satu penjualan yg dilunasi hitung dan jumlahkan poin nya
      vaValue = Array("OP" & GetNull(dba!kodeanggota), Date, GetNull(dba!kodeanggota), GetProsesPoinHadiahPending(GetNull(dba!kodeanggota)), DateAdd("M", 2, Date), "1")
      lSave = IIf(lSave, objData.Delete(GetDSN, "poinhadiah", "faktur", sisAssign, "OP" & GetNull(dba!kodeanggota)), False)
      lSave = IIf(lSave, objData.Add(GetDSN, "poinhadiah", vaField, vaValue), False)
      dba.MoveNext
    Loop
  End If
  
  If lSave Then
    objData.Save GetDSN
  Else
    objData.Cancel GetDSN
  End If
  
  MsgBox "PROSES UPDATE SELESAI"
End Sub

Private Function GetProsesPoinHadiahPending(ByVal cClient As String) As Double
Dim cSQL As String
Dim dbPoin As ADODB.Recordset
Dim nJumlahPoinPending As Double

  nJumlahPoinPending = 0
  'fungsi ini akan mengakumulasi jumlah poin dari member
  cSQL = cSQL & " select p.nomorpenjualan,t.tgl,a.kodeanggota,a.nama,a.telp from pelunasanpiutang p"
  cSQL = cSQL & " LEFT JOIN totpelunasanpiutang t  on t.nomorpelunasanpiutang = p.nomorpelunasanpiutang"
  cSQL = cSQL & " LEFT JOIN anggota a on a.kodeanggota = t.kodeanggota"
  cSQL = cSQL & " where t.tgl >= '2015-10-1' and a.kodeanggota = '" & cClient & "'"
  cSQL = cSQL & " ORDER BY a.kodeanggota"

  Set dbPoin = objData.Sql(GetDSN, cSQL)
  If Not dbPoin.EOF Then
    Do While Not dbPoin.EOF
      nJumlahPoinPending = nJumlahPoinPending + GetPoinDong(GetNull(dbPoin!nomorpenjualan))
      dbPoin.MoveNext
    Loop
  End If
  GetProsesPoinHadiahPending = nJumlahPoinPending
End Function

Private Function GetPoinDong(ByVal cNomorFakturJual As String) As Double
Dim cSQL As String
Dim dbP As ADODB.Recordset

  cSQL = cSQL & " select sum(qty*harga) as total from penjualan"
  cSQL = cSQL & " where nomorpenjualan = '" & cNomorFakturJual & "' and discount >= 20"
  cSQL = cSQL & " and tgl >='" & Format(DateAdd("D", -7, "2015-10-1"), "yyyy-MM-dd") & "'"
  'function ini akan mengecek, faktur penjualan ini dapat berapa poin sih.
  Set dbP = objData.Sql(GetDSN, cSQL)
  If Not dbP.EOF Then
    GetPoinDong = GetNull(dbP!Total) \ 100000
  End If

End Function

Private Sub cCustomer_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "anggota a", "a.kodeanggota,a.nama,a.alamat,d.keterangan as namadep", "a.kodeanggota", sisContent, cCustomer.Text, , , Array("left join dep d on d.kodedep = a.kodedep"))
  If Not dbData.EOF Then
    cCustomer.Text = cCustomer.Browse(dbData)
    cCustomer.Text = GetNull(dbData!kodeanggota)
    cNama.Text = GetNull(dbData!nama)
    cNamaDepartmen.Text = GetNull(dbData!namadep)
    cAlamat.Text = GetNull(dbData!alamat)
    If nPos = Add Then
      GetData
    End If
  End If
End Sub

Private Sub cKodeHadiah_ButtonClick()
 Set dbData = objData.Browse(GetDSN, "msthadiah", "kodehadiah,keterangan,poin", "kodehadiah", sisContent, cKodeHadiah.Text)
  If Not dbData.EOF Then
    cKodeHadiah.Text = cKodeHadiah.Browse(dbData)
    cKodeHadiah.Text = GetNull(dbData!kodehadiah)
    cNamaHadiah.Text = GetNull(dbData!keterangan)
    nPoinHadiah.Value = GetNull(dbData!poin)
    GetJumlahKanPoin
  End If
End Sub

Private Sub cmdAdd_Click()
  nPos = Add
  GetEdit True
  initvalue
  dTanggal.SetFocus
  GetBrowseFaktur False
  cFaktur.Text = CreateNomorFaktur(objData, sisModulTransaksi.PenukaranPoin, "totpoin", "nomortukarpoin")
  'cFaktur.Text = GetNomor("totpoin", "nomortukarpoin", GetID, sisModulTransaksi.PenukaranPoin)
End Sub

Private Sub cmdAktivasi_Click()
  frmAktivasi.Action Me
End Sub

Private Sub cmdKeluar_Click()
  If lEdit Then
    initvalue
    GetEdit False
  Else
    Unload Me
  End If
End Sub

Private Sub cmdSimpan_Click()
Dim cSQL As String
Dim db As New ADODB.Recordset
Dim n As Single
Dim vaField
Dim vaValue
Dim lSave As Boolean
  
  If nJumlah.Value > 0 Then
    If nPoinDitukar.Value >= nJumlah.Value Then
      lSave = True
      vaField = Array("tukar", "poinhadiah", "tukardate", "status")
      
      For n = 0 To vaArray.UpperBound(1)
        If vaArray(n, 5) > 0 Then
          vaValue = Array(vaArray(n, 5) + getSisaPoinFaktur(vaArray(n, 1)), vaArray(n, 4), Format(dTanggal.Value, "yyyy-MM-dd"), "0")
          lSave = IIf(lSave, objData.Update(GetDSN, "poinhadiah", "faktur='" & vaArray(n, 1) & "'", vaField, vaValue), False)
        End If
      Next n
      
      'simpan di tabel totpoin dan poinhadiah
      vaField = Array("nomortukarpoin", "tgl", "kodeanggota", "kodehadiah", "qty", "poin", "keterangan")
      vaValue = Array(cFaktur.Text, Format(dTanggal.Value, "yyyy-MM-dd"), cCustomer.Text, cKodeHadiah.Text, nQty.Value, nPoinHadiah.Value, "")
      lSave = IIf(lSave, objData.Add(GetDSN, "totpoin", vaField, vaValue), False)
      
      lSave = IIf(lSave, objData.Delete(GetDSN, "tukarpoin", "nomortukarpoin", sisAssign, cFaktur.Text), False)
      vaField = Array("nomortukarpoin", "faktur", "poin", "tgl")
      
      For n = 0 To vaArray.UpperBound(1)
        If vaArray(n, 5) > 0 Then
          'simpan di tabel poinhadiah
          vaValue = Array(cFaktur.Text, vaArray(n, 1), vaArray(n, 5), Format(dTanggal.Value, "yyyy-MM-dd"))
          lSave = IIf(lSave, objData.Add(GetDSN, "tukarpoin", vaField, vaValue), False)
        End If
      Next n
      
      
      If lSave Then
        objData.Save GetDSN
        MsgBox "Data sudah berhasil disimpan"
        initvalue
      Else
        objData.Cancel GetDSN
      End If
    Else
      MsgBox "Maaf Poin Tidak Mencukupi, Data tidak bisa disimpan"
      GetData
    End If
  Else
    MsgBox "Pilih hadiah terlebih dahulu, data tidak bisa disimpan"
  End If
  initvalue
  GetEdit False
End Sub

Private Function getSisaPoinFaktur(ByVal cFak As String) As Double
Dim db As New ADODB.Recordset

  getSisaPoinFaktur = 0
  Set db = objData.Sql(GetDSN, "select tukar from poinhadiah where faktur = '" & cFak & "'")
  If Not db.EOF Then
    getSisaPoinFaktur = GetNull(db!tukar)
  End If
End Function

Private Sub cNama_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "anggota a", "a.kodeanggota,a.nama,a.alamat,d.keterangan as namadep", "a.nama", sisContent, cNama.Text, , , Array("left join dep d on d.kodedep = a.kodedep"))
  If Not dbData.EOF Then
    cNama.Text = cNama.Browse(dbData)
    cCustomer.Text = GetNull(dbData!kodeanggota)
    cNama.Text = GetNull(dbData!nama)
    cNamaDepartmen.Text = GetNull(dbData!namadep)
    cAlamat.Text = GetNull(dbData!alamat)
    If nPos = Add Then
      GetData
    End If
  End If
End Sub

Private Sub GetData()
Dim db As New ADODB.Recordset
Dim cSQL As String
Dim n As Integer
Dim nPoin As Integer
  
  vaArray.ReDim 0, -1, 0, 6
  nPoin = 0
  cSQL = "SELECT faktur,tgl,kodeanggota,poinhadiah,exdate,status FROM poinhadiah WHERE kodeanggota = '" & cCustomer.Text & "'" & _
  " AND exdate >='" & Format(dTanggal.Value, "yyyy-MM-dd") & "'" & _
  " AND poinhadiah > 0"
  
  'status 1 artinya poin masih valid/belum ditukar
  
  Set db = objData.Sql(GetDSN, cSQL)
  If Not db.EOF Then
    Do While Not db.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = 0
      vaArray(n, 1) = GetNull(db!Faktur)
      vaArray(n, 2) = GetNull(db!Tgl)
      vaArray(n, 3) = GetNull(db!exdate)
      vaArray(n, 4) = GetNull(db!poinhadiah)
      vaArray(n, 6) = vaArray(n, 4)
      nPoin = nPoin + vaArray(n, 4)
      db.MoveNext
    Loop
  End If
  
  Set TDBGrid1.Array = vaArray
  TDBGrid1.Columns(4).FooterText = Format(nPoin, "###,###,###,##")
  TDBGrid1.ReBind
  TDBGrid1.Refresh
End Sub

Private Sub GetBrowseFaktur(ByVal lStat As Boolean)
  cFaktur.Button = lStat
  cFaktur.Enabled = lStat
End Sub

Private Sub cNamaHadiah_ButtonClick()
 Set dbData = objData.Browse(GetDSN, "msthadiah", "keterangan,kodehadiah,poin", "keterangan", sisContent, cNamaHadiah.Text)
  If Not dbData.EOF Then
    cNamaHadiah.Text = cNamaHadiah.Browse(dbData)
    cNamaHadiah.Text = GetNull(dbData!keterangan)
    cKodeHadiah.Text = GetNull(dbData!kodehadiah)
    nPoinHadiah.Value = GetNull(dbData!poin)
    GetJumlahKanPoin
  End If
End Sub

Private Sub Form_Activate()
  If nPos = Add Then
'    GetData
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd, "SIKD"
  initvalue
  GetEdit False
  CenterForm Me
  
  TabIndex dTanggal, n
  TabIndex cCustomer, n
  TabIndex cNama, n
  TabIndex cFaktur, n
  
  TabIndex cmdAdd, n
  TabIndex cmdEdit, n
  TabIndex cmdHapus, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
End Sub

Sub GetEdit(lPar As Boolean)
  lEdit = lPar
  BiSAFrame1.Enabled = lPar
  BiSAFrame2.Enabled = lPar
  SetButton cmdSimpan, cmdKeluar, cmdAdd, cmdEdit, cmdHapus, nPos, lPar, cmdAktivasi
End Sub

Private Sub initvalue()
  cFaktur.Default
  dTanggal.Value = Date
  cCustomer.Default
  cNama.Default
  cAlamat.Default
  nPoinDitukar.Default
  cNamaDepartmen.Default
  nQty.Value = 1
  cKodeHadiah.Default
  cNamaHadiah.Default
  nPoinHadiah.Default
  nJumlah.Default
  ClearTdbgrid
  nSisaPoin.Default
  'lClose = True
End Sub

Private Sub ClearTdbgrid()
  vaArray.ReDim 0, -1, 0, 6
  TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
End Sub


Private Sub SumTDB()
Dim n As Integer
Dim nPoinReg As Double
Dim nHadiah As Integer
Dim nPoinSisa As Integer

  
  nPoinSisa = 0
  If nJumlah.Value > 0 Then
  nHadiah = nJumlah.Value
  nPoinDitukar.Value = 0
  For n = 0 To vaArray.UpperBound(1)
  
    If vaArray(n, 0) = -1 Then
      If nHadiah - vaArray(n, 5) > 0 Then
        nHadiah = nHadiah - vaArray(n, 5)
        vaArray(n, 5) = vaArray(n, 4)
        vaArray(n, 4) = 0
      Else
        If vaArray(n, 5) >= nHadiah Then
          vaArray(n, 5) = nHadiah
          vaArray(n, 4) = vaArray(n, 5) - nHadiah
        End If
      End If
    Else
      vaArray(n, 4) = vaArray(n, 6)
      vaArray(n, 5) = 0
    End If
  Next n
  
  nPoinDitukar.Value = nPoinReg
  Else
    MsgBox "Tentukan Hadiah Terlebih Dahulu"
    GetData
  End If

End Sub


Private Sub nQty_Change()
  GetJumlahKanPoin
End Sub

Private Sub GetJumlahKanPoin()
  nJumlah.Value = nPoinHadiah.Value * nQty.Value
End Sub

Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
  TDBGrid1.Update
  SumTDB
End Sub
