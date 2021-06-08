VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form mstAkunKas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AKUN KAS"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   12975
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   12975
   Begin VB.OptionButton optModePenjualan 
      Caption         =   "&0 Full"
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
      Left            =   1185
      TabIndex        =   13
      Top             =   2745
      Width           =   795
   End
   Begin VB.OptionButton optModePenjualan 
      Caption         =   "&1 Compact"
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
      Left            =   2025
      TabIndex        =   12
      Top             =   2745
      Width           =   1524
   End
   Begin BiSATextBoxProject.BiSABrowse cAkunKas 
      Height          =   330
      Left            =   60
      TabIndex        =   4
      Top             =   570
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
      Caption         =   "Akun Kas"
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
   Begin BiSATextBoxProject.BiSATextBox cUserName 
      Height          =   330
      Left            =   60
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   225
      Width           =   3900
      _ExtentX        =   6879
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
      Caption         =   "User Name"
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
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   45
      Top             =   6735
      Width           =   12855
      _ExtentX        =   22675
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
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   11655
         TabIndex        =   0
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
         Picture         =   "mstAkunKas.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   10575
         TabIndex        =   1
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
         Picture         =   "mstAkunKas.frx":00A6
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   3510
      Left            =   30
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3165
      Width           =   12870
      _ExtentX        =   22701
      _ExtentY        =   6191
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Username"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Akun Kas"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Cost Center"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Gudang"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Akun Setoran"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Mode Penjualan"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Kode Salesman"
      Columns(6).DataField=   ""
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=4207"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4128"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=3334"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3254"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=2540"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2461"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=2566"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2487"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(25)=   "Column(4).Width=2725"
      Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=516"
      Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(31)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=516"
      Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(37)=   "Column(6).Width=4180"
      Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=4101"
      Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=516"
      Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageFooterFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=900,.italic=0"
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
      _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=50,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
      _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
      _StyleDefs(62)  =   "Named:id=33:Normal"
      _StyleDefs(63)  =   ":id=33,.parent=0"
      _StyleDefs(64)  =   "Named:id=34:Heading"
      _StyleDefs(65)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(66)  =   ":id=34,.wraptext=-1"
      _StyleDefs(67)  =   "Named:id=35:Footing"
      _StyleDefs(68)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(69)  =   "Named:id=36:Selected"
      _StyleDefs(70)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(71)  =   "Named:id=37:Caption"
      _StyleDefs(72)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(73)  =   "Named:id=38:HighlightRow"
      _StyleDefs(74)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(75)  =   "Named:id=39:EvenRow"
      _StyleDefs(76)  =   ":id=39,.parent=33,.bgcolor=&HE6E6E6&"
      _StyleDefs(77)  =   "Named:id=40:OddRow"
      _StyleDefs(78)  =   ":id=40,.parent=33"
      _StyleDefs(79)  =   "Named:id=41:RecordSelector"
      _StyleDefs(80)  =   ":id=41,.parent=34"
      _StyleDefs(81)  =   "Named:id=42:FilterBar"
      _StyleDefs(82)  =   ":id=42,.parent=33"
   End
   Begin BiSATextBoxProject.BiSATextBox cNamaAkunKas 
      Height          =   330
      Left            =   2880
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   570
      Width           =   2775
      _ExtentX        =   4895
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
   Begin BiSATextBoxProject.BiSABrowse cKodeCostCenter 
      Height          =   330
      Left            =   60
      TabIndex        =   6
      Top             =   930
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
      Caption         =   "Cost Center"
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
   Begin BiSATextBoxProject.BiSATextBox cNamaCostCenter 
      Height          =   330
      Left            =   2880
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   930
      Width           =   2775
      _ExtentX        =   4895
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
   Begin BiSATextBoxProject.BiSATextBox cNamaGudang 
      Height          =   330
      Left            =   2880
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1290
      Width           =   2775
      _ExtentX        =   4895
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
   Begin BiSATextBoxProject.BiSABrowse cKodeGudang 
      Height          =   330
      Left            =   60
      TabIndex        =   9
      Top             =   1290
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
      Caption         =   "Gudang"
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
   Begin BiSATextBoxProject.BiSATextBox cNamaAkunSetoran 
      Height          =   330
      Left            =   2880
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1635
      Width           =   2775
      _ExtentX        =   4895
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
   Begin BiSATextBoxProject.BiSABrowse cAkunSetoran 
      Height          =   330
      Left            =   60
      TabIndex        =   11
      Top             =   1635
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
      Caption         =   "Setoran"
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
   Begin BiSATextBoxProject.BiSABrowse cKodeSalesman 
      Height          =   330
      Left            =   60
      TabIndex        =   15
      Top             =   1995
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
      Caption         =   "Salesman"
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
   Begin BiSATextBoxProject.BiSATextBox cNamaSalesman 
      Height          =   330
      Left            =   2880
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1995
      Width           =   2775
      _ExtentX        =   4895
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
   Begin VB.Label Label1 
      Caption         =   "MODE PENJUALAN"
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
      Left            =   1200
      TabIndex        =   14
      Top             =   2490
      Width           =   1755
   End
End
Attribute VB_Name = "mstAkunKas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub cAkunKas_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", "kodeakun,keterangan", "kodeakun", sisPrefix, "1.", " and jenis = 'D'", "kodeakun")
  If Not dbData.EOF Then
    cAkunKas.Text = cAkunKas.Browse(dbData, Array("Kode Akun", "Keterangan"), , Array(15, 25))
    cNamaAkunKas.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub cAkunSetoran_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", "kodeakun,keterangan", "(kodeakun", sisContent, cAkunSetoran.Text, " or keterangan like '%" & cAkunSetoran.Text & "%') and jenis = 'D' and kodeakun like '1.%'", "kodeakun")
  If Not dbData.EOF Then
    cAkunSetoran.Text = cAkunSetoran.Browse(dbData, Array("KODE", "KETERANGAN"), , Array(15, 25))
    cNamaAkunSetoran.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub cKodeCostCenter_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "costcenter", "kodecostcenter,keterangan", "kodecostcenter", sisContent, cKodeCostCenter.Text, " or keterangan like '%" & cKodeCostCenter.Text & "%'", "kodecostcenter")
  If Not dbData.EOF Then
    cKodeCostCenter.Text = cAkunKas.Browse(dbData, Array("KODE", "KETERANGAN"), , Array(15, 25))
    cNamaCostCenter.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub cKodeGudang_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "gudang", "kodegudang,keterangan", "(kodegudang", sisContent, cKodeGudang.Text, " or keterangan like '%" & cKodeGudang.Text & "%') and lstatus = 'A'", "kodegudang")
  If Not dbData.EOF Then
    cKodeGudang.Text = cKodeGudang.Browse(dbData, Array("KODE", "KETERANGAN"), , Array(15, 25))
    cNamaGudang.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub cKodeSalesman_ButtonClick()
 Set dbData = objData.Browse(GetDSN, "salesman", "kodesalesman,nama", "(kodesalesman", sisContent, cKodeSalesman.Text, " or nama like '%" & cKodeSalesman.Text & "%')", "kodesalesman")
  If Not dbData.EOF Then
    cKodeSalesman.Text = cKodeSalesman.Browse(dbData, Array("KODE", "NAMA"), , Array(15, 25))
    cNamaSalesman.Text = GetNull(dbData!nama)
  End If
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdSimpan_Click()
Dim lSave As Boolean
  
  lSave = True
  objData.Start GetDSN
  lSave = IIf(lSave, objData.Update(GetDSN, "akunkas", "username = '" & cUserName.Text & "'", Array("username", "kodeakun", "kodegudang", "kodecostcenter", "akunsetoran", "modepenjualan", "kodesalesman"), Array(cUserName.Text, cAkunKas.Text, cKodeGudang.Text, cKodeCostCenter.Text, cAkunSetoran.Text, GetOpt(optModePenjualan), cKodeSalesman.Text)), False)
  If lSave Then
    objData.Save GetDSN
  Else
    objData.Cancel GetDSN
  End If
  GetLoadRows
  initvalue
End Sub

Private Sub Form_Load()
Dim n As Single
  
  SetIcon Me.hWnd
  CenterForm Me
  initvalue
  TabIndex cUserName, n
  TabIndex cAkunKas, n
  TabIndex cKodeCostCenter, n
  TabIndex cKodeGudang, n
  TabIndex cAkunSetoran, n
  TabIndex cKodeSalesman, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  GetLoadRows
End Sub

Private Sub initvalue()
  cUserName.Default
  cAkunKas.Default
  cKodeCostCenter.Default
  cKodeGudang.Default
  cAkunSetoran.Default
End Sub

Private Sub GetLoadRows()
Dim n As Integer

  vaArray.ReDim 0, -1, 0, 6
  Set dbData = objData.Browse(GetDSN, "username u", "u.username,a.kodeakun,a.kodegudang,a.kodecostcenter,a.akunsetoran,a.modepenjualan,a.kodesalesman", , , , , , Array("LEFT JOIN akunkas a ON a.username = u.username"))
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!UserName)
      vaArray(n, 1) = GetNull(dbData!kodeakun, "")
      vaArray(n, 2) = GetNull(dbData!kodecostcenter, "")
      vaArray(n, 3) = GetNull(dbData!kodegudang, "")
      vaArray(n, 4) = GetNull(dbData!akunsetoran, "")
      vaArray(n, 5) = GetNull(dbData!modepenjualan, 0)
      vaArray(n, 6) = GetNull(dbData!kodesalesman, "")
      dbData.MoveNext
    Loop
  End If
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
  TDBGrid1.Refresh
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim n As Integer
Dim nQtyTmp As Integer

On Error Resume Next


Dim lSave As Boolean
  

  If KeyCode = vbKeyDelete Then
    If vaArray.UpperBound(1) >= 0 Then
    
    lSave = True
    objData.Start GetDSN
    lSave = IIf(lSave, objData.Delete(GetDSN, "akunkas", "username", sisAssign, cUserName.Text), False)
    If lSave Then
      If MsgBox("Yakin data akan dihapus", vbYesNo, "Hapus data") = vbYes Then
      objData.Save GetDSN
      Else
        objData.Cancel GetDSN
      End If
    Else
      objData.Cancel GetDSN
    End If
    GetLoadRows
    initvalue

    
    
'      TDBGrid1.Delete
'      TDBGrid1.Update
'      SumTotal
'      For n = 0 To vaArray.UpperBound(1)
'        vaArray(n, 0) = n + 1
'        nQtyTmp = nQtyTmp + vaArray(n, 3)
'      Next
'      TDBGrid1.Columns(3).FooterText = Format(nQtyTmp, "###,###,##0")
'      nNomor.Value = vaArray.UpperBound(1) + 2
'      nPoinReguler.Value = GetHitungPoinHadiah(aCfg(objData, msKelipatan))
      TDBGrid1.ReBind
    End If
'    If vaArray.UpperBound(1) < 0 Then
'      cNamaCustomer.Enabled = True
'      cNamaCustomer.Button = True
'    End If
  End If
  
'  If lEdit = True Then
'    If KeyCode = vbKeyF3 Then
'        If vaArray.UpperBound(1) >= 0 Then
'          nNomor.Value = TDBGrid1.Columns(0).Text
'          nNomor_Validate True
'          nQty.SetFocus
'        End If
'    End If
'    If KeyCode = vbKeyReturn Then
'        If vaArray.UpperBound(1) >= 0 Then
'          nNomor.Value = TDBGrid1.Columns(0).Text
'          nNomor_Validate True
'          nQty.SetFocus
'        End If
'    End If
'  End If
'
'  If KeyCode = vbKeyEscape Then
'    InitValue1
'    nNomor.Value = vaArray.UpperBound(1) + 2
'    cBarcode.SetFocus
'    nPoinReguler.Value = GetHitungPoinHadiah(aCfg(objData, msKelipatan))
'  End If
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim db As New ADODB.Recordset
  
  cUserName.Text = GetNull(TDBGrid1.Columns(0).value)
  cAkunKas.Text = GetNull(TDBGrid1.Columns(1).value)
  cKodeCostCenter.Text = GetNull(TDBGrid1.Columns(2).value)
  cKodeGudang.Text = GetNull(TDBGrid1.Columns(3).value)
  cAkunSetoran.Text = GetNull(TDBGrid1.Columns(4).value)
  SetOpt optModePenjualan, GetNull(TDBGrid1.Columns(5).value)
  Set db = objData.Browse(GetDSN, "akun", , "kodeakun", sisAssign, GetNull(TDBGrid1.Columns(1).value))
  If Not db.EOF Then
    cNamaAkunKas.Text = GetNull(db!keterangan)
  Else
    cNamaAkunKas.Default
  End If
  Set db = objData.Browse(GetDSN, "costcenter", , "kodecostcenter", sisAssign, GetNull(TDBGrid1.Columns(2).value))
  If Not db.EOF Then
    cNamaCostCenter.Text = GetNull(db!keterangan)
  Else
    cNamaCostCenter.Default
  End If
  Set db = objData.Browse(GetDSN, "gudang", , "kodegudang", sisAssign, GetNull(TDBGrid1.Columns(3).value))
  If Not db.EOF Then
    cNamaGudang.Text = GetNull(db!keterangan)
  Else
    cNamaGudang.Default
  End If
  Set db = objData.Browse(GetDSN, "akun", , "kodeakun", sisAssign, GetNull(TDBGrid1.Columns(4).value))
  If Not db.EOF Then
    cNamaAkunSetoran.Text = GetNull(db!keterangan)
  Else
    cNamaAkunSetoran.Default
  End If
  
  Set db = objData.Browse(GetDSN, "salesman", , "kodesalesman", sisAssign, GetNull(TDBGrid1.Columns(6).value))
  If Not db.EOF Then
    cNamaSalesman.Text = GetNull(db!nama)
    cKodeSalesman.Text = GetNull(db!kodesalesman)
  Else
    cKodeSalesman.Text = ""
    cNamaSalesman.Default
  End If
End Sub
