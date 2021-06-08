VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trJurnalUmum 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Jurnal Umum"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   17100
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   17100
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   4575
      Left            =   135
      Top             =   2070
      Width           =   16800
      _ExtentX        =   29633
      _ExtentY        =   8070
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
      Begin BiSANumberBoxProject.BiSANumberBox nKredit 
         Height          =   330
         Left            =   14700
         TabIndex        =   0
         Top             =   90
         Width           =   1560
         _ExtentX        =   2752
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
      Begin BiSANumberBoxProject.BiSANumberBox nDebet 
         Height          =   330
         Left            =   13065
         TabIndex        =   1
         Top             =   90
         Width           =   1605
         _ExtentX        =   2831
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
      Begin BiSATextBoxProject.BiSATextBox cKeterangan 
         Height          =   330
         Left            =   7290
         TabIndex        =   2
         Top             =   90
         Width           =   5730
         _ExtentX        =   10107
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
         FontName        =   "Tahoma"
         MaxLength       =   150
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
      Begin BiSATextBoxProject.BiSATextBox cNama 
         Height          =   330
         Left            =   2835
         TabIndex        =   3
         Top             =   90
         Width           =   4395
         _ExtentX        =   7752
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
         FontName        =   "Tahoma"
         BackColor       =   12632256
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
      Begin BiSATextBoxProject.BiSABrowse cRekening 
         Height          =   330
         Left            =   480
         TabIndex        =   4
         Top             =   90
         Width           =   2325
         _ExtentX        =   4101
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
         FontName        =   "Tahoma"
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
      Begin BiSANumberBoxProject.BiSANumberBox nUrut 
         Height          =   330
         Left            =   75
         TabIndex        =   5
         Top             =   90
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   582
         Appearance      =   0
         Decimals        =   0
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         Height          =   3630
         Left            =   60
         TabIndex        =   6
         Top             =   465
         Width           =   16620
         _ExtentX        =   29316
         _ExtentY        =   6403
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "No."
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Rekening"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Nama Rekening"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Keterangan Buku Besar"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Debet"
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "FormatText Event"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Kredit"
         Columns(5).DataField=   ""
         Columns(5).NumberFormat=   "FormatText Event"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=661"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=582"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=4260"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4180"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=7805"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=7726"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=10186"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=10107"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2963"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2884"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=514"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=2858"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2778"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=514"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
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
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1,5
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=112,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFCFCED&,.bold=0,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bgcolor=&HEBDACB&"
         _StyleDefs(11)  =   ":id=2,.fgcolor=&H8000000D&,.bold=0,.fontsize=825,.italic=0,.underline=0"
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
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=58,.parent=13"
         _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
         _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
         _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
         _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
         _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
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
         _StyleDefs(79)  =   ":id=41,.parent=34,.alignment=3"
         _StyleDefs(80)  =   "Named:id=42:FilterBar"
         _StyleDefs(81)  =   ":id=42,.parent=33,.alignment=3"
      End
      Begin BiSANumberBoxProject.BiSANumberBox nTotDebet 
         Height          =   390
         Left            =   10095
         TabIndex        =   7
         Top             =   4125
         Width           =   3240
         _ExtentX        =   5715
         _ExtentY        =   688
         Appearance      =   0
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12632256
         ForeColor       =   -2147483635
         Caption         =   "DEBET"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin BiSANumberBoxProject.BiSANumberBox nTotKredit 
         Height          =   390
         Left            =   13455
         TabIndex        =   8
         Top             =   4125
         Width           =   3240
         _ExtentX        =   5715
         _ExtentY        =   688
         Appearance      =   0
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12632256
         ForeColor       =   -2147483635
         Caption         =   "KREDIT"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin BiSAButtonProject.BiSAButton cmdAddDetail 
         Height          =   345
         Left            =   16290
         TabIndex        =   18
         Top             =   90
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   609
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
         Picture         =   "trJurnalUmum.frx":0000
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1980
      Left            =   135
      Top             =   75
      Width           =   16800
      _ExtentX        =   29633
      _ExtentY        =   3493
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
      Begin BiSADateProject.BiSADate dTgl 
         Height          =   330
         Left            =   405
         TabIndex        =   9
         Top             =   495
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
         Caption         =   "Tanggal"
         CaptionWidth    =   1300
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin BiSATextBoxProject.BiSABrowse cCostCenter 
         Height          =   330
         Left            =   405
         TabIndex        =   16
         Top             =   1260
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
         Caption         =   "Cost Centre"
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
         Left            =   405
         TabIndex        =   17
         Top             =   870
         Width           =   3675
         _ExtentX        =   6482
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
         Caption         =   "Faktur/No"
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
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   645
      Left            =   120
      Top             =   6675
      Width           =   16845
      _ExtentX        =   29713
      _ExtentY        =   1138
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
         Left            =   2235
         TabIndex        =   10
         Top             =   105
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
         Picture         =   "trJurnalUmum.frx":059A
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   3405
         TabIndex        =   11
         Top             =   105
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
         Picture         =   "trJurnalUmum.frx":0824
      End
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   1185
         TabIndex        =   12
         Top             =   105
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
         Picture         =   "trJurnalUmum.frx":09C3
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   435
         Left            =   120
         TabIndex        =   13
         Top             =   105
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
         Picture         =   "trJurnalUmum.frx":0AEF
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   15585
         TabIndex        =   14
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
         Picture         =   "trJurnalUmum.frx":0C9A
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   14490
         TabIndex        =   15
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
         Picture         =   "trJurnalUmum.frx":0D40
      End
   End
End
Attribute VB_Name = "trJurnalUmum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vaArray As New XArrayDB
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim lEdit As Boolean
Dim nPos As SisPos

Private Sub cCostCenter_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "costcenter", "kodecostcenter,keterangan", "kodecostcenter", sisContent, cCostCenter.Text)
  If Not dbData.EOF Then
    cCostCenter.Text = cCostCenter.Browse(dbData)
  End If
End Sub

Private Sub cFaktur_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "totjurnalumum", "nomorjurnalumum,kodecostcenter", "tgl", sisAssign, Format(dTgl.Value, "yyyy-MM-dd"))
  If Not dbData.EOF Then
    cFaktur.Text = cFaktur.Browse(dbData, Array("Nomor Jurnal", "Cost Centre"), , Array(25, 10))
    GetMemory
  End If
End Sub

Private Sub cFaktur_Validate(Cancel As Boolean)
  If cFaktur.LastKey = 13 Then
'      cFaktur.Text = SisModulTransaksi.JurnalUmum & Format(dTgl.Value, "YYYYMM") & Padl(cFaktur.Text, 6, "0")
      Set dbData = objData.Browse(GetDSN, "totjurnalumum", , "nomorjurnalumum", sisAssign, cFaktur.Text)
      If Not dbData.EOF Then
        If nPos = Add Then 'Add
          MsgBox "Data Sudah Ada, silahkan ulangi pengisian  !", vbInformation
          Cancel = True
          initvalue
          cFaktur.SetFocus
          Exit Sub
        End If
        cCostCenter.Text = GetNull(dbData!kodecostcenter)
        GetMemory
        If nPos = Delete Then DeleteTransaksi
      ElseIf dbData.EOF And nPos <> Add Then
        MsgBox "Data tidak Ada....!", vbInformation
        Cancel = True
        initvalue
        cFaktur.SetFocus
        Exit Sub
      End If
  End If
End Sub

Private Sub GetMemory()
Dim n As Single, cSQL As String
  
  vaArray.ReDim 0, -1, 0, 5
  cSQL = "Select j.kodeakun,j.keterangan as KeteranganBukuBesar,j.debet,j.kredit,r.keterangan as NamaRekening"
  cSQL = cSQL & " From jurnalumum j"
  cSQL = cSQL & " Left join akun r on r.kodeakun = j.kodeakun"
  cSQL = cSQL & " Where nomorjurnalumum = '" & cFaktur.Text & "'"
  cSQL = cSQL & " Order By j.nomorjurnalumum desc"
  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.EOF Then
    dbData.MoveFirst
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      
      vaArray(n, 0) = n + 1
      vaArray(n, 1) = GetNull((dbData!kodeakun), "")
      vaArray(n, 2) = GetNull((dbData!NamaRekening), "")
      vaArray(n, 3) = GetNull((dbData!KeteranganBukuBesar), "")
      vaArray(n, 4) = GetNull(dbData!debet)
      vaArray(n, 5) = GetNull(dbData!kredit)
      dbData.MoveNext
    Loop
  End If
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
  SumJumlah
End Sub

Private Sub cmdAdd_Click()
  initvalue
  GetEdit True
  nPos = Add
  dTgl.SetFocus
  cFaktur.Text = CreateNomorFaktur(objData, sisModulTransaksi.jurnalumum, "totjurnalumum", "nomorjurnalumum")
  cFaktur.Enabled = False
  cFaktur.Button = False
End Sub

Private Sub cmdAddDetail_Click()
Dim n As Integer

  If Trim(cRekening.Text) = "" Then
    MsgBox "Rekening harus diisi...", vbInformation
    cRekening.SetFocus
    Exit Sub
  End If
  
  If nDebet.Value < 0 Then
    MsgBox "Nilai Debet tidak valid...", vbInformation
    nDebet.SetFocus
    Exit Sub
  End If
  
  If nKredit.Value < 0 Then
    MsgBox "Nilai Kredit tidak valid...", vbInformation
    nKredit.SetFocus
    Exit Sub
  End If
  
  If nDebet.Value <= 0 And nKredit.Value <= 0 Then
    MsgBox "Nilai Debet atau Kredit harus diisi...", vbInformation
    nDebet.SetFocus
    Exit Sub
  End If
  
  If nDebet.Value > 0 And nKredit.Value > 0 Then
    MsgBox "Nilai Debet atau Kredit tidak boleh diisi dua-duanya (Harus salah satu !)...", vbInformation
    nDebet.SetFocus
    Exit Sub
  End If
  
  If Not GetValidDataBrowse(objData, "akun", "kodeakun", cRekening.Text) Then
    MsgBox "Akun Jurnal Tidak Ada Dalam Database"
    cRekening.SetFocus
    Exit Sub
  End If
  
  If nUrut.Value > (vaArray.UpperBound(1) + 1) Then
    vaArray.InsertRows vaArray.UpperBound(1) + 1
    n = nUrut.Value - 1
  ElseIf vaArray.UpperBound(1) = -1 Then
    vaArray.InsertRows vaArray.UpperBound(1) + 1
    n = nUrut.Value - 1
  Else
    n = nUrut.Value - 1
  End If
  
  vaArray(n, 0) = n + 1
  vaArray(n, 1) = cRekening.Text
  vaArray(n, 2) = cNama.Text
  vaArray(n, 3) = cKeterangan.Text
  vaArray(n, 4) = nDebet.Value
  vaArray(n, 5) = nKredit.Value
  
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
  
  SumJumlah
  Initdetail
  cRekening.SetFocus
  Exit Sub
End Sub

Private Sub cmdAktivasi_Click()
  frmAktivasi.Action Me
End Sub

Private Sub cmdEdit_Click()
  initvalue
  GetEdit True
  nPos = Edit
  dTgl.SetFocus
  cFaktur.Enabled = True
  cFaktur.Button = True
End Sub

Private Sub cmdHapus_Click()
  initvalue
  GetEdit True
  nPos = Delete
  dTgl.SetFocus
  cFaktur.Enabled = True
  cFaktur.Button = True
End Sub

Private Sub DeleteTransaksi()
Dim lSave As Boolean
  lSave = True
  Me.Refresh
  If MsgBox("Data Benar-benar Dihapus ?", vbQuestion + vbYesNo) = vbYes Then
    objData.Start GetDSN
    lSave = IIf(lSave, objData.Delete(GetDSN, "jurnalumum", "nomorjurnalumum", sisAssign, cFaktur.Text), False)
    lSave = IIf(lSave, DelKodeTr(objData, msJurnalUmum, cFaktur.Text), False)
    lSave = IIf(lSave, objData.Delete(GetDSN, "totjurnalumum", "nomorjurnalumum", sisAssign, cFaktur.Text), False)
    If lSave Then
      objData.Save GetDSN
    Else
      objData.Cancel GetDSN
    End If
  End If
  initvalue
  GetEdit False
End Sub

Private Sub cmdKeluar_Click()
  If Not lEdit Then
    Unload Me
  Else
    GetEdit False
    initvalue
    cFaktur.Button = False
  End If
End Sub

Private Sub cmdOK_Click()

End Sub

Private Sub cmdSimpan_Click()
Dim vaField
Dim vaValue
Dim n As Single
Dim lSave As Boolean
Dim Faktur As String

  lSave = True
  If nTotDebet.Value <> nTotKredit.Value Then
    If MsgBox("Jurnal Tidak Balance, Apakah akan dilanjutkan") = vbNo Then
      nUrut.SetFocus
      Exit Sub
    End If
  End If
  
  If ValidSaving() Then
    If MsgBox("Data benar-benar sudah VALID ?", vbYesNo) = vbYes Then
        objData.Start GetDSN
        
        Faktur = cFaktur.Text
        lSave = IIf(lSave, DelKodeTr(objData, msJurnalUmum, Faktur), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "totjurnalumum", "nomorjurnalumum", sisAssign, Faktur), False)
        lSave = IIf(lSave, objData.Update(GetDSN, "totjurnalumum", "nomorjurnalumum = '" & Faktur & "'", Array("nomorjurnalumum", "kodecostcenter", "tgl", "username"), Array(Faktur, cCostCenter.Text, dTgl.Value, GetRegistry(reg_Username))), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "jurnalumum", "nomorjurnalumum", sisAssign, Faktur), False)
        
        vaField = Array("nomorjurnalumum", "tgl", "kodeakun", "keterangan", "debet", "kredit")
        For n = 0 To vaArray.UpperBound(1)
          vaValue = Array(Faktur, dTgl.Value, vaArray(n, 1), vaArray(n, 3), vaArray(n, 4), vaArray(n, 5))
          lSave = IIf(lSave, objData.Add(GetDSN, "jurnalumum", vaField, vaValue), False)
          lSave = IIf(lSave, UpdKodeTr(objData, msJurnalUmum, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), vaArray(n, 1), cCostCenter.Text, vaArray(n, 3), vaArray(n, 4), vaArray(n, 5)), False)
        Next
        If lSave Then
          objData.Save GetDSN
        Else
          objData.Cancel GetDSN
        End If
      Else
        nUrut.SetFocus
        Exit Sub
      End If
    End If
    initvalue
    GetEdit False
End Sub

Private Function ValidSaving() As Boolean
  ValidSaving = True

End Function

Private Sub cRekening_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", "kodeakun,keterangan,jenis", "kodeakun", sisContent, cRekening.Text, " or keterangan like '%" & cRekening.Text & "%' AND jenis ='D'", "kodeakun")
  cRekening.Text = cRekening.Browse(dbData)
  If Not dbData.EOF Then
    cNama.Text = GetNull(dbData!keterangan, "")
  End If
End Sub

Private Sub dTgl_Validate(Cancel As Boolean)
  If Not IsInPeriod(dTgl.Value) Then
    Cancel = True
    dTgl.SetFocus
    cFaktur_Validate True
  End If
End Sub

Private Sub Form_Load()
Dim n As Single
  
'  If CheckTrial(nRecordsTrial, TrialJurnalUmum) = True Then
'    End
'  End If
  
  CenterForm Me
  SetIcon Me.hWnd
  initvalue
  GetEdit False
  
  TabIndex dTgl, n
  TabIndex cFaktur, n
  TabIndex cCostCenter, n
  
  TabIndex nUrut, n
  TabIndex cRekening, n
  TabIndex cNama, n
  TabIndex cKeterangan, n
  TabIndex nDebet, n
  TabIndex nKredit, n
  TabIndex cmdAddDetail, n

  TabIndex cmdAdd, n
  TabIndex cmdEdit, n
  TabIndex cmdHapus, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
End Sub

Private Sub SumJumlah()
Dim n

  nTotDebet.Value = 0
  nTotKredit.Value = 0
  For n = 0 To vaArray.UpperBound(1)
    nTotDebet.Value = nTotDebet.Value + vaArray(n, 4)
    nTotKredit.Value = nTotKredit.Value + vaArray(n, 5)
  Next
End Sub

Private Sub Initdetail()
  cRekening.Default
  cNama.Default
  cKeterangan.Default
  nDebet.Value = 0
  nKredit.Value = 0
  nUrut.Value = vaArray.UpperBound(1) + 2
End Sub

Private Sub initvalue()
  cFaktur.Default
  dTgl.Value = Date
  nTotDebet.Value = 0
  nTotKredit.Value = 0
  Set dbData = objData.Browse(GetDSN, "costcenter", , "kodecostcenter", sisAssign, GetCostCenterUser(objData, GetRegistry(reg_Username)))
  If Not dbData.EOF Then
    cCostCenter.Text = GetNull(dbData!kodecostcenter)
  End If
  vaArray.ReDim 0, -1, 0, 5
  Set TDBGrid1.Array = vaArray
  TDBGrid1.Refresh
  TDBGrid1.ReBind
  Initdetail
End Sub

Private Sub nUrut_Change()
  If nUrut.Value <= 0 Then
    nUrut.Value = vaArray.UpperBound(1) + 2
  End If
End Sub

Private Sub nUrut_Validate(Cancel As Boolean)


  If nUrut.Value - 1 <= vaArray.UpperBound(1) And nUrut.Value >= 1 Then
    cRekening.Text = vaArray(nUrut.Value - 1, 1)
    cNama.Text = vaArray(nUrut.Value - 1, 2)
    cKeterangan.Text = vaArray(nUrut.Value - 1, 3)
    nDebet.Value = vaArray(nUrut.Value - 1, 4)
    nKredit.Value = vaArray(nUrut.Value - 1, 5)
  ElseIf nUrut.Value - 2 > vaArray.UpperBound(1) Or nUrut.Value <= 0 Then
    nUrut.Value = vaArray.UpperBound(1) + 2
  End If
End Sub

Private Sub TDBGrid1_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
  If Val(Value) = 0 Then
    Value = ""
  Else
    Value = Format(Value, "###,###,###,###,##0.00")
  End If
End Sub

Private Sub GetEdit(lPar As Boolean)
  BiSAFrame1.Enabled = lPar
  BisaFrame2.Enabled = lPar
  lEdit = lPar
  SetButton cmdSimpan, cmdKeluar, cmdAdd, cmdEdit, cmdHapus, nPos, lPar, cmdAktivasi
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim n As Integer

  If KeyCode = vbKeyDelete Then
    TDBGrid1.Delete
    
    For n = 0 To vaArray.UpperBound(1)
      vaArray(n, 0) = n + 1
    Next
    
    Set TDBGrid1.Array = vaArray
    TDBGrid1.ReBind
    nUrut.Value = vaArray.UpperBound(1) + 2
  End If
  SumJumlah
End Sub
