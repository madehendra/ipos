VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trMemberWithdraw 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Member Withdraw"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   12090
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   4065
      Left            =   135
      Top             =   1380
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   7170
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
      Begin BiSAButtonProject.BiSAButton cmdAddDetail 
         Height          =   330
         Left            =   11370
         TabIndex        =   0
         Top             =   90
         Width           =   375
         _ExtentX        =   661
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
         BackColor       =   -2147483633
         Picture         =   "trMemberWithdraw.frx":0000
      End
      Begin BiSANumberBoxProject.BiSANumberBox nTopUp 
         Height          =   330
         Left            =   9735
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
         Left            =   6390
         TabIndex        =   2
         Top             =   90
         Width           =   3330
         _ExtentX        =   5874
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
      Begin BiSATextBoxProject.BiSABrowse cKodeMember 
         Height          =   330
         Left            =   480
         TabIndex        =   3
         Top             =   90
         Width           =   1545
         _ExtentX        =   2725
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
      Begin BiSANumberBoxProject.BiSANumberBox nUrut 
         Height          =   330
         Left            =   75
         TabIndex        =   4
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
         Height          =   3045
         Left            =   60
         TabIndex        =   5
         Top             =   465
         Width           =   11670
         _ExtentX        =   20585
         _ExtentY        =   5371
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
         Columns(1).Caption=   "Member"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Nama"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Keterangan Buku Besar"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Withdraw"
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "###,###,###,##0.00"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=661"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=582"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2752"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2672"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=7752"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=7673"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=5900"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=5821"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2963"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2884"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=514"
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
         BorderStyle     =   0
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
         _StyleDefs(75)  =   ":id=41,.parent=34,.alignment=3"
         _StyleDefs(76)  =   "Named:id=42:FilterBar"
         _StyleDefs(77)  =   ":id=42,.parent=33,.alignment=3"
      End
      Begin BiSANumberBoxProject.BiSANumberBox nTotalTopUp 
         Height          =   390
         Left            =   8265
         TabIndex        =   6
         Top             =   3570
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
         Caption         =   "TOTAL"
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
      Begin BiSATextBoxProject.BiSABrowse cNamaMember 
         Height          =   330
         Left            =   2040
         TabIndex        =   7
         Top             =   90
         Width           =   4335
         _ExtentX        =   7646
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
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1260
      Left            =   135
      Top             =   105
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   2223
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
         Left            =   225
         TabIndex        =   8
         Top             =   120
         Width           =   2865
         _ExtentX        =   5054
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
      Begin BiSATextBoxProject.BiSABrowse cFaktur 
         Height          =   330
         Left            =   225
         TabIndex        =   9
         Top             =   480
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
      Begin VB.Label Label1 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   285
         TabIndex        =   16
         Top             =   900
         Width           =   3585
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   135
      Top             =   5460
      Width           =   11805
      _ExtentX        =   20823
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
         Picture         =   "trMemberWithdraw.frx":059A
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   8970
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
         Picture         =   "trMemberWithdraw.frx":0824
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
         Picture         =   "trMemberWithdraw.frx":09C3
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   435
         Left            =   105
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
         Picture         =   "trMemberWithdraw.frx":0AEF
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   10515
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
         Picture         =   "trMemberWithdraw.frx":0C9A
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   9420
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
         Picture         =   "trMemberWithdraw.frx":0D40
      End
   End
End
Attribute VB_Name = "trMemberWithdraw"
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

Private Sub cFaktur_ButtonClick()
Dim cFilterUsername As String
Dim objMenu As New CodeSuiteLibrary.Menu

  If aCfg(objData, msOtorisasiPenuh) = "Y" Then
    If GetRegistry(reg_UserLevel) <> 0 Then
      If objMenu.GetPassword("", Me, GetDSN) Then
        If objMenu.UserLevel <> 0 Then
            MsgBox "Maaf, Anda tidak diberikan wewenang untuk melakukan PENGEDITAN." & vbCrLf & _
                   "Hanya user dengan LEVEL 0 (SUPERVISOR) yg diijinkan", vbInformation, "OTORISASI not ALLOWED"
            Exit Sub
'        Else
'          MsgBox "OTORISASI DIBATALKAN", vbCritical
'          Exit Sub
        End If
      Else
        Exit Sub
      End If
    End If
  End If

  cFilterUsername = ""
  If GetRegistry(reg_UserLevel) <> 0 Then
    cFilterUsername = " and username = '" & GetRegistry(reg_Username) & "'"
  End If
  
  Set dbData = objData.Browse(GetDSN, "totmembertopup", "nomormembertopup", "tgl", sisAssign, Format(dTgl.value, "yyyy-MM-dd"), " and status = 'K'" & cFilterUsername)
  If Not dbData.EOF Then
    cFaktur.Text = cFaktur.Browse(dbData, Array("Nomor Top Up"), , Array(30))
  End If
  cFaktur_Validate True
End Sub

Private Sub GetMemory()
Dim n As Single, cSQL As String
  
  vaArray.ReDim 0, -1, 0, 5
  cSQL = "Select m.kodeanggota,a.nama,m.keterangan,m.kredit "
  cSQL = cSQL & " From membertopup m "
  cSQL = cSQL & " Left join anggota a on a.kodeanggota = m.kodeanggota "
  cSQL = cSQL & " Where m.nomormembertopup = '" & cFaktur.Text & "'"
  
  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.EOF Then
    dbData.MoveFirst
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      
      vaArray(n, 0) = n + 1
      vaArray(n, 1) = GetNull((dbData!kodeanggota), "")
      vaArray(n, 2) = GetNull((dbData!nama), "")
      vaArray(n, 3) = GetNull((dbData!keterangan), "")
      vaArray(n, 4) = GetNull(dbData!kredit)
      dbData.MoveNext
    Loop
  End If
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
  SumJumlah
End Sub

Private Sub cFaktur_Validate(Cancel As Boolean)
'   If cFaktur.LastKey = 13 Then
'      Set dbData = objData.Browse(GetDSN, "totmembertopup", , "nomormembertopup", sisAssign, cFaktur.Text)
'      If Not dbData.EOF Then
'        If nPos = Add Then 'Add
'          MsgBox "Data Sudah Ada, silahkan ulangi pengisian  !", vbInformation
'          Cancel = True
'          initvalue
'          cFaktur.SetFocus
'          Exit Sub
'        End If
'        GetMemory
'        If nPos = Delete Then DeleteTransaksi
'      ElseIf dbData.EOF And nPos <> Add Then
'        MsgBox "Data tidak Ada....!", vbInformation
'        Cancel = True
'        initvalue
'        cFaktur.SetFocus
'        Exit Sub
'      End If
'  End If

  Set dbData = objData.Browse(GetDSN, "totmembertopup", , "nomormembertopup", sisAssign, cFaktur.Text)
  If Not dbData.EOF Then
    If nPos = Add Then 'Add
      MsgBox "Data Sudah Ada, silahkan ulangi pengisian  !", vbInformation
      Cancel = True
      initvalue
      cFaktur.SetFocus
      Exit Sub
    End If
    GetMemory
    If nPos = Delete Then DeleteTransaksi
  ElseIf dbData.EOF And nPos <> Add Then
    MsgBox "Data tidak Ada....!", vbInformation
    Cancel = True
    initvalue
    cFaktur.SetFocus
    Exit Sub
  End If

End Sub

Private Sub cKodeMember_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "anggota", "kodeanggota,nama,alamat", "kodeanggota", sisContent, cKodeMember.Text)
  cKodeMember.Text = cKodeMember.Browse(dbData)
  If Not dbData.EOF Then
    cNamaMember.Text = GetNull(dbData!nama, "")
  End If
End Sub

Private Sub cmdAdd_Click()
  initvalue
  GetEdit True
  nPos = Add
  dTgl.SetFocus
  cFaktur.Text = CreateNomorFaktur(objData, sisModulTransaksi.MemberTopUp, "totmembertopup", "nomormembertopup")
  cFaktur.Enabled = False
  cFaktur.Button = False
End Sub

Private Sub cmdAddDetail_Click()
Dim n As Integer
  
  If isInGrid(vaArray, 1, cKodeMember.Text, , n) And nUrut.value > vaArray.UpperBound(1) + 1 Then
    MsgBox "Data sudah pernah dimasukkan", vbCritical
    Exit Sub
  End If
  
  If nTopUp.value <= GetSaldoTopUpMember(objData, cKodeMember.Text) Then
    If nTopUp.value <= GetSaldoTopUpMember(objData, cKodeMember.Text) - GetSaldoPiutang(objData, cKodeMember.Text) Then
      If nUrut.value > (vaArray.UpperBound(1) + 1) Then
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        n = nUrut.value - 1
      ElseIf vaArray.UpperBound(1) = -1 Then
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        n = nUrut.value - 1
      Else
        n = nUrut.value - 1
      End If
      
      vaArray(n, 0) = n + 1
      vaArray(n, 1) = cKodeMember.Text
      vaArray(n, 2) = cNamaMember.Text
      vaArray(n, 3) = cKeterangan.Text
      vaArray(n, 4) = nTopUp.value
      
      Set TDBGrid1.Array = vaArray
      TDBGrid1.ReBind
      
      SumJumlah
      Initdetail
      cNamaMember.SetFocus
      Exit Sub
    Else
      MsgBox "Maaf, masih ada piutang yg harus dilunasi sebesar " & Format(GetSaldoPiutang(objData, cKodeMember.Text), "###,###,##0.00") & vbCrLf & _
      "Sedangkan saldo top up " & cNamaMember.Text & " adalah " & Format(GetSaldoTopUpMember(objData, cKodeMember.Text), "###,###,##0.00") & vbCrLf & _
      "Maksimal uang yg boleh ditarik adalah " & Format(GetSaldoTopUpMember(objData, cKodeMember.Text) - GetSaldoPiutang(objData, cKodeMember.Text), "###,###,##0.00") & vbCrLf & vbCrLf & _
      "Menarik uang sejumlah " & Format(nTopUp.value, "###,###,##0.00") & " TIDAK DIPERBOLEHKAN", vbCritical
    End If
  Else
    MsgBox "Maaf, nilai withdraw lebih besar dari uang yg ada"
  End If
  
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
    lSave = IIf(lSave, objData.Delete(GetDSN, "totmembertopup", "nomormembertopup", sisAssign, cFaktur.Text), False)
    lSave = IIf(lSave, DelKodeTr(objData, msMemberTopUp, cFaktur.Text), False)
    lSave = IIf(lSave, objData.Delete(GetDSN, "membertopup", "nomormembertopup", sisAssign, cFaktur.Text), False)
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

Private Sub cmdSimpan_Click()
Dim vaField
Dim vaValue
Dim n As Single
Dim lSave As Boolean
Dim Faktur As String

  lSave = True
  If ValidSaving() Then
    If MsgBox("Data benar-benar sudah VALID ?", vbYesNo) = vbYes Then
        objData.Start GetDSN
        Faktur = cFaktur.Text
        lSave = IIf(lSave, DelKodeTr(objData, msMemberTopUp, Faktur), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "totmembertopup", "nomormembertopup", sisAssign, Faktur), False)
        lSave = IIf(lSave, objData.Update(GetDSN, "totmembertopup", "nomormembertopup = '" & Faktur & "'", Array("nomormembertopup", "tgl", "username", "status"), Array(Faktur, dTgl.value, GetRegistry(reg_Username), "K")), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "membertopup", "nomormembertopup ", sisAssign, Faktur), False)
        
        vaField = Array("nomormembertopup", "tgl", "kodeanggota", "keterangan", "kredit")
        
        For n = 0 To vaArray.UpperBound(1)
          
          If aCfg(objData, msMinimumDeposit) > 0 Then
            If vaArray(n, 4) > GetSaldoTopUpMember(objData, vaArray(n, 1)) - GetSaldoPiutang(objData, vaArray(n, 1)) Then
              MsgBox vaArray(n, 2) & " Masih memiliki outstanding hutang sebesar " & Format(GetSaldoPiutang(objData, vaArray(n, 1)), "###,###,##0") & vbCrLf _
              & "Maksimal dana yg boleh ditarik adalah sebesar " & Format(GetSaldoTopUpMember(objData, vaArray(n, 1)) - GetSaldoPiutang(objData, vaArray(n, 1)), "###,###,##00") & vbCrLf _
              & "Maaf, data tidak bisa disimpan "
              lSave = False
            End If
          End If

          vaValue = Array(Faktur, dTgl.value, vaArray(n, 1), vaArray(n, 3), vaArray(n, 4))
          lSave = IIf(lSave, objData.Add(GetDSN, "membertopup", vaField, vaValue), False)
          
          lSave = IIf(lSave, UpdKodeTr(objData, msMemberTopUp, Faktur, Format(dTgl.value, "yyyy-MM-dd"), GetAkunKas(objData, GetRegistry(reg_Username)), "", "Penarikan " & vaArray(n, 3) & " an " & vaArray(n, 2), 0, vaArray(n, 4)), False)
'          GetAkunKas(objData, GetRegistry(reg_UserName))
          lSave = IIf(lSave, UpdKodeTr(objData, msMemberTopUp, Faktur, Format(dTgl.value, "yyyy-MM-dd"), aCfg(objData, msRekeningTopUp), "", "Penarikan" & vaArray(n, 3) & " an " & vaArray(n, 2), vaArray(n, 4), 0), False)
          'cek satu persatu, ada hutang outstanding tidak? kalau masih ada tidak boleh disimpan
                    
        Next
        
        If lSave Then
          objData.Save GetDSN
        Else
          objData.Cancel GetDSN
        End If
        
        'Print
        trPrint4.noOrder = Faktur
        Set dbData = objData.Browse(GetDSN, "totmembertopup t", "t.*", "t.nomormembertopup", sisAssign, Faktur)
        If Not dbData.EOF Then
          Load trPrint4
          trPrint4.Show vbModal
        End If
             
      Else
        nUrut.SetFocus
        Exit Sub
      End If
    End If
    initvalue
    GetEdit False
End Sub

'Private Sub PrintStruk(ByVal Faktur As String)
'Dim n As Double
'Dim nBruto As Double
'Dim nTotQty As Double
'Dim nTotalTopUp As Double
'
'  With aMainmenu.IO1
'    .Open GetRegistry(reg_PortPrinterKasir), ""
'    .WriteString Chr(27) & Chr(15) & vbCrLf
'    .WriteString Padc(Trim("STRUK PENARIKAN UANG"), 40) & vbCrLf
'    .WriteString Padc(Trim(aCfg(objData, msNamaPerusahaan)), 40) & vbCrLf
'    .WriteString Padc(Trim(aCfg(objData, msAlamatPerusahaan)), 40) & vbCrLf
'    .WriteString Padc(Trim(aCfg(objData, msTelepon) & " " & aCfg(objData, msFax)), 40) & vbCrLf
'    .WriteString Padc(aCfg(objData, msKota), 40) & vbCrLf
'    .WriteString "Print By: " & aCfg(objData, msNama) & vbCrLf
'    .WriteString Replicate("-", 40) & vbCrLf
'    Set dbData = objData.Browse(GetDSN, "membertopup m", "m.kodeanggota,a.nama,m.keterangan,m.kredit", "m.nomormembertopup", sisAssign, Faktur, , , Array("left join anggota a on a.kodeanggota = m.kodeanggota"))
'    If Not dbData.EOF Then
'      Do While Not dbData.EOF
'        .WriteString Padl(GetNull(dbData!kodeanggota) & " " & GetNull(dbData!nama), 40) & vbCrLf
'        .WriteString Padl(GetNull(dbData!keterangan) & " " & Format(GetNull(dbData!kredit), "###,###,##0"), 40) & vbCrLf
'        n = n + 1
'        nTotalTopUp = nTotalTopUp + GetNull(dbData!kredit)
'        dbData.MoveNext
'      Loop
'    End If
'    .WriteString Replicate("-", 40) & vbCrLf
'    .WriteString Replicate("Total : " & Format(nTotalTopUp, "###,###,##0"), 40) & vbCrLf
'    .WriteString Replicate("Terimakasih, tolong disimpan bukti ini", 40) & vbCrLf
'    .WriteString Replicate("Sebagai tanda penarikan uang.", 40) & vbCrLf
'    .Close
'    OpenDrawer GetRegistry(reg_PortPrinterKasir)
'  End With
'End Sub

Private Function ValidSaving() As Boolean
  ValidSaving = True
  If vaArray.UpperBound(1) < 0 Then
    MsgBox "tidak ada data disimpan", vbCritical
    ValidSaving = False
  End If
End Function

Private Sub cNamaMember_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "anggota", "nama,kodeanggota,alamat", "nama", sisContent, cNamaMember.Text, " or kodeanggota like '%" & cNamaMember.Text & "%'")
  cNamaMember.Text = cNamaMember.Browse(dbData)
  If Not dbData.EOF Then
    cKodeMember.Text = GetNull(dbData!kodeanggota, "")
  End If
End Sub

Private Sub dTgl_Validate(Cancel As Boolean)
  If Not IsInPeriod(dTgl.value) Then
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
  
  
  TabIndex nUrut, n
  TabIndex cKodeMember, n
  TabIndex cNamaMember, n
  TabIndex cKeterangan, n
  TabIndex nTopUp, n
  TabIndex cmdAddDetail, n

  TabIndex cmdAdd, n
  TabIndex cmdEdit, n
  TabIndex cmdHapus, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
End Sub

Private Sub SumJumlah()
Dim n

  nTotalTopUp.value = 0
  For n = 0 To vaArray.UpperBound(1)
    nTotalTopUp.value = nTotalTopUp.value + vaArray(n, 4)
  Next
End Sub

Private Sub Initdetail()
  cKodeMember.Default
  cNamaMember.Default
  cKeterangan.Default
  nTopUp.Default
  nUrut.value = vaArray.UpperBound(1) + 2
End Sub

Private Sub initvalue()
  cFaktur.Default
  dTgl.value = Date
  nTopUp.value = 0
  nTotalTopUp.value = 0
  vaArray.ReDim 0, -1, 0, 4
  Set TDBGrid1.Array = vaArray
  TDBGrid1.Refresh
  TDBGrid1.ReBind
  Initdetail
End Sub

Private Sub nUrut_Change()
  If nUrut.value <= 0 Then
    nUrut.value = vaArray.UpperBound(1) + 2
  End If
End Sub

Private Sub nUrut_Validate(Cancel As Boolean)
  If nUrut.value - 1 <= vaArray.UpperBound(1) And nUrut.value >= 1 Then
    cKodeMember.Text = vaArray(nUrut.value - 1, 1)
    cNamaMember.Text = vaArray(nUrut.value - 1, 2)
    cKeterangan.Text = vaArray(nUrut.value - 1, 3)
    nTopUp.value = vaArray(nUrut.value - 1, 4)
  ElseIf nUrut.value - 2 > vaArray.UpperBound(1) Or nUrut.value <= 0 Then
    nUrut.value = vaArray.UpperBound(1) + 2
  End If
End Sub


Private Sub GetEdit(lPar As Boolean)
  BiSAFrame1.Enabled = lPar
  BiSAFrame2.Enabled = lPar
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
    nUrut.value = vaArray.UpperBound(1) + 2
  End If
  SumJumlah
End Sub


