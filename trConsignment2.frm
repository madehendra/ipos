VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trConsignment2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consigment Payment.."
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11820
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   11820
   Begin BiSAFramProject.BiSAFrame BisaFrame2 
      Height          =   4065
      Left            =   0
      Top             =   1920
      Width           =   11790
      _ExtentX        =   20796
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
      BorderStyle     =   4
      BackColor       =   -2147483633
      Begin VB.CheckBox Check2 
         Caption         =   "Clear All"
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
         Left            =   1170
         TabIndex        =   23
         Top             =   105
         Width           =   1230
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Select All"
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
         Left            =   105
         TabIndex        =   22
         Top             =   105
         Width           =   1230
      End
      Begin BiSANumberBoxProject.BiSANumberBox nTunai 
         Height          =   330
         Left            =   9735
         TabIndex        =   0
         Top             =   3630
         Width           =   1980
         _ExtentX        =   3493
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
         Height          =   3030
         Left            =   90
         TabIndex        =   1
         Top             =   330
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   5345
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   4
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "NO."
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "KODE BRG"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "NAMA BARANG"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "QTY"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "SATUAN"
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "HARGA"
         Columns(6).DataField=   ""
         Columns(6).NumberFormat=   "###,###,###,###,##0.00"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "DISC"
         Columns(7).DataField=   ""
         Columns(7).NumberFormat=   "###,###,###,###,##0.00"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "JUMLAH"
         Columns(8).DataField=   ""
         Columns(8).NumberFormat=   "###,###,###,###,##0.00"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   9
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=9"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=582"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=503"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=926"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=847"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=514"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=3016"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2937"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=512"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=5477"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=5398"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=512"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=1455"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=1376"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=197122"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=1482"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=1402"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=513"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=2593"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2514"
         Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=514"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(36)=   "Column(7).Width=1296"
         Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=1217"
         Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=514"
         Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(41)=   "Column(8).Width=3149"
         Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=3069"
         Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=514"
         Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
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
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=70,.parent=13"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=67,.parent=14"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=68,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=69,.parent=17"
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.alignment=1"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=0"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=0"
         _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
         _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
         _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
         _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=66,.parent=13,.alignment=1"
         _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=14"
         _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=15,.alignment=1"
         _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=17"
         _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=2"
         _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
         _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
         _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
         _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=14"
         _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=15"
         _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=17"
         _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=58,.parent=13,.alignment=1"
         _StyleDefs(66)  =   "Splits(0).Columns(7).HeadingStyle:id=55,.parent=14"
         _StyleDefs(67)  =   "Splits(0).Columns(7).FooterStyle:id=56,.parent=15"
         _StyleDefs(68)  =   "Splits(0).Columns(7).EditorStyle:id=57,.parent=17"
         _StyleDefs(69)  =   "Splits(0).Columns(8).Style:id=62,.parent=13,.alignment=1"
         _StyleDefs(70)  =   "Splits(0).Columns(8).HeadingStyle:id=59,.parent=14"
         _StyleDefs(71)  =   "Splits(0).Columns(8).FooterStyle:id=60,.parent=15"
         _StyleDefs(72)  =   "Splits(0).Columns(8).EditorStyle:id=61,.parent=17"
         _StyleDefs(73)  =   "Named:id=33:Normal"
         _StyleDefs(74)  =   ":id=33,.parent=0"
         _StyleDefs(75)  =   "Named:id=34:Heading"
         _StyleDefs(76)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(77)  =   ":id=34,.wraptext=-1"
         _StyleDefs(78)  =   "Named:id=35:Footing"
         _StyleDefs(79)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(80)  =   "Named:id=36:Selected"
         _StyleDefs(81)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(82)  =   "Named:id=37:Caption"
         _StyleDefs(83)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(84)  =   "Named:id=38:HighlightRow"
         _StyleDefs(85)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(86)  =   "Named:id=39:EvenRow"
         _StyleDefs(87)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(88)  =   "Named:id=40:OddRow"
         _StyleDefs(89)  =   ":id=40,.parent=33"
         _StyleDefs(90)  =   "Named:id=41:RecordSelector"
         _StyleDefs(91)  =   ":id=41,.parent=34"
         _StyleDefs(92)  =   "Named:id=42:FilterBar"
         _StyleDefs(93)  =   ":id=42,.parent=33"
      End
      Begin BiSAFramProject.BiSAFrame BiSAFrame4 
         Height          =   645
         Left            =   75
         Top             =   3390
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1138
         Caption         =   "SUB TOTAL"
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
         Begin BiSANumberBoxProject.BiSANumberBox nSubTotal 
            Height          =   315
            Left            =   60
            TabIndex        =   2
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
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
      End
      Begin BiSAFramProject.BiSAFrame BiSAFrame5 
         Height          =   645
         Left            =   1635
         Top             =   3390
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   1138
         Caption         =   "DISC"
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
         Begin BiSANumberBoxProject.BiSANumberBox nDiscount 
            Height          =   315
            Left            =   45
            TabIndex        =   3
            Top             =   240
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
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
      End
      Begin BiSAFramProject.BiSAFrame BiSAFrame7 
         Height          =   645
         Left            =   3090
         Top             =   3390
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1138
         Caption         =   "PPN"
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
         Begin BiSANumberBoxProject.BiSANumberBox nPajak 
            Height          =   315
            Left            =   45
            TabIndex        =   4
            Top             =   240
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
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
      End
      Begin BiSAFramProject.BiSAFrame BiSAFrame8 
         Height          =   645
         Left            =   4440
         Top             =   3390
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1138
         Caption         =   "TOTAL"
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
         Begin BiSANumberBoxProject.BiSANumberBox nTotal 
            Height          =   315
            Left            =   45
            TabIndex        =   5
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
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
      End
      Begin VB.Label Label4 
         Caption         =   "TUNAI"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   9735
         TabIndex        =   6
         Top             =   3420
         Width           =   570
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1935
      Left            =   0
      Top             =   0
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   3413
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
      Begin BiSATextBoxProject.BiSABrowse cAkunKas 
         Height          =   330
         Left            =   6915
         TabIndex        =   7
         Top             =   1365
         Width           =   3255
         _ExtentX        =   5741
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
      Begin BiSATextBoxProject.BiSABrowse cNamaSupplier 
         Height          =   330
         Left            =   3330
         TabIndex        =   8
         Top             =   705
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
      Begin BiSATextBoxProject.BiSATextBox cKota 
         Height          =   330
         Left            =   4095
         TabIndex        =   9
         Top             =   1020
         Width           =   1905
         _ExtentX        =   3360
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
         CaptionWidth    =   0
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
         TabIndex        =   10
         Top             =   1020
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
      Begin BiSATextBoxProject.BiSABrowse cSupplier 
         Height          =   330
         Left            =   75
         TabIndex        =   11
         Top             =   705
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
         Caption         =   "Supplier"
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
         TabIndex        =   12
         Top             =   390
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
      Begin BiSADateProject.BiSADate dPeriode 
         Height          =   330
         Index           =   0
         Left            =   6915
         TabIndex        =   13
         Top             =   420
         Width           =   2490
         _ExtentX        =   4392
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
         Caption         =   "Periode"
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
      Begin BiSANumberBoxProject.BiSANumberBox nPPn 
         Height          =   330
         Left            =   6915
         TabIndex        =   14
         Top             =   1050
         Width           =   2070
         _ExtentX        =   3651
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
         Caption         =   "PPn"
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
      Begin BiSANumberBoxProject.BiSANumberBox nPersDisc 
         Height          =   330
         Left            =   6915
         TabIndex        =   15
         Top             =   735
         Width           =   2070
         _ExtentX        =   3651
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
         Caption         =   "Discount"
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
      Begin BiSATextBoxProject.BiSATextBox cFakturAsli 
         Height          =   330
         Left            =   6915
         TabIndex        =   16
         Top             =   105
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   582
         Text            =   "12345678901234567890"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Verdana"
         BackColor       =   16777215
         MaxLength       =   20
         Appearance      =   0
         Caption         =   "Fak. Asli"
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
         TabIndex        =   17
         Top             =   1335
         Width           =   3750
         _ExtentX        =   6615
         _ExtentY        =   582
         Text            =   "12345678"
         BorderStyle     =   0
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
      Begin BiSADateProject.BiSADate dPeriode 
         Height          =   330
         Index           =   1
         Left            =   9390
         TabIndex        =   21
         Top             =   420
         Width           =   1380
         _ExtentX        =   2434
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
         CaptionWidth    =   0
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
         TabIndex        =   20
         Top             =   60
         Width           =   6030
      End
      Begin VB.Label Label3 
         Caption         =   "%"
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
         Left            =   9060
         TabIndex        =   19
         Top             =   1155
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "%"
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
         Left            =   9045
         TabIndex        =   18
         Top             =   810
         Width           =   240
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   600
      Left            =   0
      Top             =   6015
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   1058
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
         TabIndex        =   24
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
         Picture         =   "trConsignment2.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   3390
         TabIndex        =   25
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
         Picture         =   "trConsignment2.frx":028A
      End
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   1170
         TabIndex        =   26
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
         Picture         =   "trConsignment2.frx":0429
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   435
         Left            =   105
         TabIndex        =   27
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
         Picture         =   "trConsignment2.frx":0555
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   10635
         TabIndex        =   28
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
         Picture         =   "trConsignment2.frx":0700
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   9555
         TabIndex        =   29
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
         Picture         =   "trConsignment2.frx":07A6
      End
   End
End
Attribute VB_Name = "trConsignment2"
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
Dim cKode As String
Dim nSaldoStock As Double

Private Sub cAkunKas_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", "kodeakun,keterangan", "jenis", sisAssign, "D", , "kodeakun")
  If Not dbData.EOF Then
    cAkunKas.Text = cAkunKas.Browse(dbData)
  End If
End Sub

Private Sub Check1_Click()
Dim n As Integer

  If Check1.Value = 1 Then
    Check2.Value = 0
    For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
      vaArray(n, 0) = -1
    Next n
  End If
  tdbgrid1.ReBind
  tdbgrid1.Refresh
  SumTotal
End Sub

Private Sub Check2_Click()
Dim n As Integer

  If Check2.Value = 1 Then
    Check1.Value = 0
    For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
      vaArray(n, 0) = 0
    Next n
  End If
  tdbgrid1.ReBind
  tdbgrid1.Refresh
  SumTotal
End Sub

Private Sub cmdKeluar_Click()
  If lEdit Then
    InitValue
  Else
    Unload Me
  End If
End Sub

Private Sub SumTotal()
Dim n As Double
  
  nSubTotal.Value = 0
  For n = 0 To vaArray.UpperBound(1)
    If vaArray(n, 0) <> 0 Then
      nSubTotal.Value = nSubTotal.Value + vaArray(n, 8)
    End If
  Next
  
  If nPersDisc.Enabled = True Then
    nDiscount.Value = nPersDisc.Value / 100 * (nSubTotal.Value)
  End If
  
  nPajak.Value = (nPPn.Value / 100) * (nSubTotal.Value - (nDiscount.Value + nDiscount.Value))
  nTotal.Value = nSubTotal.Value + nPajak.Value - nDiscount.Value
  nTunai.Value = nTotal.Value
End Sub

Private Sub cmdSaveOK_Click()
  Simpan
End Sub

Private Sub Simpan()
End Sub

Private Sub DeleteInvoice()
End Sub

Private Function ValidSaving() As Boolean
  ValidSaving = True
  
  If cSupplier.Text = "" Then
    MsgBox "Kode Supplier tidak terisi" & vbCrLf & "Data tidak bisa disimpan", vbCritical
    ValidSaving = False
    Exit Function
  End If
  
  If cAkunKas.Text = "" Then
    MsgBox "Akun Kas tidak terisi" & vbCrLf & "Data tidak bisa disimpan", vbCritical
    ValidSaving = False
    Exit Function
  End If
End Function

Private Sub cmdSimpan_Click()
Dim lSave As Boolean
Dim Faktur As String
Dim n As Integer
lSave = True
  
  If ValidSaving Then
    If MsgBox("Data akan disimpan?", vbInformation + vbYesNo) = vbYes Then
      objData.Start GetDSN
      Faktur = cFaktur.Text
      If Not GetAvailable(cFaktur.Text, "totkonsinyasi", "nomorkonsinyasi") Then
        Faktur = GetNomor("totkonsinyasi", "nomorkonsinyasi", GetID, SisModulTransaksi.Konsinyasi)
      End If
      
      lSave = IIf(lSave, objData.Update(GetDSN, "totkonsinyasi", "nomorkonsinyasi = '" & Faktur & "'", Array("nomorkonsinyasi", "fakturasli", "tgl", "dtglawal", "dtglakhir", "kodesupplier", "ppn", "persdisc", "persdisc2", "subtotal", "pajak", "discount", "discount2", "total", "tunai", "hutang", "datetime", "username", "kodeakun", "kodecostcenter"), Array(Faktur, cFakturAsli.Text, Format(dTgl.Value, "yyyy-MM-dd"), Format(dPeriode(0).Value, "yyyy-MM-dd"), Format(dPeriode(1).Value, "yyyy-MM-dd"), cSupplier.Text, nPPn.Value, nPersDisc.Value, 0, nSubTotal.Value, nPajak.Value, nDiscount.Value, 0, nTotal.Value, nTunai.Value, 0, SNow, GetRegistry(reg_UserName), cAkunKas.Text, aCfg(objData, msCostCenterJualBeli))), False)
      lSave = IIf(lSave, objData.Delete(GetDSN, "konsinyasi", "nomorkonsinyasi", sisAssign, Faktur), False)
      For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
        If vaArray(n, 0) <> 0 Then
          lSave = IIf(lSave, objData.Add(GetDSN, "konsinyasi", Array("nomorkonsinyasi", "tgl", "kodestock", "qty", "harga", "kodesatuan", "discount", "jumlah", "nomorpenjualan"), Array(Faktur, Format(dTgl.Value, "yyyy-MM-dd"), vaArray(n, 2), vaArray(n, 4), vaArray(n, 6), vaArray(n, 5), vaArray(n, 7), vaArray(n, 8), vaArray(n, 9))), False)
          
          'Update harga beli dengan harga beli terakhir
          lSave = IIf(lSave, objData.Edit(GetDSN, "stock", "kodestock = '" & vaArray(n, 2) & "'", Array("hargabeli"), Array(vaArray(n, 8))), False)
          
          'Update status penjualan menjadi lunas
          lSave = IIf(lSave, objData.Edit(GetDSN, "penjualan", "nomorpenjualan = '" & vaArray(n, 9) & "'", Array("statuslunas"), Array(1)), False)
          
        End If
      Next n
      
      '==========
      'AKUNTANSI
      '==========
      
      'Hapus dulu di bukubesar
      lSave = IIf(lSave, DelKodeTr(objData, msKonsinyasi, Faktur), False)
      
      'Debet
      'Biaya pada akun biaya konsinyasi
      lSave = IIf(lSave, UpdKodeTr(objData, msKonsinyasi, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningKonsinyasi), aCfg(objData, msCostCenterJualBeli), "Konsinyasi an " & cNamaSupplier.Text, nSubTotal.Value, 0, "", SNow), False)
      'PPn
      lSave = IIf(lSave, UpdKodeTr(objData, msKonsinyasi, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningPPnkonsinyasi), aCfg(objData, msCostCenterJualBeli), "PPn konsinyasi an " & cNamaSupplier.Text, nPajak.Value, 0, "", SNow), False)
      
      'Kredit
      'Kas Bank
      lSave = IIf(lSave, UpdKodeTr(objData, msKonsinyasi, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), cAkunKas.Text, aCfg(objData, msCostCenterJualBeli), "Kas untuk konsinyasi an " & cNamaSupplier.Text, 0, nTunai.Value, "", SNow), False)
      'Discount seluruhnya
      lSave = IIf(lSave, UpdKodeTr(objData, msKonsinyasi, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningDiscountkonsinyasi), aCfg(objData, msCostCenterJualBeli), "Dsc Tot konsinyasi an " & cNamaSupplier.Text, 0, nDiscount.Value, "", SNow), False)
      
      If lSave Then
        objData.Save GetDSN
      Else
        objData.Cancel GetDSN
      End If
      InitValue
    End If
  End If
End Sub

Private Sub cSupplier_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "supplier", "kodesupplier,nama,alamat,kota", "kodesupplier", sisContent, cSupplier.Text, , "kodesupplier,nama")
  If Not dbData.EOF Then
    cSupplier.Text = cSupplier.Browse(dbData)
    cSupplier.Text = GetNull(dbData!kodesupplier)
    cNamaSupplier.Text = GetNull(dbData!nama, "")
    cAlamat.Text = GetNull(dbData!alamat, "")
    cKota.Text = GetNull(dbData!kota, "")
  End If
End Sub

Private Sub cnamasupplier_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "supplier", "kodesupplier,nama,alamat,kota", "nama", sisContent, cNamaSupplier.Text, , "kodesupplier,nama")
  If Not dbData.EOF Then
    cNamaSupplier.Text = cNamaSupplier.Browse(dbData)
    cSupplier.Text = GetNull(dbData!kodesupplier)
    cNamaSupplier.Text = GetNull(dbData!nama, "")
    cAlamat.Text = GetNull(dbData!alamat, "")
    cKota.Text = GetNull(dbData!kota, "")
  End If
End Sub

Private Sub dPeriode_Validate(Index As Integer, Cancel As Boolean)
  GetNewDataPenjualan
  SumTotal
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd, "SIKD"
  CenterForm Me, True
  InitValue
  
  Set dbData = objData.Browse(GetDSN, "costcenter", , "kodecostcenter", sisAssign, aCfg(objData, msCostCenterJualBeli))
  If Not dbData.EOF Then
    lbCostCenter.Caption = "Cost Centre : " & GetNull(dbData!Keterangan)
  End If
  
  TabIndex dTgl, n
  TabIndex cSupplier, n
  TabIndex cNamaSupplier, n
  TabIndex cAlamat, n
  TabIndex cFaktur, n
  
  TabIndex cFakturAsli, n
  TabIndex dPeriode(0), n
  TabIndex dPeriode(1), n
  TabIndex nPersDisc, n
  TabIndex nPPn, n
  TabIndex cAkunKas, n

  TabIndex nTunai, n
  
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
End Sub

Private Sub InitValue()
  cFaktur.Default
  cFaktur.Text = GetNomor("totkonsinyasi", "nomorkonsinyasi", GetID, SisModulTransaksi.Konsinyasi)
  dTgl.Value = Date
  dPeriode(0).Value = BOM(Date)
  dPeriode(1).Value = EOM(Date)
  nPersDisc.Value = 0
  nPPn.Value = 0
  cFakturAsli.Default
  cSupplier.Default
  cNamaSupplier.Default
  cAlamat.Default
  cKota.Default
  nSubTotal.Value = 0
  nPajak.Value = 0
  nDiscount.Value = 0
  nTotal.Value = 0
  nTunai.Value = 0
  cAkunKas.Text = cKasTeller
  
  vaArray.ReDim 0, -1, 0, 9
  Set tdbgrid1.Array = vaArray
  tdbgrid1.ReBind
End Sub

Private Sub GetNewDataPenjualan()
Dim cSQL As String
Dim n As Integer

  cSQL = ""
  cSQL = cSQL & " select p.kodestock,s.nama,p.qty,p.kodesatuan,p.hb,p.discount,p.nomorpenjualan from penjualan p"
  cSQL = cSQL & " left join stock s on s.kodestock = p.kodestock"
  cSQL = cSQL & " where s.kodesupplier = '" & cSupplier.Text & "'"
  cSQL = cSQL & " and p.tgl >= '" & Format(dPeriode(0).Value, "yyyy-MM-dd") & "' and p.tgl <= '" & Format(dPeriode(1).Value, "yyyy-MM-dd") & "'"
  cSQL = cSQL & " and p.statuslunas=0 "
  Set dbData = objData.Sql(GetDSN, cSQL)
  vaArray.ReDim 0, -1, 0, 9
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = -1
      vaArray(n, 1) = n + 1
      vaArray(n, 2) = GetNull(dbData!KodeStock)
      vaArray(n, 3) = GetNull(dbData!nama)
      vaArray(n, 4) = GetNull(dbData!qty)
      vaArray(n, 5) = GetNull(dbData!kodesatuan)
      vaArray(n, 6) = GetNull(dbData!hb)
      vaArray(n, 7) = GetNull(dbData!Discount)
      vaArray(n, 8) = vaArray(n, 4) * (vaArray(n, 6) - (vaArray(n, 6) * vaArray(n, 7)))
      vaArray(n, 9) = GetNull(dbData!nomorpenjualan)
      dbData.MoveNext
    Loop
  End If
  Set tdbgrid1.Array = vaArray
  tdbgrid1.ReBind
  tdbgrid1.Refresh
End Sub

Private Sub nPersDisc_Validate(Cancel As Boolean)
  SumTotal
End Sub

Private Sub nPPn_Validate(Cancel As Boolean)
  SumTotal
End Sub

Private Sub tdbgrid1_AfterColUpdate(ByVal ColIndex As Integer)
  tdbgrid1.Update
  SumTotal
End Sub

Private Sub tdbgrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
  If ColIndex <> 0 Then
    Cancel = True
  End If
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim n As Integer

  If KeyCode = vbKeyDelete Then
    If vaArray.UpperBound(1) >= 0 Then
      tdbgrid1.Delete
      SumTotal
      For n = 0 To vaArray.UpperBound(1)
        vaArray(n, 0) = n + 1
      Next
      tdbgrid1.ReBind
    End If
  End If
End Sub
