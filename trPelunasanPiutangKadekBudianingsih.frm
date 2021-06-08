VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trPelunasanPiutangKadekBudianingsih 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pelunasan Piutang BC KADEK BUDIANINGSIH"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13950
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   13950
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   4215
      Left            =   15
      Top             =   2520
      Width           =   13905
      _ExtentX        =   24527
      _ExtentY        =   7435
      Caption         =   "DATA PIUTANG"
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   3915
         Left            =   60
         TabIndex        =   0
         Top             =   255
         Width           =   13740
         _ExtentX        =   24236
         _ExtentY        =   6906
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   4
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "No"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "FAKTUR"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "TGL"
         Columns(3).DataField=   ""
         Columns(3).NumberFormat=   "dd-MM-yyyy"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "PIUTANG"
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "###,###,###,###,##0.00"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "JATUH TEMPO"
         Columns(5).DataField=   ""
         Columns(5).NumberFormat=   "dd-MM-yyyy"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "DISC RP"
         Columns(6).DataField=   ""
         Columns(6).NumberFormat=   "###,###,###,###,##0.00"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "PELUNASAN"
         Columns(7).DataField=   ""
         Columns(7).NumberFormat=   "###,###,###,###,##0.00"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   8
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=609"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=529"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=197124"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=1005"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=926"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=197124"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=4551"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=4471"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=197124"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=3016"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2937"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=197121"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=4022"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=3942"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=197122"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=3096"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=3016"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=197121"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=2302"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2223"
         Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=197122"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(36)=   "Column(7).Width=4604"
         Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=4524"
         Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=197122"
         Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
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
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=70,.parent=13,.alignment=2"
         _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=67,.parent=14"
         _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=68,.parent=15"
         _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=69,.parent=17"
         _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=58,.parent=13,.alignment=1"
         _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
         _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
         _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=74,.parent=13,.alignment=2"
         _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=71,.parent=14"
         _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=72,.parent=15"
         _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=73,.parent=17"
         _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=14"
         _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=15"
         _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=17"
         _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=54,.parent=13,.alignment=1,.bgcolor=&HFFFF80&"
         _StyleDefs(66)  =   ":id=54,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(67)  =   ":id=54,.fontname=Tahoma"
         _StyleDefs(68)  =   "Splits(0).Columns(7).HeadingStyle:id=51,.parent=14"
         _StyleDefs(69)  =   "Splits(0).Columns(7).FooterStyle:id=52,.parent=15"
         _StyleDefs(70)  =   "Splits(0).Columns(7).EditorStyle:id=53,.parent=17"
         _StyleDefs(71)  =   "Named:id=33:Normal"
         _StyleDefs(72)  =   ":id=33,.parent=0"
         _StyleDefs(73)  =   "Named:id=34:Heading"
         _StyleDefs(74)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(75)  =   ":id=34,.wraptext=-1"
         _StyleDefs(76)  =   "Named:id=35:Footing"
         _StyleDefs(77)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(78)  =   "Named:id=36:Selected"
         _StyleDefs(79)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(80)  =   "Named:id=37:Caption"
         _StyleDefs(81)  =   ":id=37,.parent=34,.alignment=2,.bold=0,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(82)  =   ":id=37,.strikethrough=0,.charset=0"
         _StyleDefs(83)  =   ":id=37,.fontname=Tahoma"
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
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   2565
      Left            =   15
      Top             =   0
      Width           =   13905
      _ExtentX        =   24527
      _ExtentY        =   4524
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
      Begin BiSANumberBoxProject.BiSANumberBox nSaldoPiutang 
         Height          =   330
         Left            =   11520
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   420
         Width           =   2310
         _ExtentX        =   4075
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
      Begin BiSATextBoxProject.BiSABrowse cAkunKas 
         Height          =   330
         Left            =   75
         TabIndex        =   2
         Top             =   2130
         Visible         =   0   'False
         Width           =   3225
         _ExtentX        =   5689
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
      Begin BiSAFramProject.BiSAFrame BiSAFrame3 
         Height          =   1290
         Left            =   10260
         Top             =   1065
         Width           =   3570
         _ExtentX        =   6297
         _ExtentY        =   2275
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
         Begin BiSANumberBoxProject.BiSANumberBox nDiscount 
            Height          =   330
            Left            =   45
            TabIndex        =   3
            Top             =   60
            Width           =   3420
            _ExtentX        =   6033
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
            Caption         =   "DISCOUNT"
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
         Begin BiSANumberBoxProject.BiSANumberBox nTotal 
            Height          =   330
            Left            =   45
            TabIndex        =   4
            Top             =   375
            Width           =   3420
            _ExtentX        =   6033
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
            Caption         =   "TOTAL"
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
         Begin BiSANumberBoxProject.BiSANumberBox nLunas 
            Height          =   330
            Left            =   45
            TabIndex        =   5
            Top             =   840
            Width           =   3420
            _ExtentX        =   6033
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
            Caption         =   "LUNAS"
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
         Begin VB.Line Line1 
            BorderWidth     =   2
            X1              =   135
            X2              =   3510
            Y1              =   795
            Y2              =   795
         End
      End
      Begin BiSATextBoxProject.BiSABrowse cCustomer 
         Height          =   330
         Left            =   75
         TabIndex        =   6
         Top             =   720
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
      Begin BiSADateProject.BiSADate dTanggal 
         Height          =   330
         Left            =   75
         TabIndex        =   7
         Top             =   405
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
         Left            =   3330
         TabIndex        =   8
         Top             =   720
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
         Left            =   75
         TabIndex        =   9
         Top             =   1035
         Width           =   6570
         _ExtentX        =   11589
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
         Left            =   75
         TabIndex        =   10
         Top             =   1770
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
      Begin BiSATextBoxProject.BiSABrowse cNamaDepartmen 
         Height          =   330
         Left            =   1500
         TabIndex        =   11
         Top             =   1350
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
      Begin BiSANumberBoxProject.BiSANumberBox nSaldoTopUpMember 
         Height          =   330
         Left            =   11520
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   75
         Width           =   2310
         _ExtentX        =   4075
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
         TabIndex        =   15
         Top             =   120
         Width           =   6030
      End
      Begin VB.Label Label1 
         Caption         =   "Saldo Piutang Terikini (update)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   7470
         TabIndex        =   14
         Top             =   495
         Width           =   4005
      End
      Begin VB.Label Label2 
         Caption         =   "Saldo Top Up : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   10275
         TabIndex        =   13
         Top             =   135
         Width           =   1170
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame4 
      Height          =   630
      Left            =   0
      Top             =   6720
      Width           =   13905
      _ExtentX        =   24527
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
      Begin BiSAButtonProject.BiSAButton cmdPerbaikan 
         Height          =   435
         Left            =   4320
         TabIndex        =   16
         Top             =   105
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   767
         Caption         =   "X"
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
      Begin BiSAButtonProject.BiSAButton cmdPrint 
         Height          =   435
         Left            =   3840
         TabIndex        =   17
         Top             =   105
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   767
         Caption         =   "P"
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
      Begin BiSAButtonProject.BiSAButton cmdHapus 
         Height          =   435
         Left            =   2220
         TabIndex        =   18
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
         Picture         =   "trPelunasanPiutangKadekBudianingsih.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   3390
         TabIndex        =   19
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
         Picture         =   "trPelunasanPiutangKadekBudianingsih.frx":028A
      End
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   1170
         TabIndex        =   20
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
         Picture         =   "trPelunasanPiutangKadekBudianingsih.frx":0429
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   435
         Left            =   105
         TabIndex        =   21
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
         Picture         =   "trPelunasanPiutangKadekBudianingsih.frx":0555
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   12735
         TabIndex        =   22
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
         Picture         =   "trPelunasanPiutangKadekBudianingsih.frx":0700
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   11655
         TabIndex        =   23
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
         Picture         =   "trPelunasanPiutangKadekBudianingsih.frx":07A6
      End
   End
End
Attribute VB_Name = "trPelunasanPiutangKadekBudianingsih"
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
Public lPubStatus As Boolean
Public vaPubReff As New XArrayDB
Public nPubTotal As Double
Public cPubAkun As String
Public lClose As Double
Public nTarikTunai As Double
Public nWithDraw As Double
Public nSisaKurangTopUp As Double
Public lTarikTunai As Boolean
Public nSaldoTopUp As Double
Public nTunai As Double
Public nKembalian As Double
Public nMetodePembayaran As Integer

Sub GetEdit(lPar As Boolean)
  lEdit = lPar
  BiSAFrame1.Enabled = lPar
  BiSAFrame2.Enabled = lPar
  SetButton cmdSimpan, cmdKeluar, cmdAdd, cmdEdit, cmdHapus, nPos, lPar, cmdAktivasi
End Sub

Private Sub cAkunKas_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", "kodeakun,keterangan", "jenis", sisAssign, "D", " and left(kodeakun,1)='1'")
  If Not dbData.EOF Then
    cAkunKas.Text = cAkunKas.Browse(dbData, Array("Kode Akun", "Keterangan"), , Array(25, 30))
  End If
End Sub

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

Private Sub GetMark()
Dim n As Double
  
  n = TDBGrid1.Bookmark
  If n >= 0 Then
    vaArray(n, 0) = Not vaArray(n, 0)
    TDBGrid1.Columns(0) = vaArray(n, 0)
  End If
End Sub

Private Sub GetData()
Dim n As Integer
Dim nSisaPiutang As Double
Dim nTmpSisaPiutang As Double

  vaArray.ReDim 0, -1, 0, 7
  nTmpSisaPiutang = 0
  nSaldoPiutang.Value = 0
  Set dbData = objData.Browse(GetDSN, "totpenjualan", "nomorpenjualan,tgl,piutang,jthtmp", "kodeanggota", sisAssign, cCustomer.Text, , "tgl desc")
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      If Not isLunas(objData, GetNull(dbData!nomorpenjualan), nSisaPiutang) Then
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        n = vaArray.UpperBound(1)
        vaArray(n, 0) = 0
        vaArray(n, 1) = n + 1
        vaArray(n, 2) = GetNull(dbData!nomorpenjualan)
        vaArray(n, 3) = GetNull(dbData!Tgl)
        isLunas objData, vaArray(n, 2), nSisaPiutang
        vaArray(n, 4) = nSisaPiutang 'GetNull(dbData!Piutang)
        nTmpSisaPiutang = nTmpSisaPiutang + vaArray(n, 4)
        vaArray(n, 5) = GetNull(dbData!jthtmp)
        vaArray(n, 6) = 0
        'awalny nilai ini diset 0
        'kebijakan baru, tidak ada lagi yg boleh melunasi hutang separo separo, vaArray(n,7) diset = vaArray(n,4)
        vaArray(n, 7) = vaArray(n, 4)
      End If
      dbData.MoveNext
    Loop
  End If
  Set TDBGrid1.Array = vaArray
  TDBGrid1.Columns(4).FooterText = Format(nTmpSisaPiutang, "###,###,###,##0.00")
  TDBGrid1.ReBind
  TDBGrid1.Refresh
  
  'cari jumlah/saldo piutang terakhir setelah dipotong retur
  Set dbData = objData.Browse(GetDSN, "kartupiutang", "sum(debet-kredit) as saldopiutang", "kodeanggota", sisAssign, cCustomer.Text, " and tgl <='" & Format(Date, "yyyy-MM-dd") & "'")
  If Not dbData.EOF Then
    nSaldoPiutang.Value = GetNull(dbData!saldopiutang)
  End If
  
  'cari saldo top up member
  Set dbData = objData.Browse(GetDSN, "membertopup m", "m.kodeanggota,a.nama,a.alamat,sum(debet) as debet,sum(kredit) as kredit,sum(m.debet-m.kredit) as saldo", "m.kodeanggota", sisAssign, cCustomer.Text, " GROUP BY m.kodeanggota", , Array("left join anggota a on a.kodeanggota = m.kodeanggota"))
  If Not dbData.EOF Then
    nSaldoTopUpMember.Value = GetNull(dbData!saldo)
  Else
    nSaldoTopUpMember.Value = 0
  End If

End Sub


Private Sub cFaktur_ButtonClick()
Dim n As Integer
Dim lSave As Boolean
lSave = True

  vaArray.ReDim 0, -1, 0, 7
  Set dbData = objData.Browse(GetDSN, "totpelunasanpiutang", "nomorpelunasanpiutang,pelunasan", "tgl", sisAssign, Format(dTanggal.Value, "yyyy-MM-dd"), " and kodeanggota = '" & cCustomer.Text & "'", "nomorpelunasanpiutang")
  If Not dbData.EOF Then
    cFaktur.Text = cFaktur.Browse(dbData)
    Set dbData = objData.Browse(GetDSN, "totpelunasanpiutang", , "nomorpelunasanpiutang", sisAssign, cFaktur.Text)
    If Not dbData.EOF Then
      nTotal.Value = GetNull(dbData!Total)
      nDiscount.Value = GetNull(dbData!Discount)
      nLunas.Value = GetNull(dbData!Pelunasan)
      cAkunKas.Text = GetNull(dbData!kodeakun)
    End If
    Set dbData = objData.Browse(GetDSN, "pelunasanpiutang p", "p.nomorpenjualan,t.tgl,t.jthtmp,p.piutang,p.discount,p.pelunasan", "nomorpelunasanpiutang", sisAssign, cFaktur.Text, , , Array("LEFT JOIN totpenjualan t ON t.nomorpenjualan = p.nomorpenjualan"))
    If Not dbData.EOF Then
      Do While Not dbData.EOF
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        n = vaArray.UpperBound(1)
        vaArray(n, 0) = -1
        vaArray(n, 1) = n
        vaArray(n, 2) = GetNull(dbData!nomorpenjualan)
        vaArray(n, 3) = GetNull(dbData!Tgl)
        vaArray(n, 4) = GetNull(dbData!Piutang)
        vaArray(n, 5) = GetNull(dbData!jthtmp)
        vaArray(n, 6) = GetNull(dbData!Discount)
        vaArray(n, 7) = GetNull(dbData!Pelunasan)
        dbData.MoveNext
      Loop
      SumTDB
    End If
    Set TDBGrid1.Array = vaArray
    TDBGrid1.ReBind
    TDBGrid1.Refresh
    Me.Refresh
    If nPos = Delete Then
      Me.Refresh
      If MsgBox("Yakin data akan dihapus?", vbYesNo) = vbYes Then
        objData.Start GetDSN
        lStart = True
        'CEK
        'tidak boleh dilakukan penghapusan atau pengkoreksian apabila faktur bg yg bersangkutang sudah dilunasi
        Set dbData = objData.Browse(GetDSN, "pencairanbg", , "nomorpelunasanpiutang", sisAssign, cFaktur.Text)
        If Not dbData.EOF Then
          MsgBox "Maaf transaksi ini tidak bisa dikoreksi kembali. BG/Cek sudah dicairkan"
          Exit Sub
        End If
        
        lSave = IIf(lSave, DelKodeTr(objData, msPelunasanPiutang, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "pelunasanpiutang", "nomorpelunasanpiutang", sisAssign, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "kartupiutang", "nomorkartupiutang", sisAssign, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "totpelunasanpiutang", "nomorpelunasanpiutang", sisAssign, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "bukubesar", "faktur", sisAssign, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "bg", "nomorpelunasanpiutang", sisAssign, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "membertopup", "nomormembertopup ", sisAssign, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "totmembertopup", "nomormembertopup ", sisAssign, cFaktur.Text), False)
      
        For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
          If lCekStatusLunas(objData, vaArray(n, 2)) = True Then
            lSave = IIf(lSave, objData.Edit(GetDSN, "penjualan", "nomorpenjualan = '" & vaArray(n, 2) & "'", Array("statuslunas"), Array(1)), False)
            lSave = IIf(lSave, objData.Edit(GetDSN, "totpenjualan", "nomorpenjualan = '" & vaArray(n, 2) & "'", Array("flaglunas"), Array(1)), False)
          Else
            lSave = IIf(lSave, objData.Edit(GetDSN, "penjualan", "nomorpenjualan = '" & vaArray(n, 2) & "'", Array("statuslunas"), Array(0)), False)
            lSave = IIf(lSave, objData.Edit(GetDSN, "totpenjualan", "nomorpenjualan = '" & vaArray(n, 2) & "'", Array("flaglunas"), Array(0)), False)
          End If
        Next n
        
        If lSave Then
          objData.Save GetDSN
        Else
          objData.Cancel GetDSN
        End If
      End If
      GetEdit False
      initvalue
    End If
    If nPos = Edit Then
      SendKeysA vbKeyReturn, True
    End If
  End If
End Sub

Private Sub cmdAdd_Click()
  nPos = Add
  GetEdit True
  initvalue
  dTanggal.SetFocus
  GetBrowseFaktur False
  cFaktur.Text = GetNomor("totpelunasanpiutang", "nomorpelunasanpiutang", GetID, SisModulTransaksi.pelunasanpiutang)
End Sub

Private Sub cmdAktivasi_Click()
  frmAktivasi.Action Me
End Sub

Private Sub cmdEdit_Click()
  If GetRegistry(reg_UserLevel) <> 0 Then
    If objMenu.GetPassword("", Me, GetDSN) Then
      If objMenu.UserLevel <> 0 Then
        MsgBox "Maaf, Anda tidak diberikan wewenang untuk melakukan PENGEDITAN." & vbCrLf & _
               "Hanya user dengan LEVEL 0 (SUPERVISOR) yg diijinkan", vbInformation, "OTORISASI not ALLOWED"
        Exit Sub
      End If
    Else
      Unload Me
      GetEdit False
      Exit Sub
    End If
  End If
  
  nPos = Edit
  GetEdit True
  initvalue
  dTanggal.SetFocus
  GetBrowseFaktur True
End Sub

Private Sub cmdHapus_Click()
  If GetRegistry(reg_UserLevel) <> 0 Then
    If objMenu.GetPassword("", Me, GetDSN) Then
      If objMenu.UserLevel <> 0 Then
        MsgBox "Maaf, Anda tidak diberikan wewenang untuk melakukan PENGEDITAN." & vbCrLf & _
               "Hanya user dengan LEVEL 0 (SUPERVISOR) yg diijinkan", vbInformation, "OTORISASI not ALLOWED"
        Exit Sub
      End If
    Else
      Unload Me
      GetEdit False
      Exit Sub
    End If
  End If
  nPos = Delete
  GetEdit True
  initvalue
  dTanggal.SetFocus
  GetBrowseFaktur True
End Sub

Private Sub GetBrowseFaktur(ByVal lStat As Boolean)
  cFaktur.Button = lStat
  cFaktur.Enabled = lStat
End Sub

Private Sub cmdKeluar_Click()
  If lEdit Then
    initvalue
    GetEdit False
  Else
    Unload Me
  End If
End Sub

Private Sub initvalue()
  cFaktur.Default
  dTanggal.Value = Date
  cCustomer.Default
  cNama.Default
  cAlamat.Default
  cAkunKas.Text = cKasTeller
  nDiscount.Default
  nTotal.Default
  nLunas.Default
  nSaldoPiutang.Default
  cNamaDepartmen.Default
  nSaldoTopUpMember.Default
  
  ClearTdbgrid
  TDBGrid1.Columns(7).FooterText = ""
  TDBGrid1.Columns(4).FooterText = ""
  vaPubReff.ReDim 0, 2, 0, 2
  Label1.Caption = "Saldo Piutang Terikini (update) per tgl " & Format(Date, "dd-MM-yyyy")
  lClose = True
End Sub

Private Sub ClearTdbgrid()
  vaArray.ReDim 0, -1, 0, 7
  TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
End Sub

Private Function ValidSaving() As Boolean
  ValidSaving = True
  If Not CheckData(cFaktur.Text, "Nomor Faktur harus diisi...?") Then
    ValidSaving = False
    cFaktur.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cCustomer.Text, "Kode Customer harus diisi...?") Then
    ValidSaving = False
    cCustomer.SetFocus
    Exit Function
  End If
End Function

Private Function validOK() As Boolean
  validOK = True
End Function

Private Sub cmdPerbaikan_Click()
Dim a As New exportExcel
Dim cSQL As String
Dim n As Single
Dim lSave As Boolean

cSQL = " select t.nomorpelunasanpiutang,tt.nomorpenjualan,t.kodeanggota,a.nama,t.tgl,tt.total as totalan from totpelunasanpiutang t"
cSQL = cSQL & " LEFT JOIN anggota a on a.kodeanggota = t.kodeanggota"
cSQL = cSQL & " LEFT JOIN pelunasanpiutang p on p.nomorpelunasanpiutang = t.nomorpelunasanpiutang"
cSQL = cSQL & " LEFT JOIN totpenjualan tt on tt.kodeanggota = t.kodeanggota"
cSQL = cSQL & " Where t.tgl >= '2012-11-01' And p.Pelunasan Is Null And tt.flaglunas <> 1"
cSQL = cSQL & " ORDER BY t.nomorpelunasanpiutang"

vaArray.ReDim 0, -1, 0, 4
Set dbData = objData.Sql(GetDSN, cSQL)
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!nomorpelunasanpiutang)
      vaArray(n, 1) = GetNull(dbData!nomorpenjualan)
      vaArray(n, 2) = GetNull(dbData!kodeanggota)
      vaArray(n, 3) = GetNull(dbData!nama)
      vaArray(n, 4) = GetNull(dbData!totalan)
      dbData.MoveNext
    Loop
    
    If MsgBox("UPS, ada yg salah dalam proses pelunasan, apakah akan dilihat?", vbYesNo + vbCritical) = vbYes Then
      a.RecordSource = vaArray
      a.ExportToExcel
    End If
    
    lSave = True
    objData.Start GetDSN

    If MsgBox("Apakah akan diperbaki?", vbYesNo + vbInformation) = vbYes Then
      For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
        lSave = IIf(lSave, objData.Add(GetDSN, "pelunasanpiutang", _
        Array("nomorpelunasanpiutang", "nomorpenjualan", "piutang", "discount", "pelunasan"), _
        Array(vaArray(n, 0), vaArray(n, 1), vaArray(n, 4), 0, vaArray(n, 4))), False)
        
        lSave = IIf(lSave, objData.Edit(GetDSN, "penjualan", "nomorpenjualan = '" & vaArray(n, 1) & "'", Array("statuslunas"), Array(1)), False)
        lSave = IIf(lSave, objData.Edit(GetDSN, "totpenjualan", "nomorpenjualan = '" & vaArray(n, 1) & "'", Array("flaglunas"), Array(1)), False)
      Next n
      
      If lSave Then
        objData.Save GetDSN
        MsgBox "Ok, data sudah selesai diperbaiki", vbInformation
      Else
        objData.Cancel GetDSN
        MsgBox "Maaf terjadi kesalahan dalam proses perbaikan, data tidak jadi diperbaiki", vbInformation
      End If
    
    End If
    
  Else
    MsgBox "Hore.. tidak ditemukan satupun yg salah dalam proses pelunasan" & vbCrLf & "SISTEM OK", vbInformation
  End If
  
End Sub

Private Sub cmdPrint_Click()
  trPrint2.noOrder = TDBGrid1.Columns(2).Text
  Set dbData = objData.Browse(GetDSN, "totpenjualan t", "t.*,a.nama,a.telp", "t.nomorpenjualan", sisAssign, TDBGrid1.Columns(2).Text, , , Array("left join anggota a on a.kodeanggota = t.kodeanggota"))
  If Not dbData.EOF Then
    trPrint2.nSubTotal = GetNull(dbData!Subtotal)
    trPrint2.nDiscount = GetNull(dbData!dp)
    trPrint2.nCash = GetNull(dbData!Tunai)
    trPrint2.nChange = GetNull(dbData!Piutang)
    trPrint2.cKodeMember = GetNull(dbData!kodeanggota)
    trPrint2.cMember = GetNull(dbData!nama)
    trPrint2.cTeleponMember = GetNull(dbData!telp)
    trPrint2.Ups = GetNull(dbData!upkepada)
    trPrint2.dTgNota = Format(GetNull(dbData!Tgl), "dd/MM/yyyy")
    trPrint2.dJthTempoNota = Format(GetNull(dbData!jthtmp), "dd/MM/yyyy")
    Load trPrint2
    trPrint2.Show vbModal
  End If
End Sub

Private Sub cmdSimpan_Click()
Dim n As Single
Dim lSave As Boolean
Dim Faktur As String

  lSave = True
  objData.Start GetDSN
  lStart = True
  Faktur = cFaktur.Text
  If nPos = Add Then
    If Not GetAvailable(cFaktur.Text, "totpelunasanpiutang", "nomorpelunasanpiutang") Then
      Faktur = GetNomor("totpelunasanpiutang", "nomorpelunasanpiutang", GetID, SisModulTransaksi.pelunasanpiutang)
    End If
  End If
  
' please load form pelunasan

  trLunasPiutang.nTotalYangHarusDibayar.Value = nLunas.Value
  trLunasPiutang.nTunai.Value = nLunas.Value
  trLunasPiutang.Label3.Caption = cFaktur.Text
  trLunasPiutang.cKodeAnggota.Text = cCustomer.Text
  trLunasPiutang.cNamaAnggota.Text = cNama.Text
  
  If nSaldoTopUpMember.Value > 0 Then
    trLunasPiutang.opt(2).Value = True
  Else
    trLunasPiutang.opt(0).Value = True
  End If
    
  Load trLunasPiutang
  trLunasPiutang.Show vbModal
  
  If lClose = True Then
    objData.Cancel GetDSN
    Exit Sub
  End If
  
  'simpan di tabel totpenjualan dan penjualan
  
  lSave = IIf(lSave, objData.Delete(GetDSN, "totpelunasanpiutang", "nomorpelunasanpiutang", sisAssign, Faktur), False)
  lSave = IIf(lSave, objData.Delete(GetDSN, "pelunasanpiutang", "nomorpelunasanpiutang", sisAssign, Faktur), False)
  lSave = IIf(lSave, objData.Delete(GetDSN, "kartupiutang", "nomorkartupiutang", sisAssign, Faktur), False)
  lSave = IIf(lSave, objData.Delete(GetDSN, "membertopup", "nomormembertopup ", sisAssign, Faktur), False)
  
  lSave = IIf(lSave, objData.Update(GetDSN, "totpelunasanpiutang", "nomorpelunasanpiutang = '" & Faktur & "'", Array("nomorpelunasanpiutang", "kodeanggota", "tgl", "discount", "total", "pelunasan", "datetime", "username", "kodeakun", "kodecostcenter"), Array(Faktur, cCustomer.Text, Format(dTanggal.Value, "yyyy-MM-dd"), nDiscount.Value, nTotal.Value, nLunas.Value, SNow, GetRegistry(reg_UserName), cPubAkun, aCfg(objData, msCostCenterJualBeli))), False)
  For n = 0 To vaArray.UpperBound(1)
    If vaArray(n, 0) = -1 Then
      lSave = IIf(lSave, objData.Add(GetDSN, "pelunasanpiutang", Array("nomorpelunasanpiutang", "nomorpenjualan", "piutang", "discount", "pelunasan"), Array(Faktur, vaArray(n, 2), vaArray(n, 4), vaArray(n, 6), vaArray(n, 7))), False)
      'Set status lunas pada table penjualan
      'Cek dulu apakah faktur penjualan ini sudah dilunasi apa belum
      If lCekStatusLunas(objData, vaArray(n, 2)) = True Then
        lSave = IIf(lSave, objData.Edit(GetDSN, "penjualan", "nomorpenjualan = '" & vaArray(n, 2) & "'", Array("statuslunas"), Array(1)), False)
        lSave = IIf(lSave, objData.Edit(GetDSN, "totpenjualan", "nomorpenjualan = '" & vaArray(n, 2) & "'", Array("flaglunas"), Array(1)), False)
      Else
        lSave = IIf(lSave, objData.Edit(GetDSN, "penjualan", "nomorpenjualan = '" & vaArray(n, 2) & "'", Array("statuslunas"), Array(0)), False)
        lSave = IIf(lSave, objData.Edit(GetDSN, "totpenjualan", "nomorpenjualan = '" & vaArray(n, 2) & "'", Array("flaglunas"), Array(0)), False)
      End If
    End If
    
    If nPos <> Add Then
      If vaArray(n, 0) <> -1 Then
        lSave = IIf(lSave, objData.Edit(GetDSN, "penjualan", "nomorpenjualan = '" & vaArray(n, 2) & "'", Array("statuslunas"), Array(0)), False)
      End If
    End If
  Next n
  
  lSave = IIf(lSave, UpdKartuHutang(objData, SisKartuHutang.SisPelunasanPiutang, Faktur, dTanggal.Value, cCustomer.Text, "Pelunasan Piutang an " & cNama.Text, nTotal.Value - nDiscount.Value), False)
  lSave = IIf(lSave, UpdKartuHutang(objData, SisKartuHutang.SisDiscountPelunasanPiutang, Faktur, dTanggal.Value, cCustomer.Text, "Discount Pelunasan Piutang an " & cNama.Text, nDiscount.Value), False)
  
  'CEK
  'tidak boleh dilakukan penghapusan atau pengkoreksian apabila faktur bg yg bersangkutang sudah dilunasi
  Set dbData = objData.Browse(GetDSN, "pencairanbg", , "nomorpelunasanpiutang", sisAssign, Faktur)
  If Not dbData.EOF Then
    MsgBox "Maaf transaksi ini tidak bisa dikoreksi kembali. BG/Cek sudah dicairkan"
    objData.Cancel GetDSN
    Exit Sub
  End If
  
  'Simpan di table BG
  'hapus dulu record yg pernah ada
      
  Dim i As Integer
  lSave = IIf(lSave, objData.Delete(GetDSN, "bg", "nomorpelunasanpiutang", sisAssign, Faktur), False)
  For i = vaPubReff.LowerBound(1) To vaPubReff.UpperBound(1)
    If vaPubReff(i, 1) <> 0 Then
      lSave = IIf(lSave, objData.Add(GetDSN, "bg", Array("nomorpelunasanpiutang", "reff", "jumlah", "jatuhtempo"), Array(Faktur, vaPubReff(i, 0), vaPubReff(i, 1), vaPubReff(i, 2))), False)
    End If
  Next i
  
  'Jurnal
  'Kas
  'diskon
  '   Piutang
  
  lSave = IIf(lSave, DelKodeTr(objData, msPelunasanPiutang, Faktur), False)
  
  If trLunasPiutang.opt(0).Value = True Then
    'Jika pembayaran tunai maka lawannya Kas
    lSave = IIf(lSave, UpdKodeTr(objData, msPelunasanPiutang, Faktur, Format(dTanggal.Value, "yyyy-MM-dd"), cPubAkun, aCfg(objData, msCostCenterJualBeli), "Pelunasan piutang an " & cNama.Text, nLunas.Value, 0), False)
  ElseIf trLunasPiutang.opt(1).Value = True Then
    'Jika pembayaran BG maka lawannya akun BG
    lSave = IIf(lSave, UpdKodeTr(objData, msPelunasanPiutang, Faktur, Format(dTanggal.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningBG), aCfg(objData, msCostCenterJualBeli), "Pelunasan piutang an " & cNama.Text, nLunas.Value, 0), False)
  End If
  
  lSave = IIf(lSave, UpdKodeTr(objData, msPelunasanPiutang, Faktur, Format(dTanggal.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningPotonganPiutang), aCfg(objData, msCostCenterJualBeli), "Pelunasan piutang an " & cNama.Text, nDiscount.Value, 0), False)
      lSave = IIf(lSave, UpdKodeTr(objData, msPelunasanPiutang, Faktur, Format(dTanggal.Value, "yyyy-MM-dd"), GetAkunMember(objData, cCustomer.Text), aCfg(objData, msCostCenterJualBeli), "Pelunasan piutang an " & cNama.Text, 0, nTotal.Value), False)
  
  If trLunasPiutang.opt(2).Value = True Then
    'pelunasan dengan menggunakan topup
    'simpan di buku besar
    'simpan di tabel membertopup
    
    Dim vaField
    Dim vaValue
    
    
    vaField = Array("nomormembertopup", "tgl", "kodeanggota", "keterangan", "kredit")
    If lTarikTunai = True Then
      'bagi dua pembyaran
      'vaValue = Array(Faktur, dTanggal.Value, cCustomer.Text, "Tarik Tunai an " & cNama.Text, nTarikTunai + nWithDraw)
      vaValue = Array(Faktur, dTanggal.Value, cCustomer.Text, "Tarik Tunai Dana Top Up an " & cNama.Text, nTarikTunai)
      lSave = IIf(lSave, objData.Add(GetDSN, "membertopup", vaField, vaValue), False)
      vaValue = Array(Faktur, dTanggal.Value, cCustomer.Text, "Pelunasan Piutang via Top Up an " & cNama.Text, nWithDraw)
      lSave = IIf(lSave, objData.Add(GetDSN, "membertopup", vaField, vaValue), False)
      
      lSave = IIf(lSave, UpdKodeTr(objData, msMemberTopUp, Faktur, Format(dTanggal.Value, "yyyy-MM-dd"), GetAkunKas(objData, GetRegistry(reg_UserName)), "", "Tarik tunai an " & cNama.Text, 0, nTarikTunai), False)
      lSave = IIf(lSave, UpdKodeTr(objData, msMemberTopUp, Faktur, Format(dTanggal.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningTopUp), "", "Tarik tunai an " & cNama.Text, nTarikTunai, 0), False)

    ElseIf lTarikTunai = False Then
      vaValue = Array(Faktur, dTanggal.Value, cCustomer.Text, "Tarik Tunai an " & cNama.Text, nWithDraw)
      lSave = IIf(lSave, objData.Add(GetDSN, "membertopup", vaField, vaValue), False)
    End If
    
    lSave = IIf(lSave, UpdKodeTr(objData, msMemberTopUp, Faktur, Format(dTanggal.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningTopUp), "", "Pelunasan via Top Up an " & cNama.Text, nWithDraw, 0), False)
    lSave = IIf(lSave, UpdKodeTr(objData, msMemberTopUp, Faktur, Format(dTanggal.Value, "yyyy-MM-dd"), GetAkunKas(objData, GetRegistry(reg_UserName)), "", "Bayar Sisa Pelunasan " & cNama.Text, nSisaKurangTopUp, 0), False)
    
  End If
  
  
  
  If lSave Then
    objData.Save GetDSN
  Else
    objData.Cancel GetDSN
  End If
    
  'Lakukan pengecekan apakah semua data disimpan sesuai dengan tabel yg dituju
  'langkah awal:
  
  Set dbData = objData.Browse(GetDSN, "totpelunasanpiutang", , "nomorpelunasanpiutang", sisAssign, Faktur)
  If Not dbData.EOF Then
    'cek apakah jumlah nya sudah sesuai dengan yg ada di tabel pelunasan piutang
    
  End If
  
  For i = 1 To aCfg(objData, msJumlahCetakanPenjualanNonTunai)
    trPrintPelunasanPiutang.noOrder = Faktur
    Set dbData = objData.Browse(GetDSN, "totpelunasanpiutang t", "t.*,a.*", "t.nomorpelunasanpiutang", sisAssign, Faktur, , , Array("left join anggota a on a.kodeanggota = t.kodeanggota"))
    If Not dbData.EOF Then
      trPrintPelunasanPiutang.nSubTotal = GetNull(dbData!Total)
      trPrintPelunasanPiutang.nDiscount = 0
      trPrintPelunasanPiutang.nCash = 0
      trPrintPelunasanPiutang.nChange = 0
      trPrintPelunasanPiutang.cKodeMember = GetNull(dbData!kodeanggota)
      trPrintPelunasanPiutang.cMember = GetNull(dbData!nama)
      trPrintPelunasanPiutang.cTeleponMember = GetNull(dbData!telp)
      trPrintPelunasanPiutang.Ups = 0
      
      trPrintPelunasanPiutang.nKembali1 = nTarikTunai 'berapa uang yg ditarik
      trPrintPelunasanPiutang.nSaldoTopUp = nSaldoTopUp 'saldo top up
      trPrintPelunasanPiutang.nSisa = nSisaKurangTopUp 'kurang
      trPrintPelunasanPiutang.nTunai = nTunai
      trPrintPelunasanPiutang.nKembali2 = nKembalian
      trPrintPelunasanPiutang.lKembali = lTarikTunai
      trPrintPelunasanPiutang.nMetodePembayaran = nMetodePembayaran
      
      Load trPrintPelunasanPiutang
      trPrintPelunasanPiutang.Show vbModal
      
    End If
  Next i
  PrintStruk Faktur
  
  Unload trLunasPiutang
  
  GetEdit False
  initvalue
  
End Sub

Private Sub PrintStruk(ByVal Faktur As String)
Dim n As Double
Dim nBruto As Double
Dim nTotQty As Double

  With aMainmenu.IO1
    .Open GetRegistry(reg_PortPrinterKasir), ""
    .WriteString Chr(27) & Chr(15) & vbCrLf
    .WriteString Padc(Trim("STRUK PELUNASAN"), 40) & vbCrLf
    .WriteString Padc(Trim(aCfg(objData, msNamaPerusahaan)), 40) & vbCrLf
    .WriteString Padc(Trim(aCfg(objData, msAlamatPerusahaan)), 40) & vbCrLf
    .WriteString Padc(Trim(aCfg(objData, msTelepon) & " " & aCfg(objData, msFax)), 40) & vbCrLf
    .WriteString Padc(aCfg(objData, msKota), 40) & vbCrLf
    .WriteString Replicate("-", 40) & vbCrLf
    .WriteString "Member. " & cCustomer.Text & " " & cNama.Text & vbCrLf
    .WriteString Replicate("-", 40) & vbCrLf
    .WriteString "Print By: " & aCfg(objData, msNama) & vbCrLf

'    trPrintPelunasanPiutang.nSubTotal = GetNull(dbData!Total)
'    trPrintPelunasanPiutang.nDiscount = 0
'    trPrintPelunasanPiutang.nCash = 0
'    trPrintPelunasanPiutang.nChange = 0
'    trPrintPelunasanPiutang.cKodeMember = GetNull(dbData!kodeanggota)
'    trPrintPelunasanPiutang.cMember = GetNull(dbData!nama)
'    trPrintPelunasanPiutang.cTeleponMember = GetNull(dbData!telp)
'    trPrintPelunasanPiutang.Ups = 0
'
'    trPrintPelunasanPiutang.nKembali1 = nTarikTunai 'berapa uang yg ditarik
'    trPrintPelunasanPiutang.nSaldoTopUp = nSaldoTopUp 'saldo top up
'    trPrintPelunasanPiutang.nSisa = nSisaKurangTopUp 'kurang
'    trPrintPelunasanPiutang.nTunai = nTunai
'    trPrintPelunasanPiutang.nKembali2 = nKembalian
'    trPrintPelunasanPiutang.lKembali = lTarikTunai
'    trPrintPelunasanPiutang.nMetodePembayaran = nMetodePembayaran
    
    .WriteString Padl("Tarik        : " & Padl(Format(nTarikTunai, "###,###,##0"), 11), 40) & vbCrLf
    .WriteString Padl("S.Top Up     : " & Padl(Format(nSaldoTopUp, "###,###,##0"), 11), 40) & vbCrLf
    .WriteString Padl("Sisa Kurang  : " & Padl(Format(nSisaKurangTopUp, "###,###,##0"), 11), 40) & vbCrLf
    .WriteString Padl("Tunai        : " & Padl(Format(nTunai, "###,###,##0"), 11), 40) & vbCrLf
    .WriteString Padl("Kembalian    : " & Padl(Format(nKembalian, "###,###,##0"), 11), 40) & vbCrLf

'    For n = 0 To vaArray.UpperBound(1)
'      If vaArray(n, 3) <> 0 Then
'        vaArray(n, 5) = vaArray(n, 5) - (vaArray(n, 5) * vaArray(n, 6) / 100)
'        nBruto = nBruto + (vaArray(n, 3) * vaArray(n, 5))
'        nTotQty = nTotQty + vaArray(n, 3)
'        .WriteString Left(vaArray(n, 2), 40) & vbCrLf
'        .WriteString Padl(Left(vaArray(n, 1), 10) & Padl(Format(vaArray(n, 3), "#,##0.00"), 8) & " x" & Padl(Format(vaArray(n, 5), "#,###,##0"), 9) & " =" & Padl(Format(vaArray(n, 3) * vaArray(n, 5), "#,###,##0"), 9), 40) & vbCrLf
'      End If
'    Next
'    .WriteString Replicate("-", 40) & vbCrLf
'    .WriteString Padr("=> " & Format(nTotQty, "###,###,##0.00") & " Items", 20) & Padl("Sub   : " & Padl(Format(nBruto, "###,###,##0"), 11), 20) & vbCrLf
'
'    .WriteString Padl("Disc  : " & Padl(Format(nDiscount.Value, "###,###,##0"), 11), 40) & vbCrLf
'    .WriteString Padl("Total : " & Padl(Format(nTotal.Value, "###,###,##0"), 11), 40) & vbCrLf
'    .WriteString Padl("DP    : " & Padl(Format(nDP.Value, "###,###,##0"), 11), 40) & vbCrLf
'    .WriteString Padl("Bayar Tunai .... " & Padl(Format(nTunai.Value, "###,###,##0"), 11), 40) & vbCrLf
'    .WriteString Padl("Hutang .... " & Padl(Format(nPiutang.Value, "###,###,##0"), 11), 40) & vbCrLf
'    .WriteString Replicate("-", 40) & vbCrLf
'
'    .WriteString "No. " & Faktur & Padl(Format(Now, "dd-MM-yyyy HH:MM:SS"), 22) & vbCrLf
'
'    .WriteString "Print by " & Padl(GetRegistry(reg_UserName), 26) & vbCrLf
'    .WriteString Replicate("-", 40) & vbCrLf
    
    
    .Close
    OpenDrawer GetRegistry(reg_PortPrinterKasir)
  End With
End Sub


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

Private Sub dTanggal_Validate(Cancel As Boolean)
  If Not IsInPeriod(dTanggal.Value) Or (dTanggal.Value > Date) Then
    Cancel = True
    dTanggal.SetFocus
    GetEdit False
  End If
End Sub

Private Sub Form_Activate()
  If nPos = Add Then
    GetData
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd, "SIKD"
  initvalue
  GetEdit False
  CenterForm Me
  
  Set dbData = objData.Browse(GetDSN, "costcenter", , "kodecostcenter", sisAssign, aCfg(objData, msCostCenterJualBeli))
  If Not dbData.EOF Then
    lbCostCenter.Caption = "Cost Centre : " & GetNull(dbData!keterangan)
  End If
  
  TabIndex dTanggal, n
  TabIndex cCustomer, n
  TabIndex cNama, n
  TabIndex cFaktur, n
  TabIndex cAkunKas, n
  
  TabIndex cmdAdd, n
  TabIndex cmdEdit, n
  TabIndex cmdHapus, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
End Sub

Private Sub SumTDB()
Dim n As Integer
Dim Discount As Double
Dim Pelunasan As Double
  
  For n = 0 To vaArray.UpperBound(1)
    If vaArray(n, 0) = -1 Then
      Discount = Discount + vaArray(n, 6)
      Pelunasan = Pelunasan + vaArray(n, 7)
    End If
  Next n
  
  nTotal.Value = Pelunasan + Discount
  nDiscount.Value = Discount
  nLunas.Value = nTotal.Value - nDiscount.Value
  
End Sub

Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
  TDBGrid1.Update
  SumTDB
'  MsgBox TDBGrid1.Columns(0).Value & " - " & TDBGrid1.Columns(1).Value
End Sub

Private Sub tdbgrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim nSisaPiutang As Double
  
  isLunas objData, TDBGrid1.Columns(2).Value, nSisaPiutang
  
  If Not IsNumeric(TDBGrid1.Columns(6).Value) Or Not IsNumeric(TDBGrid1.Columns(7).Value) Or TDBGrid1.Columns(6).Value < 0 Then
    Cancel = True
    Exit Sub
  End If

  If ColIndex = 0 Or ColIndex = 6 Or ColIndex = 7 Then
    If TDBGrid1.Columns(0).Value = -1 Then
      If ColIndex <> 7 Then
        TDBGrid1.Columns(7).Value = TDBGrid1.Columns(4).Value - TDBGrid1.Columns(6).Value
      Else
        TDBGrid1.Columns(6).Value = 0
      End If
'    Else
'      TDBGrid1.Columns(7).Value = 0
'      TDBGrid1.Columns(6).Value = 0
    End If
  Else
    Cancel = True
  End If
  
  'pelunasan piutang tidak boleh lebih dari sisa piutang
  If TDBGrid1.Columns(7).Value > TDBGrid1.Columns(4).Value Then
    MsgBox "Maaf, nilai pelunasan tidak boleh melebihi dari sisa piutang faktur" & vbCrLf & "Silahkan ulangi pengisian. Terimakasih"
    TDBGrid1.Refresh
    Cancel = True
  End If
End Sub

'Private Sub TDBGrid1_DblClick()
'  GetCetakFakturpenjualan objData, TDBGrid1.Columns(2).Text, False
'End Sub


