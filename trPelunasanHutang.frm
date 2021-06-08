VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trPelunasanHutang 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PELUNASAN HUTANG"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   12435
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   12435
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   4830
      Left            =   225
      Top             =   2340
      Width           =   12090
      _ExtentX        =   21325
      _ExtentY        =   8520
      Caption         =   "DATA HUTANG"
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      BackColor       =   -2147483633
      Begin SizerOneLibCtl.TabOne TabOne1 
         Height          =   4605
         Left            =   15
         TabIndex        =   18
         Top             =   210
         Width           =   11970
         _cx             =   21114
         _cy             =   8123
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   2
         MousePointer    =   0
         _ConvInfo       =   1
         Version         =   700
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FrontTabColor   =   -2147483633
         BackTabColor    =   -2147483633
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   "Hutang Belum Dibayar|Retur Beli (Titipan)"
         Align           =   0
         CurrTab         =   1
         FirstTab        =   0
         Style           =   3
         Position        =   0
         AutoSwitch      =   -1  'True
         AutoScroll      =   -1  'True
         TabPreview      =   -1  'True
         ShowFocusRect   =   -1  'True
         TabsPerPage     =   0
         BorderWidth     =   0
         BoldCurrent     =   -1  'True
         DogEars         =   -1  'True
         MultiRow        =   0   'False
         MultiRowOffset  =   200
         CaptionStyle    =   0
         TabHeight       =   0
         TabCaptionPos   =   4
         TabPicturePos   =   0
         CaptionEmpty    =   ""
         Separators      =   0   'False
         Begin VB.Frame Frame2 
            Height          =   4230
            Left            =   45
            TabIndex        =   21
            Top             =   330
            Width           =   11880
            Begin TrueOleDBGrid70.TDBGrid gridRetur 
               Height          =   3870
               Left            =   90
               TabIndex        =   22
               Top             =   210
               Width           =   11580
               _ExtentX        =   20426
               _ExtentY        =   6826
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
               Columns(4).Caption=   "HUTANG"
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
               Splits(0)._ColumnProps(11)=   "Column(2).Width=3281"
               Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
               Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=3201"
               Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=197124"
               Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
               Splits(0)._ColumnProps(16)=   "Column(3).Width=2619"
               Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
               Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2540"
               Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=197121"
               Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
               Splits(0)._ColumnProps(21)=   "Column(4).Width=3228"
               Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
               Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=3149"
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
               Splits(0)._ColumnProps(36)=   "Column(7).Width=3625"
               Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
               Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=3545"
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
               BorderStyle     =   0
               ColumnFooters   =   -1  'True
               DataMode        =   4
               DefColWidth     =   0
               HeadLines       =   1
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
         Begin VB.Frame Frame1 
            Height          =   4230
            Left            =   -12525
            TabIndex        =   19
            Top             =   330
            Width           =   11880
            Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
               Height          =   3855
               Left            =   105
               TabIndex        =   20
               Top             =   195
               Width           =   11580
               _ExtentX        =   20426
               _ExtentY        =   6800
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
               Columns(4).Caption=   "HUTANG"
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
               Splits(0)._ColumnProps(11)=   "Column(2).Width=3281"
               Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
               Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=3201"
               Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=197124"
               Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
               Splits(0)._ColumnProps(16)=   "Column(3).Width=2619"
               Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
               Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2540"
               Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=197121"
               Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
               Splits(0)._ColumnProps(21)=   "Column(4).Width=3228"
               Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
               Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=3149"
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
               Splits(0)._ColumnProps(36)=   "Column(7).Width=3625"
               Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
               Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=3545"
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
               BorderStyle     =   0
               ColumnFooters   =   -1  'True
               DataMode        =   4
               DefColWidth     =   0
               HeadLines       =   1
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
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   2190
      Left            =   210
      Top             =   120
      Width           =   6870
      _ExtentX        =   12118
      _ExtentY        =   3863
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
      BackColor       =   -2147483644
      Begin BiSATextBoxProject.BiSABrowse cSupplier 
         Height          =   330
         Left            =   75
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   765
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
         Enabled         =   0   'False
         Appearance      =   0
         Caption         =   "Kode Customer"
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
         TabIndex        =   1
         Top             =   450
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
         BackColor       =   -2147483644
         ForeColor       =   -2147483640
         Caption         =   "Tanggal"
         CaptionWidth    =   1300
         CaptionBackColor=   -2147483644
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
         TabIndex        =   2
         Top             =   765
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
         TabIndex        =   3
         Top             =   1080
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
         TabIndex        =   4
         Top             =   1395
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
      Begin BiSATextBoxProject.BiSABrowse cAkunKas 
         Height          =   330
         Left            =   75
         TabIndex        =   11
         Top             =   1740
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
      Begin BiSATextBoxProject.BiSABrowse cNamaAkun 
         Height          =   336
         Left            =   3312
         TabIndex        =   13
         Top             =   1740
         Width           =   3348
         _ExtentX        =   5900
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
         CaptionBackColor=   -2147483637
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
         Left            =   150
         TabIndex        =   12
         Top             =   135
         Width           =   6030
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame4 
      Height          =   630
      Left            =   225
      Top             =   7305
      Width           =   12015
      _ExtentX        =   21193
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
         Left            =   1185
         TabIndex        =   5
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
         Picture         =   "trPelunasanHutang.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   9000
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
         Picture         =   "trPelunasanHutang.frx":028A
      End
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   2340
         TabIndex        =   7
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
         Enabled         =   0   'False
         Picture         =   "trPelunasanHutang.frx":0429
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   435
         Left            =   105
         TabIndex        =   8
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
         Picture         =   "trPelunasanHutang.frx":0555
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   10530
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
         Picture         =   "trPelunasanHutang.frx":0700
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   9450
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
         Picture         =   "trPelunasanHutang.frx":07A6
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   2190
      Left            =   7110
      Top             =   120
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   3863
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
         Left            =   870
         TabIndex        =   14
         Top             =   840
         Width           =   3135
         _ExtentX        =   5530
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
         Left            =   870
         TabIndex        =   15
         Top             =   480
         Width           =   3135
         _ExtentX        =   5530
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
         Left            =   870
         TabIndex        =   16
         Top             =   1755
         Width           =   3150
         _ExtentX        =   5556
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
      Begin BiSANumberBoxProject.BiSANumberBox nRetur 
         Height          =   330
         Left            =   855
         TabIndex        =   17
         Top             =   1200
         Width           =   3165
         _ExtentX        =   5583
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
         Caption         =   "RETUR"
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
         X1              =   960
         X2              =   4335
         Y1              =   1665
         Y2              =   1665
      End
   End
End
Attribute VB_Name = "trPelunasanHutang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nPos As SisPos
Dim lClick As Boolean
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB
Dim vaRetur As New XArrayDB
Dim lEdit As Boolean

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

Private Sub cNama_Validate(Cancel As Boolean)
  cNama.Enabled = False
End Sub

Private Sub cSupplier_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "supplier", "kodesupplier,nama,alamat", "kodesupplier", sisContent, cSupplier.Text)
  If Not dbData.EOF Then
    cSupplier.Text = cSupplier.Browse(dbData)
    cSupplier.Text = GetNull(dbData!kodesupplier)
    cNama.Text = GetNull(dbData!nama)
    cAlamat.Text = GetNull(dbData!alamat)
    If nPos = Add Then
      GetData
    End If
  End If
End Sub

Private Sub GetData()
Dim n As Integer
Dim nSisaHutang As Double
Dim nTotalHutang As Double
Dim cCaption As String

  nTotalHutang = 0
  n = -1
  vaArray.ReDim 0, -1, 0, 7
  Set dbData = objData.Browse(GetDSN, "totpembelian", "nomorpembelian,tgl,hutang,jthtmp", "kodesupplier", sisAssign, cSupplier.Text, , "tgl desc")
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      If Not isLunas(objData, GetNull(dbData!nomorpembelian), nSisaHutang) Then
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        n = vaArray.UpperBound(1)
        vaArray(n, 0) = 0
        vaArray(n, 1) = n + 1
        vaArray(n, 2) = GetNull(dbData!nomorpembelian)
        vaArray(n, 3) = GetNull(dbData!tgl)
        isLunas objData, vaArray(n, 2), nSisaHutang
        vaArray(n, 4) = nSisaHutang 'GetNull(dbData!hutang)
        vaArray(n, 5) = GetNull(dbData!jthtmp)
        vaArray(n, 6) = 0
        vaArray(n, 7) = 0
        nTotalHutang = nTotalHutang + vaArray(n, 4)
        End If
      dbData.MoveNext
    Loop
  End If
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
  TDBGrid1.Refresh
  TDBGrid1.Columns(4).FooterText = Format(nTotalHutang, "###,###,###,##0.00")
  cCaption = "Hutang yg belum dibayar "
  TabOne1.TabCaption(0) = cCaption
  If n >= 0 Then
    TabOne1.TabCaption(0) = cCaption & IIf(n >= 0, "(" & n + 1 & ")", "")
  End If
End Sub

Private Sub GetDataReturLunas()
Dim nTotalRetur As Double
Dim n As Single

  nTotalRetur = 0
  n = -1
  vaRetur.ReDim 0, -1, 0, 7
  Set dbData = objData.Browse(GetDSN, "totrtnpembelian", "nomorreturpembelian,tgl,hutang,jthtmp", "kodesupplier", sisAssign, cSupplier.Text, " and flag_posting = 1 and fakturlunas = '" & cFaktur.Text & "'", "tgl desc")
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      'If Not isLunas(objData, GetNull(dbData!nomorpembelian), nSisaHutang) Then
        vaRetur.InsertRows vaRetur.UpperBound(1) + 1
        n = vaRetur.UpperBound(1)
        vaRetur(n, 0) = -1
        vaRetur(n, 1) = n + 1
        vaRetur(n, 2) = GetNull(dbData!nomorreturpembelian)
        vaRetur(n, 3) = GetNull(dbData!tgl)
        'isLunas objData, vaRetur(n, 2), nSisaHutang
        vaRetur(n, 4) = GetNull(dbData!hutang) 'GetNull(dbData!hutang)
        vaRetur(n, 5) = GetNull(dbData!jthtmp)
        vaRetur(n, 6) = 0
        vaRetur(n, 7) = 0
        nTotalRetur = nTotalRetur + vaRetur(n, 4)
      'End If
      dbData.MoveNext
    Loop
  End If
  Set gridRetur.Array = vaRetur
  gridRetur.ReBind
  gridRetur.Refresh
  gridRetur.Columns(4).FooterText = Format(nTotalRetur, "###,###,###,##0.00")
End Sub

Private Sub GetDataRetur()
Dim n As Integer
Dim nSisaHutang As Double
Dim nTotalRetur As Double
Dim cCaption As String

  nTotalRetur = 0
  n = -1
  vaRetur.ReDim 0, -1, 0, 7
  Set dbData = objData.Browse(GetDSN, "totrtnpembelian", "nomorreturpembelian,tgl,hutang,jthtmp", "kodesupplier", sisAssign, cSupplier.Text, " and flag_posting = 0 and hutang <> 0 and jenis_retur = '" & SisModelReturPembelian.vTitip & "'", "tgl desc")
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      'If Not isLunas(objData, GetNull(dbData!nomorpembelian), nSisaHutang) Then
        vaRetur.InsertRows vaRetur.UpperBound(1) + 1
        n = vaRetur.UpperBound(1)
        vaRetur(n, 0) = 0
        vaRetur(n, 1) = n + 1
        vaRetur(n, 2) = GetNull(dbData!nomorreturpembelian)
        vaRetur(n, 3) = GetNull(dbData!tgl)
        'isLunas objData, vaRetur(n, 2), nSisaHutang
        vaRetur(n, 4) = GetNull(dbData!hutang) 'GetNull(dbData!hutang)
        vaRetur(n, 5) = GetNull(dbData!jthtmp)
        vaRetur(n, 6) = 0
        vaRetur(n, 7) = 0
        nTotalRetur = nTotalRetur + vaRetur(n, 4)
      'End If
      dbData.MoveNext
    Loop
  End If
  Set gridRetur.Array = vaRetur
  gridRetur.ReBind
  gridRetur.Refresh
  gridRetur.Columns(4).FooterText = Format(nTotalRetur, "###,###,###,##0.00")
  
  cCaption = "Retur Beli "
  TabOne1.TabCaption(1) = cCaption
  If n >= 0 Then
    TabOne1.TabCaption(1) = cCaption & IIf(n >= 0, "(" & n + 1 & ")", "")
  End If
End Sub

Private Function isLunas(ByVal obj As CodeSuiteLibrary.Data, ByVal nomorpembelian As String, ByRef Sisahutang As Double) As Boolean
Dim db As New ADODB.Recordset
Dim hutang As Double
Dim Lunas As Double

  isLunas = True
  hutang = 0
  Lunas = 0
  Sisahutang = 0
  Set db = obj.Browse(GetDSN, "totpembelian", "hutang", "nomorpembelian", sisAssign, nomorpembelian)
  If Not db.EOF Then
    hutang = GetNull(db!hutang)
  End If
  Set db = obj.Browse(GetDSN, "pelunasanhutang", "sum(discount+pelunasan) as totallunas", "nomorpembelian", sisAssign, nomorpembelian)
  If Not db.EOF Then
    Lunas = GetNull(db!totalLunas)
  End If
  If hutang - Lunas > 0 Then
    Sisahutang = hutang - Lunas
    isLunas = False
  End If
End Function

Private Sub cFaktur_ButtonClick()
Dim n As Integer
Dim lSave As Boolean
Dim nTotalHutang As Double
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
  
  lSave = True

  nTotalHutang = 0
  vaArray.ReDim 0, -1, 0, 7
  Set dbData = objData.Browse(GetDSN, "totpelunasanhutang", "nomorpelunasanhutang,pelunasan", "tgl", sisAssign, Format(dTanggal.value, "yyyy-MM-dd"), " and kodesupplier = '" & cSupplier.Text & "'", "nomorpelunasanhutang desc")
  If Not dbData.EOF Then
    cFaktur.Text = cFaktur.Browse(dbData)
    Set dbData = objData.Browse(GetDSN, "totpelunasanhutang", , "nomorpelunasanhutang", sisAssign, cFaktur.Text)
    If Not dbData.EOF Then
      nTotal.value = GetNull(dbData!Total)
      nDiscount.value = GetNull(dbData!Discount)
      nLunas.value = GetNull(dbData!Pelunasan)
      cAkunKas.Text = GetNull(dbData!kodeakun)
    End If
    
    Set dbData = objData.Browse(GetDSN, "pelunasanhutang p", "p.nomorpembelian,t.tgl,t.jthtmp,p.hutang,p.discount,p.pelunasan", "nomorpelunasanhutang", sisAssign, cFaktur.Text, , , Array("LEFT JOIN totpembelian t ON t.nomorpembelian = p.nomorpembelian"))
    If Not dbData.EOF Then
      Do While Not dbData.EOF
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        n = vaArray.UpperBound(1)
        vaArray(n, 0) = -1
        vaArray(n, 1) = n
        vaArray(n, 2) = GetNull(dbData!nomorpembelian)
        vaArray(n, 3) = GetNull(dbData!tgl)
        vaArray(n, 4) = GetNull(dbData!hutang)
        vaArray(n, 5) = GetNull(dbData!jthtmp)
        vaArray(n, 6) = GetNull(dbData!Discount)
        vaArray(n, 7) = GetNull(dbData!Pelunasan)
        nTotalHutang = nTotalHutang + vaArray(n, 4)
        dbData.MoveNext
      Loop
      SumTDB
    End If
    Set TDBGrid1.Array = vaArray
    TDBGrid1.ReBind
    TDBGrid1.Refresh
    TDBGrid1.Columns(4).FooterText = Format(nTotalHutang, "###,###,###,##0.00")
    
    GetDataReturLunas
    
    Me.Refresh
    If nPos = Delete Then
      Me.Refresh
      If MsgBox("Yakin data akan dihapus?", vbYesNo) = vbYes Then
        objData.Start GetDSN
        lSave = IIf(lSave, objData.Delete(GetDSN, "totpelunasanhutang", "nomorpelunasanhutang", sisAssign, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "pelunasanhutang", "nomorpelunasanhutang", sisAssign, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "kartuhutang", "nomorkartuHutang", sisAssign, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "bukubesar", "faktur", sisAssign, cFaktur.Text), False)
          Dim i
          For i = vaArray.LowerBound(1) To vaArray.UpperBound(1)
            'Update flaglunas
            If lCekStatusLunasHutang(objData, vaArray(i, 2)) = True Then
              lSave = IIf(lSave, objData.Edit(GetDSN, "totpembelian", "nomorpembelian = '" & vaArray(i, 2) & "'", Array("flaglunas"), Array(1)), False)
            Else
              lSave = IIf(lSave, objData.Edit(GetDSN, "totpembelian", "nomorpembelian = '" & vaArray(i, 2) & "'", Array("flaglunas"), Array(0)), False)
            End If
          Next i
          
          For i = vaRetur.LowerBound(1) To vaRetur.UpperBound(1)
            'Kembalikan status flag_lunas = 0
            lSave = IIf(lSave, objData.Edit(GetDSN, "totrtnpembelian", "nomorreturpembelian = '" & vaRetur(i, 2) & "'", Array("flag_posting", "fakturlunas"), Array(0, "")), False)
          Next i
          
        If lSave Then
          objData.Save GetDSN
        Else
          objData.Cancel GetDSN
        End If
      End If
      GetEdit False
      initvalue
    End If
  End If
End Sub

Private Sub cmdAdd_Click()
  nPos = Add
  GetEdit True
  initvalue
  dTanggal.SetFocus
  GetBrowseFaktur False
  cFaktur.Text = CreateNomorFaktur(objData, sisModulTransaksi.pelunasanhutang, "totpelunasanhutang", "nomorpelunasanhutang")
End Sub

Private Sub cmdAktivasi_Click()
  frmAktivasi.Action Me
End Sub

Private Sub cmdEdit_Click()
  nPos = Edit
  GetEdit True
  initvalue
  dTanggal.SetFocus
  GetBrowseFaktur True
End Sub

Private Sub cmdHapus_Click()
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
  dTanggal.value = Date
  cSupplier.Default
  cNama.Default
  cAlamat.Default
  nDiscount.Default
  nTotal.Default
  nLunas.Default
  nRetur.Default
  ClearTdbgrid
  TDBGrid1.Columns(7).FooterText = ""
  TDBGrid1.Columns(4).FooterText = ""
  cNama.Enabled = True
  cAkunKas.Text = cKasTeller
  
  cAkunKas.Enabled = True
  cAkunKas.BackColor = vbWhite
  If GetRegistry(reg_UserLevel) <> 0 Then
    cAkunKas.Enabled = False
    cAkunKas.BackColor = vbButtonFace
  End If

  Set dbData = objData.Browse(GetDSN, "akun", "kodeakun,keterangan", "kodeakun", sisAssign, cAkunKas.Text)
  If Not dbData.EOF Then
    cNamaAkun.Text = GetNull(dbData!keterangan)
  Else
    cNamaAkun.Default
  End If
End Sub

Private Sub ClearTdbgrid()
  vaArray.ReDim 0, -1, 0, 7
  TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
  
  vaRetur.ReDim 0, -1, 0, 7
  gridRetur.Array = vaArray
  gridRetur.ReBind
  TabOne1 = 0
  TabOne1.TabCaption(0) = "Hutang Belum Dibayar"
  TabOne1.TabCaption(1) = "Retur Pembelian"
  
End Sub

Private Function ValidSaving() As Boolean
  ValidSaving = True
  If Not CheckData(cFaktur.Text, "Nomor Faktur harus diisi...?") Then
    ValidSaving = False
    cFaktur.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cSupplier.Text, "Kode Supplier harus diisi...?") Then
    ValidSaving = False
    cSupplier.SetFocus
    Exit Function
  End If
End Function

Private Function validOK() As Boolean
  validOK = True
End Function

Private Sub DeleteData()
End Sub

Private Sub cmdSimpan_Click()
Dim n As Single
Dim lSave As Boolean
Dim Faktur As String

  
  lSave = True
  objData.Start GetDSN
  Faktur = cFaktur.Text
  
  'cek apakah ada data yg akan dilunasi
  If Trim(Faktur) = "" Then
    MsgBox "Maaf Nomor Faktur Kosong/Tidak Valid" & vbCrLf & "Data tidak bisa disimpan"
    Exit Sub
  End If
  
  If GetCekCentang = False Then
    MsgBox "Maaf tidak ada data untuk di proses", vbExclamation
    Exit Sub
  End If
  
  If nLunas.value < 0 Then
    MsgBox "Maaf nilai pelunasan tidak ada atau minus. Proses tidak bisa dilanjutkan", vbExclamation
    Exit Sub
  End If
  
  'Simpan di tabel totpembelian dan pembelian
  lSave = IIf(lSave, objData.Delete(GetDSN, "totpelunasanhutang", "nomorpelunasanhutang", sisAssign, Faktur), False)
  lSave = IIf(lSave, objData.Delete(GetDSN, "pelunasanhutang", "nomorpelunasanhutang", sisAssign, Faktur), False)
  lSave = IIf(lSave, objData.Delete(GetDSN, "kartuhutang", "nomorkartuhutang", sisAssign, Faktur), False)
  
  lSave = IIf(lSave, objData.Update(GetDSN, "totpelunasanhutang", "nomorpelunasanhutang = '" & Faktur & "'", Array("nomorpelunasanhutang", "kodesupplier", "tgl", "discount", "total", "pelunasan", "datetime", "username", "kodeakun", "kodecostcenter"), Array(Faktur, cSupplier.Text, Format(dTanggal.value, "yyyy-MM-dd"), nDiscount.value, nTotal.value, nLunas.value, SNow, GetRegistry(reg_Username), cAkunKas.Text, GetCostCenterUser(objData, GetRegistry(reg_Username)))), False)
  For n = 0 To vaArray.UpperBound(1)
    If vaArray(n, 0) = -1 Then
      lSave = IIf(lSave, objData.Add(GetDSN, "pelunasanhutang", Array("nomorpelunasanhutang", "nomorpembelian", "hutang", "discount", "pelunasan"), Array(Faktur, vaArray(n, 2), vaArray(n, 4), vaArray(n, 6), vaArray(n, 7))), False)
      'Update flaglunas
      If lCekStatusLunasHutang(objData, vaArray(n, 2)) = True Then
        lSave = IIf(lSave, objData.Edit(GetDSN, "totpembelian", "nomorpembelian = '" & vaArray(n, 2) & "'", Array("flaglunas"), Array(1)), False)
      Else
        lSave = IIf(lSave, objData.Edit(GetDSN, "totpembelian", "nomorpembelian = '" & vaArray(n, 2) & "'", Array("flaglunas"), Array(0)), False)
      End If
    End If
  Next n
  
  'update faktur vaRetur set flag = 1 jika dicentang
    For n = 0 To vaRetur.UpperBound(1)
    If vaRetur(n, 0) = -1 Then
      lSave = IIf(lSave, objData.Edit(GetDSN, "totrtnpembelian", "nomorreturpembelian = '" & vaRetur(n, 2) & "'", Array("flag_posting", "fakturlunas"), Array(1, Faktur)), False)
    End If
  Next n
  
  lSave = IIf(lSave, UpdKartuHutang(objData, SisKartuHutang.SisPelunasanHutang, Faktur, dTanggal.value, cSupplier.Text, "Pelunasan hutang an " & cNama.Text, nTotal.value - nDiscount.value), False)
  lSave = IIf(lSave, UpdKartuHutang(objData, SisKartuHutang.SisDiscountPelunasanHutang, Faktur, dTanggal.value, cSupplier.Text, "Discount Pelunasan hutang an " & cNama.Text, nDiscount.value), False)
  'lSave = IIf(lSave, UpdKartuHutang(objData, SisKartuHutang.SisReturPembelian, Faktur, dTanggal.Value, cSupplier.Text, "Retur Pelunasan hutang an " & cNama.Text, nRetur.Value), False)

  '==========
  '  Jurnal +
  '==========
  ' Hutang
  '   Kas
  '   Diskon
  '   Retur
  '==========
  '==========
  '  Jurnal -
  '==========
  ' Kas
  ' Hutang
  '   Diskon
  '   Retur
  '==========
  
  
  lSave = IIf(lSave, DelKodeTr(objData, msPelunasanHutang, Faktur), False)
  If nLunas.value > 0 Then
    lSave = IIf(lSave, UpdKodeTr(objData, msPelunasanHutang, Faktur, Format(dTanggal.value, "yyyy-MM-dd"), cAkunKas.Text, GetCostCenterUser(objData, GetRegistry(reg_Username)), "Pelunasan Hutang an " & cNama.Text, 0, nLunas.value), False)
  Else
    lSave = IIf(lSave, UpdKodeTr(objData, msPelunasanHutang, Faktur, Format(dTanggal.value, "yyyy-MM-dd"), cAkunKas.Text, GetCostCenterUser(objData, GetRegistry(reg_Username)), "Pelunasan Hutang an " & cNama.Text, nLunas.value * -1, 0), False)
  End If
  lSave = IIf(lSave, UpdKodeTr(objData, msPelunasanHutang, Faktur, Format(dTanggal.value, "yyyy-MM-dd"), GetAkunSupplier(objData, cSupplier.Text), GetCostCenterUser(objData, GetRegistry(reg_Username)), "Pelunasan Hutang an " & cNama.Text, nTotal.value, 0), False)
    lSave = IIf(lSave, UpdKodeTr(objData, msPelunasanHutang, Faktur, Format(dTanggal.value, "yyyy-MM-dd"), aCfg(objData, msRekeningPotonganHutang), GetCostCenterUser(objData, GetRegistry(reg_Username)), "Pelunasan Hutang an " & cNama.Text, 0, nDiscount.value), False)
    lSave = IIf(lSave, UpdKodeTr(objData, msPelunasanHutang, Faktur, Format(dTanggal.value, "yyyy-MM-dd"), aCfg(objData, msRekeningReturPembelian), GetCostCenterUser(objData, GetRegistry(reg_Username)), "Pelunasan Hutang an " & cNama.Text, 0, nRetur.value), False)

  If lSave Then
    objData.Save GetDSN
  Else
    objData.Cancel GetDSN
  End If
  
  GetEdit False
  initvalue
  
End Sub

Private Function GetCekCentang() As Boolean
Dim n As Integer

  GetCekCentang = False
  For n = 0 To vaArray.UpperBound(1)
    If vaArray(n, 0) = -1 Then
      GetCekCentang = True
      Exit For
    End If
  Next n
  
  For n = 0 To vaRetur.UpperBound(1)
    If vaRetur(n, 0) = -1 Then
      GetCekCentang = True
      Exit For
    End If
  Next n
End Function

Private Sub cNama_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "supplier", "kodesupplier,nama,alamat", "nama", sisContent, cNama.Text)
  If Not dbData.EOF Then
    cNama.Text = cNama.Browse(dbData)
    cSupplier.Text = GetNull(dbData!kodesupplier)
    cNama.Text = GetNull(dbData!nama)
    cAlamat.Text = GetNull(dbData!alamat)
    If nPos = Add Then
      GetData
      GetDataRetur
    End If
  End If
End Sub

Private Sub dTanggal_Validate(Cancel As Boolean)
  If Not IsInPeriod(dTanggal.value) Or (dTanggal.value > Date) Then
    Cancel = True
    dTanggal.SetFocus
    GetEdit False
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd, "SIKD"
  initvalue
  GetEdit False
  CenterForm Me
  
  Set dbData = objData.Browse(GetDSN, "costcenter", , "kodecostcenter", sisAssign, GetCostCenterUser(objData, GetRegistry(reg_Username)))
  If Not dbData.EOF Then
    lbCostCenter.Caption = "Cost Centre : " & GetNull(dbData!keterangan)
  End If
  
  TabIndex dTanggal, n
  TabIndex cSupplier, n
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
Dim nTotalHutang As Double

  
  nTotalHutang = 0
  For n = 0 To vaArray.UpperBound(1)
    Discount = Discount + vaArray(n, 6)
    Pelunasan = Pelunasan + vaArray(n, 7)
    nTotalHutang = nTotalHutang + vaArray(n, 4)
  Next n
  nTotal.value = Pelunasan + Discount
  nDiscount.value = Discount
  nLunas.value = nTotal.value - nDiscount.value - nRetur.value
  TDBGrid1.Columns(4).FooterText = Format(nTotalHutang, "###,###,###,##0.00")
End Sub

Private Sub SumTDBRetur()
Dim n As Integer
'Dim Discount As Double
Dim Pelunasan As Double
Dim nTotalRetur As Double

  
  nTotalRetur = 0
  For n = 0 To vaRetur.UpperBound(1)
    'Discount = Discount + vaRetur(n, 6)
    Pelunasan = Pelunasan + vaRetur(n, 7)
    nTotalRetur = nTotalRetur + vaRetur(n, 4)
  Next n
  
  nRetur.value = Pelunasan '+ Discount
  'nDiscount.Value = Discount
  nLunas.value = nTotal.value - nDiscount.value - nRetur.value
  gridRetur.Columns(4).FooterText = Format(nTotalRetur, "###,###,###,##0.00")
End Sub

Private Sub gridRetur_AfterColUpdate(ByVal ColIndex As Integer)
  gridRetur.Update
  SumTDBRetur
End Sub

Private Sub gridRetur_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim nSisaPiutang As Double
  
  'isLunas objData, gridRetur.Columns(2).Value, nSisaPiutang
  
  If Not IsNumeric(gridRetur.Columns(6).value) Or Not IsNumeric(gridRetur.Columns(7).value) Or gridRetur.Columns(6).value < 0 Or gridRetur.Columns(7).value < 0 Then
    Cancel = True
    Exit Sub
  End If
  
  If ColIndex = 0 Or ColIndex = 6 Or ColIndex = 7 Then
    If gridRetur.Columns(0).value = -1 Then
      If ColIndex <> 7 Then
        gridRetur.Columns(7).value = gridRetur.Columns(4).value - gridRetur.Columns(6).value
      Else
        gridRetur.Columns(6).value = 0
      End If
    Else
      gridRetur.Columns(7).value = 0
      gridRetur.Columns(6).value = 0
    End If
  Else
    Cancel = True
  End If
End Sub

Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
  TDBGrid1.Update
  SumTDB
End Sub

Private Sub tdbgrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim nSisaPiutang As Double
  
  isLunas objData, TDBGrid1.Columns(2).value, nSisaPiutang
  
  If Not IsNumeric(TDBGrid1.Columns(6).value) Or Not IsNumeric(TDBGrid1.Columns(7).value) Or TDBGrid1.Columns(6).value < 0 Or TDBGrid1.Columns(7).value < 0 Then
    Cancel = True
    Exit Sub
  End If
  
  If ColIndex = 0 Or ColIndex = 6 Or ColIndex = 7 Then
    If TDBGrid1.Columns(0).value = -1 Then
      If ColIndex <> 7 Then
        TDBGrid1.Columns(7).value = TDBGrid1.Columns(4).value - TDBGrid1.Columns(6).value
      Else
        TDBGrid1.Columns(6).value = 0
      End If
    Else
      TDBGrid1.Columns(7).value = 0
      TDBGrid1.Columns(6).value = 0
    End If
  Else
    Cancel = True
  End If
End Sub
