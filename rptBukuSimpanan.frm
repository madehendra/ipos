VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptBukuSimpanan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buku Simpanan Harian"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   11775
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   4050
      Left            =   0
      Top             =   975
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   7144
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   3915
         Left            =   90
         TabIndex        =   0
         Top             =   60
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   6906
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "TGL"
         Columns(0).DataField=   "Tgl"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "NO. TRANSAKSI"
         Columns(1).DataField=   "Faktur"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "SD"
         Columns(2).DataField=   "KodeTransaksi"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "KETERANGAN"
         Columns(3).FooterText=   "Total:"
         Columns(3).DataField=   "Keterangan"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "DEBET"
         Columns(4).DataField=   "Debet"
         Columns(4).NumberFormat=   "Standard"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "KREDIT"
         Columns(5).DataField=   "Kredit"
         Columns(5).NumberFormat=   "Standard"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "SALDO AKHIR"
         Columns(6).DataField=   "Saldo"
         Columns(6).NumberFormat=   "Standard"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   4
         Splits(0).SizeMode=   2
         Splits(0).Size  =   4
         Splits(0).Size.vt=   2
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   0
         Splits(0).Caption=   " "
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1720"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1640"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=197121"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=3466"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3387"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=197124"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=661"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=582"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=197121"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=6112"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=6033"
         Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=197124"
         Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(25)=   "Column(4).Width=344"
         Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=265"
         Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=197122"
         Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(31)=   "Column(5).Width=2302"
         Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=2223"
         Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=197122"
         Splits(0)._ColumnProps(36)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(38)=   "Column(6).Width=2699"
         Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=2619"
         Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=197122"
         Splits(0)._ColumnProps(43)=   "Column(6).Visible=0"
         Splits(0)._ColumnProps(44)=   "Column(6).Order=7"
         Splits(1)._UserFlags=   0
         Splits(1).MarqueeStyle=   4
         Splits(1).SizeMode=   2
         Splits(1).Size  =   3
         Splits(1).Size.vt=   2
         Splits(1).RecordSelectors=   0   'False
         Splits(1).RecordSelectorWidth=   688
         Splits(1)._SavedRecordSelectors=   0   'False
         Splits(1).ScrollBars=   2
         Splits(1).Caption=   "Mutasi"
         Splits(1).DividerColor=   14215660
         Splits(1).SpringMode=   0   'False
         Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(1)._ColumnProps(0)=   "Columns.Count=7"
         Splits(1)._ColumnProps(1)=   "Column(0).Width=2170"
         Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=2090"
         Splits(1)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(1)._ColumnProps(5)=   "Column(0)._ColStyle=197121"
         Splits(1)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(1)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(1)._ColumnProps(8)=   "Column(1).Width=3704"
         Splits(1)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(1)._ColumnProps(10)=   "Column(1)._WidthInPix=3625"
         Splits(1)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(1)._ColumnProps(12)=   "Column(1)._ColStyle=197124"
         Splits(1)._ColumnProps(13)=   "Column(1).Visible=0"
         Splits(1)._ColumnProps(14)=   "Column(1).Order=2"
         Splits(1)._ColumnProps(15)=   "Column(2).Width=688"
         Splits(1)._ColumnProps(16)=   "Column(2).DividerColor=0"
         Splits(1)._ColumnProps(17)=   "Column(2)._WidthInPix=609"
         Splits(1)._ColumnProps(18)=   "Column(2)._EditAlways=0"
         Splits(1)._ColumnProps(19)=   "Column(2)._ColStyle=197121"
         Splits(1)._ColumnProps(20)=   "Column(2).Visible=0"
         Splits(1)._ColumnProps(21)=   "Column(2).Order=3"
         Splits(1)._ColumnProps(22)=   "Column(3).Width=4789"
         Splits(1)._ColumnProps(23)=   "Column(3).DividerColor=0"
         Splits(1)._ColumnProps(24)=   "Column(3)._WidthInPix=4710"
         Splits(1)._ColumnProps(25)=   "Column(3)._EditAlways=0"
         Splits(1)._ColumnProps(26)=   "Column(3)._ColStyle=197124"
         Splits(1)._ColumnProps(27)=   "Column(3).Visible=0"
         Splits(1)._ColumnProps(28)=   "Column(3).Order=4"
         Splits(1)._ColumnProps(29)=   "Column(4).Width=2778"
         Splits(1)._ColumnProps(30)=   "Column(4).DividerColor=0"
         Splits(1)._ColumnProps(31)=   "Column(4)._WidthInPix=2699"
         Splits(1)._ColumnProps(32)=   "Column(4)._EditAlways=0"
         Splits(1)._ColumnProps(33)=   "Column(4)._ColStyle=197122"
         Splits(1)._ColumnProps(34)=   "Column(4).Order=5"
         Splits(1)._ColumnProps(35)=   "Column(5).Width=2461"
         Splits(1)._ColumnProps(36)=   "Column(5).DividerColor=0"
         Splits(1)._ColumnProps(37)=   "Column(5)._WidthInPix=2381"
         Splits(1)._ColumnProps(38)=   "Column(5)._EditAlways=0"
         Splits(1)._ColumnProps(39)=   "Column(5)._ColStyle=197122"
         Splits(1)._ColumnProps(40)=   "Column(5).Order=6"
         Splits(1)._ColumnProps(41)=   "Column(6).Width=2752"
         Splits(1)._ColumnProps(42)=   "Column(6).DividerColor=0"
         Splits(1)._ColumnProps(43)=   "Column(6)._WidthInPix=2672"
         Splits(1)._ColumnProps(44)=   "Column(6)._EditAlways=0"
         Splits(1)._ColumnProps(45)=   "Column(6)._ColStyle=197122"
         Splits(1)._ColumnProps(46)=   "Column(6).Order=7"
         Splits.Count    =   2
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         Appearance      =   2
         ColumnFooters   =   -1  'True
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1.5
         FootLines       =   1.5
         MarqueeUnique   =   0   'False
         MultipleLines   =   0
         CellTipsWidth   =   0
         InsertMode      =   0   'False
         DeadAreaBackColor=   -2147483637
         RowDividerColor =   14215660
         RowSubDividerColor=   14215660
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
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bgcolor=&HC0C0C0&,.fgcolor=&H0&"
         _StyleDefs(10)  =   ":id=4,.bold=-1,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(11)  =   ":id=4,.fontname=Times New Roman"
         _StyleDefs(12)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bgcolor=&HEBDACB&"
         _StyleDefs(13)  =   ":id=2,.fgcolor=&H8000000D&,.bold=0,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(14)  =   ":id=2,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(16)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.alignment=1,.bgcolor=&HEBDACB&"
         _StyleDefs(17)  =   ":id=3,.fgcolor=&H8000000D&,.bold=0,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(18)  =   ":id=3,.strikethrough=0,.charset=0"
         _StyleDefs(19)  =   ":id=3,.fontname=MS Sans Serif"
         _StyleDefs(20)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(21)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&H80000014&"
         _StyleDefs(22)  =   ":id=6,.fgcolor=&H80000012&"
         _StyleDefs(23)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(24)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(25)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(26)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(27)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(28)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(29)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(30)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(31)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(32)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(33)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(34)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(35)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(36)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(37)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(38)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(39)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(40)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(41)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=2"
         _StyleDefs(42)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(1).Style:id=62,.parent=13"
         _StyleDefs(46)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=17"
         _StyleDefs(49)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=2"
         _StyleDefs(50)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(51)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(52)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(53)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(54)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(55)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(57)  =   "Splits(0).Columns(4).Style:id=46,.parent=13,.alignment=1"
         _StyleDefs(58)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
         _StyleDefs(59)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
         _StyleDefs(60)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
         _StyleDefs(61)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=1"
         _StyleDefs(62)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
         _StyleDefs(63)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
         _StyleDefs(64)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
         _StyleDefs(65)  =   "Splits(0).Columns(6).Style:id=70,.parent=13,.alignment=1"
         _StyleDefs(66)  =   "Splits(0).Columns(6).HeadingStyle:id=67,.parent=14"
         _StyleDefs(67)  =   "Splits(0).Columns(6).FooterStyle:id=68,.parent=15"
         _StyleDefs(68)  =   "Splits(0).Columns(6).EditorStyle:id=69,.parent=17"
         _StyleDefs(69)  =   "Splits(1).Style:id=55,.parent=1"
         _StyleDefs(70)  =   "Splits(1).CaptionStyle:id=72,.parent=4"
         _StyleDefs(71)  =   "Splits(1).HeadingStyle:id=56,.parent=2"
         _StyleDefs(72)  =   "Splits(1).FooterStyle:id=57,.parent=3"
         _StyleDefs(73)  =   "Splits(1).InactiveStyle:id=58,.parent=5"
         _StyleDefs(74)  =   "Splits(1).SelectedStyle:id=64,.parent=6"
         _StyleDefs(75)  =   "Splits(1).EditorStyle:id=63,.parent=7"
         _StyleDefs(76)  =   "Splits(1).HighlightRowStyle:id=65,.parent=8"
         _StyleDefs(77)  =   "Splits(1).EvenRowStyle:id=66,.parent=9"
         _StyleDefs(78)  =   "Splits(1).OddRowStyle:id=71,.parent=10"
         _StyleDefs(79)  =   "Splits(1).RecordSelectorStyle:id=73,.parent=11"
         _StyleDefs(80)  =   "Splits(1).FilterBarStyle:id=74,.parent=12"
         _StyleDefs(81)  =   "Splits(1).Columns(0).Style:id=78,.parent=55,.alignment=2"
         _StyleDefs(82)  =   "Splits(1).Columns(0).HeadingStyle:id=75,.parent=56"
         _StyleDefs(83)  =   "Splits(1).Columns(0).FooterStyle:id=76,.parent=57"
         _StyleDefs(84)  =   "Splits(1).Columns(0).EditorStyle:id=77,.parent=63"
         _StyleDefs(85)  =   "Splits(1).Columns(1).Style:id=82,.parent=55"
         _StyleDefs(86)  =   "Splits(1).Columns(1).HeadingStyle:id=79,.parent=56"
         _StyleDefs(87)  =   "Splits(1).Columns(1).FooterStyle:id=80,.parent=57"
         _StyleDefs(88)  =   "Splits(1).Columns(1).EditorStyle:id=81,.parent=63"
         _StyleDefs(89)  =   "Splits(1).Columns(2).Style:id=86,.parent=55,.alignment=2"
         _StyleDefs(90)  =   "Splits(1).Columns(2).HeadingStyle:id=83,.parent=56"
         _StyleDefs(91)  =   "Splits(1).Columns(2).FooterStyle:id=84,.parent=57"
         _StyleDefs(92)  =   "Splits(1).Columns(2).EditorStyle:id=85,.parent=63"
         _StyleDefs(93)  =   "Splits(1).Columns(3).Style:id=90,.parent=55"
         _StyleDefs(94)  =   "Splits(1).Columns(3).HeadingStyle:id=87,.parent=56"
         _StyleDefs(95)  =   "Splits(1).Columns(3).FooterStyle:id=88,.parent=57"
         _StyleDefs(96)  =   "Splits(1).Columns(3).EditorStyle:id=89,.parent=63"
         _StyleDefs(97)  =   "Splits(1).Columns(4).Style:id=94,.parent=55,.alignment=1"
         _StyleDefs(98)  =   "Splits(1).Columns(4).HeadingStyle:id=91,.parent=56"
         _StyleDefs(99)  =   "Splits(1).Columns(4).FooterStyle:id=92,.parent=57"
         _StyleDefs(100) =   "Splits(1).Columns(4).EditorStyle:id=93,.parent=63"
         _StyleDefs(101) =   "Splits(1).Columns(5).Style:id=98,.parent=55,.alignment=1"
         _StyleDefs(102) =   "Splits(1).Columns(5).HeadingStyle:id=95,.parent=56"
         _StyleDefs(103) =   "Splits(1).Columns(5).FooterStyle:id=96,.parent=57"
         _StyleDefs(104) =   "Splits(1).Columns(5).EditorStyle:id=97,.parent=63"
         _StyleDefs(105) =   "Splits(1).Columns(6).Style:id=102,.parent=55,.alignment=1"
         _StyleDefs(106) =   "Splits(1).Columns(6).HeadingStyle:id=99,.parent=56"
         _StyleDefs(107) =   "Splits(1).Columns(6).FooterStyle:id=100,.parent=57"
         _StyleDefs(108) =   "Splits(1).Columns(6).EditorStyle:id=101,.parent=63"
         _StyleDefs(109) =   "Named:id=33:Normal"
         _StyleDefs(110) =   ":id=33,.parent=0"
         _StyleDefs(111) =   "Named:id=34:Heading"
         _StyleDefs(112) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(113) =   ":id=34,.wraptext=-1"
         _StyleDefs(114) =   "Named:id=35:Footing"
         _StyleDefs(115) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(116) =   "Named:id=36:Selected"
         _StyleDefs(117) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(118) =   "Named:id=37:Caption"
         _StyleDefs(119) =   ":id=37,.parent=34,.alignment=2,.bold=-1,.fontsize=975,.italic=0,.underline=0"
         _StyleDefs(120) =   ":id=37,.strikethrough=0,.charset=0"
         _StyleDefs(121) =   ":id=37,.fontname=MS Sans Serif"
         _StyleDefs(122) =   "Named:id=38:HighlightRow"
         _StyleDefs(123) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(124) =   "Named:id=39:EvenRow"
         _StyleDefs(125) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(126) =   "Named:id=40:OddRow"
         _StyleDefs(127) =   ":id=40,.parent=33"
         _StyleDefs(128) =   "Named:id=41:RecordSelector"
         _StyleDefs(129) =   ":id=41,.parent=34"
         _StyleDefs(130) =   "Named:id=42:FilterBar"
         _StyleDefs(131) =   ":id=42,.parent=33"
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   1720
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
      Begin BiSATextBoxProject.BiSABrowse cKodeAnggota 
         Height          =   345
         Left            =   105
         TabIndex        =   8
         Top             =   465
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   609
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
         Button          =   -1  'True
         Caption         =   "Kode Anggota"
         CaptionWidth    =   1700
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
         Left            =   5520
         TabIndex        =   1
         Top             =   105
         Width           =   4575
         _ExtentX        =   8070
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
         Button          =   -1  'True
         Caption         =   "NAMA"
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
      Begin BiSADateProject.BiSADate dAwal 
         Height          =   330
         Left            =   105
         TabIndex        =   2
         Top             =   105
         Width           =   3165
         _ExtentX        =   5583
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
         ForeColor       =   -2147483640
         Caption         =   "ANTARA TANGGAL"
         CaptionWidth    =   1700
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
         Left            =   3345
         TabIndex        =   3
         Top             =   105
         Width           =   2010
         _ExtentX        =   3545
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
         ForeColor       =   -2147483640
         Caption         =   "S.D"
         CaptionWidth    =   500
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
         Left            =   5520
         TabIndex        =   4
         Top             =   495
         Width           =   6030
         _ExtentX        =   10636
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
         BackColor       =   12632256
         Enabled         =   0   'False
         Appearance      =   0
         Caption         =   "ALAMAT"
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
      Height          =   630
      Left            =   0
      Top             =   5010
      Width           =   11760
      _ExtentX        =   20743
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
         Left            =   10515
         TabIndex        =   5
         Top             =   90
         Width           =   1140
         _ExtentX        =   2011
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
         Picture         =   "rptBukuSimpanan.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   9345
         TabIndex        =   6
         Top             =   90
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   767
         Caption         =   "     &Preview"
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
         Picture         =   "rptBukuSimpanan.frx":00A6
      End
      Begin BiSAButtonProject.BiSAButton cmdRefresh 
         Height          =   435
         Left            =   8175
         TabIndex        =   7
         Top             =   90
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   767
         Caption         =   "     &Refresh"
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
         Picture         =   "rptBukuSimpanan.frx":032C
      End
   End
End
Attribute VB_Name = "rptBukuSimpanan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim dbRekening As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim xarray As New XArrayDB
Dim nAwal As Double
Dim lPreview As Boolean

Private Sub cKodeAnggota_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "anggota", "kodeanggota,nama,alamat", "kodeanggota", sisContent, cKodeAnggota.Text, , "kodeanggota")
  If Not dbData.EOF Then
    cKodeAnggota.Text = cKodeAnggota.Browse(dbData, Array("Kode", "Nama", "Alamat"), , Array(15, 20, 25))
    cAlamat.Text = GetNull(dbData!alamat)
    cNama.Text = GetNull(dbData!nama)
  End If
  GetSQL
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
  lPreview = True
  GetSQL
End Sub

Private Sub cmdRefresh_Click()
  lPreview = False
  GetSQL
End Sub

Private Sub cNama_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "anggota", "nama,kodeanggota,alamat", "nama", sisContent, cNama.Text, , "nama")
  If Not dbData.EOF Then
    cNama.Text = cNama.Browse(dbData, Array("Nama", "Kode", "Alamat"), , Array(20, 15, 25))
    cAlamat.Text = GetNull(dbData!alamat)
    cKodeAnggota.Text = GetNull(dbData!kodeanggota)
  End If
  GetSQL
End Sub

Private Sub Form_Load()
Dim n As Single
  
  SetIcon Me.hWnd
  ValidDisplay
  CenterForm Me
  dAwal.Value = BOM(Date)
  dAkhir.Value = EOM(Date)
  cNama.Default
  cAlamat.Default
  
  TabIndex dAwal, n
  TabIndex dAkhir, n
  TabIndex cKodeAnggota, n
  TabIndex cNama, n
  TabIndex cAlamat, n
  TabIndex cmdRefresh, n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub

Private Sub ValidDisplay()
Dim n As Integer
Dim i As Integer
  For n = 0 To tdbgrid1.Splits.Count - 1
    For i = 0 To tdbgrid1.Splits(n).Columns.Count - 1
      If n = 0 Then
        tdbgrid1.Splits(n).Columns(i).Visible = i <= 3
      Else
        tdbgrid1.Splits(n).Columns(i).Visible = i > 3
      End If
    Next
  Next
End Sub

Private Sub GetSQL()
Dim nSaldo As Double
Dim n As Long
Dim nDebet As Double
Dim nKredit As Double

  xarray.ReDim 0, 0, 0, 6
  xarray(0, 3) = "SALDO PER " & Format(dAwal.Value - 1, "dd/MM/yy")
  xarray(0, 6) = 0
  
  Set dbData = objData.Browse(GetDSN, "MutasiTabungan", , "Tgl", sisLTEqual, Format(dAkhir.Value, "yyyy-mm-dd"), " and kodeanggota = '" & cKodeAnggota.Text & "'", "Tgl,kodeanggota")
  If Not dbData.EOF Then
    dbData.MoveFirst
    Do While Not dbData.EOF
      If dbData!tgl < dAwal.Value Then
        xarray(0, 6) = (GetNull(xarray(0, 6) + IIf(dbData!DK = "K", dbData!Jumlah, -dbData!Jumlah)))
      Else
        n = n + 1
        xarray.InsertRows n
        xarray(n, 0) = Format((dbData!tgl), "dd/MM/yy")
        xarray(n, 1) = GetNull(dbData!nomormutasitabungan)
        xarray(n, 2) = GetNull(dbData!KodeTransaksi)
        xarray(n, 3) = (dbData!Keterangan)
        xarray(n, 4) = GetNull(IIf(GetNull(dbData!DK) = "D", GetNull(dbData!Jumlah), 0))
        xarray(n, 5) = GetNull(IIf(GetNull(dbData!DK) = "K", GetNull(dbData!Jumlah), 0))
        nDebet = nDebet + xarray(n, 4)
        nKredit = nKredit + xarray(n, 5)
        xarray(n, 6) = (xarray(n - 1, 6) - xarray(n, 4) + xarray(n, 5))
      End If
      dbData.MoveNext
    Loop
    tdbgrid1.Columns(4).FooterText = Format(nDebet, "###,###,###,###,##0.00")
    tdbgrid1.Columns(5).FooterText = Format(nKredit, "###,###,###,###,##0.00")
    tdbgrid1.Array = xarray
    tdbgrid1.ReBind
    If lPreview = True Then
      Rpt
    End If
  Else
    MsgBox "Data tidak ada", vbInformation
    Exit Sub
  End If
End Sub

Private Sub TDBGrid1_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
Dim dTanggal As Date

  Select Case ColIndex
    Case 0
      If Value <> "" Then
        dTanggal = Format(Value, "dd-MM-yyyy")
        Value = Format(dTanggal, "dd-MM-yyyy")
      Else
        Value = ""
      End If
      
    Case Else
      If Value = 0 Then
        Value = ""
      Else
        Value = Format(Value, "###,###,##0.00")
      End If
  End Select
End Sub

Private Sub Rpt()
  With FrmRPT
    .AddPageHeader UCase("Buku SIMPANAN sukarela"), tdbHalignCenter, , , , , 10, True
    .AddPageHeader UCase(aCfg(objData, msNamaPerusahaan)), tdbHalignCenter, , , True, , 12, True
    .AddPageHeader "Antara Tanggal : " & Format(dAwal.Value, "dd-MM-yyyy") & " s.d " & Format(dAkhir.Value, "dd-MM-yyyy"), tdbHalignCenter, , , True, , 10, True
    .AddPageHeader " ", , , , True
    .AddPageHeader " ", , , , True
    
    .AddPageHeader "Nama Anggota", tdbHalignLeft, , 15, , , , , , True, , tdbPageHeaderSect
    .AddPageHeader ": " & cNama.Text
    .AddPageHeader "Kode", tdbHalignLeft, , 15, True, , , , , True, , tdbPageHeaderSect
    .AddPageHeader ": " & cKodeAnggota.Text
    .AddPageHeader "Alamat", tdbHalignLeft, , 15, True, , , , , True, , tdbPageHeaderSect
    .AddPageHeader ": " & cAlamat.Text
    
    .AddTableHeader "TANGGAL", , , , 8, , , , , , True, tdbTableHeaderSect
    .AddTableHeader "NO. TRANSAKSI", , , , 17
    .AddTableHeader "SD", , , , 5
    .AddTableHeader "KETERANGAN"
    .AddTableHeader "DEBET", , , , 11
    .AddTableHeader "KREDIT", , , , 11
    .AddTableHeader "SALDO AKHIR", , , , 13
    
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    
    .AddTableFooter "Total", , tdbHalignCenter, , , , , , , , , , , , 4
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter
    
    .Preview xarray, True
  End With
End Sub



