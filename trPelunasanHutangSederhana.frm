VERSION 5.00
Object = "{0D6235E7-DBA2-11D1-B5DF-0060976089D0}#1.0#0"; "tdbr6.ocx"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trPelunasanHutangSederhana 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ":. PELUNASAN HUTANG SEDERHANA"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17550
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   17550
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   6915
      Left            =   6150
      Top             =   195
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   12197
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
      BorderStyle     =   0
      BackColor       =   -2147483633
      Begin BiSAButtonProject.BiSAButton BiSAButton1 
         Height          =   330
         Left            =   4515
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   360
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   582
         Caption         =   ".:"
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
         Height          =   2940
         Left            =   60
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   765
         Width           =   11010
         _ExtentX        =   19420
         _ExtentY        =   5186
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
         Columns(4).Caption=   "KETERANGAN"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "DEBET"
         Columns(5).DataField=   ""
         Columns(5).NumberFormat=   "###,###,###,###,##0.00"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "KREDIT"
         Columns(6).DataField=   ""
         Columns(6).NumberFormat=   "###,###,###,###,##0.00"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "FLAG"
         Columns(7).DataField=   ""
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "FLAGID"
         Columns(8).DataField=   ""
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   9
         Splits(0)._UserFlags=   0
         Splits(0).AllowSizing=   -1  'True
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).FetchRowStyle=   -1  'True
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=9"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=609"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=529"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=197124"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=1164"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1085"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=197124"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2408"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2328"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=197124"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=2355"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2275"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=197121"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=6509"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=6429"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=197124"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=2593"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2514"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=197122"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=2752"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2672"
         Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=197122"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(36)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=197124"
         Splits(0)._ColumnProps(40)=   "Column(7).FetchStyle=1"
         Splits(0)._ColumnProps(41)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(42)=   "Column(8).Width=2725"
         Splits(0)._ColumnProps(43)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(44)=   "Column(8)._WidthInPix=2646"
         Splits(0)._ColumnProps(45)=   "Column(8)._ColStyle=197124"
         Splits(0)._ColumnProps(46)=   "Column(8).Order=9"
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
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=46,.parent=13,.bgcolor=&H80000005&"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17,.bgcolor=&H8000000F&"
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
         _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=50,.parent=13"
         _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
         _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
         _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=1"
         _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.alignment=1"
         _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=54,.parent=13"
         _StyleDefs(66)  =   "Splits(0).Columns(7).HeadingStyle:id=51,.parent=14"
         _StyleDefs(67)  =   "Splits(0).Columns(7).FooterStyle:id=52,.parent=15"
         _StyleDefs(68)  =   "Splits(0).Columns(7).EditorStyle:id=53,.parent=17"
         _StyleDefs(69)  =   "Splits(0).Columns(8).Style:id=66,.parent=13"
         _StyleDefs(70)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
         _StyleDefs(71)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
         _StyleDefs(72)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
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
         _StyleDefs(83)  =   ":id=37,.parent=34,.alignment=2,.bold=0,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(84)  =   ":id=37,.strikethrough=0,.charset=0"
         _StyleDefs(85)  =   ":id=37,.fontname=Tahoma"
         _StyleDefs(86)  =   "Named:id=38:HighlightRow"
         _StyleDefs(87)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(88)  =   "Named:id=39:EvenRow"
         _StyleDefs(89)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(90)  =   "Named:id=40:OddRow"
         _StyleDefs(91)  =   ":id=40,.parent=33"
         _StyleDefs(92)  =   "Named:id=41:RecordSelector"
         _StyleDefs(93)  =   ":id=41,.parent=34"
         _StyleDefs(94)  =   "Named:id=42:FilterBar"
         _StyleDefs(95)  =   ":id=42,.parent=33"
      End
      Begin BiSADateProject.BiSADate dTglMutasi 
         Height          =   324
         Index           =   0
         Left            =   180
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   360
         Width           =   2796
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
         Caption         =   "Dari Tanggal"
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
      Begin BiSADateProject.BiSADate dTglMutasi 
         Height          =   324
         Index           =   1
         Left            =   3048
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   360
         Width           =   1440
         _ExtentX        =   2540
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
      Begin TrueOleDBGrid70.TDBGrid GridVoucher 
         Height          =   3105
         Left            =   60
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   3750
         Width           =   11025
         _ExtentX        =   19447
         _ExtentY        =   5477
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
         Columns(2).NumberFormat=   "dd-MM-yyyy"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "KETERANGAN"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "DEBET"
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "###,###,###,###,##0.00"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "FLAG"
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "FLAGID"
         Columns(6).DataField=   ""
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
         Splits(0)._UserFlags=   0
         Splits(0).AllowSizing=   -1  'True
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).FetchRowStyle=   -1  'True
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=609"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=529"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=197124"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2408"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2328"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=197124"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2355"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2275"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=197121"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=6271"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=6191"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=197124"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2593"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2514"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=197122"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=1402"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=1323"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=197124"
         Splits(0)._ColumnProps(30)=   "Column(5).FetchStyle=1"
         Splits(0)._ColumnProps(31)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(32)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(33)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(34)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(35)=   "Column(6)._ColStyle=197124"
         Splits(0)._ColumnProps(36)=   "Column(6).Order=7"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         ColumnFooters   =   -1  'True
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         Caption         =   "VOUCHER -TOP UP"
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
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=46,.parent=13,.bgcolor=&H80000005&"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17,.bgcolor=&H8000000F&"
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=70,.parent=13,.alignment=2"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=67,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=68,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=69,.parent=17"
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=58,.parent=13,.alignment=1"
         _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
         _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
         _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=54,.parent=13"
         _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
         _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
         _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
         _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=66,.parent=13"
         _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=63,.parent=14"
         _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=64,.parent=15"
         _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=65,.parent=17"
         _StyleDefs(65)  =   "Named:id=33:Normal"
         _StyleDefs(66)  =   ":id=33,.parent=0"
         _StyleDefs(67)  =   "Named:id=34:Heading"
         _StyleDefs(68)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(69)  =   ":id=34,.wraptext=-1"
         _StyleDefs(70)  =   "Named:id=35:Footing"
         _StyleDefs(71)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(72)  =   "Named:id=36:Selected"
         _StyleDefs(73)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(74)  =   "Named:id=37:Caption"
         _StyleDefs(75)  =   ":id=37,.parent=34,.alignment=2,.bold=0,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(76)  =   ":id=37,.strikethrough=0,.charset=0"
         _StyleDefs(77)  =   ":id=37,.fontname=Tahoma"
         _StyleDefs(78)  =   "Named:id=38:HighlightRow"
         _StyleDefs(79)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(80)  =   "Named:id=39:EvenRow"
         _StyleDefs(81)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(82)  =   "Named:id=40:OddRow"
         _StyleDefs(83)  =   ":id=40,.parent=33"
         _StyleDefs(84)  =   "Named:id=41:RecordSelector"
         _StyleDefs(85)  =   ":id=41,.parent=34"
         _StyleDefs(86)  =   "Named:id=42:FilterBar"
         _StyleDefs(87)  =   ":id=42,.parent=33"
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   2790
      Left            =   60
      Top             =   75
      Width           =   6060
      _ExtentX        =   10689
      _ExtentY        =   4921
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
      Begin BiSATextBoxProject.BiSABrowse cNama 
         Height          =   330
         Left            =   75
         TabIndex        =   1
         Top             =   735
         Width           =   4440
         _ExtentX        =   7832
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
         Height          =   324
         Left            =   72
         TabIndex        =   2
         Top             =   372
         Width           =   2796
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
      Begin BiSATextBoxProject.BiSABrowse cCustomer 
         Height          =   330
         Left            =   1500
         TabIndex        =   3
         Top             =   1095
         Width           =   4065
         _ExtentX        =   7170
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
      Begin BiSATextBoxProject.BiSABrowse cFaktur 
         Height          =   330
         Left            =   75
         TabIndex        =   4
         Top             =   2160
         Width           =   3435
         _ExtentX        =   6059
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
      Begin BiSATextBoxProject.BiSABrowse cDepartement 
         Height          =   330
         Left            =   1500
         TabIndex        =   19
         Top             =   1455
         Width           =   4065
         _ExtentX        =   7170
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
      Begin BiSANumberBoxProject.BiSANumberBox nOutstanding 
         Height          =   330
         Left            =   75
         TabIndex        =   22
         Top             =   1800
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   582
         Appearance      =   0
         Enabled         =   0   'False
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         Caption         =   "TOTAL BON"
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
         TabIndex        =   5
         Top             =   90
         Width           =   4890
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame4 
      Height          =   630
      Left            =   45
      Top             =   7095
      Width           =   17445
      _ExtentX        =   30771
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
         Left            =   1170
         TabIndex        =   6
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
         Picture         =   "trPelunasanHutangSederhana.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   14730
         TabIndex        =   7
         Top             =   90
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
         Picture         =   "trPelunasanHutangSederhana.frx":028A
      End
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   2325
         TabIndex        =   8
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
         Picture         =   "trPelunasanHutangSederhana.frx":0429
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   435
         Left            =   90
         TabIndex        =   9
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
         Picture         =   "trPelunasanHutangSederhana.frx":0555
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   16275
         TabIndex        =   10
         Top             =   90
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
         Picture         =   "trPelunasanHutangSederhana.frx":0700
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   15180
         TabIndex        =   11
         Top             =   90
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
         Picture         =   "trPelunasanHutangSederhana.frx":07A6
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   4245
      Left            =   60
      Top             =   2865
      Width           =   6060
      _ExtentX        =   10689
      _ExtentY        =   7488
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
      Begin BiSANumberBoxProject.BiSANumberBox nTunai 
         Height          =   330
         Left            =   2145
         TabIndex        =   12
         Top             =   1920
         Width           =   3420
         _ExtentX        =   6033
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
         Caption         =   "TUNAI"
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
      Begin BiSANumberBoxProject.BiSANumberBox nVoucher 
         Height          =   330
         Left            =   2145
         TabIndex        =   13
         Top             =   1515
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
         BackColor       =   -2147483633
         Caption         =   "VOUCHER"
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
         Left            =   2145
         TabIndex        =   15
         Top             =   2430
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   582
         Appearance      =   0
         Enabled         =   0   'False
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
      Begin BiSANumberBoxProject.BiSANumberBox nSaldoPiutang 
         Height          =   330
         Left            =   1935
         TabIndex        =   20
         Top             =   1125
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   582
         Appearance      =   0
         Enabled         =   0   'False
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         Caption         =   "TOTAL BON"
         CaptionWidth    =   1200
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin BiSATextBoxProject.BiSABrowse cAkunKas 
         Height          =   345
         Left            =   1845
         TabIndex        =   23
         Top             =   2820
         Width           =   3735
         _ExtentX        =   6588
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
      Begin TrueDBReports60Ctl.TDBReports rptKuitansiLunas 
         Height          =   570
         Left            =   240
         TabIndex        =   24
         Top             =   105
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   1005
         Caption         =   "Kuitansi Lunas"
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         ErrorMsgCaption =   ""
         Filtered        =   0   'False
         DataMode        =   1
         DataMember      =   ""
         LinkSequence    =   1
         LinkOrder       =   0
         NameSubstitute  =   ""
         ConnectionString=   "DSN=MySalemba"
         ConnectStringType=   3
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "MySalemba"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         CursorLocation  =   3
         ConnectionTimeout=   15
         CommandTimeout  =   30
         RecordSource    =   ""
         CursorType      =   1
         CommandType     =   8
         MaxRecords      =   0
         LinkType        =   0
         Master          =   ""
         CallDataRead    =   0   'False
         ConvertNullToEmpty=   -1  'True
         DesignConnection=   -1  'True
         DesignTimeout   =   5
         UnitsOfMeasurement=   4
         Vedit_ShowGrid  =   -1  'True
         Vedit_SnapToGrid=   0   'False
         Vedit_GridUnitWidth=   2.822
         Vedit_GridUnitHeight=   2.822
         Vedit_ShowCellExpressions=   -1  'True
         Norm_rect_left  =   0
         Norm_rect_top   =   0
         Norm_rect_right =   0
         Norm_rect_bottom=   0
         Virgin          =   0   'False
         Parameters.Count=   29
         Parameters(0).Name=   "cSE"
         Parameters(0).ValueExpression=   """"""
         Parameters(1).Name=   "cNama"
         Parameters(1).ValueExpression=   """"""
         Parameters(2).Name=   "cAlamat"
         Parameters(2).ValueExpression=   """"""
         Parameters(3).Name=   "cKota"
         Parameters(3).ValueExpression=   """"""
         Parameters(4).Name=   "cTerbilang"
         Parameters(4).ValueExpression=   """"""
         Parameters(5).Name=   "dTgl"
         Parameters(6).Name=   "dJTHTMP"
         Parameters(6).Type=   7
         Parameters(7).Name=   "cTTD"
         Parameters(8).Name=   "nSubTotal"
         Parameters(8).Type=   5
         Parameters(8).ValueExpression=   "0"
         Parameters(9).Name=   "nTotal"
         Parameters(9).Type=   5
         Parameters(9).ValueExpression=   "0"
         Parameters(10).Name=   "nPPn"
         Parameters(10).ValueExpression=   "0"
         Parameters(11).Name=   "nPajak"
         Parameters(11).Type=   5
         Parameters(11).ValueExpression=   "0"
         Parameters(12).Name=   "cNamaPerusahaan"
         Parameters(13).Name=   "cAlamatPerusahaan"
         Parameters(14).Name=   "cTeleponPerusahaan"
         Parameters(15).Name=   "cReceived"
         Parameters(16).Name=   "cKetReceived"
         Parameters(17).Name=   "cRef"
         Parameters(18).Name=   "cPerusahaanLine"
         Parameters(19).Name=   "cPayment"
         Parameters(20).Name=   "cUserName"
         Parameters(21).Name=   "nDiscount"
         Parameters(21).Type=   5
         Parameters(22).Name=   "cJudul"
         Parameters(23).Name=   "cSales"
         Parameters(24).Name=   "cFooter"
         Parameters(25).Name=   "nDp"
         Parameters(26).Name=   "cFooter2"
         Parameters(27).Name=   "cKodeAnggota"
         Parameters(28).Name=   "keAkun"
         Sections.Count  =   2
         Sections(0).Name=   "SECTION_2"
         Sections(0).Type=   1
         Sections(0).StyleExp=   "'Tdb_Base'"
         Sections(0).Cells.Count=   16
         Sections(0).Cells(0).Name=   "CELL_22"
         Sections(0).Cells(0).Exp=   "cNamaPerusahaan"
         Sections(0).Cells(0).NewLine=   -1  'True
         Sections(0).Cells(0).PrivateStyle=   -1  'True
         Sections(0).Cells(0).Style.Name=   "<private>"
         Sections(0).Cells(0).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(0).Style.Font_Name=   "Courier"
         Sections(0).Cells(0).Style.Font_Size=   12
         Sections(0).Cells(0).Style.Font_Bold=   -1  'True
         Sections(0).Cells(0).Style.Font_Italic=   0   'False
         Sections(0).Cells(0).Style.Font_Underline=   0   'False
         Sections(0).Cells(0).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(0).Style.Font_Charset=   0
         Sections(0).Cells(0).Style.TextAlign=   0
         Sections(0).Cells(0).Style.TextVAlign=   1
         Sections(0).Cells(0).Style.TextWrap=   -1  'True
         Sections(0).Cells(0).Style.ForeColor=   0
         Sections(0).Cells(0).Style.BackColor=   16777215
         Sections(0).Cells(0).Style.NoFill=   -1  'True
         Sections(0).Cells(0).Style.BackPicFile=   ""
         Sections(0).Cells(0).Style.ForePicFile=   ""
         Sections(0).Cells(0).Style.BackPicVertPlacement=   0
         Sections(0).Cells(0).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(0).Style.ForePicPlacement=   0
         Sections(0).Cells(0).Style.ForePicDrawMode=   0
         Sections(0).Cells(0).Style.MarginLeft=   6
         Sections(0).Cells(0).Style.MarginTop=   1
         Sections(0).Cells(0).Style.MarginRight=   6
         Sections(0).Cells(0).Style.MarginBottom=   1
         Sections(0).Cells(0).Style.HasBorders=   -1  'True
         Sections(0).Cells(0).Style.BorderHT=   ""
         Sections(0).Cells(0).Style.BorderHI=   ""
         Sections(0).Cells(0).Style.BorderHB=   ""
         Sections(0).Cells(0).Style.BorderVL=   ""
         Sections(0).Cells(0).Style.BorderVI=   ""
         Sections(0).Cells(0).Style.BorderVR=   ""
         Sections(0).Cells(0).Style.NoClipping=   0   'False
         Sections(0).Cells(0).Style.RTF=   0   'False
         Sections(0).Cells(0).Style.fprops=   89391105
         Sections(0).Cells(1).Name=   "CELL_25"
         Sections(0).Cells(1).Exp=   "cAlamatPerusahaan"
         Sections(0).Cells(1).NewLine=   -1  'True
         Sections(0).Cells(1).Height=   5
         Sections(0).Cells(1).AutoHeight=   0   'False
         Sections(0).Cells(1).PrivateStyle=   -1  'True
         Sections(0).Cells(1).Style.Name=   "<private>"
         Sections(0).Cells(1).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(1).Style.Font_Name=   "Courier"
         Sections(0).Cells(1).Style.Font_Size=   9.75
         Sections(0).Cells(1).Style.Font_Bold=   0   'False
         Sections(0).Cells(1).Style.Font_Italic=   0   'False
         Sections(0).Cells(1).Style.Font_Underline=   0   'False
         Sections(0).Cells(1).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(1).Style.Font_Charset=   0
         Sections(0).Cells(1).Style.TextAlign=   0
         Sections(0).Cells(1).Style.TextVAlign=   1
         Sections(0).Cells(1).Style.TextWrap=   -1  'True
         Sections(0).Cells(1).Style.ForeColor=   0
         Sections(0).Cells(1).Style.BackColor=   16777215
         Sections(0).Cells(1).Style.NoFill=   -1  'True
         Sections(0).Cells(1).Style.BackPicFile=   ""
         Sections(0).Cells(1).Style.ForePicFile=   ""
         Sections(0).Cells(1).Style.BackPicVertPlacement=   0
         Sections(0).Cells(1).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(1).Style.ForePicPlacement=   0
         Sections(0).Cells(1).Style.ForePicDrawMode=   0
         Sections(0).Cells(1).Style.MarginLeft=   6
         Sections(0).Cells(1).Style.MarginTop=   1
         Sections(0).Cells(1).Style.MarginRight=   6
         Sections(0).Cells(1).Style.MarginBottom=   1
         Sections(0).Cells(1).Style.HasBorders=   0   'False
         Sections(0).Cells(1).Style.BorderHT=   ""
         Sections(0).Cells(1).Style.BorderHI=   ""
         Sections(0).Cells(1).Style.BorderHB=   ""
         Sections(0).Cells(1).Style.BorderVL=   ""
         Sections(0).Cells(1).Style.BorderVI=   ""
         Sections(0).Cells(1).Style.BorderVR=   ""
         Sections(0).Cells(1).Style.NoClipping=   0   'False
         Sections(0).Cells(1).Style.RTF=   0   'False
         Sections(0).Cells(1).Style.fprops=   22413313
         Sections(0).Cells(2).Name=   "CELL_2"
         Sections(0).Cells(2).Exp=   """"""
         Sections(0).Cells(2).NewLine=   -1  'True
         Sections(0).Cells(2).PrivateStyle=   -1  'True
         Sections(0).Cells(2).Style.Name=   "<private>"
         Sections(0).Cells(2).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(2).Style.Font_Name=   "Courier"
         Sections(0).Cells(2).Style.Font_Size=   9.75
         Sections(0).Cells(2).Style.Font_Bold=   -1  'True
         Sections(0).Cells(2).Style.Font_Italic=   0   'False
         Sections(0).Cells(2).Style.Font_Underline=   0   'False
         Sections(0).Cells(2).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(2).Style.Font_Charset=   0
         Sections(0).Cells(2).Style.TextAlign=   3
         Sections(0).Cells(2).Style.TextVAlign=   1
         Sections(0).Cells(2).Style.TextWrap=   -1  'True
         Sections(0).Cells(2).Style.ForeColor=   0
         Sections(0).Cells(2).Style.BackColor=   16777215
         Sections(0).Cells(2).Style.NoFill=   -1  'True
         Sections(0).Cells(2).Style.BackPicFile=   ""
         Sections(0).Cells(2).Style.ForePicFile=   ""
         Sections(0).Cells(2).Style.BackPicVertPlacement=   0
         Sections(0).Cells(2).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(2).Style.ForePicPlacement=   0
         Sections(0).Cells(2).Style.ForePicDrawMode=   0
         Sections(0).Cells(2).Style.MarginLeft=   6
         Sections(0).Cells(2).Style.MarginTop=   1
         Sections(0).Cells(2).Style.MarginRight=   6
         Sections(0).Cells(2).Style.MarginBottom=   1
         Sections(0).Cells(2).Style.HasBorders=   -1  'True
         Sections(0).Cells(2).Style.BorderHT=   ""
         Sections(0).Cells(2).Style.BorderHI=   ""
         Sections(0).Cells(2).Style.BorderHB=   ""
         Sections(0).Cells(2).Style.BorderVL=   ""
         Sections(0).Cells(2).Style.BorderVI=   ""
         Sections(0).Cells(2).Style.BorderVR=   ""
         Sections(0).Cells(2).Style.NoClipping=   0   'False
         Sections(0).Cells(2).Style.RTF=   0   'False
         Sections(0).Cells(2).Style.fprops=   131072
         Sections(0).Cells(3).Name=   "CELL_26"
         Sections(0).Cells(3).Exp=   """ """
         Sections(0).Cells(3).NewLine=   -1  'True
         Sections(0).Cells(3).PrivateStyle=   -1  'True
         Sections(0).Cells(3).Style.Name=   "<private>"
         Sections(0).Cells(3).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(3).Style.Font_Name=   "Courier"
         Sections(0).Cells(3).Style.Font_Size=   9.75
         Sections(0).Cells(3).Style.Font_Bold=   -1  'True
         Sections(0).Cells(3).Style.Font_Italic=   0   'False
         Sections(0).Cells(3).Style.Font_Underline=   0   'False
         Sections(0).Cells(3).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(3).Style.Font_Charset=   0
         Sections(0).Cells(3).Style.TextAlign=   1
         Sections(0).Cells(3).Style.TextVAlign=   1
         Sections(0).Cells(3).Style.TextWrap=   -1  'True
         Sections(0).Cells(3).Style.ForeColor=   0
         Sections(0).Cells(3).Style.BackColor=   16777215
         Sections(0).Cells(3).Style.NoFill=   -1  'True
         Sections(0).Cells(3).Style.BackPicFile=   ""
         Sections(0).Cells(3).Style.ForePicFile=   ""
         Sections(0).Cells(3).Style.BackPicVertPlacement=   0
         Sections(0).Cells(3).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(3).Style.ForePicPlacement=   0
         Sections(0).Cells(3).Style.ForePicDrawMode=   0
         Sections(0).Cells(3).Style.MarginLeft=   6
         Sections(0).Cells(3).Style.MarginTop=   1
         Sections(0).Cells(3).Style.MarginRight=   6
         Sections(0).Cells(3).Style.MarginBottom=   1
         Sections(0).Cells(3).Style.HasBorders=   -1  'True
         Sections(0).Cells(3).Style.BorderHT=   ""
         Sections(0).Cells(3).Style.BorderHI=   ""
         Sections(0).Cells(3).Style.BorderHB=   ""
         Sections(0).Cells(3).Style.BorderVL=   ""
         Sections(0).Cells(3).Style.BorderVI=   ""
         Sections(0).Cells(3).Style.BorderVR=   ""
         Sections(0).Cells(3).Style.NoClipping=   0   'False
         Sections(0).Cells(3).Style.RTF=   0   'False
         Sections(0).Cells(3).Style.fprops=   68419585
         Sections(0).Cells(4).Name=   "CELL_3"
         Sections(0).Cells(4).Exp=   """ """
         Sections(0).Cells(4).NewLine=   -1  'True
         Sections(0).Cells(4).Width=   30
         Sections(0).Cells(4).PrivateStyle=   -1  'True
         Sections(0).Cells(4).Style.Name=   "<private>"
         Sections(0).Cells(4).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(4).Style.Font_Name=   "Courier"
         Sections(0).Cells(4).Style.Font_Size=   9.75
         Sections(0).Cells(4).Style.Font_Bold=   0   'False
         Sections(0).Cells(4).Style.Font_Italic=   0   'False
         Sections(0).Cells(4).Style.Font_Underline=   0   'False
         Sections(0).Cells(4).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(4).Style.Font_Charset=   0
         Sections(0).Cells(4).Style.TextAlign=   3
         Sections(0).Cells(4).Style.TextVAlign=   1
         Sections(0).Cells(4).Style.TextWrap=   -1  'True
         Sections(0).Cells(4).Style.ForeColor=   0
         Sections(0).Cells(4).Style.BackColor=   16777215
         Sections(0).Cells(4).Style.NoFill=   -1  'True
         Sections(0).Cells(4).Style.BackPicFile=   ""
         Sections(0).Cells(4).Style.ForePicFile=   ""
         Sections(0).Cells(4).Style.BackPicVertPlacement=   0
         Sections(0).Cells(4).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(4).Style.ForePicPlacement=   0
         Sections(0).Cells(4).Style.ForePicDrawMode=   0
         Sections(0).Cells(4).Style.MarginLeft=   6
         Sections(0).Cells(4).Style.MarginTop=   1
         Sections(0).Cells(4).Style.MarginRight=   6
         Sections(0).Cells(4).Style.MarginBottom=   1
         Sections(0).Cells(4).Style.HasBorders=   -1  'True
         Sections(0).Cells(4).Style.BorderHT=   ""
         Sections(0).Cells(4).Style.BorderHI=   ""
         Sections(0).Cells(4).Style.BorderHB=   ""
         Sections(0).Cells(4).Style.BorderVL=   ""
         Sections(0).Cells(4).Style.BorderVI=   ""
         Sections(0).Cells(4).Style.BorderVR=   ""
         Sections(0).Cells(4).Style.NoClipping=   0   'False
         Sections(0).Cells(4).Style.RTF=   0   'False
         Sections(0).Cells(4).Style.fprops=   22282240
         Sections(0).Cells(5).Name=   "CELL_27"
         Sections(0).Cells(5).Exp=   """Kuitansi"""
         Sections(0).Cells(5).PrivateStyle=   -1  'True
         Sections(0).Cells(5).Style.Name=   "<private>"
         Sections(0).Cells(5).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(5).Style.Font_Name=   "Courier"
         Sections(0).Cells(5).Style.Font_Size=   9.75
         Sections(0).Cells(5).Style.Font_Bold=   -1  'True
         Sections(0).Cells(5).Style.Font_Italic=   0   'False
         Sections(0).Cells(5).Style.Font_Underline=   0   'False
         Sections(0).Cells(5).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(5).Style.Font_Charset=   0
         Sections(0).Cells(5).Style.TextAlign=   1
         Sections(0).Cells(5).Style.TextVAlign=   1
         Sections(0).Cells(5).Style.TextWrap=   -1  'True
         Sections(0).Cells(5).Style.ForeColor=   0
         Sections(0).Cells(5).Style.BackColor=   16777215
         Sections(0).Cells(5).Style.NoFill=   -1  'True
         Sections(0).Cells(5).Style.BackPicFile=   ""
         Sections(0).Cells(5).Style.ForePicFile=   ""
         Sections(0).Cells(5).Style.BackPicVertPlacement=   0
         Sections(0).Cells(5).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(5).Style.ForePicPlacement=   0
         Sections(0).Cells(5).Style.ForePicDrawMode=   0
         Sections(0).Cells(5).Style.MarginLeft=   6
         Sections(0).Cells(5).Style.MarginTop=   1
         Sections(0).Cells(5).Style.MarginRight=   6
         Sections(0).Cells(5).Style.MarginBottom=   1
         Sections(0).Cells(5).Style.HasBorders=   -1  'True
         Sections(0).Cells(5).Style.BorderHT=   ""
         Sections(0).Cells(5).Style.BorderHI=   ""
         Sections(0).Cells(5).Style.BorderHB=   "None"
         Sections(0).Cells(5).Style.BorderVL=   ""
         Sections(0).Cells(5).Style.BorderVI=   ""
         Sections(0).Cells(5).Style.BorderVR=   ""
         Sections(0).Cells(5).Style.NoClipping=   0   'False
         Sections(0).Cells(5).Style.RTF=   0   'False
         Sections(0).Cells(5).Style.fprops=   84017153
         Sections(0).Cells(6).Name=   "CELL_12"
         Sections(0).Cells(6).Exp=   """Tgl : "" & dTgl"
         Sections(0).Cells(6).Width=   30
         Sections(0).Cells(6).PrivateStyle=   -1  'True
         Sections(0).Cells(6).Style.Name=   "<private>"
         Sections(0).Cells(6).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(6).Style.Font_Name=   "Courier"
         Sections(0).Cells(6).Style.Font_Size=   9.75
         Sections(0).Cells(6).Style.Font_Bold=   0   'False
         Sections(0).Cells(6).Style.Font_Italic=   0   'False
         Sections(0).Cells(6).Style.Font_Underline=   0   'False
         Sections(0).Cells(6).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(6).Style.Font_Charset=   0
         Sections(0).Cells(6).Style.TextAlign=   2
         Sections(0).Cells(6).Style.TextVAlign=   1
         Sections(0).Cells(6).Style.TextWrap=   -1  'True
         Sections(0).Cells(6).Style.ForeColor=   0
         Sections(0).Cells(6).Style.BackColor=   16777215
         Sections(0).Cells(6).Style.NoFill=   -1  'True
         Sections(0).Cells(6).Style.BackPicFile=   ""
         Sections(0).Cells(6).Style.ForePicFile=   ""
         Sections(0).Cells(6).Style.BackPicVertPlacement=   0
         Sections(0).Cells(6).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(6).Style.ForePicPlacement=   0
         Sections(0).Cells(6).Style.ForePicDrawMode=   0
         Sections(0).Cells(6).Style.MarginLeft=   6
         Sections(0).Cells(6).Style.MarginTop=   1
         Sections(0).Cells(6).Style.MarginRight=   6
         Sections(0).Cells(6).Style.MarginBottom=   1
         Sections(0).Cells(6).Style.HasBorders=   -1  'True
         Sections(0).Cells(6).Style.BorderHT=   ""
         Sections(0).Cells(6).Style.BorderHI=   ""
         Sections(0).Cells(6).Style.BorderHB=   ""
         Sections(0).Cells(6).Style.BorderVL=   ""
         Sections(0).Cells(6).Style.BorderVI=   ""
         Sections(0).Cells(6).Style.BorderVR=   ""
         Sections(0).Cells(6).Style.NoClipping=   0   'False
         Sections(0).Cells(6).Style.RTF=   0   'False
         Sections(0).Cells(6).Style.fprops=   18087937
         Sections(0).Cells(7).Name=   "CELL_13"
         Sections(0).Cells(7).NewLine=   -1  'True
         Sections(0).Cells(7).Width=   30
         Sections(0).Cells(7).PrivateStyle=   -1  'True
         Sections(0).Cells(7).Style.Name=   "<private>"
         Sections(0).Cells(7).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(7).Style.Font_Name=   "Courier"
         Sections(0).Cells(7).Style.Font_Size=   9.75
         Sections(0).Cells(7).Style.Font_Bold=   0   'False
         Sections(0).Cells(7).Style.Font_Italic=   0   'False
         Sections(0).Cells(7).Style.Font_Underline=   0   'False
         Sections(0).Cells(7).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(7).Style.Font_Charset=   0
         Sections(0).Cells(7).Style.TextAlign=   3
         Sections(0).Cells(7).Style.TextVAlign=   1
         Sections(0).Cells(7).Style.TextWrap=   -1  'True
         Sections(0).Cells(7).Style.ForeColor=   0
         Sections(0).Cells(7).Style.BackColor=   16777215
         Sections(0).Cells(7).Style.NoFill=   -1  'True
         Sections(0).Cells(7).Style.BackPicFile=   ""
         Sections(0).Cells(7).Style.ForePicFile=   ""
         Sections(0).Cells(7).Style.BackPicVertPlacement=   0
         Sections(0).Cells(7).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(7).Style.ForePicPlacement=   0
         Sections(0).Cells(7).Style.ForePicDrawMode=   0
         Sections(0).Cells(7).Style.MarginLeft=   6
         Sections(0).Cells(7).Style.MarginTop=   1
         Sections(0).Cells(7).Style.MarginRight=   6
         Sections(0).Cells(7).Style.MarginBottom=   1
         Sections(0).Cells(7).Style.HasBorders=   -1  'True
         Sections(0).Cells(7).Style.BorderHT=   ""
         Sections(0).Cells(7).Style.BorderHI=   ""
         Sections(0).Cells(7).Style.BorderHB=   ""
         Sections(0).Cells(7).Style.BorderVL=   ""
         Sections(0).Cells(7).Style.BorderVI=   ""
         Sections(0).Cells(7).Style.BorderVR=   ""
         Sections(0).Cells(7).Style.NoClipping=   0   'False
         Sections(0).Cells(7).Style.RTF=   0   'False
         Sections(0).Cells(7).Style.fprops=   22413312
         Sections(0).Cells(8).Name=   "CELL_14"
         Sections(0).Cells(8).Exp=   """No. "" & cSE"
         Sections(0).Cells(8).PrivateStyle=   -1  'True
         Sections(0).Cells(8).Style.Name=   "<private>"
         Sections(0).Cells(8).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(8).Style.Font_Name=   "Courier"
         Sections(0).Cells(8).Style.Font_Size=   9.75
         Sections(0).Cells(8).Style.Font_Bold=   0   'False
         Sections(0).Cells(8).Style.Font_Italic=   0   'False
         Sections(0).Cells(8).Style.Font_Underline=   0   'False
         Sections(0).Cells(8).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(8).Style.Font_Charset=   0
         Sections(0).Cells(8).Style.TextAlign=   1
         Sections(0).Cells(8).Style.TextVAlign=   1
         Sections(0).Cells(8).Style.TextWrap=   -1  'True
         Sections(0).Cells(8).Style.ForeColor=   0
         Sections(0).Cells(8).Style.BackColor=   16777215
         Sections(0).Cells(8).Style.NoFill=   -1  'True
         Sections(0).Cells(8).Style.BackPicFile=   ""
         Sections(0).Cells(8).Style.ForePicFile=   ""
         Sections(0).Cells(8).Style.BackPicVertPlacement=   0
         Sections(0).Cells(8).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(8).Style.ForePicPlacement=   0
         Sections(0).Cells(8).Style.ForePicDrawMode=   0
         Sections(0).Cells(8).Style.MarginLeft=   6
         Sections(0).Cells(8).Style.MarginTop=   1
         Sections(0).Cells(8).Style.MarginRight=   6
         Sections(0).Cells(8).Style.MarginBottom=   1
         Sections(0).Cells(8).Style.HasBorders=   -1  'True
         Sections(0).Cells(8).Style.BorderHT=   ""
         Sections(0).Cells(8).Style.BorderHI=   ""
         Sections(0).Cells(8).Style.BorderHB=   ""
         Sections(0).Cells(8).Style.BorderVL=   ""
         Sections(0).Cells(8).Style.BorderVI=   ""
         Sections(0).Cells(8).Style.BorderVR=   ""
         Sections(0).Cells(8).Style.NoClipping=   0   'False
         Sections(0).Cells(8).Style.RTF=   0   'False
         Sections(0).Cells(8).Style.fprops=   16908289
         Sections(0).Cells(9).Name=   "CELL_15"
         Sections(0).Cells(9).Exp=   """Page "" & PageNo()"
         Sections(0).Cells(9).Width=   30
         Sections(0).Cells(9).PrivateStyle=   -1  'True
         Sections(0).Cells(9).Style.Name=   "<private>"
         Sections(0).Cells(9).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(9).Style.Font_Name=   "Courier"
         Sections(0).Cells(9).Style.Font_Size=   9.75
         Sections(0).Cells(9).Style.Font_Bold=   0   'False
         Sections(0).Cells(9).Style.Font_Italic=   0   'False
         Sections(0).Cells(9).Style.Font_Underline=   0   'False
         Sections(0).Cells(9).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(9).Style.Font_Charset=   0
         Sections(0).Cells(9).Style.TextAlign=   2
         Sections(0).Cells(9).Style.TextVAlign=   1
         Sections(0).Cells(9).Style.TextWrap=   -1  'True
         Sections(0).Cells(9).Style.ForeColor=   0
         Sections(0).Cells(9).Style.BackColor=   16777215
         Sections(0).Cells(9).Style.NoFill=   -1  'True
         Sections(0).Cells(9).Style.BackPicFile=   ""
         Sections(0).Cells(9).Style.ForePicFile=   ""
         Sections(0).Cells(9).Style.BackPicVertPlacement=   0
         Sections(0).Cells(9).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(9).Style.ForePicPlacement=   0
         Sections(0).Cells(9).Style.ForePicDrawMode=   0
         Sections(0).Cells(9).Style.MarginLeft=   6
         Sections(0).Cells(9).Style.MarginTop=   1
         Sections(0).Cells(9).Style.MarginRight=   6
         Sections(0).Cells(9).Style.MarginBottom=   1
         Sections(0).Cells(9).Style.HasBorders=   -1  'True
         Sections(0).Cells(9).Style.BorderHT=   ""
         Sections(0).Cells(9).Style.BorderHI=   ""
         Sections(0).Cells(9).Style.BorderHB=   ""
         Sections(0).Cells(9).Style.BorderVL=   ""
         Sections(0).Cells(9).Style.BorderVI=   ""
         Sections(0).Cells(9).Style.BorderVR=   ""
         Sections(0).Cells(9).Style.NoClipping=   0   'False
         Sections(0).Cells(9).Style.RTF=   0   'False
         Sections(0).Cells(9).Style.fprops=   17956865
         Sections(0).Cells(10).Name=   "CELL_4"
         Sections(0).Cells(10).NewLine=   -1  'True
         Sections(0).Cells(10).Height=   6
         Sections(0).Cells(10).AutoHeight=   0   'False
         Sections(0).Cells(10).PrivateStyle=   -1  'True
         Sections(0).Cells(10).Style.Name=   "<private>"
         Sections(0).Cells(10).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(10).Style.Font_Name=   "Courier"
         Sections(0).Cells(10).Style.Font_Size=   9.75
         Sections(0).Cells(10).Style.Font_Bold=   0   'False
         Sections(0).Cells(10).Style.Font_Italic=   0   'False
         Sections(0).Cells(10).Style.Font_Underline=   0   'False
         Sections(0).Cells(10).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(10).Style.Font_Charset=   0
         Sections(0).Cells(10).Style.TextAlign=   3
         Sections(0).Cells(10).Style.TextVAlign=   1
         Sections(0).Cells(10).Style.TextWrap=   -1  'True
         Sections(0).Cells(10).Style.ForeColor=   0
         Sections(0).Cells(10).Style.BackColor=   16777215
         Sections(0).Cells(10).Style.NoFill=   -1  'True
         Sections(0).Cells(10).Style.BackPicFile=   ""
         Sections(0).Cells(10).Style.ForePicFile=   ""
         Sections(0).Cells(10).Style.BackPicVertPlacement=   0
         Sections(0).Cells(10).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(10).Style.ForePicPlacement=   0
         Sections(0).Cells(10).Style.ForePicDrawMode=   0
         Sections(0).Cells(10).Style.MarginLeft=   6
         Sections(0).Cells(10).Style.MarginTop=   1
         Sections(0).Cells(10).Style.MarginRight=   6
         Sections(0).Cells(10).Style.MarginBottom=   1
         Sections(0).Cells(10).Style.HasBorders=   -1  'True
         Sections(0).Cells(10).Style.BorderHT=   ""
         Sections(0).Cells(10).Style.BorderHI=   ""
         Sections(0).Cells(10).Style.BorderHB=   ""
         Sections(0).Cells(10).Style.BorderVL=   ""
         Sections(0).Cells(10).Style.BorderVI=   ""
         Sections(0).Cells(10).Style.BorderVR=   ""
         Sections(0).Cells(10).Style.NoClipping=   0   'False
         Sections(0).Cells(10).Style.RTF=   0   'False
         Sections(0).Cells(10).Style.fprops=   18087936
         Sections(0).Cells(11).Name=   "CELL_17"
         Sections(0).Cells(11).NewLine=   -1  'True
         Sections(0).Cells(11).PrivateStyle=   -1  'True
         Sections(0).Cells(11).Style.Name=   "<private>"
         Sections(0).Cells(11).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(11).Style.Font_Name=   "Courier"
         Sections(0).Cells(11).Style.Font_Size=   9.75
         Sections(0).Cells(11).Style.Font_Bold=   0   'False
         Sections(0).Cells(11).Style.Font_Italic=   0   'False
         Sections(0).Cells(11).Style.Font_Underline=   0   'False
         Sections(0).Cells(11).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(11).Style.Font_Charset=   0
         Sections(0).Cells(11).Style.TextAlign=   3
         Sections(0).Cells(11).Style.TextVAlign=   1
         Sections(0).Cells(11).Style.TextWrap=   -1  'True
         Sections(0).Cells(11).Style.ForeColor=   0
         Sections(0).Cells(11).Style.BackColor=   16777215
         Sections(0).Cells(11).Style.NoFill=   -1  'True
         Sections(0).Cells(11).Style.BackPicFile=   ""
         Sections(0).Cells(11).Style.ForePicFile=   ""
         Sections(0).Cells(11).Style.BackPicVertPlacement=   0
         Sections(0).Cells(11).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(11).Style.ForePicPlacement=   0
         Sections(0).Cells(11).Style.ForePicDrawMode=   0
         Sections(0).Cells(11).Style.MarginLeft=   6
         Sections(0).Cells(11).Style.MarginTop=   1
         Sections(0).Cells(11).Style.MarginRight=   6
         Sections(0).Cells(11).Style.MarginBottom=   1
         Sections(0).Cells(11).Style.HasBorders=   -1  'True
         Sections(0).Cells(11).Style.BorderHT=   ""
         Sections(0).Cells(11).Style.BorderHI=   ""
         Sections(0).Cells(11).Style.BorderHB=   ""
         Sections(0).Cells(11).Style.BorderVL=   ""
         Sections(0).Cells(11).Style.BorderVI=   ""
         Sections(0).Cells(11).Style.BorderVR=   ""
         Sections(0).Cells(11).Style.NoClipping=   0   'False
         Sections(0).Cells(11).Style.RTF=   0   'False
         Sections(0).Cells(11).Style.fprops=   18087936
         Sections(0).Cells(12).Name=   "CELL_20"
         Sections(0).Cells(12).Exp=   """Print By : ""& cUserName"
         Sections(0).Cells(12).PrivateStyle=   -1  'True
         Sections(0).Cells(12).Style.Name=   "<private>"
         Sections(0).Cells(12).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(12).Style.Font_Name=   "Courier"
         Sections(0).Cells(12).Style.Font_Size=   9.75
         Sections(0).Cells(12).Style.Font_Bold=   0   'False
         Sections(0).Cells(12).Style.Font_Italic=   0   'False
         Sections(0).Cells(12).Style.Font_Underline=   0   'False
         Sections(0).Cells(12).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(12).Style.Font_Charset=   0
         Sections(0).Cells(12).Style.TextAlign=   2
         Sections(0).Cells(12).Style.TextVAlign=   1
         Sections(0).Cells(12).Style.TextWrap=   -1  'True
         Sections(0).Cells(12).Style.ForeColor=   0
         Sections(0).Cells(12).Style.BackColor=   16777215
         Sections(0).Cells(12).Style.NoFill=   -1  'True
         Sections(0).Cells(12).Style.BackPicFile=   ""
         Sections(0).Cells(12).Style.ForePicFile=   ""
         Sections(0).Cells(12).Style.BackPicVertPlacement=   0
         Sections(0).Cells(12).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(12).Style.ForePicPlacement=   0
         Sections(0).Cells(12).Style.ForePicDrawMode=   0
         Sections(0).Cells(12).Style.MarginLeft=   6
         Sections(0).Cells(12).Style.MarginTop=   1
         Sections(0).Cells(12).Style.MarginRight=   6
         Sections(0).Cells(12).Style.MarginBottom=   1
         Sections(0).Cells(12).Style.HasBorders=   -1  'True
         Sections(0).Cells(12).Style.BorderHT=   ""
         Sections(0).Cells(12).Style.BorderHI=   ""
         Sections(0).Cells(12).Style.BorderHB=   ""
         Sections(0).Cells(12).Style.BorderVL=   ""
         Sections(0).Cells(12).Style.BorderVI=   ""
         Sections(0).Cells(12).Style.BorderVR=   ""
         Sections(0).Cells(12).Style.NoClipping=   0   'False
         Sections(0).Cells(12).Style.RTF=   0   'False
         Sections(0).Cells(12).Style.fprops=   16777217
         Sections(0).Cells(13).Name=   "CELL_16"
         Sections(0).Cells(13).NewLine=   -1  'True
         Sections(0).Cells(13).PrivateStyle=   -1  'True
         Sections(0).Cells(13).Style.Name=   "<private>"
         Sections(0).Cells(13).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(13).Style.Font_Name=   "Courier"
         Sections(0).Cells(13).Style.Font_Size=   9.75
         Sections(0).Cells(13).Style.Font_Bold=   0   'False
         Sections(0).Cells(13).Style.Font_Italic=   0   'False
         Sections(0).Cells(13).Style.Font_Underline=   0   'False
         Sections(0).Cells(13).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(13).Style.Font_Charset=   0
         Sections(0).Cells(13).Style.TextAlign=   3
         Sections(0).Cells(13).Style.TextVAlign=   1
         Sections(0).Cells(13).Style.TextWrap=   -1  'True
         Sections(0).Cells(13).Style.ForeColor=   0
         Sections(0).Cells(13).Style.BackColor=   16777215
         Sections(0).Cells(13).Style.NoFill=   -1  'True
         Sections(0).Cells(13).Style.BackPicFile=   ""
         Sections(0).Cells(13).Style.ForePicFile=   ""
         Sections(0).Cells(13).Style.BackPicVertPlacement=   0
         Sections(0).Cells(13).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(13).Style.ForePicPlacement=   0
         Sections(0).Cells(13).Style.ForePicDrawMode=   0
         Sections(0).Cells(13).Style.MarginLeft=   6
         Sections(0).Cells(13).Style.MarginTop=   1
         Sections(0).Cells(13).Style.MarginRight=   6
         Sections(0).Cells(13).Style.MarginBottom=   1
         Sections(0).Cells(13).Style.HasBorders=   -1  'True
         Sections(0).Cells(13).Style.BorderHT=   ""
         Sections(0).Cells(13).Style.BorderHI=   ""
         Sections(0).Cells(13).Style.BorderHB=   ""
         Sections(0).Cells(13).Style.BorderVL=   ""
         Sections(0).Cells(13).Style.BorderVI=   ""
         Sections(0).Cells(13).Style.BorderVR=   ""
         Sections(0).Cells(13).Style.NoClipping=   0   'False
         Sections(0).Cells(13).Style.RTF=   0   'False
         Sections(0).Cells(13).Style.fprops=   18087936
         Sections(0).Cells(14).Name=   "CELL_18"
         Sections(0).Cells(14).PrivateStyle=   -1  'True
         Sections(0).Cells(14).Style.Name=   "<private>"
         Sections(0).Cells(14).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(14).Style.Font_Name=   "Courier"
         Sections(0).Cells(14).Style.Font_Size=   9.75
         Sections(0).Cells(14).Style.Font_Bold=   0   'False
         Sections(0).Cells(14).Style.Font_Italic=   0   'False
         Sections(0).Cells(14).Style.Font_Underline=   0   'False
         Sections(0).Cells(14).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(14).Style.Font_Charset=   0
         Sections(0).Cells(14).Style.TextAlign=   2
         Sections(0).Cells(14).Style.TextVAlign=   1
         Sections(0).Cells(14).Style.TextWrap=   -1  'True
         Sections(0).Cells(14).Style.ForeColor=   0
         Sections(0).Cells(14).Style.BackColor=   16777215
         Sections(0).Cells(14).Style.NoFill=   -1  'True
         Sections(0).Cells(14).Style.BackPicFile=   ""
         Sections(0).Cells(14).Style.ForePicFile=   ""
         Sections(0).Cells(14).Style.BackPicVertPlacement=   0
         Sections(0).Cells(14).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(14).Style.ForePicPlacement=   0
         Sections(0).Cells(14).Style.ForePicDrawMode=   0
         Sections(0).Cells(14).Style.MarginLeft=   6
         Sections(0).Cells(14).Style.MarginTop=   1
         Sections(0).Cells(14).Style.MarginRight=   6
         Sections(0).Cells(14).Style.MarginBottom=   1
         Sections(0).Cells(14).Style.HasBorders=   -1  'True
         Sections(0).Cells(14).Style.BorderHT=   ""
         Sections(0).Cells(14).Style.BorderHI=   ""
         Sections(0).Cells(14).Style.BorderHB=   ""
         Sections(0).Cells(14).Style.BorderVL=   ""
         Sections(0).Cells(14).Style.BorderVI=   ""
         Sections(0).Cells(14).Style.BorderVR=   ""
         Sections(0).Cells(14).Style.NoClipping=   0   'False
         Sections(0).Cells(14).Style.RTF=   0   'False
         Sections(0).Cells(14).Style.fprops=   17825793
         Sections(0).Cells(15).Name=   "CELL_19"
         Sections(0).Cells(15).Exp=   "Now"
         Sections(0).Cells(15).PrivateStyle=   -1  'True
         Sections(0).Cells(15).Format=   "dd-MM-yyyy HH:MM:SS"
         Sections(0).Cells(15).Style.Name=   "<private>"
         Sections(0).Cells(15).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(15).Style.Font_Name=   "Courier"
         Sections(0).Cells(15).Style.Font_Size=   9.75
         Sections(0).Cells(15).Style.Font_Bold=   0   'False
         Sections(0).Cells(15).Style.Font_Italic=   0   'False
         Sections(0).Cells(15).Style.Font_Underline=   0   'False
         Sections(0).Cells(15).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(15).Style.Font_Charset=   0
         Sections(0).Cells(15).Style.TextAlign=   2
         Sections(0).Cells(15).Style.TextVAlign=   1
         Sections(0).Cells(15).Style.TextWrap=   -1  'True
         Sections(0).Cells(15).Style.ForeColor=   0
         Sections(0).Cells(15).Style.BackColor=   16777215
         Sections(0).Cells(15).Style.NoFill=   -1  'True
         Sections(0).Cells(15).Style.BackPicFile=   ""
         Sections(0).Cells(15).Style.ForePicFile=   ""
         Sections(0).Cells(15).Style.BackPicVertPlacement=   0
         Sections(0).Cells(15).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(15).Style.ForePicPlacement=   0
         Sections(0).Cells(15).Style.ForePicDrawMode=   0
         Sections(0).Cells(15).Style.MarginLeft=   6
         Sections(0).Cells(15).Style.MarginTop=   1
         Sections(0).Cells(15).Style.MarginRight=   6
         Sections(0).Cells(15).Style.MarginBottom=   1
         Sections(0).Cells(15).Style.HasBorders=   -1  'True
         Sections(0).Cells(15).Style.BorderHT=   ""
         Sections(0).Cells(15).Style.BorderHI=   ""
         Sections(0).Cells(15).Style.BorderHB=   ""
         Sections(0).Cells(15).Style.BorderVL=   ""
         Sections(0).Cells(15).Style.BorderVI=   ""
         Sections(0).Cells(15).Style.BorderVR=   ""
         Sections(0).Cells(15).Style.NoClipping=   0   'False
         Sections(0).Cells(15).Style.RTF=   0   'False
         Sections(0).Cells(15).Style.fprops=   16777217
         Sections(1).Name=   "SECTION_3"
         Sections(1).StyleExp=   "'STYLE_1'"
         Sections(1).AutoHeight=   0   'False
         Sections(1).Height=   5
         Sections(1).dtopts=   2
         Sections(1).Cells.Count=   12
         Sections(1).Cells(0).Name=   "CELL_9"
         Sections(1).Cells(0).Exp=   """Sudah Terima uang dari : "" & cNama & cAlamat  & "" Kode Anggota : ""& cKodeAnggota & ""  """
         Sections(1).Cells(0).NewLine=   -1  'True
         Sections(1).Cells(0).Width=   100
         Sections(1).Cells(1).Name=   "CELL_10"
         Sections(1).Cells(1).Exp=   """Sebesar "" & cTerbilang"
         Sections(1).Cells(1).NewLine=   -1  'True
         Sections(1).Cells(1).PrivateStyle=   -1  'True
         Sections(1).Cells(1).Style.Name=   "<private>"
         Sections(1).Cells(1).Style.ParentName=   "STYLE_1"
         Sections(1).Cells(1).Style.Font_Name=   "Courier"
         Sections(1).Cells(1).Style.Font_Size=   9.75
         Sections(1).Cells(1).Style.Font_Bold=   -1  'True
         Sections(1).Cells(1).Style.Font_Italic=   -1  'True
         Sections(1).Cells(1).Style.Font_Underline=   0   'False
         Sections(1).Cells(1).Style.Font_Strikeout=   0   'False
         Sections(1).Cells(1).Style.Font_Charset=   0
         Sections(1).Cells(1).Style.TextAlign=   0
         Sections(1).Cells(1).Style.TextVAlign=   1
         Sections(1).Cells(1).Style.TextWrap=   -1  'True
         Sections(1).Cells(1).Style.ForeColor=   0
         Sections(1).Cells(1).Style.BackColor=   16777215
         Sections(1).Cells(1).Style.NoFill=   -1  'True
         Sections(1).Cells(1).Style.BackPicFile=   ""
         Sections(1).Cells(1).Style.ForePicFile=   ""
         Sections(1).Cells(1).Style.BackPicVertPlacement=   0
         Sections(1).Cells(1).Style.BackPicHorzPlacement=   0
         Sections(1).Cells(1).Style.ForePicPlacement=   0
         Sections(1).Cells(1).Style.ForePicDrawMode=   0
         Sections(1).Cells(1).Style.MarginLeft=   6
         Sections(1).Cells(1).Style.MarginTop=   1
         Sections(1).Cells(1).Style.MarginRight=   6
         Sections(1).Cells(1).Style.MarginBottom=   1
         Sections(1).Cells(1).Style.HasBorders=   -1  'True
         Sections(1).Cells(1).Style.BorderHT=   ""
         Sections(1).Cells(1).Style.BorderHI=   ""
         Sections(1).Cells(1).Style.BorderHB=   ""
         Sections(1).Cells(1).Style.BorderVL=   ""
         Sections(1).Cells(1).Style.BorderVI=   ""
         Sections(1).Cells(1).Style.BorderVR=   ""
         Sections(1).Cells(1).Style.NoClipping=   0   'False
         Sections(1).Cells(1).Style.RTF=   0   'False
         Sections(1).Cells(1).Style.fprops=   50331649
         Sections(1).Cells(2).Name=   "CELL_11"
         Sections(1).Cells(2).Exp=   """Sebagai Bukti Pelunasan Hutang/Piutang/Bon - Mohon disimpan dengan baik"""
         Sections(1).Cells(2).NewLine=   -1  'True
         Sections(1).Cells(3).Name=   "CELL_3"
         Sections(1).Cells(3).Exp=   """"""
         Sections(1).Cells(3).NewLine=   -1  'True
         Sections(1).Cells(4).Name=   "CELL_0"
         Sections(1).Cells(4).Exp=   """                 Kasir"""
         Sections(1).Cells(4).NewLine=   -1  'True
         Sections(1).Cells(4).PrivateStyle=   -1  'True
         Sections(1).Cells(4).Style.Name=   "<private>"
         Sections(1).Cells(4).Style.ParentName=   "STYLE_1"
         Sections(1).Cells(4).Style.Font_Name=   "Courier"
         Sections(1).Cells(4).Style.Font_Size=   9.75
         Sections(1).Cells(4).Style.Font_Bold=   0   'False
         Sections(1).Cells(4).Style.Font_Italic=   0   'False
         Sections(1).Cells(4).Style.Font_Underline=   0   'False
         Sections(1).Cells(4).Style.Font_Strikeout=   0   'False
         Sections(1).Cells(4).Style.Font_Charset=   0
         Sections(1).Cells(4).Style.TextAlign=   3
         Sections(1).Cells(4).Style.TextVAlign=   1
         Sections(1).Cells(4).Style.TextWrap=   -1  'True
         Sections(1).Cells(4).Style.ForeColor=   0
         Sections(1).Cells(4).Style.BackColor=   16777215
         Sections(1).Cells(4).Style.NoFill=   -1  'True
         Sections(1).Cells(4).Style.BackPicFile=   ""
         Sections(1).Cells(4).Style.ForePicFile=   ""
         Sections(1).Cells(4).Style.BackPicVertPlacement=   0
         Sections(1).Cells(4).Style.BackPicHorzPlacement=   0
         Sections(1).Cells(4).Style.ForePicPlacement=   0
         Sections(1).Cells(4).Style.ForePicDrawMode=   0
         Sections(1).Cells(4).Style.MarginLeft=   6
         Sections(1).Cells(4).Style.MarginTop=   1
         Sections(1).Cells(4).Style.MarginRight=   6
         Sections(1).Cells(4).Style.MarginBottom=   1
         Sections(1).Cells(4).Style.HasBorders=   -1  'True
         Sections(1).Cells(4).Style.BorderHT=   ""
         Sections(1).Cells(4).Style.BorderHI=   ""
         Sections(1).Cells(4).Style.BorderHB=   ""
         Sections(1).Cells(4).Style.BorderVL=   ""
         Sections(1).Cells(4).Style.BorderVI=   ""
         Sections(1).Cells(4).Style.BorderVR=   ""
         Sections(1).Cells(4).Style.NoClipping=   0   'False
         Sections(1).Cells(4).Style.RTF=   0   'False
         Sections(1).Cells(4).Style.fprops=   294912
         Sections(1).Cells(5).Name=   "CELL_15"
         Sections(1).Cells(5).Exp=   """                           """
         Sections(1).Cells(5).PrivateStyle=   -1  'True
         Sections(1).Cells(5).Style.Name=   "<private>"
         Sections(1).Cells(5).Style.ParentName=   "STYLE_1"
         Sections(1).Cells(5).Style.Font_Name=   "Courier"
         Sections(1).Cells(5).Style.Font_Size=   9.75
         Sections(1).Cells(5).Style.Font_Bold=   0   'False
         Sections(1).Cells(5).Style.Font_Italic=   0   'False
         Sections(1).Cells(5).Style.Font_Underline=   0   'False
         Sections(1).Cells(5).Style.Font_Strikeout=   0   'False
         Sections(1).Cells(5).Style.Font_Charset=   0
         Sections(1).Cells(5).Style.TextAlign=   3
         Sections(1).Cells(5).Style.TextVAlign=   1
         Sections(1).Cells(5).Style.TextWrap=   -1  'True
         Sections(1).Cells(5).Style.ForeColor=   0
         Sections(1).Cells(5).Style.BackColor=   16777215
         Sections(1).Cells(5).Style.NoFill=   -1  'True
         Sections(1).Cells(5).Style.BackPicFile=   ""
         Sections(1).Cells(5).Style.ForePicFile=   ""
         Sections(1).Cells(5).Style.BackPicVertPlacement=   0
         Sections(1).Cells(5).Style.BackPicHorzPlacement=   0
         Sections(1).Cells(5).Style.ForePicPlacement=   0
         Sections(1).Cells(5).Style.ForePicDrawMode=   0
         Sections(1).Cells(5).Style.MarginLeft=   6
         Sections(1).Cells(5).Style.MarginTop=   1
         Sections(1).Cells(5).Style.MarginRight=   6
         Sections(1).Cells(5).Style.MarginBottom=   1
         Sections(1).Cells(5).Style.HasBorders=   -1  'True
         Sections(1).Cells(5).Style.BorderHT=   ""
         Sections(1).Cells(5).Style.BorderHI=   ""
         Sections(1).Cells(5).Style.BorderHB=   ""
         Sections(1).Cells(5).Style.BorderVL=   ""
         Sections(1).Cells(5).Style.BorderVI=   ""
         Sections(1).Cells(5).Style.BorderVR=   ""
         Sections(1).Cells(5).Style.NoClipping=   0   'False
         Sections(1).Cells(5).Style.RTF=   0   'False
         Sections(1).Cells(5).Style.fprops=   294912
         Sections(1).Cells(6).Name=   "CELL_1"
         Sections(1).Cells(6).Exp=   """Total : """
         Sections(1).Cells(6).Width=   14
         Sections(1).Cells(6).PrivateStyle=   -1  'True
         Sections(1).Cells(6).Style.Name=   "<private>"
         Sections(1).Cells(6).Style.ParentName=   "STYLE_1"
         Sections(1).Cells(6).Style.Font_Name=   "Courier"
         Sections(1).Cells(6).Style.Font_Size=   9.75
         Sections(1).Cells(6).Style.Font_Bold=   0   'False
         Sections(1).Cells(6).Style.Font_Italic=   0   'False
         Sections(1).Cells(6).Style.Font_Underline=   0   'False
         Sections(1).Cells(6).Style.Font_Strikeout=   0   'False
         Sections(1).Cells(6).Style.Font_Charset=   0
         Sections(1).Cells(6).Style.TextAlign=   2
         Sections(1).Cells(6).Style.TextVAlign=   1
         Sections(1).Cells(6).Style.TextWrap=   -1  'True
         Sections(1).Cells(6).Style.ForeColor=   0
         Sections(1).Cells(6).Style.BackColor=   16777215
         Sections(1).Cells(6).Style.NoFill=   -1  'True
         Sections(1).Cells(6).Style.BackPicFile=   ""
         Sections(1).Cells(6).Style.ForePicFile=   ""
         Sections(1).Cells(6).Style.BackPicVertPlacement=   0
         Sections(1).Cells(6).Style.BackPicHorzPlacement=   0
         Sections(1).Cells(6).Style.ForePicPlacement=   0
         Sections(1).Cells(6).Style.ForePicDrawMode=   0
         Sections(1).Cells(6).Style.MarginLeft=   6
         Sections(1).Cells(6).Style.MarginTop=   1
         Sections(1).Cells(6).Style.MarginRight=   6
         Sections(1).Cells(6).Style.MarginBottom=   1
         Sections(1).Cells(6).Style.HasBorders=   -1  'True
         Sections(1).Cells(6).Style.BorderHT=   ""
         Sections(1).Cells(6).Style.BorderHI=   ""
         Sections(1).Cells(6).Style.BorderHB=   ""
         Sections(1).Cells(6).Style.BorderVL=   ""
         Sections(1).Cells(6).Style.BorderVI=   ""
         Sections(1).Cells(6).Style.BorderVR=   ""
         Sections(1).Cells(6).Style.NoClipping=   0   'False
         Sections(1).Cells(6).Style.RTF=   0   'False
         Sections(1).Cells(6).Style.fprops=   32769
         Sections(1).Cells(7).Name=   "CELL_2"
         Sections(1).Cells(7).Exp=   "nSubTotal"
         Sections(1).Cells(7).Width=   15
         Sections(1).Cells(7).PrivateStyle=   -1  'True
         Sections(1).Cells(7).Format=   "###,###,###,###,###,##0.00"
         Sections(1).Cells(7).Style.Name=   "<private>"
         Sections(1).Cells(7).Style.ParentName=   "STYLE_1"
         Sections(1).Cells(7).Style.Font_Name=   "Courier"
         Sections(1).Cells(7).Style.Font_Size=   9.75
         Sections(1).Cells(7).Style.Font_Bold=   0   'False
         Sections(1).Cells(7).Style.Font_Italic=   0   'False
         Sections(1).Cells(7).Style.Font_Underline=   0   'False
         Sections(1).Cells(7).Style.Font_Strikeout=   0   'False
         Sections(1).Cells(7).Style.Font_Charset=   0
         Sections(1).Cells(7).Style.TextAlign=   3
         Sections(1).Cells(7).Style.TextVAlign=   1
         Sections(1).Cells(7).Style.TextWrap=   -1  'True
         Sections(1).Cells(7).Style.ForeColor=   0
         Sections(1).Cells(7).Style.BackColor=   16777215
         Sections(1).Cells(7).Style.NoFill=   -1  'True
         Sections(1).Cells(7).Style.BackPicFile=   ""
         Sections(1).Cells(7).Style.ForePicFile=   ""
         Sections(1).Cells(7).Style.BackPicVertPlacement=   0
         Sections(1).Cells(7).Style.BackPicHorzPlacement=   0
         Sections(1).Cells(7).Style.ForePicPlacement=   0
         Sections(1).Cells(7).Style.ForePicDrawMode=   0
         Sections(1).Cells(7).Style.MarginLeft=   6
         Sections(1).Cells(7).Style.MarginTop=   1
         Sections(1).Cells(7).Style.MarginRight=   6
         Sections(1).Cells(7).Style.MarginBottom=   1
         Sections(1).Cells(7).Style.HasBorders=   -1  'True
         Sections(1).Cells(7).Style.BorderHT=   ""
         Sections(1).Cells(7).Style.BorderHI=   ""
         Sections(1).Cells(7).Style.BorderHB=   ""
         Sections(1).Cells(7).Style.BorderVL=   ""
         Sections(1).Cells(7).Style.BorderVI=   ""
         Sections(1).Cells(7).Style.BorderVR=   ""
         Sections(1).Cells(7).Style.NoClipping=   0   'False
         Sections(1).Cells(7).Style.RTF=   0   'False
         Sections(1).Cells(7).Style.fprops=   1081344
         Sections(1).Cells(8).Name=   "CELL_4"
         Sections(1).Cells(8).Exp=   "keAkun"
         Sections(1).Cells(8).NewLine=   -1  'True
         Sections(1).Cells(8).PrivateStyle=   -1  'True
         Sections(1).Cells(8).Style.Name=   "<private>"
         Sections(1).Cells(8).Style.ParentName=   "STYLE_1"
         Sections(1).Cells(8).Style.Font_Name=   "Courier"
         Sections(1).Cells(8).Style.Font_Size=   9.75
         Sections(1).Cells(8).Style.Font_Bold=   0   'False
         Sections(1).Cells(8).Style.Font_Italic=   0   'False
         Sections(1).Cells(8).Style.Font_Underline=   0   'False
         Sections(1).Cells(8).Style.Font_Strikeout=   0   'False
         Sections(1).Cells(8).Style.Font_Charset=   0
         Sections(1).Cells(8).Style.TextAlign=   2
         Sections(1).Cells(8).Style.TextVAlign=   1
         Sections(1).Cells(8).Style.TextWrap=   -1  'True
         Sections(1).Cells(8).Style.ForeColor=   0
         Sections(1).Cells(8).Style.BackColor=   16777215
         Sections(1).Cells(8).Style.NoFill=   -1  'True
         Sections(1).Cells(8).Style.BackPicFile=   ""
         Sections(1).Cells(8).Style.ForePicFile=   ""
         Sections(1).Cells(8).Style.BackPicVertPlacement=   0
         Sections(1).Cells(8).Style.BackPicHorzPlacement=   0
         Sections(1).Cells(8).Style.ForePicPlacement=   0
         Sections(1).Cells(8).Style.ForePicDrawMode=   0
         Sections(1).Cells(8).Style.MarginLeft=   6
         Sections(1).Cells(8).Style.MarginTop=   1
         Sections(1).Cells(8).Style.MarginRight=   6
         Sections(1).Cells(8).Style.MarginBottom=   1
         Sections(1).Cells(8).Style.HasBorders=   -1  'True
         Sections(1).Cells(8).Style.BorderHT=   ""
         Sections(1).Cells(8).Style.BorderHI=   ""
         Sections(1).Cells(8).Style.BorderHB=   ""
         Sections(1).Cells(8).Style.BorderVL=   ""
         Sections(1).Cells(8).Style.BorderVI=   ""
         Sections(1).Cells(8).Style.BorderVR=   ""
         Sections(1).Cells(8).Style.NoClipping=   0   'False
         Sections(1).Cells(8).Style.RTF=   0   'False
         Sections(1).Cells(8).Style.fprops=   1
         Sections(1).Cells(9).Name=   "CELL_5"
         Sections(1).Cells(9).NewLine=   -1  'True
         Sections(1).Cells(10).Name=   "CELL_16"
         Sections(1).Cells(10).NewLine=   -1  'True
         Sections(1).Cells(10).Height=   5
         Sections(1).Cells(10).PrivateStyle=   -1  'True
         Sections(1).Cells(10).Style.Name=   "<private>"
         Sections(1).Cells(10).Style.ParentName=   "STYLE_1"
         Sections(1).Cells(10).Style.Font_Name=   "Courier"
         Sections(1).Cells(10).Style.Font_Size=   9.75
         Sections(1).Cells(10).Style.Font_Bold=   0   'False
         Sections(1).Cells(10).Style.Font_Italic=   0   'False
         Sections(1).Cells(10).Style.Font_Underline=   0   'False
         Sections(1).Cells(10).Style.Font_Strikeout=   0   'False
         Sections(1).Cells(10).Style.Font_Charset=   0
         Sections(1).Cells(10).Style.TextAlign=   2
         Sections(1).Cells(10).Style.TextVAlign=   1
         Sections(1).Cells(10).Style.TextWrap=   -1  'True
         Sections(1).Cells(10).Style.ForeColor=   0
         Sections(1).Cells(10).Style.BackColor=   16777215
         Sections(1).Cells(10).Style.NoFill=   -1  'True
         Sections(1).Cells(10).Style.BackPicFile=   ""
         Sections(1).Cells(10).Style.ForePicFile=   ""
         Sections(1).Cells(10).Style.BackPicVertPlacement=   0
         Sections(1).Cells(10).Style.BackPicHorzPlacement=   0
         Sections(1).Cells(10).Style.ForePicPlacement=   0
         Sections(1).Cells(10).Style.ForePicDrawMode=   0
         Sections(1).Cells(10).Style.MarginLeft=   6
         Sections(1).Cells(10).Style.MarginTop=   1
         Sections(1).Cells(10).Style.MarginRight=   6
         Sections(1).Cells(10).Style.MarginBottom=   1
         Sections(1).Cells(10).Style.HasBorders=   -1  'True
         Sections(1).Cells(10).Style.BorderHT=   ""
         Sections(1).Cells(10).Style.BorderHI=   ""
         Sections(1).Cells(10).Style.BorderHB=   ""
         Sections(1).Cells(10).Style.BorderVL=   ""
         Sections(1).Cells(10).Style.BorderVI=   ""
         Sections(1).Cells(10).Style.BorderVR=   ""
         Sections(1).Cells(10).Style.NoClipping=   0   'False
         Sections(1).Cells(10).Style.RTF=   0   'False
         Sections(1).Cells(10).Style.fprops=   2064389
         Sections(1).Cells(11).Name=   "CELL_20"
         Sections(1).Cells(11).Exp=   "cFooter2"
         Sections(1).Cells(11).NewLine=   -1  'True
         Sections(1).Cells(11).PrivateStyle=   -1  'True
         Sections(1).Cells(11).Style.Name=   "<private>"
         Sections(1).Cells(11).Style.ParentName=   "STYLE_1"
         Sections(1).Cells(11).Style.Font_Name=   "Courier"
         Sections(1).Cells(11).Style.Font_Size=   9.75
         Sections(1).Cells(11).Style.Font_Bold=   0   'False
         Sections(1).Cells(11).Style.Font_Italic=   0   'False
         Sections(1).Cells(11).Style.Font_Underline=   0   'False
         Sections(1).Cells(11).Style.Font_Strikeout=   0   'False
         Sections(1).Cells(11).Style.Font_Charset=   0
         Sections(1).Cells(11).Style.TextAlign=   2
         Sections(1).Cells(11).Style.TextVAlign=   1
         Sections(1).Cells(11).Style.TextWrap=   -1  'True
         Sections(1).Cells(11).Style.ForeColor=   0
         Sections(1).Cells(11).Style.BackColor=   16777215
         Sections(1).Cells(11).Style.NoFill=   -1  'True
         Sections(1).Cells(11).Style.BackPicFile=   ""
         Sections(1).Cells(11).Style.ForePicFile=   ""
         Sections(1).Cells(11).Style.BackPicVertPlacement=   0
         Sections(1).Cells(11).Style.BackPicHorzPlacement=   0
         Sections(1).Cells(11).Style.ForePicPlacement=   0
         Sections(1).Cells(11).Style.ForePicDrawMode=   0
         Sections(1).Cells(11).Style.MarginLeft=   6
         Sections(1).Cells(11).Style.MarginTop=   1
         Sections(1).Cells(11).Style.MarginRight=   6
         Sections(1).Cells(11).Style.MarginBottom=   1
         Sections(1).Cells(11).Style.HasBorders=   -1  'True
         Sections(1).Cells(11).Style.BorderHT=   ""
         Sections(1).Cells(11).Style.BorderHI=   ""
         Sections(1).Cells(11).Style.BorderHB=   ""
         Sections(1).Cells(11).Style.BorderVL=   ""
         Sections(1).Cells(11).Style.BorderVI=   ""
         Sections(1).Cells(11).Style.BorderVR=   ""
         Sections(1).Cells(11).Style.NoClipping=   0   'False
         Sections(1).Cells(11).Style.RTF=   0   'False
         Sections(1).Cells(11).Style.fprops=   2064385
         Styles.Count    =   6
         Styles(0).Name  =   "Tdb_Base"
         Styles(0).ParentName=   ""
         Styles(0).Font_Name=   "Courier"
         Styles(0).Font_Size=   9.75
         Styles(0).Font_Bold=   -1  'True
         Styles(0).Font_Charset=   0
         Styles(0).TextVAlign=   1
         Styles(0).MarginTop=   1
         Styles(0).MarginBottom=   1
         Styles(1).Name  =   "STYLE_1"
         Styles(1).ParentName=   "Tdb_Base"
         Styles(1).Font_Name=   "Courier"
         Styles(1).Font_Size=   9.75
         Styles(1).Font_Charset=   0
         Styles(1).TextVAlign=   1
         Styles(1).MarginTop=   1
         Styles(1).MarginBottom=   1
         Styles(1).fprops=   18087936
         Styles(2).Name  =   "Tdb_Body"
         Styles(2).ParentName=   "Tdb_Base"
         Styles(2).Font_Name=   "Courier"
         Styles(2).Font_Size=   9.75
         Styles(2).Font_Charset=   0
         Styles(2).TextVAlign=   1
         Styles(2).MarginTop=   0
         Styles(2).MarginBottom=   0
         Styles(2).fprops=   18862080
         Styles(3).Name  =   "Tdb_Header"
         Styles(3).ParentName=   "Tdb_Base"
         Styles(3).Font_Name=   "Courier"
         Styles(3).Font_Size=   9.75
         Styles(3).Font_Bold=   -1  'True
         Styles(3).Font_Charset=   0
         Styles(3).TextAlign=   0
         Styles(3).TextVAlign=   1
         Styles(3).MarginTop=   1
         Styles(3).MarginBottom=   1
         Styles(3).BorderHT=   "Single"
         Styles(3).BorderHI=   "Single"
         Styles(3).BorderHB=   "Single"
         Styles(3).fprops=   2064385
         Styles(4).Name  =   "Tdb_PageFooter"
         Styles(4).ParentName=   "Tdb_Base"
         Styles(4).Font_Name=   "Courier"
         Styles(4).Font_Size=   9.75
         Styles(4).Font_Bold=   -1  'True
         Styles(4).Font_Charset=   0
         Styles(4).TextVAlign=   1
         Styles(4).MarginTop=   1
         Styles(4).MarginBottom=   1
         Styles(4).BorderHT=   "Single"
         Styles(4).fprops=   163840
         Styles(5).Name  =   "Garis"
         Styles(5).ParentName=   "Tdb_Base"
         Styles(5).Font_Name=   "Courier"
         Styles(5).Font_Size=   9.75
         Styles(5).Font_Bold=   -1  'True
         Styles(5).Font_Charset=   0
         Styles(5).TextAlign=   2
         Styles(5).TextVAlign=   1
         Styles(5).MarginTop=   1
         Styles(5).MarginBottom=   1
         Styles(5).BorderHT=   "Single"
         Styles(5).fprops=   32769
         Lines.Count     =   4
         Lines(0).Name   =   "Single"
         Lines(0).Thickness=   4
         Lines(1).Name   =   "Double"
         Lines(1).Thickness=   5
         Lines(2).Name   =   "Quarter"
         Lines(2).Thickness=   1
         Lines(2).Color  =   8421504
         Lines(3).Name   =   "None"
         Profiles.Count  =   1
         Profiles(0).Name=   "PROFILE_0"
         Profiles(0).Active=   -1  'True
         Profiles(0).PreviewNoMinimize=   -1  'True
         Profiles(0).PreviewNoMaximize=   -1  'True
         Profiles(0).PreviewNoResize=   -1  'True
         Profiles(0).PreviewMaximized=   -1  'True
         Profiles(0).PreviewNoSaveLoad=   -1  'True
         Profiles(0).PrinterMarginLeft=   10
         Profiles(0).PrinterMarginTop=   5
         Profiles(0).PrinterMarginRight=   10
         Profiles(0).PrinterMarginBottom=   5
         Profiles(0).PrinterPaperSize=   256
         Profiles(0).PrinterPaperHeight=   139
         Profiles(0).PrinterPaperWidth=   215
         Profiles(0).PrinterMargins_set=   -1  'True
         Profiles(0).PrinterPaperSize_set=   -1  'True
         Profiles(0).PrinterPaperUserSize_set=   -1  'True
      End
      Begin TrueDBReports60Ctl.TDBReports rptKuitansiLunas2 
         Height          =   570
         Left            =   270
         TabIndex        =   25
         Top             =   690
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1005
         Caption         =   "Kuitansi Lunas2"
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         ErrorMsgCaption =   ""
         Filtered        =   0   'False
         DataMode        =   1
         DataMember      =   ""
         LinkSequence    =   1
         LinkOrder       =   0
         NameSubstitute  =   ""
         ConnectionString=   "DSN=MySalemba"
         ConnectStringType=   3
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "MySalemba"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         CursorLocation  =   3
         ConnectionTimeout=   15
         CommandTimeout  =   30
         RecordSource    =   ""
         CursorType      =   1
         CommandType     =   8
         MaxRecords      =   0
         LinkType        =   0
         Master          =   ""
         CallDataRead    =   0   'False
         ConvertNullToEmpty=   -1  'True
         DesignConnection=   -1  'True
         DesignTimeout   =   5
         UnitsOfMeasurement=   4
         Vedit_ShowGrid  =   -1  'True
         Vedit_SnapToGrid=   0   'False
         Vedit_GridUnitWidth=   2.822
         Vedit_GridUnitHeight=   2.822
         Vedit_ShowCellExpressions=   -1  'True
         Norm_rect_left  =   0
         Norm_rect_top   =   0
         Norm_rect_right =   0
         Norm_rect_bottom=   0
         Virgin          =   0   'False
         Parameters.Count=   29
         Parameters(0).Name=   "cSE"
         Parameters(0).ValueExpression=   """"""
         Parameters(1).Name=   "cNama"
         Parameters(1).ValueExpression=   """"""
         Parameters(2).Name=   "cAlamat"
         Parameters(2).ValueExpression=   """"""
         Parameters(3).Name=   "cKota"
         Parameters(3).ValueExpression=   """"""
         Parameters(4).Name=   "cTerbilang"
         Parameters(4).ValueExpression=   """"""
         Parameters(5).Name=   "dTgl"
         Parameters(6).Name=   "dJTHTMP"
         Parameters(6).Type=   7
         Parameters(7).Name=   "cTTD"
         Parameters(8).Name=   "nSubTotal"
         Parameters(8).Type=   5
         Parameters(8).ValueExpression=   "0"
         Parameters(9).Name=   "nTotal"
         Parameters(9).Type=   5
         Parameters(9).ValueExpression=   "0"
         Parameters(10).Name=   "nPPn"
         Parameters(10).ValueExpression=   "0"
         Parameters(11).Name=   "nPajak"
         Parameters(11).Type=   5
         Parameters(11).ValueExpression=   "0"
         Parameters(12).Name=   "cNamaPerusahaan"
         Parameters(13).Name=   "cAlamatPerusahaan"
         Parameters(14).Name=   "cTeleponPerusahaan"
         Parameters(15).Name=   "cReceived"
         Parameters(16).Name=   "cKetReceived"
         Parameters(17).Name=   "cRef"
         Parameters(18).Name=   "cPerusahaanLine"
         Parameters(19).Name=   "cPayment"
         Parameters(20).Name=   "cUserName"
         Parameters(21).Name=   "nDiscount"
         Parameters(21).Type=   5
         Parameters(22).Name=   "cJudul"
         Parameters(23).Name=   "cSales"
         Parameters(24).Name=   "cFooter"
         Parameters(25).Name=   "nDp"
         Parameters(26).Name=   "cFooter2"
         Parameters(27).Name=   "cKodeAnggota"
         Parameters(28).Name=   "keAkun"
         Fields.Count    =   5
         Fields(0).Name  =   "Nomor"
         Fields(0).DisplayName=   "Nomor"
         Fields(0).Type  =   2
         Fields(1).Name  =   "NoInvoice"
         Fields(1).DisplayName=   "NoInvoice"
         Fields(2).Name  =   "Total"
         Fields(2).DisplayName=   "Total"
         Fields(2).Type  =   5
         Fields(3).Name  =   "Kredit"
         Fields(3).DisplayName=   "Kredit"
         Fields(4).Name  =   "JatuhTempo"
         Fields(4).DisplayName=   "JatuhTempo"
         Fields(4).Type  =   7
         Sections.Count  =   6
         Sections(0).Name=   "SECTION_2"
         Sections(0).Type=   1
         Sections(0).StyleExp=   "'Tdb_Base'"
         Sections(0).Cells.Count=   16
         Sections(0).Cells(0).Name=   "CELL_22"
         Sections(0).Cells(0).Exp=   "cNamaPerusahaan"
         Sections(0).Cells(0).NewLine=   -1  'True
         Sections(0).Cells(0).PrivateStyle=   -1  'True
         Sections(0).Cells(0).Style.Name=   "<private>"
         Sections(0).Cells(0).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(0).Style.Font_Name=   "Courier"
         Sections(0).Cells(0).Style.Font_Size=   12
         Sections(0).Cells(0).Style.Font_Bold=   -1  'True
         Sections(0).Cells(0).Style.Font_Italic=   0   'False
         Sections(0).Cells(0).Style.Font_Underline=   0   'False
         Sections(0).Cells(0).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(0).Style.Font_Charset=   0
         Sections(0).Cells(0).Style.TextAlign=   0
         Sections(0).Cells(0).Style.TextVAlign=   1
         Sections(0).Cells(0).Style.TextWrap=   -1  'True
         Sections(0).Cells(0).Style.ForeColor=   0
         Sections(0).Cells(0).Style.BackColor=   16777215
         Sections(0).Cells(0).Style.NoFill=   -1  'True
         Sections(0).Cells(0).Style.BackPicFile=   ""
         Sections(0).Cells(0).Style.ForePicFile=   ""
         Sections(0).Cells(0).Style.BackPicVertPlacement=   0
         Sections(0).Cells(0).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(0).Style.ForePicPlacement=   0
         Sections(0).Cells(0).Style.ForePicDrawMode=   0
         Sections(0).Cells(0).Style.MarginLeft=   6
         Sections(0).Cells(0).Style.MarginTop=   1
         Sections(0).Cells(0).Style.MarginRight=   6
         Sections(0).Cells(0).Style.MarginBottom=   1
         Sections(0).Cells(0).Style.HasBorders=   -1  'True
         Sections(0).Cells(0).Style.BorderHT=   ""
         Sections(0).Cells(0).Style.BorderHI=   ""
         Sections(0).Cells(0).Style.BorderHB=   ""
         Sections(0).Cells(0).Style.BorderVL=   ""
         Sections(0).Cells(0).Style.BorderVI=   ""
         Sections(0).Cells(0).Style.BorderVR=   ""
         Sections(0).Cells(0).Style.NoClipping=   0   'False
         Sections(0).Cells(0).Style.RTF=   0   'False
         Sections(0).Cells(0).Style.fprops=   89391105
         Sections(0).Cells(1).Name=   "CELL_25"
         Sections(0).Cells(1).Exp=   "cAlamatPerusahaan"
         Sections(0).Cells(1).NewLine=   -1  'True
         Sections(0).Cells(1).Height=   5
         Sections(0).Cells(1).AutoHeight=   0   'False
         Sections(0).Cells(1).PrivateStyle=   -1  'True
         Sections(0).Cells(1).Style.Name=   "<private>"
         Sections(0).Cells(1).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(1).Style.Font_Name=   "Courier"
         Sections(0).Cells(1).Style.Font_Size=   9.75
         Sections(0).Cells(1).Style.Font_Bold=   0   'False
         Sections(0).Cells(1).Style.Font_Italic=   0   'False
         Sections(0).Cells(1).Style.Font_Underline=   0   'False
         Sections(0).Cells(1).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(1).Style.Font_Charset=   0
         Sections(0).Cells(1).Style.TextAlign=   0
         Sections(0).Cells(1).Style.TextVAlign=   1
         Sections(0).Cells(1).Style.TextWrap=   -1  'True
         Sections(0).Cells(1).Style.ForeColor=   0
         Sections(0).Cells(1).Style.BackColor=   16777215
         Sections(0).Cells(1).Style.NoFill=   -1  'True
         Sections(0).Cells(1).Style.BackPicFile=   ""
         Sections(0).Cells(1).Style.ForePicFile=   ""
         Sections(0).Cells(1).Style.BackPicVertPlacement=   0
         Sections(0).Cells(1).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(1).Style.ForePicPlacement=   0
         Sections(0).Cells(1).Style.ForePicDrawMode=   0
         Sections(0).Cells(1).Style.MarginLeft=   6
         Sections(0).Cells(1).Style.MarginTop=   1
         Sections(0).Cells(1).Style.MarginRight=   6
         Sections(0).Cells(1).Style.MarginBottom=   1
         Sections(0).Cells(1).Style.HasBorders=   0   'False
         Sections(0).Cells(1).Style.BorderHT=   ""
         Sections(0).Cells(1).Style.BorderHI=   ""
         Sections(0).Cells(1).Style.BorderHB=   ""
         Sections(0).Cells(1).Style.BorderVL=   ""
         Sections(0).Cells(1).Style.BorderVI=   ""
         Sections(0).Cells(1).Style.BorderVR=   ""
         Sections(0).Cells(1).Style.NoClipping=   0   'False
         Sections(0).Cells(1).Style.RTF=   0   'False
         Sections(0).Cells(1).Style.fprops=   22413313
         Sections(0).Cells(2).Name=   "CELL_2"
         Sections(0).Cells(2).Exp=   """"""
         Sections(0).Cells(2).NewLine=   -1  'True
         Sections(0).Cells(2).PrivateStyle=   -1  'True
         Sections(0).Cells(2).Style.Name=   "<private>"
         Sections(0).Cells(2).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(2).Style.Font_Name=   "Courier"
         Sections(0).Cells(2).Style.Font_Size=   9.75
         Sections(0).Cells(2).Style.Font_Bold=   -1  'True
         Sections(0).Cells(2).Style.Font_Italic=   0   'False
         Sections(0).Cells(2).Style.Font_Underline=   0   'False
         Sections(0).Cells(2).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(2).Style.Font_Charset=   0
         Sections(0).Cells(2).Style.TextAlign=   3
         Sections(0).Cells(2).Style.TextVAlign=   1
         Sections(0).Cells(2).Style.TextWrap=   -1  'True
         Sections(0).Cells(2).Style.ForeColor=   0
         Sections(0).Cells(2).Style.BackColor=   16777215
         Sections(0).Cells(2).Style.NoFill=   -1  'True
         Sections(0).Cells(2).Style.BackPicFile=   ""
         Sections(0).Cells(2).Style.ForePicFile=   ""
         Sections(0).Cells(2).Style.BackPicVertPlacement=   0
         Sections(0).Cells(2).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(2).Style.ForePicPlacement=   0
         Sections(0).Cells(2).Style.ForePicDrawMode=   0
         Sections(0).Cells(2).Style.MarginLeft=   6
         Sections(0).Cells(2).Style.MarginTop=   1
         Sections(0).Cells(2).Style.MarginRight=   6
         Sections(0).Cells(2).Style.MarginBottom=   1
         Sections(0).Cells(2).Style.HasBorders=   -1  'True
         Sections(0).Cells(2).Style.BorderHT=   ""
         Sections(0).Cells(2).Style.BorderHI=   ""
         Sections(0).Cells(2).Style.BorderHB=   ""
         Sections(0).Cells(2).Style.BorderVL=   ""
         Sections(0).Cells(2).Style.BorderVI=   ""
         Sections(0).Cells(2).Style.BorderVR=   ""
         Sections(0).Cells(2).Style.NoClipping=   0   'False
         Sections(0).Cells(2).Style.RTF=   0   'False
         Sections(0).Cells(2).Style.fprops=   131072
         Sections(0).Cells(3).Name=   "CELL_26"
         Sections(0).Cells(3).Exp=   """ """
         Sections(0).Cells(3).NewLine=   -1  'True
         Sections(0).Cells(3).PrivateStyle=   -1  'True
         Sections(0).Cells(3).Style.Name=   "<private>"
         Sections(0).Cells(3).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(3).Style.Font_Name=   "Courier"
         Sections(0).Cells(3).Style.Font_Size=   9.75
         Sections(0).Cells(3).Style.Font_Bold=   -1  'True
         Sections(0).Cells(3).Style.Font_Italic=   0   'False
         Sections(0).Cells(3).Style.Font_Underline=   0   'False
         Sections(0).Cells(3).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(3).Style.Font_Charset=   0
         Sections(0).Cells(3).Style.TextAlign=   1
         Sections(0).Cells(3).Style.TextVAlign=   1
         Sections(0).Cells(3).Style.TextWrap=   -1  'True
         Sections(0).Cells(3).Style.ForeColor=   0
         Sections(0).Cells(3).Style.BackColor=   16777215
         Sections(0).Cells(3).Style.NoFill=   -1  'True
         Sections(0).Cells(3).Style.BackPicFile=   ""
         Sections(0).Cells(3).Style.ForePicFile=   ""
         Sections(0).Cells(3).Style.BackPicVertPlacement=   0
         Sections(0).Cells(3).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(3).Style.ForePicPlacement=   0
         Sections(0).Cells(3).Style.ForePicDrawMode=   0
         Sections(0).Cells(3).Style.MarginLeft=   6
         Sections(0).Cells(3).Style.MarginTop=   1
         Sections(0).Cells(3).Style.MarginRight=   6
         Sections(0).Cells(3).Style.MarginBottom=   1
         Sections(0).Cells(3).Style.HasBorders=   -1  'True
         Sections(0).Cells(3).Style.BorderHT=   ""
         Sections(0).Cells(3).Style.BorderHI=   ""
         Sections(0).Cells(3).Style.BorderHB=   ""
         Sections(0).Cells(3).Style.BorderVL=   ""
         Sections(0).Cells(3).Style.BorderVI=   ""
         Sections(0).Cells(3).Style.BorderVR=   ""
         Sections(0).Cells(3).Style.NoClipping=   0   'False
         Sections(0).Cells(3).Style.RTF=   0   'False
         Sections(0).Cells(3).Style.fprops=   68419585
         Sections(0).Cells(4).Name=   "CELL_3"
         Sections(0).Cells(4).Exp=   """ """
         Sections(0).Cells(4).NewLine=   -1  'True
         Sections(0).Cells(4).Width=   30
         Sections(0).Cells(4).PrivateStyle=   -1  'True
         Sections(0).Cells(4).Style.Name=   "<private>"
         Sections(0).Cells(4).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(4).Style.Font_Name=   "Courier"
         Sections(0).Cells(4).Style.Font_Size=   9.75
         Sections(0).Cells(4).Style.Font_Bold=   0   'False
         Sections(0).Cells(4).Style.Font_Italic=   0   'False
         Sections(0).Cells(4).Style.Font_Underline=   0   'False
         Sections(0).Cells(4).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(4).Style.Font_Charset=   0
         Sections(0).Cells(4).Style.TextAlign=   3
         Sections(0).Cells(4).Style.TextVAlign=   1
         Sections(0).Cells(4).Style.TextWrap=   -1  'True
         Sections(0).Cells(4).Style.ForeColor=   0
         Sections(0).Cells(4).Style.BackColor=   16777215
         Sections(0).Cells(4).Style.NoFill=   -1  'True
         Sections(0).Cells(4).Style.BackPicFile=   ""
         Sections(0).Cells(4).Style.ForePicFile=   ""
         Sections(0).Cells(4).Style.BackPicVertPlacement=   0
         Sections(0).Cells(4).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(4).Style.ForePicPlacement=   0
         Sections(0).Cells(4).Style.ForePicDrawMode=   0
         Sections(0).Cells(4).Style.MarginLeft=   6
         Sections(0).Cells(4).Style.MarginTop=   1
         Sections(0).Cells(4).Style.MarginRight=   6
         Sections(0).Cells(4).Style.MarginBottom=   1
         Sections(0).Cells(4).Style.HasBorders=   -1  'True
         Sections(0).Cells(4).Style.BorderHT=   ""
         Sections(0).Cells(4).Style.BorderHI=   ""
         Sections(0).Cells(4).Style.BorderHB=   ""
         Sections(0).Cells(4).Style.BorderVL=   ""
         Sections(0).Cells(4).Style.BorderVI=   ""
         Sections(0).Cells(4).Style.BorderVR=   ""
         Sections(0).Cells(4).Style.NoClipping=   0   'False
         Sections(0).Cells(4).Style.RTF=   0   'False
         Sections(0).Cells(4).Style.fprops=   22282240
         Sections(0).Cells(5).Name=   "CELL_27"
         Sections(0).Cells(5).Exp=   """Kuitansi"""
         Sections(0).Cells(5).PrivateStyle=   -1  'True
         Sections(0).Cells(5).Style.Name=   "<private>"
         Sections(0).Cells(5).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(5).Style.Font_Name=   "Courier"
         Sections(0).Cells(5).Style.Font_Size=   9.75
         Sections(0).Cells(5).Style.Font_Bold=   -1  'True
         Sections(0).Cells(5).Style.Font_Italic=   0   'False
         Sections(0).Cells(5).Style.Font_Underline=   0   'False
         Sections(0).Cells(5).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(5).Style.Font_Charset=   0
         Sections(0).Cells(5).Style.TextAlign=   1
         Sections(0).Cells(5).Style.TextVAlign=   1
         Sections(0).Cells(5).Style.TextWrap=   -1  'True
         Sections(0).Cells(5).Style.ForeColor=   0
         Sections(0).Cells(5).Style.BackColor=   16777215
         Sections(0).Cells(5).Style.NoFill=   -1  'True
         Sections(0).Cells(5).Style.BackPicFile=   ""
         Sections(0).Cells(5).Style.ForePicFile=   ""
         Sections(0).Cells(5).Style.BackPicVertPlacement=   0
         Sections(0).Cells(5).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(5).Style.ForePicPlacement=   0
         Sections(0).Cells(5).Style.ForePicDrawMode=   0
         Sections(0).Cells(5).Style.MarginLeft=   6
         Sections(0).Cells(5).Style.MarginTop=   1
         Sections(0).Cells(5).Style.MarginRight=   6
         Sections(0).Cells(5).Style.MarginBottom=   1
         Sections(0).Cells(5).Style.HasBorders=   -1  'True
         Sections(0).Cells(5).Style.BorderHT=   ""
         Sections(0).Cells(5).Style.BorderHI=   ""
         Sections(0).Cells(5).Style.BorderHB=   "None"
         Sections(0).Cells(5).Style.BorderVL=   ""
         Sections(0).Cells(5).Style.BorderVI=   ""
         Sections(0).Cells(5).Style.BorderVR=   ""
         Sections(0).Cells(5).Style.NoClipping=   0   'False
         Sections(0).Cells(5).Style.RTF=   0   'False
         Sections(0).Cells(5).Style.fprops=   84017153
         Sections(0).Cells(6).Name=   "CELL_12"
         Sections(0).Cells(6).Exp=   """Tgl : "" & dTgl"
         Sections(0).Cells(6).Width=   30
         Sections(0).Cells(6).PrivateStyle=   -1  'True
         Sections(0).Cells(6).Style.Name=   "<private>"
         Sections(0).Cells(6).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(6).Style.Font_Name=   "Courier"
         Sections(0).Cells(6).Style.Font_Size=   9.75
         Sections(0).Cells(6).Style.Font_Bold=   0   'False
         Sections(0).Cells(6).Style.Font_Italic=   0   'False
         Sections(0).Cells(6).Style.Font_Underline=   0   'False
         Sections(0).Cells(6).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(6).Style.Font_Charset=   0
         Sections(0).Cells(6).Style.TextAlign=   2
         Sections(0).Cells(6).Style.TextVAlign=   1
         Sections(0).Cells(6).Style.TextWrap=   -1  'True
         Sections(0).Cells(6).Style.ForeColor=   0
         Sections(0).Cells(6).Style.BackColor=   16777215
         Sections(0).Cells(6).Style.NoFill=   -1  'True
         Sections(0).Cells(6).Style.BackPicFile=   ""
         Sections(0).Cells(6).Style.ForePicFile=   ""
         Sections(0).Cells(6).Style.BackPicVertPlacement=   0
         Sections(0).Cells(6).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(6).Style.ForePicPlacement=   0
         Sections(0).Cells(6).Style.ForePicDrawMode=   0
         Sections(0).Cells(6).Style.MarginLeft=   6
         Sections(0).Cells(6).Style.MarginTop=   1
         Sections(0).Cells(6).Style.MarginRight=   6
         Sections(0).Cells(6).Style.MarginBottom=   1
         Sections(0).Cells(6).Style.HasBorders=   -1  'True
         Sections(0).Cells(6).Style.BorderHT=   ""
         Sections(0).Cells(6).Style.BorderHI=   ""
         Sections(0).Cells(6).Style.BorderHB=   ""
         Sections(0).Cells(6).Style.BorderVL=   ""
         Sections(0).Cells(6).Style.BorderVI=   ""
         Sections(0).Cells(6).Style.BorderVR=   ""
         Sections(0).Cells(6).Style.NoClipping=   0   'False
         Sections(0).Cells(6).Style.RTF=   0   'False
         Sections(0).Cells(6).Style.fprops=   18087937
         Sections(0).Cells(7).Name=   "CELL_13"
         Sections(0).Cells(7).NewLine=   -1  'True
         Sections(0).Cells(7).Width=   30
         Sections(0).Cells(7).PrivateStyle=   -1  'True
         Sections(0).Cells(7).Style.Name=   "<private>"
         Sections(0).Cells(7).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(7).Style.Font_Name=   "Courier"
         Sections(0).Cells(7).Style.Font_Size=   9.75
         Sections(0).Cells(7).Style.Font_Bold=   0   'False
         Sections(0).Cells(7).Style.Font_Italic=   0   'False
         Sections(0).Cells(7).Style.Font_Underline=   0   'False
         Sections(0).Cells(7).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(7).Style.Font_Charset=   0
         Sections(0).Cells(7).Style.TextAlign=   3
         Sections(0).Cells(7).Style.TextVAlign=   1
         Sections(0).Cells(7).Style.TextWrap=   -1  'True
         Sections(0).Cells(7).Style.ForeColor=   0
         Sections(0).Cells(7).Style.BackColor=   16777215
         Sections(0).Cells(7).Style.NoFill=   -1  'True
         Sections(0).Cells(7).Style.BackPicFile=   ""
         Sections(0).Cells(7).Style.ForePicFile=   ""
         Sections(0).Cells(7).Style.BackPicVertPlacement=   0
         Sections(0).Cells(7).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(7).Style.ForePicPlacement=   0
         Sections(0).Cells(7).Style.ForePicDrawMode=   0
         Sections(0).Cells(7).Style.MarginLeft=   6
         Sections(0).Cells(7).Style.MarginTop=   1
         Sections(0).Cells(7).Style.MarginRight=   6
         Sections(0).Cells(7).Style.MarginBottom=   1
         Sections(0).Cells(7).Style.HasBorders=   -1  'True
         Sections(0).Cells(7).Style.BorderHT=   ""
         Sections(0).Cells(7).Style.BorderHI=   ""
         Sections(0).Cells(7).Style.BorderHB=   ""
         Sections(0).Cells(7).Style.BorderVL=   ""
         Sections(0).Cells(7).Style.BorderVI=   ""
         Sections(0).Cells(7).Style.BorderVR=   ""
         Sections(0).Cells(7).Style.NoClipping=   0   'False
         Sections(0).Cells(7).Style.RTF=   0   'False
         Sections(0).Cells(7).Style.fprops=   22413312
         Sections(0).Cells(8).Name=   "CELL_14"
         Sections(0).Cells(8).Exp=   """No. "" & cSE"
         Sections(0).Cells(8).PrivateStyle=   -1  'True
         Sections(0).Cells(8).Style.Name=   "<private>"
         Sections(0).Cells(8).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(8).Style.Font_Name=   "Courier"
         Sections(0).Cells(8).Style.Font_Size=   9.75
         Sections(0).Cells(8).Style.Font_Bold=   0   'False
         Sections(0).Cells(8).Style.Font_Italic=   0   'False
         Sections(0).Cells(8).Style.Font_Underline=   0   'False
         Sections(0).Cells(8).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(8).Style.Font_Charset=   0
         Sections(0).Cells(8).Style.TextAlign=   1
         Sections(0).Cells(8).Style.TextVAlign=   1
         Sections(0).Cells(8).Style.TextWrap=   -1  'True
         Sections(0).Cells(8).Style.ForeColor=   0
         Sections(0).Cells(8).Style.BackColor=   16777215
         Sections(0).Cells(8).Style.NoFill=   -1  'True
         Sections(0).Cells(8).Style.BackPicFile=   ""
         Sections(0).Cells(8).Style.ForePicFile=   ""
         Sections(0).Cells(8).Style.BackPicVertPlacement=   0
         Sections(0).Cells(8).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(8).Style.ForePicPlacement=   0
         Sections(0).Cells(8).Style.ForePicDrawMode=   0
         Sections(0).Cells(8).Style.MarginLeft=   6
         Sections(0).Cells(8).Style.MarginTop=   1
         Sections(0).Cells(8).Style.MarginRight=   6
         Sections(0).Cells(8).Style.MarginBottom=   1
         Sections(0).Cells(8).Style.HasBorders=   -1  'True
         Sections(0).Cells(8).Style.BorderHT=   ""
         Sections(0).Cells(8).Style.BorderHI=   ""
         Sections(0).Cells(8).Style.BorderHB=   ""
         Sections(0).Cells(8).Style.BorderVL=   ""
         Sections(0).Cells(8).Style.BorderVI=   ""
         Sections(0).Cells(8).Style.BorderVR=   ""
         Sections(0).Cells(8).Style.NoClipping=   0   'False
         Sections(0).Cells(8).Style.RTF=   0   'False
         Sections(0).Cells(8).Style.fprops=   16908289
         Sections(0).Cells(9).Name=   "CELL_15"
         Sections(0).Cells(9).Exp=   """Page "" & PageNo()"
         Sections(0).Cells(9).Width=   30
         Sections(0).Cells(9).PrivateStyle=   -1  'True
         Sections(0).Cells(9).Style.Name=   "<private>"
         Sections(0).Cells(9).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(9).Style.Font_Name=   "Courier"
         Sections(0).Cells(9).Style.Font_Size=   9.75
         Sections(0).Cells(9).Style.Font_Bold=   0   'False
         Sections(0).Cells(9).Style.Font_Italic=   0   'False
         Sections(0).Cells(9).Style.Font_Underline=   0   'False
         Sections(0).Cells(9).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(9).Style.Font_Charset=   0
         Sections(0).Cells(9).Style.TextAlign=   2
         Sections(0).Cells(9).Style.TextVAlign=   1
         Sections(0).Cells(9).Style.TextWrap=   -1  'True
         Sections(0).Cells(9).Style.ForeColor=   0
         Sections(0).Cells(9).Style.BackColor=   16777215
         Sections(0).Cells(9).Style.NoFill=   -1  'True
         Sections(0).Cells(9).Style.BackPicFile=   ""
         Sections(0).Cells(9).Style.ForePicFile=   ""
         Sections(0).Cells(9).Style.BackPicVertPlacement=   0
         Sections(0).Cells(9).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(9).Style.ForePicPlacement=   0
         Sections(0).Cells(9).Style.ForePicDrawMode=   0
         Sections(0).Cells(9).Style.MarginLeft=   6
         Sections(0).Cells(9).Style.MarginTop=   1
         Sections(0).Cells(9).Style.MarginRight=   6
         Sections(0).Cells(9).Style.MarginBottom=   1
         Sections(0).Cells(9).Style.HasBorders=   -1  'True
         Sections(0).Cells(9).Style.BorderHT=   ""
         Sections(0).Cells(9).Style.BorderHI=   ""
         Sections(0).Cells(9).Style.BorderHB=   ""
         Sections(0).Cells(9).Style.BorderVL=   ""
         Sections(0).Cells(9).Style.BorderVI=   ""
         Sections(0).Cells(9).Style.BorderVR=   ""
         Sections(0).Cells(9).Style.NoClipping=   0   'False
         Sections(0).Cells(9).Style.RTF=   0   'False
         Sections(0).Cells(9).Style.fprops=   17956865
         Sections(0).Cells(10).Name=   "CELL_4"
         Sections(0).Cells(10).Exp=   """Cust ID : "" & cKodeAnggota"
         Sections(0).Cells(10).NewLine=   -1  'True
         Sections(0).Cells(10).Height=   6
         Sections(0).Cells(10).AutoHeight=   0   'False
         Sections(0).Cells(10).PrivateStyle=   -1  'True
         Sections(0).Cells(10).Style.Name=   "<private>"
         Sections(0).Cells(10).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(10).Style.Font_Name=   "Courier"
         Sections(0).Cells(10).Style.Font_Size=   9.75
         Sections(0).Cells(10).Style.Font_Bold=   0   'False
         Sections(0).Cells(10).Style.Font_Italic=   0   'False
         Sections(0).Cells(10).Style.Font_Underline=   0   'False
         Sections(0).Cells(10).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(10).Style.Font_Charset=   0
         Sections(0).Cells(10).Style.TextAlign=   3
         Sections(0).Cells(10).Style.TextVAlign=   1
         Sections(0).Cells(10).Style.TextWrap=   -1  'True
         Sections(0).Cells(10).Style.ForeColor=   0
         Sections(0).Cells(10).Style.BackColor=   16777215
         Sections(0).Cells(10).Style.NoFill=   -1  'True
         Sections(0).Cells(10).Style.BackPicFile=   ""
         Sections(0).Cells(10).Style.ForePicFile=   ""
         Sections(0).Cells(10).Style.BackPicVertPlacement=   0
         Sections(0).Cells(10).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(10).Style.ForePicPlacement=   0
         Sections(0).Cells(10).Style.ForePicDrawMode=   0
         Sections(0).Cells(10).Style.MarginLeft=   6
         Sections(0).Cells(10).Style.MarginTop=   1
         Sections(0).Cells(10).Style.MarginRight=   6
         Sections(0).Cells(10).Style.MarginBottom=   1
         Sections(0).Cells(10).Style.HasBorders=   -1  'True
         Sections(0).Cells(10).Style.BorderHT=   ""
         Sections(0).Cells(10).Style.BorderHI=   ""
         Sections(0).Cells(10).Style.BorderHB=   ""
         Sections(0).Cells(10).Style.BorderVL=   ""
         Sections(0).Cells(10).Style.BorderVI=   ""
         Sections(0).Cells(10).Style.BorderVR=   ""
         Sections(0).Cells(10).Style.NoClipping=   0   'False
         Sections(0).Cells(10).Style.RTF=   0   'False
         Sections(0).Cells(10).Style.fprops=   18087936
         Sections(0).Cells(11).Name=   "CELL_17"
         Sections(0).Cells(11).Exp=   """Cust Name : "" & cNama"
         Sections(0).Cells(11).NewLine=   -1  'True
         Sections(0).Cells(11).PrivateStyle=   -1  'True
         Sections(0).Cells(11).Style.Name=   "<private>"
         Sections(0).Cells(11).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(11).Style.Font_Name=   "Courier"
         Sections(0).Cells(11).Style.Font_Size=   9.75
         Sections(0).Cells(11).Style.Font_Bold=   0   'False
         Sections(0).Cells(11).Style.Font_Italic=   0   'False
         Sections(0).Cells(11).Style.Font_Underline=   0   'False
         Sections(0).Cells(11).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(11).Style.Font_Charset=   0
         Sections(0).Cells(11).Style.TextAlign=   3
         Sections(0).Cells(11).Style.TextVAlign=   1
         Sections(0).Cells(11).Style.TextWrap=   -1  'True
         Sections(0).Cells(11).Style.ForeColor=   0
         Sections(0).Cells(11).Style.BackColor=   16777215
         Sections(0).Cells(11).Style.NoFill=   -1  'True
         Sections(0).Cells(11).Style.BackPicFile=   ""
         Sections(0).Cells(11).Style.ForePicFile=   ""
         Sections(0).Cells(11).Style.BackPicVertPlacement=   0
         Sections(0).Cells(11).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(11).Style.ForePicPlacement=   0
         Sections(0).Cells(11).Style.ForePicDrawMode=   0
         Sections(0).Cells(11).Style.MarginLeft=   6
         Sections(0).Cells(11).Style.MarginTop=   1
         Sections(0).Cells(11).Style.MarginRight=   6
         Sections(0).Cells(11).Style.MarginBottom=   1
         Sections(0).Cells(11).Style.HasBorders=   -1  'True
         Sections(0).Cells(11).Style.BorderHT=   ""
         Sections(0).Cells(11).Style.BorderHI=   ""
         Sections(0).Cells(11).Style.BorderHB=   ""
         Sections(0).Cells(11).Style.BorderVL=   ""
         Sections(0).Cells(11).Style.BorderVI=   ""
         Sections(0).Cells(11).Style.BorderVR=   ""
         Sections(0).Cells(11).Style.NoClipping=   0   'False
         Sections(0).Cells(11).Style.RTF=   0   'False
         Sections(0).Cells(11).Style.fprops=   18087936
         Sections(0).Cells(12).Name=   "CELL_20"
         Sections(0).Cells(12).Exp=   """Print By : ""& cUserName"
         Sections(0).Cells(12).PrivateStyle=   -1  'True
         Sections(0).Cells(12).Style.Name=   "<private>"
         Sections(0).Cells(12).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(12).Style.Font_Name=   "Courier"
         Sections(0).Cells(12).Style.Font_Size=   9.75
         Sections(0).Cells(12).Style.Font_Bold=   0   'False
         Sections(0).Cells(12).Style.Font_Italic=   0   'False
         Sections(0).Cells(12).Style.Font_Underline=   0   'False
         Sections(0).Cells(12).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(12).Style.Font_Charset=   0
         Sections(0).Cells(12).Style.TextAlign=   2
         Sections(0).Cells(12).Style.TextVAlign=   1
         Sections(0).Cells(12).Style.TextWrap=   -1  'True
         Sections(0).Cells(12).Style.ForeColor=   0
         Sections(0).Cells(12).Style.BackColor=   16777215
         Sections(0).Cells(12).Style.NoFill=   -1  'True
         Sections(0).Cells(12).Style.BackPicFile=   ""
         Sections(0).Cells(12).Style.ForePicFile=   ""
         Sections(0).Cells(12).Style.BackPicVertPlacement=   0
         Sections(0).Cells(12).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(12).Style.ForePicPlacement=   0
         Sections(0).Cells(12).Style.ForePicDrawMode=   0
         Sections(0).Cells(12).Style.MarginLeft=   6
         Sections(0).Cells(12).Style.MarginTop=   1
         Sections(0).Cells(12).Style.MarginRight=   6
         Sections(0).Cells(12).Style.MarginBottom=   1
         Sections(0).Cells(12).Style.HasBorders=   -1  'True
         Sections(0).Cells(12).Style.BorderHT=   ""
         Sections(0).Cells(12).Style.BorderHI=   ""
         Sections(0).Cells(12).Style.BorderHB=   ""
         Sections(0).Cells(12).Style.BorderVL=   ""
         Sections(0).Cells(12).Style.BorderVI=   ""
         Sections(0).Cells(12).Style.BorderVR=   ""
         Sections(0).Cells(12).Style.NoClipping=   0   'False
         Sections(0).Cells(12).Style.RTF=   0   'False
         Sections(0).Cells(12).Style.fprops=   16777217
         Sections(0).Cells(13).Name=   "CELL_16"
         Sections(0).Cells(13).Exp=   "cKota"
         Sections(0).Cells(13).NewLine=   -1  'True
         Sections(0).Cells(13).PrivateStyle=   -1  'True
         Sections(0).Cells(13).Style.Name=   "<private>"
         Sections(0).Cells(13).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(13).Style.Font_Name=   "Courier"
         Sections(0).Cells(13).Style.Font_Size=   9.75
         Sections(0).Cells(13).Style.Font_Bold=   0   'False
         Sections(0).Cells(13).Style.Font_Italic=   0   'False
         Sections(0).Cells(13).Style.Font_Underline=   0   'False
         Sections(0).Cells(13).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(13).Style.Font_Charset=   0
         Sections(0).Cells(13).Style.TextAlign=   3
         Sections(0).Cells(13).Style.TextVAlign=   1
         Sections(0).Cells(13).Style.TextWrap=   -1  'True
         Sections(0).Cells(13).Style.ForeColor=   0
         Sections(0).Cells(13).Style.BackColor=   16777215
         Sections(0).Cells(13).Style.NoFill=   -1  'True
         Sections(0).Cells(13).Style.BackPicFile=   ""
         Sections(0).Cells(13).Style.ForePicFile=   ""
         Sections(0).Cells(13).Style.BackPicVertPlacement=   0
         Sections(0).Cells(13).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(13).Style.ForePicPlacement=   0
         Sections(0).Cells(13).Style.ForePicDrawMode=   0
         Sections(0).Cells(13).Style.MarginLeft=   6
         Sections(0).Cells(13).Style.MarginTop=   1
         Sections(0).Cells(13).Style.MarginRight=   6
         Sections(0).Cells(13).Style.MarginBottom=   1
         Sections(0).Cells(13).Style.HasBorders=   -1  'True
         Sections(0).Cells(13).Style.BorderHT=   ""
         Sections(0).Cells(13).Style.BorderHI=   ""
         Sections(0).Cells(13).Style.BorderHB=   ""
         Sections(0).Cells(13).Style.BorderVL=   ""
         Sections(0).Cells(13).Style.BorderVI=   ""
         Sections(0).Cells(13).Style.BorderVR=   ""
         Sections(0).Cells(13).Style.NoClipping=   0   'False
         Sections(0).Cells(13).Style.RTF=   0   'False
         Sections(0).Cells(13).Style.fprops=   18087936
         Sections(0).Cells(14).Name=   "CELL_18"
         Sections(0).Cells(14).PrivateStyle=   -1  'True
         Sections(0).Cells(14).Style.Name=   "<private>"
         Sections(0).Cells(14).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(14).Style.Font_Name=   "Courier"
         Sections(0).Cells(14).Style.Font_Size=   9.75
         Sections(0).Cells(14).Style.Font_Bold=   0   'False
         Sections(0).Cells(14).Style.Font_Italic=   0   'False
         Sections(0).Cells(14).Style.Font_Underline=   0   'False
         Sections(0).Cells(14).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(14).Style.Font_Charset=   0
         Sections(0).Cells(14).Style.TextAlign=   2
         Sections(0).Cells(14).Style.TextVAlign=   1
         Sections(0).Cells(14).Style.TextWrap=   -1  'True
         Sections(0).Cells(14).Style.ForeColor=   0
         Sections(0).Cells(14).Style.BackColor=   16777215
         Sections(0).Cells(14).Style.NoFill=   -1  'True
         Sections(0).Cells(14).Style.BackPicFile=   ""
         Sections(0).Cells(14).Style.ForePicFile=   ""
         Sections(0).Cells(14).Style.BackPicVertPlacement=   0
         Sections(0).Cells(14).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(14).Style.ForePicPlacement=   0
         Sections(0).Cells(14).Style.ForePicDrawMode=   0
         Sections(0).Cells(14).Style.MarginLeft=   6
         Sections(0).Cells(14).Style.MarginTop=   1
         Sections(0).Cells(14).Style.MarginRight=   6
         Sections(0).Cells(14).Style.MarginBottom=   1
         Sections(0).Cells(14).Style.HasBorders=   -1  'True
         Sections(0).Cells(14).Style.BorderHT=   ""
         Sections(0).Cells(14).Style.BorderHI=   ""
         Sections(0).Cells(14).Style.BorderHB=   ""
         Sections(0).Cells(14).Style.BorderVL=   ""
         Sections(0).Cells(14).Style.BorderVI=   ""
         Sections(0).Cells(14).Style.BorderVR=   ""
         Sections(0).Cells(14).Style.NoClipping=   0   'False
         Sections(0).Cells(14).Style.RTF=   0   'False
         Sections(0).Cells(14).Style.fprops=   17825793
         Sections(0).Cells(15).Name=   "CELL_19"
         Sections(0).Cells(15).Exp=   "Now"
         Sections(0).Cells(15).PrivateStyle=   -1  'True
         Sections(0).Cells(15).Format=   "dd-MM-yyyy HH:MM:SS"
         Sections(0).Cells(15).Style.Name=   "<private>"
         Sections(0).Cells(15).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(15).Style.Font_Name=   "Courier"
         Sections(0).Cells(15).Style.Font_Size=   9.75
         Sections(0).Cells(15).Style.Font_Bold=   0   'False
         Sections(0).Cells(15).Style.Font_Italic=   0   'False
         Sections(0).Cells(15).Style.Font_Underline=   0   'False
         Sections(0).Cells(15).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(15).Style.Font_Charset=   0
         Sections(0).Cells(15).Style.TextAlign=   2
         Sections(0).Cells(15).Style.TextVAlign=   1
         Sections(0).Cells(15).Style.TextWrap=   -1  'True
         Sections(0).Cells(15).Style.ForeColor=   0
         Sections(0).Cells(15).Style.BackColor=   16777215
         Sections(0).Cells(15).Style.NoFill=   -1  'True
         Sections(0).Cells(15).Style.BackPicFile=   ""
         Sections(0).Cells(15).Style.ForePicFile=   ""
         Sections(0).Cells(15).Style.BackPicVertPlacement=   0
         Sections(0).Cells(15).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(15).Style.ForePicPlacement=   0
         Sections(0).Cells(15).Style.ForePicDrawMode=   0
         Sections(0).Cells(15).Style.MarginLeft=   6
         Sections(0).Cells(15).Style.MarginTop=   1
         Sections(0).Cells(15).Style.MarginRight=   6
         Sections(0).Cells(15).Style.MarginBottom=   1
         Sections(0).Cells(15).Style.HasBorders=   -1  'True
         Sections(0).Cells(15).Style.BorderHT=   ""
         Sections(0).Cells(15).Style.BorderHI=   ""
         Sections(0).Cells(15).Style.BorderHB=   ""
         Sections(0).Cells(15).Style.BorderVL=   ""
         Sections(0).Cells(15).Style.BorderVI=   ""
         Sections(0).Cells(15).Style.BorderVR=   ""
         Sections(0).Cells(15).Style.NoClipping=   0   'False
         Sections(0).Cells(15).Style.RTF=   0   'False
         Sections(0).Cells(15).Style.fprops=   16777217
         Sections(1).Name=   "DetailHeader"
         Sections(1).Type=   3
         Sections(1).StyleExp=   "'Tdb_Header'"
         Sections(1).Tabulator=   "Detail"
         Sections(1).Cells.Count=   5
         Sections(1).Cells(0).Name=   "CELL_0"
         Sections(1).Cells(0).Exp=   """No."""
         Sections(1).Cells(0).PrivateStyle=   -1  'True
         Sections(1).Cells(0).Style.Name=   "<private>"
         Sections(1).Cells(0).Style.ParentName=   "Tdb_Header"
         Sections(1).Cells(0).Style.Font_Name=   "Courier"
         Sections(1).Cells(0).Style.Font_Size=   9.75
         Sections(1).Cells(0).Style.Font_Bold=   -1  'True
         Sections(1).Cells(0).Style.Font_Italic=   0   'False
         Sections(1).Cells(0).Style.Font_Underline=   0   'False
         Sections(1).Cells(0).Style.Font_Strikeout=   0   'False
         Sections(1).Cells(0).Style.Font_Charset=   0
         Sections(1).Cells(0).Style.TextAlign=   1
         Sections(1).Cells(0).Style.TextVAlign=   1
         Sections(1).Cells(0).Style.TextWrap=   -1  'True
         Sections(1).Cells(0).Style.ForeColor=   0
         Sections(1).Cells(0).Style.BackColor=   16777215
         Sections(1).Cells(0).Style.NoFill=   -1  'True
         Sections(1).Cells(0).Style.BackPicFile=   ""
         Sections(1).Cells(0).Style.ForePicFile=   ""
         Sections(1).Cells(0).Style.BackPicVertPlacement=   0
         Sections(1).Cells(0).Style.BackPicHorzPlacement=   0
         Sections(1).Cells(0).Style.ForePicPlacement=   0
         Sections(1).Cells(0).Style.ForePicDrawMode=   0
         Sections(1).Cells(0).Style.MarginLeft=   6
         Sections(1).Cells(0).Style.MarginTop=   1
         Sections(1).Cells(0).Style.MarginRight=   6
         Sections(1).Cells(0).Style.MarginBottom=   1
         Sections(1).Cells(0).Style.HasBorders=   -1  'True
         Sections(1).Cells(0).Style.BorderHT=   "Single"
         Sections(1).Cells(0).Style.BorderHI=   "Single"
         Sections(1).Cells(0).Style.BorderHB=   "Single"
         Sections(1).Cells(0).Style.BorderVL=   "Single"
         Sections(1).Cells(0).Style.BorderVI=   "Single"
         Sections(1).Cells(0).Style.BorderVR=   "Single"
         Sections(1).Cells(0).Style.NoClipping=   0   'False
         Sections(1).Cells(0).Style.RTF=   0   'False
         Sections(1).Cells(0).Style.fprops=   1835009
         Sections(1).Cells(1).Name=   "CELL_2"
         Sections(1).Cells(1).Exp=   """No Invoice Penjualan"""
         Sections(1).Cells(1).PrivateStyle=   -1  'True
         Sections(1).Cells(1).Style.Name=   "<private>"
         Sections(1).Cells(1).Style.ParentName=   "Tdb_Header"
         Sections(1).Cells(1).Style.Font_Name=   "Courier"
         Sections(1).Cells(1).Style.Font_Size=   9.75
         Sections(1).Cells(1).Style.Font_Bold=   -1  'True
         Sections(1).Cells(1).Style.Font_Italic=   0   'False
         Sections(1).Cells(1).Style.Font_Underline=   0   'False
         Sections(1).Cells(1).Style.Font_Strikeout=   0   'False
         Sections(1).Cells(1).Style.Font_Charset=   0
         Sections(1).Cells(1).Style.TextAlign=   1
         Sections(1).Cells(1).Style.TextVAlign=   1
         Sections(1).Cells(1).Style.TextWrap=   -1  'True
         Sections(1).Cells(1).Style.ForeColor=   0
         Sections(1).Cells(1).Style.BackColor=   16777215
         Sections(1).Cells(1).Style.NoFill=   -1  'True
         Sections(1).Cells(1).Style.BackPicFile=   ""
         Sections(1).Cells(1).Style.ForePicFile=   ""
         Sections(1).Cells(1).Style.BackPicVertPlacement=   0
         Sections(1).Cells(1).Style.BackPicHorzPlacement=   0
         Sections(1).Cells(1).Style.ForePicPlacement=   0
         Sections(1).Cells(1).Style.ForePicDrawMode=   0
         Sections(1).Cells(1).Style.MarginLeft=   6
         Sections(1).Cells(1).Style.MarginTop=   1
         Sections(1).Cells(1).Style.MarginRight=   6
         Sections(1).Cells(1).Style.MarginBottom=   1
         Sections(1).Cells(1).Style.HasBorders=   -1  'True
         Sections(1).Cells(1).Style.BorderHT=   "Single"
         Sections(1).Cells(1).Style.BorderHI=   "Single"
         Sections(1).Cells(1).Style.BorderHB=   "Single"
         Sections(1).Cells(1).Style.BorderVL=   "Single"
         Sections(1).Cells(1).Style.BorderVI=   "Single"
         Sections(1).Cells(1).Style.BorderVR=   "Single"
         Sections(1).Cells(1).Style.NoClipping=   0   'False
         Sections(1).Cells(1).Style.RTF=   0   'False
         Sections(1).Cells(1).Style.fprops=   1835009
         Sections(1).Cells(2).Name=   "CELL_4"
         Sections(1).Cells(2).Exp=   """Bon"""
         Sections(1).Cells(2).PrivateStyle=   -1  'True
         Sections(1).Cells(2).Style.Name=   "<private>"
         Sections(1).Cells(2).Style.ParentName=   "Tdb_Header"
         Sections(1).Cells(2).Style.Font_Name=   "Courier"
         Sections(1).Cells(2).Style.Font_Size=   9.75
         Sections(1).Cells(2).Style.Font_Bold=   -1  'True
         Sections(1).Cells(2).Style.Font_Italic=   0   'False
         Sections(1).Cells(2).Style.Font_Underline=   0   'False
         Sections(1).Cells(2).Style.Font_Strikeout=   0   'False
         Sections(1).Cells(2).Style.Font_Charset=   0
         Sections(1).Cells(2).Style.TextAlign=   1
         Sections(1).Cells(2).Style.TextVAlign=   1
         Sections(1).Cells(2).Style.TextWrap=   -1  'True
         Sections(1).Cells(2).Style.ForeColor=   0
         Sections(1).Cells(2).Style.BackColor=   16777215
         Sections(1).Cells(2).Style.NoFill=   -1  'True
         Sections(1).Cells(2).Style.BackPicFile=   ""
         Sections(1).Cells(2).Style.ForePicFile=   ""
         Sections(1).Cells(2).Style.BackPicVertPlacement=   0
         Sections(1).Cells(2).Style.BackPicHorzPlacement=   0
         Sections(1).Cells(2).Style.ForePicPlacement=   0
         Sections(1).Cells(2).Style.ForePicDrawMode=   0
         Sections(1).Cells(2).Style.MarginLeft=   6
         Sections(1).Cells(2).Style.MarginTop=   1
         Sections(1).Cells(2).Style.MarginRight=   6
         Sections(1).Cells(2).Style.MarginBottom=   1
         Sections(1).Cells(2).Style.HasBorders=   -1  'True
         Sections(1).Cells(2).Style.BorderHT=   "Single"
         Sections(1).Cells(2).Style.BorderHI=   "Single"
         Sections(1).Cells(2).Style.BorderHB=   "Single"
         Sections(1).Cells(2).Style.BorderVL=   "Single"
         Sections(1).Cells(2).Style.BorderVI=   "Single"
         Sections(1).Cells(2).Style.BorderVR=   "Single"
         Sections(1).Cells(2).Style.NoClipping=   0   'False
         Sections(1).Cells(2).Style.RTF=   0   'False
         Sections(1).Cells(2).Style.fprops=   1835009
         Sections(1).Cells(3).Name=   "CELL_3"
         Sections(1).Cells(3).Exp=   """Bayar"""
         Sections(1).Cells(3).PrivateStyle=   -1  'True
         Sections(1).Cells(3).Style.Name=   "<private>"
         Sections(1).Cells(3).Style.ParentName=   "Tdb_Header"
         Sections(1).Cells(3).Style.Font_Name=   "Courier"
         Sections(1).Cells(3).Style.Font_Size=   9.75
         Sections(1).Cells(3).Style.Font_Bold=   -1  'True
         Sections(1).Cells(3).Style.Font_Italic=   0   'False
         Sections(1).Cells(3).Style.Font_Underline=   0   'False
         Sections(1).Cells(3).Style.Font_Strikeout=   0   'False
         Sections(1).Cells(3).Style.Font_Charset=   0
         Sections(1).Cells(3).Style.TextAlign=   1
         Sections(1).Cells(3).Style.TextVAlign=   1
         Sections(1).Cells(3).Style.TextWrap=   -1  'True
         Sections(1).Cells(3).Style.ForeColor=   0
         Sections(1).Cells(3).Style.BackColor=   16777215
         Sections(1).Cells(3).Style.NoFill=   -1  'True
         Sections(1).Cells(3).Style.BackPicFile=   ""
         Sections(1).Cells(3).Style.ForePicFile=   ""
         Sections(1).Cells(3).Style.BackPicVertPlacement=   0
         Sections(1).Cells(3).Style.BackPicHorzPlacement=   0
         Sections(1).Cells(3).Style.ForePicPlacement=   0
         Sections(1).Cells(3).Style.ForePicDrawMode=   0
         Sections(1).Cells(3).Style.MarginLeft=   6
         Sections(1).Cells(3).Style.MarginTop=   1
         Sections(1).Cells(3).Style.MarginRight=   6
         Sections(1).Cells(3).Style.MarginBottom=   1
         Sections(1).Cells(3).Style.HasBorders=   -1  'True
         Sections(1).Cells(3).Style.BorderHT=   "Single"
         Sections(1).Cells(3).Style.BorderHI=   "Single"
         Sections(1).Cells(3).Style.BorderHB=   "Single"
         Sections(1).Cells(3).Style.BorderVL=   "Single"
         Sections(1).Cells(3).Style.BorderVI=   "Single"
         Sections(1).Cells(3).Style.BorderVR=   "Single"
         Sections(1).Cells(3).Style.NoClipping=   0   'False
         Sections(1).Cells(3).Style.RTF=   0   'False
         Sections(1).Cells(3).Style.fprops=   1835009
         Sections(1).Cells(4).Name=   "CELL_1"
         Sections(1).Cells(4).Exp=   """Tanggal"""
         Sections(1).Cells(4).PrivateStyle=   -1  'True
         Sections(1).Cells(4).Style.Name=   "<private>"
         Sections(1).Cells(4).Style.ParentName=   "Tdb_Header"
         Sections(1).Cells(4).Style.Font_Name=   "Courier"
         Sections(1).Cells(4).Style.Font_Size=   9.75
         Sections(1).Cells(4).Style.Font_Bold=   -1  'True
         Sections(1).Cells(4).Style.Font_Italic=   0   'False
         Sections(1).Cells(4).Style.Font_Underline=   0   'False
         Sections(1).Cells(4).Style.Font_Strikeout=   0   'False
         Sections(1).Cells(4).Style.Font_Charset=   0
         Sections(1).Cells(4).Style.TextAlign=   1
         Sections(1).Cells(4).Style.TextVAlign=   1
         Sections(1).Cells(4).Style.TextWrap=   -1  'True
         Sections(1).Cells(4).Style.ForeColor=   0
         Sections(1).Cells(4).Style.BackColor=   16777215
         Sections(1).Cells(4).Style.NoFill=   -1  'True
         Sections(1).Cells(4).Style.BackPicFile=   ""
         Sections(1).Cells(4).Style.ForePicFile=   ""
         Sections(1).Cells(4).Style.BackPicVertPlacement=   0
         Sections(1).Cells(4).Style.BackPicHorzPlacement=   0
         Sections(1).Cells(4).Style.ForePicPlacement=   0
         Sections(1).Cells(4).Style.ForePicDrawMode=   0
         Sections(1).Cells(4).Style.MarginLeft=   6
         Sections(1).Cells(4).Style.MarginTop=   1
         Sections(1).Cells(4).Style.MarginRight=   6
         Sections(1).Cells(4).Style.MarginBottom=   1
         Sections(1).Cells(4).Style.HasBorders=   -1  'True
         Sections(1).Cells(4).Style.BorderHT=   "Single"
         Sections(1).Cells(4).Style.BorderHI=   "Single"
         Sections(1).Cells(4).Style.BorderHB=   "Single"
         Sections(1).Cells(4).Style.BorderVL=   "Single"
         Sections(1).Cells(4).Style.BorderVI=   "Single"
         Sections(1).Cells(4).Style.BorderVR=   "Single"
         Sections(1).Cells(4).Style.NoClipping=   0   'False
         Sections(1).Cells(4).Style.RTF=   0   'False
         Sections(1).Cells(4).Style.fprops=   1835009
         Sections(2).Name=   "Detail"
         Sections(2).Type=   4
         Sections(2).StyleExp=   "'Tdb_Body'"
         Sections(2).AutoHeight=   0   'False
         Sections(2).Height=   5
         Sections(2).Cells.Count=   5
         Sections(2).Cells(0).Name=   "CELL_0"
         Sections(2).Cells(0).Exp=   "Nomor"
         Sections(2).Cells(0).Width=   4
         Sections(2).Cells(0).PrivateStyle=   -1  'True
         Sections(2).Cells(0).Style.Name=   "<private>"
         Sections(2).Cells(0).Style.ParentName=   "Tdb_Body"
         Sections(2).Cells(0).Style.Font_Name=   "Courier"
         Sections(2).Cells(0).Style.Font_Size=   9.75
         Sections(2).Cells(0).Style.Font_Bold=   0   'False
         Sections(2).Cells(0).Style.Font_Italic=   0   'False
         Sections(2).Cells(0).Style.Font_Underline=   0   'False
         Sections(2).Cells(0).Style.Font_Strikeout=   0   'False
         Sections(2).Cells(0).Style.Font_Charset=   0
         Sections(2).Cells(0).Style.TextAlign=   0
         Sections(2).Cells(0).Style.TextVAlign=   1
         Sections(2).Cells(0).Style.TextWrap=   0   'False
         Sections(2).Cells(0).Style.ForeColor=   0
         Sections(2).Cells(0).Style.BackColor=   16777215
         Sections(2).Cells(0).Style.NoFill=   -1  'True
         Sections(2).Cells(0).Style.BackPicFile=   ""
         Sections(2).Cells(0).Style.ForePicFile=   ""
         Sections(2).Cells(0).Style.BackPicVertPlacement=   0
         Sections(2).Cells(0).Style.BackPicHorzPlacement=   0
         Sections(2).Cells(0).Style.ForePicPlacement=   0
         Sections(2).Cells(0).Style.ForePicDrawMode=   0
         Sections(2).Cells(0).Style.MarginLeft=   6
         Sections(2).Cells(0).Style.MarginTop=   0
         Sections(2).Cells(0).Style.MarginRight=   6
         Sections(2).Cells(0).Style.MarginBottom=   0
         Sections(2).Cells(0).Style.HasBorders=   -1  'True
         Sections(2).Cells(0).Style.BorderHT=   ""
         Sections(2).Cells(0).Style.BorderHI=   ""
         Sections(2).Cells(0).Style.BorderHB=   ""
         Sections(2).Cells(0).Style.BorderVL=   "Single"
         Sections(2).Cells(0).Style.BorderVI=   "Single"
         Sections(2).Cells(0).Style.BorderVR=   "Single"
         Sections(2).Cells(0).Style.NoClipping=   0   'False
         Sections(2).Cells(0).Style.RTF=   0   'False
         Sections(2).Cells(0).Style.fprops=   1855493
         Sections(2).Cells(1).Name=   "CELL_2"
         Sections(2).Cells(1).Exp=   "NoInvoice"
         Sections(2).Cells(1).PrivateStyle=   -1  'True
         Sections(2).Cells(1).Style.Name=   "<private>"
         Sections(2).Cells(1).Style.ParentName=   "Tdb_Body"
         Sections(2).Cells(1).Style.Font_Name=   "Courier"
         Sections(2).Cells(1).Style.Font_Size=   9.75
         Sections(2).Cells(1).Style.Font_Bold=   0   'False
         Sections(2).Cells(1).Style.Font_Italic=   0   'False
         Sections(2).Cells(1).Style.Font_Underline=   0   'False
         Sections(2).Cells(1).Style.Font_Strikeout=   0   'False
         Sections(2).Cells(1).Style.Font_Charset=   0
         Sections(2).Cells(1).Style.TextAlign=   3
         Sections(2).Cells(1).Style.TextVAlign=   1
         Sections(2).Cells(1).Style.TextWrap=   0   'False
         Sections(2).Cells(1).Style.ForeColor=   0
         Sections(2).Cells(1).Style.BackColor=   16777215
         Sections(2).Cells(1).Style.NoFill=   -1  'True
         Sections(2).Cells(1).Style.BackPicFile=   ""
         Sections(2).Cells(1).Style.ForePicFile=   ""
         Sections(2).Cells(1).Style.BackPicVertPlacement=   0
         Sections(2).Cells(1).Style.BackPicHorzPlacement=   0
         Sections(2).Cells(1).Style.ForePicPlacement=   0
         Sections(2).Cells(1).Style.ForePicDrawMode=   0
         Sections(2).Cells(1).Style.MarginLeft=   6
         Sections(2).Cells(1).Style.MarginTop=   0
         Sections(2).Cells(1).Style.MarginRight=   6
         Sections(2).Cells(1).Style.MarginBottom=   0
         Sections(2).Cells(1).Style.HasBorders=   -1  'True
         Sections(2).Cells(1).Style.BorderHT=   ""
         Sections(2).Cells(1).Style.BorderHI=   ""
         Sections(2).Cells(1).Style.BorderHB=   ""
         Sections(2).Cells(1).Style.BorderVL=   "Single"
         Sections(2).Cells(1).Style.BorderVI=   "Single"
         Sections(2).Cells(1).Style.BorderVR=   "Single"
         Sections(2).Cells(1).Style.NoClipping=   0   'False
         Sections(2).Cells(1).Style.RTF=   0   'False
         Sections(2).Cells(1).Style.fprops=   1835012
         Sections(2).Cells(2).Name=   "CELL_3"
         Sections(2).Cells(2).Exp=   "Total"
         Sections(2).Cells(2).Width=   13
         Sections(2).Cells(2).PrivateStyle=   -1  'True
         Sections(2).Cells(2).Format=   "###,###,###,###,###,##0.00"
         Sections(2).Cells(2).Style.Name=   "<private>"
         Sections(2).Cells(2).Style.ParentName=   "Tdb_Body"
         Sections(2).Cells(2).Style.Font_Name=   "Courier"
         Sections(2).Cells(2).Style.Font_Size=   9.75
         Sections(2).Cells(2).Style.Font_Bold=   0   'False
         Sections(2).Cells(2).Style.Font_Italic=   0   'False
         Sections(2).Cells(2).Style.Font_Underline=   0   'False
         Sections(2).Cells(2).Style.Font_Strikeout=   0   'False
         Sections(2).Cells(2).Style.Font_Charset=   0
         Sections(2).Cells(2).Style.TextAlign=   2
         Sections(2).Cells(2).Style.TextVAlign=   1
         Sections(2).Cells(2).Style.TextWrap=   0   'False
         Sections(2).Cells(2).Style.ForeColor=   0
         Sections(2).Cells(2).Style.BackColor=   16777215
         Sections(2).Cells(2).Style.NoFill=   -1  'True
         Sections(2).Cells(2).Style.BackPicFile=   ""
         Sections(2).Cells(2).Style.ForePicFile=   ""
         Sections(2).Cells(2).Style.BackPicVertPlacement=   0
         Sections(2).Cells(2).Style.BackPicHorzPlacement=   0
         Sections(2).Cells(2).Style.ForePicPlacement=   0
         Sections(2).Cells(2).Style.ForePicDrawMode=   0
         Sections(2).Cells(2).Style.MarginLeft=   6
         Sections(2).Cells(2).Style.MarginTop=   0
         Sections(2).Cells(2).Style.MarginRight=   6
         Sections(2).Cells(2).Style.MarginBottom=   0
         Sections(2).Cells(2).Style.HasBorders=   -1  'True
         Sections(2).Cells(2).Style.BorderHT=   ""
         Sections(2).Cells(2).Style.BorderHI=   ""
         Sections(2).Cells(2).Style.BorderHB=   ""
         Sections(2).Cells(2).Style.BorderVL=   "Single"
         Sections(2).Cells(2).Style.BorderVI=   "Single"
         Sections(2).Cells(2).Style.BorderVR=   "Single"
         Sections(2).Cells(2).Style.NoClipping=   0   'False
         Sections(2).Cells(2).Style.RTF=   0   'False
         Sections(2).Cells(2).Style.fprops=   1835013
         Sections(2).Cells(3).Name=   "CELL_5"
         Sections(2).Cells(3).Exp=   "Kredit"
         Sections(2).Cells(3).Width=   13
         Sections(2).Cells(3).PrivateStyle=   -1  'True
         Sections(2).Cells(3).Format=   "###,###,###,###,###,##0.00"
         Sections(2).Cells(3).Style.Name=   "<private>"
         Sections(2).Cells(3).Style.ParentName=   "Tdb_Body"
         Sections(2).Cells(3).Style.Font_Name=   "Courier"
         Sections(2).Cells(3).Style.Font_Size=   9.75
         Sections(2).Cells(3).Style.Font_Bold=   0   'False
         Sections(2).Cells(3).Style.Font_Italic=   0   'False
         Sections(2).Cells(3).Style.Font_Underline=   0   'False
         Sections(2).Cells(3).Style.Font_Strikeout=   0   'False
         Sections(2).Cells(3).Style.Font_Charset=   0
         Sections(2).Cells(3).Style.TextAlign=   2
         Sections(2).Cells(3).Style.TextVAlign=   1
         Sections(2).Cells(3).Style.TextWrap=   -1  'True
         Sections(2).Cells(3).Style.ForeColor=   0
         Sections(2).Cells(3).Style.BackColor=   16777215
         Sections(2).Cells(3).Style.NoFill=   -1  'True
         Sections(2).Cells(3).Style.BackPicFile=   ""
         Sections(2).Cells(3).Style.ForePicFile=   ""
         Sections(2).Cells(3).Style.BackPicVertPlacement=   0
         Sections(2).Cells(3).Style.BackPicHorzPlacement=   0
         Sections(2).Cells(3).Style.ForePicPlacement=   0
         Sections(2).Cells(3).Style.ForePicDrawMode=   0
         Sections(2).Cells(3).Style.MarginLeft=   6
         Sections(2).Cells(3).Style.MarginTop=   0
         Sections(2).Cells(3).Style.MarginRight=   6
         Sections(2).Cells(3).Style.MarginBottom=   0
         Sections(2).Cells(3).Style.HasBorders=   -1  'True
         Sections(2).Cells(3).Style.BorderHT=   ""
         Sections(2).Cells(3).Style.BorderHI=   ""
         Sections(2).Cells(3).Style.BorderHB=   ""
         Sections(2).Cells(3).Style.BorderVL=   "Single"
         Sections(2).Cells(3).Style.BorderVI=   "Single"
         Sections(2).Cells(3).Style.BorderVR=   "Single"
         Sections(2).Cells(3).Style.NoClipping=   0   'False
         Sections(2).Cells(3).Style.RTF=   0   'False
         Sections(2).Cells(3).Style.fprops=   1835009
         Sections(2).Cells(4).Name=   "CELL_4"
         Sections(2).Cells(4).Exp=   "JatuhTempo"
         Sections(2).Cells(4).Width=   13
         Sections(2).Cells(4).PrivateStyle=   -1  'True
         Sections(2).Cells(4).Style.Name=   "<private>"
         Sections(2).Cells(4).Style.ParentName=   "Tdb_Body"
         Sections(2).Cells(4).Style.Font_Name=   "Courier"
         Sections(2).Cells(4).Style.Font_Size=   9.75
         Sections(2).Cells(4).Style.Font_Bold=   0   'False
         Sections(2).Cells(4).Style.Font_Italic=   0   'False
         Sections(2).Cells(4).Style.Font_Underline=   0   'False
         Sections(2).Cells(4).Style.Font_Strikeout=   0   'False
         Sections(2).Cells(4).Style.Font_Charset=   0
         Sections(2).Cells(4).Style.TextAlign=   2
         Sections(2).Cells(4).Style.TextVAlign=   1
         Sections(2).Cells(4).Style.TextWrap=   0   'False
         Sections(2).Cells(4).Style.ForeColor=   0
         Sections(2).Cells(4).Style.BackColor=   16777215
         Sections(2).Cells(4).Style.NoFill=   -1  'True
         Sections(2).Cells(4).Style.BackPicFile=   ""
         Sections(2).Cells(4).Style.ForePicFile=   ""
         Sections(2).Cells(4).Style.BackPicVertPlacement=   0
         Sections(2).Cells(4).Style.BackPicHorzPlacement=   0
         Sections(2).Cells(4).Style.ForePicPlacement=   0
         Sections(2).Cells(4).Style.ForePicDrawMode=   0
         Sections(2).Cells(4).Style.MarginLeft=   6
         Sections(2).Cells(4).Style.MarginTop=   0
         Sections(2).Cells(4).Style.MarginRight=   6
         Sections(2).Cells(4).Style.MarginBottom=   0
         Sections(2).Cells(4).Style.HasBorders=   -1  'True
         Sections(2).Cells(4).Style.BorderHT=   ""
         Sections(2).Cells(4).Style.BorderHI=   ""
         Sections(2).Cells(4).Style.BorderHB=   ""
         Sections(2).Cells(4).Style.BorderVL=   "Single"
         Sections(2).Cells(4).Style.BorderVI=   "Single"
         Sections(2).Cells(4).Style.BorderVR=   "Single"
         Sections(2).Cells(4).Style.NoClipping=   0   'False
         Sections(2).Cells(4).Style.RTF=   0   'False
         Sections(2).Cells(4).Style.fprops=   1835013
         Sections(3).Name=   "SECTION_7"
         Sections(3).Type=   5
         Sections(3).Condition=   "IsLastRec()"
         Sections(3).Cells.Count=   5
         Sections(3).Cells(0).Name=   "CELL_0"
         Sections(3).Cells(0).PrivateStyle=   -1  'True
         Sections(3).Cells(0).Style.Name=   "<private>"
         Sections(3).Cells(0).Style.ParentName=   "<null>"
         Sections(3).Cells(0).Style.Font_Name=   "Times New Roman"
         Sections(3).Cells(0).Style.Font_Size=   10
         Sections(3).Cells(0).Style.Font_Bold=   0   'False
         Sections(3).Cells(0).Style.Font_Italic=   0   'False
         Sections(3).Cells(0).Style.Font_Underline=   0   'False
         Sections(3).Cells(0).Style.Font_Strikeout=   0   'False
         Sections(3).Cells(0).Style.Font_Charset=   1
         Sections(3).Cells(0).Style.TextAlign=   3
         Sections(3).Cells(0).Style.TextVAlign=   0
         Sections(3).Cells(0).Style.TextWrap=   -1  'True
         Sections(3).Cells(0).Style.ForeColor=   0
         Sections(3).Cells(0).Style.BackColor=   16777215
         Sections(3).Cells(0).Style.NoFill=   -1  'True
         Sections(3).Cells(0).Style.BackPicFile=   ""
         Sections(3).Cells(0).Style.ForePicFile=   ""
         Sections(3).Cells(0).Style.BackPicVertPlacement=   0
         Sections(3).Cells(0).Style.BackPicHorzPlacement=   0
         Sections(3).Cells(0).Style.ForePicPlacement=   0
         Sections(3).Cells(0).Style.ForePicDrawMode=   0
         Sections(3).Cells(0).Style.MarginLeft=   6
         Sections(3).Cells(0).Style.MarginTop=   6
         Sections(3).Cells(0).Style.MarginRight=   6
         Sections(3).Cells(0).Style.MarginBottom=   6
         Sections(3).Cells(0).Style.HasBorders=   -1  'True
         Sections(3).Cells(0).Style.BorderHT=   ""
         Sections(3).Cells(0).Style.BorderHI=   ""
         Sections(3).Cells(0).Style.BorderHB=   "Single"
         Sections(3).Cells(0).Style.BorderVL=   "Single"
         Sections(3).Cells(0).Style.BorderVI=   ""
         Sections(3).Cells(0).Style.BorderVR=   ""
         Sections(3).Cells(0).Style.NoClipping=   0   'False
         Sections(3).Cells(0).Style.RTF=   0   'False
         Sections(3).Cells(0).Style.fprops=   1441792
         Sections(3).Cells(1).Name=   "CELL_1"
         Sections(3).Cells(1).PrivateStyle=   -1  'True
         Sections(3).Cells(1).Style.Name=   "<private>"
         Sections(3).Cells(1).Style.ParentName=   "<null>"
         Sections(3).Cells(1).Style.Font_Name=   "Times New Roman"
         Sections(3).Cells(1).Style.Font_Size=   10
         Sections(3).Cells(1).Style.Font_Bold=   0   'False
         Sections(3).Cells(1).Style.Font_Italic=   0   'False
         Sections(3).Cells(1).Style.Font_Underline=   0   'False
         Sections(3).Cells(1).Style.Font_Strikeout=   0   'False
         Sections(3).Cells(1).Style.Font_Charset=   1
         Sections(3).Cells(1).Style.TextAlign=   3
         Sections(3).Cells(1).Style.TextVAlign=   0
         Sections(3).Cells(1).Style.TextWrap=   -1  'True
         Sections(3).Cells(1).Style.ForeColor=   0
         Sections(3).Cells(1).Style.BackColor=   16777215
         Sections(3).Cells(1).Style.NoFill=   -1  'True
         Sections(3).Cells(1).Style.BackPicFile=   ""
         Sections(3).Cells(1).Style.ForePicFile=   ""
         Sections(3).Cells(1).Style.BackPicVertPlacement=   0
         Sections(3).Cells(1).Style.BackPicHorzPlacement=   0
         Sections(3).Cells(1).Style.ForePicPlacement=   0
         Sections(3).Cells(1).Style.ForePicDrawMode=   0
         Sections(3).Cells(1).Style.MarginLeft=   6
         Sections(3).Cells(1).Style.MarginTop=   6
         Sections(3).Cells(1).Style.MarginRight=   6
         Sections(3).Cells(1).Style.MarginBottom=   6
         Sections(3).Cells(1).Style.HasBorders=   -1  'True
         Sections(3).Cells(1).Style.BorderHT=   ""
         Sections(3).Cells(1).Style.BorderHI=   ""
         Sections(3).Cells(1).Style.BorderHB=   "Single"
         Sections(3).Cells(1).Style.BorderVL=   ""
         Sections(3).Cells(1).Style.BorderVI=   ""
         Sections(3).Cells(1).Style.BorderVR=   ""
         Sections(3).Cells(1).Style.NoClipping=   0   'False
         Sections(3).Cells(1).Style.RTF=   0   'False
         Sections(3).Cells(1).Style.fprops=   1441792
         Sections(3).Cells(2).Name=   "CELL_2"
         Sections(3).Cells(2).PrivateStyle=   -1  'True
         Sections(3).Cells(2).Style.Name=   "<private>"
         Sections(3).Cells(2).Style.ParentName=   "<null>"
         Sections(3).Cells(2).Style.Font_Name=   "Times New Roman"
         Sections(3).Cells(2).Style.Font_Size=   10
         Sections(3).Cells(2).Style.Font_Bold=   0   'False
         Sections(3).Cells(2).Style.Font_Italic=   0   'False
         Sections(3).Cells(2).Style.Font_Underline=   0   'False
         Sections(3).Cells(2).Style.Font_Strikeout=   0   'False
         Sections(3).Cells(2).Style.Font_Charset=   1
         Sections(3).Cells(2).Style.TextAlign=   3
         Sections(3).Cells(2).Style.TextVAlign=   0
         Sections(3).Cells(2).Style.TextWrap=   -1  'True
         Sections(3).Cells(2).Style.ForeColor=   0
         Sections(3).Cells(2).Style.BackColor=   16777215
         Sections(3).Cells(2).Style.NoFill=   -1  'True
         Sections(3).Cells(2).Style.BackPicFile=   ""
         Sections(3).Cells(2).Style.ForePicFile=   ""
         Sections(3).Cells(2).Style.BackPicVertPlacement=   0
         Sections(3).Cells(2).Style.BackPicHorzPlacement=   0
         Sections(3).Cells(2).Style.ForePicPlacement=   0
         Sections(3).Cells(2).Style.ForePicDrawMode=   0
         Sections(3).Cells(2).Style.MarginLeft=   6
         Sections(3).Cells(2).Style.MarginTop=   6
         Sections(3).Cells(2).Style.MarginRight=   6
         Sections(3).Cells(2).Style.MarginBottom=   6
         Sections(3).Cells(2).Style.HasBorders=   -1  'True
         Sections(3).Cells(2).Style.BorderHT=   ""
         Sections(3).Cells(2).Style.BorderHI=   ""
         Sections(3).Cells(2).Style.BorderHB=   "Single"
         Sections(3).Cells(2).Style.BorderVL=   ""
         Sections(3).Cells(2).Style.BorderVI=   ""
         Sections(3).Cells(2).Style.BorderVR=   ""
         Sections(3).Cells(2).Style.NoClipping=   0   'False
         Sections(3).Cells(2).Style.RTF=   0   'False
         Sections(3).Cells(2).Style.fprops=   1441792
         Sections(3).Cells(3).Name=   "CELL_5"
         Sections(3).Cells(3).PrivateStyle=   -1  'True
         Sections(3).Cells(3).Style.Name=   "<private>"
         Sections(3).Cells(3).Style.ParentName=   "<null>"
         Sections(3).Cells(3).Style.Font_Name=   "Times New Roman"
         Sections(3).Cells(3).Style.Font_Size=   10
         Sections(3).Cells(3).Style.Font_Bold=   0   'False
         Sections(3).Cells(3).Style.Font_Italic=   0   'False
         Sections(3).Cells(3).Style.Font_Underline=   0   'False
         Sections(3).Cells(3).Style.Font_Strikeout=   0   'False
         Sections(3).Cells(3).Style.Font_Charset=   1
         Sections(3).Cells(3).Style.TextAlign=   3
         Sections(3).Cells(3).Style.TextVAlign=   0
         Sections(3).Cells(3).Style.TextWrap=   -1  'True
         Sections(3).Cells(3).Style.ForeColor=   0
         Sections(3).Cells(3).Style.BackColor=   16777215
         Sections(3).Cells(3).Style.NoFill=   -1  'True
         Sections(3).Cells(3).Style.BackPicFile=   ""
         Sections(3).Cells(3).Style.ForePicFile=   ""
         Sections(3).Cells(3).Style.BackPicVertPlacement=   0
         Sections(3).Cells(3).Style.BackPicHorzPlacement=   0
         Sections(3).Cells(3).Style.ForePicPlacement=   0
         Sections(3).Cells(3).Style.ForePicDrawMode=   0
         Sections(3).Cells(3).Style.MarginLeft=   6
         Sections(3).Cells(3).Style.MarginTop=   6
         Sections(3).Cells(3).Style.MarginRight=   6
         Sections(3).Cells(3).Style.MarginBottom=   6
         Sections(3).Cells(3).Style.HasBorders=   -1  'True
         Sections(3).Cells(3).Style.BorderHT=   ""
         Sections(3).Cells(3).Style.BorderHI=   ""
         Sections(3).Cells(3).Style.BorderHB=   "Single"
         Sections(3).Cells(3).Style.BorderVL=   ""
         Sections(3).Cells(3).Style.BorderVI=   ""
         Sections(3).Cells(3).Style.BorderVR=   ""
         Sections(3).Cells(3).Style.NoClipping=   0   'False
         Sections(3).Cells(3).Style.RTF=   0   'False
         Sections(3).Cells(3).Style.fprops=   131072
         Sections(3).Cells(4).Name=   "CELL_3"
         Sections(3).Cells(4).PrivateStyle=   -1  'True
         Sections(3).Cells(4).Style.Name=   "<private>"
         Sections(3).Cells(4).Style.ParentName=   "<null>"
         Sections(3).Cells(4).Style.Font_Name=   "Times New Roman"
         Sections(3).Cells(4).Style.Font_Size=   10
         Sections(3).Cells(4).Style.Font_Bold=   0   'False
         Sections(3).Cells(4).Style.Font_Italic=   0   'False
         Sections(3).Cells(4).Style.Font_Underline=   0   'False
         Sections(3).Cells(4).Style.Font_Strikeout=   0   'False
         Sections(3).Cells(4).Style.Font_Charset=   1
         Sections(3).Cells(4).Style.TextAlign=   3
         Sections(3).Cells(4).Style.TextVAlign=   0
         Sections(3).Cells(4).Style.TextWrap=   -1  'True
         Sections(3).Cells(4).Style.ForeColor=   0
         Sections(3).Cells(4).Style.BackColor=   16777215
         Sections(3).Cells(4).Style.NoFill=   -1  'True
         Sections(3).Cells(4).Style.BackPicFile=   ""
         Sections(3).Cells(4).Style.ForePicFile=   ""
         Sections(3).Cells(4).Style.BackPicVertPlacement=   0
         Sections(3).Cells(4).Style.BackPicHorzPlacement=   0
         Sections(3).Cells(4).Style.ForePicPlacement=   0
         Sections(3).Cells(4).Style.ForePicDrawMode=   0
         Sections(3).Cells(4).Style.MarginLeft=   6
         Sections(3).Cells(4).Style.MarginTop=   6
         Sections(3).Cells(4).Style.MarginRight=   6
         Sections(3).Cells(4).Style.MarginBottom=   6
         Sections(3).Cells(4).Style.HasBorders=   -1  'True
         Sections(3).Cells(4).Style.BorderHT=   ""
         Sections(3).Cells(4).Style.BorderHI=   ""
         Sections(3).Cells(4).Style.BorderHB=   "Single"
         Sections(3).Cells(4).Style.BorderVL=   ""
         Sections(3).Cells(4).Style.BorderVI=   ""
         Sections(3).Cells(4).Style.BorderVR=   "Single"
         Sections(3).Cells(4).Style.NoClipping=   0   'False
         Sections(3).Cells(4).Style.RTF=   0   'False
         Sections(3).Cells(4).Style.fprops=   1441792
         Sections(4).Name=   "SECTION_6"
         Sections(4).Type=   5
         Sections(4).Condition=   "IsLastRec()=false"
         Sections(4).StyleExp=   "'STYLE_1'"
         Sections(4).AutoHeight=   0   'False
         Sections(4).Height=   5
         Sections(4).Cells.Count=   1
         Sections(4).Cells(0).Name=   "CELL_1"
         Sections(4).Cells(0).Exp=   "IIF(IsLastRec(),"""",""Continued to page.."" & PageNo()+1)"
         Sections(4).Cells(0).NewLine=   -1  'True
         Sections(4).Cells(0).PrivateStyle=   -1  'True
         Sections(4).Cells(0).Style.Name=   "<private>"
         Sections(4).Cells(0).Style.ParentName=   "STYLE_1"
         Sections(4).Cells(0).Style.Font_Name=   "Courier"
         Sections(4).Cells(0).Style.Font_Size=   9.75
         Sections(4).Cells(0).Style.Font_Bold=   0   'False
         Sections(4).Cells(0).Style.Font_Italic=   0   'False
         Sections(4).Cells(0).Style.Font_Underline=   0   'False
         Sections(4).Cells(0).Style.Font_Strikeout=   0   'False
         Sections(4).Cells(0).Style.Font_Charset=   0
         Sections(4).Cells(0).Style.TextAlign=   3
         Sections(4).Cells(0).Style.TextVAlign=   1
         Sections(4).Cells(0).Style.TextWrap=   -1  'True
         Sections(4).Cells(0).Style.ForeColor=   0
         Sections(4).Cells(0).Style.BackColor=   16777215
         Sections(4).Cells(0).Style.NoFill=   -1  'True
         Sections(4).Cells(0).Style.BackPicFile=   ""
         Sections(4).Cells(0).Style.ForePicFile=   ""
         Sections(4).Cells(0).Style.BackPicVertPlacement=   0
         Sections(4).Cells(0).Style.BackPicHorzPlacement=   0
         Sections(4).Cells(0).Style.ForePicPlacement=   0
         Sections(4).Cells(0).Style.ForePicDrawMode=   0
         Sections(4).Cells(0).Style.MarginLeft=   6
         Sections(4).Cells(0).Style.MarginTop=   1
         Sections(4).Cells(0).Style.MarginRight=   6
         Sections(4).Cells(0).Style.MarginBottom=   1
         Sections(4).Cells(0).Style.HasBorders=   -1  'True
         Sections(4).Cells(0).Style.BorderHT=   "Single"
         Sections(4).Cells(0).Style.BorderHI=   ""
         Sections(4).Cells(0).Style.BorderHB=   "Single"
         Sections(4).Cells(0).Style.BorderVL=   "Single"
         Sections(4).Cells(0).Style.BorderVI=   ""
         Sections(4).Cells(0).Style.BorderVR=   "Single"
         Sections(4).Cells(0).Style.NoClipping=   0   'False
         Sections(4).Cells(0).Style.RTF=   0   'False
         Sections(4).Cells(0).Style.fprops=   1474560
         Sections(5).Name=   "SECTION_3"
         Sections(5).Condition=   "IsLastRec()"
         Sections(5).StyleExp=   "'STYLE_1'"
         Sections(5).AutoHeight=   0   'False
         Sections(5).Height=   5
         Sections(5).Cells.Count=   8
         Sections(5).Cells(0).Name=   "CELL_0"
         Sections(5).Cells(0).Exp=   """                 Kasir"""
         Sections(5).Cells(0).NewLine=   -1  'True
         Sections(5).Cells(0).PrivateStyle=   -1  'True
         Sections(5).Cells(0).Style.Name=   "<private>"
         Sections(5).Cells(0).Style.ParentName=   "STYLE_1"
         Sections(5).Cells(0).Style.Font_Name=   "Courier"
         Sections(5).Cells(0).Style.Font_Size=   9.75
         Sections(5).Cells(0).Style.Font_Bold=   0   'False
         Sections(5).Cells(0).Style.Font_Italic=   0   'False
         Sections(5).Cells(0).Style.Font_Underline=   0   'False
         Sections(5).Cells(0).Style.Font_Strikeout=   0   'False
         Sections(5).Cells(0).Style.Font_Charset=   0
         Sections(5).Cells(0).Style.TextAlign=   3
         Sections(5).Cells(0).Style.TextVAlign=   1
         Sections(5).Cells(0).Style.TextWrap=   -1  'True
         Sections(5).Cells(0).Style.ForeColor=   0
         Sections(5).Cells(0).Style.BackColor=   16777215
         Sections(5).Cells(0).Style.NoFill=   -1  'True
         Sections(5).Cells(0).Style.BackPicFile=   ""
         Sections(5).Cells(0).Style.ForePicFile=   ""
         Sections(5).Cells(0).Style.BackPicVertPlacement=   0
         Sections(5).Cells(0).Style.BackPicHorzPlacement=   0
         Sections(5).Cells(0).Style.ForePicPlacement=   0
         Sections(5).Cells(0).Style.ForePicDrawMode=   0
         Sections(5).Cells(0).Style.MarginLeft=   6
         Sections(5).Cells(0).Style.MarginTop=   1
         Sections(5).Cells(0).Style.MarginRight=   6
         Sections(5).Cells(0).Style.MarginBottom=   1
         Sections(5).Cells(0).Style.HasBorders=   -1  'True
         Sections(5).Cells(0).Style.BorderHT=   "Single"
         Sections(5).Cells(0).Style.BorderHI=   ""
         Sections(5).Cells(0).Style.BorderHB=   ""
         Sections(5).Cells(0).Style.BorderVL=   ""
         Sections(5).Cells(0).Style.BorderVI=   ""
         Sections(5).Cells(0).Style.BorderVR=   ""
         Sections(5).Cells(0).Style.NoClipping=   0   'False
         Sections(5).Cells(0).Style.RTF=   0   'False
         Sections(5).Cells(0).Style.fprops=   294912
         Sections(5).Cells(1).Name=   "CELL_15"
         Sections(5).Cells(1).Exp=   """                           """
         Sections(5).Cells(1).PrivateStyle=   -1  'True
         Sections(5).Cells(1).Style.Name=   "<private>"
         Sections(5).Cells(1).Style.ParentName=   "STYLE_1"
         Sections(5).Cells(1).Style.Font_Name=   "Courier"
         Sections(5).Cells(1).Style.Font_Size=   9.75
         Sections(5).Cells(1).Style.Font_Bold=   0   'False
         Sections(5).Cells(1).Style.Font_Italic=   0   'False
         Sections(5).Cells(1).Style.Font_Underline=   0   'False
         Sections(5).Cells(1).Style.Font_Strikeout=   0   'False
         Sections(5).Cells(1).Style.Font_Charset=   0
         Sections(5).Cells(1).Style.TextAlign=   3
         Sections(5).Cells(1).Style.TextVAlign=   1
         Sections(5).Cells(1).Style.TextWrap=   -1  'True
         Sections(5).Cells(1).Style.ForeColor=   0
         Sections(5).Cells(1).Style.BackColor=   16777215
         Sections(5).Cells(1).Style.NoFill=   -1  'True
         Sections(5).Cells(1).Style.BackPicFile=   ""
         Sections(5).Cells(1).Style.ForePicFile=   ""
         Sections(5).Cells(1).Style.BackPicVertPlacement=   0
         Sections(5).Cells(1).Style.BackPicHorzPlacement=   0
         Sections(5).Cells(1).Style.ForePicPlacement=   0
         Sections(5).Cells(1).Style.ForePicDrawMode=   0
         Sections(5).Cells(1).Style.MarginLeft=   6
         Sections(5).Cells(1).Style.MarginTop=   1
         Sections(5).Cells(1).Style.MarginRight=   6
         Sections(5).Cells(1).Style.MarginBottom=   1
         Sections(5).Cells(1).Style.HasBorders=   -1  'True
         Sections(5).Cells(1).Style.BorderHT=   "Single"
         Sections(5).Cells(1).Style.BorderHI=   ""
         Sections(5).Cells(1).Style.BorderHB=   ""
         Sections(5).Cells(1).Style.BorderVL=   ""
         Sections(5).Cells(1).Style.BorderVI=   ""
         Sections(5).Cells(1).Style.BorderVR=   ""
         Sections(5).Cells(1).Style.NoClipping=   0   'False
         Sections(5).Cells(1).Style.RTF=   0   'False
         Sections(5).Cells(1).Style.fprops=   294912
         Sections(5).Cells(2).Name=   "CELL_1"
         Sections(5).Cells(2).Exp=   """Total : """
         Sections(5).Cells(2).Width=   14
         Sections(5).Cells(2).PrivateStyle=   -1  'True
         Sections(5).Cells(2).Style.Name=   "<private>"
         Sections(5).Cells(2).Style.ParentName=   "STYLE_1"
         Sections(5).Cells(2).Style.Font_Name=   "Courier"
         Sections(5).Cells(2).Style.Font_Size=   9.75
         Sections(5).Cells(2).Style.Font_Bold=   0   'False
         Sections(5).Cells(2).Style.Font_Italic=   0   'False
         Sections(5).Cells(2).Style.Font_Underline=   0   'False
         Sections(5).Cells(2).Style.Font_Strikeout=   0   'False
         Sections(5).Cells(2).Style.Font_Charset=   0
         Sections(5).Cells(2).Style.TextAlign=   2
         Sections(5).Cells(2).Style.TextVAlign=   1
         Sections(5).Cells(2).Style.TextWrap=   -1  'True
         Sections(5).Cells(2).Style.ForeColor=   0
         Sections(5).Cells(2).Style.BackColor=   16777215
         Sections(5).Cells(2).Style.NoFill=   -1  'True
         Sections(5).Cells(2).Style.BackPicFile=   ""
         Sections(5).Cells(2).Style.ForePicFile=   ""
         Sections(5).Cells(2).Style.BackPicVertPlacement=   0
         Sections(5).Cells(2).Style.BackPicHorzPlacement=   0
         Sections(5).Cells(2).Style.ForePicPlacement=   0
         Sections(5).Cells(2).Style.ForePicDrawMode=   0
         Sections(5).Cells(2).Style.MarginLeft=   6
         Sections(5).Cells(2).Style.MarginTop=   1
         Sections(5).Cells(2).Style.MarginRight=   6
         Sections(5).Cells(2).Style.MarginBottom=   1
         Sections(5).Cells(2).Style.HasBorders=   -1  'True
         Sections(5).Cells(2).Style.BorderHT=   "Single"
         Sections(5).Cells(2).Style.BorderHI=   ""
         Sections(5).Cells(2).Style.BorderHB=   ""
         Sections(5).Cells(2).Style.BorderVL=   ""
         Sections(5).Cells(2).Style.BorderVI=   ""
         Sections(5).Cells(2).Style.BorderVR=   ""
         Sections(5).Cells(2).Style.NoClipping=   0   'False
         Sections(5).Cells(2).Style.RTF=   0   'False
         Sections(5).Cells(2).Style.fprops=   32769
         Sections(5).Cells(3).Name=   "CELL_2"
         Sections(5).Cells(3).Exp=   "nSubTotal"
         Sections(5).Cells(3).Width=   15
         Sections(5).Cells(3).PrivateStyle=   -1  'True
         Sections(5).Cells(3).Format=   "###,###,###,###,###,##0.00"
         Sections(5).Cells(3).Style.Name=   "<private>"
         Sections(5).Cells(3).Style.ParentName=   "STYLE_1"
         Sections(5).Cells(3).Style.Font_Name=   "Courier"
         Sections(5).Cells(3).Style.Font_Size=   9.75
         Sections(5).Cells(3).Style.Font_Bold=   0   'False
         Sections(5).Cells(3).Style.Font_Italic=   0   'False
         Sections(5).Cells(3).Style.Font_Underline=   0   'False
         Sections(5).Cells(3).Style.Font_Strikeout=   0   'False
         Sections(5).Cells(3).Style.Font_Charset=   0
         Sections(5).Cells(3).Style.TextAlign=   3
         Sections(5).Cells(3).Style.TextVAlign=   1
         Sections(5).Cells(3).Style.TextWrap=   -1  'True
         Sections(5).Cells(3).Style.ForeColor=   0
         Sections(5).Cells(3).Style.BackColor=   16777215
         Sections(5).Cells(3).Style.NoFill=   -1  'True
         Sections(5).Cells(3).Style.BackPicFile=   ""
         Sections(5).Cells(3).Style.ForePicFile=   ""
         Sections(5).Cells(3).Style.BackPicVertPlacement=   0
         Sections(5).Cells(3).Style.BackPicHorzPlacement=   0
         Sections(5).Cells(3).Style.ForePicPlacement=   0
         Sections(5).Cells(3).Style.ForePicDrawMode=   0
         Sections(5).Cells(3).Style.MarginLeft=   6
         Sections(5).Cells(3).Style.MarginTop=   1
         Sections(5).Cells(3).Style.MarginRight=   6
         Sections(5).Cells(3).Style.MarginBottom=   1
         Sections(5).Cells(3).Style.HasBorders=   -1  'True
         Sections(5).Cells(3).Style.BorderHT=   "Single"
         Sections(5).Cells(3).Style.BorderHI=   ""
         Sections(5).Cells(3).Style.BorderHB=   ""
         Sections(5).Cells(3).Style.BorderVL=   ""
         Sections(5).Cells(3).Style.BorderVI=   ""
         Sections(5).Cells(3).Style.BorderVR=   ""
         Sections(5).Cells(3).Style.NoClipping=   0   'False
         Sections(5).Cells(3).Style.RTF=   0   'False
         Sections(5).Cells(3).Style.fprops=   1081344
         Sections(5).Cells(4).Name=   "CELL_4"
         Sections(5).Cells(4).Exp=   "keAkun"
         Sections(5).Cells(4).NewLine=   -1  'True
         Sections(5).Cells(4).PrivateStyle=   -1  'True
         Sections(5).Cells(4).Style.Name=   "<private>"
         Sections(5).Cells(4).Style.ParentName=   "STYLE_1"
         Sections(5).Cells(4).Style.Font_Name=   "Courier"
         Sections(5).Cells(4).Style.Font_Size=   9.75
         Sections(5).Cells(4).Style.Font_Bold=   0   'False
         Sections(5).Cells(4).Style.Font_Italic=   0   'False
         Sections(5).Cells(4).Style.Font_Underline=   0   'False
         Sections(5).Cells(4).Style.Font_Strikeout=   0   'False
         Sections(5).Cells(4).Style.Font_Charset=   0
         Sections(5).Cells(4).Style.TextAlign=   2
         Sections(5).Cells(4).Style.TextVAlign=   1
         Sections(5).Cells(4).Style.TextWrap=   -1  'True
         Sections(5).Cells(4).Style.ForeColor=   0
         Sections(5).Cells(4).Style.BackColor=   16777215
         Sections(5).Cells(4).Style.NoFill=   -1  'True
         Sections(5).Cells(4).Style.BackPicFile=   ""
         Sections(5).Cells(4).Style.ForePicFile=   ""
         Sections(5).Cells(4).Style.BackPicVertPlacement=   0
         Sections(5).Cells(4).Style.BackPicHorzPlacement=   0
         Sections(5).Cells(4).Style.ForePicPlacement=   0
         Sections(5).Cells(4).Style.ForePicDrawMode=   0
         Sections(5).Cells(4).Style.MarginLeft=   6
         Sections(5).Cells(4).Style.MarginTop=   1
         Sections(5).Cells(4).Style.MarginRight=   6
         Sections(5).Cells(4).Style.MarginBottom=   1
         Sections(5).Cells(4).Style.HasBorders=   -1  'True
         Sections(5).Cells(4).Style.BorderHT=   ""
         Sections(5).Cells(4).Style.BorderHI=   ""
         Sections(5).Cells(4).Style.BorderHB=   ""
         Sections(5).Cells(4).Style.BorderVL=   ""
         Sections(5).Cells(4).Style.BorderVI=   ""
         Sections(5).Cells(4).Style.BorderVR=   ""
         Sections(5).Cells(4).Style.NoClipping=   0   'False
         Sections(5).Cells(4).Style.RTF=   0   'False
         Sections(5).Cells(4).Style.fprops=   1
         Sections(5).Cells(5).Name=   "CELL_5"
         Sections(5).Cells(5).NewLine=   -1  'True
         Sections(5).Cells(6).Name=   "CELL_16"
         Sections(5).Cells(6).NewLine=   -1  'True
         Sections(5).Cells(6).Height=   5
         Sections(5).Cells(6).PrivateStyle=   -1  'True
         Sections(5).Cells(6).Style.Name=   "<private>"
         Sections(5).Cells(6).Style.ParentName=   "STYLE_1"
         Sections(5).Cells(6).Style.Font_Name=   "Courier"
         Sections(5).Cells(6).Style.Font_Size=   9.75
         Sections(5).Cells(6).Style.Font_Bold=   0   'False
         Sections(5).Cells(6).Style.Font_Italic=   0   'False
         Sections(5).Cells(6).Style.Font_Underline=   0   'False
         Sections(5).Cells(6).Style.Font_Strikeout=   0   'False
         Sections(5).Cells(6).Style.Font_Charset=   0
         Sections(5).Cells(6).Style.TextAlign=   2
         Sections(5).Cells(6).Style.TextVAlign=   1
         Sections(5).Cells(6).Style.TextWrap=   -1  'True
         Sections(5).Cells(6).Style.ForeColor=   0
         Sections(5).Cells(6).Style.BackColor=   16777215
         Sections(5).Cells(6).Style.NoFill=   -1  'True
         Sections(5).Cells(6).Style.BackPicFile=   ""
         Sections(5).Cells(6).Style.ForePicFile=   ""
         Sections(5).Cells(6).Style.BackPicVertPlacement=   0
         Sections(5).Cells(6).Style.BackPicHorzPlacement=   0
         Sections(5).Cells(6).Style.ForePicPlacement=   0
         Sections(5).Cells(6).Style.ForePicDrawMode=   0
         Sections(5).Cells(6).Style.MarginLeft=   6
         Sections(5).Cells(6).Style.MarginTop=   1
         Sections(5).Cells(6).Style.MarginRight=   6
         Sections(5).Cells(6).Style.MarginBottom=   1
         Sections(5).Cells(6).Style.HasBorders=   -1  'True
         Sections(5).Cells(6).Style.BorderHT=   ""
         Sections(5).Cells(6).Style.BorderHI=   ""
         Sections(5).Cells(6).Style.BorderHB=   ""
         Sections(5).Cells(6).Style.BorderVL=   ""
         Sections(5).Cells(6).Style.BorderVI=   ""
         Sections(5).Cells(6).Style.BorderVR=   ""
         Sections(5).Cells(6).Style.NoClipping=   0   'False
         Sections(5).Cells(6).Style.RTF=   0   'False
         Sections(5).Cells(6).Style.fprops=   2064389
         Sections(5).Cells(7).Name=   "CELL_20"
         Sections(5).Cells(7).Exp=   "cFooter2"
         Sections(5).Cells(7).NewLine=   -1  'True
         Sections(5).Cells(7).PrivateStyle=   -1  'True
         Sections(5).Cells(7).Style.Name=   "<private>"
         Sections(5).Cells(7).Style.ParentName=   "STYLE_1"
         Sections(5).Cells(7).Style.Font_Name=   "Courier"
         Sections(5).Cells(7).Style.Font_Size=   9.75
         Sections(5).Cells(7).Style.Font_Bold=   0   'False
         Sections(5).Cells(7).Style.Font_Italic=   0   'False
         Sections(5).Cells(7).Style.Font_Underline=   0   'False
         Sections(5).Cells(7).Style.Font_Strikeout=   0   'False
         Sections(5).Cells(7).Style.Font_Charset=   0
         Sections(5).Cells(7).Style.TextAlign=   2
         Sections(5).Cells(7).Style.TextVAlign=   1
         Sections(5).Cells(7).Style.TextWrap=   -1  'True
         Sections(5).Cells(7).Style.ForeColor=   0
         Sections(5).Cells(7).Style.BackColor=   16777215
         Sections(5).Cells(7).Style.NoFill=   -1  'True
         Sections(5).Cells(7).Style.BackPicFile=   ""
         Sections(5).Cells(7).Style.ForePicFile=   ""
         Sections(5).Cells(7).Style.BackPicVertPlacement=   0
         Sections(5).Cells(7).Style.BackPicHorzPlacement=   0
         Sections(5).Cells(7).Style.ForePicPlacement=   0
         Sections(5).Cells(7).Style.ForePicDrawMode=   0
         Sections(5).Cells(7).Style.MarginLeft=   6
         Sections(5).Cells(7).Style.MarginTop=   1
         Sections(5).Cells(7).Style.MarginRight=   6
         Sections(5).Cells(7).Style.MarginBottom=   1
         Sections(5).Cells(7).Style.HasBorders=   -1  'True
         Sections(5).Cells(7).Style.BorderHT=   ""
         Sections(5).Cells(7).Style.BorderHI=   ""
         Sections(5).Cells(7).Style.BorderHB=   ""
         Sections(5).Cells(7).Style.BorderVL=   ""
         Sections(5).Cells(7).Style.BorderVI=   ""
         Sections(5).Cells(7).Style.BorderVR=   ""
         Sections(5).Cells(7).Style.NoClipping=   0   'False
         Sections(5).Cells(7).Style.RTF=   0   'False
         Sections(5).Cells(7).Style.fprops=   2064385
         Styles.Count    =   6
         Styles(0).Name  =   "Tdb_Base"
         Styles(0).ParentName=   ""
         Styles(0).Font_Name=   "Courier"
         Styles(0).Font_Size=   9.75
         Styles(0).Font_Bold=   -1  'True
         Styles(0).Font_Charset=   0
         Styles(0).TextVAlign=   1
         Styles(0).MarginTop=   1
         Styles(0).MarginBottom=   1
         Styles(1).Name  =   "STYLE_1"
         Styles(1).ParentName=   "Tdb_Base"
         Styles(1).Font_Name=   "Courier"
         Styles(1).Font_Size=   9.75
         Styles(1).Font_Charset=   0
         Styles(1).TextVAlign=   1
         Styles(1).MarginTop=   1
         Styles(1).MarginBottom=   1
         Styles(1).fprops=   18087936
         Styles(2).Name  =   "Tdb_Body"
         Styles(2).ParentName=   "Tdb_Base"
         Styles(2).Font_Name=   "Courier"
         Styles(2).Font_Size=   9.75
         Styles(2).Font_Charset=   0
         Styles(2).TextVAlign=   1
         Styles(2).MarginTop=   0
         Styles(2).MarginBottom=   0
         Styles(2).fprops=   18862080
         Styles(3).Name  =   "Tdb_Header"
         Styles(3).ParentName=   "Tdb_Base"
         Styles(3).Font_Name=   "Courier"
         Styles(3).Font_Size=   9.75
         Styles(3).Font_Bold=   -1  'True
         Styles(3).Font_Charset=   0
         Styles(3).TextAlign=   0
         Styles(3).TextVAlign=   1
         Styles(3).MarginTop=   1
         Styles(3).MarginBottom=   1
         Styles(3).BorderHT=   "Single"
         Styles(3).BorderHI=   "Single"
         Styles(3).BorderHB=   "Single"
         Styles(3).fprops=   2064385
         Styles(4).Name  =   "Tdb_PageFooter"
         Styles(4).ParentName=   "Tdb_Base"
         Styles(4).Font_Name=   "Courier"
         Styles(4).Font_Size=   9.75
         Styles(4).Font_Bold=   -1  'True
         Styles(4).Font_Charset=   0
         Styles(4).TextVAlign=   1
         Styles(4).MarginTop=   1
         Styles(4).MarginBottom=   1
         Styles(4).BorderHT=   "Single"
         Styles(4).fprops=   163840
         Styles(5).Name  =   "Garis"
         Styles(5).ParentName=   "Tdb_Base"
         Styles(5).Font_Name=   "Courier"
         Styles(5).Font_Size=   9.75
         Styles(5).Font_Bold=   -1  'True
         Styles(5).Font_Charset=   0
         Styles(5).TextAlign=   2
         Styles(5).TextVAlign=   1
         Styles(5).MarginTop=   1
         Styles(5).MarginBottom=   1
         Styles(5).BorderHT=   "Single"
         Styles(5).fprops=   32769
         Lines.Count     =   4
         Lines(0).Name   =   "Single"
         Lines(0).Thickness=   4
         Lines(1).Name   =   "Double"
         Lines(1).Thickness=   5
         Lines(2).Name   =   "Quarter"
         Lines(2).Thickness=   1
         Lines(2).Color  =   8421504
         Lines(3).Name   =   "None"
         Profiles.Count  =   1
         Profiles(0).Name=   "PROFILE_0"
         Profiles(0).Active=   -1  'True
         Profiles(0).PreviewNoMinimize=   -1  'True
         Profiles(0).PreviewNoMaximize=   -1  'True
         Profiles(0).PreviewNoResize=   -1  'True
         Profiles(0).PreviewMaximized=   -1  'True
         Profiles(0).PreviewNoSaveLoad=   -1  'True
         Profiles(0).PrinterMarginLeft=   10
         Profiles(0).PrinterMarginTop=   5
         Profiles(0).PrinterMarginRight=   10
         Profiles(0).PrinterMarginBottom=   5
         Profiles(0).PrinterPaperSize=   256
         Profiles(0).PrinterPaperHeight=   139
         Profiles(0).PrinterPaperWidth=   215
         Profiles(0).PrinterMargins_set=   -1  'True
         Profiles(0).PrinterPaperSize_set=   -1  'True
         Profiles(0).PrinterPaperUserSize_set=   -1  'True
      End
      Begin BiSATextBoxProject.BiSABrowse cKodeKeteranganBayar 
         Height          =   345
         Left            =   1845
         TabIndex        =   27
         Top             =   3180
         Width           =   2835
         _ExtentX        =   5001
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
         Appearance      =   0
         Button          =   -1  'True
         Caption         =   "Ket. Bayar"
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
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1995
         TabIndex        =   21
         Top             =   675
         Width           =   3555
      End
      Begin VB.Line Line1 
         X1              =   1635
         X2              =   5595
         Y1              =   2355
         Y2              =   2355
      End
      Begin VB.Label Label1 
         Caption         =   "PEMBAYARAN"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3510
         TabIndex        =   14
         Top             =   255
         Width           =   2100
      End
   End
End
Attribute VB_Name = "trPelunasanHutangSederhana"
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
Dim vaVoucher As New XArrayDB
Dim lEdit As Boolean
Dim cPosKas As String
Dim Prospective As New TrueOleDBGrid70.Style
Dim Distributors As New TrueOleDBGrid70.Style



Public nDebet As Double
Public lPubStatus As Boolean
Public vaPubReff As New XArrayDB
Public nPubTotal As Double

Public lClose As Double
Public nTarikTunai As Double
Public nWithDraw As Double
Public nSisaKurangTopUp As Double
Public lTarikTunai As Boolean
Public nSaldoTopUp As Double

Public nKembalian As Double
Public nJaminan As Double
Public nTotYgHarusDibayar As Double
Public nMetodePembayaran As Integer

Sub GetEdit(lPar As Boolean)
  lEdit = lPar
  BiSAFrame1.Enabled = lPar
  BiSAFrame2.Enabled = lPar
  BiSAFrame3.Enabled = lPar
  SetButton cmdSimpan, cmdKeluar, cmdAdd, cmdEdit, cmdHapus, nPos, lPar, cmdAktivasi
End Sub

Private Sub BiSAButton1_Click()
Dim n As Single
Dim nTotDebet As Double
Dim nTotKredit As Double

  vaArray.ReDim 0, -1, 0, 8
  nTotDebet = 0
  nTotKredit = 0
  'ambil saldo awal sebelum periode
  Set dbData = objData.Browse(GetDSN, "kartupiutang", "sum(debet) as totdebetawal,  sum(kredit) as totkreditawal", "kodeanggota", sisAssign, cCustomer.Text, " and tgl < '" & Format(dTglMutasi(0).value, "yyyy-MM-dd") & "' and groupsales ='" & GetRegistry(reg_KodeGroupPenjualan) & "'")
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = 0
      vaArray(n, 1) = ""
      vaArray(n, 2) = ""
      vaArray(n, 3) = ""
      vaArray(n, 4) = "Saldo sebelum periode"
      vaArray(n, 5) = GetNull(dbData!totdebetawal)
      vaArray(n, 6) = GetNull(dbData!totkreditawal)
      vaArray(n, 7) = sisFlag.Nul
      vaArray(n, 8) = ""
      nTotDebet = nTotDebet + vaArray(n, 5)
      nTotKredit = nTotKredit + vaArray(n, 6)
      dbData.MoveNext
    Loop
  End If
  Set dbData = objData.Browse(GetDSN, "kartupiutang", , "kodeanggota", sisAssign, cCustomer.Text, " and tgl >= '" & Format(dTglMutasi(0).value, "yyyy-MM-dd") & "' and tgl <= '" & Format(dTglMutasi(1).value, "yyyy-MM-dd") & "' and groupsales ='" & GetRegistry(reg_KodeGroupPenjualan) & "' order by tgl,id asc")
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = IIf(GetNull(dbData!flag) = sisFlag.Posting, TDBGrid1.Columns(0).BackColor = vbButtonFace, 0)
      vaArray(n, 1) = IIf(GetNull(dbData!flag) = sisFlag.Posting, "POST", "")
      vaArray(n, 2) = GetNull(dbData!nomorkartupiutang)
      vaArray(n, 3) = GetNull(dbData!tgl)
      vaArray(n, 4) = GetNull(dbData!keterangan)
      vaArray(n, 5) = GetNull(dbData!debet)
      vaArray(n, 6) = GetNull(dbData!kredit)
      vaArray(n, 7) = GetNull(dbData!flag)
      vaArray(n, 8) = GetNull(dbData!flagid)
      nTotDebet = nTotDebet + vaArray(n, 5)
      nTotKredit = nTotKredit + vaArray(n, 6)
      dbData.MoveNext
    Loop
  End If
  'ambil saldo setelah periode
  Set dbData = objData.Browse(GetDSN, "kartupiutang", "sum(debet) as totdebetakhir, sum(kredit) as totkreditakhir", "kodeanggota", sisAssign, cCustomer.Text, " and tgl > '" & Format(dTglMutasi(1).value, "yyyy-MM-dd") & "' and groupsales ='" & GetRegistry(reg_KodeGroupPenjualan) & "'")
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = 0
      vaArray(n, 1) = ""
      vaArray(n, 2) = ""
      vaArray(n, 3) = ""
      vaArray(n, 4) = "Saldo Setelah periode"
      vaArray(n, 5) = GetNull(dbData!totdebetakhir)
      vaArray(n, 6) = GetNull(dbData!totkreditakhir)
      vaArray(n, 7) = sisFlag.Nul
      vaArray(n, 8) = ""
      nTotDebet = nTotDebet + vaArray(n, 5)
      nTotKredit = nTotKredit + vaArray(n, 6)
      dbData.MoveNext
    Loop
  End If
  'ambil data voucher yg masih ada
  vaVoucher.ReDim 0, -1, 0, 6
  Set dbData = objData.Browse(GetDSN, "membertopup", , "kodeanggota", sisAssign, cCustomer.Text, " and lstatus = '0'")
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaVoucher.InsertRows vaVoucher.UpperBound(1) + 1
      n = vaVoucher.UpperBound(1)
      vaVoucher(n, 0) = 0
      vaVoucher(n, 1) = GetNull(dbData!nomormembertopup)
      vaVoucher(n, 2) = GetNull(dbData!tgl)
      vaVoucher(n, 3) = GetNull(dbData!keterangan)
      vaVoucher(n, 4) = GetNull(dbData!debet)
      vaVoucher(n, 5) = GetNull(dbData!lStatus)
      vaVoucher(n, 6) = GetNull(dbData!lstatusid)
      dbData.MoveNext
    Loop
  End If
  Set GridVoucher.Array = vaVoucher
  GridVoucher.ReBind
  GridVoucher.Refresh
  
  nOutstanding.value = GetSaldoPiutang(objData, cCustomer.Text)
  nSaldoPiutang.value = GetSaldoPiutang(objData, cCustomer.Text, GetRegistry(reg_KodeGroupPenjualan))
  TDBGrid1.FetchRowStyle = True
  Set TDBGrid1.Array = vaArray
  TDBGrid1.Columns(5).FooterText = Format(nTotDebet, "###,###,###,##0.00")
  TDBGrid1.Columns(6).FooterText = Format(nTotKredit, "###,###,###,##0.00")
  TDBGrid1.ReBind
  TDBGrid1.Refresh
  
  
  GetSumTunai
  nTunai_Change
End Sub

Private Sub cAkunKas_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", "kodeakun,keterangan", "jenis", sisAssign, "D", " and left(kodeakun,1)='1' and keterangan like '%" & cAkunKas.Text & "%'")
  If Not dbData.EOF Then
    cAkunKas.Text = cAkunKas.Browse(dbData, Array("Kode Akun", "Keterangan"), , Array(25, 30))
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
  nSaldoPiutang.value = 0
  Set dbData = objData.Browse(GetDSN, "totpenjualan", "nomorpenjualan,tgl,piutang,jthtmp", "kodeanggota", sisAssign, cCustomer.Text, " and flaglunas = 0", "tgl desc")
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      If Not isLunas(objData, GetNull(dbData!nomorpenjualan), nSisaPiutang) Then
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        n = vaArray.UpperBound(1)
        vaArray(n, 0) = 0
        vaArray(n, 1) = n + 1
        vaArray(n, 2) = GetNull(dbData!nomorpenjualan)
        vaArray(n, 3) = GetNull(dbData!tgl)
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
    nSaldoPiutang.value = GetNull(dbData!saldopiutang)
  End If

  'cari saldo top up member
'  Set dbData = objData.Browse(GetDSN, "membertopup m", "m.kodeanggota,a.nama,a.alamat,sum(debet) as debet,sum(kredit) as kredit,sum(m.debet-m.kredit) as saldo", "m.kodeanggota", sisAssign, cCustomer.Text, " GROUP BY m.kodeanggota", , Array("left join anggota a on a.kodeanggota = m.kodeanggota"))
'  If Not dbData.EOF Then
'    nSaldoTopUpMember.Value = GetNull(dbData!saldo)
'  Else
'    nSaldoTopUpMember.Value = 0
'  End If
  
End Sub

Private Sub cFaktur_ButtonClick()
Dim n As Integer
Dim lSave As Boolean

lSave = True

  Set dbData = objData.Browse(GetDSN, "totpelunasanpiutangsederhana", , "kodeanggota", sisAssign, cCustomer.Text, " and username='" & GetRegistry(reg_Username) & "' and tgl = '" & Format(dTanggal.value, "yyyy-MM-dd") & "'")
  If Not dbData.EOF Then
    'jika sudah dipilih maka tampilkan datanya
    cFaktur.Text = cFaktur.Browse(dbData)
    nVoucher.value = GetNull(dbData!voucher)
    nTunai.value = GetNull(dbData!Tunai)
    nTotal.value = GetNull(dbData!Total)
    BiSAButton1_Click
    Me.Refresh
    If MsgBox("Data akan dihapus", vbYesNo + vbCritical) = vbYes Then
      'rutin penghapusan
      lSave = True
      objData.Start GetDSN
      lStart = True
      
      'Hapus
      lSave = IIf(lSave, DelKodeTr(objData, msPelunasanPiutangSederhana, cFaktur.Text), False)
      lSave = IIf(lSave, objData.Delete(GetDSN, "totpelunasanpiutangsederhana", "nomorpelunasanpiutang", sisAssign, cFaktur.Text), False)
      lSave = IIf(lSave, objData.Delete(GetDSN, "kartupiutang", "nomorkartupiutang", sisAssign, cFaktur.Text), False)
      lSave = IIf(lSave, objData.Delete(GetDSN, "membertopup", "nomormembertopup ", sisAssign, cFaktur.Text), False)
      lSave = IIf(lSave, objData.Edit(GetDSN, "kartupiutang", "flagid='" & cFaktur.Text & "'", Array("flag", "flagid"), Array(sisFlag.Nul, "")), False)
      
      'update status di voucher, jika ada
      Set dbData = objData.Browse(GetDSN, "membertopup", , "lstatusid", sisAssign, cFaktur.Text)
      If Not dbData.EOF Then
        Do While Not dbData.EOF
          lSave = IIf(lSave, objData.Edit(GetDSN, "membertopup", " nomormembertopup = '" & GetNull(dbData!nomormembertopup) & "'", Array("lstatus", "lstatusid"), Array(sisFlag.Nul, "")), False)
          dbData.MoveNext
        Loop
      End If
      
      If lSave Then
        objData.Save GetDSN
      Else
        MsgBox "Maaf data tidak berhasil dihapus", vbExclamation
        objData.Cancel GetDSN
      End If
      initvalue
      GetEdit False
    
    End If
  End If

End Sub

Private Sub cKodeKeteranganBayar_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "keteranganbayar", "kode,keterangan")
  If Not dbData.EOF Then
    cKodeKeteranganBayar.Text = cKodeKeteranganBayar.Browse(dbData, Array("Kode", "Keterangan"), , Array(25, 30))
  End If
End Sub

Private Sub cmdAdd_Click()
  nPos = Add
  GetEdit True
  initvalue
  dTanggal.SetFocus
  GetBrowseFaktur False
  cFaktur.Text = CreateNomorFaktur(objData, sisModulTransaksi.PelunasanPiutangSederhana, "totpelunasanpiutang", "nomorpelunasanpiutang")
End Sub

Private Sub cmdAktivasi_Click()
  frmAktivasi.Action Me
End Sub

Private Sub GetDataKartuPiutang()
Dim nDebet As Double
Dim nKredit As Double
Dim n As Integer

  nDebet = 0
  nKredit = 0
  vaArray.ReDim 0, -1, 0, 6
  Set dbData = objData.Browse(GetDSN, "kartupiutang", , "kodeanggota", sisAssign, cCustomer.Text)
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = 0
      vaArray(n, 1) = n
      vaArray(n, 2) = GetNull(dbData!nomorkartupiutang)
      vaArray(n, 3) = GetNull(dbData!tgl)
      vaArray(n, 4) = GetNull(dbData!keterangan)
      vaArray(n, 5) = GetNull(dbData!debet)
      vaArray(n, 6) = GetNull(dbData!kredit)
      nDebet = nDebet + vaArray(n, 5)
      nKredit = nKredit + vaArray(n, 6)
      dbData.MoveNext
    Loop
  End If
  nSaldoPiutang.value = GetSaldoPiutang(objData, cCustomer.Text)
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
  TDBGrid1.Refresh

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
'      Unload Me
'      GetEdit False
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
'      Unload Me
'      GetEdit False
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
  Unload trLunasPiutang
  If lEdit Then
    initvalue
    GetEdit False
  Else
    Unload Me
  End If
End Sub

Private Sub initvalue()
  
  nDebet = 0
  dTanggal.value = Date
  cCustomer.Default
  cNama.Default
  cFaktur.Default
  cDepartement.Default
  cAkunKas.Text = cKasTeller
  cNama.Enabled = True
  
  nVoucher.Default
  nTunai.Default
  nTotal.Default
  nSaldoPiutang.Default
  nOutstanding.Default
  cKodeKeteranganBayar.Default
  
  ClearTdbgrid
  dTglMutasi(0).value = Format(Year(Now) & "-" & Month(Now) & "-01", "yyyy-MM-dd")
  dTglMutasi(1).value = Format(Year(Now) & "-" & Month(Now) & "-28", "yyyy-MM-dd")
  lClose = True
  trPelunasanHutangSederhana.Caption = "PELUNASAN PIUTANG - GROUP SALES " & GetRegistry(reg_KodeGroupPenjualan)
  Label2.Caption = "BON : " & GetRegistry(reg_KodeGroupPenjualan)
  'Kuncu akun kas
  cAkunKas.Enabled = True
  cAkunKas.BackColor = vbWindowBackground
  cAkunKas.Button = True
  If GetRegistry(reg_UserLevel) <> 0 Then
    cAkunKas.Enabled = False
    cAkunKas.BackColor = vbButtonFace
    cAkunKas.Button = False
  End If
  TDBGrid1.Columns(5).FooterText = Format(0, "###,###,###,##0.00")
  TDBGrid1.Columns(6).FooterText = Format(0, "###,###,###,##0.00")
End Sub

Private Sub ClearTdbgrid()
  vaArray.ReDim 0, -1, 0, 6
  TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
  vaVoucher.ReDim 0, -1, 0, 6
  GridVoucher.Array = vaVoucher
  GridVoucher.ReBind
End Sub

Private Function ValidSaving() As Boolean
Dim db As New ADODB.Recordset

  ValidSaving = True
  If Trim(cKodeKeteranganBayar.Text) <> "" Then
    Set db = objData.Browse(GetDSN, "keteranganbayar", , "kode", sisAssign, cKodeKeteranganBayar.Text)
    If db.EOF Then
      MsgBox "Maaf kode keterangan bayar tidak valid", vbCritical
      ValidSaving = False
      Exit Function
    End If
  Else
    MsgBox "Err. Keteragan pembayaran belum diisi"
    ValidSaving = False
    Exit Function
  End If
End Function

Private Function validOK() As Boolean
  validOK = True
End Function

Private Sub cmdSimpan_Click()
Dim vaField
Dim vaValue
Dim n As Single
Dim lSave As Boolean
Dim Faktur As String
Dim cSQL As String
  
  
  If Not ValidSaving Then
    Exit Sub
  End If
  
  Faktur = cFaktur.Text
  If nPos = Add Then
    If Not GetAvailable(cFaktur.Text, "totpelunasanpiutangsederhana", "nomorpelunasanpiutang") Then
      Faktur = GetNomor("totpelunasanpiutangsederhana", "nomorpelunasanpiutang", GetID, sisModulTransaksi.PelunasanPiutangSederhana)
    End If
  End If

  'cek apakah ada data yg akan dilunasi
  If Trim(cFaktur.Text) = "" Then
    MsgBox "Maaf Nomor Faktur Kosong/Tidak Valid" & vbCrLf & "Data tidak bisa disimpan", vbCritical
    Exit Sub
  End If
    
  If nTotal.value > 0 And nTotal.value <= nSaldoPiutang.value And nTunai.value >= 0 Then
    lSave = True
    objData.Start GetDSN
    lStart = True
    
    'Hapus
    lSave = IIf(lSave, DelKodeTr(objData, msPelunasanPiutangSederhana, Faktur), False)
    lSave = IIf(lSave, objData.Delete(GetDSN, "totpelunasanpiutangsederhana", "nomorpelunasanpiutang", sisAssign, Faktur), False)
    lSave = IIf(lSave, objData.Delete(GetDSN, "kartupiutang", "nomorkartupiutang", sisAssign, Faktur), False)
    lSave = IIf(lSave, objData.Delete(GetDSN, "membertopup", "nomormembertopup ", sisAssign, Faktur), False)
    
    'Simpan
    lSave = IIf(lSave, objData.Update(GetDSN, "totpelunasanpiutangsederhana", "nomorpelunasanpiutang = '" & Faktur & "'", _
    Array("nomorpelunasanpiutang", "kodeanggota", "username", "kodeakun", "kodecostcenter", "tgl", "tunai", "voucher", "total", "datetime", "kodeketeranganbayar"), _
    Array(Faktur, cCustomer.Text, GetRegistry(reg_Username), cAkunKas.Text, GetCostCenterUser(objData, GetRegistry(reg_Username)), Format(dTanggal.value, "yyyy-MM-dd"), nTunai.value, nVoucher.value, nTotal.value, SNow, cKodeKeteranganBayar.Text)), False)
    
    If nVoucher.value > 0 Then
      vaField = Array("nomormembertopup", "tgl", "kodeanggota", "keterangan", "kredit", "lstatus", "lstatusid")
      vaValue = Array(Faktur, dTanggal.value, cCustomer.Text, "Pelunasan", nVoucher.value, sisFlag.Posting, Faktur)
      lSave = IIf(lSave, objData.Add(GetDSN, "membertopup", vaField, vaValue), False)
      'edit di masing2 voucher
      For n = 0 To vaVoucher.UpperBound(1)
        If vaVoucher(n, 0) = -1 Then
          lSave = IIf(lSave, objData.Edit(GetDSN, "membertopup", "nomormembertopup = '" & vaVoucher(n, 1) & "'", Array("lstatus", "lstatusid"), Array(sisFlag.Posting, Faktur)), False)
        End If
      Next n
    End If
    
    lSave = IIf(lSave, UpdKartuHutang(objData, SisKartuHutang.SisPelunasanPiutangSederhana, Faktur, dTanggal.value, cCustomer.Text, "Pelunasan Piutang an " & cNama.Text, nTotal.value), False)
    
    'Jika ada yg dicentang, update flag
    
    For n = 0 To vaArray.UpperBound(1)
      If vaArray(n, 0) = -1 Then
        lSave = IIf(lSave, objData.Edit(GetDSN, "kartupiutang", "nomorkartupiutang='" & vaArray(n, 2) & "'", Array("flag", "flagid"), Array(sisFlag.Posting, cFaktur.Text)), False)
        lSave = IIf(lSave, objData.Edit(GetDSN, "totpenjualan", "nomorpenjualan='" & vaArray(n, 2) & "'", Array("flaglunas"), Array(sisFlag.Posting, cFaktur.Text)), False)
        'kasi flag juga untuk pelunasan nya
        If nDebet = nTotal.value Then
          lSave = IIf(lSave, objData.Edit(GetDSN, "kartupiutang", "nomorkartupiutang='" & cFaktur.Text & "'", Array("flag", "flagid"), Array(sisFlag.Posting, cFaktur.Text)), False)
        End If
      End If
    Next n
    
    'update juga di table totpembelian kasi flag lunas

    'kas
    'hutang voucher
    '       piutang
    lSave = IIf(lSave, UpdKodeTr(objData, msPelunasanPiutangSederhana, Faktur, Format(dTanggal.value, "yyyy-MM-dd"), cAkunKas.Text, GetCostCenterUser(objData, GetRegistry(reg_Username)), "Pelunasan piutang an " & cNama.Text, nTunai.value, 0), False)
    lSave = IIf(lSave, UpdKodeTr(objData, msPelunasanPiutangSederhana, Faktur, Format(dTanggal.value, "yyyy-MM-dd"), aCfg(objData, msRekeningTopUp), GetCostCenterUser(objData, GetRegistry(reg_Username)), "Pelunasan piutang an " & cNama.Text, nVoucher.value, 0), False)
        lSave = IIf(lSave, UpdKodeTr(objData, msPelunasanPiutangSederhana, Faktur, Format(dTanggal.value, "yyyy-MM-dd"), GetAkunMember(objData, cCustomer.Text), GetCostCenterUser(objData, GetRegistry(reg_Username)), "Pelunasan piutang an " & cNama.Text, 0, nTotal.value), False)


    If lSave Then
      objData.Save GetDSN
      cSQL = ""
      cSQL = " select * from kartupiutang k where k.flag = 1 and flagid = '" & cFaktur.Text & "' and debet > 0 and k.kodeanggota = '" & cCustomer.Text & "'"
      Set dbData = objData.SQL(GetDSN, cSQL)
      If Not dbData.EOF Then
        GetCetakPelunasan2 objData, cFaktur.Text, cCustomer.Text
      Else
        GetCetakPelunasan objData, cFaktur.Text
      End If
      
    Else
      MsgBox "Maaf data tidak berhasil disimpan", vbExclamation
      objData.Cancel GetDSN
    End If
    
    initvalue
    GetEdit False
    
  Else
    MsgBox "Maaf tidak ada data untuk di proses", vbExclamation
  End If
End Sub

Private Sub cNama_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "anggota a", "a.kodeanggota,a.nama,a.alamat,d.keterangan as namadep", "a.kodeanggota", sisContent, cNama.Text, " or a.nama like '%" & cNama.Text & "%' ", , Array("left join dep d on d.kodedep = a.kodedep"))
  If Not dbData.EOF Then
    cNama.Text = cCustomer.Browse(dbData)
    cCustomer.Text = GetNull(dbData!kodeanggota)
    cNama.Text = GetNull(dbData!nama)
    cDepartement.Text = GetNull(dbData!namadep)
    If nPos = Add Then
      BiSAButton1_Click
    End If
  End If
End Sub

Private Sub cNama_Validate(Cancel As Boolean)
  cNama.Enabled = False
End Sub

Private Sub dTanggal_Validate(Cancel As Boolean)
  If Not IsInPeriod(dTanggal.value) Or (dTanggal.value > Date) Then
    Cancel = True
    dTanggal.SetFocus
    GetEdit False
  End If
End Sub

Private Sub Form_Activate()
'  If nPos = Add Then
'    GetData
'  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  TDBGrid1.FetchRowStyle = True
  SetIcon Me.hWnd, "SIKD"
  initvalue
  GetEdit False
  CenterForm Me
  
  Set dbData = objData.Browse(GetDSN, "costcenter", , "kodecostcenter", sisAssign, GetCostCenterUser(objData, GetRegistry(reg_Username)))
  If Not dbData.EOF Then
    lbCostCenter.Caption = "Cost Centre : " & GetNull(dbData!keterangan)
  End If
  
  TabIndex dTanggal, n
  TabIndex cCustomer, n
  TabIndex cNama, n
  TabIndex cFaktur, n
  TabIndex nTunai, n
'  TabIndex cKeteranganBayar, n
  TabIndex cAkunKas, n
  TabIndex cKodeKeteranganBayar, n
  
  TabIndex cmdAdd, n
  TabIndex cmdEdit, n
  TabIndex cmdHapus, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  
  If GetRegistry(reg_OptModelPelunasanPiutang) = 1 Then
    '1 = faktur
    '2 = bebas
    nTunai.Enabled = False
  Else
    nTunai.Enabled = True
  End If
  
  Set Prospective = TDBGrid1.Styles.Add("Prospective")
'    Prospective.Font.Italic = True
'    Prospective.Font.Bold = True
'    Prospective.ForeColor = vbBlue
    Prospective.BackColor = vbBlue
    Prospective.ForeColor = vbWhite
  Set Distributors = TDBGrid1.Styles.Add("Distributors")
    Distributors.BackColor = vbRed
    Distributors.ForeColor = vbWhite

End Sub

Private Function GetNamaAkun(ByVal obj As CodeSuiteLibrary.Data, ByVal kodeakun As String) As String
Dim db As New ADODB.Recordset
  GetNamaAkun = ""
  Set db = obj.Browse(GetDSN, "akun", , "kodeakun", sisAssign, kodeakun)
  If Not db.EOF Then
     GetNamaAkun = "Rekening Kas : " & kodeakun & " " & GetNull(db!keterangan)
  End If
End Function


Private Sub GridVoucher_AfterColUpdate(ByVal ColIndex As Integer)
  If ColIndex = 0 Then
    GridVoucher.Update
    GetSumTunai
  End If
  GridVoucher.Update
  Me.Refresh
End Sub

Private Sub nTunai_Change()
  nTotal.value = nVoucher.value + nTunai.value
End Sub

Private Sub nTunai_Validate(Cancel As Boolean)
  If nTotal.value > nSaldoPiutang.value Then
    MsgBox "Error. Tunai lebih banyak"
    Cancel = True
    nTunai.value = 0
    nTunai.SetFocus
  End If
  'kalau dicentang, cek jumlah totalnya. harus sama persis
  If nDebet > 0 Then
    If nDebet <> nTotal.value Then
      MsgBox "Error. Pelunasan tidak sama"
      Cancel = True
      nTunai.value = 0
      GetSumTunai
      nTunai.SetFocus
    End If
  End If
End Sub

Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
  If ColIndex = 0 Then
    If TDBGrid1.Columns(8).value <> "" Or TDBGrid1.Columns(2).value = "" Then
      MsgBox "Tidak bisa di edit"
      TDBGrid1.Columns(0).value = 0
      TDBGrid1.Update
    Else
      TDBGrid1.Update
      GetSumTunai
    End If
  End If
  TDBGrid1.Update
  Me.Refresh
End Sub

Private Sub tdbgrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
'  If ColIndex = 0 Then
'    If TDBGrid1.Columns(0).Value = -1 Then
'      If ColIndex <> 7 Then
'        TDBGrid1.Columns(7).Value = TDBGrid1.Columns(4).Value - TDBGrid1.Columns(6).Value
'      Else
'        TDBGrid1.Columns(6).Value = 0
'      End If
'    End If
'  Else
'    Cancel = True
'  End If
End Sub

Private Sub GetSumTunai()
Dim n As Integer
Dim nTmp As Double

  nDebet = 0
  nTmp = 0
  For n = 0 To vaArray.UpperBound(1)
    If vaArray(n, 0) = -1 Then
      nDebet = nDebet + vaArray(n, 5) - vaArray(n, 6)
'      If CDbl(vaArray(n, 4)) = CDbl(vaArray(n, 7)) Then
'        nPoinReg = getPoinReguler(vaArray(n, 2)) + nPoinReg
'      End If
    End If
  Next n
  
  'hitung penggunaan voucher
  For n = 0 To vaVoucher.UpperBound(1)
    If vaVoucher(n, 0) = -1 Then
      nTmp = nTmp + vaVoucher(n, 4)
    End If
  Next n
  nVoucher.value = nTmp
  nTotal.value = nVoucher.value + nTunai.value
  nTunai.value = nDebet - nVoucher.value
  
End Sub

Private Sub TDBGrid1_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal col As Integer, ByVal CellStyle As TrueOleDBGrid70.StyleDisp)
  Select Case col
      Case 7
          If TDBGrid1.Columns(7).CellText(Bookmark) = "1" Then CellStyle.ForeColor = vbRed
  End Select
End Sub

Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
    If CDbl(TDBGrid1.Columns(5).CellText(Bookmark)) > 0 Then
      RowStyle = Distributors
    End If
    
    If CDbl(TDBGrid1.Columns(6).CellText(Bookmark)) > 0 Then
      RowStyle = Prospective
    End If
End Sub

Sub GetCetakPelunasan(ByVal obj As CodeSuiteLibrary.Data, ByVal Faktur As String)
Dim n As Integer
Dim cTerbilang As String
Dim cField As String
Dim vaJoin
Dim vaGrid As New XArrayDB
Dim cHead As String
Dim cSQL As String

'  cSQL = ""
'  cSQL = cSQL & " select p.nomorpelunasanpiutang,p.nomorpenjualan,p.piutang,p.pelunasan,t.jthtmp from pelunasanpiutang p"
'  cSQL = cSQL & " LEFT JOIN totpenjualan t on t.nomorpenjualan = p.nomorpenjualan"
'  cSQL = cSQL & " where p.nomorpelunasanpiutang = '" & Faktur & "'"
'
'  Set dbData = obj.SQL(GetDSN, cSQL)
'  If Not dbData.EOF Then
'    n = 0
'    vaGrid.ReDim 0, dbData.RecordCount - 1, 0, 5
'    Do While Not dbData.EOF
'       vaGrid(n, 0) = n + 1
'       vaGrid(n, 1) = (dbData!nomorpenjualan)
'       vaGrid(n, 2) = (dbData!Pelunasan)
'       vaGrid(n, 3) = (dbData!jthtmp)
'       vaGrid(n, 4) = 0
'       vaGrid(n, 5) = vaGrid(n, 2) - vaGrid(n, 4)
'       dbData.MoveNext
'      n = n + 1
'    Loop
'
'
'  End If


    'AMBIL INFORMASI customer
    cSQL = ""
    cSQL = " select a.kodeanggota,a.nama,a.alamat,a.telp,t.nomorpelunasanpiutang,t.tgl,t.total from totpelunasanpiutangsederhana t"
    cSQL = cSQL & " LEFT JOIN anggota a on a.kodeanggota = t.kodeanggota where t.nomorpelunasanpiutang = '" & Faktur & "'"
    
    Set dbData = obj.SQL(GetDSN, cSQL)
    cTerbilang = "# " & Dec2Text(GetNull(dbData!Total)) & "Rupiah #"
    cHead = "Kuitansi Lunas"
    With rptKuitansiLunas
      .Parameters("dTgl").ValueExpression = "'" & Format(GetNull(dbData!tgl), "dd-MM-yyyy") & "'"
      .Parameters("cSE").ValueExpression = "'" & Faktur & "'"
      
      .Parameters("cNama").ValueExpression = "'" & GetNull(dbData!nama, "") & "'"
      .Parameters("cAlamat").ValueExpression = "'" & GetNull(dbData!alamat, "") & "'"
      .Parameters("cKota").ValueExpression = "'" & GetNull(dbData!telp) & "'"
      .Parameters("cKodeAnggota").ValueExpression = "'" & GetNull(dbData!kodeanggota, "") & "'"
      
      .Parameters("cTerbilang").ValueExpression = "'" & cTerbilang & "'"
      .Parameters("cTTD").ValueExpression = "'" & Padc(GetRegistry(reg_FullName), 45) & "'"
      .Parameters("cReceived").ValueExpression = "'" & Padc("", 45) & "'"
      
      .Parameters("nSubtotal").ValueExpression = GetNull(dbData!Total)
      .Parameters("nTotal").ValueExpression = GetNull(dbData!Total)
      .Parameters("cNamaPerusahaan").ValueExpression = "'" & aCfg(obj, msNamaPerusahaan) & "'"
      .Parameters("cAlamatPerusahaan").ValueExpression = "'Alamat : " & aCfg(obj, msAlamatPerusahaan) & " Telp/Fax " & aCfg(objData, msTelepon) & "/" & aCfg(objData, msFax) & "'"
      .Parameters("cUserName").ValueExpression = "'" & GetRegistry(reg_FullName) & "'"
      .Parameters("cJudul").ValueExpression = "'" & cHead & "'"
      .Parameters("keAkun").ValueExpression = "'" & GetNamaAkun(objData, cAkunKas.Text) & "'"
      
     ' Set .Array = vaGrid
      .Refresh
      If MsgBox("Apakah cetakan mau dalam bentuk kertas A4?!!" & vbCrLf & "Jika tidak maka cetakan akan dalam bentuk 1/2 kertas kuarto", vbYesNo) = vbYes Then
        .Profiles(0).PrinterPaperSize = tdbPPS_A4
      End If
      .PrintPreview
    End With
End Sub

Sub GetCetakPelunasan2(ByVal obj As CodeSuiteLibrary.Data, ByVal Faktur As String, ByVal cKodeAnggota As String)
Dim n As Integer
Dim cTerbilang As String
Dim cField As String
Dim vaJoin
Dim vaGrid As New XArrayDB
Dim cHead As String
Dim cSQL As String


    cSQL = ""
    cSQL = " select * from kartupiutang k where k.flag = 1 and flagid = '" & Faktur & "' and nomorkartupiutang<> '" & Faktur & "' and k.kodeanggota = '" & cKodeAnggota & "'"
    Set dbData = objData.SQL(GetDSN, cSQL)
    Set dbData = obj.SQL(GetDSN, cSQL)
    If Not dbData.EOF Then
      n = 0
      vaGrid.ReDim 0, dbData.RecordCount - 1, 0, 4
      Do While Not dbData.EOF
         vaGrid(n, 0) = n + 1
         vaGrid(n, 1) = (dbData!nomorkartupiutang)
         vaGrid(n, 2) = (dbData!debet)
         vaGrid(n, 3) = (dbData!kredit)
         vaGrid(n, 4) = (dbData!tgl)
         dbData.MoveNext
        n = n + 1
      Loop
    End If


    'AMBIL INFORMASI customer
    cSQL = ""
    cSQL = " select a.kodeanggota,a.nama,a.alamat,a.telp,t.nomorpelunasanpiutang,t.tgl,t.total from totpelunasanpiutangsederhana t"
    cSQL = cSQL & " LEFT JOIN anggota a on a.kodeanggota = t.kodeanggota where t.nomorpelunasanpiutang = '" & Faktur & "'"
    
    Set dbData = obj.SQL(GetDSN, cSQL)
    cTerbilang = "# " & Dec2Text(GetNull(dbData!Total)) & "Rupiah #"
    cHead = "Kuitansi Lunas"
    With rptKuitansiLunas2
      .Parameters("dTgl").ValueExpression = "'" & Format(GetNull(dbData!tgl), "dd-MM-yyyy") & "'"
      .Parameters("cSE").ValueExpression = "'" & Faktur & "'"
      
      .Parameters("cNama").ValueExpression = "'" & GetNull(dbData!nama, "") & "'"
      .Parameters("cAlamat").ValueExpression = "'" & GetNull(dbData!alamat, "") & "'"
      .Parameters("cKota").ValueExpression = "'" & GetNull(dbData!telp) & "'"
      .Parameters("cKodeAnggota").ValueExpression = "'" & GetNull(dbData!kodeanggota, "") & "'"
      
      .Parameters("cTerbilang").ValueExpression = "'" & cTerbilang & "'"
      .Parameters("cTTD").ValueExpression = "'" & Padc(GetRegistry(reg_FullName), 45) & "'"
      .Parameters("cReceived").ValueExpression = "'" & Padc("", 45) & "'"
      
      .Parameters("nSubtotal").ValueExpression = GetNull(dbData!Total)
      .Parameters("nTotal").ValueExpression = GetNull(dbData!Total)
      .Parameters("cNamaPerusahaan").ValueExpression = "'" & aCfg(obj, msNamaPerusahaan) & "'"
      .Parameters("cAlamatPerusahaan").ValueExpression = "'Alamat : " & aCfg(obj, msAlamatPerusahaan) & " Telp/Fax " & aCfg(objData, msTelepon) & "/" & aCfg(objData, msFax) & "'"
      .Parameters("cUserName").ValueExpression = "'" & GetRegistry(reg_FullName) & "'"
      .Parameters("cJudul").ValueExpression = "'" & cHead & "'"
      .Parameters("keAkun").ValueExpression = "'" & GetNamaAkun(objData, cAkunKas.Text) & "'"
      
      Set .Array = vaGrid
      .Refresh
      If MsgBox("Apakah cetakan mau dalam bentuk kertas A4?!!" & vbCrLf & "Jika tidak maka cetakan akan dalam bentuk 1/2 kertas kuarto", vbYesNo) = vbYes Then
        .Profiles(0).PrinterPaperSize = tdbPPS_A4
      End If
      .PrintPreview
    End With
End Sub

