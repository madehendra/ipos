VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trUpdateKartuPiutangMember 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Kartu Piutang Member"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   11220
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   675
      Left            =   15
      Top             =   7020
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   1191
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
         TabIndex        =   0
         Top             =   120
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
         Picture         =   "trUpdateKartuPiutangMember.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   8310
         TabIndex        =   1
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
         Picture         =   "trUpdateKartuPiutangMember.frx":028A
      End
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   2325
         TabIndex        =   2
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
         Picture         =   "trUpdateKartuPiutangMember.frx":0429
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   435
         Left            =   105
         TabIndex        =   3
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
         Picture         =   "trUpdateKartuPiutangMember.frx":0555
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   9855
         TabIndex        =   4
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
         Picture         =   "trUpdateKartuPiutangMember.frx":0700
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   8775
         TabIndex        =   5
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
         Picture         =   "trUpdateKartuPiutangMember.frx":07A6
      End
   End
   Begin BiSAFramProject.BiSAFrame BisaFrame2 
      Height          =   5580
      Left            =   15
      Top             =   1470
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   9843
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
      Begin BiSANumberBoxProject.BiSANumberBox nPenyesuaian 
         Height          =   330
         Left            =   8265
         TabIndex        =   6
         Top             =   75
         Width           =   2160
         _ExtentX        =   3810
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
      Begin BiSATextBoxProject.BiSABrowse cNamaMember 
         Height          =   330
         Left            =   2595
         TabIndex        =   7
         Top             =   75
         Width           =   4050
         _ExtentX        =   7144
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
      Begin BiSATextBoxProject.BiSABrowse cKodeMember 
         Height          =   330
         Left            =   645
         TabIndex        =   8
         Top             =   75
         Width           =   1965
         _ExtentX        =   3466
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
      Begin BiSANumberBoxProject.BiSANumberBox nNomor 
         Height          =   330
         Left            =   105
         TabIndex        =   9
         Top             =   75
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   582
         Appearance      =   0
         Decimals        =   0
         DecimalPoint    =   ""
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
         Height          =   5025
         Left            =   90
         TabIndex        =   10
         Top             =   420
         Width           =   10860
         _ExtentX        =   19156
         _ExtentY        =   8864
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "NO."
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Kode"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Nama Member"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Saldo Piutang"
         Columns(3).DataField=   ""
         Columns(3).NumberFormat=   "###,###,###,##0.00"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Penyesuaian"
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "###,###,###,##0.00"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   5
         Splits(0)._UserFlags=   0
         Splits(0).PartialRightColumn=   0   'False
         Splits(0).MarqueeStyle=   3
         Splits(0).AllowRowSizing=   0   'False
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AllowColSelect=   0   'False
         Splits(0).AllowRowSelect=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=5"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=926"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=847"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=514"
         Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=3493"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=3413"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=512"
         Splits(0)._ColumnProps(13)=   "Column(1).WrapText=1"
         Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(15)=   "Column(2).Width=7091"
         Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=7011"
         Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(19)=   "Column(2)._ColStyle=512"
         Splits(0)._ColumnProps(20)=   "Column(2).WrapText=1"
         Splits(0)._ColumnProps(21)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(22)=   "Column(3).Width=2884"
         Splits(0)._ColumnProps(23)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(3)._WidthInPix=2805"
         Splits(0)._ColumnProps(25)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(26)=   "Column(3)._ColStyle=197122"
         Splits(0)._ColumnProps(27)=   "Column(3).WrapText=1"
         Splits(0)._ColumnProps(28)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(29)=   "Column(4).Width=3254"
         Splits(0)._ColumnProps(30)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(4)._WidthInPix=3175"
         Splits(0)._ColumnProps(32)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(33)=   "Column(4)._ColStyle=514"
         Splits(0)._ColumnProps(34)=   "Column(4).WrapText=1"
         Splits(0)._ColumnProps(35)=   "Column(4).Order=5"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         Appearance      =   0
         ColumnFooters   =   -1  'True
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1.5
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         InsertMode      =   0   'False
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
         _StyleDefs(16)  =   ":id=3,.fontname=Tahoma"
         _StyleDefs(17)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(18)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1,.wraptext=-1"
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
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=1"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=0"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=66,.parent=13,.alignment=1"
         _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=14"
         _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=15,.alignment=1"
         _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=17"
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
         _StyleDefs(75)  =   ":id=41,.parent=34"
         _StyleDefs(76)  =   "Named:id=42:FilterBar"
         _StyleDefs(77)  =   ":id=42,.parent=33"
      End
      Begin BiSANumberBoxProject.BiSANumberBox nSaldoPiutang 
         Height          =   330
         Left            =   6630
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   75
         Width           =   1620
         _ExtentX        =   2858
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
      Begin BiSAButtonProject.BiSAButton cmdOK 
         Height          =   330
         Left            =   10485
         TabIndex        =   12
         Top             =   75
         Width           =   420
         _ExtentX        =   741
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
         Alignment       =   1
         BackColor       =   -2147483633
         Picture         =   "trUpdateKartuPiutangMember.frx":0A2C
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1485
      Left            =   15
      Top             =   0
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   2619
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
      Begin BiSADateProject.BiSADate dTgl 
         Height          =   330
         Left            =   75
         TabIndex        =   13
         Top             =   405
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
      Begin BiSATextBoxProject.BiSABrowse cFaktur 
         Height          =   330
         Left            =   75
         TabIndex        =   14
         Top             =   750
         Width           =   3405
         _ExtentX        =   6006
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
   End
End
Attribute VB_Name = "trUpdateKartuPiutangMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cNoOrder As String

Dim nPos As SisPos
Dim lEdit As Boolean
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB
Dim cKode As String
Dim cJenis  As String
Dim nSaldoStock As Double
Dim cTelp As String

Private Sub GetFakturBrowse(ByVal lStat As Boolean)
  cFaktur.Button = lStat
End Sub

Private Sub cFaktur_ButtonClick()
Dim db As New ADODB.Recordset
Dim n As Integer
Dim lSave As Boolean
Dim nQtyTmp As Single

lSave = True
nQtyTmp = 0

  Set db = objData.Browse(GetDSN, "totupdatekartupiutang", "nomorupdatekartupiutang", "nomorupdatekartupiutang", sisContent, cFaktur.Text, " and tgl = '" & Format(dTgl.Value, "yyyy-MM-dd") & "'")
  If Not db.EOF Then
    cFaktur.Text = cFaktur.Browse(db)
    'ambil nilai detail
    Set db = objData.Browse(GetDSN, "updatekartupiutang p", "p.kodeanggota,a.nama,p.jumlah", "nomorupdatekartupiutang", sisAssign, cFaktur.Text, , , Array("LEFT JOIN anggota a ON a.kodeanggota = p.kodeanggota"))
    If Not db.EOF Then
      vaArray.ReDim 0, -1, 0, 9
      Do While Not db.EOF
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        n = vaArray.UpperBound(1)
        vaArray(n, 0) = n + 1
        vaArray(n, 1) = GetNull(db!kodeanggota)
        vaArray(n, 2) = GetNull(db!nama)
        vaArray(n, 3) = GetPiutangMemberDong(objData, vaArray(n, 1))
        vaArray(n, 4) = GetNull(db!jumlah)
        db.MoveNext
      Loop
      Set TDBGrid1.Array = vaArray
      TDBGrid1.ReBind
      TDBGrid1.Refresh
      Me.Refresh
      nNomor.Value = vaArray.UpperBound(1) + 2
    End If

    If nPos = Delete Then
      If MsgBox("Data akan dihapus?", vbYesNo) = vbYes Then
        objData.Start GetDSN

        'rutin menghapus transaksi

        lSave = IIf(lSave, objData.Delete(GetDSN, "totupdatekartupiutang", "nomorupdatekartupiutang", sisAssign, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "updatekartupiutang", "nomorupdatekartupiutang", sisAssign, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "kartupiutang", "nomorkartupiutang", sisAssign, cFaktur.Text), False)
        

        If lSave Then
          objData.Save GetDSN
        Else
          objData.Cancel GetDSN
        End If

      End If
      initvalue
      GetEdit False
    End If
    If nPos = Edit Then
      SendKeysA vbKeyReturn, True
    End If
  End If
End Sub

Private Sub cFaktur_Validate(Cancel As Boolean)
  If nPos = Edit Then
      cFaktur.Enabled = False
      cFaktur.BackColor = vbButtonFace
  End If
End Sub

Private Sub cKodeMember_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "anggota a", "a.kodeanggota,a.nama,a.kodeanggota", "a.kodeanggota", sisContent, cKodeMember.Text, , "a.nama")
  If Not dbData.EOF Then
    cKodeMember.Text = cKodeMember.Browse(dbData)
    cNamaMember.Text = GetNull(dbData!nama)
    nSaldoPiutang.Value = GetPiutangMemberDong(objData, GetNull(dbData!kodeanggota))
  End If
End Sub

Private Sub cmdAktivasi_Click()
  frmAktivasi.Action Me
End Sub

Private Sub cmdAdd_Click()
Dim i As Integer

  nPos = Add
  GetEdit True
  GetFakturBrowse False
  cFaktur.Text = GetNomor("totupdatekartupiutang", "nomorupdatekartupiutang", GetID, sisModulTransaksi.UpdateKartuPiutang)
End Sub

Private Sub cmdEdit_Click()
  nPos = Edit
  GetEdit True
  GetFakturBrowse True
End Sub

Private Sub cmdHapus_Click()
  nPos = Delete
  GetEdit True
  GetFakturBrowse True

End Sub

Private Sub cmdKeluar_Click()
  If lEdit Then
    GetEdit False
    initvalue
  Else
    Unload Me
  End If
End Sub

Private Function validOK() As Boolean
  validOK = True
End Function

Private Sub cmdOK_Click()
Dim n As Integer
Dim nJumlah1 As Double
Dim nQtyTmp As Single

  
  If validOK() Then
    'jika baris <= Nomor
    If vaArray.UpperBound(1) + 2 <= nNomor.Value Then
      vaArray.ReDim 0, nNomor.Value - 1, 0, 4
      n = vaArray.UpperBound(1)
    'jika baris=0
    ElseIf vaArray.UpperBound(1) = -1 Then
      nNomor.Value = 1
      vaArray.ReDim 0, nNomor.Value - 1, 0, 4
      n = vaArray.UpperBound(1)
    Else
      n = nNomor.Value - 1
    End If
        
    vaArray(n, 0) = nNomor.Value
    vaArray(n, 1) = cKodeMember.Text
    vaArray(n, 2) = cNamaMember.Text
    vaArray(n, 3) = nSaldoPiutang.Value
    vaArray(n, 4) = nPenyesuaian.Value
    
    
    TDBGrid1.Array = vaArray
    TDBGrid1.ReBind
    TDBGrid1.MoveNext
    
    InitValue1
    nNomor.Value = vaArray.UpperBound(1) + 2
    cKodeMember.SetFocus
  End If
End Sub

Private Function ValidSaving() As Boolean
  ValidSaving = True
End Function

Private Sub cmdSimpan_Click()
Dim lSave As Boolean
Dim Faktur As String
Dim n As Integer
Dim i As Integer
Dim nValueTunai As Double
Dim nValueKredit As Double

lSave = True

  If isValidSaving Then
    objData.Start GetDSN
    Faktur = cFaktur.Text
    'simpan di tabel totupdatekartupiutang
    lSave = IIf(lSave, objData.Update(GetDSN, "totupdatekartupiutang", "nomorupdatekartupiutang = '" & Faktur & "'", Array("nomorupdatekartupiutang", "username", "tgl", "datetime"), Array(Faktur, GetRegistry(reg_UserName), Format(Date, "yyyy-MM-dd"), SNow)), False)
    'simpan di tabel updatekartupiutang
    lSave = IIf(lSave, objData.Delete(GetDSN, "updatekartupiutang", "nomorupdatekartupiutang", sisAssign, Faktur), False)
    lSave = IIf(lSave, objData.Delete(GetDSN, "kartupiutang", "nomorkartupiutang", sisAssign, Faktur), False)
    
    For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)

      lSave = IIf(lSave, objData.Add(GetDSN, "updatekartupiutang", Array("nomorupdatekartupiutang", "kodeanggota", "jumlah", "tgl"), Array(Faktur, vaArray(n, 1), vaArray(n, 4), Format(dTgl.Value, "yyyy-MM-dd"))), False)
      
      
      'update kartupiutang
      
      If vaArray(n, 4) - vaArray(n, 3) > 0 Then
        lSave = IIf(lSave, UpdKartuHutang(objData, SisUpdatePiutangDebet, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), vaArray(n, 1), "Penyesuaian Piutang " & vaArray(n, 2), vaArray(n, 4) - vaArray(n, 3), SNow, GetRegistry(reg_UserName), False), False)
      End If
      If vaArray(n, 4) - vaArray(n, 3) < 0 Then
        lSave = IIf(lSave, UpdKartuHutang(objData, SisUpdatePiutangKredit, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), vaArray(n, 1), "Penyesuaian Piutang " & vaArray(n, 2), vaArray(n, 3) - vaArray(n, 4), SNow, GetRegistry(reg_UserName), False), False)
      End If
  
    Next n

    If lSave Then
      objData.Save GetDSN
    Else
      objData.Cancel GetDSN
    End If

    initvalue
    GetEdit False
  End If
End Sub

Private Function isValidSaving() As Boolean
Dim dba As New ADODB.Recordset
Dim nPernahBayar As Double

isValidSaving = True
End Function

Private Sub cNamaMember_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "anggota a", "a.nama,a.kodeanggota", "a.nama", sisContent, cNamaMember.Text)
  If Not dbData.EOF Then
    cNamaMember.Text = cNamaMember.Browse(dbData)
    cKodeMember.Text = GetNull(dbData!kodeanggota)
    nSaldoPiutang.Value = GetPiutangMemberDong(objData, GetNull(dbData!kodeanggota))
  End If
End Sub

Private Function GetPiutangMemberDong(ByVal obj As CodeSuiteLibrary.Data, ByVal cAnggota As String) As Double
Dim db As New ADODB.Recordset

  GetPiutangMemberDong = 0
  Set db = obj.Browse(GetDSN, "kartupiutang", "sum(debet-kredit) as saldopiutang", "kodeanggota", sisAssign, cAnggota)
  If Not db.EOF Then
    GetPiutangMemberDong = GetNull(db!saldopiutang)
  End If
End Function

Private Sub dTgl_Validate(Cancel As Boolean)
  If Not IsInPeriod(dTgl.Value) Or (dTgl.Value > Date) Then
    Cancel = True
    dTgl.SetFocus
    GetEdit False
  End If
End Sub

Private Sub Form_Load()
Dim n As Single
  
  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  initvalue
  GetEdit False

  TabIndex dTgl, n
  TabIndex cFaktur, n
   
  TabIndex nNomor, n
  TabIndex cKodeMember, n
  TabIndex cNamaMember, n
  TabIndex nSaldoPiutang, n
  TabIndex nPenyesuaian, n
  
  TabIndex cmdOK, n
  
  TabIndex cmdAdd, n
  TabIndex cmdEdit, n
  TabIndex cmdHapus, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  TabIndex cmdAktivasi, n
End Sub

Private Sub initvalue()
  cFaktur.Default
  dTgl.Value = Date
  vaArray.ReDim 0, -1, 0, 5
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
  InitValue1
End Sub

Private Sub InitValue1()
  nNomor.Value = 1
  cKodeMember.Default
  cNamaMember.Default
  nSaldoPiutang.Default
  nPenyesuaian.Default
  
   
End Sub

Private Sub GetEdit(lPar As Boolean)
  BiSAFrame1.Enabled = lPar
  BiSAFrame2.Enabled = lPar
  lEdit = lPar
  initvalue
  SetButton cmdSimpan, cmdKeluar, cmdAdd, cmdEdit, cmdHapus, nPos, lPar, cmdAktivasi
  
  If lPar Then
    dTgl.SetFocus
    If nPos = Add Then
      cFaktur.Enabled = False
      cFaktur.BackColor = vbButtonFace
    Else
      cFaktur.Enabled = True
      cFaktur.BackColor = vbWindowBackground
      cFaktur.CaptionBackColor = vbButtonFace
    End If
  End If
End Sub

Private Sub lbCostCenter_Click()

End Sub

Private Sub nNomor_Validate(Cancel As Boolean)
Dim n As Single
  
  If GetValidNomorUrut(nNomor, vaArray) Then
    n = nNomor.Value - 1
    If n <= vaArray.UpperBound(1) Then
      cKodeMember.Text = vaArray(n, 1)
      cNamaMember.Text = vaArray(n, 2)
      nSaldoPiutang.Value = vaArray(n, 3)
      nPenyesuaian.Value = vaArray(n, 4)
      
    End If
  End If
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim n As Integer

  If KeyCode = vbKeyDelete Then
    If vaArray.UpperBound(1) >= 0 Then
      TDBGrid1.Delete
      
      For n = 0 To vaArray.UpperBound(1)
        vaArray(n, 0) = n + 1
      Next
      nNomor.Value = vaArray.UpperBound(1) + 2
      TDBGrid1.ReBind
    End If
  End If
End Sub

