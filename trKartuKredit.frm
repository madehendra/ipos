VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form trKartuKredit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kartu Kredit / Debit"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9600
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   9600
   StartUpPosition =   3  'Windows Default
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   2235
      Left            =   0
      Top             =   0
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   3942
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
      Begin VB.OptionButton optAktif 
         Caption         =   "Tidak (&0)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   3180
         TabIndex        =   13
         Top             =   1815
         Width           =   1125
      End
      Begin VB.OptionButton optAktif 
         Caption         =   "Aktif (&1)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   2070
         TabIndex        =   12
         Top             =   1815
         Width           =   1125
      End
      Begin BiSATextBoxProject.BiSATextBox cKode 
         Height          =   330
         Left            =   420
         TabIndex        =   0
         Top             =   255
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   582
         Text            =   "123456789012345"
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
         MaxLength       =   15
         Appearance      =   0
         GetPicture      =   1
         Caption         =   "Kode"
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
      Begin BiSATextBoxProject.BiSATextBox cKeterangan 
         Height          =   330
         Left            =   420
         TabIndex        =   1
         Top             =   630
         Width           =   3930
         _ExtentX        =   6932
         _ExtentY        =   582
         Text            =   "12345678901234567890"
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
         MaxLength       =   40
         Appearance      =   0
         GetPicture      =   1
         Caption         =   "Keterangan"
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
      Begin BiSATextBoxProject.BiSABrowse cNamaKodeAkun 
         Height          =   330
         Left            =   4065
         TabIndex        =   2
         Top             =   1395
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   582
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
      Begin BiSATextBoxProject.BiSATextBox cKodeAkun 
         Height          =   330
         Left            =   435
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1395
         Width           =   3570
         _ExtentX        =   6297
         _ExtentY        =   582
         Text            =   "123456789012345"
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
         MaxLength       =   15
         Appearance      =   0
         GetPicture      =   1
         Caption         =   "Kode Akun"
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
      Begin BiSANumberBoxProject.BiSANumberBox nPersen 
         Height          =   330
         Left            =   915
         TabIndex        =   11
         Top             =   1020
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   582
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
         Caption         =   "Fee (%)"
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
      Height          =   2880
      Left            =   0
      Top             =   2250
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   5080
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
      Begin TrueOleDBGrid70.TDBGrid DataGrid1 
         Height          =   2625
         Left            =   75
         TabIndex        =   4
         Top             =   60
         Width           =   9435
         _ExtentX        =   16642
         _ExtentY        =   4630
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Kode"
         Columns(0).DataField=   "Kode"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Keterangan"
         Columns(1).DataField=   "Keterangan"
         Columns(1).NumberFormat=   "General Date"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Fee(%)"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Kode Akun"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Status"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   5
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=5"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2117"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2037"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=6932"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=6853"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=1958"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1879"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=3043"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2963"
         Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(25)=   "Column(4).Width=1693"
         Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1614"
         Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=516"
         Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Microsoft Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Microsoft Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         TabAction       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   -2147483633
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000018&,.bold=0,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bgcolor=&HFF8080&"
         _StyleDefs(10)  =   ":id=4,.fgcolor=&H80000014&,.appearance=0,.ellipsis=0,.borderColor=&HFF8000&"
         _StyleDefs(11)  =   ":id=4,.bold=-1,.fontsize=1425,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=4,.fontname=Times New Roman"
         _StyleDefs(13)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&HFF8080&"
         _StyleDefs(14)  =   ":id=2,.fgcolor=&H80000014&,.bold=0,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(15)  =   ":id=2,.strikethrough=0,.charset=0"
         _StyleDefs(16)  =   ":id=2,.fontname=Microsoft Sans Serif"
         _StyleDefs(17)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(18)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(19)  =   ":id=3,.fontname=Microsoft Sans Serif"
         _StyleDefs(20)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(21)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(22)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(23)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(24)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&H8000000F&"
         _StyleDefs(25)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(26)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(27)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(28)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(29)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(30)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.alignment=2"
         _StyleDefs(31)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(32)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(33)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(34)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(35)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(36)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(37)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(38)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(39)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(40)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(60)  =   "Named:id=33:Normal"
         _StyleDefs(61)  =   ":id=33,.parent=0"
         _StyleDefs(62)  =   "Named:id=34:Heading"
         _StyleDefs(63)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(64)  =   ":id=34,.wraptext=-1,.bold=-1,.fontsize=975,.italic=0,.underline=0"
         _StyleDefs(65)  =   ":id=34,.strikethrough=0,.charset=0"
         _StyleDefs(66)  =   ":id=34,.fontname=Times New Roman"
         _StyleDefs(67)  =   "Named:id=35:Footing"
         _StyleDefs(68)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(69)  =   "Named:id=36:Selected"
         _StyleDefs(70)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(71)  =   "Named:id=37:Caption"
         _StyleDefs(72)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(73)  =   "Named:id=38:HighlightRow"
         _StyleDefs(74)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(75)  =   "Named:id=39:EvenRow"
         _StyleDefs(76)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(77)  =   "Named:id=40:OddRow"
         _StyleDefs(78)  =   ":id=40,.parent=33"
         _StyleDefs(79)  =   "Named:id=41:RecordSelector"
         _StyleDefs(80)  =   ":id=41,.parent=34"
         _StyleDefs(81)  =   "Named:id=42:FilterBar"
         _StyleDefs(82)  =   ":id=42,.parent=33"
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   0
      Top             =   5130
      Width           =   9555
      _ExtentX        =   16854
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
         Left            =   2220
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
         Picture         =   "trKartuKredit.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   6855
         TabIndex        =   6
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
         Picture         =   "trKartuKredit.frx":028A
      End
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   1170
         TabIndex        =   7
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
         Picture         =   "trKartuKredit.frx":0429
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
         Picture         =   "trKartuKredit.frx":0555
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   8385
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
         Picture         =   "trKartuKredit.frx":0700
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   7305
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
         Picture         =   "trKartuKredit.frx":07A6
      End
   End
End
Attribute VB_Name = "trKartuKredit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lClick As Boolean
Dim nPos As SisPos
Dim lEdit As Boolean
Dim vaArray As New XArrayDB
Dim objData As New CodeSuiteLibrary.Data
Dim dbData As New ADODB.Recordset


Private Sub GetEdit(lPar As Boolean)
  lEdit = lPar
  BiSAFrame1.Enabled = lPar
  SetButton cmdSimpan, cmdKeluar, cmdAdd, cmdEdit, cmdHapus, nPos, lPar
End Sub

Private Sub cKode_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    keybd_event VK_TAB, 0, 0, 0
  End If
End Sub

Private Sub cmdAktivasi_Click()
  frmAktivasi.Action Me
End Sub

Private Sub cmdAdd_Click()
  nPos = Add
  GetEdit True
  initvalue
  cKeterangan.SetFocus
End Sub

Private Sub initvalue()
  cKode.Text = ""
  cKeterangan.Default
  cKodeAkun.Default
  nPersen.Default
  cKodeAkun.Default
  cNamaKodeAkun.Default
End Sub

Private Sub HapusData()
Dim cInfo As String

  cInfo = "Kode: " & cKode.Text & vbCrLf
  cInfo = cInfo & "Keterangan: " & cKeterangan.Text & vbCrLf

  If Trim(cKode.Text) <> "" Then
    If MsgBox("Data benar-benar dihapus...?" & vbCrLf & vbCrLf & cInfo, vbExclamation + vbYesNo) = vbYes Then
      If Not lExist(objData, "trkartu", "kodekartu", cKode.Text) Then
        objData.Delete GetDSN, "kartu", "kodekartu", sisAssign, cKode.Text
      Else
        MsgBox "Maaf, data ini masih digunakan dalam sistem" & vbCrLf & "Data tidak bisa dihapus"
      End If
    End If
    initvalue
    getSQL
  End If
  GetEdit False
End Sub

Private Sub cmdEdit_Click()
  nPos = Edit
  GetEdit True
  cKeterangan.SetFocus
End Sub

Private Sub cmdHapus_Click()
  nPos = Delete
  GetEdit True
  cKeterangan.SetFocus
  HapusData
End Sub

Private Sub cmdKeluar_Click()
  If lEdit Then
    GetEdit False
    initvalue
  Else
    Unload Me
  End If
End Sub

Private Sub cmdSimpan_Click()
Dim vaField, vaValue

  If ValidSaving() Then
    vaField = Array("namakartu", "kodeakun", "fee", "status")
    vaValue = Array(cKeterangan.Text, cKodeAkun.Text, nPersen.Value, GetOpt(optAktif))
    objData.Update GetDSN, "kartu", "kodekartu = '" & cKode.Text & "'", vaField, vaValue
    getSQL
    initvalue
    GetEdit False
   Else
    cKode.SetFocus
    Exit Sub
  End If
End Sub

Private Function ValidSaving() As Boolean

  ValidSaving = True
  If Trim(cKodeAkun.Text) = "" Then
    ValidSaving = False
    cNamaKodeAkun.SetFocus
    Exit Function
  End If
  If Trim(cKeterangan.Text) = "" Then
    ValidSaving = False
    cKeterangan.SetFocus
    Exit Function
  End If
End Function

Private Sub GetMemory()
  cKode.Text = DataGrid1.Columns(0).Text
  cKeterangan.Text = Trim(DataGrid1.Columns(1).Text)
  nPersen.Value = DataGrid1.Columns(2).Text
  cKodeAkun.Text = DataGrid1.Columns(3).Text
  SetOpt optAktif, DataGrid1.Columns(4).Text
End Sub

Private Sub cNamaKodeAkun_ButtonClick()
Dim db As New ADODB.Recordset

  Set db = objData.Browse(GetDSN, "akun", "kodeakun,keterangan", "(kodeakun ", sisContent, cNamaKodeAkun.Text, " OR keterangan like '" & cNamaKodeAkun.Text & "') and jenis = 'D' and kodeakun like  '1.%'")
  If Not db.EOF Then
    cNamaKodeAkun.Text = cNamaKodeAkun.Browse(db)
    cNamaKodeAkun.Text = GetNull(db!keterangan)
    cKodeAkun.Text = GetNull(db!kodeakun)
  End If
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  GetMemory
End Sub

Private Sub Form_Load()
Dim n As Single
  
  SetIcon Me.hWnd, "SIKD"
  initvalue
  getSQL
  GetEdit False
  CenterForm Me
  
  TabIndex cKode, n
  TabIndex cKeterangan, n
  TabIndex nPersen, n
  TabIndex cNamaKodeAkun, n
  
  TabIndex cmdAdd, n
  TabIndex cmdEdit, n
  TabIndex cmdHapus, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
End Sub

Private Sub getSQL()
Dim n As Single

  vaArray.ReDim 0, -1, 0, 4
  Set dbData = objData.Browse(GetDSN, "kartu", , , , , , "namakartu")
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!kodekartu, "")
      vaArray(n, 1) = GetNull(dbData!namakartu, "")
      vaArray(n, 2) = GetNull(dbData!fee, "")
      vaArray(n, 3) = GetNull(dbData!kodeakun, "")
      vaArray(n, 4) = GetNull(dbData!Status, "")
      dbData.MoveNext
    Loop
  End If
  Set DataGrid1.Array = vaArray
  DataGrid1.ReBind
  DataGrid1.Refresh
End Sub




