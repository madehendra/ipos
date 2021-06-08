VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trPengeluaranBiaya 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PENGELUARAN/ BIAYA..."
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   15525
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   15525
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   6855
      TabIndex        =   18
      Top             =   30
      Width           =   8490
      Begin BiSANumberBoxProject.BiSANumberBox nTotal 
         Height          =   1395
         Left            =   105
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   390
         Width           =   8235
         _ExtentX        =   14526
         _ExtentY        =   2461
         Appearance      =   0
         Enabled         =   0   'False
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12648447
         Caption         =   " "
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   810
      Left            =   120
      Top             =   6165
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   1429
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
         TabIndex        =   0
         Top             =   270
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
         Picture         =   "trPengeluaranBiaya.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   12405
         TabIndex        =   1
         Top             =   210
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
         Picture         =   "trPengeluaranBiaya.frx":028A
      End
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   1170
         TabIndex        =   2
         Top             =   270
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
         Picture         =   "trPengeluaranBiaya.frx":0429
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   435
         Left            =   90
         TabIndex        =   3
         Top             =   270
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
         Picture         =   "trPengeluaranBiaya.frx":0555
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   13950
         TabIndex        =   4
         Top             =   210
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
         Picture         =   "trPengeluaranBiaya.frx":0700
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   12870
         TabIndex        =   5
         Top             =   210
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
         Picture         =   "trPengeluaranBiaya.frx":07A6
      End
   End
   Begin BiSAFramProject.BiSAFrame BisaFrame2 
      Height          =   4080
      Left            =   45
      Top             =   2055
      Width           =   15405
      _ExtentX        =   27173
      _ExtentY        =   7197
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
      Begin BiSANumberBoxProject.BiSANumberBox nJumlah 
         Height          =   330
         Left            =   11010
         TabIndex        =   6
         Top             =   75
         Width           =   3810
         _ExtentX        =   6720
         _ExtentY        =   582
         Appearance      =   0
         Decimals        =   0
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
      Begin BiSATextBoxProject.BiSATextBox cKeterangan 
         Height          =   330
         Left            =   7170
         TabIndex        =   7
         Top             =   75
         Width           =   3810
         _ExtentX        =   6720
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
         Height          =   330
         Left            =   3510
         TabIndex        =   8
         Top             =   75
         Width           =   3645
         _ExtentX        =   6429
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
      Begin BiSATextBoxProject.BiSABrowse cKodeAkun 
         Height          =   330
         Left            =   630
         TabIndex        =   9
         Top             =   75
         Width           =   2850
         _ExtentX        =   5027
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
         Left            =   75
         TabIndex        =   10
         Top             =   75
         Width           =   570
         _ExtentX        =   1005
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
         Height          =   3450
         Left            =   90
         TabIndex        =   11
         Top             =   420
         Width           =   15180
         _ExtentX        =   26776
         _ExtentY        =   6085
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
         Columns(1).Caption=   "AKUN"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "NAMA"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "KETT"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "JUMLAH"
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "###,###,###,##0.00"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   5
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=5"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=926"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=847"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=514"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=5054"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4974"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=512"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=6482"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=6403"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=512"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=6773"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=6694"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=197120"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=1482"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=1402"
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
         Appearance      =   0
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1,5
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
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=66,.parent=13,.alignment=0"
         _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=14"
         _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=15,.alignment=1"
         _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=17"
         _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=1"
         _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
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
      Begin BiSAButtonProject.BiSAButton cmdOK 
         Height          =   330
         Left            =   14850
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
         Picture         =   "trPengeluaranBiaya.frx":0A2C
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1920
      Left            =   120
      Top             =   105
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   3387
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
         Left            =   150
         TabIndex        =   13
         Top             =   285
         Width           =   2880
         _ExtentX        =   5080
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
         Left            =   150
         TabIndex        =   14
         Top             =   645
         Width           =   3750
         _ExtentX        =   6615
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
      Begin BiSATextBoxProject.BiSABrowse cAkunKas 
         Height          =   330
         Left            =   150
         TabIndex        =   15
         Top             =   1005
         Width           =   3090
         _ExtentX        =   5450
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
         Caption         =   "Akun Kas"
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
      Begin BiSATextBoxProject.BiSABrowse cNamaAkunKas 
         Height          =   330
         Left            =   3300
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1005
         Width           =   3150
         _ExtentX        =   5556
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
      Begin BiSATextBoxProject.BiSABrowse cCostCenter 
         Height          =   330
         Left            =   150
         TabIndex        =   17
         Top             =   1365
         Width           =   3090
         _ExtentX        =   5450
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
      Begin BiSATextBoxProject.BiSABrowse cNamaCostCenter 
         Height          =   330
         Left            =   3300
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1365
         Width           =   3150
         _ExtentX        =   5556
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
Attribute VB_Name = "trPengeluaranBiaya"
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

Private Sub GetFakturBrowse(ByVal lStat As Boolean)
  cFaktur.Button = lStat
End Sub

Private Sub cAkunKas_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", "kodeakun,keterangan", "kodeakun", sisContent, cAkunKas.Text, " and jenis = 'D'")
  If Not dbData.EOF Then
    cAkunKas.Text = cAkunKas.Browse(dbData)
    cNamaAkunKas.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub cAkunKas_Validate(Cancel As Boolean)
  cAkunKas.Enabled = False
End Sub

Private Sub cCostCenter_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "costcenter", "kodecostcenter,keterangan")
  If Not dbData.EOF Then
    cCostCenter.Text = cCostCenter.Browse(dbData, Array("Kode", "Keterangan"), , Array(15, 25))
  End If
End Sub

Private Sub cCostCenter_Validate(Cancel As Boolean)
  cCostCenter.Enabled = False
End Sub

Private Sub cFaktur_ButtonClick()
Dim db As New ADODB.Recordset
Dim n As Integer
Dim lSave As Boolean
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
  vaArray.ReDim 0, -1, 0, 4
  
  Set dbData = objData.Browse(GetDSN, "totbiaya t", "t.nomorbiaya,t.kodeakun,a.keterangan,t.kodecostcenter,t.jumlah", "t.tgl", sisAssign, Format(dTgl.value, "yyyy-MM-dd"), , "t.nomorbiaya desc", Array("LEFT JOIN akun a ON a.kodeakun = t.kodeakun"))
  If Not dbData.EOF Then
    cFaktur.Text = cFaktur.Browse(dbData, Array("Nomor", "Akun Kas", "Ket.", "Cost Center", "Jumlah"), , Array(15, 15, 15, 15))
    cAkunKas.Text = GetNull(dbData!kodeakun)
    cNamaAkunKas.Text = GetNull(dbData!keterangan)
    cCostCenter.Text = GetNull(dbData!kodecostcenter)
    nTotal.value = GetNull(dbData!jumlah)
    
    Set dbData = objData.Browse(GetDSN, "biaya b", "b.nomorbiaya,b.kodeakun,a.keterangan as namaakun,b.keterangan,b.jumlah", "b.nomorbiaya", sisAssign, cFaktur.Text, , , Array("LEFT JOIN akun a ON a.kodeakun = b.kodeakun"))
    If Not dbData.EOF Then
      Do While Not dbData.EOF
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        n = vaArray.UpperBound(1)
        vaArray(n, 0) = n + 1
        vaArray(n, 1) = GetNull(dbData!kodeakun)
        vaArray(n, 2) = GetNull(dbData!namaakun)
        vaArray(n, 3) = GetNull(dbData!keterangan)
        vaArray(n, 4) = GetNull(dbData!jumlah)
        dbData.MoveNext
      Loop
      nNomor.value = vaArray.UpperBound(1) + 2
    End If
    Set TDBGrid1.Array = vaArray
    TDBGrid1.ReBind
    TDBGrid1.Refresh
    Me.Refresh
    
    If nPos = Delete Then
      If MsgBox("Data akan dihapus?", vbYesNo) = vbYes Then
        objData.Start GetDSN
        lSave = IIf(lSave, DelKodeTr(objData, msBiaya, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "biaya", "nomorbiaya", sisAssign, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "totbiaya", "nomorbiaya", sisAssign, cFaktur.Text), False)
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

Private Sub cKodeAkun_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akunbiaya b", "b.kodeakun,a.keterangan", "b.kodeakun", sisContent, cKodeAkun.Text, "or a.keterangan like '%" & cKodeAkun.Text & "%' AND jenis = 'D'", , Array("LEFT JOIN akun a ON a.kodeakun = b.kodeakun"))
  If Not dbData.EOF Then
    cKodeAkun.Text = cKodeAkun.Browse(dbData, Array("Kode Akun", "Keterangan"), , Array(15, 25))
    cNamaAkun.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub cmdAktivasi_Click()
  frmAktivasi.Action Me
End Sub

Private Sub cmdAdd_Click()
  nPos = Add
  GetEdit True
  GetFakturBrowse False
  cFaktur.Text = CreateNomorFaktur(objData, sisModulTransaksi.biaya, "totbiaya", "nomorbiaya")
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
Dim db As New ADODB.Recordset
Dim nKe As Integer
  
  validOK = True
  
'  If isInGrid(vaArray, 1, cKodeAkun.Text, , nKe) And nNomor.value > vaArray.UpperBound(1) + 1 Then
'    MsgBox "Data sudah pernah dimasukkan sebelumnya dan akan dijumlahkan dengan data sebelumnya", vbExclamation
'    validOK = False
'
'    'jika barang yg sama diinput 2x dalam waktu bersamaan, maka akan qty akan
'    'dijumlahkan dengan yg sebelumnya, baik harga dan diskon akan sesuai dengan data
'    'yg diinput terakhir kali
'
'    vaArray(nKe, 3) = cKeterangan.Text
'    vaArray(nKe, 4) = vaArray(nKe, 4) + nJumlah.value
'
'    TDBGrid1.Update
'    TDBGrid1.Refresh
'
'    SumTotal
'    InitValue1
'    Exit Function
'  End If
  
  If Not GetValidDataBrowse(objData, "akun", "kodeakun", cKodeAkun.Text) Then
    MsgBox "Data akun tidak benar"
    validOK = False
    Exit Function
  End If
  If nJumlah.value = 0 Then
    MsgBox "Value 0", vbExclamation
    nJumlah.SetFocus
    validOK = False
    Exit Function
  End If
End Function

Private Function isOnBudget(ByVal obj As CodeSuiteLibrary.Data, ByVal cKodeAkunBudget As String, ByVal nBulanBudget As Integer, ByVal nBudgetVariable As Double, Optional ByRef nBudgetIs As Double) As Double
Dim db As New ADODB.Recordset
Dim cSQL As String

isOnBudget = True
nBudgetIs = 0
cSQL = cSQL & " select b.kodeakun,sum(b.jumlah) as jumlah,a.budget from biaya b"
cSQL = cSQL & " LEFT JOIN totbiaya t on t.nomorbiaya = b.nomorbiaya"
cSQL = cSQL & " LEFT JOIN akun a on a.kodeakun = b.kodeakun"
cSQL = cSQL & " where b.kodeakun= '" & cKodeAkunBudget & "'"
cSQL = cSQL & " and t.tgl > DATE_SUB(LAST_DAY(NOW()),INTERVAL  1 MONTH)"
cSQL = cSQL & " GROUP BY b.kodeakun"

  Set db = obj.SQL(GetDSN, cSQL)
  If Not db.EOF Then
    nBudgetIs = GetNull(db!budget)
    If GetNull(db!jumlah) + nBudgetVariable > GetNull(db!budget) Then
      isOnBudget = False
    End If
  End If

End Function


Private Sub cmdOK_Click()
Dim n As Integer
Dim nJumlah1 As Double

  If validOK() Then
   
    'jika baris <= Nomor
    If vaArray.UpperBound(1) + 2 <= nNomor.value Then
      vaArray.ReDim 0, nNomor.value - 1, 0, 9
      n = vaArray.UpperBound(1)
    'jika baris=0
    ElseIf vaArray.UpperBound(1) = -1 Then
      nNomor.value = 1
      vaArray.ReDim 0, nNomor.value - 1, 0, 9
      n = vaArray.UpperBound(1)
    Else
      n = nNomor.value - 1
    End If
        
    vaArray(n, 0) = nNomor.value
    vaArray(n, 1) = cKodeAkun.Text
    vaArray(n, 2) = cNamaAkun.Text
    vaArray(n, 3) = cKeterangan.Text
    vaArray(n, 4) = nJumlah.value
      
    TDBGrid1.Array = vaArray
    TDBGrid1.ReBind
    
    InitValue1
    SumTotal
    nNomor.value = vaArray.UpperBound(1) + 2
    cNamaAkun.SetFocus
  End If
End Sub

Private Sub SumTotal()
Dim n As Double
Dim nTmp As Double
  
  nTmp = 0
  For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
    nTmp = nTmp + vaArray(n, 4)
  Next n
  nTotal.value = nTmp
End Sub


Private Sub cmdSimpan_Click()
Dim lSave As Boolean
Dim Faktur As String
Dim n As Integer
Dim i As Integer
lSave = True
  
  If isValidSaving Then
    objData.Start GetDSN
    Faktur = cFaktur.Text
    lSave = IIf(lSave, objData.Update(GetDSN, "totbiaya", "nomorbiaya = '" & Faktur & "'", Array("nomorbiaya", "username", "kodeakun", "kodecostcenter", "jumlah", "tgl", "datetime"), Array(Faktur, GetRegistry(reg_Username), cAkunKas.Text, cCostCenter.Text, nTotal.value, dTgl.value, SNow)), False)
    lSave = IIf(lSave, objData.Delete(GetDSN, "biaya", "nomorbiaya", sisAssign, Faktur), False)
    For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
      lSave = IIf(lSave, objData.Add(GetDSN, "biaya", Array("nomorbiaya", "kodeakun", "keterangan", "jumlah"), Array(Faktur, vaArray(n, 1), vaArray(n, 3), vaArray(n, 4))), False)
    Next n
    
    DelKodeTr objData, msBiaya, Faktur
    'biaya
    ' kas
    UpdKodeTr objData, msBiaya, Faktur, Format(dTgl.value, "yyyy-MM-dd"), cAkunKas.Text, cCostCenter.Text, "Pengeluaran biaya " & Faktur, 0, nTotal.value, "K", SNow
    
    For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
      UpdKodeTr objData, msBiaya, Faktur, Format(dTgl.value, "yyyy-MM-dd"), vaArray(n, 1), cCostCenter.Text, vaArray(n, 3), vaArray(n, 4), 0, "K", SNow
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
Dim na As Integer
isValidSaving = True

  
  If Trim(cFaktur.Text) = "" Then
    MsgBox "Nomor transaksi tidak boleh kosong", vbCritical
    isValidSaving = False
    Exit Function
  End If
  
  If Trim(cAkunKas.Text) = "" Then
    MsgBox "Akun kas tidak boleh kosong", vbCritical
    isValidSaving = False
    Exit Function
  End If
  
  If Trim(cCostCenter.Text) = "" Then
    MsgBox "Cost Center tidak boleh kosong", vbCritical
    isValidSaving = False
    Exit Function
  End If
  
  If Not GetValidDataBrowse(objData, "costcenter", "kodecostcenter", cCostCenter.Text) Then
    MsgBox "Data cost center tidak benar", vbCritical
    isValidSaving = False
    Exit Function
  End If
  
  'cek budget masing masing pos
'  For na = vaArray.LowerBound(1) To vaArray.UpperBound(1)
'    If isOnBudget(objData, vaArray(na, 1), 1, vaArray(na, 4)) = False Then
'      isValidSaving = False
'      MsgBox "Rekening " & vaArray(na, 1) & " is over Budget"
'      Exit Function
'    End If
'  Next
  
End Function

Private Sub cNamaAkun_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akunbiaya b", "a.keterangan,b.kodeakun", "a.keterangan", sisContent, cNamaAkun.Text, " AND jenis = 'D' or b.kodeakun like '%" & cNamaAkun.Text & "%'", , Array("LEFT JOIN akun a ON a.kodeakun = b.kodeakun"))
  If Not dbData.EOF Then
    cNamaAkun.Text = cNamaAkun.Browse(dbData, Array("Keterangan", "Kode Akun"), , Array(25, 15))
    cKodeAkun.Text = GetNull(dbData!kodeakun)
  End If
End Sub

Private Sub dTgl_Validate(Cancel As Boolean)
  If Not IsInPeriod(dTgl.value) Or (dTgl.value > Date) Then
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
  TabIndex cAkunKas, n
  TabIndex cCostCenter, n
  
  TabIndex nNomor, n
  TabIndex cKodeAkun, n
  TabIndex cNamaAkun, n
  TabIndex cKeterangan, n
  TabIndex nJumlah, n
  TabIndex cmdOK, n
  
  TabIndex cmdAdd, n
  TabIndex cmdEdit, n
  TabIndex cmdHapus, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  TabIndex cmdAktivasi, n
End Sub

Private Sub initvalue()
Dim objMenu As New CodeSuiteLibrary.Menu

  cFaktur.Default
  dTgl.value = Date
  cAkunKas.Enabled = True
  cAkunKas.BackColor = vbWhite
  
'  Set dbData = objData.Browse(GetDSN, "costcenter", , "kodecostcenter", sisAssign, aCfg(objData, msCostCenterJualBeli))
'  If Not dbData.EOF Then
'    cCostCenter.Text = GetNull(dbData!kodecostcenter)
'  End If
  
  cCostCenter.Text = GetCostCenterUser(objData, GetRegistry(reg_Username))
  
  If objMenu.UserLevel <> 0 Then
    cAkunKas.Enabled = False
    cAkunKas.BackColor = vbButtonFace
  End If
  
  cAkunKas.Text = cKasTeller
  cNamaAkunKas.Text = cNamaKasTeller
  nTotal.Default
  InitValue1
  vaArray.ReDim 0, -1, 0, 4
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind


  cCostCenter.Enabled = True
  cCostCenter.BackColor = vbWhite
  If GetRegistry(reg_UserLevel) <> 0 Then
    cCostCenter.Enabled = False
    cCostCenter.BackColor = vbButtonFace
  End If
  
  cNamaCostCenter.Default
  cCostCenter.Text = GetCostCenterUser(objData, GetRegistry(reg_Username))
  Set dbData = objData.Browse(GetDSN, "costcenter", "kodecostcenter,keterangan", "kodecostcenter", sisAssign, cCostCenter.Text)
  If Not dbData.EOF Then
    cNamaCostCenter.Text = GetNull(dbData!keterangan)
  Else
    cNamaCostCenter.Default
  End If
  
  cAkunKas.Enabled = True
  cCostCenter.Enabled = True
  
End Sub

Private Sub InitValue1()
  nNomor.value = 1
  cKodeAkun.Default
  cNamaAkun.Default
  cKeterangan.Default
  nJumlah.Default
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


Private Sub nNomor_Validate(Cancel As Boolean)
Dim n As Single
  
  If GetValidNomorUrut(nNomor, vaArray) Then
    n = nNomor.value - 1
    If n <= vaArray.UpperBound(1) Then
      cKodeAkun.Text = vaArray(n, 1)
      cNamaAkun.Text = vaArray(n, 2)
      cKeterangan.Text = vaArray(n, 3)
      nJumlah.value = vaArray(n, 4)
    End If
  End If
End Sub


Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim n As Integer

  If KeyCode = vbKeyDelete Then
    If vaArray.UpperBound(1) >= 0 Then
      TDBGrid1.Delete
      SumTotal
      For n = 0 To vaArray.UpperBound(1)
        vaArray(n, 0) = n + 1
      Next
      nNomor.value = vaArray.UpperBound(1) + 2
      TDBGrid1.ReBind
    End If
  End If
End Sub

