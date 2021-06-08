VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trPembatalanKasir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CETAK ATAU PEMBATALAN KASIR"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   11640
   Begin BiSAFramProject.BiSAFrame frame 
      Height          =   1380
      Left            =   0
      Top             =   0
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   2434
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
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1860
         TabIndex        =   12
         Top             =   585
         Width           =   210
      End
      Begin BiSATextBoxProject.BiSABrowse cFaktur 
         Height          =   330
         Left            =   240
         TabIndex        =   0
         Top             =   975
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   582
         Text            =   "12345678901234567890"
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
         MaxLength       =   20
         Appearance      =   0
         Button          =   -1  'True
         Caption         =   "NO. TRANSAKSI"
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
      Begin BiSADateProject.BiSADate dTgl 
         Height          =   330
         Left            =   240
         TabIndex        =   1
         Top             =   225
         Width           =   2970
         _ExtentX        =   5239
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
         Caption         =   "TGL TRANSAKSI"
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
      Begin BiSATextBoxProject.BiSABrowse cKasir 
         Height          =   360
         Left            =   2085
         TabIndex        =   10
         Top             =   585
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   635
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
      Begin VB.Label Label1 
         Caption         =   "Kasir"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   525
         TabIndex        =   11
         Top             =   555
         Width           =   495
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   645
      Left            =   0
      Top             =   5700
      Width           =   11610
      _ExtentX        =   20479
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
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   10365
         TabIndex        =   2
         Top             =   135
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
         Picture         =   "trPembatalanKasir.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdHapus 
         Height          =   435
         Left            =   9285
         TabIndex        =   3
         Top             =   135
         Width           =   1065
         _ExtentX        =   1879
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
         Picture         =   "trPembatalanKasir.frx":00A6
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   8835
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   135
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
         Picture         =   "trPembatalanKasir.frx":032C
      End
   End
   Begin BiSAFramProject.BiSAFrame BisaFrame2 
      Height          =   4380
      Left            =   0
      Top             =   1350
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   7726
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
      Begin BiSANumberBoxProject.BiSANumberBox nTotal 
         Height          =   330
         Left            =   7875
         TabIndex        =   4
         Top             =   3645
         Width           =   3630
         _ExtentX        =   6403
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   3525
         Left            =   75
         TabIndex        =   5
         Top             =   75
         Width           =   11445
         _ExtentX        =   20188
         _ExtentY        =   6218
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
         Columns(1).Caption=   "BARCODE"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "NAMA BARANG"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "QTY"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "SATUAN"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "HARGA"
         Columns(5).DataField=   ""
         Columns(5).NumberFormat=   "Standard"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "JUMLAH"
         Columns(6).DataField=   ""
         Columns(6).NumberFormat=   "Standard"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=953"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=873"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=514"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=3228"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3149"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=512"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=5239"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=5159"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=512"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=1720"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1640"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=514"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=1746"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=1667"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=513"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=2990"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2910"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=514"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=3149"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=3069"
         Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=514"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=2,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bold=0,.fontsize=825"
         _StyleDefs(11)  =   ":id=2,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(27)  =   ":id=14,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(28)  =   ":id=14,.fontname=Tahoma"
         _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=1"
         _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
         _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=0"
         _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=66,.parent=13,.alignment=1"
         _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=2"
         _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
         _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.alignment=1"
         _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(66)  =   "Named:id=33:Normal"
         _StyleDefs(67)  =   ":id=33,.parent=0"
         _StyleDefs(68)  =   "Named:id=34:Heading"
         _StyleDefs(69)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(70)  =   ":id=34,.wraptext=-1"
         _StyleDefs(71)  =   "Named:id=35:Footing"
         _StyleDefs(72)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(73)  =   "Named:id=36:Selected"
         _StyleDefs(74)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(75)  =   "Named:id=37:Caption"
         _StyleDefs(76)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(77)  =   "Named:id=38:HighlightRow"
         _StyleDefs(78)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(79)  =   "Named:id=39:EvenRow"
         _StyleDefs(80)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(81)  =   "Named:id=40:OddRow"
         _StyleDefs(82)  =   ":id=40,.parent=33"
         _StyleDefs(83)  =   "Named:id=41:RecordSelector"
         _StyleDefs(84)  =   ":id=41,.parent=34"
         _StyleDefs(85)  =   "Named:id=42:FilterBar"
         _StyleDefs(86)  =   ":id=42,.parent=33"
      End
      Begin BiSANumberBoxProject.BiSANumberBox nDiscount 
         Height          =   330
         Left            =   3825
         TabIndex        =   6
         Top             =   3660
         Width           =   3630
         _ExtentX        =   6403
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
      Begin BiSANumberBoxProject.BiSANumberBox nSubTotal 
         Height          =   330
         Left            =   120
         TabIndex        =   7
         Top             =   3660
         Width           =   3630
         _ExtentX        =   6403
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
         Caption         =   "SUB TOTAL"
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
      Begin BiSANumberBoxProject.BiSANumberBox nCash 
         Height          =   330
         Left            =   7890
         TabIndex        =   8
         Top             =   4005
         Width           =   3630
         _ExtentX        =   6403
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
         Caption         =   "TUNAI"
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
   End
End
Attribute VB_Name = "trPembatalanKasir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaGrid As New XArrayDB
Dim nTunai As Double
Dim dDateTime As Date
 
Private Sub initvalue()
  dTgl.Value = Date
  cFaktur.Default
  nSubTotal.Default
  nDiscount.Default
  nTotal.Default
  nCash.Default
  ClearGrid
End Sub

Private Sub cFaktur_ButtonClick()
Dim cWhere As String

  cWhere = " AND username = '" & cKasir.Text & "'"
  If Check1.Value = 1 Then
    Set dbData = objData.Browse(GetDSN, "totkasir", "nomorkasir,tgl,username", "tgl", sisAssign, Format(dTgl.Value, "yyyy-mm-dd"), cWhere, "username,nomorkasir")
  Else
    Set dbData = objData.Browse(GetDSN, "totkasir", "nomorkasir,tgl,username", "tgl", sisAssign, Format(dTgl.Value, "yyyy-mm-dd"), , "username,nomorkasir")
  End If
  cFaktur.Text = cFaktur.Browse(dbData, Array("No", "Tanggal", "UserName"), , Array(20, 20, 20))
  If Not dbData.EOF Then
    GetDataInduk
    GetDataDetail
  Else
    initvalue
    Exit Sub
  End If
End Sub

Private Sub GetDataInduk()
  Set dbData = objData.Browse(GetDSN, "totkasir", , "nomorkasir", sisAssign, cFaktur.Text)
  If Not dbData.EOF Then
    nSubTotal.Value = GetNull(dbData!Subtotal)
    nDiscount.Value = GetNull(dbData!Discount)
    nTotal.Value = GetNull(dbData!Total)
    nTunai = GetNull(dbData!Tunai)
    nCash.Value = nTunai
  End If
End Sub

Private Sub ClearGrid()
  vaGrid.Clear
  vaGrid.ReDim 0, -1, 0, 7
  Set TDBGrid1.Array = vaGrid
  TDBGrid1.ReBind
  TDBGrid1.Refresh
End Sub

Private Sub GetDataDetail()
Dim n As Integer
  
  ClearGrid
  Set dbData = objData.Browse(GetDSN, "kasir k", "k.*,s.barcode,s.nama,s.kodesatuan", "k.nomorkasir", sisAssign, cFaktur.Text, , , _
                            Array("left Join stock s on s.kodestock = k.kodestock"))
  If Not dbData.EOF Then
    n = 0
    vaGrid.ReDim 0, dbData.RecordCount - 1, 0, 7
    Do While Not dbData.EOF
      vaGrid(n, 0) = n + 1
      vaGrid(n, 1) = GetNull(dbData!barcode, "")
      vaGrid(n, 2) = GetNull(dbData!nama)
      vaGrid(n, 3) = GetNull(dbData!qty)
      vaGrid(n, 4) = GetNull(dbData!kodesatuan, "")
      vaGrid(n, 5) = GetNull(dbData!Harga)
      vaGrid(n, 6) = GetNull(dbData!jumlah)
      vaGrid(n, 7) = GetNull(dbData!KodeStock)
      dbData.MoveNext
      n = n + 1
    Loop
    Set TDBGrid1.Array = vaGrid
    TDBGrid1.ReBind
    TDBGrid1.Refresh
  End If
End Sub

Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    keybd_event VK_TAB, 0, 0, 0
  End If
End Sub

Private Sub Check1_Validate(Cancel As Boolean)
  cFaktur.Default
End Sub

Private Sub cKasir_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "username", "username,fullname", "username", sisContent, cKasir.Text, , "username")
  If Not dbData.EOF Then
    cKasir.Text = cKasir.Browse(dbData)
  End If
End Sub

Private Sub cmdHapus_Click()
Dim n As Integer
Dim vaField, vaValue
Dim lSave As Boolean
lSave = True

  If ValidDelete Then
    If MsgBox("Anda yakin transaksi ini akan dibatalkan?", vbYesNo + vbInformation) = vbYes Then
      objData.Start GetDSN
      lSave = IIf(lSave, objData.Delete(GetDSN, "batalkasir", "nomorkasir", sisAssign, cFaktur.Text), False)
      vaField = Array("nomorkasir", "kodestock", "qty", "harga", "jumlah")
      For n = vaGrid.LowerBound(1) To vaGrid.UpperBound(1)
        vaValue = Array(cFaktur.Text, vaGrid(n, 7), vaGrid(n, 3), vaGrid(n, 5), vaGrid(n, 6))
        lSave = IIf(lSave, objData.Add(GetDSN, "batalkasir", vaField, vaValue), False)
      Next n
      
      vaField = Array("nomorkasir", "subtotal", "discount", "total", "tunai", "username", "tgl", "datetime")
      vaValue = Array(cFaktur.Text, nSubTotal.Value, nDiscount.Value, nTotal.Value, nCash.Value, GetRegistry(reg_UserName), Format(Date, "yyyy-MM-dd"), SNow)
      lSave = IIf(lSave, objData.Update(GetDSN, "totbatalkasir", "nomorkasir = '" & cFaktur.Text & "'", vaField, vaValue), False)
      
      lSave = IIf(lSave, objData.Delete(GetDSN, "kasir", "nomorkasir", sisAssign, cFaktur.Text), False)
      lSave = IIf(lSave, objData.Delete(GetDSN, "totkasir", "nomorkasir", sisAssign, cFaktur.Text), False)
      lSave = IIf(lSave, objData.Delete(GetDSN, "kartustock", "nomor", sisAssign, cFaktur.Text), False)
      
      lSave = IIf(lSave, DelKodeTr(objData, msPenjualanKasir, cFaktur.Text), False)
      
      If lSave Then
        objData.Save GetDSN
      Else
        objData.Cancel GetDSN
      End If
      initvalue
      dTgl.SetFocus
      Exit Sub
    Else
      initvalue
      Exit Sub
    End If
  End If
End Sub

Private Function ValidDelete() As Boolean
ValidDelete = True
  
  If cFaktur.Text = "" Then
    MsgBox "Nomor Tidak Ada, Data Tidak Bisa Dihapus..", vbExclamation, Me.Caption
    cFaktur.SetFocus
    ValidDelete = False
    Exit Function
  End If
End Function

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
  If MsgBox("Akan mencetak transaksi?", vbYesNo + vbInformation) = vbYes Then
    PrintStruk
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd, "SIKD"
  InitGrid TDBGrid1
  CenterForm Me
  Me.Top = 0
  initvalue
  
  TabIndex dTgl, n
  TabIndex Check1, n
  TabIndex cKasir, n
  TabIndex cFaktur, n
  TabIndex cmdHapus, n
  TabIndex cmdKeluar, n
End Sub

Private Sub PrintStruk()
Dim n As Double
Dim i As Double
Dim nBruto As Double
Dim nTotQty As Double
Dim vaArray As New XArrayDB
vaArray.ReDim 0, -1, 0, 7
      
  If vaGrid.UpperBound(1) >= 0 Then
    For i = vaGrid.LowerBound(1) To vaGrid.UpperBound(1)
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 1) = vaGrid(n, 1)
      vaArray(n, 2) = vaGrid(n, 2)
      vaArray(n, 3) = vaGrid(n, 3)
      vaArray(n, 4) = vaGrid(n, 4)
      vaArray(n, 5) = vaGrid(n, 5)
      vaArray(n, 6) = vaGrid(n, 6)
      vaArray(n, 7) = vaGrid(n, 7)
    Next i
      
    Open "lpt1" For Output As #1
    Print #1, Chr(27); Chr(33); Chr(4);
    Print #1, Chr(27) & Chr(97) & Chr(1)
    Print #1, "STRUK KASIR"
    Print #1, aCfg(objData, msNamaPerusahaan)
    Print #1, aCfg(objData, msAlamatPerusahaan)
    Select Case GetRegistry(reg_AlignmentThermal)
      Case 1 ' rata kiri
                Print #1, Chr(27) & Chr(97) & Chr(3)
      Case 2 ' rata kanan
          Print #1, Chr(27) & Chr(97) & Chr(2)
    End Select

    Print #1, ""
    Print #1, "No. " & cFaktur.Text
    Print #1, Format(Now, "dd-MM-yyyy HH:MM:SS")
    Print #1, ""
    
    Print #1, Replicate("-", 27)
    Print #1, Padl("Qty", 6); Padl("Hrg Net", 11); Padl("Jml", 10)
    Print #1, Replicate("-", 27)
    For n = 0 To vaArray.UpperBound(1)
      If vaArray(n, 3) <> 0 Then
        nBruto = nBruto + (vaArray(n, 3) * vaArray(n, 5))
        nTotQty = nTotQty + vaArray(n, 3)
        Print #1, vaArray(n, 2) ' vaArray(n, 1) kolom REF barang ditiadakan karena terlalu panjang
        Print #1, Padl(Format(vaArray(n, 3), "#,##0"), 3) & " x " & Padl(Format(vaArray(n, 5), "#,###,##0"), 8) & " = " & Padl(Format(vaArray(n, 3) * vaArray(n, 5), "#,###,##0"), 10)
      End If
    Next
    
    Print #1, Replicate("-", 27)
    Print #1, Format(nTotQty, "###,###,##0") & " Items"
    Print #1, Padl("Bruto  : ", 9); Padl(Format(nBruto, "###,###,##0"), 10)
    Print #1, Padl("Disc   : ", 9); Padl(Format(nDiscount.Value, "###,###,##0"), 10)
    Print #1, Padl("Total  : ", 9); Padl(Format(nTotal.Value, "###,###,##0"), 10)
    Print #1, Padl("Tunai  : ", 9); Padl(Format(nCash.Value, "###,###,##0"), 10)
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, Chr(27); Chr(33); Chr(0);
    Close #1
  End If
End Sub
