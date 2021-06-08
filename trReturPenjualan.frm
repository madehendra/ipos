VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trReturPenjualan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RETUR PENJUALAN"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   11805
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   585
      Left            =   0
      Top             =   6525
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   1032
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
         Picture         =   "trReturPenjualan.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   3390
         TabIndex        =   1
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
         Picture         =   "trReturPenjualan.frx":028A
      End
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   1170
         TabIndex        =   2
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
         Picture         =   "trReturPenjualan.frx":0429
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   435
         Left            =   105
         TabIndex        =   3
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
         Picture         =   "trReturPenjualan.frx":0555
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   10635
         TabIndex        =   4
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
         Picture         =   "trReturPenjualan.frx":0700
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   9555
         TabIndex        =   5
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
         Picture         =   "trReturPenjualan.frx":07A6
      End
   End
   Begin BiSAFramProject.BiSAFrame BisaFrame2 
      Height          =   3975
      Left            =   0
      Top             =   2580
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   7011
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
         Left            =   8595
         TabIndex        =   6
         Top             =   3585
         Width           =   1560
         _ExtentX        =   2752
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
      Begin BiSANumberBoxProject.BiSANumberBox nQty 
         Height          =   330
         Left            =   6030
         TabIndex        =   7
         Top             =   75
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   582
         Appearance      =   0
         MinValue        =   0
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
      Begin BiSANumberBoxProject.BiSANumberBox nHarga 
         Height          =   330
         Left            =   7710
         TabIndex        =   8
         Top             =   75
         Width           =   1485
         _ExtentX        =   2619
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
      Begin BiSATextBoxProject.BiSATextBox cSatuan 
         Height          =   330
         Left            =   6855
         TabIndex        =   9
         Top             =   75
         Width           =   870
         _ExtentX        =   1535
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
      Begin BiSATextBoxProject.BiSABrowse cNama 
         Height          =   330
         Left            =   2580
         TabIndex        =   10
         Top             =   75
         Width           =   3465
         _ExtentX        =   6112
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
      Begin BiSATextBoxProject.BiSABrowse cBarcode 
         Height          =   330
         Left            =   630
         TabIndex        =   11
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
         Left            =   90
         TabIndex        =   12
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
         Height          =   2955
         Left            =   90
         TabIndex        =   13
         Top             =   420
         Width           =   11640
         _ExtentX        =   20532
         _ExtentY        =   5212
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
         Columns(1).Caption=   "KODE BRG"
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
         Columns(5).NumberFormat=   "###,###,###,###,##0.00"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "DISC"
         Columns(6).DataField=   ""
         Columns(6).NumberFormat=   "###,###,###,###,##0.00"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "JUMLAH"
         Columns(7).DataField=   ""
         Columns(7).NumberFormat=   "###,###,###,###,##0.00"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   8
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=926"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=847"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=514"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=3493"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3413"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=512"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=6085"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=6006"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=512"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=1455"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1376"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=197122"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=1482"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=1402"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=513"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=2593"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2514"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=514"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=1296"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=1217"
         Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=514"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(36)=   "Column(7).Width=3149"
         Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=3069"
         Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=514"
         Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
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
         HeadLines       =   1.5
         FootLines       =   0
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
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=66,.parent=13,.alignment=1"
         _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=14"
         _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=15,.alignment=1"
         _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=17"
         _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=2"
         _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
         _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
         _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
         _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=58,.parent=13,.alignment=1"
         _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
         _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
         _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
         _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=62,.parent=13,.alignment=1"
         _StyleDefs(66)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
         _StyleDefs(67)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
         _StyleDefs(68)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
         _StyleDefs(69)  =   "Named:id=33:Normal"
         _StyleDefs(70)  =   ":id=33,.parent=0"
         _StyleDefs(71)  =   "Named:id=34:Heading"
         _StyleDefs(72)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(73)  =   ":id=34,.wraptext=-1"
         _StyleDefs(74)  =   "Named:id=35:Footing"
         _StyleDefs(75)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(76)  =   "Named:id=36:Selected"
         _StyleDefs(77)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(78)  =   "Named:id=37:Caption"
         _StyleDefs(79)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(80)  =   "Named:id=38:HighlightRow"
         _StyleDefs(81)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(82)  =   "Named:id=39:EvenRow"
         _StyleDefs(83)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(84)  =   "Named:id=40:OddRow"
         _StyleDefs(85)  =   ":id=40,.parent=33"
         _StyleDefs(86)  =   "Named:id=41:RecordSelector"
         _StyleDefs(87)  =   ":id=41,.parent=34"
         _StyleDefs(88)  =   "Named:id=42:FilterBar"
         _StyleDefs(89)  =   ":id=42,.parent=33"
      End
      Begin BiSANumberBoxProject.BiSANumberBox nJumlah 
         Height          =   330
         Left            =   9915
         TabIndex        =   14
         Top             =   75
         Width           =   1395
         _ExtentX        =   2461
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
         Left            =   11325
         TabIndex        =   15
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
         Picture         =   "trReturPenjualan.frx":0A2C
      End
      Begin BiSANumberBoxProject.BiSANumberBox nDisc1 
         Height          =   330
         Left            =   9180
         TabIndex        =   16
         Top             =   75
         Width           =   750
         _ExtentX        =   1323
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
      Begin BiSAFramProject.BiSAFrame BiSAFrame4 
         Height          =   555
         Left            =   75
         Top             =   3390
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   979
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
            TabIndex        =   17
            Top             =   195
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
         Height          =   555
         Left            =   1635
         Top             =   3390
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   979
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
            TabIndex        =   18
            Top             =   195
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
         Height          =   555
         Left            =   3090
         Top             =   3390
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   979
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
            TabIndex        =   19
            Top             =   195
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
         Height          =   555
         Left            =   4440
         Top             =   3390
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   979
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
            TabIndex        =   20
            Top             =   195
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
      Begin BiSANumberBoxProject.BiSANumberBox nPiutang 
         Height          =   330
         Left            =   10170
         TabIndex        =   21
         Top             =   3585
         Width           =   1560
         _ExtentX        =   2752
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
         Left            =   8565
         TabIndex        =   23
         Top             =   3390
         Width           =   570
      End
      Begin VB.Label Label5 
         Caption         =   "PIUTANG"
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
         Left            =   10200
         TabIndex        =   22
         Top             =   3390
         Width           =   885
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   2625
      Left            =   0
      Top             =   0
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   4630
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
      Begin BiSAButtonProject.BiSAButton cmdFind 
         Height          =   330
         Left            =   10560
         TabIndex        =   38
         Top             =   135
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   582
         Caption         =   "GET"
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
      Begin BiSATextBoxProject.BiSABrowse cNamaCustomer 
         Height          =   330
         Left            =   3330
         TabIndex        =   24
         Top             =   750
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
         TabIndex        =   25
         Top             =   1065
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
         TabIndex        =   26
         Top             =   1065
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
      Begin BiSATextBoxProject.BiSABrowse cCustomer 
         Height          =   330
         Left            =   75
         TabIndex        =   27
         Top             =   750
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
         Caption         =   "Member"
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
         TabIndex        =   28
         Top             =   435
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
      Begin BiSADateProject.BiSADate dJthTmp 
         Height          =   330
         Left            =   6915
         TabIndex        =   29
         Top             =   1365
         Width           =   2580
         _ExtentX        =   4551
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
         Caption         =   "Due Date"
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
         TabIndex        =   30
         Top             =   1995
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
         TabIndex        =   31
         Top             =   1680
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
         TabIndex        =   32
         Top             =   1050
         Width           =   3660
         _ExtentX        =   6456
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
         TabIndex        =   33
         Top             =   1380
         Width           =   3750
         _ExtentX        =   6615
         _ExtentY        =   582
         Text            =   "12345678"
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
      Begin BiSATextBoxProject.BiSABrowse cNomorPembelian 
         Height          =   330
         Left            =   6420
         TabIndex        =   37
         Top             =   135
         Width           =   4155
         _ExtentX        =   7329
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
         Caption         =   "Penjualan No."
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
      Begin BiSAButtonProject.BiSAButton cmdClear 
         Height          =   330
         Left            =   11070
         TabIndex        =   39
         Top             =   135
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   582
         Caption         =   "CLR"
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
      Begin BiSANumberBoxProject.BiSANumberBox nSisaPiutang 
         Height          =   330
         Left            =   6930
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   495
         Width           =   3675
         _ExtentX        =   6482
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
         Caption         =   "Saldo"
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
         TabIndex        =   36
         Top             =   120
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
         Height          =   255
         Left            =   9060
         TabIndex        =   35
         Top             =   2085
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
         Height          =   255
         Left            =   9045
         TabIndex        =   34
         Top             =   1755
         Width           =   240
      End
   End
End
Attribute VB_Name = "trReturPenjualan"
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

Private Sub GetFakturBrowse(ByVal lStat As Boolean)
  cFaktur.Button = lStat
End Sub

Private Sub cBarcode_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "stock s", "s.kodestock,s.barcode,s.nama,s.kodesatuan,s.hargajual", "s.barcode", sisContent, cBarcode.Text, " and s.statusnonaktif <> 1")
  If Not dbData.EOF Then
    cBarcode.Text = cBarcode.Browse(dbData)
    GetDataStock
  Else
    cBarcode.Default
  End If
End Sub

Private Sub SumBayar()
  nPiutang.Value = nTotal.Value - IIf(nTunai.Value > nTotal.Value, nTotal.Value, nTunai.Value)
End Sub

Private Sub cFaktur_ButtonClick()
Dim db As New ADODB.Recordset
Dim n As Integer
Dim lSave As Boolean

lSave = True
  
  Set db = objData.Browse(GetDSN, "totrtnpenjualan", "nomorreturpenjualan,tgl,subtotal,total,piutang,nomorpenjualan", "nomorreturpenjualan", sisContent, cFaktur.Text, " and tgl = '" & Format(dTgl.Value, "yyyy-MM-dd") & "' and kodeanggota = '" & cCustomer.Text & "'")
  If Not db.EOF Then
    cFaktur.Text = cFaktur.Browse(db)
    'ambil nilai total
    Set db = objData.Browse(GetDSN, "totrtnpenjualan", , "nomorreturpenjualan", sisAssign, cFaktur.Text)
    If Not db.EOF Then
      cFakturAsli.Text = GetNull(db!fakturasli, "")
      dJthTmp.Value = GetNull(db!jthtmp)
      nPersDisc.Value = GetNull(db!PersDisc, 0)
      nPPn.Value = GetNull(db!ppn, 0)
      nSubTotal.Value = GetNull(db!Subtotal, 0)
      nDiscount.Value = GetNull(db!Discount, 0)
      nPajak.Value = GetNull(db!PAJAK, 0)
      nTotal.Value = GetNull(db!Total, 0)
      nTunai.Value = GetNull(db!Tunai, 0)
      nPiutang.Value = GetNull(db!Piutang, "")
      cNomorPembelian.Text = GetNull(db!nomorpenjualan, "")
    End If
    
    'ambil nilai detail
    Set db = objData.Browse(GetDSN, "returpenjualan p", "s.barcode,p.kodestock,s.Nama,p.qty,p.kodesatuan,p.harga,p.discount,p.jumlah", "nomorreturpenjualan", sisAssign, cFaktur.Text, , , Array("Left join stock s on s.kodestock = p.kodestock"))
    If Not db.EOF Then
      vaArray.ReDim 0, -1, 0, 8
      Do While Not db.EOF
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        n = vaArray.UpperBound(1)
        vaArray(n, 0) = n + 1
        vaArray(n, 1) = GetNull(db!barcode)
        vaArray(n, 2) = GetNull(db!nama)
        vaArray(n, 3) = GetNull(db!qty)
        vaArray(n, 4) = GetNull(db!kodesatuan)
        vaArray(n, 5) = GetNull(db!Harga)
        vaArray(n, 6) = GetNull(db!Discount)
        vaArray(n, 7) = GetNull(db!jumlah)
        vaArray(n, 8) = GetNull(db!KodeStock)
        db.MoveNext
      Loop
      Set TDBGrid1.Array = vaArray
      TDBGrid1.ReBind
      TDBGrid1.Refresh
      Me.Refresh
    End If
    
    If nPos = Delete Then
      If MsgBox("Data akan dihapus?", vbYesNo) = vbYes Then
        objData.Start GetDSN
        lSave = IIf(lSave, DelKodeTr(objData, msReturPenjualan, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "totrtnpenjualan", "nomorreturpenjualan", sisAssign, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "returpenjualan", "nomorreturpenjualan", sisAssign, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "kartupiutang", "nomorkartupiutang", sisAssign, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "kartustock", "nomor", sisAssign, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "bukubesar", "faktur", sisAssign, cFaktur.Text), False)
        
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

Private Sub cmdAktivasi_Click()
  frmAktivasi.Action Me
End Sub

Private Sub GetDataStock()
  cBarcode.Text = GetNull(dbData!barcode, "")
  cKode = GetNull(dbData!KodeStock, "")
  cNama.Text = GetNull(dbData!nama, "")
  cSatuan.Text = GetNull(dbData!kodesatuan, "")
  nHarga.Value = GetNull(dbData!hargajual, 0)
End Sub

Private Sub cmdAdd_Click()
  nPos = Add
  GetEdit True
  GetFakturBrowse False
  cFaktur.Text = CreateNomorFaktur(objData, sisModulTransaksi.returpenjualan, "totrtnpenjualan", "nomorreturpenjualan")
End Sub

Private Sub cmdClear_Click()
  cNomorPembelian.Default
  nSisaPiutang.Default
End Sub

Private Sub cmdEdit_Click()
  nPos = Edit
  GetEdit True
  GetFakturBrowse True
End Sub

Private Sub cmdFind_Click()
  Load trFindFakturPembelian
  trFindFakturPembelian.cCustomer.Text = cCustomer.Text
  trFindFakturPembelian.chk(1).Value = 1
  trFindFakturPembelian.Show vbModal
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
  If Not GetValidDataBrowse(objData, "stock", "kodestock", cKode) Then
'  If Trim(cKode) = "" Then
    MsgBox "Barang tersebut tidak ada dalam database"
    cBarcode.SetFocus
    validOK = False
    Exit Function
  End If
  
'  If isInGrid(vaArray, 8, cKode) And nNomor.Value > vaArray.UpperBound(1) + 1 Then
'    MsgBox "Data sudah pernah dimasukkan sebelumnya ..", vbExclamation
'    cBarcode.SetFocus
'    validOK = False
'    Exit Function
'  End If

  If Not GetValidDataBrowse(objData, "stock", "kodestock", cKode) Then
    MsgBox "Data barang tidak ada dalam database"
    validOK = False
    Exit Function
  End If
  
End Function

Private Sub cmdOK_Click()
Dim n As Integer
Dim nJumlah1 As Double

  If validOK() Then
   
    'jika baris <= Nomor
    If vaArray.UpperBound(1) + 2 <= nNomor.Value Then
      vaArray.ReDim 0, nNomor.Value - 1, 0, 8
      n = vaArray.UpperBound(1)
    'jika baris=0
    ElseIf vaArray.UpperBound(1) = -1 Then
      nNomor.Value = 1
      vaArray.ReDim 0, nNomor.Value - 1, 0, 8
      n = vaArray.UpperBound(1)
    Else
      n = nNomor.Value - 1
    End If
        
    vaArray(n, 0) = nNomor.Value
    vaArray(n, 1) = cBarcode.Text
    vaArray(n, 2) = cNama.Text
    vaArray(n, 3) = nQty.Value
    vaArray(n, 4) = cSatuan.Text
    vaArray(n, 5) = nHarga.Value
    vaArray(n, 6) = nDisc1.Value
    vaArray(n, 7) = nJumlah.Value
    vaArray(n, 8) = cKode
      
    TDBGrid1.Array = vaArray
    TDBGrid1.ReBind
    
    nJumlah1 = 0
    For n = 0 To vaArray.UpperBound(1)
      nJumlah1 = nJumlah1 + vaArray(n, 7)
    Next
    nSubTotal.Value = nJumlah1
    
    SumTotal
    
    InitValue1
    
    nNomor.Value = vaArray.UpperBound(1) + 2
    cBarcode.SetFocus
  End If
End Sub

Private Sub SumTotal()
Dim n As Double
  
  nSubTotal.Value = 0
  For n = 0 To vaArray.UpperBound(1)
    nSubTotal.Value = nSubTotal.Value + vaArray(n, 7)
  Next
  
  If nPersDisc.Enabled = True Then
    nDiscount.Value = nPersDisc.Value / 100 * (nSubTotal.Value)
  End If
  
  nPajak.Value = (nPPn.Value / 100) * (nSubTotal.Value - (nDiscount.Value + nDiscount.Value))
  nTotal.Value = nSubTotal.Value + nPajak.Value - nDiscount.Value
  nPiutang.Value = nTotal.Value
End Sub


Private Function ValidSaving() As Boolean
Dim nSisaLunas As Double

  ValidSaving = True
  
  If vaArray.UpperBound(1) < 0 Then
    MsgBox "Nota kosong, data tidak disimpan"
    ValidSaving = False
    Exit Function
  End If
  
  'cek validitas
  If Not GetValidDataBrowse(objData, "anggota", "kodeanggota", cCustomer.Text) Then
    MsgBox "Kode member tidak tertera dengan benar" & vbCrLf & "Data tidak bisa disimpan"
    cCustomer.SetFocus
    ValidSaving = False
    Exit Function
  End If
  
  'Cek apakah nilai retur ini melebihi dari nilai yg tertera pada faktur penjualan
  If Trim(cNomorPembelian.Text) <> "" Then
    If isLunas(objData, cNomorPembelian.Text, nSisaLunas) = True Then
      MsgBox "Maaf, Nomor penjualan " & cNomorPembelian.Text & " sudah lunas. Tidak bisa dilakukan retur"
      cNomorPembelian.SetFocus
      ValidSaving = False
      Exit Function
    Else
      If nPiutang.Value > nSisaLunas Then
        MsgBox "Tidak bisa dilanjutkan. Sisa piutang tidak mencukupi untuk dilakukan retur" & vbCrLf & "Sisa piutang untuk nomor penjualan " & cNomorPembelian.Text & " adalah " & Format(nSisaLunas, "###,###,###,##0.00")
        ValidSaving = False
        Exit Function
      End If
    End If
  End If
  
'  If Trim(cNomorPembelian.Value) <> "" Then
'    If nPiutang.Value > nSisaPiutang.Value Then
'      'stop
'      MsgBox "Maaf, nilai retur lebih dari saldo piutang yg telah dipilih"
'      ValidSaving = False
'      Exit Function
'    End If
'  End If
  
End Function

Private Sub cCustomer_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "anggota", "kodeanggota,nama,kodedep,alamat", "kodeanggota", sisContent, cCustomer.Text, , "kodeanggota,nama")
  If Not dbData.EOF Then
    cCustomer.Text = cCustomer.Browse(dbData, Array("Kode", "Nama", "Dep", "Alamat"), , Array(6, 20, 6, 10))
    cCustomer.Text = GetNull(dbData!kodeanggota)
    cNamaCustomer.Text = GetNull(dbData!nama, "")
    cAlamat.Text = GetNull(dbData!alamat, "")
  End If
End Sub

Private Sub cmdSimpan_Click()
Dim lSave As Boolean
Dim Faktur As String
Dim n As Integer
lSave = True

  If ValidSaving Then
      'simpan pada tabel totpenjualan
      'simpan pada tabel penjualan
      'simpan pada tabel kartustock
      'simpan pada tabel kartupiutang
      
      objData.Start GetDSN
      Faktur = cFaktur.Text
      lSave = IIf(lSave, objData.Update(GetDSN, "totrtnpenjualan", "nomorreturpenjualan = '" & Faktur & "'", Array("nomorreturpenjualan", "fakturasli", "tgl", "jthtmp", "kodeanggota", "ppn", "persdisc", "persdisc2", "subtotal", "pajak", "discount", "discount2", "total", "tunai", "piutang", "datetime", "username", "nomorpenjualan"), Array(Faktur, cFakturAsli.Text, Format(dTgl.Value, "yyyy-MM-dd"), Format(dJthTmp.Value, "yyyy-MM-dd"), cCustomer.Text, nPPn.Value, nPersDisc.Value, 0, nSubTotal.Value, nPajak.Value, nDiscount.Value, 0, nTotal.Value, nTunai.Value, nPiutang.Value, SNow, GetRegistry(reg_username), cNomorPembelian.Text)), False)
      lSave = IIf(lSave, objData.Delete(GetDSN, "returpenjualan", "nomorreturpenjualan", sisAssign, Faktur), False)
      lSave = IIf(lSave, objData.Delete(GetDSN, "kartustock", "nomor", sisAssign, Faktur), False)
      For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
        lSave = IIf(lSave, objData.Add(GetDSN, "returpenjualan", Array("nomorreturpenjualan", "tgl", "kodestock", "qty", "harga", "kodesatuan", "discount", "jumlah"), Array(Faktur, Format(dTgl.Value, "yyyy-MM-dd"), vaArray(n, 8), vaArray(n, 3), vaArray(n, 5), vaArray(n, 4), vaArray(n, 6), vaArray(n, 7))), False)
        lSave = IIf(lSave, UpdKartuStock(objData, SisKartuStock.returpenjualan, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), vaArray(n, 8), vaArray(n, 3), vaArray(n, 5), vaArray(n, 6), "Retur Penjualan an. " & cNamaCustomer.Text, aCfg(objData, msGudangPenjualan), GetHargaPokok(objData, vaArray(n, 8))), False)
      Next n
      lSave = IIf(lSave, UpdKartuHutang(objData, SisKartuHutang.SisReturPenjualan, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), cCustomer.Text, "Retur Penjualan an. " & cNamaCustomer.Text, nPiutang.Value, SNow, GetRegistry(reg_username)), False)
      
      lSave = IIf(lSave, DelKodeTr(objData, vbTrigger.msReturPenjualan, Faktur), False)
      
      'Piutang, Kas
      'Kas, piutang
      '   Penjualan
      
      'Diskon Penjualan
      '   Penjualan
      
      'PPn Penjualan
      '   Penjualan
      
      'Inventory
      '
      lSave = IIf(lSave, UpdKodeTr(objData, msReturPenjualan, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), GetAkunMember(objData, cCustomer.Text), GetCostCenterUser(objData, GetRegistry(reg_username)), "Retur Penjualan an " & cNamaCustomer.Text, 0, nPiutang.Value), False)
          lSave = IIf(lSave, UpdKodeTr(objData, msReturPenjualan, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningPenjualan), GetCostCenterUser(objData, GetRegistry(reg_username)), "Retur Penjualan an " & cNamaCustomer.Text, nTunai.Value + nPiutang.Value, 0), False)
      
      'Debet
      Dim nTmp As Double
      Dim nSaldoTmp As Double
      Dim nTmpCOGS As Double
      Dim nTmpSaldoCOGS As Double
      Dim db As New ADODB.Recordset
      
      For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
        'Discount Pembelian per item
        nTmp = vaArray(n, 3) * vaArray(n, 5) - vaArray(n, 7)
        nSaldoTmp = nSaldoTmp + nTmp
        lSave = IIf(lSave, UpdKodeTr(objData, msReturPenjualan, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningDiscountPenjualan), GetCostCenterUser(objData, GetRegistry(reg_username)), "Dsc Item Retur Penjualan an " & cNamaCustomer.Text, 0, nTmp, "", SNow), False)
        
        Set db = objData.Browse(GetDSN, "stock", "kodestock,asbiaya", "kodestock", sisAssign, vaArray(n, 8))
        If Not db.EOF Then
          If GetNull(db!asbiaya) <> "1" Then
            'posting cogs
            nTmpCOGS = vaArray(n, 3) * GetHargaPokok(objData, vaArray(n, 8))
            lSave = IIf(lSave, UpdKodeTr(objData, msReturPenjualan, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningCOGS), GetCostCenterUser(objData, GetRegistry(reg_username)), "COGS Retur Penjualan an " & vaArray(n, 2), 0, nTmpCOGS, "", SNow), False)
              lSave = IIf(lSave, UpdKodeTr(objData, msReturPenjualan, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), GetAkunInventory(objData, vaArray(n, 8)), GetCostCenterUser(objData, GetRegistry(reg_username)), "COGS Retur Penjualan an " & vaArray(n, 2), nTmpCOGS, 0, "", SNow), False)
          End If
        End If
      Next n
      
      'Kredit
      lSave = IIf(lSave, UpdKodeTr(objData, msReturPenjualan, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningPenjualan), GetCostCenterUser(objData, GetRegistry(reg_username)), "Dsc Item Retur Penjualan an  " & cNamaCustomer.Text, nSaldoTmp, 0), False)
      
      'PPn
      lSave = IIf(lSave, UpdKodeTr(objData, msReturPenjualan, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, SisCfg.msRekeningPPnPenjualan), GetCostCenterUser(objData, GetRegistry(reg_username)), "PPn Retur Penjualan an " & cNamaCustomer.Text, nPajak.Value, 0, "", SNow), False)
          lSave = IIf(lSave, UpdKodeTr(objData, msReturPenjualan, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningPenjualan), GetCostCenterUser(objData, GetRegistry(reg_username)), "PPn Retur Penjualan an " & cNamaCustomer.Text, 0, nPajak.Value), False)
          
      'Discount seluruhnya
      lSave = IIf(lSave, UpdKodeTr(objData, msReturPenjualan, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningDiscountPenjualan), GetCostCenterUser(objData, GetRegistry(reg_username)), "Dsc Total Retur Penjualan an " & cNamaCustomer.Text, 0, nDiscount.Value, "", SNow), False)
          lSave = IIf(lSave, UpdKodeTr(objData, msReturPenjualan, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningPenjualan), GetCostCenterUser(objData, GetRegistry(reg_username)), "Dsc Total Retur Penjualan an " & cNamaCustomer.Text, nDiscount.Value, 0, "", SNow), False)
          
      If lSave Then
        objData.Save GetDSN
      Else
        objData.Cancel GetDSN
      End If
      initvalue
      GetEdit False
  End If
End Sub

Private Sub cNama_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "stock s", "s.kodestock,s.Barcode,s.nama,s.kodesatuan,s.hargajual", "s.nama", sisContent, cNama.Text, " and s.statusnonaktif <> 1")
  If Not dbData.EOF Then
    cNama.Text = cNama.Browse(dbData)
    GetDataStock
  Else
    cNama.Default
  End If
End Sub

Private Sub cNamaCustomer_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "anggota", "kodeanggota,nama,kodedep,alamat", "nama", sisContent, cNamaCustomer.Text, , "kodeanggota,nama")
  If Not dbData.EOF Then
    cNamaCustomer.Text = cNamaCustomer.Browse(dbData, Array("Kode", "Nama", "Dep", "Alamat"), , Array(6, 20, 6, 10))
    cCustomer.Text = GetNull(dbData!kodeanggota)
    cNamaCustomer.Text = GetNull(dbData!nama, "")
    cAlamat.Text = GetNull(dbData!alamat, "")
  End If
End Sub

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
  
  Set dbData = objData.Browse(GetDSN, "costcenter", , "kodecostcenter", sisAssign, GetCostCenterUser(objData, GetRegistry(reg_username)))
  If Not dbData.EOF Then
    lbCostCenter.Caption = "Cost Centre : " & GetNull(dbData!keterangan)
  End If
  
  TabIndex dTgl, n

  TabIndex cCustomer, n
  TabIndex cNamaCustomer, n
  TabIndex cAlamat, n
  TabIndex cFaktur, n
  
  TabIndex cFakturAsli, n
  TabIndex dJthTmp, n
  TabIndex nPersDisc, n
  TabIndex nPPn, n
  
  TabIndex nNomor, n
  TabIndex cBarcode, n
  TabIndex cNama, n
  TabIndex nQty, n
  TabIndex cSatuan, n
  TabIndex nHarga, n
  TabIndex nDisc1, n
  TabIndex nJumlah, n
  TabIndex cmdOK, n
  TabIndex nTunai, n
  
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
  dJthTmp.Value = Date
  nPersDisc.Value = 0
  nPPn.Value = 0
  cFakturAsli.Default
  cCustomer.Default
  cNamaCustomer.Default
  cAlamat.Default
  cKota.Default
  nSubTotal.Value = 0
  nPajak.Value = 0
  nDiscount.Value = 0
  nTotal.Value = 0
  nTunai.Value = 0
  nPiutang.Value = 0
  
  vaArray.ReDim 0, -1, 0, 8
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
  InitValue1
End Sub

Private Sub nDisc1_Validate(Cancel As Boolean)
  SumJumlah
End Sub

Private Sub nDisc2_Validate(Cancel As Boolean)
  SumJumlah
End Sub

Private Sub nDiscount_Validate(Cancel As Boolean)
  SumTotal
End Sub

Private Sub nDiscount2_Validate(Cancel As Boolean)
  SumTotal
End Sub

Private Sub InitValue1()
  nNomor.Value = 1
  cBarcode.Default
  cNama.Default
  nQty.Value = 1
  cSatuan.Default
  nHarga.Value = 0
  nDisc1.Default
  nJumlah.Value = 0
  cKode = ""
End Sub

Private Sub GetEdit(lPar As Boolean)
  BiSAFrame1.Enabled = lPar
  BisaFrame2.Enabled = lPar
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

Private Sub nBiaya_Validate(Cancel As Boolean)
  SumTotal
End Sub

Private Sub SumJumlah()
Dim nSubJumlah As Double

  nSubJumlah = nHarga.Value * nQty.Value
  nSubJumlah = nSubJumlah - (nSubJumlah * (nDisc1.Value / 100))
  nJumlah.Value = nSubJumlah
End Sub

Private Sub nHarga_Validate(Cancel As Boolean)
  SumJumlah
End Sub

Private Sub nNomor_Validate(Cancel As Boolean)
Dim n As Single
  
  If GetValidNomorUrut(nNomor, vaArray) Then
    n = nNomor.Value - 1
    If n <= vaArray.UpperBound(1) Then
      cBarcode.Text = vaArray(n, 1)
      cNama.Text = vaArray(n, 2)
      nQty.Value = vaArray(n, 3)
      cSatuan.Text = vaArray(n, 4)
      nHarga.Value = vaArray(n, 5)
      nDisc1.Value = vaArray(n, 6)
      nJumlah.Value = vaArray(n, 7)
      cKode = vaArray(n, 8)
    End If
  End If
End Sub

Private Sub nPersDisc_Validate(Cancel As Boolean)
  SumTotal
End Sub

Private Sub nPersDisc2_Validate(Cancel As Boolean)
  SumTotal
End Sub

Private Sub nPPn_Validate(Cancel As Boolean)
  SumTotal
End Sub

Private Sub nQty_Validate(Cancel As Boolean)
  SumJumlah
End Sub

Private Sub nTunai_Validate(Cancel As Boolean)
  SumBayar
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
      nNomor.Value = vaArray.UpperBound(1) + 2
      TDBGrid1.ReBind
    End If
  End If
End Sub

