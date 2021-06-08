VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trBuyBack 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRANSAKSI BUY BACK"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   17190
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   17190
   Begin BiSAFramProject.BiSAFrame BisaFrame2 
      Height          =   4515
      Left            =   75
      Top             =   2775
      Width           =   17040
      _ExtentX        =   30057
      _ExtentY        =   7964
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
      Begin BiSANumberBoxProject.BiSANumberBox nTunai 
         Height          =   330
         Left            =   15330
         TabIndex        =   0
         Top             =   3960
         Width           =   1560
         _ExtentX        =   2752
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
         Left            =   8115
         TabIndex        =   1
         Top             =   45
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   582
         Appearance      =   0
         Decimals        =   0
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
         Left            =   10260
         TabIndex        =   2
         Top             =   45
         Width           =   1620
         _ExtentX        =   2858
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
      Begin BiSATextBoxProject.BiSATextBox cSatuan 
         Height          =   330
         Left            =   8925
         TabIndex        =   3
         Top             =   45
         Width           =   1305
         _ExtentX        =   2302
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
         Left            =   2655
         TabIndex        =   4
         Top             =   45
         Width           =   5445
         _ExtentX        =   9604
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
         Left            =   675
         TabIndex        =   5
         Top             =   45
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
         GetPicture      =   1
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
         TabIndex        =   6
         Top             =   45
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
         Height          =   3375
         Left            =   105
         TabIndex        =   7
         Top             =   420
         Width           =   16785
         _ExtentX        =   29607
         _ExtentY        =   5953
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
         Columns(5).NumberFormat=   "###,###,###,###"
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
         Columns(7).NumberFormat=   "###,###,###,###"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "KODESTOCK"
         Columns(8).DataField=   ""
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "ID"
         Columns(9).DataField=   ""
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "Selisih COGS"
         Columns(10).DataField=   ""
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "H Pokok"
         Columns(11).DataField=   ""
         Columns(11).NumberFormat=   "###,###,###,###"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   12
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=12"
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
         Splits(0)._ColumnProps(11)=   "Column(2).Width=9710"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=9631"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=512"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=1455"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1376"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=197122"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2275"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2196"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=513"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=2990"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2910"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=514"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=1852"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=1773"
         Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=514"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(36)=   "Column(7).Width=3466"
         Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=3387"
         Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=514"
         Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(41)=   "Column(8).Width=714"
         Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=635"
         Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=516"
         Splits(0)._ColumnProps(45)=   "Column(8).Visible=0"
         Splits(0)._ColumnProps(46)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(47)=   "Column(9).Width=132"
         Splits(0)._ColumnProps(48)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(49)=   "Column(9)._WidthInPix=53"
         Splits(0)._ColumnProps(50)=   "Column(9)._ColStyle=516"
         Splits(0)._ColumnProps(51)=   "Column(9).Visible=0"
         Splits(0)._ColumnProps(52)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(53)=   "Column(10).Width=2725"
         Splits(0)._ColumnProps(54)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(55)=   "Column(10)._WidthInPix=2646"
         Splits(0)._ColumnProps(56)=   "Column(10)._ColStyle=516"
         Splits(0)._ColumnProps(57)=   "Column(10).Visible=0"
         Splits(0)._ColumnProps(58)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(59)=   "Column(11).Width=2725"
         Splits(0)._ColumnProps(60)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(61)=   "Column(11)._WidthInPix=2646"
         Splits(0)._ColumnProps(62)=   "Column(11)._ColStyle=514"
         Splits(0)._ColumnProps(63)=   "Column(11).Order=12"
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
         HeadLines       =   1,5
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
         _StyleDefs(69)  =   "Splits(0).Columns(8).Style:id=74,.parent=13"
         _StyleDefs(70)  =   "Splits(0).Columns(8).HeadingStyle:id=71,.parent=14"
         _StyleDefs(71)  =   "Splits(0).Columns(8).FooterStyle:id=72,.parent=15"
         _StyleDefs(72)  =   "Splits(0).Columns(8).EditorStyle:id=73,.parent=17"
         _StyleDefs(73)  =   "Splits(0).Columns(9).Style:id=70,.parent=13"
         _StyleDefs(74)  =   "Splits(0).Columns(9).HeadingStyle:id=67,.parent=14"
         _StyleDefs(75)  =   "Splits(0).Columns(9).FooterStyle:id=68,.parent=15"
         _StyleDefs(76)  =   "Splits(0).Columns(9).EditorStyle:id=69,.parent=17"
         _StyleDefs(77)  =   "Splits(0).Columns(10).Style:id=86,.parent=13"
         _StyleDefs(78)  =   "Splits(0).Columns(10).HeadingStyle:id=83,.parent=14"
         _StyleDefs(79)  =   "Splits(0).Columns(10).FooterStyle:id=84,.parent=15"
         _StyleDefs(80)  =   "Splits(0).Columns(10).EditorStyle:id=85,.parent=17"
         _StyleDefs(81)  =   "Splits(0).Columns(11).Style:id=78,.parent=13,.alignment=1"
         _StyleDefs(82)  =   "Splits(0).Columns(11).HeadingStyle:id=75,.parent=14"
         _StyleDefs(83)  =   "Splits(0).Columns(11).FooterStyle:id=76,.parent=15"
         _StyleDefs(84)  =   "Splits(0).Columns(11).EditorStyle:id=77,.parent=17"
         _StyleDefs(85)  =   "Named:id=33:Normal"
         _StyleDefs(86)  =   ":id=33,.parent=0"
         _StyleDefs(87)  =   "Named:id=34:Heading"
         _StyleDefs(88)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(89)  =   ":id=34,.wraptext=-1"
         _StyleDefs(90)  =   "Named:id=35:Footing"
         _StyleDefs(91)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(92)  =   "Named:id=36:Selected"
         _StyleDefs(93)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(94)  =   "Named:id=37:Caption"
         _StyleDefs(95)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(96)  =   "Named:id=38:HighlightRow"
         _StyleDefs(97)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(98)  =   "Named:id=39:EvenRow"
         _StyleDefs(99)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(100) =   "Named:id=40:OddRow"
         _StyleDefs(101) =   ":id=40,.parent=33"
         _StyleDefs(102) =   "Named:id=41:RecordSelector"
         _StyleDefs(103) =   ":id=41,.parent=34"
         _StyleDefs(104) =   "Named:id=42:FilterBar"
         _StyleDefs(105) =   ":id=42,.parent=33"
      End
      Begin BiSANumberBoxProject.BiSANumberBox nJumlah 
         Height          =   330
         Left            =   13005
         TabIndex        =   8
         Top             =   45
         Width           =   1905
         _ExtentX        =   3360
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
         Left            =   16530
         TabIndex        =   9
         Top             =   45
         Width           =   360
         _ExtentX        =   635
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
         Picture         =   "trBuyBack.frx":0000
      End
      Begin BiSANumberBoxProject.BiSANumberBox nDisc1 
         Height          =   330
         Left            =   11910
         TabIndex        =   10
         Top             =   45
         Width           =   1080
         _ExtentX        =   1905
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
         BackColor       =   -2147483634
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
         Height          =   645
         Left            =   105
         Top             =   3780
         Width           =   1635
         _ExtentX        =   2884
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
            Left            =   90
            TabIndex        =   11
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
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
         Left            =   1755
         Top             =   3780
         Width           =   1515
         _ExtentX        =   2672
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
            Left            =   75
            TabIndex        =   12
            Top             =   240
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
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
         Left            =   3300
         Top             =   3780
         Width           =   1455
         _ExtentX        =   2566
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
            Left            =   90
            TabIndex        =   13
            Top             =   240
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
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
         Left            =   4785
         Top             =   3780
         Width           =   1635
         _ExtentX        =   2884
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
            Left            =   90
            TabIndex        =   14
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
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
      Begin BiSANumberBoxProject.BiSANumberBox nHargaPokok 
         Height          =   330
         Left            =   14925
         TabIndex        =   37
         Top             =   45
         Width           =   1575
         _ExtentX        =   2778
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
         Left            =   14400
         TabIndex        =   15
         Top             =   4050
         Width           =   570
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   2775
      Left            =   0
      Top             =   0
      Width           =   17040
      _ExtentX        =   30057
      _ExtentY        =   4895
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
      Begin VB.Frame Frame2 
         Height          =   2775
         Left            =   180
         TabIndex        =   30
         Top             =   0
         Width           =   9000
         Begin BiSATextBoxProject.BiSABrowse cNamaSupplier 
            Height          =   330
            Left            =   3915
            TabIndex        =   31
            Top             =   1140
            Width           =   3045
            _ExtentX        =   5371
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
         Begin BiSATextBoxProject.BiSATextBox cAlamat 
            Height          =   330
            Left            =   630
            TabIndex        =   32
            Top             =   1500
            Width           =   6315
            _ExtentX        =   11139
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
            Left            =   630
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   1140
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
            Enabled         =   0   'False
            Appearance      =   0
            Caption         =   "Customer"
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
            Left            =   630
            TabIndex        =   34
            Top             =   765
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
            Left            =   630
            TabIndex        =   35
            Top             =   1875
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
            Left            =   705
            TabIndex        =   36
            Top             =   480
            Width           =   4965
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2775
         Left            =   9210
         TabIndex        =   22
         Top             =   0
         Width           =   7755
         Begin BiSATextBoxProject.BiSATextBox cFakturAsli 
            Height          =   330
            Left            =   720
            TabIndex        =   23
            Top             =   375
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
         Begin BiSADateProject.BiSADate dJthTmp 
            Height          =   330
            Left            =   735
            TabIndex        =   24
            Top             =   735
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
         Begin BiSANumberBoxProject.BiSANumberBox nPersDisc 
            Height          =   330
            Left            =   750
            TabIndex        =   25
            Top             =   1110
            Width           =   1980
            _ExtentX        =   3493
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
         Begin BiSANumberBoxProject.BiSANumberBox nPPn 
            Height          =   345
            Left            =   2835
            TabIndex        =   26
            Top             =   1095
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   609
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
         Begin BiSATextBoxProject.BiSABrowse cAkunKas 
            Height          =   330
            Left            =   750
            TabIndex        =   27
            Top             =   1485
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
         Begin BiSATextBoxProject.BiSABrowse cGudang 
            Height          =   330
            Left            =   750
            TabIndex        =   28
            Top             =   1860
            Width           =   2310
            _ExtentX        =   4075
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
            Caption         =   "Gudang"
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
         Begin BiSATextBoxProject.BiSATextBox cNamaGudang 
            Height          =   330
            Left            =   3105
            TabIndex        =   29
            Top             =   1860
            Width           =   2445
            _ExtentX        =   4313
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
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   4725
            Top             =   825
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   825
      Left            =   150
      Top             =   7290
      Width           =   16875
      _ExtentX        =   29766
      _ExtentY        =   1455
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
         TabIndex        =   16
         Top             =   240
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
         Picture         =   "trBuyBack.frx":059A
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   14145
         TabIndex        =   17
         Top             =   240
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
         Picture         =   "trBuyBack.frx":0824
      End
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   1170
         TabIndex        =   18
         Top             =   240
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
         Picture         =   "trBuyBack.frx":09C3
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   435
         Left            =   90
         TabIndex        =   19
         Top             =   240
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
         Picture         =   "trBuyBack.frx":0AEF
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   15675
         TabIndex        =   20
         Top             =   240
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
         Picture         =   "trBuyBack.frx":0C9A
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   14595
         TabIndex        =   21
         Top             =   240
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
         Picture         =   "trBuyBack.frx":0D40
      End
   End
End
Attribute VB_Name = "trBuyBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nPos As SisPos
Dim lEdit As Boolean

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim objMenu As New CodeSuiteLibrary.Menu

Dim vaArray As New XArrayDB
Dim vaDelete As New XArrayDB
Dim vaExport As New XArrayDB

Dim cKode As String
Dim cID
Dim nSaldoStock As Double

Private Sub cAkunKas_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", "kodeakun,keterangan", "jenis", sisAssign, "D", , "kodeakun")
  If Not dbData.EOF Then
    cAkunKas.Text = cAkunKas.Browse(dbData)
  End If
End Sub

Private Sub cBarcode_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "stock s", "s.kodestock,s.barcode,s.nama,s.kodesatuan,s.hargajual,s.diskonpenjualan,s.cogs,s.hargabeli", "s.barcode", sisContent, cBarcode.Text, " AND s.jenis < 9 and s.statusnonaktif <> 1")
  If Not dbData.EOF Then
    cBarcode.Text = cBarcode.Browse(dbData)
    GetDataStock
  Else
    cBarcode.Default
  End If
End Sub

Private Sub cFaktur_ButtonClick()
Dim db As New ADODB.Recordset
Dim n As Integer
Dim lSave As Boolean

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
  
  Set db = objData.Browse(GetDSN, "totbuyback", "nomorbuyback,tgl,subtotal,total,hutang", "nomorbuyback", sisContent, cFaktur.Text, " and tgl = '" & Format(dTgl.value, "yyyy-MM-dd") & "' and kodeanggota = '" & cSupplier.Text & "'")
  If Not db.EOF Then
    cFaktur.Text = cFaktur.Browse(db)
    'ambil nilai total
    Set db = objData.Browse(GetDSN, "totbuyback t", "t.*,g.keterangan as namagudang", "t.nomorbuyback", sisAssign, cFaktur.Text, , , Array("left join gudang g on g.kodegudang = t.kodegudang"))
    If Not db.EOF Then
      
      cFakturAsli.Text = GetNull(db!fakturasli, "")
      dJthTmp.value = GetNull(db!jthtmp)
      nPersDisc.value = GetNull(db!PersDisc, 0)
      nPPn.value = GetNull(db!ppn, 0)
      nSubTotal.value = GetNull(db!Subtotal, 0)
      nDiscount.value = GetNull(db!Discount, 0)
      nPajak.value = GetNull(db!PAJAK, 0)
      nTotal.value = GetNull(db!Total, 0)
      nTunai.value = GetNull(db!Tunai, 0)

      cAkunKas.Text = GetNull(db!kodeakun)
      
      cGudang.Text = GetNull(db!kodegudang, "")
      cNamaGudang.Text = GetNull(db!namagudang, "")
      
      
    End If
    'ambil nilai detail
    Dim nQtyTmp As Single
    nQtyTmp = 0
    Set db = objData.Browse(GetDSN, "buyback p", "s.barcode,p.kodestock,s.Nama,p.qty,p.kodesatuan,p.harga,p.discount,p.jumlah", "nomorbuyback", sisAssign, cFaktur.Text, , , Array("Left join stock s on s.kodestock = p.kodestock"))
    If Not db.EOF Then
      vaArray.ReDim 0, -1, 0, 10
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
        nQtyTmp = nQtyTmp + vaArray(n, 3)
        db.MoveNext
      Loop
      Set TDBGrid1.Array = vaArray
      TDBGrid1.ReBind
      TDBGrid1.Refresh
      TDBGrid1.Columns(3).FooterText = Format(nQtyTmp, "###,###,##0")
      Me.Refresh
      nNomor.value = vaArray.UpperBound(1) + 2
    End If
    
    If nPos = Delete Then
      If MsgBox("Data akan dihapus?", vbYesNo) = vbYes Then
        objData.Start GetDSN
        lSave = IIf(lSave, DelKodeTr(objData, msBuyBack, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "buyback", "nomorbuyback", sisAssign, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "kartustock", "nomor", sisAssign, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "totbuyback", "nomorbuyback", sisAssign, cFaktur.Text), False)
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

Private Sub cFaktur_Validate(Cancel As Boolean)
  If nPos = Edit Then
      cFaktur.Enabled = False
      cFaktur.BackColor = vbButtonFace
  End If
End Sub

Private Sub cGudang_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "gudang", "kodegudang,keterangan", "lstatus", sisAssign, "A")
  If Not dbData.EOF Then
    cGudang.Text = cGudang.Browse(dbData)
    cGudang.Text = GetNull(dbData!kodegudang)
    cNamaGudang.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub GetDataStock()
  cBarcode.Text = GetNull(dbData!barcode, "")
  cKode = GetNull(dbData!KodeStock, "")
  cNama.Text = GetNull(dbData!nama, "")
  cSatuan.Text = GetNull(dbData!kodesatuan, "")
  nHarga.value = GetNull(dbData!HargaJual, 0)
  nDisc1.value = GetNull(dbData!diskonpenjualan)
  nHargaPokok.value = IIf(GetNull(dbData!cogs, 0) = 0, GetNull(dbData!hargabeli), GetNull(dbData!cogs))
End Sub

Private Sub cGudang_Validate(Cancel As Boolean)
  cGudang.Enabled = False
End Sub

Private Sub cmdAdd_Click()
  nPos = Add
  GetEdit True
  GetFakturBrowse False
  cFaktur.Text = CreateNomorFaktur(objData, sisModulTransaksi.BuyBack, "totbuyback", "nomorbuyback")
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
  GetFakturBrowse True

End Sub

Private Sub GetFakturBrowse(ByVal lStat As Boolean)
  cFaktur.Button = lStat
  If lStat = True Then
    cFaktur.BackColor = vbWindowBackground
  Else
    cFaktur.BackColor = vbButtonFace
  End If
End Sub

Private Sub GetEdit(lPar As Boolean)
  BiSAFrame1.Enabled = lPar
  BiSAFrame2.Enabled = lPar
  lEdit = lPar
  initvalue
  SetButton cmdSimpan, cmdKeluar, cmdAdd, cmdEdit, cmdHapus, nPos, lPar, cmdAktivasi
  GetFakturBrowse False
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

Private Sub cmdHapus_Click()
  If GetRegistry(reg_UserLevel) <> 0 Then
    If objMenu.GetPassword("", Me, GetDSN) Then
      If objMenu.UserLevel <> 0 Then
        MsgBox "Maaf, Anda tidak diberikan wewenang untuk melakukan PENGHAPUSAN." & vbCrLf & _
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
    MsgBox "Barang tersebut tidak ada dalam database"
    cBarcode.SetFocus
    validOK = False
    Exit Function
'  Else
'    'cek apakah harga yg diberikan sudah benar
'    Set dbData = objData.Browse(GetDSN, "stock s", "kodestock,hargabeli,diskonpenjualan,cogs", "kodestock", sisAssign, cKode)
'    If Not dbData.EOF Then
'      'If nHarga.Value - (nHarga.Value * nDisc1.Value / 100) > dbData!hargabeli - (dbData!hargabeli * dbData!diskonpenjualan / 100) Then
'      If nHarga.Value - (nHarga.Value * nDisc1.Value / 100) > dbData!cogs Then
'        MsgBox "Barang tersebut terlalu mahal untuk di buyback" & vbCrLf & "Paling mahal yg bisa kami buyback " & Format(dbData!cogs, "###,###,##0.00")
'        cBarcode.SetFocus
'        validOK = False
'        Exit Function
'      End If
'    End If
  End If
End Function

Private Sub cmdOK_Click()
Dim n As Integer
Dim nJumlah1 As Double
Dim nQtyTmp As Single

  If validOK() Then
   
    'jika baris <= Nomor
    If vaArray.UpperBound(1) + 2 <= nNomor.value Then
      vaArray.ReDim 0, nNomor.value - 1, 0, 11
      n = vaArray.UpperBound(1)
    'jika baris=0
    ElseIf vaArray.UpperBound(1) = -1 Then
      nNomor.value = 1
      vaArray.ReDim 0, nNomor.value - 1, 0, 11
      n = vaArray.UpperBound(1)
    Else
      n = nNomor.value - 1
    End If
        
    vaArray(n, 0) = nNomor.value
    vaArray(n, 1) = cBarcode.Text
    vaArray(n, 2) = cNama.Text
    vaArray(n, 3) = nQty.value
    vaArray(n, 4) = cSatuan.Text
    vaArray(n, 5) = nHarga.value
    vaArray(n, 6) = nDisc1.value
    vaArray(n, 7) = nJumlah.value
    vaArray(n, 8) = cKode
    vaArray(n, 9) = cID
    vaArray(n, 10) = 0 'untuk mendapatkan selisih cogs yg baru
    vaArray(n, 11) = nHargaPokok.value 'cogs lama
          
    TDBGrid1.Array = vaArray
    TDBGrid1.ReBind
    TDBGrid1.MoveNext
    
    nJumlah1 = 0
    For n = 0 To vaArray.UpperBound(1)
      nJumlah1 = nJumlah1 + vaArray(n, 7)
      nQtyTmp = nQtyTmp + vaArray(n, 3)
    Next
    nSubTotal.value = nJumlah1
    TDBGrid1.Columns(3).FooterText = Format(nQtyTmp, "###,###,##0")
    
    SumTotal
    
    InitValue1
    
    nNomor.value = vaArray.UpperBound(1) + 2
    cBarcode.SetFocus
  End If
End Sub

Private Sub SumTotal()
Dim n As Double
  
  nSubTotal.value = 0
  For n = 0 To vaArray.UpperBound(1)
    nSubTotal.value = nSubTotal.value + vaArray(n, 7)
  Next
  
  If nPersDisc.Enabled = True Then
    nDiscount.value = nPersDisc.value / 100 * (nSubTotal.value)
  End If
  
  nPajak.value = (nPPn.value / 100) * (nSubTotal.value - (nDiscount.value + nDiscount.value))
  nTotal.value = nSubTotal.value + nPajak.value - nDiscount.value
  nTunai.value = nTotal.value

End Sub

Private Function ValidSaving() As Boolean
Dim n As Integer

  ValidSaving = True
  
  If cSupplier.Text = "" Then
    MsgBox "Kode Supplier tidak terisi" & vbCrLf & "Data tidak bisa disimpan", vbExclamation
    ValidSaving = False
    Exit Function
  End If
  
  If cAkunKas.Text = "" Then
    MsgBox "Akun Kas tidak terisi" & vbCrLf & "Data tidak bisa disimpan", vbExclamation
    ValidSaving = False
    Exit Function
  End If
  
  If Not GetValidDataBrowse(objData, "anggota", "kodeanggota", cSupplier.Text) Then
    MsgBox "Maaf, data supplier tidak tertera dengan benar" & vbCrLf & "Data tidak bisa disimpan", vbExclamation
    ValidSaving = False
    Exit Function
  End If
  
  If Not GetValidDataBrowse(objData, "gudang", "kodegudang", cGudang.Text) Then
    MsgBox "Kode gudang tidak tertera dengan benar" & vbCrLf & "Data tidak bisa disimpan", vbExclamation
    cGudang.SetFocus
    ValidSaving = False
    Exit Function
  End If
  
  'Jika kode gudang tidak aktif, maka penyimpanan data tidak diijinkan
  Set dbData = objData.Browse(GetDSN, "gudang", "lstatus", "kodegudang", sisAssign, cGudang.Text)
  If Not dbData.EOF Then
    If GetNull(dbData!lStatus) <> "A" Then
      MsgBox "Kode Gudang tidak valid, atau tidak aktif, Data tidak bisa disimpan", vbExclamation
      ValidSaving = False
      Exit Function
    End If
  End If

  
  If Not GetValidDataBrowse(objData, "akun", "kodeakun", cAkunKas.Text) Then
    MsgBox "Kode akun tidak tertera dengan benar" & vbCrLf & "Data tidak bisa disimpan", vbExclamation
    ValidSaving = False
    Exit Function
  End If
  
    
End Function

Private Sub cmdSimpan_Click()
Dim lSave As Boolean
Dim Faktur As String
Dim n As Integer
lSave = True

  'simpan pada tabel totbuyback
  'simpan pada tabel buyback
  'simpan pada tabel kartustock
  'simpan pada tabel kartuhutang
  'tidak diperlukan update harga beli ke master jika harga buyback tidak sama dengan harga dimaster
  
  If ValidSaving Then
    objData.Start GetDSN
    Faktur = cFaktur.Text
    lSave = IIf(lSave, objData.Update(GetDSN, "totbuyBack", "nomorbuyBack = '" & Faktur & "'", Array("nomorbuyBack", "fakturasli", "tgl", "jthtmp", "kodeanggota", "ppn", "persdisc", "persdisc2", "subtotal", "pajak", "discount", "discount2", "total", "tunai", "hutang", "datetime", "username", "kodeakun", "kodecostcenter", "kodesalesman", "statusbuyBack", "kodegudang"), Array(Faktur, cFakturAsli.Text, Format(dTgl.value, "yyyy-MM-dd"), Format(dJthTmp.value, "yyyy-MM-dd"), cSupplier.Text, nPPn.value, nPersDisc.value, 0, nSubTotal.value, nPajak.value, nDiscount.value, 0, nTotal.value, nTunai.value, 0, SNow, GetRegistry(reg_Username), cAkunKas.Text, GetCostCenterUser(objData, GetRegistry(reg_Username)), "", "", cGudang.Text)), False)
    lSave = IIf(lSave, objData.Delete(GetDSN, "buyBack", "nomorbuyBack", sisAssign, Faktur), False)
    lSave = IIf(lSave, objData.Delete(GetDSN, "kartustock", "nomor", sisAssign, Faktur), False)
    
    For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
      lSave = IIf(lSave, objData.Add(GetDSN, "buyBack", Array("nomorbuyBack", "kodegudang", "tgl", "kodestock", "qty", "harga", "kodesatuan", "discount", "jumlah"), Array(Faktur, cGudang.Text, Format(dTgl.value, "yyyy-MM-dd"), vaArray(n, 8), vaArray(n, 3), vaArray(n, 5), vaArray(n, 4), vaArray(n, 6), vaArray(n, 7))), False)
      '** PENTING **
      'UPDATE KARTU STOCK
      'Cek nilai persediaan terlebih dahulu
      'Jika nilai persediaan minus, gunakan HPP jika tidak gunakan Harga beli
      If GetSaldoStock(objData, "", vaArray(n, 8)) < 0 Then
'        vaArray(n, 5) = NewUpdHargaPokok(objData, vaArray(n, 8))
        'Update harga cogs dengan yg terakhir
        '** PENTING **
        vaArray(n, 10) = vaArray(n, 7)
        vaArray(n, 7) = NewUpdHargaPokok(objData, vaArray(n, 8)) * vaArray(n, 3)
        vaArray(n, 10) = vaArray(n, 10) - vaArray(n, 7)
        
        lSave = IIf(lSave, objData.Edit(GetDSN, "stock", "kodestock = '" & vaArray(n, 8) & "'", Array("cogs"), Array(NewUpdHargaPokok(objData, vaArray(n, 8)))), False)
        lSave = IIf(lSave, UpdKartuStock(objData, SisKartuStock.BuyBack, Faktur, Format(dTgl.value, "yyyy-MM-dd"), vaArray(n, 8), vaArray(n, 3), vaArray(n, 5), vaArray(n, 6), "Buy Back an. " & cNamaSupplier.Text & " Gudang " & cGudang.Text, cGudang.Text, NewUpdHargaPokok(objData, vaArray(n, 8))), False)
      Else
        'Update harga cogs dengan yg terakhir
        'lSave = IIf(lSave, objData.Edit(GetDSN, "stock", "kodestock = '" & vaArray(n, 8) & "'", Array("cogs"), Array(vaArray(n, 5) - (vaArray(n, 5) * vaArray(n, 6) / 100))), False)
        '** PENTING **
        'lSave = IIf(lSave, UpdKartuStock(objData, SisKartuStock.BuyBack, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), vaArray(n, 8), vaArray(n, 3), vaArray(n, 5), vaArray(n, 6), "Buy Back an. " & cNamaSupplier.Text & " Gudang " & cGudang.Text, cGudang.Text, vaArray(n, 5) - (vaArray(n, 5) * vaArray(n, 6) / 100)), False)
        lSave = IIf(lSave, UpdKartuStock(objData, SisKartuStock.BuyBack, Faktur, Format(dTgl.value, "yyyy-MM-dd"), vaArray(n, 8), vaArray(n, 3), vaArray(n, 5), vaArray(n, 6), "Buy Back an. " & cNamaSupplier.Text & " Gudang " & cGudang.Text, cGudang.Text, vaArray(n, 11)), False)
      End If
    Next n

    
    ' Inventory (1)
    ' Purchase Tax (2)
    ' Non Inventory Expenses (5)
    '    Acc Payable (2)
    '    Cash Bank (1)
    
    
    'Posting inventory
    'Hapus dulu di bukubesar
    lSave = IIf(lSave, DelKodeTr(objData, msBuyBack, Faktur), False)
    
    'Debet
    For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
      Dim db As New ADODB.Recordset
      
      Set db = objData.Browse(GetDSN, "stock", "kodestock,asbiaya", "kodestock", sisAssign, vaArray(n, 8))
      If Not db.EOF Then
        If GetNull(db!asbiaya) = "1" Then
'          lSave = IIf(lSave, UpdKodeTr(objData, msBuyBack, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningBiayaBarang), aCfg(objData, msCostCenterJualBeli), "BuyBack Inventory an " & cNamaSupplier.Text, vaArray(n, 3) * vaArray(n, 5), 0, "", SNow, vaArray(n, 8)), False)
          lSave = IIf(lSave, UpdKodeTr(objData, msBuyBack, Faktur, Format(dTgl.value, "yyyy-MM-dd"), aCfg(objData, msRekeningBiayaBarang), GetCostCenterUser(objData, GetRegistry(reg_Username)), "BuyBack Inventory an " & cNamaSupplier.Text, vaArray(n, 7), 0, "", SNow, vaArray(n, 8)), False)

          'Discount Pembelian per item
'          lSave = IIf(lSave, UpdKodeTr(objData, msBuyBack, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningDiscountPembelian), aCfg(objData, msCostCenterJualBeli), "Dsc Itm buyBack an " & cNamaSupplier.Text, 0, vaArray(n, 3) * vaArray(n, 5) - vaArray(n, 7), "", SNow, vaArray(n, 8)), False)
        Else
'          lSave = IIf(lSave, UpdKodeTr(objData, msBuyBack, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), GetAkunInventory(objData, vaArray(n, 8)), aCfg(objData, msCostCenterJualBeli), "BuyBack Inventory an " & cNamaSupplier.Text, vaArray(n, 3) * vaArray(n, 5), 0, "", SNow, vaArray(n, 8)), False)
          lSave = IIf(lSave, UpdKodeTr(objData, msBuyBack, Faktur, Format(dTgl.value, "yyyy-MM-dd"), GetAkunInventory(objData, vaArray(n, 8)), GetCostCenterUser(objData, GetRegistry(reg_Username)), "BuyBack Inventory an " & vaArray(n, 2), vaArray(n, 7), 0, "", SNow, vaArray(n, 8)), False)
          If vaArray(n, 10) <> 0 Then
            lSave = IIf(lSave, UpdKodeTr(objData, msBuyBack, Faktur, Format(dTgl.value, "yyyy-MM-dd"), aCfg(objData, msRekeningCOGS), GetCostCenterUser(objData, GetRegistry(reg_Username)), "COGS BuyBack an " & vaArray(n, 2), vaArray(n, 10), 0, "", SNow, vaArray(n, 8)), False)
          End If
          'Discount Pembelian per item
'          lSave = IIf(lSave, UpdKodeTr(objData, msBuyBack, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningDiscountPembelian), aCfg(objData, msCostCenterJualBeli), "Dsc Itm Pembelian an " & cNamaSupplier.Text, 0, vaArray(n, 3) * vaArray(n, 5) - vaArray(n, 7), "", SNow, vaArray(n, 8)), False)
        End If
      End If
      
      'Update COGS pada tabel stock
      'lSave = UpdHargaPokok(objData, vaArray(n, 8), vaArray(n, 3), vaArray(n, 5))
    Next n
    
    
    'PPn
    lSave = IIf(lSave, UpdKodeTr(objData, msBuyBack, Faktur, Format(dTgl.value, "yyyy-MM-dd"), aCfg(objData, msRekeningPPnPembelian), GetCostCenterUser(objData, GetRegistry(reg_Username)), "PPn Pembelian an " & cNamaSupplier.Text, nPajak.value, 0, "", SNow), False)
    'Discount seluruhnya
    lSave = IIf(lSave, UpdKodeTr(objData, msBuyBack, Faktur, Format(dTgl.value, "yyyy-MM-dd"), aCfg(objData, msRekeningDiscountPembelian), GetCostCenterUser(objData, GetRegistry(reg_Username)), "Dsc Tot Pembelian an " & cNamaSupplier.Text, 0, nDiscount.value, "", SNow), False)
    
    'Kredit
    'Tidak ada Hutang hanya kas
    'kas
    lSave = IIf(lSave, UpdKodeTr(objData, msBuyBack, Faktur, Format(dTgl.value, "yyyy-MM-dd"), cAkunKas.Text, GetCostCenterUser(objData, GetRegistry(reg_Username)), "Kas untuk Buy Back " & cNamaSupplier.Text, 0, nTunai.value, "", SNow), False)
    
    If lSave Then
      objData.Save GetDSN
    Else
      MsgBox "Maaf data tidak bisa disimpan", vbCritical, "Error"
      objData.Cancel GetDSN
    End If
    
'    GetCetak Faktur
    initvalue
    GetEdit False
  End If
End Sub

Private Sub GetCetak(ByVal cFak As String)
  trPrintBuyBack.noOrder = cFak
  Set dbData = objData.Browse(GetDSN, "totbuyback t", "t.*,a.*", "t.nomorbuyback", sisAssign, cFak, , , Array("left join anggota a on a.kodeanggota = t.kodeanggota"))
  If Not dbData.EOF Then
    trPrintBuyBack.nSubTotal = GetNull(dbData!Subtotal)
    trPrintBuyBack.nDiscount = 0
    trPrintBuyBack.nCash = GetNull(dbData!Tunai)
    trPrintBuyBack.nChange = GetNull(dbData!hutang)
    trPrintBuyBack.cKodeMember = GetNull(dbData!kodeanggota)
    trPrintBuyBack.cMember = GetNull(dbData!nama)
    trPrintBuyBack.cTeleponMember = 0
    trPrintBuyBack.Ups = 0
    Load trPrintBuyBack
    trPrintBuyBack.Show vbModal
  End If
End Sub

Private Sub cNamaSupplier_Validate(Cancel As Boolean)
  cNamaSupplier.Enabled = False
End Sub

Private Sub cSupplier_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "anggota", "kodeanggota,nama,alamat", "kodeanggota", sisContent, cSupplier.Text, , "kodeanggota,nama")
  If Not dbData.EOF Then
    cSupplier.Text = cSupplier.Browse(dbData)
    cSupplier.Text = GetNull(dbData!kodeanggota)
    cNamaSupplier.Text = GetNull(dbData!nama, "")
    cAlamat.Text = GetNull(dbData!alamat, "")
    
  End If
End Sub

Private Sub cNama_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "stock s", "s.kodestock,s.Barcode,s.nama,s.kodesatuan,s.hargajual,s.diskonpenjualan,s.cogs,s.hargabeli", "s.nama", sisContent, cNama.Text, " AND s.jenis < 9 and s.statusnonaktif <> 1")
  If Not dbData.EOF Then
    cNama.Text = cNama.Browse(dbData)
    GetDataStock
  Else
    cNama.Default
  End If
End Sub

Private Sub cNamaSupplier_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "anggota", "kodeanggota,nama,alamat", "nama", sisContent, cNamaSupplier.Text, , "kodeanggota,nama")
  If Not dbData.EOF Then
    cNamaSupplier.Text = cNamaSupplier.Browse(dbData)
    cSupplier.Text = GetNull(dbData!kodeanggota)
    cNamaSupplier.Text = GetNull(dbData!nama, "")
    cAlamat.Text = GetNull(dbData!alamat, "")
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

'  If CheckTrial(nRecordsTrial, TrialPembelian) = True Then
'    End
'  End If

  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  GetEdit False
  initvalue
  
  
  Set dbData = objData.Browse(GetDSN, "costcenter", , "kodecostcenter", sisAssign, GetCostCenterUser(objData, GetRegistry(reg_Username)))
  If Not dbData.EOF Then
    lbCostCenter.Caption = "Cost Centre : " & GetNull(dbData!keterangan)
  End If
  
  TabIndex dTgl, n
  TabIndex cSupplier, n
  TabIndex cNamaSupplier, n
  TabIndex cFaktur, n
  
  TabIndex cAlamat, n
  
  
  TabIndex cFakturAsli, n
  TabIndex dJthTmp, n
  TabIndex nPersDisc, n
  TabIndex nPPn, n
  TabIndex cAkunKas, n
  TabIndex cGudang, n
  TabIndex cNamaGudang, n
  
  
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
Dim dbgudang As New ADODB.Recordset

  

  cFaktur.Default
  dTgl.value = Date
  dJthTmp.value = Date
  nPersDisc.value = 0
  nPPn.value = 0
  cFakturAsli.Default
  cSupplier.Default
  cNamaSupplier.Default
  cNamaSupplier.Enabled = True
  cGudang.Enabled = True
  cAlamat.Default

  nSubTotal.value = 0
  nPajak.value = 0
  nDiscount.value = 0
  nTotal.value = 0
  nTunai.value = 0


  cAkunKas.Text = cKasTeller
  cGudang.Text = aCfg(objData, msGudangPembelian)
  cNamaGudang.Default
  Set dbgudang = objData.Browse(GetDSN, "gudang", "kodegudang,keterangan", "kodegudang", sisAssign, cGudang.Text)
  If Not dbgudang.EOF Then
    cNamaGudang.Text = GetNull(dbgudang!keterangan)
  End If

  
  vaArray.ReDim 0, -1, 0, 9
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
  InitValue1
  vaDelete.ReDim 0, -1, 0, 1
  
  TDBGrid1.Columns(3).FooterText = ""
  nQty.Decimals = aCfg(objData, msNilaiDecimals)
  
  cGudang.Enabled = True
  cGudang.BackColor = vbWhite
  If GetRegistry(reg_UserLevel) <> 0 Then
    cGudang.Enabled = False
    cGudang.BackColor = vbButtonFace
  End If
  cGudang.Text = GetGudangUser(objData, GetRegistry(reg_Username))
  Set dbData = objData.Browse(GetDSN, "gudang", "kodegudang,keterangan", "kodegudang", sisAssign, cGudang.Text)
  If Not dbData.EOF Then
    cNamaGudang.Text = GetNull(dbData!keterangan)
  Else
    cNamaGudang.Default
  End If
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
  nNomor.value = 1
  cBarcode.Default
  cNama.Default
  nQty.value = 1
  cSatuan.Default
  nHarga.value = 0
  nDisc1.value = aCfg(objData, msDiscountItemPembelian, 0)
  nJumlah.value = 0
  cKode = ""
  nHargaPokok.value = 0
End Sub

Private Sub nBiaya_Validate(Cancel As Boolean)
  SumTotal
End Sub

Private Sub SumJumlah()
Dim nSubJumlah As Double

  nSubJumlah = nHarga.value * nQty.value
  nSubJumlah = nSubJumlah - (nSubJumlah * (nDisc1.value / 100))
  nJumlah.value = nSubJumlah
End Sub

Private Sub nHarga_Validate(Cancel As Boolean)
  SumJumlah
End Sub

Private Sub nNomor_Validate(Cancel As Boolean)
Dim n As Single
  
  If GetValidNomorUrut(nNomor, vaArray) Then
    n = nNomor.value - 1
    If n <= vaArray.UpperBound(1) Then
      cBarcode.Text = vaArray(n, 1)
      cNama.Text = vaArray(n, 2)
      nQty.value = vaArray(n, 3)
      cSatuan.Text = vaArray(n, 4)
      nHarga.value = vaArray(n, 5)
      nDisc1.value = vaArray(n, 6)
      nJumlah.value = vaArray(n, 7)
      cKode = vaArray(n, 8)
      cID = vaArray(n, 9)
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

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim n As Integer
Dim nQtyTmp As Integer

  If KeyCode = vbKeyDelete Then
    If vaArray.UpperBound(1) >= 0 Then
      vaDelete.InsertRows vaDelete.UpperBound(1) + 1
      n = vaDelete.UpperBound(1)
      vaDelete(n, 0) = TDBGrid1.Columns(1).Text
      vaDelete(n, 1) = TDBGrid1.Columns(9).Text
      
      TDBGrid1.Delete
      SumTotal
      For n = 0 To vaArray.UpperBound(1)
        vaArray(n, 0) = n + 1
        nQtyTmp = nQtyTmp + vaArray(n, 3)
      Next
      TDBGrid1.Columns(3).FooterText = Format(nQtyTmp, "###,###,##0")
      nNomor.value = vaArray.UpperBound(1) + 2
      TDBGrid1.ReBind
    End If
  End If
End Sub


