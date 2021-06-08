VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trRefund 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Refund/Cash Back"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17100
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   17100
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   4890
      Left            =   135
      TabIndex        =   13
      Top             =   2775
      Width           =   16905
      Begin BiSANumberBoxProject.BiSANumberBox nQty 
         Height          =   330
         Left            =   9750
         TabIndex        =   17
         Top             =   180
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   582
         BorderStyle     =   0
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
         Left            =   11715
         TabIndex        =   19
         Top             =   180
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
         BorderStyle     =   0
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
         BackColor       =   12632319
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
         Left            =   10665
         TabIndex        =   18
         Top             =   180
         Width           =   1035
         _ExtentX        =   1826
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
         BackColor       =   12640511
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
         Left            =   2805
         TabIndex        =   16
         Top             =   180
         Width           =   6915
         _ExtentX        =   12197
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
         Left            =   615
         TabIndex        =   15
         Top             =   180
         Width           =   2175
         _ExtentX        =   3836
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
         Left            =   60
         TabIndex        =   14
         Top             =   180
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   582
         BorderStyle     =   0
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
      Begin BiSANumberBoxProject.BiSANumberBox nJumlah 
         Height          =   330
         Left            =   14715
         TabIndex        =   21
         Top             =   180
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   582
         BorderStyle     =   0
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
         BackColor       =   12640511
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
         Left            =   16440
         TabIndex        =   22
         Top             =   180
         Width           =   390
         _ExtentX        =   688
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
         Picture         =   "trRefund.frx":0000
      End
      Begin BiSANumberBoxProject.BiSANumberBox nRefund 
         Height          =   330
         Left            =   13245
         TabIndex        =   20
         Top             =   180
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   582
         BorderStyle     =   0
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   3780
         Left            =   75
         TabIndex        =   23
         Top             =   555
         Width           =   16740
         _ExtentX        =   29528
         _ExtentY        =   6668
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
         Columns(5).NumberFormat=   "###,###,###,###"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "REFUND (RP)"
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
         Columns.Count   =   10
         Splits(0)._UserFlags=   0
         Splits(0).Size  =   2
         Splits(0).Size.vt=   2
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=10"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=926"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=847"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=514"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=3863"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3784"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=512"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=12330"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=12250"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=512"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=1588"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1508"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=197122"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=1905"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=1826"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=513"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=2593"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2514"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=514"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=2699"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2619"
         Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=514"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(36)=   "Column(7).Width=3334"
         Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=3254"
         Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=514"
         Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(41)=   "Column(8).Width=1508"
         Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=1429"
         Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=516"
         Splits(0)._ColumnProps(45)=   "Column(8).Visible=0"
         Splits(0)._ColumnProps(46)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(47)=   "Column(9).Width=2725"
         Splits(0)._ColumnProps(48)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(49)=   "Column(9)._WidthInPix=2646"
         Splits(0)._ColumnProps(50)=   "Column(9)._ColStyle=516"
         Splits(0)._ColumnProps(51)=   "Column(9).Visible=0"
         Splits(0)._ColumnProps(52)=   "Column(9).Order=10"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         BorderStyle     =   0
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
         _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.transparentBmp=0,.borderSize=0"
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
         _StyleDefs(77)  =   "Named:id=33:Normal"
         _StyleDefs(78)  =   ":id=33,.parent=0"
         _StyleDefs(79)  =   "Named:id=34:Heading"
         _StyleDefs(80)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(81)  =   ":id=34,.wraptext=-1"
         _StyleDefs(82)  =   "Named:id=35:Footing"
         _StyleDefs(83)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(84)  =   "Named:id=36:Selected"
         _StyleDefs(85)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(86)  =   "Named:id=37:Caption"
         _StyleDefs(87)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(88)  =   "Named:id=38:HighlightRow"
         _StyleDefs(89)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(90)  =   "Named:id=39:EvenRow"
         _StyleDefs(91)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(92)  =   "Named:id=40:OddRow"
         _StyleDefs(93)  =   ":id=40,.parent=33"
         _StyleDefs(94)  =   "Named:id=41:RecordSelector"
         _StyleDefs(95)  =   ":id=41,.parent=34"
         _StyleDefs(96)  =   "Named:id=42:FilterBar"
         _StyleDefs(97)  =   ":id=42,.parent=33"
      End
      Begin BiSANumberBoxProject.BiSANumberBox nJumlahTotal 
         Height          =   330
         Left            =   13245
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   4455
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   582
         BorderStyle     =   0
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
         BackColor       =   12640511
         Caption         =   "Jumlah Total : "
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
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2820
      Left            =   150
      TabIndex        =   0
      Top             =   -30
      Width           =   16875
      Begin BiSATextBoxProject.BiSABrowse cNamaSupplier 
         Height          =   330
         Left            =   3540
         TabIndex        =   3
         Top             =   1035
         Width           =   3075
         _ExtentX        =   5424
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
         Left            =   4335
         TabIndex        =   5
         Top             =   1395
         Width           =   2280
         _ExtentX        =   4022
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
         Left            =   270
         TabIndex        =   4
         Top             =   1395
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
         Left            =   270
         TabIndex        =   2
         Top             =   1035
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
         BackColor       =   -2147483633
         Enabled         =   0   'False
         Appearance      =   0
         Caption         =   "Supplier"
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
         Left            =   285
         TabIndex        =   1
         Top             =   675
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   582
         Value           =   "16-01-2016"
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
         Left            =   270
         TabIndex        =   6
         Top             =   1755
         Width           =   4050
         _ExtentX        =   7144
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
         Left            =   2970
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   2130
         Width           =   2880
         _ExtentX        =   5080
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
         Left            =   270
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   2130
         Width           =   2655
         _ExtentX        =   4683
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
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   660
      Left            =   135
      Top             =   7680
      Width           =   16905
      _ExtentX        =   29819
      _ExtentY        =   1164
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
         Left            =   12945
         TabIndex        =   9
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
         Picture         =   "trRefund.frx":059A
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   14115
         TabIndex        =   10
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
         Picture         =   "trRefund.frx":0824
      End
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   75
         TabIndex        =   8
         Top             =   90
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
         Picture         =   "trRefund.frx":09C3
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   435
         Left            =   11805
         TabIndex        =   7
         Top             =   105
         Width           =   1110
         _ExtentX        =   1958
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
         Picture         =   "trRefund.frx":0AEF
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   15645
         TabIndex        =   12
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
         Picture         =   "trRefund.frx":0C9A
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   14550
         TabIndex        =   11
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
         Picture         =   "trRefund.frx":0D40
      End
   End
End
Attribute VB_Name = "trRefund"
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
Dim cStatusPembelian As String
Dim nDiskonExcel As Double
Dim nQtyTmp As Single

Dim nHargaJual As Double

Private Sub cBarcode_ButtonClick()
  If Len(cBarcode.Text) >= 3 Then
    Set dbData = objData.Browse(GetDSN, "stock s", "s.barcode,s.nama,s.hargabeli,s.kodesatuan,s.hargajual,s.kodestock", "s.barcode", sisContent, cBarcode.Text, " AND s.jenis < 9 and s.statusnonaktif <> 1")
    If Not dbData.EOF Then
      cBarcode.Text = cBarcode.Browse(dbData, Array("BARCODE", "NAMA", "BELI", "SATUAN"), , Array(10, 35, 10, 8))
      GetDataStock
      SumJumlah
    Else
      cBarcode.Default
    End If
  Else
    MsgBox "Ketikkan 3 karakter atau lebih pencarian", vbCritical
  End If
End Sub


Private Sub cFaktur_ButtonClick()
Dim db As New ADODB.Recordset
Dim n As Integer
Dim lSave As Boolean
Dim nTotal As Double

  nTotal = 0

  If aCfg(objData, msOtorisasiPenuh) = "Y" Then
    If objMenu.GetPassword("", Me, GetDSN) Then
      If objMenu.UserLevel <> 0 Then
          MsgBox "Maaf, Anda tidak diberikan wewenang untuk melakukan PENGEDITAN." & vbCrLf & _
                 "Hanya user dengan LEVEL 0 (SUPERVISOR) yg diijinkan", vbInformation, "OTORISASI not ALLOWED"
          Exit Sub
      End If
    Else
      Exit Sub
    End If
  End If

  lSave = True
  
  Set db = objData.Browse(GetDSN, "totrefund", "nomorrefund,tgl,subtotal,total,hutang,kodeakun", "nomorrefund", sisContent, cFaktur.Text, " and tgl = '" & Format(dTgl.value, "yyyy-MM-dd") & "' and kodesupplier = '" & cSupplier.Text & "'")
  If Not db.EOF Then
    cFaktur.Text = cFaktur.Browse(db)
    cFaktur.Text = GetNull(db!nomorrefund)
    cAkunKas.Text = GetNull(db!kodeakun)
    'ambil nilai total
    Set db = objData.Browse(GetDSN, "totrefund t", "t.*,g.keterangan as namagudang", "t.nomorrefund", sisAssign, cFaktur.Text, , , Array("left join gudang g on g.kodegudang = t.kodegudang"))
    If Not db.EOF Then
      cStatusPembelian = GetNull(db!statuspembelian)
    End If
    'ambil nilai detail
    Dim nQtyTmp As Single
    nQtyTmp = 0
    Set db = objData.Browse(GetDSN, "refund p", "s.barcode,p.kodestock,s.Nama,s.hargabeli,p.qty,p.kodesatuan,p.refund,p.jumlah", "nomorrefund", sisAssign, cFaktur.Text, , , Array("Left join stock s on s.kodestock = p.kodestock"))
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
        vaArray(n, 5) = GetNull(db!hargabeli)
        vaArray(n, 6) = GetNull(db!refund)
        vaArray(n, 7) = GetNull(db!jumlah)
        vaArray(n, 8) = GetNull(db!KodeStock)
        nTotal = nTotal + vaArray(n, 7)
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
      If MsgBox("Data akan dihapus?", vbYesNo + vbCritical, "Hapus Data") = vbYes Then
        objData.Start GetDSN
        
        Dim cSQL As String
        cSQL = ""
'        cSQL = " select distinct(nomorpelunasanhutang) as nomorpelunasanhutang from pelunasanhutang where nomorpembelian = '" & cFaktur.Text & "'"
'        Set db = objData.SQL(GetDSN, cSQL)
'        If Not db.EOF Then
'
'          If MsgBox("Transaksi ini sudah pernah dilunasi sebelumnya!" & vbCrLf & "Dengan menghapus berarti seluruh data pelunasan yg berkenaan dengan transaksi ini akan ikut terhapus juga" & vbCrLf & "Apakah anda yakin akan menghapus?", vbYesNo) = vbYes Then
'            lSave = IIf(lSave, objData.Delete(GetDSN, "totpelunasanhutang", "nomorpelunasanhutang", sisAssign, GetNull(db!nomorpelunasanhutang)), False)
'            lSave = IIf(lSave, objData.Delete(GetDSN, "pelunasanhutang", "nomorpelunasanhutang", sisAssign, GetNull(db!nomorpelunasanhutang)), False)
'            lSave = IIf(lSave, objData.Delete(GetDSN, "kartuHutang", "nomorkartuHutang", sisAssign, GetNull(db!nomorpelunasanhutang)), False)
'            lSave = IIf(lSave, objData.Delete(GetDSN, "bukubesar", "faktur", sisAssign, GetNull(db!nomorpelunasanhutang)), False)
'          Else
'            MsgBox "Data tidak bisa dihapus. Penghapusan dibatalkan", vbCritical, "Data "
'            GetEdit False
'            initvalue
'            Exit Sub
'          End If
'        End If
        
        lSave = IIf(lSave, DelKodeTr(objData, vbTrigger.msRefund, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "refund", "nomorrefund", sisAssign, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "totrefund", "nomorrefund", sisAssign, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "kartustock", "nomor", sisAssign, cFaktur.Text), False)
        For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
          'kembalikan harga beli ke semula = hargabeli sekarang + refund (per item)
          lSave = IIf(lSave, objData.Edit(GetDSN, "stock", "kodestock = '" & vaArray(n, 8) & "'", Array("hargabeli"), Array(vaArray(n, 5) + vaArray(n, 6))), False)
        Next n
        If lSave Then
          objData.Save GetDSN
          
          lSave = True
          objData.Start GetDSN
          
          'LAKUKAN UPDATE HARGA POKOK UNTUK MASING MASING PRODUK
          For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
            lSave = IIf(lSave, objData.Edit(GetDSN, "stock", "kodestock = '" & vaArray(n, 8) & "'", Array("cogs"), Array(NewUpdHargaPokok(objData, vaArray(n, 8)))), False)
          Next n
          'NewUpdHargaPokok objData, vaArray(n, 8)
          If lSave Then
            objData.Save GetDSN
          Else
            objData.Cancel GetDSN
          End If
        
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
  nJumlahTotal.value = nTotal
End Sub

Private Sub cFaktur_Validate(Cancel As Boolean)
  If nPos = Edit Then
      cFaktur.Enabled = False
      cFaktur.BackColor = vbButtonFace
  End If
End Sub

Private Sub GetDataStock()
  cBarcode.Text = GetNull(dbData!barcode, "")
  cKode = GetNull(dbData!KodeStock, "")
  cNama.Text = GetNull(dbData!nama, "")
  cSatuan.Text = GetNull(dbData!kodesatuan, "")
  nHarga.value = GetNull(dbData!hargabeli, 0)
  
End Sub

Private Function GetReplaceDataMySQL(cData) As Double
  GetReplaceDataMySQL = Replace(cData, ",", "")
  GetReplaceDataMySQL = Replace(cData, ".", "")
End Function

Private Sub cmdAdd_Click()
  nPos = Add
  GetEdit True
  GetFakturBrowse False
  cFaktur.Text = CreateNomorFaktur(objData, sisModulTransaksi.refund, "totrefund", "nomorrefund")
End Sub

Private Sub cmdAktivasi_Click()
  frmAktivasi.Action Me
End Sub

Private Sub cmdEdit_Click()
  If aCfg(objData, msBisaEditPembelian) = "T" Then
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
  Frame3.Enabled = lPar
  Frame2.Enabled = lPar
  
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
  If aCfg(objData, msBisaEditPembelian) = "T" Then
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
  End If
  If nJumlah.value <= 0 Then
    MsgBox "Nilai jumlah tidak valid", vbExclamation
    nQty.SetFocus
    validOK = False
    Exit Function
  End If
  If Trim(cAkunKas.Text) = "" Or Trim(cSupplier.Text) = "" Then
    MsgBox "Masukkan kode supplier dan Akun Kas", vbExclamation
    validOK = False
    Exit Function
  End If
End Function

Private Sub cmdOK_Click()
Dim n As Integer
Dim nJumlah1 As Double
Dim nTotal As Double

  nTotal = 0
  If validOK() Then
    
    
    'jika baris <= Nomor
    If vaArray.UpperBound(1) + 2 <= nNomor.value Then
      vaArray.ReDim 0, nNomor.value - 1, 0, 10
      n = vaArray.UpperBound(1)
    'jika baris=0
    ElseIf vaArray.UpperBound(1) = -1 Then
      nNomor.value = 1
      vaArray.ReDim 0, nNomor.value - 1, 0, 10
      n = vaArray.UpperBound(1)
    Else
      n = nNomor.value - 1
    End If
        
    vaArray(n, 0) = nNomor.value
    vaArray(n, 1) = cBarcode.Text
    vaArray(n, 2) = GetNamaBarang(cKode)
    vaArray(n, 3) = nQty.value
    vaArray(n, 4) = cSatuan.Text
    vaArray(n, 5) = nHarga.value
    vaArray(n, 6) = nRefund.value
    vaArray(n, 7) = nJumlah.value
    vaArray(n, 8) = cKode
    vaArray(n, 9) = cID
    vaArray(n, 10) = 0 'Untuk menampung selisih harga pokok
    nTotal = nTotal + vaArray(n, 7)
    
    TDBGrid1.Array = vaArray
    TDBGrid1.ReBind
    TDBGrid1.MoveNext
    
'    nJumlah1 = 0
'    nQtyTmp = 0
'    For n = 0 To vaArray.UpperBound(1)
'      nJumlah1 = nJumlah1 + vaArray(n, 7)
'      nQtyTmp = nQtyTmp + vaArray(n, 3)
'    Next
'    TDBGrid1.Columns(3).FooterText = Format(nQtyTmp, "###,###,##0")
'    TDBGrid1.Columns(7).FooterText = Format(nJumlah1, "###,###,##0")
    SumTotal
    
    InitValue1
    
    nNomor.value = vaArray.UpperBound(1) + 2
    cBarcode.SetFocus
  End If
  
End Sub
Private Function GetNamaBarang(cKode As String) As String
Dim db As New ADODB.Recordset
  
  Set db = objData.Browse(GetDSN, "stock", "nama", "kodestock", sisAssign, cKode)
  If Not db.EOF Then
    GetNamaBarang = GetNull(db!nama)
  End If
  
End Function

Private Sub SumTotal()
Dim n As Double
Dim nJumlah As Double
Dim nQtyTmp As Double

    'Looping isi tabel
    For n = 0 To vaArray.UpperBound(1)
      nQtyTmp = nQtyTmp + vaArray(n, 3)
      nJumlah = nJumlah + vaArray(n, 7)
    Next
    'Jumlahkan data kolom yg diperlukan
    nJumlahTotal.value = nJumlah
    
    TDBGrid1.Columns(3).FooterText = Format(nQtyTmp, "###,###,##0")
    TDBGrid1.Columns(7).FooterText = Format(nJumlah, "###,###,##0")
End Sub

Private Function ValidSaving() As Boolean
Dim n As Integer

  ValidSaving = True
  
  If vaArray.UpperBound(1) < 0 Then
    MsgBox "Nota kosong, data tidak disimpan", vbExclamation
    ValidSaving = False
    Exit Function
  End If
  
  If Trim(cFaktur.Text) = "" Then
     MsgBox "Maaf Nomor Faktur Kosong/Tidak Valid" & vbCrLf & "Data tidak bisa disimpan", vbExclamation
     ValidSaving = False
     Exit Function
  End If
  
  If cSupplier.Text = "" Then
    MsgBox "Kode Supplier tidak terisi" & vbCrLf & "Data tidak bisa disimpan", vbExclamation
    ValidSaving = False
    Exit Function
  End If
  
  If Not GetValidDataBrowse(objData, "supplier", "kodesupplier", cSupplier.Text) Then
    MsgBox "Maaf, data supplier tidak tertera dengan benar" & vbCrLf & "Data tidak bisa disimpan", vbExclamation
    ValidSaving = False
    Exit Function
  End If
  
  If Trim(cAkunKas.Text) = "" Then
    MsgBox "Maaf. Akun Kas Belum Diisi" & vbCrLf & "Data tidak bisa disimpan", vbExclamation
    ValidSaving = False
    Exit Function
  End If
  
'  'Jika kode gudang tidak valid, maka penyimpanan data tidak diijinkan
'  Set dbData = objData.Browse(GetDSN, "gudang", "lstatus", "kodegudang", sisAssign, cGudang.Text)
'  If Not dbData.EOF Then
'    If GetNull(dbData!lStatus) <> "A" Then
'      MsgBox "Kode Gudang tidak valid, atau tidak aktif, Data tidak bisa disimpan", vbExclamation
'      ValidSaving = False
'      Exit Function
'    End If
'  End If

End Function

Private Sub cmdSimpan_Click()
 Dim lSave As Boolean
Dim Faktur As String
Dim n As Integer
Dim vaField, vaValue

lSave = True


  
  If ValidSaving Then
    GetNotifikasiAdd "Menyimpan Refund"
    objData.Start GetDSN
    Faktur = cFaktur.Text
    
    vaField = Array("nomorrefund", "fakturasli", "tgl", "jthtmp", "kodesupplier", _
                    "ppn", "persdisc", "persdisc2", "subtotal", "pajak", _
                    "discount", "discount2", "total", "tunai", "hutang", _
                    "datetime", "username", "kodeakun", "kodecostcenter", "kodesalesman", _
                    "statuspembelian", "kodegudang", "kodegroupsales")
    vaValue = Array(Faktur, "", Format(dTgl.value, "yyyy-MM-dd"), Format("2020-06-13", "yyyy-MM-dd"), cSupplier.Text, _
                    0, 0, 0, TDBGrid1.Columns(7).value, 0, _
                    0, 0, TDBGrid1.Columns(7).value, TDBGrid1.Columns(7).value, 0, _
                    SNow, GetRegistry(reg_Username), cAkunKas.Text, GetCostCenterUser(objData, GetRegistry(reg_Username)), "", _
                    cStatusPembelian, aCfg(objData, msGudangPembelian), GetRegistry(reg_KodeGroupSalesPembelian))
    
    lSave = IIf(lSave, objData.Update(GetDSN, "totrefund", "nomorrefund = '" & Faktur & "'", vaField, vaValue), False)
    lSave = IIf(lSave, objData.Delete(GetDSN, "refund", "nomorrefund", sisAssign, Faktur), False)
    lSave = IIf(lSave, objData.Delete(GetDSN, "kartustock", "nomor", sisAssign, Faktur), False)
    
    For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
    
      '    vaArray(n, 0) = nNomor.Value
      '    vaArray(n, 1) = cBarcode.Text
      '    vaArray(n, 2) = cNama.Text
      '    vaArray(n, 3) = nQty.Value
      '    vaArray(n, 4) = cSatuan.Text
      '    vaArray(n, 5) = nHarga.Value
      '    vaArray(n, 6) = nRefund.Value
      '    vaArray(n, 7) = nJumlah.Value
      '    vaArray(n, 8) = cKode
      '    vaArray(n, 9) = cID
      
      vaField = Array("nomorrefund", "kodegudang", "tgl", _
                      "kodestock", "qty", "refund", _
                      "kodesatuan", "jumlah")
      vaValue = Array(Faktur, aCfg(objData, msGudangPembelian), Format(dTgl.value, "yyyy-MM-dd"), _
                      vaArray(n, 8), vaArray(n, 3), vaArray(n, 6), _
                      vaArray(n, 4), vaArray(n, 7))
      
      lSave = IIf(lSave, objData.Add(GetDSN, "refund", vaField, vaValue), False)
      
      '***PENTING***
      'UPDATE KARTU STOCK
      'Cek nilai persediaan terlebih dahulu
      'Jika nilai persediaan minus, gunakan HPP baru dan jika tidak gunakan Harga beli untuk menambah nilai persediaan
      '------------------------------------------------------------------------
      If GetSaldoStock(objData, "", vaArray(n, 8)) < 0 Then
        'vaArray(n, 5) = NewUpdHargaPokok(objData, vaArray(n, 8))
        vaArray(n, 10) = vaArray(n, 7)
        
        '***PENTING***
        vaArray(n, 7) = NewUpdHargaPokok(objData, vaArray(n, 8)) * vaArray(n, 3)
        vaArray(n, 10) = vaArray(n, 10) - vaArray(n, 7) 'Temukan selisih cogs akibat stok mines
        
        lSave = IIf(lSave, UpdKartuStock(objData, SisKartuStock.refund, Faktur, Format(dTgl.value, "yyyy-MM-dd"), vaArray(n, 8), vaArray(n, 3), vaArray(n, 6), 0, "Refund " & cNamaSupplier.Text & " Gudang " & "", aCfg(objData, msGudangPembelian), NewUpdHargaPokok(objData, vaArray(n, 8)), "0"), False)
        'Update FIELD cogs di TABLE stock Dengan yg terakhir
        lSave = IIf(lSave, objData.Edit(GetDSN, "stock", "kodestock = '" & vaArray(n, 8) & "'", Array("cogs", "hargabeli"), Array(NewUpdHargaPokok(objData, vaArray(n, 8)), vaArray(n, 5) - vaArray(n, 6))), False)
      Else
        '***PENTING***
        'Update TABLE kartustock set FIELD ljenis = 0 JANGAN LUPA
        
        lSave = IIf(lSave, UpdKartuStock(objData, SisKartuStock.refund, Faktur, Format(dTgl.value, "yyyy-MM-dd"), vaArray(n, 8), vaArray(n, 3), vaArray(n, 6), 0, "Refund " & cNamaSupplier.Text & " Gudang " & "", aCfg(objData, msGudangPembelian), vaArray(n, 6) - (vaArray(n, 6) * 0 / 100), "0"), False)
        'Update FIELD cogs di TABLE stock Dengan yg terakhir
        lSave = IIf(lSave, objData.Edit(GetDSN, "stock", "kodestock = '" & vaArray(n, 8) & "'", Array("cogs", "hargabeli"), Array(NewUpdHargaPokok(objData, vaArray(n, 8)), vaArray(n, 5) - vaArray(n, 6))), False)
      End If
    Next n
    
    lSave = IIf(lSave, UpdKartuHutang(objData, Sispembelian, Faktur, Format(dTgl.value, "yyyy-MM-dd"), cSupplier.Text, "Refund Non Tunai an. " & cNamaSupplier.Text, 0, SNow, GetRegistry(reg_Username)), False)
    
    ' Inventory (1)
    ' Purchase Tax (2)
    ' Non Inventory Expenses (5)
    '    Acc Payable (2)
    '    Cash Bank (1)
    
    'Posting inventory
    'Hapus dulu di bukubesar
    
    lSave = IIf(lSave, DelKodeTr(objData, vbTrigger.msRefund, Faktur), False)
    'Debet
    For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
'      Dim db As New ADODB.Recordset
'
'      Set db = objData.Browse(GetDSN, "stock", "kodestock,asbiaya", "kodestock", sisAssign, vaArray(n, 8))
'      If Not db.EOF Then
'
'        'lSave = IIf(lSave, UpdKodeTr(objData, vbTrigger.msRefund, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), GetAkunInventory(objData, vaArray(n, 8)), GetCostCenterUser(objData, GetRegistry(reg_Username)), "Pembelian Inventory an " & vaArray(n, 2), vaArray(n, 7), 0, "", SNow, vaArray(n, 8)), False)
'        If vaArray(n, 10) <> 0 Then
'          lSave = IIf(lSave, UpdKodeTr(objData, vbTrigger.msRefund, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningCOGS), GetCostCenterUser(objData, GetRegistry(reg_Username)), "Refund : " & vaArray(n, 2), vaArray(n, 10), 0, "", SNow, vaArray(n, 8)), False)
'        End If
'
''        If GetNull(db!asbiaya) = "1" Then
''          lSave = IIf(lSave, UpdKodeTr(objData, vbTrigger.msRefund, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningBiayaBarang), GetCostCenterUser(objData, GetRegistry(reg_Username)), "Pembelian Inventory an " & vaArray(n, 2), vaArray(n, 7), 0, "", SNow, vaArray(n, 8)), False)
''        Else
'
'        End If
'      End If
      
      'Persediaan (-)
      lSave = IIf(lSave, UpdKodeTr(objData, vbTrigger.msRefund, Faktur, Format(dTgl.value, "yyyy-MM-dd"), GetAkunInventory(objData, vaArray(n, 8)), GetCostCenterUser(objData, GetRegistry(reg_Username)), "Refund " & vaArray(n, 2), 0, vaArray(n, 7), "", SNow, vaArray(n, 8)), False)
    Next n
    
    'Kas (+)
    lSave = IIf(lSave, UpdKodeTr(objData, vbTrigger.msRefund, Faktur, Format(dTgl.value, "yyyy-MM-dd"), cAkunKas.Text, GetCostCenterUser(objData, GetRegistry(reg_Username)), "Terima Kas Refund : " & cNamaSupplier.Text, nJumlahTotal.value, 0, "", SNow), False)
    
    If lSave Then
      objData.Save GetDSN
      UpdateHargaPokokStockTrPembelian vaArray
      GetNotifikasiRemove
    Else
      MsgBox "Data tidak berhasil disimpan", vbCritical
      objData.Cancel GetDSN
    End If
    
    initvalue
    GetEdit False
    
  End If
End Sub

Private Sub UpdateHargaPokokStockTrPembelian(ByVal vaArrayHP As XArrayDB)
Dim n As Single

'    vaArray(n, 0) = nNomor.Value
'    vaArray(n, 1) = cBarcode.Text
'    vaArray(n, 2) = cNama.Text
'    vaArray(n, 3) = nQty.Value
'    vaArray(n, 4) = cSatuan.Text
'    vaArray(n, 5) = nHarga.Value
'    vaArray(n, 6) = nRefund.Value
'    vaArray(n, 7) = nJumlah.Value
'    vaArray(n, 8) = cKode
'    vaArray(n, 9) = cID

  'update harga pokok pada tabel stock untuk masing masing barang
  For n = vaArrayHP.LowerBound(1) To vaArrayHP.UpperBound(1)
      objData.Edit GetDSN, "stock", "kodestock = '" & vaArray(n, 9) & "'", Array("cogs"), Array(NewUpdHargaPokok(objData, vaArray(n, 9)))
  Next n
End Sub

Private Sub cNama_ButtonClick()
' Tampilkan data stock yg Real/Bukan Dummy atau Non Inventory dan statusnya masih aktif

  Set dbData = objData.Browse(GetDSN, "stock s", "s.Barcode,s.nama,s.hargabeli,s.kodesatuan,s.hargajual,s.kodestock", "s.nama", sisContent, cNama.Text, " AND s.jenis < 9 and s.statusnonaktif <> 1")
  If Not dbData.EOF Then
    cNama.Text = cNama.Browse(dbData, Array("BARCODE", "NAMA", "BELI", "SATUAN"), , Array(10, 35, 10, 8))
    GetDataStock
    SumJumlah
  Else
    cNama.Default
  End If
End Sub

Private Sub cNamaAkunKas_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", "kodeakun,keterangan", "jenis", sisAssign, "D", " and keterangan like '%" & cNamaAkunKas.Text & "%'", "kodeakun")
  If Not dbData.EOF Then
    cNamaAkunKas.Text = cNamaAkunKas.Browse(dbData)
    cAkunKas.Text = GetNull(dbData!kodeakun)
    cNamaAkunKas.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub cNamaSupplier_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "supplier", "kodesupplier,nama,alamat,kota", "nama", sisContent, cNamaSupplier.Text, " or kodesupplier like '%" & cNamaSupplier.Text & "%'", "kodesupplier,nama")
  If Not dbData.EOF Then
    cNamaSupplier.Text = cNamaSupplier.Browse(dbData)
    cSupplier.Text = GetNull(dbData!kodesupplier)
    cNamaSupplier.Text = GetNull(dbData!nama, "")
    cAlamat.Text = GetNull(dbData!alamat, "")
    cKota.Text = GetNull(dbData!kota, "")
  End If
End Sub

Private Sub cNamaSupplier_Validate(Cancel As Boolean)
  cNamaSupplier.Enabled = False
End Sub

Private Sub dTgl_Validate(Cancel As Boolean)
  If Not IsInPeriod(dTgl.value) Or (dTgl.value > Date) Then
    Cancel = True
    dTgl.SetFocus
    'GetEdit False
  End If
End Sub

Private Sub Form_Activate()
  Me.Caption = "REFUND * " & GetRegistry(reg_KodeGroupSalesPembelian)
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  GetEdit False
  initvalue
  
  
  Set dbData = objData.Browse(GetDSN, "costcenter", , "kodecostcenter", sisAssign, GetCostCenterUser(objData, GetRegistry(reg_Username)))
  If Not dbData.EOF Then
    Frame3.Caption = "Cost Centre : " & GetNull(dbData!keterangan)
  End If
  
  TabIndex dTgl, n
  TabIndex cSupplier, n
  TabIndex cNamaSupplier, n
  TabIndex cFaktur, n
  TabIndex cAlamat, n
  
  TabIndex nNomor, n
  TabIndex cBarcode, n
  TabIndex cNama, n
  TabIndex nQty, n
  TabIndex cSatuan, n
  TabIndex nHarga, n
  TabIndex nRefund, n
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
Dim dbgudang As New ADODB.Recordset

  
  cStatusPembelian = 0
  cFaktur.Default
  dTgl.value = Date
  cSupplier.Default
  cNamaSupplier.Default
  cAlamat.Default
  cKota.Default
  vaArray.ReDim 0, -1, 0, 10
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
  InitValue1
  vaDelete.ReDim 0, -1, 0, 1
  TDBGrid1.Columns(3).FooterText = ""
  nQty.Decimals = aCfg(objData, msNilaiDecimals)
  cAkunKas.Text = GetAkunKas(objData, GetRegistry(reg_Username))
  nJumlahTotal.value = 0
  cAkunKas.Default
  cNamaAkunKas.Default
  cNamaSupplier.Enabled = True
End Sub

Private Sub InitValue1()
  nNomor.value = 1
  cBarcode.Default
  cNama.Default
  nQty.value = 1
  cSatuan.Default
  nHarga.value = 0
  nRefund.value = 0
  nJumlah.value = 0
  cKode = ""
End Sub

Private Sub nBiaya_Validate(Cancel As Boolean)
  SumTotal
End Sub

Private Sub SumJumlah()
Dim nSubJumlah As Double

  nSubJumlah = nRefund.value * nQty.value
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
      nRefund.value = vaArray(n, 6)
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

Private Sub nRefund_Validate(Cancel As Boolean)
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
      SumTotal
    End If
  End If
End Sub


