VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trKomplimen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compliment"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14130
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   14130
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   2100
      Left            =   15
      Top             =   75
      Width           =   14085
      _ExtentX        =   24844
      _ExtentY        =   3704
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
      Begin BiSADateProject.BiSADate dTgl 
         Height          =   330
         Left            =   450
         TabIndex        =   1
         Top             =   240
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
         Left            =   465
         TabIndex        =   2
         Top             =   600
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
      Begin BiSATextBoxProject.BiSATextBox cKeterangan 
         Height          =   330
         Left            =   465
         TabIndex        =   3
         Top             =   960
         Width           =   6600
         _ExtentX        =   11642
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
         BackColor       =   16777215
         MaxLength       =   100
         Appearance      =   0
         Caption         =   "Keterangan"
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
      Begin BiSATextBoxProject.BiSABrowse cNamaAkun 
         Height          =   330
         Left            =   3825
         TabIndex        =   20
         Top             =   1335
         Width           =   4245
         _ExtentX        =   7488
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
      Begin BiSATextBoxProject.BiSATextBox cKodeAkun 
         Height          =   330
         Left            =   480
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1335
         Width           =   3315
         _ExtentX        =   5847
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
         BackColor       =   -2147483633
         Enabled         =   0   'False
         MaxLength       =   20
         Appearance      =   0
         Caption         =   "Rek. Biaya"
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
         Left            =   7275
         TabIndex        =   0
         Top             =   105
         Width           =   4950
      End
   End
   Begin BiSAFramProject.BiSAFrame BisaFrame2 
      Height          =   4590
      Left            =   15
      Top             =   2190
      Width           =   14100
      _ExtentX        =   24871
      _ExtentY        =   8096
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
      Begin BiSANumberBoxProject.BiSANumberBox nQty 
         Height          =   330
         Left            =   7170
         TabIndex        =   7
         Top             =   135
         Width           =   1110
         _ExtentX        =   1958
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
         Left            =   9540
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   135
         Width           =   1635
         _ExtentX        =   2884
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
      Begin BiSATextBoxProject.BiSATextBox cSatuan 
         Height          =   330
         Left            =   8310
         TabIndex        =   8
         Top             =   135
         Width           =   1200
         _ExtentX        =   2117
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
         Left            =   2670
         TabIndex        =   6
         Top             =   135
         Width           =   4470
         _ExtentX        =   7885
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
         Left            =   720
         TabIndex        =   5
         Top             =   135
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
         Left            =   135
         TabIndex        =   4
         Top             =   135
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
         Height          =   3495
         Left            =   150
         TabIndex        =   12
         Top             =   495
         Width           =   13875
         _ExtentX        =   24474
         _ExtentY        =   6165
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
         Columns(6).Caption=   "JUMLAH"
         Columns(6).DataField=   ""
         Columns(6).NumberFormat=   "###,###,###,###,##0.00"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
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
         Splits(0)._ColumnProps(11)=   "Column(2).Width=7911"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=7832"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=512"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=2037"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1958"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=197122"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2170"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2090"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=513"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=2884"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2805"
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
         _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.alignment=1"
         _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
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
         _StyleDefs(75)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(76)  =   "Named:id=38:HighlightRow"
         _StyleDefs(77)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(78)  =   "Named:id=39:EvenRow"
         _StyleDefs(79)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(80)  =   "Named:id=40:OddRow"
         _StyleDefs(81)  =   ":id=40,.parent=33"
         _StyleDefs(82)  =   "Named:id=41:RecordSelector"
         _StyleDefs(83)  =   ":id=41,.parent=34"
         _StyleDefs(84)  =   "Named:id=42:FilterBar"
         _StyleDefs(85)  =   ":id=42,.parent=33"
      End
      Begin BiSANumberBoxProject.BiSANumberBox nJumlah 
         Height          =   330
         Left            =   11205
         TabIndex        =   10
         Top             =   135
         Width           =   1920
         _ExtentX        =   3387
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
         Left            =   13155
         TabIndex        =   11
         Top             =   135
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
         Picture         =   "frmCompliment.frx":0000
      End
      Begin BiSANumberBoxProject.BiSANumberBox nTotal 
         Height          =   315
         Left            =   10170
         TabIndex        =   13
         Top             =   4125
         Width           =   3855
         _ExtentX        =   6800
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
         BackColor       =   -2147483633
         Caption         =   "Total"
         CaptionWidth    =   900
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
      Begin BiSAButtonProject.BiSAButton Command1 
         Height          =   330
         Left            =   13590
         TabIndex        =   22
         Top             =   135
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   582
         Caption         =   "-"
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
         Picture         =   "frmCompliment.frx":059A
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   675
      Left            =   0
      Top             =   6795
      Width           =   14100
      _ExtentX        =   24871
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
      BorderStyle     =   4
      BackColor       =   -2147483633
      Begin BiSAButtonProject.BiSAButton cmdHapus 
         Height          =   435
         Left            =   2220
         TabIndex        =   14
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
         Picture         =   "frmCompliment.frx":0B34
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   3390
         TabIndex        =   15
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
         Picture         =   "frmCompliment.frx":0DBE
      End
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   1170
         TabIndex        =   16
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
         Picture         =   "frmCompliment.frx":0F5D
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   435
         Left            =   90
         TabIndex        =   17
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
         Picture         =   "frmCompliment.frx":1089
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   12900
         TabIndex        =   18
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
         Picture         =   "frmCompliment.frx":1234
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   11820
         TabIndex        =   19
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
         Picture         =   "frmCompliment.frx":12DA
      End
   End
End
Attribute VB_Name = "trKomplimen"
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
Dim objMenu As New CodeSuiteLibrary.Menu
Dim cAkunKasUser As String

Private Sub GetFakturBrowse(ByVal lStat As Boolean)
  cFaktur.Button = lStat
End Sub

Private Sub cBarcode_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "stock s", "s.kodestock,s.barcode,s.nama,s.kodesatuan,s.hargabeli", "s.barcode", sisContent, cBarcode.Text, " AND s.jenis < 9 and s.statusnonaktif <> 1 and s.konsi = '0'", , , 0, 10)
  If Not dbData.EOF Then
    cBarcode.Text = cBarcode.Browse(dbData)
    GetDataStock
  Else
    cBarcode.Default
  End If
End Sub

Private Sub cBarcode_LostFocus()
Dim kdestock As String

  If Trim(cBarcode.Text) <> "" Then
    Set dbData = objData.Browse(GetDSN, "stock s", "s.kodestock,s.barcode,s.nama,s.kodesatuan,s.hargajual,s.jenis,s.diskonpenjualan,s.bv,s.stok,s.kategori,s.hargabeli", "s.barcode", sisAssign, cBarcode.Text, " AND s.statusnonaktif <> 1 and s.konsi = '0'")
    If Not dbData.EOF Then
      'cBarcode.Text = cBarcode.Browse(dbData)
      kdestock = GetNull(dbData!KodeStock)
      GetDataStock
      SumJumlah
      cmdOK_Click
    Else
      MsgBox "Maaf data yg dicari tidak ada", vbCritical
      cBarcode.Default
      cBarcode.SetFocus
    End If
  End If
End Sub

Private Sub cFaktur_ButtonClick()
Dim db As New ADODB.Recordset
Dim n As Integer
Dim lSave As Boolean
Dim cFilterUsername As String


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

  cFilterUsername = ""
  If objMenu.UserLevel <> 0 Then
    cFilterUsername = " and username = '" & GetRegistry(reg_Username) & "'"
  End If

  Set db = objData.Browse(GetDSN, "totkomplimen", "nomorkomplimen,tgl,total,keterangan", "nomorkomplimen", sisContent, cFaktur.Text, " and tgl = '" & Format(dTgl.value, "yyyy-MM-dd") & "' " & cFilterUsername)
  If Not db.EOF Then
    cFaktur.Text = cFaktur.Browse(db)
    'ambil nilai total
    Set db = objData.Browse(GetDSN, "totkomplimen", , "nomorkomplimen", sisAssign, cFaktur.Text)
    If Not db.EOF Then
      
      nTotal.value = GetNull(db!Total, 0)
     
    End If
    'ambil nilai detail
    Set db = objData.Browse(GetDSN, "komplimen p", "s.barcode,p.kodestock,s.Nama,p.qty,p.kodesatuan,p.harga,p.jumlah", "nomorkomplimen", sisAssign, cFaktur.Text, , , Array("Left join stock s on s.kodestock = p.kodestock"))
    If Not db.EOF Then
      vaArray.ReDim 0, -1, 0, 7
      Do While Not db.EOF
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        n = vaArray.UpperBound(1)
        vaArray(n, 0) = n + 1
        vaArray(n, 1) = GetNull(db!barcode)
        vaArray(n, 2) = GetNull(db!nama)
        vaArray(n, 3) = GetNull(db!qty)
        vaArray(n, 4) = GetNull(db!kodesatuan)
        vaArray(n, 5) = GetNull(db!Harga)
        vaArray(n, 6) = GetNull(db!jumlah)
        vaArray(n, 7) = GetNull(db!KodeStock)
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
        lSave = IIf(lSave, DelKodeTr(objData, msReturPembelian, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "totkomplimen", "nomorkomplimen", sisAssign, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "komplimen", "nomorkomplimen", sisAssign, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "kartustock", "nomor", sisAssign, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "bukubesar", "faktur", sisAssign, cFaktur.Text), False)
        If lSave Then
          objData.Save GetDSN
        Else
          objData.Cancel GetDSN
          MsgBox "Data gagal dihapus", vbExclamation
        End If
      End If
      GetEdit False
      initvalue
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
  nHarga.value = GetNull(dbData!hargabeli, 0)
End Sub

Private Sub cmdAdd_Click()
  nPos = Add
  GetEdit True
  GetFakturBrowse False
  cFaktur.Text = CreateNomorFaktur(objData, sisModulTransaksi.Kompliment, "totkomplimen", "nomorkomplimen")
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
Dim nKe As Integer

  validOK = True
  
  If Not GetValidDataBrowse(objData, "stock", "kodestock", cKode) Then
    MsgBox "Barang tersebut tidak ada dalam database"
    cBarcode.SetFocus
    validOK = False
    Exit Function
  End If
  
  If Trim(cKodeAkun.Text) = "" Then
    MsgBox "Maaf, Masukkan atau isi terlebih dahulu Kode Akun", vbInformation
    validOK = False
    Exit Function
  End If
  
  If isInGrid(vaArray, 7, cKode, , nKe) And nNomor.value > vaArray.UpperBound(1) Then
    'MsgBox "Data sudah pernah dimasukkan sebelumnya dan akan dijumlahkan dengan data sebelumnya", vbExclamation
    cBarcode.SetFocus
    validOK = False
    
    'jika barang yg sama diinput 2x dalam waktu bersamaan, maka akan qty akan
    'dijumlahkan dengan yg sebelumnya, baik harga dan diskon akan sesuai dengan data
    'yg diinput terakhir kali
    
    If nNomor.value > nKe + 1 Then
      vaArray(nKe, 3) = vaArray(nKe, 3) + nQty.value
    Else
      vaArray(nKe, 3) = nQty.value
    End If
    vaArray(nKe, 5) = nHarga.value 'GetHargaGrosir(cKode, nHarga.Value, vaArray(nKe, 3))
    vaArray(nKe, 6) = vaArray(nKe, 3) * vaArray(nKe, 5)
    TDBGrid1.Update
    TDBGrid1.Refresh
    InitValue1
    nNomor.value = vaArray.UpperBound(1) + 2
    cBarcode.SetFocus
    SumTotal
  End If
End Function

Private Sub cmdOK_Click()
Dim n As Integer
Dim nJumlah1 As Double

  If validOK() Then
   
    'jika baris <= Nomor
    If vaArray.UpperBound(1) + 2 <= nNomor.value Then
      vaArray.ReDim 0, nNomor.value - 1, 0, 7
      n = vaArray.UpperBound(1)
    'jika baris=0
    ElseIf vaArray.UpperBound(1) = -1 Then
      nNomor.value = 1
      vaArray.ReDim 0, nNomor.value - 1, 0, 7
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
    vaArray(n, 6) = nJumlah.value
    vaArray(n, 7) = cKode
      
    TDBGrid1.Array = vaArray
    TDBGrid1.ReBind
    
    nJumlah1 = 0
    For n = 0 To vaArray.UpperBound(1)
      nJumlah1 = nJumlah1 + vaArray(n, 6)
    Next
    nTotal.value = nJumlah1
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
Dim nT As Double
  
  nTotal.value = 0
  For n = 0 To vaArray.UpperBound(1)
    nT = nT + vaArray(n, 6)
  Next
  nTotal.value = nT
End Sub

Private Sub Simpan()
Dim lSave As Boolean
Dim Faktur As String
Dim n As Integer
Dim cMetodeRetur As SisModelReturPembelian
lSave = True


  If ValidSaving Then
    If MsgBox("Data akan disimpan?", vbInformation + vbYesNo) = vbYes Then
      objData.Start GetDSN
      
      Faktur = cFaktur.Text
      
      
      lSave = IIf(lSave, objData.Update(GetDSN, "totkomplimen", "nomorkomplimen = '" & Faktur & "'", _
              Array("nomorkomplimen", "tgl", "total", "keterangan", "kodecostcenter", "username", "datetime"), _
              Array(Faktur, Format(dTgl.value, "yyyy-MM-dd"), nTotal.value, cKeterangan.Text, lbCostCenter.Caption, GetRegistry(reg_Username), SNow)), False)
      lSave = IIf(lSave, objData.Delete(GetDSN, "komplimen", "nomorkomplimen", sisAssign, Faktur), False)
      lSave = IIf(lSave, objData.Delete(GetDSN, "kartustock", "nomor", sisAssign, Faktur), False)

      For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
        lSave = IIf(lSave, objData.Add(GetDSN, "komplimen", _
              Array("nomorkomplimen", "tgl", "kodestock", "kodesatuan", "qty", "harga", "jumlah"), _
              Array(Faktur, Format(dTgl.value, "yyyy-MM-dd"), vaArray(n, 7), vaArray(n, 4), vaArray(n, 3), vaArray(n, 5), vaArray(n, 6))), False)
        lSave = IIf(lSave, UpdKartuStock(objData, SisKartuStock.Komplimen, Faktur, Format(dTgl.value, "yyyy-MM-dd"), vaArray(n, 7), vaArray(n, 3), vaArray(n, 5), 0, "Komplimen : " & cKeterangan.Text, GetGudangUser(objData, GetRegistry(reg_Username)), vaArray(n, 5)), False)
      Next n



      'JURNALNYA
      'Biaya (D)
      '     Persediaan     (K)
      '     Jika itu jasa/Konsinyasi
      '     Pendapatan Tidak Tertagih (Komplimen)

      'Hapus dulu di bukubesar
      lSave = IIf(lSave, DelKodeTr(objData, msKomplimen, Faktur), False)
      
      
      'Biayakan sesuai kode akun
      lSave = IIf(lSave, UpdKodeTr(objData, msKomplimen, Faktur, Format(dTgl.value, "yyyy-MM-dd"), cKodeAkun.Text, GetCostCenterUser(objData, GetRegistry(reg_Username)), "Komplimen " & cKeterangan.Text, nTotal.value, 0, "", SNow), False)

      'Persediaan     (K)
      For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
        Dim db As New ADODB.Recordset
'        Set db = objData.Browse(GetDSN, "stock", "kodestock,asbiaya", "kodestock", sisAssign, vaArray(n, 7))
'        If Not db.EOF Then
'          If GetNull(db!asbiaya) = "1" Then
'            lSave = IIf(lSave, UpdKodeTr(objData, msKomplimen, Faktur, Format(dTgl.value, "yyyy-MM-dd"), aCfg(objData, msRekeningBiayaBarang), GetCostCenterUser(objData, GetRegistry(reg_Username)), "Komplimen " & cKeterangan.Text, 0, vaArray(n, 3) * vaArray(n, 5), "", SNow), False)
'          Else
'            lSave = IIf(lSave, UpdKodeTr(objData, msKomplimen, Faktur, Format(dTgl.value, "yyyy-MM-dd"), GetAkunInventory(objData, vaArray(n, 7)), GetCostCenterUser(objData, GetRegistry(reg_Username)), "Komplimen " & cKeterangan.Text, 0, vaArray(n, 3) * vaArray(n, 5), "", SNow), False)
'          End If
'        End If

        lSave = IIf(lSave, UpdKodeTr(objData, msKomplimen, Faktur, Format(dTgl.value, "yyyy-MM-dd"), GetAkunInventory(objData, vaArray(n, 7)), GetCostCenterUser(objData, GetRegistry(reg_Username)), "Komplimen " & cKeterangan.Text, 0, vaArray(n, 3) * vaArray(n, 5), "", SNow), False)
      Next n
      
      If lSave Then
        objData.Save GetDSN
      Else
        MsgBox "Maaf data tidak berhasil disimpan", vbExclamation
        objData.Cancel GetDSN
      End If
      initvalue
      GetEdit False
    End If
  End If
End Sub

Private Function ValidSaving() As Boolean
Dim nCekMetode As Boolean

  ValidSaving = True
  
  If vaArray.UpperBound(1) < 0 Or nTotal.value <= 0 Then
    MsgBox "Nota kosong, data tidak disimpan"
    ValidSaving = False
    Exit Function
  End If
  
  If Trim(cKodeAkun.Text) = "" Then
    MsgBox "Kode akun belum terisi", vbInformation
    ValidSaving = False
    Exit Function
  End If
End Function

Private Sub cmdSimpan_Click()
  Simpan
End Sub

Private Sub cNama_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "stock s", "s.Barcode,s.nama,s.kodesatuan,s.hargabeli,s.kodestock", "s.nama", sisContent, cNama.Text, " AND s.jenis < 9 and s.statusnonaktif <> 1 and s.konsi = '0'", , , 0, 10)
  If Not dbData.EOF Then
    cNama.Text = cNama.Browse(dbData, Array("BARCODE", "NAMA", "SATUAN", "BELI", "KODE"), , Array(13, 35, 10, 8))
    GetDataStock
  Else
    cNama.Default
  End If
End Sub

Private Sub cNamaAkun_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akunbiaya b", "a.keterangan,b.kodeakun", "a.keterangan", sisContent, cNamaAkun.Text, " AND jenis = 'D' or b.kodeakun like '%" & cNamaAkun.Text & "%'", , Array("LEFT JOIN akun a ON a.kodeakun = b.kodeakun"))
  If Not dbData.EOF Then
    cNamaAkun.Text = cNamaAkun.Browse(dbData, Array("Keterangan", "Kode Akun"), , Array(25, 15))
    cKodeAkun.Text = GetNull(dbData!kodeakun)
  End If
End Sub

Private Sub cNamaAkun_Validate(Cancel As Boolean)
  If Trim(cNamaAkun.Text) = "" Then
    cKodeAkun.Text = ""
  End If
  cNamaAkun.Enabled = False
  cNamaAkun.BackColor = vbButtonFace
End Sub

Private Sub Command1_Click()
Dim n As Integer
    If vaArray.UpperBound(1) >= 0 Then
      TDBGrid1.Delete
      For n = 0 To vaArray.UpperBound(1)
        vaArray(n, 0) = n + 1
      Next
      nNomor.value = vaArray.UpperBound(1) + 2
      TDBGrid1.ReBind
      SumTotal
    End If
End Sub

Private Sub dTgl_Validate(Cancel As Boolean)
  If Not IsInPeriod(dTgl.value) Or (dTgl.value > Date) Then
    Cancel = True
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  initvalue
  
  lbCostCenter.Caption = GetCostCenterUser(objData, GetRegistry(reg_Username))
  GetEdit False
  
  TabIndex dTgl, n
  TabIndex cFaktur, n
  TabIndex cKeterangan, n
  TabIndex cNamaAkun, n
  TabIndex nNomor, n
  TabIndex cBarcode, n
  TabIndex cNama, n
  TabIndex nQty, n
  TabIndex cSatuan, n
  TabIndex nHarga, n
  TabIndex nJumlah, n
  
  TabIndex cmdOK, n
  TabIndex cmdAdd, n
  TabIndex cmdEdit, n
  TabIndex cmdHapus, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  TabIndex cmdAktivasi, n
  
  cAkunKasUser = GetAkunKas(objData, GetRegistry(reg_Username))
End Sub

Private Sub initvalue()
  dTgl.value = Date
  cFaktur.Default
  cKeterangan.Default
  cKodeAkun.Default
  cNamaAkun.Default
  
  nTotal.value = 0
  cNamaAkun.Enabled = True
  cNamaAkun.BackColor = vbHighlightText
  vaArray.ReDim 0, -1, 0, 7
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
  InitValue1
End Sub

Private Sub InitValue1()
  nNomor.value = 1
  cBarcode.Default
  cNama.Default
  nQty.value = 1
  cSatuan.Default
  nHarga.value = 0
  nJumlah.value = 0
  cKode = ""
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

Private Sub SumJumlah()
Dim nSubJumlah As Double

  nSubJumlah = nHarga.value * nQty.value
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
      nJumlah.value = vaArray(n, 6)
      cKode = vaArray(n, 7)
    End If
  End If
End Sub

Private Sub nQty_Validate(Cancel As Boolean)
  SumJumlah
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim n As Integer

  If KeyCode = vbKeyDelete Then
    If vaArray.UpperBound(1) >= 0 Then
      TDBGrid1.Delete
      For n = 0 To vaArray.UpperBound(1)
        vaArray(n, 0) = n + 1
      Next
      nNomor.value = vaArray.UpperBound(1) + 2
      TDBGrid1.ReBind
      SumTotal
    End If
  End If
End Sub




