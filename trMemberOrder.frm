VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trMemberOrder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Member Order"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   11775
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   585
      Left            =   0
      Top             =   5745
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
         Top             =   75
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
         Picture         =   "trMemberOrder.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   3390
         TabIndex        =   1
         Top             =   75
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
         Picture         =   "trMemberOrder.frx":028A
      End
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   1170
         TabIndex        =   2
         Top             =   75
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
         Picture         =   "trMemberOrder.frx":0429
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   435
         Left            =   105
         TabIndex        =   3
         Top             =   75
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
         Picture         =   "trMemberOrder.frx":0555
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   10635
         TabIndex        =   4
         Top             =   75
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
         Picture         =   "trMemberOrder.frx":0700
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   9555
         TabIndex        =   5
         Top             =   75
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
         Picture         =   "trMemberOrder.frx":07A6
      End
   End
   Begin BiSAFramProject.BiSAFrame BisaFrame2 
      Height          =   3690
      Left            =   0
      Top             =   2085
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   6509
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
         Left            =   6045
         TabIndex        =   6
         Top             =   75
         Width           =   840
         _ExtentX        =   1482
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
         Left            =   7725
         TabIndex        =   7
         Top             =   75
         Width           =   1500
         _ExtentX        =   2646
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
      Begin BiSATextBoxProject.BiSATextBox cSatuan 
         Height          =   330
         Left            =   6870
         TabIndex        =   8
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
         Left            =   2595
         TabIndex        =   9
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
         Left            =   645
         TabIndex        =   10
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
         TabIndex        =   11
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
         Height          =   2880
         Left            =   105
         TabIndex        =   12
         Top             =   420
         Width           =   11625
         _ExtentX        =   20505
         _ExtentY        =   5080
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
         Columns.Count   =   8
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).MarqueeStyle=   3
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
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=514"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=3493"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3413"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=512"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=6085"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=6006"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=512"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=1455"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1376"
         Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=197122"
         Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(25)=   "Column(4).Width=1482"
         Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1402"
         Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=513"
         Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(31)=   "Column(5).Width=2593"
         Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=2514"
         Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=514"
         Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(37)=   "Column(6).Width=1296"
         Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=1217"
         Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=514"
         Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(43)=   "Column(7).Width=3149"
         Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=3069"
         Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=514"
         Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
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
         Left            =   9930
         TabIndex        =   13
         Top             =   75
         Width           =   1380
         _ExtentX        =   2434
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
         Left            =   11325
         TabIndex        =   14
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
         Picture         =   "trMemberOrder.frx":0A2C
      End
      Begin BiSANumberBoxProject.BiSANumberBox nDisc1 
         Height          =   330
         Left            =   9210
         TabIndex        =   15
         Top             =   75
         Width           =   735
         _ExtentX        =   1296
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
      Begin BiSANumberBoxProject.BiSANumberBox nSubTotal 
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   3330
         Width           =   2640
         _ExtentX        =   4657
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
         BackColor       =   -2147483633
         Caption         =   "Total"
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
      Begin BiSANumberBoxProject.BiSANumberBox nDP 
         Height          =   330
         Left            =   2805
         TabIndex        =   27
         Top             =   3315
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
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
         Caption         =   "DP"
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
      Begin BiSANumberBoxProject.BiSANumberBox nSisa 
         Height          =   330
         Left            =   5625
         TabIndex        =   28
         Top             =   3315
         Visible         =   0   'False
         Width           =   2685
         _ExtentX        =   4736
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
         Caption         =   "Sisa"
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
         Left            =   8325
         TabIndex        =   29
         Top             =   3315
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   582
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
         Button          =   -1  'True
         Caption         =   "KAS"
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
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   2115
      Left            =   0
      Top             =   0
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   3731
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
      Begin BiSAButtonProject.BiSAButton BiSAButton1 
         Height          =   390
         Left            =   6765
         TabIndex        =   32
         Top             =   1590
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   688
         Caption         =   "Label1"
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
      Begin VB.OptionButton optOrder 
         Caption         =   "&Promo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2520
         TabIndex        =   31
         Top             =   1830
         Width           =   1125
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "&Reguler"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   1560
         TabIndex        =   30
         Top             =   1830
         Width           =   1125
      End
      Begin VB.TextBox cFakturAsli 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   8235
         MultiLine       =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   615
         Width           =   3285
      End
      Begin BiSATextBoxProject.BiSABrowse cNamaCustomer 
         Height          =   330
         Left            =   3330
         TabIndex        =   16
         Top             =   720
         Width           =   2670
         _ExtentX        =   4710
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
         TabIndex        =   17
         Top             =   1035
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
         TabIndex        =   18
         Top             =   1035
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
         TabIndex        =   19
         Top             =   720
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
         TabIndex        =   20
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
         TabIndex        =   21
         Top             =   1350
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
      Begin BiSATextBoxProject.BiSABrowse cSalesman 
         Height          =   330
         Left            =   3450
         TabIndex        =   22
         Top             =   1350
         Width           =   2550
         _ExtentX        =   4498
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
         Caption         =   "Sales"
         CaptionWidth    =   700
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
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label1 
         Caption         =   "Keterangan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8235
         TabIndex        =   26
         Top             =   315
         Width           =   960
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
         TabIndex        =   23
         Top             =   135
         Width           =   6030
      End
   End
End
Attribute VB_Name = "trMemberOrder"
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
Dim cJenis  As String
Dim nSaldoStock As Double
Dim cTelp As String

Dim Excel As Excel.Application
Dim ExcelWBk As Excel.Workbook
Dim ExcelWS As Excel.Worksheet

Private Sub StartExcel()
  On Error GoTo err:
  Set Excel = GetObject(, "Excel.Application")
  Exit Sub
err:
  Set Excel = CreateObject("Excel.Application")
End Sub

Private Sub CloseWorkSheet()
  ExcelWBk.Close
  Excel.Quit
End Sub

Private Sub FinishExcel()
  'Jangan lupa, selalu bersihkan memory saat mengakhiri
  If Not ExcelWS Is Nothing Then Set ExcelWS = Nothing
  If Not ExcelWBk Is Nothing Then Set ExcelWBk = Nothing
  If Not Excel Is Nothing Then Set Excel = Nothing
End Sub

Private Sub GetFakturBrowse(ByVal lStat As Boolean)
  cFaktur.Button = lStat
End Sub

Private Sub BiSAButton1_Click()
  CommonDialog1.Filter = "Excel File (*.xls)|*.xls"
  CommonDialog1.ShowOpen
  If CommonDialog1.FileName <> "" Then
    GetLoadExcel
  End If
End Sub

Private Sub GetLoadExcel()
Dim lSave As Boolean
Dim vaField, vaValue
Dim i, j, n As Integer
Dim dbData As New ADODB.Recordset
Dim Wb As Excel.Workbook

  On Error GoTo err:
  StartExcel
  lSave = True
  
  Excel.Workbooks.Close
  Set ExcelWBk = Excel.Workbooks.Open(CommonDialog1.FileName)
  Set ExcelWS = ExcelWBk.Worksheets(1)
  
  
  FrmPB.InitPB ExcelWS.Cells.SpecialCells(xlCellTypeLastCell).Row
  Dim cBarcode
  Dim cQty

  For i = 1 To ExcelWS.Cells.SpecialCells(xlCellTypeLastCell).Row
    FrmPB.RunPB
    With ExcelWS
      Set dbData = objData.Browse(GetDSN, "stock", "kodestock,nama,hargabeli,diskonpenjualan,kodesatuan,barcode", "barcode", sisAssign, Trim(.Cells(i, 1).Value))
      If Not dbData.EOF Then
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        n = vaArray.UpperBound(1)
        vaArray(n, 0) = n + 1
        vaArray(n, 1) = .Cells(i, 1).Value
        vaArray(n, 2) = GetNull(dbData!nama)
        vaArray(n, 3) = .Cells(i, 2).Value 'IIf(.Cells(i, 2).Value Is Null Or .Cells(i, 2).Value = "", 1, .Cells(i, 2).Value)
        vaArray(n, 4) = GetNull(dbData!kodesatuan)
        
        vaArray(n, 5) = IIf(Trim(.Cells(i, 3)) = "", GetNull(dbData!hargabeli), GetNull(.Cells(i, 3).Value))
        vaArray(n, 6) = IIf(Trim(.Cells(i, 4)) = "", IIf(GetNull(dbData!diskonpenjualan) = 0, 0, GetNull(dbData!diskonpenjualan) + 3), GetNull(.Cells(i, 4).Value))

        vaArray(n, 7) = (vaArray(n, 5) - (vaArray(n, 5) * vaArray(n, 6) / 100)) * vaArray(n, 3)
                
        vaArray(n, 8) = GetNull(dbData!KodeStock)
'        vaArray(n, 9) = cID
      Else
        'jika data yg di import tidak ada dalam database simpan
        
      End If
    End With
  Next i
  nNomor.Value = vaArray.UpperBound(1) + 2
  Set tdbgrid1.Array = vaArray
  tdbgrid1.ReBind
  tdbgrid1.Refresh
  SumTotal
  FrmPB.EndPB
  CloseWorkSheet
  FinishExcel
  
err:
End Sub

Private Sub cAkunKas_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", , "jenis", sisAssign, "D", " AND LEFT(kodeakun,1) = 1")
  If Not dbData.EOF Then
    cAkunKas.Text = cAkunKas.Browse(dbData)
  End If
End Sub

Private Sub cBarcode_ButtonClick()
Dim kdestock As String

  Set dbData = objData.Browse(GetDSN, "stock s", "s.kodestock,s.barcode,s.nama,s.kodesatuan,s.hargajual,s.jenis,s.diskonpenjualan", "s.barcode", sisContent, cBarcode.Text)
  If Not dbData.EOF Then
    cBarcode.Text = cBarcode.Browse(dbData)
    kdestock = GetNull(dbData!KodeStock)
    GetDataStock
    SumJumlah
    'tampilkan info stock
'    GetInfoStockDong objData, kdestock
  Else
    cBarcode.Default
  End If
End Sub

Private Sub cFaktur_ButtonClick()
Dim db As New ADODB.Recordset
Dim n As Integer
Dim lSave As Boolean

lSave = True
  
  Set db = objData.Browse(GetDSN, "totmemberorder", "nomormemberorder,tgl,subtotal,total,piutang", "nomormemberorder", sisContent, cFaktur.Text, " and tgl = '" & Format(dTgl.Value, "yyyy-MM-dd") & "' and kodeanggota = '" & cCustomer.Text & "' AND status = 0")
  If Not db.EOF Then
    cFaktur.Text = cFaktur.Browse(db)
    'ambil nilai total
    Set db = objData.Browse(GetDSN, "totmemberorder", , "nomormemberorder", sisAssign, cFaktur.Text)
    If Not db.EOF Then
      cFakturAsli.Text = GetNull(db!fakturasli, "")
      nSubTotal.Value = GetNull(db!Subtotal, 0)
      cSalesman.Text = GetNull(db!kodesalesman, "")
      nDP.Value = GetNull(db!dp)
      cAkunKas.Text = GetNull(db!akunkas)
      SetOpt optOrder, GetNull(db!jenisorder)
    End If
    'ambil nilai detail
    Set db = objData.Browse(GetDSN, "memberorder p", "s.barcode,p.kodestock,s.Nama,p.qty,p.kodesatuan,p.harga,p.discount,p.jumlah", "nomormemberorder", sisAssign, cFaktur.Text, , "p.nourut asc", Array("Left join stock s on s.kodestock = p.kodestock"))
    If Not db.EOF Then
      vaArray.ReDim 0, -1, 0, 9
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
      Set tdbgrid1.Array = vaArray
      tdbgrid1.ReBind
      tdbgrid1.Refresh
      Me.Refresh
      nNomor.Value = vaArray.UpperBound(1) + 2
    End If
    
    If nPos = Delete Then
      If Not isInProsesPO(objData, cFaktur.Text) Then
        If MsgBox("Data akan dihapus?", vbYesNo) = vbYes Then
          objData.Start GetDSN
          'rutin menghapus transaksi memberorder
          lSave = IIf(lSave, objData.Delete(GetDSN, "memberorder", "nomormemberorder", sisAssign, cFaktur.Text), False)
          lSave = IIf(lSave, objData.Delete(GetDSN, "totmemberorder", "nomormemberorder", sisAssign, cFaktur.Text), False)
          lSave = IIf(lSave, objData.Delete(GetDSN, "po", "kodeso", sisAssign, cFaktur.Text), False)
          
          'hapus posting di bukubesar
          lSave = IIf(lSave, DelKodeTr(objData, msMemberOrder, cFaktur.Text), False)
          
          If lSave Then
            objData.Save GetDSN
          Else
            objData.Cancel GetDSN
          End If
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

Private Sub cmdAktivasi_Click()
  frmAktivasi.Action Me
End Sub

Private Sub GetDataStock()
Dim db As New ADODB.Recordset

  cBarcode.Text = GetNull(dbData!barcode, "")
  cKode = GetNull(dbData!KodeStock, "")
  cNama.Text = GetNull(dbData!nama, "")
  cSatuan.Text = GetNull(dbData!kodesatuan, "")
  nDisc1.Value = GetNull(dbData!diskonpenjualan)
  
  If aCfg(objData, msCHKdiscountPenjualan) = 1 Then
    nDisc1.Value = aCfg(objData, msDiscountPenjualan)
  End If
  
  'tentukan harga jual sesuai dengan konfigurasi yg telah di setup
  If aCfg(objData, msHargaPenjualanNonTunai) = "3" Then
    nHarga.Value = GetHargaKontrak(objData, cCustomer.Text, cKode)
  ElseIf aCfg(objData, msHargaPenjualanNonTunai) = "2" Then
    nHarga.Value = GetHargaJualLastByCustomer(objData, cKode, cCustomer.Text)
  Else
    nHarga.Value = GetNull(dbData!hargajual)
  End If
  
  'Lakukan markup harga jika non member
  nHarga.Value = MarkUpHarga(objData, cCustomer.Text, nHarga.Value)
  cJenis = GetNull(dbData!jenis)
  
  'jika di master customer tersetup diskon, maka abaikan semuanya
  Set dbData = objData.Browse(GetDSN, "anggota")
  If Not dbData.EOF Then
    If GetNull(dbData!diskon) <> 0 Then
      nDisc1.Value = GetNull(dbData!diskon)
    End If
  End If
  
  'jika di master customer tersetup diskon, maka abaikan semuanya
  Set dbData = objData.Browse(GetDSN, "anggota")
  If Not dbData.EOF Then
    If GetNull(dbData!diskon) <> 0 Then
      nDisc1.Value = GetNull(dbData!diskon)
    End If
  End If
  
End Sub

Private Function GetHargaJualLastByCustomer(ByVal obj As CodeSuiteLibrary.Data, ByVal cStock As String, ByVal cCust As String) As Double
Dim db As New ADODB.Recordset

  Set db = obj.Browse(GetDSN, "memberorder p", "p.tgl,p.kodestock,p.harga", "p.kodestock", sisAssign, cStock, " and t.kodeanggota = '" & cCust & "'", "p.tgl desc", Array("left join totmemberorder t on t.nomormemberorder = p.nomormemberorder"), 0, 1)
  If Not db.EOF Then
    GetHargaJualLastByCustomer = GetNull(db!Harga)
  Else
    Set db = obj.Browse(GetDSN, "stock", , "kodestock", sisAssign, cStock)
    If Not db.EOF Then
      GetHargaJualLastByCustomer = GetNull(db!hargajual)
    End If
  End If
End Function

Private Function GetHargaKontrak(ByVal obj As CodeSuiteLibrary.Data, ByVal cCustomer As String, ByVal cStock As String) As Double
Dim db As New ADODB.Recordset
  
  GetHargaKontrak = 0
  Set db = obj.Browse(GetDSN, "kontrakstock", , "kodeanggota", sisAssign, cCustomer, " and kodestock = '" & cStock & "'")
  If Not db.EOF Then
    GetHargaKontrak = GetNull(db!hargakontrak)
  Else
    Set db = obj.Browse(GetDSN, "stock", , "kodestock", sisAssign, cStock)
    If Not db.EOF Then
      GetHargaKontrak = GetNull(db!hargajual)
    End If
  End If
End Function

Private Function MarkUpHarga(ByVal obj As CodeSuiteLibrary.Data, ByVal anggota As String, ByVal Harga As Double) As Double
Dim db As New ADODB.Recordset
  MarkUpHarga = Harga
  Set db = obj.Browse(GetDSN, "anggota", , "kodeanggota", sisAssign, anggota)
  If Not db.EOF Then
    If GetNull(db!Status) <> "A" Then
      MarkUpHarga = Harga + (aCfg(objData, msMarkUpHargaJual) * Harga / 100)
    End If
  End If
End Function

Private Sub cmdAdd_Click()
  nPos = Add
  GetEdit True
  GetFakturBrowse False
  cFaktur.Text = GetNomor("totmemberorder", "nomormemberorder", GetID, sisModulTransaksi.MemberOrder)
  cAkunKas.Text = GetAkunKas(objData, GetRegistry(reg_UserName))
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
  If Not GetValidDataBrowse(objData, "stock", "kodestock", cKode) Then
'  If Trim(cKode) = "" Then
    MsgBox "Barang tersebut tidak ada dalam database "
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

  
  If aCfg(objData, msIjinkanHargaBeliDibawahHargajual) <> 1 Then
    Set dbData = objData.Browse(GetDSN, "stock", , "kodestock", sisAssign, cKode)
    If Not dbData.EOF Then
      If GetNull(dbData!hargabeli) > nHarga.Value Then
        MsgBox "Stop" & vbCrLf & "Maaf. tidak bisa dilanjutkan." & vbCrLf & "Harga jual tidak sesuai, silahkan hubungi supervisor untuk penjelasan lebih lanjut." & vbCrLf & "Terimaksih"
        If MsgBox("Apakah tetap akan dilanjutkan", vbYesNo) = vbYes Then
          validOK = True
          Exit Function
        Else
          validOK = False
          Exit Function
        End If
      End If
    End If
  End If

End Function

Private Sub cmdOK_Click()
Dim n As Integer
Dim nJumlah1 As Double

  If validOK() Then
   
    'jika baris <= Nomor
    If vaArray.UpperBound(1) + 2 <= nNomor.Value Then
      vaArray.ReDim 0, nNomor.Value - 1, 0, 9
      n = vaArray.UpperBound(1)
    'jika baris=0
    ElseIf vaArray.UpperBound(1) = -1 Then
      nNomor.Value = 1
      vaArray.ReDim 0, nNomor.Value - 1, 0, 9
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
    vaArray(n, 9) = cJenis
      
    tdbgrid1.Array = vaArray
    tdbgrid1.ReBind
    
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
End Sub

Private Function ValidSaving() As Boolean
  ValidSaving = True
End Function

Private Sub cCustomer_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "anggota", "kodeanggota,nama,kodedep,alamat,telp", "kodeanggota", sisContent, cCustomer.Text, , "kodeanggota,nama")
  If Not dbData.EOF Then
    cCustomer.Text = cCustomer.Browse(dbData, Array("KODE", "NAMA", "DEP", "ALAMAT"), , Array(10, 20, 6, 10))
    cCustomer.Text = GetNull(dbData!kodeanggota)
    cNamaCustomer.Text = GetNull(dbData!nama, "")
    cAlamat.Text = GetNull(dbData!alamat, "")
    cKota.Text = GetNull(dbData!kodedep, "")
    cTelp = GetNull(dbData!telp)
  End If
End Sub

Private Sub cmdSimpan_Click()
Dim lSave As Boolean
Dim Faktur As String
Dim n As Integer
Dim i As Integer
Dim nValueTunai As Double
Dim nValueKredit As Double

  lSave = True

  'simpan pada tabel totmemberorder
  'simpan pada tabel memberorder
  'simpan pada tabel kartustock
  'simpan pada tabel kartupiutang
  
  If isValidSaving Then
    objData.Start GetDSN
    Faktur = cFaktur.Text
    If nPos = Add Then
      If Not GetAvailable(cFaktur.Text, "totmemberorder", "nomormemberorder") Then
        Faktur = GetNomor("totmemberorder", "nomormemberorder", GetID, sisModulTransaksi.MemberOrder)
      End If
    End If
    lSave = IIf(lSave, objData.Update(GetDSN, "totmemberorder", "nomormemberorder = '" & Faktur & "'", Array("nomormemberorder", "fakturasli", "tgl", "jthtmp", "kodeanggota", "ppn", "persdisc", "persdisc2", "subtotal", "pajak", "discount", "discount2", "total", "tunai", "piutang", "datetime", "username", "kodeakun", "kodecostcenter", "flaglunas", "kodesalesman", "komisi", "dp", "akunkas", "jenisorder"), Array(Faktur, cFakturAsli.Text, Format(dTgl.Value, "yyyy-MM-dd"), "", cCustomer.Text, 0, 0, 0, nSubTotal.Value, 0, 0, 0, 0, 0, 0, SNow, GetRegistry(reg_UserName), "", aCfg(objData, msCostCenterJualBeli), 0, cSalesman.Text, 0, nDP.Value, cAkunKas.Text, GetOpt(optOrder))), False)
    lSave = IIf(lSave, objData.Delete(GetDSN, "memberorder", "nomormemberorder", sisAssign, Faktur), False)
    lSave = IIf(lSave, objData.Delete(GetDSN, "po", "kodeso", sisAssign, Faktur), False)
    lSave = IIf(lSave, objData.Delete(GetDSN, "kartupiutang", "nomorkartupiutang", sisAssign, Faktur), False)
    'simpan juga di kartupiutang
    lSave = IIf(lSave, UpdKartuHutang(objData, SisMemberOrder, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), cCustomer.Text, "DP Member an. " & cNamaCustomer.Text, nDP.Value, SNow, GetRegistry(reg_UserName)), False)
    
    For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
      nValueTunai = 0
      nValueKredit = 0
      lSave = IIf(lSave, objData.Add(GetDSN, "memberorder", Array("nomormemberorder", "kodegudang", "tgl", "kodestock", "qty", "harga", "kodesatuan", "discount", "jumlah", "hb", "tunai", "piutang", "nourut"), Array(Faktur, aCfg(objData, msGudangPenjualan), Format(dTgl.Value, "yyyy-MM-dd"), vaArray(n, 8), vaArray(n, 3), vaArray(n, 5), vaArray(n, 4), vaArray(n, 6), vaArray(n, 7), GetHargaBeli(objData, vaArray(n, 8)), nValueTunai, nValueKredit, vaArray(n, 0))), False)
      'update ke table po
      lSave = IIf(lSave, objData.Add(GetDSN, "po", Array("kodeso", "kodestock", "qty", "datetime", "username", "harga", "diskonpenjualan"), Array(Faktur, vaArray(n, 8), vaArray(n, 3), SNow, GetRegistry(reg_UserName), vaArray(n, 5), vaArray(n, 6))), False)
      'update status lunas
      If lCekStatusLunas(objData, Faktur) = True Then
        lSave = IIf(lSave, objData.Edit(GetDSN, "memberorder", "nomormemberorder = '" & Faktur & "'", Array("statuslunas"), Array(1)), False)
      Else
        lSave = IIf(lSave, objData.Edit(GetDSN, "memberorder", "nomormemberorder = '" & Faktur & "'", Array("statuslunas"), Array(0)), False)
      End If
    Next n
    
    'Simpan kedalam bukubesar
    'Jurnal
    '========================
    'Kas
    '   Piutang Usaha
    lSave = IIf(lSave, DelKodeTr(objData, msMemberOrder, Faktur), False)
    lSave = IIf(lSave, UpdKodeTr(objData, msMemberOrder, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), cAkunKas.Text, aCfg(objData, msCostCenterJualBeli), "Order/DP an " & cNamaCustomer.Text, nDP.Value), False)
      lSave = IIf(lSave, UpdKodeTr(objData, msMemberOrder, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), GetAkunMember(objData, cCustomer.Text), aCfg(objData, msCostCenterJualBeli), "Order/DP an " & cNamaCustomer.Text, 0, nDP.Value), False)
    
    If lSave Then
      objData.Save GetDSN
    Else
      objData.Cancel GetDSN
    End If
    
'    If MsgBox("Apakah akan mencetak transaksi ke printer?", vbYesNo + vbInformation) = vbYes Then
'      If aCfg(objData, msCetakanPenjualanNonTunai) = 1 Then
'        'memberorder Nota
'        GetCetakFakturMemberOrder objData, Faktur, False
'        Unload frmFaktur
'      Else
'        If GetRegistry(reg_PortPrinterKasir) = "USB" Then
'          'print struk
'          For i = 1 To aCfg(objData, msJumlahCetakanPenjualanNonTunai)
'            trPrint.noOrder = Faktur
'            Set dbData = objData.Browse(GetDSN, "totmemberorder t", "t.*,a.*", "t.nomormemberorder", sisAssign, Faktur, , , Array("left join anggota a on a.kodeanggota = t.kodeanggota"))
'            If Not dbData.EOF Then
'              trPrint.nSubTotal = GetNull(dbData!Subtotal)
'              trPrint.nDiscount = GetNull(dbData!dp)
'              trPrint.cKodeMember = GetNull(dbData!kodeanggota)
'              trPrint.cMember = GetNull(dbData!nama)
'              trPrint.cTeleponMember = GetNull(dbData!telp)
'              Load trPrint
'              trPrint.Show vbModal
'            End If
'          Next i
'        Else
'          PrintStruk Faktur
'        End If
'      End If
'    End If
    
    initvalue
    GetEdit False
  End If
End Sub

Private Function isInProsesPO(ByVal obj As CodeSuiteLibrary.Data, ByVal Faktur As String) As Boolean
isInProsesPO = False

    'Cek jika so ini sudah po, maka data tidak boleh disimpan/dikoreksi/dihapus
    Set dbData = objData.Browse(GetDSN, "po", , "kodeso", sisAssign, Faktur, " and statusorder = 1 ")
    If Not dbData.EOF Then
      'jika no faktur ini ditemukan dalam table po, maka cancel penyimpanan
      MsgBox "Maaf, order dengan nomor " & Faktur & " ini sudah diproses." & vbCrLf & "PENYIMPANAN DIBATALKAN"
      isInProsesPO = True
      Exit Function
    End If
End Function

'Private Sub PrintStruk(ByVal Faktur As String)
'Dim n As Double
'Dim nBruto As Double
'Dim nTotQty As Double
'
'  With aMainmenu.IO1
'    .Open GetRegistry(reg_PortPrinterKasir), ""
'
''    .WriteString Chr(27) & Chr(15) & vbCrLf
'    .WriteString Padc(Trim("STRUK ORDER/PESANAN"), 40) & vbCrLf
'    .WriteString Padc(Trim(aCfg(objData, msNamaPerusahaan)), 40) & vbCrLf
'    .WriteString Padc(Trim(aCfg(objData, msAlamatPerusahaan)), 40) & vbCrLf
'    .WriteString Padc(Trim(aCfg(objData, msTelepon) & " " & aCfg(objData, msFax)), 40) & vbCrLf
'    .WriteString Padc(aCfg(objData, msKota), 40) & vbCrLf
'    .WriteString Replicate("-", 40) & vbCrLf
'    .WriteString "Member. " & cCustomer.Text & "/" & cNamaCustomer.Text & vbCrLf
'    .WriteString "Telp. " & cTelp & vbCrLf
'    .WriteString Replicate("-", 40) & vbCrLf
'
'    For n = 0 To vaArray.UpperBound(1)
'      If vaArray(n, 3) <> 0 Then
'        vaArray(n, 5) = vaArray(n, 5) - (vaArray(n, 5) * vaArray(n, 6) / 100)
'        nBruto = nBruto + (vaArray(n, 3) * vaArray(n, 5))
'        nTotQty = nTotQty + vaArray(n, 3)
'        .WriteString Left(vaArray(n, 2), 40) & vbCrLf
'        .WriteString Padl(Left(vaArray(n, 1), 10) & Padl(Format(vaArray(n, 3), "#,##0.00"), 8) & " x" & Padl(Format(vaArray(n, 5), "#,###,##0"), 9) & " =" & Padl(Format(vaArray(n, 3) * vaArray(n, 5), "#,###,##0"), 9), 40) & vbCrLf
'      End If
'    Next
'    .WriteString Replicate("-", 40) & vbCrLf
'    .WriteString Padr("=> " & Format(nTotQty, "###,###,##0.00") & " Items", 20) & Padl("Sub   : " & Padl(Format(nBruto, "###,###,##0"), 11), 20) & vbCrLf
'    .WriteString Padl("DP    : " & Padl(Format(nDP.Value, "###,###,##0"), 11), 40) & vbCrLf
'    .WriteString Replicate("-", 40) & vbCrLf
'    .WriteString "No. " & Faktur & Padl(Format(Now, "dd-MM-yyyy HH:MM:SS"), 22) & vbCrLf
'
''   .WriteString "No. " & Faktur & Padl(Format(Now, "dd-MM-yyyy HH:MM:SS"), 22) & vbCrLf
''   .WriteString Padr("=> " & Format(nTotQty, "###,###,##0.00") & " Items", 20) & Padl("Sub   : " & Padl(Format(nBruto, "###,###,##0"), 11), 20) & vbCrLf
'
'    .WriteString "Print by " & Padl(GetRegistry(reg_UserName), 26) & vbCrLf
'    .WriteString Replicate("-", 40) & vbCrLf
'    .WriteString Padc(Trim(aCfg(objData, msKasir1, "")), 40) & vbCrLf
'    .WriteString Padc(Trim(aCfg(objData, msKasir2, "")), 40) & vbCrLf
'    .WriteString Padc(Trim(aCfg(objData, msKasir3, "")), 40) & vbCrLf
'    .WriteString Padc(Trim(aCfg(objData, msKasir4, "")), 40) & vbCrLf
'    .WriteString Padc(Trim(aCfg(objData, msKasir5, "")), 40) & vbCrLf
'    .WriteString Padc(Trim(aCfg(objData, msKasir6, "")), 40) & vbCrLf
'    .WriteString Padc(Trim(aCfg(objData, msKasir7, "")), 40) & vbCrLf
'    .WriteString Padc(Trim(aCfg(objData, msKasir7, "")), 40) & vbCrLf
'
'    .Close
'    OpenDrawer GetRegistry(reg_PortPrinterKasir)
'  End With
'End Sub

Private Function isValidSaving() As Boolean
Dim dba As New ADODB.Recordset
isValidSaving = True
  
  'pastikan kolom anggota sudah diisi lengkap
  Set dba = objData.Browse(GetDSN, "anggota", "kodeanggota", "kodeanggota", sisAssign, cCustomer.Text)
  If dba.EOF Then
    If dba.RecordCount = 0 Then
      MsgBox "Maaf, Kode Anggota yang dimasukkan tidak ada dalam database komputer" & vbCrLf & "Data tidak bisa disimpan", vbInformation
      isValidSaving = False
      Exit Function
    End If
  End If
  'cek validitas
  If Not GetValidDataBrowse(objData, "anggota", "kodeanggota", cCustomer.Text) Then
    MsgBox "Kode member tidak tertera dengan benar" & vbCrLf & "Data tidak bisa disimpan"
    cCustomer.SetFocus
    isValidSaving = False
    Exit Function
  End If
  
  If Not GetValidDataBrowse(objData, "salesman", "kodesalesman", cSalesman.Text) Then
    MsgBox "Kode salesman tidak tertera dengan benar" & vbCrLf & "Data tidak bisa disimpan"
    cSalesman.SetFocus
    isValidSaving = False
    Exit Function
  End If
  
  
  If nDP.Value > 0 Then
    If Not GetValidDataBrowse(objData, "akun", "kodeakun", cAkunKas.Text) Then
      MsgBox "Kode akun kas tidak tertera dengan benar" & vbCrLf & "Data tidak bisa disimpan"
      cAkunKas.SetFocus
      isValidSaving = False
      Exit Function
    End If
  End If
  If nPos = Edit Then
    If isInProsesPO(objData, cFaktur.Text) Then
      
      isValidSaving = False
      Exit Function
    End If
  End If
End Function

Private Sub cNama_ButtonClick()
Dim kdestock As String

  Set dbData = objData.Browse(GetDSN, "stock s", "s.kodestock,s.Barcode,s.nama,s.kodesatuan,s.hargajual,s.jenis,s.diskonpenjualan", "s.nama", sisContent, cNama.Text)
  If Not dbData.EOF Then
    cNama.Text = cNama.Browse(dbData)
    kdestock = GetNull(dbData!KodeStock)
    GetDataStock
    SumJumlah
    GetInfoStockDong objData, kdestock
  Else
    cNama.Default
  End If
End Sub

Private Sub cNamaCustomer_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "anggota", "kodeanggota,nama,kodedep,alamat,telp", "nama", sisContent, cNamaCustomer.Text, , "kodeanggota,nama")
  If Not dbData.EOF Then
    cNamaCustomer.Text = cNamaCustomer.Browse(dbData, Array("Kode", "Nama", "Dep", "Alamat"), , Array(6, 15, 6, 15))
    cCustomer.Text = GetNull(dbData!kodeanggota)
    cNamaCustomer.Text = GetNull(dbData!nama, "")
    cAlamat.Text = GetNull(dbData!alamat, "")
    cKota.Text = GetNull(dbData!kodedep, "")
    cTelp = GetNull(dbData!telp)
  End If
End Sub

Private Sub cSalesman_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "salesman", "kodesalesman,nama")
  If Not dbData.EOF Then
    cSalesman.Text = cSalesman.Browse(dbData)
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

  
'  If CheckTrial(nRecordsTrial, TrialPenjualan) = True Then
'    End
'  End If
  
  
  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  initvalue
  GetEdit False
  
  Set dbData = objData.Browse(GetDSN, "costcenter", , "kodecostcenter", sisAssign, aCfg(objData, msCostCenterJualBeli))
  If Not dbData.EOF Then
    lbCostCenter.Caption = "Cost Centre : " & GetNull(dbData!keterangan)
  End If
  
  
  TabIndex dTgl, n
  TabIndex cCustomer, n
  TabIndex cNamaCustomer, n
  TabIndex cAlamat, n
  TabIndex cFaktur, n
  TabIndex cSalesman, n
  TabIndex optOrder(0), n
  TabIndex optOrder(1), n
  
  TabIndex cFakturAsli, n
  
  TabIndex nNomor, n
  TabIndex cBarcode, n
  TabIndex cNama, n
  TabIndex nQty, n
  TabIndex cSatuan, n
  TabIndex nHarga, n
  TabIndex nDisc1, n
  TabIndex nJumlah, n
  TabIndex cmdOK, n
  TabIndex nDP, n
  TabIndex cAkunKas, n
  
  TabIndex cmdAdd, n
  TabIndex cmdEdit, n
  TabIndex cmdHapus, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  TabIndex cmdAktivasi, n
End Sub

Private Sub GetHasil()
  If nDP.Value <= nSubTotal.Value Then
    nSisa.Value = nSubTotal.Value - nDP.Value
  Else
    MsgBox "Maaf, dp untuk nota tidak boleh melebih dari total yg sudah tertera"
    nDP.Default
    nDP.SetFocus
  End If
End Sub


Private Sub initvalue()
  cFaktur.Default
  dTgl.Value = Date
  cSalesman.Default
  cFakturAsli.Text = ""
  cCustomer.Default
  cNamaCustomer.Default
  cAlamat.Default
  cKota.Default
  nSubTotal.Value = 0
  nDP.Default
  nSisa.Default
  cAkunKas.Default
  optOrder(0).Value = True
  
  If GetKunciAkunKas(objData) Then
    cAkunKas.Enabled = False
  End If

  vaArray.ReDim 0, -1, 0, 9
  Set tdbgrid1.Array = vaArray
  tdbgrid1.ReBind
  InitValue1
  If aCfg(objData, msKolomHargaPenjualanNonTunai) = 1 Then
    nHarga.Enabled = True
  Else
    nHarga.Enabled = False
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
  nNomor.Value = 1
  cBarcode.Default
  cNama.Default
  nQty.Value = 1
  cSatuan.Default
  nHarga.Value = 0
  nDisc1.Value = 0 'aCfg(objData, msDiscountItemPembelian)
  nJumlah.Value = 0
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

Private Sub nBiaya_Validate(Cancel As Boolean)
  SumTotal
End Sub

Private Sub SumJumlah()
Dim nSubJumlah As Double

  nSubJumlah = nHarga.Value * nQty.Value
  nSubJumlah = nSubJumlah - (nSubJumlah * (nDisc1.Value / 100))
  nJumlah.Value = nSubJumlah
End Sub

Private Sub nDP_Change()
  GetHasil
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
      cJenis = vaArray(n, 9)
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

Private Sub optOrder_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    keybd_event VK_TAB, 0, 0, 0
  End If
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim n As Integer

  If KeyCode = vbKeyDelete Then
    If vaArray.UpperBound(1) >= 0 Then
      tdbgrid1.Delete
      SumTotal
      For n = 0 To vaArray.UpperBound(1)
        vaArray(n, 0) = n + 1
      Next
      nNomor.Value = vaArray.UpperBound(1) + 2
      tdbgrid1.ReBind
    End If
  End If
End Sub

