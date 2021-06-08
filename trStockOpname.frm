VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trStockOpname 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "STOCK OPNAME"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   11505
   WindowState     =   2  'Maximized
   Begin SizerOneLibCtl.ElasticOne ElasticOne1 
      Height          =   1065
      Left            =   0
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   11505
      _cx             =   20294
      _cy             =   1879
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   4
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   1
      AutoSizeChildren=   8
      BorderWidth     =   6
      ChildSpacing    =   1
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   3
      GridCols        =   7
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"trStockOpname.frx":0000
      Begin VB.CheckBox Check1 
         Caption         =   "Posting Ke Buku Besar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7035
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   465
         Width           =   4365
      End
      Begin BiSATextBoxProject.BiSATextBox cKeterangan 
         Height          =   360
         Left            =   7035
         TabIndex        =   12
         Top             =   90
         Width           =   4380
         _ExtentX        =   7726
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
         Caption         =   "Ket."
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
      Begin BiSATextBoxProject.BiSABrowse cNomor 
         Height          =   345
         Left            =   255
         TabIndex        =   11
         Top             =   465
         Width           =   2895
         _ExtentX        =   5106
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
      Begin BiSADateProject.BiSADate dTgl 
         Height          =   360
         Left            =   240
         TabIndex        =   10
         Top             =   90
         Width           =   2910
         _ExtentX        =   5133
         _ExtentY        =   635
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
   End
   Begin SizerOneLibCtl.ElasticOne ElasticOne3 
      Height          =   570
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6720
      Width           =   11505
      _cx             =   20294
      _cy             =   1005
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   4
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   2
      AutoSizeChildren=   8
      BorderWidth     =   4
      ChildSpacing    =   0
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   1
      GridCols        =   12
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"trStockOpname.frx":008A
      Begin BiSAButtonProject.BiSAButton cmdHapus 
         Height          =   450
         Left            =   2085
         TabIndex        =   2
         Top             =   60
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   794
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
         Picture         =   "trStockOpname.frx":0138
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   450
         Left            =   3285
         TabIndex        =   3
         Top             =   60
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   794
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
         Picture         =   "trStockOpname.frx":03C2
      End
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   450
         Left            =   1035
         TabIndex        =   4
         Top             =   60
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   794
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
         Picture         =   "trStockOpname.frx":0561
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   450
         Left            =   60
         TabIndex        =   5
         Top             =   60
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   794
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
         Picture         =   "trStockOpname.frx":068D
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   450
         Left            =   10350
         TabIndex        =   6
         Top             =   60
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   794
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
         Picture         =   "trStockOpname.frx":0838
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   450
         Left            =   9225
         TabIndex        =   7
         Top             =   60
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   794
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
         Picture         =   "trStockOpname.frx":08DE
      End
   End
   Begin SizerOneLibCtl.ElasticOne ElasticOne2 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1065
      Width           =   11505
      _cx             =   20294
      _cy             =   9975
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   4
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   7
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      Begin TrueOleDBGrid70.TDBGrid tdbgrid1 
         Height          =   5610
         Left            =   30
         TabIndex        =   9
         Top             =   15
         Width           =   11445
         _ExtentX        =   20188
         _ExtentY        =   9895
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "KODE"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "BARCODE"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "INVENTORY"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "GOLONGAN"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "SATUAN"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "ON HAND"
         Columns(5).DataField=   ""
         Columns(5).NumberFormat=   "###,###,###,###,##0.00"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "ADJUST"
         Columns(6).DataField=   ""
         Columns(6).NumberFormat=   "###,###,###,###,##0.00"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "NEXT STOCK"
         Columns(7).DataField=   ""
         Columns(7).NumberFormat=   "###,###,###,###,##0.00"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   8
         Splits(0)._UserFlags=   0
         Splits(0).AllowSizing=   -1  'True
         Splits(0).Size  =   414
         Splits(0).Size.vt=   2
         Splits(0).RecordSelectorWidth=   873
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   15790320
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2408"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2328"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2646"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2566"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=513"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=5080"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=5001"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=3069"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2990"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2328"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2249"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=516"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=3598"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=3519"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=514"
         Splits(0)._ColumnProps(30)=   "Column(5).FetchStyle=1"
         Splits(0)._ColumnProps(31)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(32)=   "Column(6).Width=3149"
         Splits(0)._ColumnProps(33)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(34)=   "Column(6)._WidthInPix=3069"
         Splits(0)._ColumnProps(35)=   "Column(6)._ColStyle=514"
         Splits(0)._ColumnProps(36)=   "Column(6).FetchStyle=1"
         Splits(0)._ColumnProps(37)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(38)=   "Column(7).Width=3598"
         Splits(0)._ColumnProps(39)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(40)=   "Column(7)._WidthInPix=3519"
         Splits(0)._ColumnProps(41)=   "Column(7)._ColStyle=514"
         Splits(0)._ColumnProps(42)=   "Column(7).FetchStyle=1"
         Splits(0)._ColumnProps(43)=   "Column(7).Order=8"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         Appearance      =   0
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         RowDividerStyle =   0
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   15790320
         RowDividerColor =   15790320
         RowSubDividerColor=   15790320
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=255,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2"
         _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=2"
         _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(5).Style:id=66,.parent=13,.alignment=1"
         _StyleDefs(53)  =   "Splits(0).Columns(5).HeadingStyle:id=63,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(5).FooterStyle:id=64,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(5).EditorStyle:id=65,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.alignment=1"
         _StyleDefs(57)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(7).Style:id=58,.parent=13,.alignment=1"
         _StyleDefs(61)  =   "Splits(0).Columns(7).HeadingStyle:id=55,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(7).FooterStyle:id=56,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(7).EditorStyle:id=57,.parent=17"
         _StyleDefs(64)  =   "Named:id=33:Normal"
         _StyleDefs(65)  =   ":id=33,.parent=0"
         _StyleDefs(66)  =   "Named:id=34:Heading"
         _StyleDefs(67)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(68)  =   ":id=34,.wraptext=-1"
         _StyleDefs(69)  =   "Named:id=35:Footing"
         _StyleDefs(70)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(71)  =   "Named:id=36:Selected"
         _StyleDefs(72)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(73)  =   "Named:id=37:Caption"
         _StyleDefs(74)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(75)  =   "Named:id=38:HighlightRow"
         _StyleDefs(76)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(77)  =   "Named:id=39:EvenRow"
         _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&HE9E9E9&"
         _StyleDefs(79)  =   "Named:id=40:OddRow"
         _StyleDefs(80)  =   ":id=40,.parent=33"
         _StyleDefs(81)  =   "Named:id=41:RecordSelector"
         _StyleDefs(82)  =   ":id=41,.parent=34"
         _StyleDefs(83)  =   "Named:id=42:FilterBar"
         _StyleDefs(84)  =   ":id=42,.parent=33"
      End
   End
End
Attribute VB_Name = "trStockOpname"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Gudang As String

Dim nPos As SisPos
Dim lEdit As Boolean
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB
Dim cKode As String


Private Sub GetFakturBrowse(ByVal lStat As Boolean)
  cNomor.Button = lStat
End Sub

Private Sub cmdAktivasi_Click()
  frmAktivasi.Action Me
End Sub

Private Sub cmdAdd_Click()
  nPos = Add
  GetEdit True
  GetFakturBrowse False
  cNomor.Text = GetNomor("totstockopname", "nomorstockopname", GetID, sisModulTransaksi.StockOpname)
  GetLoadRows
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

Private Function ValidSaving() As Boolean
  ValidSaving = True
End Function

Private Sub cmdSimpan_Click()
Dim lSave As Boolean
Dim Faktur As String
Dim n As Integer

  lSave = True
  objData.Start GetDSN
  Faktur = cNomor.Text
  lSave = IIf(lSave, objData.Delete(GetDSN, "totstockopname", "nomorstockopname", sisAssign, Faktur), False)
  lSave = IIf(lSave, objData.Delete(GetDSN, "stockopname", "nomorstockopname", sisAssign, Faktur), False)
  lSave = IIf(lSave, objData.Delete(GetDSN, "kartustock", "nomor", sisAssign, Faktur), False)
  lSave = IIf(lSave, objData.Delete(GetDSN, "bukubesar", "faktur", sisAssign, Faktur), False)
  lSave = IIf(lSave, objData.Update(GetDSN, "totstockopname", "nomorstockopname = '" & Faktur & "'", Array("nomorstockopname", "tgl", "keterangan", "username", "datetime", "kodegudang"), Array(Faktur, Format(dTgl.Value, "yyyy-MM-dd"), cKeterangan.Text, GetRegistry(reg_Username), SNow, Gudang)), False)
  
  For n = 0 To vaArray.UpperBound(1)
    If vaArray(n, 6) <> 0 Then
      lSave = IIf(lSave, objData.Add(GetDSN, "stockopname", Array("nomorstockopname", "kodestock", "adjust", "kodegudang"), Array(Faktur, vaArray(n, 0), vaArray(n, 6), Gudang)), False)
      lSave = IIf(lSave, UpdKartuStock(objData, SisKartuStock.PenyesuaianStock, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), vaArray(n, 0), vaArray(n, 6), GetHargaPokok(objData, vaArray(n, 0)), 0, "Penyesuaian : " & cKeterangan.Text, Gudang, GetHargaPokok(objData, vaArray(n, 0))), False)
      'Posting
      If Check1.Value = 1 Then
        'jika - = biaya
        'jika + = pendapatan/modal
        If vaArray(n, 6) < 0 Then
          'jurnal
          ' biaya
          '  persediaan
          lSave = IIf(lSave, UpdKodeTr(objData, msPenyesuaian, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningPenyesuaianKurang), GetCostCenterUser(objData, GetRegistry(reg_Username)), "Penyesuaian " & vaArray(n, 2), Abs(vaArray(n, 6)) * GetHargaBeli(objData, vaArray(n, 0)), 0, "", SNow), False)
            lSave = IIf(lSave, UpdKodeTr(objData, msPenyesuaian, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), GetAkunInventory(objData, vaArray(n, 6)), GetCostCenterUser(objData, GetRegistry(reg_Username)), "Penyesuaian " & vaArray(n, 2), 0, Abs(vaArray(n, 6)) * GetHargaBeli(objData, vaArray(n, 0)), "", SNow), False)
        Else
          'jurnal
          ' persediaan
          '  pendapatan/modal
          lSave = IIf(lSave, UpdKodeTr(objData, msPenyesuaian, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), GetAkunInventory(objData, vaArray(n, 6)), GetCostCenterUser(objData, GetRegistry(reg_Username)), "Penyesuaian " & vaArray(n, 2), Abs(vaArray(n, 6)) * GetHargaBeli(objData, vaArray(n, 0)), 0, "", SNow), False)
            lSave = IIf(lSave, UpdKodeTr(objData, msPenyesuaian, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningPenyesuian), GetCostCenterUser(objData, GetRegistry(reg_Username)), "Penyesuaian " & vaArray(n, 2), 0, Abs(vaArray(n, 6)) * GetHargaBeli(objData, vaArray(n, 0)), "", SNow), False)
        End If
      End If
    
    End If
  Next
  
  
  If lSave Then
    objData.Save GetDSN
  Else
    objData.Cancel GetDSN
  End If
  GetLoadRows
  GetEdit False
End Sub

Private Sub cNomor_ButtonClick()
Dim lSave As Boolean
lSave = True

  Set dbData = objData.Browse(GetDSN, "totstockopname", "nomorstockopname,keterangan,username", "tgl", sisAssign, Format(dTgl.Value, "yyyy-MM-dd"), " AND kodegudang = '" & Gudang & "'")
  If Not dbData.EOF Then
    cNomor.Text = cNomor.Browse(dbData)
    cKeterangan.Text = GetNull(dbData!keterangan)
    GetLoadRows
    Me.Refresh
    objData.Start GetDSN
    If nPos = Delete Then
      If MsgBox("Yakin data akan dihapus?", vbYesNo) = vbYes Then
        lSave = IIf(lSave, objData.Delete(GetDSN, "totstockopname", "nomorstockopname", sisAssign, cNomor.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "stockopname", "nomorstockopname", sisAssign, cNomor.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "kartustock", "nomor", sisAssign, cNomor.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "bukubesar", "faktur", sisAssign, cNomor.Text), False)
      End If
      If lSave Then
        objData.Save GetDSN
      Else
        objData.Cancel GetDSN
      End If
      GetEdit False
      initvalue
    End If
  End If
End Sub

Private Sub Form_Activate()
  Me.Refresh
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  initvalue
  Me.Caption = "STOCK OPNAME GUDANG " & UCase(Gudang)
  GetEdit False
  TabIndex dTgl, n
  TabIndex cNomor, n
  TabIndex cKeterangan, n
  TabIndex cmdAdd, n
  TabIndex cmdEdit, n
  TabIndex cmdHapus, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  TabIndex cmdAktivasi, n
End Sub

Private Sub GetLoadRows()
Dim n As Integer
Dim cSQL As String
  
  cSQL = ""
  If Trim(Gudang) <> "" Then
    cSQL = " AND k.kodegudang = '" & Gudang & "'"
  End If
  If nPos = Edit Or nPos = Delete Then
    GetRows2
  Else
    vaArray.ReDim 0, -1, 0, 7
    Set dbData = objData.Browse(GetDSN, "stock s", "s.kodestock,s.barcode,s.nama,g.keterangan,s.kodesatuan,sum(k.debet-k.kredit) as saldostock", "s.kodestock", sisContent, TDBGrid1.Columns(0).FilterText, " AND barcode LIKE '%" & TDBGrid1.Columns(1).FilterText & "%' AND nama LIKE '%" & TDBGrid1.Columns(2).FilterText & "%' AND g.keterangan LIKE '%" & TDBGrid1.Columns(3).FilterText & "%' AND s.kodesatuan LIKE '%" & TDBGrid1.Columns(4).FilterText & "%' AND s.jenis = 1 GROUP BY s.kodestock", "s.kodestock desc", Array("LEFT JOIN golongan g on g.kodegolongan = s.kodegolongan", "LEFT JOIN kartustock k on k.kodestock = s.kodestock " & cSQL))
    If Not dbData.EOF Then
      Do While Not dbData.EOF
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        n = vaArray.UpperBound(1)
        vaArray(n, 0) = GetNull(dbData!KodeStock)
        vaArray(n, 1) = GetNull(dbData!barcode)
        vaArray(n, 2) = GetNull(dbData!nama)
        vaArray(n, 3) = GetNull(dbData!keterangan)
        vaArray(n, 4) = GetNull(dbData!kodesatuan)
        vaArray(n, 5) = GetNull(dbData!saldostock)
        vaArray(n, 6) = 0
        vaArray(n, 7) = vaArray(n, 5)
        dbData.MoveNext
      Loop
    End If
    Set TDBGrid1.Array = vaArray
    TDBGrid1.ReBind
    TDBGrid1.Refresh
  End If
End Sub
Private Sub GetRows2()
Dim n As Integer
Dim cSQL As String
  
  cSQL = ""
  If Trim(Gudang) <> "" Then
    cSQL = " AND k.kodegudang = '" & Gudang & "'"
  End If
  vaArray.ReDim 0, -1, 0, 7
  Set dbData = objData.Browse(GetDSN, "stock s", "s.kodestock,s.barcode,s.nama,g.keterangan,s.kodesatuan,p.adjust,sum(k.debet-k.kredit) as saldostock", "s.kodestock", sisContent, TDBGrid1.Columns(0).FilterText, " AND barcode LIKE '%" & TDBGrid1.Columns(1).FilterText & "%' AND nama LIKE '%" & TDBGrid1.Columns(2).FilterText & "%' AND g.keterangan LIKE '%" & TDBGrid1.Columns(3).FilterText & "%' AND s.kodesatuan LIKE '%" & TDBGrid1.Columns(4).FilterText & "%'  AND p.nomorstockopname = '" & cNomor.Text & "' GROUP BY s.kodestock", "s.kodestock desc", Array("LEFT JOIN golongan g on g.kodegolongan = s.kodegolongan", "LEFT JOIN stockopname p on p.kodestock = s.kodestock", "LEFT JOIN kartustock k on k.kodestock = s.kodestock " & cSQL))
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!KodeStock)
      vaArray(n, 1) = GetNull(dbData!barcode)
      vaArray(n, 2) = GetNull(dbData!nama)
      vaArray(n, 3) = GetNull(dbData!keterangan)
      vaArray(n, 4) = GetNull(dbData!kodesatuan)
      vaArray(n, 6) = GetNull(dbData!adjust)
      vaArray(n, 7) = GetNull(dbData!saldostock)
      vaArray(n, 5) = vaArray(n, 7) - vaArray(n, 6)
      dbData.MoveNext
    Loop
  End If
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
  TDBGrid1.Refresh
End Sub

Private Sub initvalue()
  dTgl.Value = Date
  cNomor.Default
  cKeterangan.Default
  Check1.Value = 1
  vaArray.ReDim 0, -1, 0, 7
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
  TDBGrid1.Refresh
End Sub

Private Sub GetEdit(lPar As Boolean)
  ElasticOne1.Enabled = lPar
  ElasticOne2.Enabled = lPar
  lEdit = lPar
  initvalue
  SetButton cmdSimpan, cmdKeluar, cmdAdd, cmdEdit, cmdHapus, nPos, lPar, cmdAktivasi
  
  If lPar Then
    dTgl.SetFocus
    If nPos = Add Then
      cNomor.Enabled = False
      cNomor.BackColor = vbButtonFace
    Else
      cNomor.Enabled = True
      cNomor.BackColor = vbWindowBackground
      cNomor.CaptionBackColor = vbButtonFace
    End If
  End If
End Sub

Private Sub Form_Resize()
Dim nSisaLebar As Double

  If Me.WindowState = 2 Then
    Me.Refresh
    nSisaLebar = TDBGrid1.Width - TDBGrid1.Columns(0).Width - TDBGrid1.Columns(1).Width - TDBGrid1.Columns(3).Width - TDBGrid1.Columns(4).Width - TDBGrid1.Columns(5).Width - TDBGrid1.Columns(6).Width - TDBGrid1.Columns(7).Width
    TDBGrid1.Columns(2).Width = nSisaLebar - 1000
  End If
End Sub

Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
  TDBGrid1.Update
End Sub

Private Sub tdbgrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
  If Not IsNumeric(TDBGrid1.Columns(6).Value) Then
    Cancel = True
    Exit Sub
  End If
  If Not IsNumeric(TDBGrid1.Columns(7).Value) Then
    Cancel = True
    Exit Sub
  End If
  If ColIndex < 6 Then
    Cancel = True
    Exit Sub
  End If
  TDBGrid1.Columns(7).Value = Val(TDBGrid1.Columns(5).Value) + Val(TDBGrid1.Columns(6).Value)
End Sub

Private Sub TDBGrid1_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueOleDBGrid70.StyleDisp)
On Error Resume Next

  Select Case Col
      Case 5
          Dim Col1 As Long
          Col1 = CLng(TDBGrid1.Columns(5).CellText(Bookmark))
          If Col1 < 0 Then CellStyle.ForeColor = vbRed
      Case 6
          Dim Col2 As Long
          Col2 = CLng(TDBGrid1.Columns(6).CellText(Bookmark))
          If Col2 < 0 Then CellStyle.ForeColor = vbRed
      Case 7
          Dim Col3 As Long
          Col3 = CLng(TDBGrid1.Columns(7).CellText(Bookmark))
          If Col3 < 0 Then CellStyle.ForeColor = vbRed
          
  End Select
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    GetLoadRows
  End If
End Sub

