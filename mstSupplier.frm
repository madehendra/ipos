VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form mstSupplier 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "DATA SUPPLIER"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   2520
      Left            =   180
      Top             =   120
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   4445
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      BackColor       =   -2147483633
      Begin BiSATextBoxProject.BiSABrowse cKode 
         Height          =   330
         Left            =   60
         TabIndex        =   0
         Top             =   60
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   582
         Text            =   "AAAAAA"
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
         MaxLength       =   6
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
      Begin BiSATextBoxProject.BiSATextBox cNama 
         Height          =   330
         Left            =   60
         TabIndex        =   1
         Top             =   435
         Width           =   6015
         _ExtentX        =   10610
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
         Caption         =   "Nama"
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
      Begin BiSATextBoxProject.BiSATextBox cAlamat 
         Height          =   330
         Left            =   60
         TabIndex        =   2
         Top             =   810
         Width           =   6285
         _ExtentX        =   11086
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
         MaxLength       =   50
         Appearance      =   0
         Caption         =   "Alamat"
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
      Begin BiSATextBoxProject.BiSATextBox cTelepon 
         Height          =   330
         Left            =   60
         TabIndex        =   3
         Top             =   1185
         Width           =   5025
         _ExtentX        =   8864
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
         MaxLength       =   30
         Appearance      =   0
         Caption         =   "Telepon"
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
      Begin BiSATextBoxProject.BiSATextBox cFax 
         Height          =   330
         Left            =   60
         TabIndex        =   4
         Top             =   1560
         Width           =   4785
         _ExtentX        =   8440
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
         MaxLength       =   30
         Appearance      =   0
         Caption         =   "Fax"
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
      Begin BiSATextBoxProject.BiSATextBox cKota 
         Height          =   330
         Left            =   60
         TabIndex        =   5
         Top             =   1935
         Width           =   4530
         _ExtentX        =   7990
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
         MaxLength       =   20
         Appearance      =   0
         Caption         =   "Kota"
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
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   705
      Left            =   165
      Top             =   6465
      Width           =   11370
      _ExtentX        =   20055
      _ExtentY        =   1244
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
         TabIndex        =   6
         Top             =   150
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
         Picture         =   "mstSupplier.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   8655
         TabIndex        =   7
         Top             =   150
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
         Picture         =   "mstSupplier.frx":028A
      End
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   1170
         TabIndex        =   8
         Top             =   150
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
         Picture         =   "mstSupplier.frx":0429
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   435
         Left            =   105
         TabIndex        =   9
         Top             =   150
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
         Picture         =   "mstSupplier.frx":0555
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   10185
         TabIndex        =   10
         Top             =   150
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
         Picture         =   "mstSupplier.frx":0700
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   9105
         TabIndex        =   11
         Top             =   150
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
         Picture         =   "mstSupplier.frx":07A6
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   3600
      Left            =   195
      TabIndex        =   12
      Top             =   2700
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   6350
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "NO"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "KODE"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "NAMA"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "ALAMAT"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "KOTA"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "TELP"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   873
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   15790320
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1244"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1164"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=2672"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2593"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=512"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=4683"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=4604"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=512"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=5741"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=5662"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=512"
      Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(25)=   "Column(4).Width=2937"
      Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2858"
      Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=512"
      Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(31)=   "Column(5).Width=2593"
      Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=2514"
      Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=512"
      Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      BorderStyle     =   0
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=58,.parent=13"
      _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=14"
      _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=15"
      _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=17"
      _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.alignment=0"
      _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=54,.parent=13,.alignment=0"
      _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=32,.parent=13,.alignment=0"
      _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=46,.parent=13,.alignment=0"
      _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(5).Style:id=50,.parent=13,.alignment=0"
      _StyleDefs(53)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
      _StyleDefs(56)  =   "Named:id=33:Normal"
      _StyleDefs(57)  =   ":id=33,.parent=0"
      _StyleDefs(58)  =   "Named:id=34:Heading"
      _StyleDefs(59)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(60)  =   ":id=34,.wraptext=-1"
      _StyleDefs(61)  =   "Named:id=35:Footing"
      _StyleDefs(62)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(63)  =   "Named:id=36:Selected"
      _StyleDefs(64)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(65)  =   "Named:id=37:Caption"
      _StyleDefs(66)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(67)  =   "Named:id=38:HighlightRow"
      _StyleDefs(68)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(69)  =   "Named:id=39:EvenRow"
      _StyleDefs(70)  =   ":id=39,.parent=33,.bgcolor=&HE6E6E6&"
      _StyleDefs(71)  =   "Named:id=40:OddRow"
      _StyleDefs(72)  =   ":id=40,.parent=33"
      _StyleDefs(73)  =   "Named:id=41:RecordSelector"
      _StyleDefs(74)  =   ":id=41,.parent=34"
      _StyleDefs(75)  =   "Named:id=42:FilterBar"
      _StyleDefs(76)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "mstSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lClick As Boolean
Dim dbData As New ADODB.Recordset
Dim dbSupplier As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim lEdit As Boolean
Dim nPos As SisPos
Dim Validasi As Variant
Dim vaArray As New XArrayDB

Private Sub HapusData()
Dim cInfo As String

  cInfo = "Kode: " & cKode.Text & vbCrLf
  cInfo = cInfo & "Nama: " & cNama.Text & vbCrLf
  cInfo = cInfo & "Alamat: " & cAlamat.Text & vbCrLf
  cInfo = cInfo & "Kota: " & cKota.Text & vbCrLf
  
  If MsgBox("Data Benar-benar dihapus ?" & vbCrLf & vbCrLf & cInfo, vbQuestion + vbYesNo) = vbYes Then
    If lExist(objData, "totpembelian", "kodesupplier", GetKode) Then
      MsgBox "Maaf, data ini masih digunakan oleh sistem" & vbCrLf & "Tidak bisa dihapus"
      InitDel
      Exit Sub
    End If
    
    If lExist(objData, "totrtnpembelian", "kodesupplier", GetKode) Then
      MsgBox "Maaf, data ini masih digunakan oleh sistem" & vbCrLf & "Tidak bisa dihapus"
      InitDel
      Exit Sub
    End If
    objData.Delete GetDSN, "supplier", "kodesupplier", sisAssign, GetKode
  End If
  InitDel
End Sub

Private Sub InitDel()
  initvalue
  GetLoadRows
  GetEdit False
End Sub

Private Function GetKode() As String
  GetKode = cKode.Text
End Function

Private Sub cKode_Validate(Cancel As Boolean)
  Set dbData = objData.Browse(GetDSN, "supplier", , "kodesupplier", sisAssign, GetKode)
  If dbData.RecordCount > 0 Then
    GetMemory
    If nPos = Delete Then
      HapusData
    End If
  End If
End Sub

Private Sub cKode1_Validate(Cancel As Boolean)
  If Trim(cKode1.Text) = "" Then
    MsgBox "Data Harus Diisi, Ulangi Pengisian", vbExclamation
    Cancel = True
    cKode.SetFocus
  End If
End Sub

Private Sub cmdAktivasi_Click()
  frmAktivasi.Action Me
End Sub

Private Sub cWilayah_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "Wilayah", "Kode,Keterangan", "Kode", sisContent, cWilayah.Text, , "Kode,Keterangan")
  If Not dbData.EOF Then
    If dbData.RecordCount > 1 Then cWilayah.Text = cWilayah.Browse(dbData)
    cWilayah.Text = GetNull(dbData!Kode)
    cNamaWilayah.Text = GetNull(dbData!keterangan, "")
  Else
    cWilayah.Default
    cNamaWilayah.Default
  End If
End Sub

Private Sub cWilayah_Validate(Cancel As Boolean)
  If Trim(cWilayah.Text) <> "" Then
    cWilayah_ButtonClick
  End If
End Sub

Private Sub cJenisUsaha_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "JenisUsaha", "Kode,Keterangan", "Kode", sisContent, cJenisUsaha.Text, , "Kode,Keterangan")
  If Not dbData.EOF Then
    If dbData.RecordCount > 1 Then
      cJenisUsaha.Text = cJenisUsaha.Browse(dbData)
    End If
    cJenisUsaha.Text = GetNull(dbData!Kode)
    cNamaJenisUsaha.Text = GetNull(dbData!keterangan, "")
  Else
    cJenisUsaha.Default
    cNamaJenisUsaha.Default
  End If
End Sub

Private Sub cJenisUsaha_Validate(Cancel As Boolean)
  If Trim(cJenisUsaha.Text) <> "" Then
    cJenisUsaha_ButtonClick
  End If
End Sub

Private Sub cmdAdd_Click()
  GetEdit True
  initvalue
  cKode.SetFocus
  nPos = Add
  
End Sub

Private Sub GetEdit(lPar As Boolean)
  BiSAFrame1.Enabled = lPar
  lEdit = lPar
  SetButton cmdSimpan, cmdKeluar, cmdAdd, cmdEdit, cmdHapus, nPos, lPar
End Sub

Private Sub cmdBrowse_Click()
  Load rptBrowseSupplier
  rptBrowseSupplier.Show
End Sub

Private Sub cmdEdit_Click()
  GetEdit True
  cKode.SetFocus
  nPos = Edit
End Sub

Private Sub cmdHapus_Click()
  nPos = Delete
  GetEdit True
  cKode.SetFocus
  HapusData
  GetLoadRows
End Sub

Private Sub DeleteData()
  If MsgBox("Data benar-benar akan dihapus ", vbExclamation + vbYesNo) = vbYes Then
     objData.Delete GetDSN, "Supplier", "Kode", sisAssign, GetKode
  End If
  initvalue
  GetEdit False
End Sub

Private Sub cmdKeluar_Click()
  If Not lEdit Then
    Unload Me
  Else
    initvalue
    GetEdit False
  End If
End Sub

Private Sub cmdSimpan_Click()
Dim vaField
Dim vaValue

  If ValidSaving() Then
    objData.Start GetDSN
    vaField = Array("kodesupplier", "nama", "alamat", "telepon", "fax", "kota")
    vaValue = Array(Trim(GetKode), cNama.Text, cAlamat.Text, cTelepon.Text, cFax.Text, cKota.Text)
    objData.Update GetDSN, "supplier", "kodesupplier = '" & GetKode & "'", vaField, vaValue
    objData.Save GetDSN
    initvalue
    GetEdit False
  End If
  GetLoadRows
End Sub

Static Function ValidSaving() As Boolean
  ValidSaving = True
  
  If Not CheckData(cKode.Text, "Kode Harus Diisi, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    cKode.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cNama.Text, "Nama Harus Diisi, Silahkan Mengulangi Pengisian") Then
    ValidSaving = False
    cNama.SetFocus
    Exit Function
  End If
  
  If InStr(cKode.Text, " ") > 0 Then
    ValidSaving = False
    MsgBox "Karakter spasi tidak diijinkan"
    cKode.SetFocus
    Exit Function
  End If
  
End Function

Private Sub initvalue()
  cKode.Default
  cNama.Default
  cAlamat.Default
  cTelepon.Default
  cFax.Default
  cKota.Default
End Sub

Private Sub GetMemory()
  Set dbData = objData.Browse(GetDSN, "supplier s", "s.*", "s.kodesupplier", sisAssign, GetKode)
  If dbData.RecordCount > 0 Then
    cNama.Text = dbData!nama
    cAlamat.Text = dbData!alamat
    cTelepon.Text = dbData!telepon
    cFax.Text = dbData!fax
    cKota.Text = dbData!kota
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  initvalue
  GetEdit False
  TabIndex cKode, n
  TabIndex cNama, n
  TabIndex cAlamat, n
  TabIndex cTelepon, n
  TabIndex cFax, n
  TabIndex cKota, n
  
  TabIndex cmdAdd, n
  TabIndex cmdEdit, n
  TabIndex cmdHapus, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  TabIndex cmdAktivasi, n
  GetLoadRows
End Sub

Private Sub GetLoadRows()
Dim n As Integer

  vaArray.ReDim 0, -1, 0, 5
  Set dbData = objData.Browse(GetDSN, "supplier", "kodesupplier,nama,alamat,kota,telepon", "kodesupplier", sisContent, TDBGrid1.Columns(1).FilterText, " and nama LIKE '%" & TDBGrid1.Columns(2).FilterText & "%' AND alamat LIKE '%" & TDBGrid1.Columns(3).FilterText & "%' AND kota LIKE '%" & TDBGrid1.Columns(4).FilterText & "%' AND telepon LIKE '%" & TDBGrid1.Columns(5).FilterText & "%'", "nama")
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = n + 1
      vaArray(n, 1) = GetNull(dbData!kodesupplier)
      vaArray(n, 2) = GetNull(dbData!nama)
      vaArray(n, 3) = GetNull(dbData!alamat)
      vaArray(n, 4) = GetNull(dbData!kota)
      vaArray(n, 5) = GetNull(dbData!telepon)
      dbData.MoveNext
    Loop
  End If
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
  TDBGrid1.Refresh
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    GetLoadRows
  End If
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim db As New ADODB.Recordset

  Set db = objData.Browse(GetDSN, "supplier s", "s.*", "s.kodesupplier", sisAssign, TDBGrid1.Columns(1).Text)
  If Not db.EOF Then
    cKode.Text = GetNull(db!kodesupplier)
    cNama.Text = GetNull(db!nama)
    cAlamat.Text = GetNull(db!alamat)
    cTelepon.Text = GetNull(db!telepon)
    cFax.Text = GetNull(db!fax)
    cKota.Text = GetNull(db!kota)
  End If
End Sub
