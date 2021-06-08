VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form mstStock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DATA INVENTORY"
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   15465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   15465
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   8100
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   15150
      _cx             =   26723
      _cy             =   14287
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   0
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   -2147483633
      TabOutlineColor =   0
      FrontTabForeColor=   -2147483630
      Caption         =   "&1 Input Inventory, Non Inventory|&2 Daftar Inventory |&3 Daftar Non Inventory"
      Align           =   0
      CurrTab         =   0
      FirstTab        =   0
      Style           =   3
      Position        =   0
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   6
      BorderWidth     =   1
      BoldCurrent     =   -1  'True
      DogEars         =   0   'False
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   3
      TabHeight       =   0
      TabCaptionPos   =   3
      TabPicturePos   =   0
      CaptionEmpty    =   ""
      Separators      =   0   'False
      Begin SizerOneLibCtl.ElasticOne ElasticOne2 
         Height          =   7710
         Left            =   16080
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   360
         Width           =   15090
         _cx             =   26617
         _cy             =   13600
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
         Appearance      =   0
         MousePointer    =   0
         _ConvInfo       =   1
         Version         =   700
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   0
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
         Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
            Height          =   6585
            Left            =   240
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   375
            Width           =   14490
            _ExtentX        =   25559
            _ExtentY        =   11615
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
            Columns(2).Caption=   "BARCODE"
            Columns(2).DataField=   ""
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "NAMA"
            Columns(3).DataField=   ""
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "GOLONGAN"
            Columns(4).DataField=   ""
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "SATUAN"
            Columns(5).DataField=   ""
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   6
            Splits(0)._UserFlags=   0
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1270"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1191"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=2408"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2328"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=2646"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2566"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=11483"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=11404"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=4948"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=4868"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(31)=   "Column(5).Width=2328"
            Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=2249"
            Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=516"
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=825"
            _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
            _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.namedParent=40"
            _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=58,.parent=13"
            _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=14"
            _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=15"
            _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=17"
            _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=50,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(5).Style:id=54,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
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
            _StyleDefs(70)  =   ":id=39,.parent=33,.bgcolor=&H9BCDFF&"
            _StyleDefs(71)  =   "Named:id=40:OddRow"
            _StyleDefs(72)  =   ":id=40,.parent=33,.bgcolor=&HFFFFFF&"
            _StyleDefs(73)  =   "Named:id=41:RecordSelector"
            _StyleDefs(74)  =   ":id=41,.parent=34"
            _StyleDefs(75)  =   "Named:id=42:FilterBar"
            _StyleDefs(76)  =   ":id=42,.parent=33"
         End
      End
      Begin SizerOneLibCtl.ElasticOne ElasticOne3 
         Height          =   7710
         Left            =   30
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   360
         Width           =   15090
         _cx             =   26617
         _cy             =   13600
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
         Appearance      =   0
         MousePointer    =   0
         _ConvInfo       =   1
         Version         =   700
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   0
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
         Begin BiSAFramProject.BiSAFrame BiSAFrame1 
            Height          =   7485
            Left            =   75
            Top             =   210
            Width           =   14760
            _ExtentX        =   26035
            _ExtentY        =   13203
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
               Caption         =   "Check1"
               Height          =   195
               Left            =   4260
               TabIndex        =   35
               Top             =   6165
               Width           =   270
            End
            Begin VB.TextBox Text1 
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
               Height          =   2205
               Left            =   7110
               MultiLine       =   -1  'True
               TabIndex        =   14
               TabStop         =   0   'False
               Text            =   "mstStock.frx":0000
               Top             =   525
               Width           =   6450
            End
            Begin VB.CheckBox chStatus 
               Caption         =   "Check1"
               Height          =   210
               Left            =   2235
               TabIndex        =   41
               TabStop         =   0   'False
               Top             =   6945
               Width           =   255
            End
            Begin BiSAFramProject.BiSAFrame BiSAFrame4 
               Height          =   525
               Left            =   2205
               Top             =   3630
               Width           =   3720
               _ExtentX        =   6562
               _ExtentY        =   926
               BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   -2147483634
               Begin VB.OptionButton optJenisInventory 
                  BackColor       =   &H8000000E&
                  Caption         =   "&9 Non Inventory"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   1
                  Left            =   1530
                  TabIndex        =   22
                  Top             =   165
                  Width           =   1524
               End
               Begin VB.OptionButton optJenisInventory 
                  BackColor       =   &H8000000E&
                  Caption         =   "&1 Inventory"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   0
                  Left            =   165
                  TabIndex        =   21
                  Top             =   165
                  Value           =   -1  'True
                  Width           =   1215
               End
            End
            Begin BiSAFramProject.BiSAFrame BiSAFrame2 
               Height          =   495
               Left            =   7335
               Top             =   4425
               Width           =   3705
               _ExtentX        =   6535
               _ExtentY        =   873
               BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   -2147483634
               Begin VB.OptionButton optBiaya 
                  BackColor       =   &H8000000E&
                  Caption         =   "&1 Ya"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   0
                  Left            =   195
                  TabIndex        =   27
                  Top             =   105
                  Width           =   1185
               End
               Begin VB.OptionButton optBiaya 
                  BackColor       =   &H8000000E&
                  Caption         =   "&2 Tidak"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   1
                  Left            =   1545
                  TabIndex        =   28
                  Top             =   105
                  Value           =   -1  'True
                  Width           =   885
               End
            End
            Begin BiSANumberBoxProject.BiSANumberBox nHargaBeli 
               Height          =   330
               Left            =   1095
               TabIndex        =   31
               Top             =   5340
               Width           =   2940
               _ExtentX        =   5186
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
               Caption         =   "Harga Beli"
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
               Left            =   1080
               TabIndex        =   5
               Top             =   870
               Width           =   5670
               _ExtentX        =   10001
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
               Caption         =   "Nama"
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
            Begin BiSATextBoxProject.BiSABrowse cSatuan 
               Height          =   330
               Left            =   1080
               TabIndex        =   8
               Top             =   1665
               Width           =   2235
               _ExtentX        =   3942
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
               Caption         =   "Satuan"
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
            Begin BiSATextBoxProject.BiSATextBox cNamaSatuan 
               Height          =   330
               Left            =   3330
               TabIndex        =   9
               Top             =   1665
               Width           =   2775
               _ExtentX        =   4895
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
            Begin BiSATextBoxProject.BiSABrowse cGolongan 
               Height          =   330
               Left            =   1080
               TabIndex        =   10
               Top             =   2025
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
               Appearance      =   0
               Button          =   -1  'True
               Caption         =   "Golongan"
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
            Begin BiSATextBoxProject.BiSATextBox cNamaGolongan 
               Height          =   330
               Left            =   3930
               TabIndex        =   11
               Top             =   2025
               Width           =   2775
               _ExtentX        =   4895
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
            Begin BiSANumberBoxProject.BiSANumberBox nHargaJual 
               Height          =   330
               Left            =   585
               TabIndex        =   29
               Top             =   4980
               Width           =   3180
               _ExtentX        =   5609
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
               Caption         =   "Harga Jual"
               CaptionWidth    =   1500
               BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin BiSATextBoxProject.BiSATextBox cKode 
               Height          =   330
               Left            =   1080
               TabIndex        =   2
               Top             =   135
               Width           =   3735
               _ExtentX        =   6588
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
               Caption         =   "Kode"
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
               Left            =   1080
               TabIndex        =   4
               Top             =   495
               Width           =   3450
               _ExtentX        =   6085
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
               Caption         =   "Barcode"
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
            Begin BiSANumberBoxProject.BiSANumberBox nPoin 
               Height          =   330
               Left            =   1095
               TabIndex        =   38
               Top             =   6555
               Width           =   2250
               _ExtentX        =   3969
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
               Caption         =   "Poin"
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
            Begin BiSANumberBoxProject.BiSANumberBox nDiskonPenjualan 
               Height          =   330
               Left            =   1095
               TabIndex        =   32
               Top             =   5700
               Width           =   2025
               _ExtentX        =   3572
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
               Caption         =   "Dsc Jual %"
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
            Begin BiSATextBoxProject.BiSABrowse cKategori 
               Height          =   330
               Left            =   1080
               TabIndex        =   12
               Top             =   2370
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
               Appearance      =   0
               Button          =   -1  'True
               Caption         =   "Kategori"
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
            Begin BiSATextBoxProject.BiSATextBox cNamaKategori 
               Height          =   330
               Left            =   3945
               TabIndex        =   13
               Top             =   2370
               Width           =   2775
               _ExtentX        =   4895
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
            Begin BiSATextBoxProject.BiSABrowse cKodeGroupSales 
               Height          =   330
               Left            =   1080
               TabIndex        =   15
               Top             =   2715
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
               Appearance      =   0
               Button          =   -1  'True
               Caption         =   "Group"
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
            Begin BiSATextBoxProject.BiSATextBox cNamaGroupSales 
               Height          =   330
               Left            =   3945
               TabIndex        =   16
               Top             =   2715
               Width           =   2775
               _ExtentX        =   4895
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
            Begin BiSAFramProject.BiSAFrame BiSAFrame5 
               Height          =   480
               Left            =   2205
               Top             =   4410
               Width           =   3690
               _ExtentX        =   6509
               _ExtentY        =   847
               BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   -2147483634
               Begin VB.OptionButton optAutoBiaya 
                  BackColor       =   &H8000000E&
                  Caption         =   "Tidak (&0)"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   1
                  Left            =   1590
                  TabIndex        =   26
                  Top             =   120
                  Value           =   -1  'True
                  Width           =   1035
               End
               Begin VB.OptionButton optAutoBiaya 
                  BackColor       =   &H8000000E&
                  Caption         =   "Ya (&1)"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   0
                  Left            =   210
                  TabIndex        =   25
                  Top             =   120
                  Width           =   1185
               End
            End
            Begin BiSAFramProject.BiSAFrame BiSAFrame6 
               Height          =   525
               Left            =   2205
               Top             =   3075
               Width           =   3720
               _ExtentX        =   6562
               _ExtentY        =   926
               BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   -2147483634
               Begin VB.OptionButton optKonsinyasi 
                  BackColor       =   &H8000000E&
                  Caption         =   "&0 Non Konsi"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   0
                  Left            =   150
                  TabIndex        =   18
                  Top             =   165
                  Value           =   -1  'True
                  Width           =   1230
               End
               Begin VB.OptionButton optKonsinyasi 
                  BackColor       =   &H8000000E&
                  Caption         =   "&1 Konsinyasi"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   1
                  Left            =   1515
                  TabIndex        =   19
                  Top             =   165
                  Width           =   1485
               End
            End
            Begin BiSATextBoxProject.BiSABrowse cNamaSupplier 
               Height          =   330
               Left            =   3795
               TabIndex        =   7
               Top             =   1275
               Width           =   2970
               _ExtentX        =   5239
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
            Begin BiSATextBoxProject.BiSATextBox cKodeSupplier 
               Height          =   330
               Left            =   1080
               TabIndex        =   6
               TabStop         =   0   'False
               Top             =   1275
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
            Begin BiSANumberBoxProject.BiSANumberBox BiSANumberBox1 
               Height          =   330
               Left            =   600
               TabIndex        =   34
               Top             =   6075
               Width           =   3450
               _ExtentX        =   6085
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
               Caption         =   "Promo Price"
               CaptionWidth    =   1500
               BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin BiSADateProject.BiSADate dTglAwal 
               Height          =   330
               Left            =   4605
               TabIndex        =   36
               Top             =   6075
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   582
               Value           =   "03-04-2011"
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
            Begin BiSADateProject.BiSADate dTglAkhir 
               Height          =   330
               Left            =   6165
               TabIndex        =   37
               Top             =   6075
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   582
               Value           =   "03-04-2011"
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
            Begin VB.Label Label10 
               Caption         =   "Berlaku Dari- Sampai Tgl"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   5040
               TabIndex        =   33
               Top             =   5790
               Width           =   2235
            End
            Begin VB.Label Label8 
               Caption         =   "Jenis"
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
               Left            =   1125
               TabIndex        =   17
               Top             =   3240
               Width           =   915
            End
            Begin VB.Label Label7 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Margin Profit Caption"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   3810
               TabIndex        =   30
               Top             =   4980
               Width           =   4530
            End
            Begin VB.Label Label6 
               Caption         =   "Keterangan :"
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
               Left            =   7110
               TabIndex        =   3
               Top             =   285
               Width           =   1005
            End
            Begin VB.Label Label3 
               Caption         =   "Jika Non Inventory : Auto Post As Biaya saat Penjualan?"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   2250
               TabIndex        =   23
               Top             =   4215
               Width           =   4950
            End
            Begin VB.Label Label1 
               Caption         =   "Tipe"
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
               Left            =   1125
               TabIndex        =   20
               Top             =   3795
               Width           =   915
            End
            Begin VB.Label Label2 
               Caption         =   "Jika Non Inventory : Auto Post As Biaya pada saat Pembelian?"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   7335
               TabIndex        =   24
               Top             =   4215
               Width           =   5550
            End
            Begin VB.Label Label4 
               Caption         =   "Status"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1140
               TabIndex        =   39
               Top             =   6930
               Width           =   645
            End
            Begin VB.Label Label5 
               Caption         =   "Berikan tanda centang untuk membuat inventori ini TIDAK AKTIF"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   2610
               TabIndex        =   40
               Top             =   6930
               Width           =   5820
            End
         End
      End
      Begin SizerOneLibCtl.ElasticOne ElasticOne1 
         Height          =   7710
         Left            =   15780
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   360
         Width           =   15090
         _cx             =   26617
         _cy             =   13600
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
         Appearance      =   0
         MousePointer    =   0
         _ConvInfo       =   1
         Version         =   700
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   0
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
         Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
            Height          =   6600
            Left            =   225
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   375
            Width           =   14610
            _ExtentX        =   25770
            _ExtentY        =   11642
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
            Columns(2).Caption=   "BARCODE"
            Columns(2).DataField=   ""
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "NAMA"
            Columns(3).DataField=   ""
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "GOLONGAN"
            Columns(4).DataField=   ""
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "STOK"
            Columns(5).DataField=   ""
            Columns(5).NumberFormat=   "###,###,##0.00"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).DataField=   ""
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "HARGA BELI"
            Columns(7).DataField=   ""
            Columns(7).NumberFormat=   "###,###,###,##0.00"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "HARGA JUAL"
            Columns(8).DataField=   ""
            Columns(8).NumberFormat=   "###,###,###,##0.00"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "HARGA POKOK"
            Columns(9).DataField=   ""
            Columns(9).NumberFormat=   "###,###,###,##0.00"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "SATUAN"
            Columns(10).DataField=   ""
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   11
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   873
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).AlternatingRowStyle=   -1  'True
            Splits(0).DividerColor=   15790320
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=11"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1244"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1164"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=2408"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2328"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=2646"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2566"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=6324"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=6244"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=2910"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2831"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(31)=   "Column(5).Width=1588"
            Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=1508"
            Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=514"
            Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(37)=   "Column(6).Width=741"
            Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=661"
            Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=516"
            Splits(0)._ColumnProps(42)=   "Column(6).Visible=0"
            Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(44)=   "Column(7).Width=2619"
            Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=2540"
            Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=514"
            Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(50)=   "Column(8).Width=2487"
            Splits(0)._ColumnProps(51)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(52)=   "Column(8)._WidthInPix=2408"
            Splits(0)._ColumnProps(53)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(54)=   "Column(8)._ColStyle=514"
            Splits(0)._ColumnProps(55)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(56)=   "Column(9).Width=2884"
            Splits(0)._ColumnProps(57)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(58)=   "Column(9)._WidthInPix=2805"
            Splits(0)._ColumnProps(59)=   "Column(9)._EditAlways=0"
            Splits(0)._ColumnProps(60)=   "Column(9)._ColStyle=514"
            Splits(0)._ColumnProps(61)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(62)=   "Column(10).Width=2725"
            Splits(0)._ColumnProps(63)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(64)=   "Column(10)._WidthInPix=2646"
            Splits(0)._ColumnProps(65)=   "Column(10)._EditAlways=0"
            Splits(0)._ColumnProps(66)=   "Column(10)._ColStyle=514"
            Splits(0)._ColumnProps(67)=   "Column(10).Order=11"
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=825"
            _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
            _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.namedParent=40"
            _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=58,.parent=13"
            _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=14"
            _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=15"
            _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=17"
            _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=50,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=1"
            _StyleDefs(53)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(6).Style:id=74,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(6).HeadingStyle:id=71,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(6).FooterStyle:id=72,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(6).EditorStyle:id=73,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(7).Style:id=62,.parent=13,.alignment=1"
            _StyleDefs(61)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(8).Style:id=66,.parent=13,.alignment=1"
            _StyleDefs(65)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(9).Style:id=70,.parent=13,.alignment=1"
            _StyleDefs(69)  =   "Splits(0).Columns(9).HeadingStyle:id=67,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(9).FooterStyle:id=68,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(9).EditorStyle:id=69,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(10).Style:id=78,.parent=13,.alignment=1"
            _StyleDefs(73)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=17"
            _StyleDefs(76)  =   "Named:id=33:Normal"
            _StyleDefs(77)  =   ":id=33,.parent=0"
            _StyleDefs(78)  =   "Named:id=34:Heading"
            _StyleDefs(79)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(80)  =   ":id=34,.wraptext=-1"
            _StyleDefs(81)  =   "Named:id=35:Footing"
            _StyleDefs(82)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(83)  =   "Named:id=36:Selected"
            _StyleDefs(84)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(85)  =   "Named:id=37:Caption"
            _StyleDefs(86)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(87)  =   "Named:id=38:HighlightRow"
            _StyleDefs(88)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(89)  =   "Named:id=39:EvenRow"
            _StyleDefs(90)  =   ":id=39,.parent=33,.bgcolor=&HE8E8E8&"
            _StyleDefs(91)  =   "Named:id=40:OddRow"
            _StyleDefs(92)  =   ":id=40,.parent=33,.bgcolor=&HFFFFFF&"
            _StyleDefs(93)  =   "Named:id=41:RecordSelector"
            _StyleDefs(94)  =   ":id=41,.parent=34"
            _StyleDefs(95)  =   "Named:id=42:FilterBar"
            _StyleDefs(96)  =   ":id=42,.parent=33"
         End
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   705
      Left            =   120
      Top             =   8250
      Width           =   15180
      _ExtentX        =   26776
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
         Left            =   2235
         TabIndex        =   51
         Top             =   180
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
         Picture         =   "mstStock.frx":0006
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   12450
         TabIndex        =   46
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
         Picture         =   "mstStock.frx":0290
      End
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   1185
         TabIndex        =   50
         Top             =   180
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
         Picture         =   "mstStock.frx":042F
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   435
         Left            =   120
         TabIndex        =   49
         Top             =   180
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
         Picture         =   "mstStock.frx":055B
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   13980
         TabIndex        =   48
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
         Picture         =   "mstStock.frx":0706
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   12900
         TabIndex        =   47
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
         Picture         =   "mstStock.frx":07AC
      End
   End
End
Attribute VB_Name = "mstStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB
Dim vaArray2 As New XArrayDB
Dim lEdit As Boolean
Dim nPos As SisPos

Private Sub GetEdit(lPar As Boolean)
  lEdit = lPar
  BiSAFrame1.Enabled = lPar
  SetButton cmdSimpan, cmdKeluar, cmdAdd, cmdEdit, cmdHapus, nPos, lPar, cmdAktivasi
End Sub

Private Sub GetLoadRows(Optional ByVal lShow = False)
Dim n As Integer

  vaArray.ReDim 0, 0, 0, 10
  If lShow = False Then
    Set dbData = objData.Browse(GetDSN, "stock s", "s.kodestock,s.barcode,s.nama,g.keterangan,s.jenis,s.hargabeli,s.hargajual,s.cogs,s.kodesatuan,s.stok", "s.kodestock", sisContent, TDBGrid1.Columns(1).FilterText, " AND barcode LIKE '%" & TDBGrid1.Columns(2).FilterText & "%' AND nama LIKE '%" & TDBGrid1.Columns(3).FilterText & "%' AND g.keterangan LIKE '%" & TDBGrid1.Columns(4).FilterText & "%' AND s.kodesatuan LIKE '%" & TDBGrid1.Columns(5).FilterText & "%' AND s.jenis < 9", "s.kodestock desc", Array("LEFT JOIN golongan g on g.kodegolongan = s.kodegolongan"), 0, 50)
  Else
    Set dbData = objData.Browse(GetDSN, "stock s", "s.kodestock,s.barcode,s.nama,g.keterangan,s.jenis,s.hargabeli,s.hargajual,s.cogs,s.kodesatuan,s.stok", "s.kodestock", sisContent, TDBGrid1.Columns(1).FilterText, " AND barcode LIKE '%" & TDBGrid1.Columns(2).FilterText & "%' AND nama LIKE '%" & TDBGrid1.Columns(3).FilterText & "%' AND g.keterangan LIKE '%" & TDBGrid1.Columns(4).FilterText & "%' AND s.kodesatuan LIKE '%" & TDBGrid1.Columns(5).FilterText & "%' AND s.jenis < 9", "s.kodestock desc", Array("LEFT JOIN golongan g on g.kodegolongan = s.kodegolongan"), 0, 50)
  End If
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = n
      vaArray(n, 1) = GetNull(dbData!KodeStock)
      vaArray(n, 2) = GetNull(dbData!barcode)
      vaArray(n, 3) = GetNull(dbData!nama)
      vaArray(n, 4) = GetNull(dbData!keterangan)
      vaArray(n, 5) = Format(GetNull(dbData!stok), "###,###,###,##0.00") 'GetSaldoStock(objData, "", vaArray(n, 1)) 'GetNull(dbData!kodesatuan)
      vaArray(n, 6) = GetNull(dbData!jenis, "1")
      vaArray(n, 7) = GetNull(dbData!hargabeli)
      vaArray(n, 8) = GetNull(dbData!HargaJual)
      vaArray(n, 9) = GetNull(dbData!cogs)
      vaArray(n, 10) = GetNull(dbData!kodesatuan)
      dbData.MoveNext
    Loop
'    vaArray.LoadRows dbData.GetRows(dbData.RecordCount)
  End If
  GetBerapaProfit
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
  TDBGrid1.Refresh
End Sub

Private Sub GetLoadRows2()
Dim n As Integer

  vaArray2.ReDim 0, 0, 0, 6
  Set dbData = objData.Browse(GetDSN, "stock s", "s.kodestock,s.barcode,s.nama,g.keterangan,s.kodesatuan,s.jenis", "s.kodestock", sisContent, TDBGrid2.Columns(1).FilterText, " AND barcode LIKE '%" & TDBGrid2.Columns(2).FilterText & "%' AND nama LIKE '%" & TDBGrid2.Columns(3).FilterText & "%' AND g.keterangan LIKE '%" & TDBGrid2.Columns(4).FilterText & "%' AND s.kodesatuan LIKE '%" & TDBGrid2.Columns(5).FilterText & "%' AND s.jenis = 9", "s.kodestock desc", Array("LEFT JOIN golongan g on g.kodegolongan = s.kodegolongan"))
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray2.InsertRows vaArray2.UpperBound(1) + 1
      n = vaArray2.UpperBound(1)
      vaArray2(n, 0) = n
      vaArray2(n, 1) = GetNull(dbData!KodeStock)
      vaArray2(n, 2) = GetNull(dbData!barcode)
      vaArray2(n, 3) = GetNull(dbData!nama)
      vaArray2(n, 4) = GetNull(dbData!keterangan)
      vaArray2(n, 5) = GetNull(dbData!kodesatuan)
      vaArray2(n, 6) = GetNull(dbData!jenis, "1")
      dbData.MoveNext
    Loop
  End If
  GetBerapaProfit
  Set TDBGrid2.Array = vaArray2
  TDBGrid2.ReBind
  TDBGrid2.Refresh
End Sub


Private Sub cBarcode_ButtonClick()
  If Len(Trim(cBarcode.Text)) >= 3 Then
    Set dbData = objData.Browse(GetDSN, "stock", "barcode,kodestock,nama", "barcode", sisContent, cBarcode.Text, , "barcode")
    If Not dbData.EOF Then
      cBarcode.Text = cBarcode.Browse(dbData)
      cKode.Text = GetNull(dbData!KodeStock, "")
      'GetMemory
      AssignData 0
    End If
  Else
    MsgBox "Ketikkan 3 atau lebih karakter dalam pencarian", vbInformation
  End If
End Sub

Private Sub cBarcode_Validate(Cancel As Boolean)
  If nPos = Add Then
    If Trim(cBarcode.Text) <> "" Then
      Set dbData = objData.Browse(GetDSN, "stock", "barcode", "barcode", sisAssign, cBarcode.Text)
      If Not dbData.EOF Then
        If MsgBox("Peringatan!!" & vbCrLf & "Barcode yg ingin anda masukkan sudah terpakai sebelumnya, silahkan gunakan Kode Barcode yg lain." & vbCrLf & "Apakah akan dilanjutkan?" & vbCrLf & "Jika Ya, maka akan ada Kode Barcode yg sama dalam komputer.", vbYesNo + vbInformation) = vbNo Then
          cBarcode.Text = ""
        End If
      End If
    End If
  End If
End Sub

Private Sub cGolongan_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "golongan", "kodegolongan,keterangan", "keterangan", sisContent, cGolongan.Text, , "kodegolongan,keterangan")
  If Not dbData.EOF Then
    If dbData.RecordCount > 1 Then
      cGolongan.Text = cGolongan.Browse(dbData)
    End If
    cGolongan.Text = GetNull(dbData!kodegolongan)
    cNamaGolongan.Text = GetNull(dbData!keterangan, "")
  Else
    cGolongan.Default
    cNamaGolongan.Default
  End If
End Sub

Private Function ValidKode() As Boolean
  ValidKode = True
End Function

Private Sub GetMemory()
Dim n As Single
   
  Set dbData = objData.Browse(GetDSN, "stock s", "s.nama,s.kodesatuan,sa.keterangan as keterangansatuan,s.hargabeli,s.hargajual,s.kodegolongan,ga.keterangan as keterangangolongan,s.jenis,s.statusnonaktif,s.cogs,s.diskonpenjualan,s.keterangan,s.autobiaya", "s.kodestock", sisAssign, cKode.Text, , , Array("left join satuan sa on sa.kodesatuan = s.kodesatuan", "left join golongan ga on ga.kodegolongan = s.kodegolongan"))
  If Not dbData.EOF Then
    cNama.Text = GetNull(dbData!nama, "")
    cSatuan.Text = GetNull(dbData!kodesatuan, "")
    cNamaSatuan.Text = GetNull(dbData!keterangansatuan, "")
    nHargaBeli.value = GetNull(dbData!hargabeli)
    nHargaJual.value = GetNull(dbData!HargaJual)
    cGolongan.Text = GetNull(dbData!kodegolongan, "")
    cNamaGolongan.Text = GetNull(dbData!KeteranganGolongan, "")
    chStatus.value = GetNull(dbData!statusnonaktif)
    SetOpt optJenisInventory, GetNull(dbData!jenis, "1")
    Text1.Text = GetNull(dbData!keterangan)
    nDiskonPenjualan.value = GetNull(dbData!diskonpenjualan)
    SetOpt optAutoBiaya, GetNull(dbData!autobiaya)
    If nPos = Delete Then
      DeleteStock
    End If
  End If
End Sub

Private Sub cKategori_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "kategori", "kategori,keterangan", "keterangan", sisContent, cKategori.Text, " or kategori like '%" & cKategori.Text & "'", "kategori,keterangan")
  If Not dbData.EOF Then
    If dbData.RecordCount > 1 Then
      cKategori.Text = cKategori.Browse(dbData)
    End If
    cKategori.Text = GetNull(dbData!kategori)
    cNamaKategori.Text = GetNull(dbData!keterangan, "")
  Else
    cKategori.Default
    cNamaKategori.Default
  End If
End Sub

Private Sub cKodeGroupSales_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "groupsales", "kode,keterangan", "keterangan", sisContent, cKodeGroupSales.Text, " and status = 1", "kode,keterangan")
  If Not dbData.EOF Then
    If dbData.RecordCount > 1 Then
      cKodeGroupSales.Text = cKodeGroupSales.Browse(dbData)
    End If
    cKodeGroupSales.Text = GetNull(dbData!Kode)
    cNamaGroupSales.Text = GetNull(dbData!keterangan, "")
  Else
    cKodeGroupSales.Default
    cNamaGroupSales.Default
  End If
End Sub

Private Sub cmdAdd_Click()
  GetEdit True
  initvalue
  cBarcode.SetFocus
  nPos = Add
  getDisableEnableJenisInventoryInput True
End Sub

Private Sub cmdAktivasi_Click()
  frmAktivasi.Action Me
End Sub

Private Sub cmdEdit_Click()
  GetEdit True
  TabOne1 = 0
  cBarcode.SetFocus
  nPos = Edit
End Sub

Private Sub cmdHapus_Click()
Dim objMenu As New CodeSuiteLibrary.Menu

  If aCfg(objData, msOtorisasiPenuh) = "Y" Then
    If nPos = Edit Or Delete Then
      If GetRegistry(reg_UserLevel) <> 0 Then
        If objMenu.GetPassword("", Me, GetDSN) Then
          If objMenu.UserLevel <> 0 Then
              MsgBox "Maaf, Anda tidak diberikan wewenang untuk melakukan PENGEDITAN." & vbCrLf & _
                     "Hanya user dengan LEVEL 0 (SUPERVISOR) yg diijinkan", vbInformation, "OTORISASI not ALLOWED"
              Exit Sub
'          Else
'            MsgBox "OTORISASI DIBATALKAN", vbCritical
'            Exit Sub
          End If
        Else
          Exit Sub
        End If
      End If
    End If
  End If
  
  GetEdit True
  nPos = Delete
  TabOne1 = 0
  cBarcode.SetFocus
  DeleteStock
  GetLoadRows
  GetLoadRows2
End Sub

Private Sub DeleteStock()
Dim cInfo As String
  
  If Trim(cKode.Text) <> "" Then
    cInfo = "Kode: " & cKode.Text & vbCrLf
    cInfo = cInfo & "Nama: " & cNama.Text & vbCrLf
    cInfo = cInfo & "Satuan: " & cSatuan.Text & vbCrLf
    cInfo = cInfo & "Golongan: " & cNamaGolongan.Text & vbCrLf
    cInfo = cInfo & "Harga Beli: " & Format(nHargaBeli.value, "###,###,###,##0.00") & vbCrLf
    cInfo = cInfo & "Harga Jual: " & Format(nHargaJual.value, "###,###,###,##0.00") & vbCrLf
    
    If MsgBox("Data Benar-benar Dihapus ?" & vbCrLf & vbCrLf & cInfo, vbQuestion + vbYesNo) = vbYes Then
      
      'Jika sedang digunakan pada kartu stock 'STOP
      If lExist(objData, "kartustock", "kodestock", cKode.Text) Then
        MsgBox "Maaf, data ini masih digunakan oleh sistem komputer" & vbCrLf & "Tidak bisa dihapus"
        Exit Sub
      End If
      
      If lExist(objData, "memberorder", "kodestock", cKode.Text) Then
        MsgBox "Maaf, data ini masih digunakan dalam transaksi order" & vbCrLf & "Tidak bisa dihapus"
        Exit Sub
      End If
      
      objData.Start GetDSN
      objData.Delete GetDSN, "stock", "kodestock", sisAssign, cKode.Text
      objData.Save GetDSN
    End If
    initvalue
    GetEdit False
  End If
  
End Sub

Private Function GetEditStock() As Boolean
Dim n As Single
Dim db As New ADODB.Recordset
  
  '*apabila stok sudah di set sebagai inventory dan sudah pernah terjadi transaksi
  '*maka tidak boleh di edit lagi jenis nya
  
  GetEditStock = True
  'cek dulu di kartu stock apakah stok ini sudah pernah terjadi transaksi
  Set dbData = objData.Browse(GetDSN, "kartustock", "kodestock", "kodestock", sisAssign, cKode.Text)
  If Not dbData.EOF Then
    'cek perubahan
      Set db = objData.Browse(GetDSN, "stock", "kodestock,jenis", "kodestock", sisAssign, cKode.Text)
      If Not db.EOF Then
        If GetNull(db!jenis) <> GetOpt(optJenisInventory) Then
          GetEditStock = False
          MsgBox "Stok ini tidak diperkenankan dirubah jenis nya", vbCritical, "error"
        End If
      End If
  End If
End Function

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
Dim nMin As Double
Dim nMax As Double
Dim n As Single
Dim cFaktur As String
Dim dTglAwal As Date
Dim lSave As Boolean
Dim objMenu As New CodeSuiteLibrary.Menu


  If aCfg(objData, msOtorisasiPenuh) = "Y" Then
    If nPos = Edit Then
      If GetRegistry(reg_UserLevel) <> 0 Then
        If objMenu.GetPassword("", Me, GetDSN) Then
          If objMenu.UserLevel <> 0 Then
              MsgBox "Maaf, Anda tidak diberikan wewenang untuk melakukan PENGEDITAN." & vbCrLf & _
                     "Hanya user dengan LEVEL 0 (SUPERVISOR) yg diijinkan", vbInformation, "OTORISASI not ALLOWED"
              Exit Sub
'          Else
'            MsgBox "OTORISASI DIBATALKAN", vbCritical
'            Exit Sub
          End If
        Else
          Exit Sub
        End If
      End If
    End If
  End If

lSave = True
  If ValidSaving Then
    objData.Start GetDSN
    
    ' Simpan table Stock
    vaField = Array("nama", _
                    "barcode", _
                    "kodegolongan", _
                    "kodesatuan", _
                    "hargabeli", _
                    "hargajual", _
                    "jenis", _
                    "asbiaya", _
                    "poin", _
                    "diskonpenjualan", _
                    "statusnonaktif", _
                    "bv", "groupsales", _
                    "kategori", "keterangan", "autobiaya", "konsi", "kodesupplier")
    vaValue = Array(cNama.Text, _
                    Trim(cBarcode.Text), _
                    cGolongan.Text, _
                    cSatuan.Text, _
                    nHargaBeli.value, _
                    nHargaJual.value, _
                    GetOpt(optJenisInventory), _
                    GetOpt(optBiaya), _
                    nPoin.value, _
                    nDiskonPenjualan.value, _
                    chStatus.value, _
                    0, cKodeGroupSales.Text, _
                    cKategori.Text, Text1.Text, GetOpt(optAutoBiaya), GetOpt(optKonsinyasi), cKodeSupplier.Text)
    
    If Trim(cKode.Text) <> "" Then
      lSave = IIf(lSave, objData.Update(GetDSN, "stock", "kodestock = '" & cKode.Text & "'", vaField, vaValue), False)
    Else
      lSave = IIf(lSave, objData.Add(GetDSN, "stock", vaField, vaValue), False)
    End If
    
    If Not lSave Then
      objData.Cancel GetDSN
    Else
      objData.Save GetDSN
    End If
    
    GetEdit False
    initvalue
    GetLoadRows
    GetLoadRows2
  End If
End Sub

Private Function ValidSaving() As Boolean
Dim db As New ADODB.Recordset

  ValidSaving = True
  Set db = objData.Browse(GetDSN, "golongan", "kodegolongan", "kodegolongan", sisAssign, cGolongan.Text)
  If db.EOF Then
    'Kodegolongan tidak ada dalam database, hentikan penyimpanan
    MsgBox "Kode Golongan tidak tertera dengan benar" & vbCrLf & "Data tidak bisa disimpan"
    ValidSaving = False
    Exit Function
  End If
  
  Set db = objData.Browse(GetDSN, "satuan", "kodesatuan", "kodesatuan", sisAssign, cSatuan.Text)
  If db.EOF Then
    'Kodesatuan tidak ada dalam database, hentikan penyimpanan
    MsgBox "Kode satuan tidak tertera dengan benar" & vbCrLf & "Data tidak bisa disimpan"
    ValidSaving = False
    Exit Function
  End If
  
  If nHargaBeli.value < 0 Or nHargaJual.value < 0 Then
    MsgBox "Harga beli atau harga jual tidak boleh kurang dari 0"
    ValidSaving = False
    Exit Function
  End If
  
  If nHargaBeli.value > nHargaJual.value Then
    If MsgBox("Yakin harga jual dibawa harga beli??", vbYesNo + vbCritical) = vbNo Then
      nHargaBeli.SetFocus
      ValidSaving = False
      Exit Function
    End If
  End If
  
  If NewGetValidJenis = False Then
    If MsgBox("Maaf penentuan jenis barang Inventory/Non masih belum benar?? ", vbCritical) Then
      ValidSaving = False
      Exit Function
    End If
  End If
  
  If nPos = Edit Then
    If GetEditStock = False Then
      ValidSaving = False
      Exit Function
    End If
  End If
  
  If GetValidDataBrowse(objData, "groupsales", "kode", cKodeGroupSales.Text) = False Then
    MsgBox "Maaf isian Group Sales belum lengkap", vbCritical
    Exit Function
  End If
  
  If GetValidDataBrowse(objData, "kategori", "kategori", cKategori.Text) = False Then
    MsgBox "Maaf isian Kategori belum lengkap", vbCritical
    Exit Function
  End If
  
  If GetValidDataBrowse(objData, "golongan", "kodegolongan", cGolongan.Text) = False Then
    MsgBox "Maaf isian Golongan belum lengkap", vbCritical
    Exit Function
  End If
End Function

Private Sub cNamaSupplier_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "supplier", "kodesupplier,nama", "kodesupplier", sisContent, cNamaSupplier.Text, " or nama like '%" & cNamaSupplier.Text & "%'")
  If Not dbData.EOF Then
    cNamaSupplier.Text = cNamaSupplier.Browse(dbData, Array("KODE", "NAMA"), , Array(10, 35))
    cKodeSupplier.Text = GetNull(dbData!kodesupplier)
    cNamaSupplier.Text = GetNull(dbData!nama)
  End If
End Sub

Private Sub cSatuan_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "satuan", "kodesatuan,keterangan", "kodesatuan", sisContent, cSatuan.Text, , "kodesatuan,keterangan")
  If Not dbData.EOF Then
    If dbData.RecordCount > 1 Then
      cSatuan.Text = cSatuan.Browse(dbData)
    End If
    cSatuan.Text = GetNull(dbData!kodesatuan)
    cNamaSatuan.Text = GetNull(dbData!keterangan, "")
  Else
    cSatuan.Default
    cNamaSatuan.Default
  End If
End Sub

Private Sub Form_Activate()
  'GetLoadRows False
End Sub

Private Sub Form_Load()
Dim n As Single
  
  SetIcon Me.hWnd, "SIKD"
  GetEdit False
  CenterForm Me
  initvalue
  
  TabIndex cKode, n
  TabIndex cBarcode, n
  TabIndex cNama, n
  TabIndex cNamaSupplier, n
  TabIndex cSatuan, n
  TabIndex cNamaSatuan, n
  TabIndex cGolongan, n
  TabIndex cKategori, n
  TabIndex cNamaKategori, n
  TabIndex cNamaGolongan, n
  TabIndex cKodeGroupSales, n
  TabIndex optKonsinyasi(0), n
  TabIndex optKonsinyasi(1), n
  TabIndex optJenisInventory(0), n
  TabIndex optJenisInventory(1), n
  TabIndex optBiaya(0), n
  TabIndex optBiaya(1), n
  TabIndex nHargaJual, n
  TabIndex nHargaBeli, n
  TabIndex nDiskonPenjualan, n
  TabIndex nPoin, n
  
  TabIndex chStatus, n
  
  'optional
  TabIndex cmdAdd, n
  TabIndex cmdEdit, n
  TabIndex cmdHapus, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  TabIndex cmdAktivasi, n
  GetLoadRows
  GetLoadRows2
End Sub

Private Sub initvalue()
  cKode.Default
  cBarcode.Default
  cNama.Default
  cSatuan.Default
  cNamaSatuan.Default
  cGolongan.Default
  cNamaGolongan.Default
  cKategori.Default
  cNamaKategori.Default
  cKodeGroupSales.Default
  cNamaGroupSales.Default
  Text1.Text = ""
  cKodeSupplier.Default
  cNamaSupplier.Default
  
  nHargaBeli.Default
  nHargaJual.Default
  nPoin.Default
  Label7.Caption = ""
  
  optKonsinyasi(0).value = True
  optJenisInventory(0).value = True
  optBiaya(1).value = True
  
  nHargaBeli.Enabled = True
  nHargaJual.Enabled = True
  TabOne1 = 0
  nDiskonPenjualan.Default
  chStatus.value = 0
  cKodeGroupSales.Enabled = True
End Sub

Private Sub nHargaBeli_Change()
On Error Resume Next
  GetBerapaProfit
End Sub

Private Sub GetBerapaProfit()
  On Error Resume Next
  Label7.Caption = "Profit : " & Format(nHargaJual.value - nHargaBeli.value, "###,###,##0") & " = " & Format(GetNull(((nHargaJual.value - nHargaBeli.value) / nHargaBeli.value) * 100), "###,###,##0") & "%"
End Sub

Private Sub nHargaJual_Change()
On Error Resume Next
  GetBerapaProfit
End Sub

Private Sub optAutoBiaya_Click(Index As Integer)
  If optAutoBiaya(0).value = True Then
    optBiaya(1).value = True
  End If
End Sub

Private Sub optAutoBiaya_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    keybd_event VK_TAB, 0, 0, 0
  End If
End Sub

Private Sub optAutoBiaya_Validate(Index As Integer, Cancel As Boolean)
  If optAutoBiaya(0).value = True Then
    optBiaya(1).value = True
  End If
End Sub

Private Sub optBiaya_Click(Index As Integer)
  If optBiaya(0).value = True Then
    optAutoBiaya(1).value = True
  End If
End Sub

Private Sub optBiaya_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    keybd_event VK_TAB, 0, 0, 0
  End If
End Sub

Private Sub optBiaya_Validate(Index As Integer, Cancel As Boolean)
  If optBiaya(0).value = True Then
    optJenisInventory(1).value = True
  End If
  If optAutoBiaya(0).value = True Then
    optBiaya(1).value = True
  End If
End Sub

Private Sub optJenisInventory_Click(Index As Integer)
  If optKonsinyasi(1).value = True Then
    optJenisInventory(0).value = True
  End If
End Sub

Private Sub optJenisInventory_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    keybd_event VK_TAB, 0, 0, 0
  End If

End Sub

Private Sub optJenisInventory_Validate(Index As Integer, Cancel As Boolean)
  If Index = 0 Then
    optBiaya(1).value = True
    optAutoBiaya(1).value = True
  End If
End Sub

Private Sub optKonsinyasi_Click(Index As Integer)
  Select Case Index
    Case 1
      optJenisInventory(0).value = True
    Case 0
      optJenisInventory(0).value = True
  End Select
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    If Trim(TDBGrid1.Columns(3).Text) = "" Then
      GetLoadRows False
    Else
      GetLoadRows True
    End If
  End If
End Sub

Private Function GetValidasiJenis() As Boolean
  GetValidasiJenis = False
  
  If optJenisInventory(0).value Then
    'Jika inventory
    If optAutoBiaya(1).value = True And optBiaya(1).value = True Then
      GetValidasiJenis = True
    Else
      GetValidasiJenis = False
    End If
  Else
    'jika non inventory
   If optJenisInventory(1).value = True Then
    If optAutoBiaya(0).value = True And optBiaya(0).value = True Then
      GetValidasiJenis = False
    Else
      GetValidasiJenis = True
    End If
   End If
   'Exit Function
  End If
  
  If optAutoBiaya(0).value = True Then
    If optBiaya(1).value = True Then
      GetValidasiJenis = True
    Else
      GetValidasiJenis = False
    End If
  End If
  
'  If optKonsinyasi(1).Value = True Then
'    If optJenisInventory(1).Value = True Then
'      GetValidasiJenis = True
'    Else
'      GetValidasiJenis = False
'    End If
'  End If
  
  
  If optKonsinyasi(1).value = True Then
    If optJenisInventory(0).value = True Then
      GetValidasiJenis = True
    Else
      GetValidasiJenis = False
    End If
  End If
  
End Function


Private Function NewGetValidJenis() As Boolean
  NewGetValidJenis = False
  If _
    SkeNario1 = True Or _
    SkeNario2 = True Or _
    SkeNario3 = True _
  Then
    NewGetValidJenis = True
  End If
End Function

Private Function SkeNario1() As Boolean
SkeNario1 = False

  'Jika Bukan Konsinyasi -> Non Inventory
  If _
    optKonsinyasi(0).value = True And _
    optJenisInventory(1).value = True And ( _
    (optAutoBiaya(0).value = True And optBiaya(1).value = True) Or (optAutoBiaya(1).value = True And optBiaya(0).value = True) Or (optAutoBiaya(1).value = True And optBiaya(1).value = True) _
    ) _
  Then
    SkeNario1 = True
  End If
  
End Function

Private Function SkeNario2() As Boolean
SkeNario2 = False

  'Jika Bukan Konsinyasi -> Inventory
  If _
    optKonsinyasi(0).value = True And _
    optJenisInventory(0).value = True _
  Then
    SkeNario2 = True
  End If
  
End Function

Private Function SkeNario3() As Boolean
SkeNario3 = False

  'Jika Konsinyasi -> Inventory
  If _
    optKonsinyasi(1).value = True And _
    optJenisInventory(0).value = True _
  Then
    SkeNario3 = True
  End If
  
End Function

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  If nPos <> Add Then
    AssignData 1
  End If
End Sub

Private Sub TDBGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    GetLoadRows2
  End If
End Sub

Private Sub TDBGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  AssignData 2
End Sub

Private Sub AssignData(ByVal nTab As Integer)
  
  If nTab = 1 Then
    Set dbData = objData.Browse(GetDSN, "stock s", "s.kodestock,s.barcode,s.nama,s.kodesatuan,sa.keterangan as keterangansatuan,s.hargabeli,s.hargajual,s.kodegolongan,ga.keterangan as keterangangolongan,s.jenis,s.asbiaya,s.poin,s.diskonpenjualan,s.statusnonaktif,s.cogs,s.kategori,ka.keterangan as namakategori,s.groupsales,gs.keterangan as namagroupsales,s.keterangan,s.autobiaya,s.konsi,s.kodesupplier,sp.nama as namasupplier", "s.kodestock", sisAssign, TDBGrid1.Columns(1).Text, , , Array("left join satuan sa on sa.kodesatuan = s.kodesatuan", "left join golongan ga on ga.kodegolongan = s.kodegolongan", "left join kategori ka on ka.kategori = s.kategori", "left join groupsales gs on gs.kode = s.groupsales", "left join supplier sp on sp.kodesupplier = s.kodesupplier"))
  ElseIf nTab = 2 Then
    Set dbData = objData.Browse(GetDSN, "stock s", "s.kodestock,s.barcode,s.nama,s.kodesatuan,sa.keterangan as keterangansatuan,s.hargabeli,s.hargajual,s.kodegolongan,ga.keterangan as keterangangolongan,s.jenis,s.asbiaya,s.poin,s.diskonpenjualan,s.statusnonaktif,s.cogs,s.kategori,ka.keterangan as namakategori,s.groupsales,gs.keterangan as namagroupsales,s.keterangan,s.autobiaya,s.konsi,s.kodesupplier,sp.nama as namasupplier", "s.kodestock", sisAssign, TDBGrid2.Columns(1).Text, , , Array("left join satuan sa on sa.kodesatuan = s.kodesatuan", "left join golongan ga on ga.kodegolongan = s.kodegolongan", "left join kategori ka on ka.kategori = s.kategori", "left join groupsales gs on gs.kode = s.groupsales", "left join supplier sp on sp.kodesupplier = s.kodesupplier"))
  ElseIf nTab = 0 Then
    Set dbData = objData.Browse(GetDSN, "stock s", "s.kodestock,s.barcode,s.nama,s.kodesatuan,sa.keterangan as keterangansatuan,s.hargabeli,s.hargajual,s.kodegolongan,ga.keterangan as keterangangolongan,s.jenis,s.asbiaya,s.poin,s.diskonpenjualan,s.statusnonaktif,s.cogs,s.kategori,ka.keterangan as namakategori,s.groupsales,gs.keterangan as namagroupsales,s.keterangan,s.autobiaya,s.konsi,s.kodesupplier,sp.nama as namasupplier", "s.kodestock", sisAssign, cKode.Text, , , Array("left join satuan sa on sa.kodesatuan = s.kodesatuan", "left join golongan ga on ga.kodegolongan = s.kodegolongan", "left join kategori ka on ka.kategori = s.kategori", "left join groupsales gs on gs.kode = s.groupsales", "left join supplier sp on sp.kodesupplier = s.kodesupplier"))
  End If
  
  If Not dbData.EOF Then
    cKode.Text = GetNull(dbData!KodeStock)
    cBarcode.Text = GetNull(dbData!barcode)
    cNama.Text = GetNull(dbData!nama, "")
    cSatuan.Text = GetNull(dbData!kodesatuan, "")
    cNamaSatuan.Text = GetNull(dbData!keterangansatuan, "")
    nHargaBeli.value = GetNull(dbData!hargabeli)
    nHargaJual.value = GetNull(dbData!HargaJual)
    cGolongan.Text = GetNull(dbData!kodegolongan, "")
    cNamaGolongan.Text = GetNull(dbData!KeteranganGolongan, "")
    SetOpt optJenisInventory, GetNull(dbData!jenis)
    SetOpt optBiaya, GetNull(dbData!asbiaya)
    nPoin.value = GetNull(dbData!poin)
    chStatus.value = GetNull(dbData!statusnonaktif)
    nDiskonPenjualan.value = GetNull(dbData!diskonpenjualan)
    cKategori.Text = GetNull(dbData!kategori)
    cNamaKategori.Text = GetNull(dbData!namakategori)
    cKodeGroupSales.Text = GetNull(dbData!GroupSales)
    GetOnOffGroupSales cKodeGroupSales.Text
    cNamaGroupSales.Text = GetNull(dbData!namagroupsales)
    Text1.Text = GetNull(dbData!keterangan)
    SetOpt optAutoBiaya, GetNull(dbData!autobiaya)
    SetOpt optKonsinyasi, GetNull(dbData!konsi)
    cKodeSupplier.Text = GetNull(dbData!kodesupplier)
    cNamaSupplier.Text = GetNull(dbData!namasupplier)
    
    '******************************************************************************
    'cari stock ini apakah sudah pernah digunakan transaksi (kartustock)
    'jika sudah disable opsi untuk mengganti jenis barang (invent or non invent)
    '******************************************************************************
    getDisableEnableJenisInventoryInput True
    
    Set dbData = objData.Browse(GetDSN, "kartustock", "kodestock", "kodestock", sisAssign, cKode.Text, , , , 0, 1)
    If Not dbData.EOF Then
      'jika stock sudah pernah di transaksikan
      getDisableEnableJenisInventoryInput False
    End If
  End If
End Sub

Private Sub getDisableEnableJenisInventoryInput(ByVal lOption As Boolean)
Dim enColor As SystemColorConstants

    If lOption = False Then
      enColor = vbButtonFace
    Else
      enColor = vbHighlightText
    End If
    
    BiSAFrame6.Enabled = lOption
    BiSAFrame4.Enabled = lOption
    BiSAFrame4.BackColor = enColor
    optJenisInventory(0).BackColor = enColor
    optJenisInventory(1).BackColor = enColor
    
    If optJenisInventory(1).value = True Then
      'khusus non inventory
      enColor = vbHighlightText
      lOption = True
    End If
    
    BiSAFrame2.Enabled = lOption
    BiSAFrame2.BackColor = enColor
    BiSAFrame5.BackColor = enColor
    BiSAFrame5.Enabled = lOption
    optBiaya(0).BackColor = enColor
    optBiaya(1).BackColor = enColor
    optAutoBiaya(0).BackColor = enColor
    optAutoBiaya(1).BackColor = enColor

    cNamaSupplier.Enabled = lOption
End Sub

Private Sub GetOnOffGroupSales(ByVal KodeGroupSales As String)
Dim dba As New ADODB.Recordset

  cKodeGroupSales.Enabled = True
  Set dba = objData.Browse(GetDSN, "groupsales", "kode,status", "kode", sisAssign, KodeGroupSales)
  If Not dba.EOF Then
    If GetNull(dba!Status) = 0 Then
      cKodeGroupSales.Enabled = False
    End If
  End If
End Sub

