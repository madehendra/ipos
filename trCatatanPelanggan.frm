VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trCatatanPelanggan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catatan Pelanggan"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13200
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   13200
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   6495
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   13020
      _cx             =   22966
      _cy             =   11456
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
      Appearance      =   2
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   -2147483633
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   "Input Catatan Baru|Catatan Pelanggan Lama"
      Align           =   0
      CurrTab         =   0
      FirstTab        =   0
      Style           =   3
      Position        =   0
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   0
      BorderWidth     =   0
      BoldCurrent     =   0   'False
      DogEars         =   -1  'True
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      TabCaptionPos   =   4
      TabPicturePos   =   0
      CaptionEmpty    =   ""
      Separators      =   0   'False
      Begin VB.Frame Frame2 
         Height          =   6120
         Left            =   45
         TabIndex        =   1
         Top             =   330
         Width           =   12930
         Begin VB.TextBox cKeterangan 
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
            Height          =   2595
            Left            =   6165
            MultiLine       =   -1  'True
            TabIndex        =   15
            Text            =   "trCatatanPelanggan.frx":0000
            Top             =   2415
            Width           =   6240
         End
         Begin BiSAFramProject.BiSAFrame BiSAFrame2 
            Height          =   1050
            Left            =   165
            Top             =   255
            Width           =   12255
            _ExtentX        =   21616
            _ExtentY        =   1852
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
            Begin BiSATextBoxProject.BiSABrowse cNamaAnggota 
               Height          =   330
               Left            =   3420
               TabIndex        =   8
               Top             =   585
               Width           =   4125
               _ExtentX        =   7276
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
            Begin BiSATextBoxProject.BiSATextBox cKodeAnggota 
               Height          =   330
               Left            =   1335
               TabIndex        =   6
               Top             =   585
               Width           =   2055
               _ExtentX        =   3625
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
            Begin BiSATextBoxProject.BiSATextBox cKodeSalesman 
               Height          =   330
               Left            =   1320
               TabIndex        =   3
               Top             =   165
               Width           =   2055
               _ExtentX        =   3625
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
            Begin BiSATextBoxProject.BiSATextBox cNamaSalesman 
               Height          =   330
               Left            =   3405
               TabIndex        =   5
               Top             =   165
               Width           =   4140
               _ExtentX        =   7303
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
               Left            =   9180
               TabIndex        =   2
               Top             =   135
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
               Caption         =   "Date"
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
            Begin VB.Label Label8 
               Caption         =   "Salesman"
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
               Left            =   270
               TabIndex        =   4
               Top             =   180
               Width           =   1110
            End
            Begin VB.Label Label6 
               Caption         =   "Pelanggan"
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
               Left            =   270
               TabIndex        =   7
               Top             =   615
               Width           =   1110
            End
         End
         Begin BiSATextBoxProject.BiSATextBox cSubject 
            Height          =   330
            Left            =   6150
            TabIndex        =   13
            Top             =   2055
            Width           =   6240
            _ExtentX        =   11007
            _ExtentY        =   582
            Text            =   "12345678901234567890123456789012345678901234567890"
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
            Height          =   4305
            Left            =   255
            TabIndex        =   9
            Top             =   1545
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   7594
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
            Columns(1).Caption=   "Tgl"
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "ID"
            Columns(2).DataField=   ""
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Subject"
            Columns(3).DataField=   ""
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   4
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=4"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=926"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=847"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=514"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=2143"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2064"
            Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=1720"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1640"
            Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(16)=   "Column(3).Width=2805"
            Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2725"
            Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
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
            _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=98,.parent=13"
            _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=95,.parent=14"
            _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=96,.parent=15"
            _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=97,.parent=17"
            _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
            _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
            _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
            _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
            _StyleDefs(53)  =   "Named:id=33:Normal"
            _StyleDefs(54)  =   ":id=33,.parent=0"
            _StyleDefs(55)  =   "Named:id=34:Heading"
            _StyleDefs(56)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(57)  =   ":id=34,.wraptext=-1"
            _StyleDefs(58)  =   "Named:id=35:Footing"
            _StyleDefs(59)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(60)  =   "Named:id=36:Selected"
            _StyleDefs(61)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(62)  =   "Named:id=37:Caption"
            _StyleDefs(63)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(64)  =   "Named:id=38:HighlightRow"
            _StyleDefs(65)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(66)  =   "Named:id=39:EvenRow"
            _StyleDefs(67)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(68)  =   "Named:id=40:OddRow"
            _StyleDefs(69)  =   ":id=40,.parent=33"
            _StyleDefs(70)  =   "Named:id=41:RecordSelector"
            _StyleDefs(71)  =   ":id=41,.parent=34"
            _StyleDefs(72)  =   "Named:id=42:FilterBar"
            _StyleDefs(73)  =   ":id=42,.parent=33"
         End
         Begin BiSATextBoxProject.BiSATextBox cID 
            Height          =   330
            Left            =   6150
            TabIndex        =   11
            Top             =   1680
            Width           =   2055
            _ExtentX        =   3625
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
         Begin BiSAFramProject.BiSAFrame BiSAFrame1 
            Height          =   630
            Left            =   4740
            Top             =   5235
            Width           =   8055
            _ExtentX        =   14208
            _ExtentY        =   1111
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
               Left            =   4620
               TabIndex        =   16
               TabStop         =   0   'False
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
               Picture         =   "trCatatanPelanggan.frx":0006
            End
            Begin BiSAButtonProject.BiSAButton cmdAktivasi 
               Height          =   435
               Left            =   4170
               TabIndex        =   17
               TabStop         =   0   'False
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
               Picture         =   "trCatatanPelanggan.frx":0290
            End
            Begin BiSAButtonProject.BiSAButton cmdKeluar 
               Height          =   435
               Left            =   6855
               TabIndex        =   19
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
               Picture         =   "trCatatanPelanggan.frx":042F
            End
            Begin BiSAButtonProject.BiSAButton cmdSimpan 
               Height          =   435
               Left            =   5775
               TabIndex        =   18
               TabStop         =   0   'False
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
               Picture         =   "trCatatanPelanggan.frx":04D5
            End
            Begin BiSAButtonProject.BiSAButton cmdAdd 
               Height          =   435
               Left            =   3105
               TabIndex        =   20
               TabStop         =   0   'False
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
               Picture         =   "trCatatanPelanggan.frx":075B
            End
         End
         Begin VB.Label Label7 
            Caption         =   "Nomor"
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
            Left            =   4920
            TabIndex        =   10
            Top             =   1725
            Width           =   1110
         End
         Begin VB.Label Label4 
            Caption         =   "Subject"
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
            Left            =   4905
            TabIndex        =   12
            Top             =   2115
            Width           =   1110
         End
         Begin VB.Label Label3 
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
            Height          =   270
            Left            =   4920
            TabIndex        =   14
            Top             =   2490
            Width           =   1005
         End
      End
   End
End
Attribute VB_Name = "trCatatanPelanggan"
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

Private Sub cmdAdd_Click()
  cID.Default
  cSubject.Default
  cKeterangan.Text = ""
  cSubject.SetFocus
End Sub

Private Sub cmdHapus_Click()
  If MsgBox("Apakah data akan dihapus?", vbYesNo) = vbYes Then
    objData.Delete GetDSN, "catatansalesman", "id", sisAssign, TDBGrid1.Columns(2).value
    GetDataGrid
  End If
End Sub

Private Sub cmdKeluar_Click()
  If lEdit Then
    initvalue
  Else
    Unload Me
  End If
End Sub

Private Function ValidSaving() As Boolean
Dim n As Integer

  ValidSaving = True
  
  If Trim(cKodeAnggota.Text) = "" Or Trim(cKodeSalesman.Text) = "" Then
    MsgBox "Kode Sales atau Pelanggan Kosong ! Data tidak bisa disimpan", vbCritical
    ValidSaving = False
    Exit Function
  End If
End Function

Private Sub cmdSimpan_Click()
Dim lSave As Boolean
Dim Faktur As String
Dim n As Integer
lSave = True

  'Simpan di table kontrakstock
  
  If ValidSaving Then
    objData.Start GetDSN
    lSave = IIf(lSave, objData.Update(GetDSN, "catatansalesman", "id='" & cID.Text & "'", Array("kodeanggota", "kodesalesman", "judul", "keterangan", "tgl", "username", "datetime"), Array(cKodeAnggota.Text, cKodeSalesman.Text, cSubject.Text, cKeterangan.Text, Format(dTgl.value, "yyyy-MM-dd"), GetRegistry(reg_Username), SNow)), False)
    If lSave Then
      objData.Save GetDSN
      MsgBox "Data berhasil disimpan", vbInformation
    Else
      objData.Cancel GetDSN
    End If
    initvalue
  End If
End Sub

Private Sub GetDataGrid()
Dim db As New ADODB.Recordset
Dim n As Single

  Set db = objData.Browse(GetDSN, "catatansalesman", "id,tgl,judul", "kodeanggota", sisAssign, cKodeAnggota.Text, " and kodesalesman = '" & cKodeSalesman.Text & "'", "id desc")
  If Not db.EOF Then
    vaArray.ReDim 0, -1, 0, 3
    Do While Not db.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = n + 1
      vaArray(n, 1) = GetNull(db!tgl)
      vaArray(n, 2) = GetNull(db!ID)
      vaArray(n, 3) = GetNull(db!judul)
      db.MoveNext
    Loop
    Set TDBGrid1.Array = vaArray
    TDBGrid1.ReBind
    TDBGrid1.Refresh
    Me.Refresh
  End If
End Sub

Private Sub GetDataCatatan(cNoId As String)
Dim db As New ADODB.Recordset
Dim n As Single

  cID.Default
  cSubject.Default
  cKeterangan.Text = ""
  Set db = objData.Browse(GetDSN, "catatansalesman", "id,judul,keterangan", "id", sisAssign, cNoId)
  If Not db.EOF Then
    cID.Text = GetNull(db!ID)
    cSubject.Text = GetNull(db!judul)
    cKeterangan.Text = GetNull(db!keterangan)
    Me.Refresh
  End If
End Sub

Private Sub cNamaanggota_ButtonClick()

Set dbData = objData.Browse(GetDSN, "anggota", "kodeanggota,nama,alamat,telp", "nama", sisContent, cNamaAnggota.Text, , "kodeanggota,nama")
  If Not dbData.EOF Then
    cNamaAnggota.Text = cNamaAnggota.Browse(dbData)
    cKodeAnggota.Text = GetNull(dbData!kodeanggota)
    cNamaAnggota.Text = GetNull(dbData!nama)
    GetDataGrid
  End If
End Sub

Private Sub cNamaAnggota_Validate(Cancel As Boolean)
  cNamaAnggota.Enabled = False
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  cNamaAnggota.Enabled = True
  initvalue
  TabIndex dTgl, n
  TabIndex cNamaAnggota, n
  TabIndex cSubject, n
  TabIndex cKeterangan, n
  TabIndex cmdKeluar, n
End Sub

Private Sub initvalue()
  dTgl.value = Date
  cKodeSalesman.Text = ""
  cNamaSalesman.Text = ""
  cID.Default
  cSubject.Text = ""
  cKeterangan.Text = ""
  Set dbData = objData.Browse(GetDSN, "akunkas u", "u.kodesalesman,s.nama", "u.username", sisAssign, GetRegistry(reg_Username), , , Array("left join salesman s on s.kodesalesman = u.kodesalesman"))
  If Not dbData.EOF Then
    cKodeSalesman.Text = GetNull(dbData!kodesalesman)
    cNamaSalesman.Text = GetNull(dbData!nama)
  End If
  GetDataGrid
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
  GetDataCatatan TDBGrid1.Columns(2).value
End Sub
