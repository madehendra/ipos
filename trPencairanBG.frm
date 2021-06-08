VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trPencairanBG 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pencairan BG"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   11880
   Begin TrueOleDBGrid70.TDBDropDown TDBDropDown1 
      Height          =   2340
      Left            =   7245
      TabIndex        =   9
      Top             =   1845
      Width           =   3360
      _ExtentX        =   5927
      _ExtentY        =   4128
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Kode Akun"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Keterangan"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   873
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   15790320
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2593"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2514"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=3149"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3069"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits.Count    =   1
      AllowRowSizing  =   0   'False
      Appearance      =   1
      BorderStyle     =   1
      ColumnHeaders   =   -1  'True
      DataMode        =   4
      DefColWidth     =   0
      Enabled         =   -1  'True
      HeadLines       =   1
      RowDividerStyle =   2
      LayoutName      =   ""
      LayoutFileName  =   ""
      LayoutURL       =   ""
      EmptyRows       =   0   'False
      ListField       =   ""
      DataField       =   ""
      IntegralHeight  =   0   'False
      FetchRowStyle   =   0   'False
      AlternatingRowStyle=   0   'False
      DataMember      =   ""
      ColumnFooters   =   0   'False
      FootLines       =   1
      DeadAreaBackColor=   15790320
      ValueTranslate  =   0   'False
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=192,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
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
      _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(40)  =   "Named:id=33:Normal"
      _StyleDefs(41)  =   ":id=33,.parent=0"
      _StyleDefs(42)  =   "Named:id=34:Heading"
      _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(44)  =   ":id=34,.wraptext=-1"
      _StyleDefs(45)  =   "Named:id=35:Footing"
      _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(47)  =   "Named:id=36:Selected"
      _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(49)  =   "Named:id=37:Caption"
      _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(51)  =   "Named:id=38:HighlightRow"
      _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=39:EvenRow"
      _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(55)  =   "Named:id=40:OddRow"
      _StyleDefs(56)  =   ":id=40,.parent=33"
      _StyleDefs(57)  =   "Named:id=41:RecordSelector"
      _StyleDefs(58)  =   ":id=41,.parent=34"
      _StyleDefs(59)  =   "Named:id=42:FilterBar"
      _StyleDefs(60)  =   ":id=42,.parent=33"
   End
   Begin BiSADateProject.BiSADate dJatuhTempo 
      Height          =   330
      Left            =   2250
      TabIndex        =   6
      Top             =   240
      Width           =   1410
      _ExtentX        =   2487
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
   Begin BiSAButtonProject.BiSAButton cmdFind 
      Height          =   705
      Left            =   4245
      TabIndex        =   5
      Top             =   240
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1244
      Caption         =   "FIND"
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
   Begin VB.CheckBox chkMember 
      Caption         =   "Member"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   195
      TabIndex        =   4
      Top             =   675
      Width           =   975
   End
   Begin VB.CheckBox chkJatuhTempo 
      Caption         =   "Tanggal Jatuh Tempo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   195
      TabIndex        =   3
      Top             =   300
      Width           =   1980
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame4 
      Height          =   630
      Left            =   150
      Top             =   5745
      Width           =   11565
      _ExtentX        =   20399
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
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   10260
         TabIndex        =   0
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
         Picture         =   "trPencairanBG.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   9180
         TabIndex        =   1
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
         Picture         =   "trPencairanBG.frx":00A6
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   4710
      Left            =   165
      TabIndex        =   2
      Top             =   1020
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   8308
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   4
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "REF"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "MEMBER"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "BG ID"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "NOMINAL"
      Columns(4).DataField=   ""
      Columns(4).NumberFormat=   "###,###,###,##0.00"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "JATUH TEMPO"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "AKUN/BANK"
      Columns(6).DataField=   ""
      Columns(6).DropDown=   "TDBDropDown1"
      Columns(6).DropDown.vt=   8
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "KETERANGAN"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "NOMOR PELUNASAN PIUTANG"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   9
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=9"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=609"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=529"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=197124"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=3916"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3836"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=197120"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=3254"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=3175"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=197120"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=2381"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2302"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=197121"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=3731"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=3651"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=197122"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=2831"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2752"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=197122"
      Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(31)=   "Column(6).Width=2910"
      Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2831"
      Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=197120"
      Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(36)=   "Column(7).Width=4207"
      Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=4128"
      Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=197120"
      Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(41)=   "Column(8).Width=2725"
      Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=2646"
      Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=197124"
      Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      Appearance      =   0
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.alignment=1,.bold=0,.fontsize=825"
      _StyleDefs(15)  =   ":id=3,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
      _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
      _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
      _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
      _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
      _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.alignment=0"
      _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
      _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
      _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=54,.parent=13,.alignment=0"
      _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
      _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
      _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
      _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=32,.parent=13,.alignment=2"
      _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
      _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
      _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
      _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=70,.parent=13,.alignment=1"
      _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=67,.parent=14"
      _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=68,.parent=15"
      _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=69,.parent=17"
      _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=1"
      _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
      _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
      _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
      _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=74,.parent=13,.alignment=0"
      _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=71,.parent=14"
      _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=72,.parent=15"
      _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=73,.parent=17"
      _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=50,.parent=13,.alignment=0"
      _StyleDefs(66)  =   "Splits(0).Columns(7).HeadingStyle:id=47,.parent=14"
      _StyleDefs(67)  =   "Splits(0).Columns(7).FooterStyle:id=48,.parent=15"
      _StyleDefs(68)  =   "Splits(0).Columns(7).EditorStyle:id=49,.parent=17"
      _StyleDefs(69)  =   "Splits(0).Columns(8).Style:id=62,.parent=13"
      _StyleDefs(70)  =   "Splits(0).Columns(8).HeadingStyle:id=59,.parent=14"
      _StyleDefs(71)  =   "Splits(0).Columns(8).FooterStyle:id=60,.parent=15"
      _StyleDefs(72)  =   "Splits(0).Columns(8).EditorStyle:id=61,.parent=17"
      _StyleDefs(73)  =   "Named:id=33:Normal"
      _StyleDefs(74)  =   ":id=33,.parent=0"
      _StyleDefs(75)  =   "Named:id=34:Heading"
      _StyleDefs(76)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(77)  =   ":id=34,.wraptext=-1"
      _StyleDefs(78)  =   "Named:id=35:Footing"
      _StyleDefs(79)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(80)  =   "Named:id=36:Selected"
      _StyleDefs(81)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(82)  =   "Named:id=37:Caption"
      _StyleDefs(83)  =   ":id=37,.parent=34,.alignment=2,.bold=0,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(84)  =   ":id=37,.strikethrough=0,.charset=0"
      _StyleDefs(85)  =   ":id=37,.fontname=Tahoma"
      _StyleDefs(86)  =   "Named:id=38:HighlightRow"
      _StyleDefs(87)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(88)  =   "Named:id=39:EvenRow"
      _StyleDefs(89)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(90)  =   "Named:id=40:OddRow"
      _StyleDefs(91)  =   ":id=40,.parent=33"
      _StyleDefs(92)  =   "Named:id=41:RecordSelector"
      _StyleDefs(93)  =   ":id=41,.parent=34"
      _StyleDefs(94)  =   "Named:id=42:FilterBar"
      _StyleDefs(95)  =   ":id=42,.parent=33"
   End
   Begin BiSATextBoxProject.BiSABrowse cMember 
      Height          =   330
      Left            =   2250
      TabIndex        =   7
      Top             =   615
      Width           =   1935
      _ExtentX        =   3413
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
   Begin BiSADateProject.BiSADate dTgl 
      Height          =   330
      Left            =   8460
      TabIndex        =   8
      Top             =   210
      Width           =   3045
      _ExtentX        =   5371
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
      Caption         =   "Tanggal Cair"
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
Attribute VB_Name = "trPencairanBG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB
Dim vaLoad As New XArrayDB

Private Sub cmdFind_Click()
  initvalue
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdSimpan_Click()
Dim n As Integer
Dim db As New ADODB.Recordset
Dim lSave As Boolean
  
  lSave = True
  objData.Start GetDSN
    
    For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
      If vaArray(n, 0) = -1 And lCekAkun(vaArray(n, 6)) Then
        'Simpan
        'tableid, nomorpelunasanpiutang, kodeakun, username, date, datetime, keterangan
        lSave = IIf(lSave, objData.Delete(GetDSN, "pencairanbg", "tableid", sisAssign, vaArray(n, 3)), False)
        lSave = IIf(lSave, objData.Add(GetDSN, "pencairanbg", Array("tableid", "nomorpelunasanpiutang", "kodeakun", "username", "date", "datetime", "keterangan"), Array(vaArray(n, 3), vaArray(n, 8), vaArray(n, 6), GetRegistry(reg_Username), Format(Now, "yyyy-MM-dd"), SNow, vaArray(n, 7))), False)
        
        'posting ke bukubesar
        'KasBank
          'BG
        lSave = IIf(lSave, UpdKodeTr(objData, msPencairanBG, vaArray(n, 3), Format(dTgl.value, "yyyy-MM-dd"), vaArray(n, 6), GetCostCenterUser(objData, GetRegistry(reg_Username)), "Pencairan BG reff no " & vaArray(n, 1), vaArray(n, 4), 0), False)
            lSave = IIf(lSave, UpdKodeTr(objData, msPencairanBG, vaArray(n, 3), Format(dTgl.value, "yyyy-MM-dd"), aCfg(objData, msRekeningBG), GetCostCenterUser(objData, GetRegistry(reg_Username)), "Pencairan BG reff no " & vaArray(n, 1), 0, vaArray(n, 4)), False)
            
      End If
    Next n
  
  If lSave Then
    objData.Save GetDSN
  Else
    objData.Cancel GetDSN
  End If
  initvalue
End Sub

Private Function lCekAkun(cAkunRek As String) As Boolean
Dim db As New ADODB.Recordset

lCekAkun = False
Set db = objData.Browse(GetDSN, "akun", , "kodeakun", sisAssign, cAkunRek)
If Not db.EOF Then
  lCekAkun = True
End If
End Function

Private Sub cMember_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "anggota", "kodeanggota,nama", "nama", sisContent, cMember.Text)
  If Not dbData.EOF Then
    cMember.Text = cMember.Browse(dbData)
  End If
End Sub

Private Sub Form_Load()
Dim n As Single
Dim db As New ADODB.Recordset

  
  CenterForm Me
  SetIcon Me.hWnd
  
  TabIndex dTgl, n
  TabIndex chkJatuhTempo, n
  TabIndex dJatuhTempo, n
  TabIndex chkMember, n
  TabIndex cMember, n
  TabIndex cmdFind, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  
  'Load dropdown table
  
  vaLoad.ReDim 0, -1, 0, 1
  Set db = objData.Browse(GetDSN, "akun", "kodeakun,keterangan", "kodeakun", sisPrefix, "1", " and jenis = 'D'")
  If Not db.EOF Then
    Do While Not db.EOF
      vaLoad.InsertRows vaLoad.UpperBound(1) + 1
      n = vaLoad.UpperBound(1)
      vaLoad(n, 0) = GetNull(db!kodeakun)
      vaLoad(n, 1) = GetNull(db!keterangan)
      db.MoveNext
    Loop
  End If
  Set TDBDropDown1.Array = vaLoad
  TDBDropDown1.ReBind
  TDBDropDown1.Refresh
   
  initvalue
End Sub

Private Sub initvalue()
Dim db As New ADODB.Recordset
Dim n As Integer
Dim cSQL As String
  
  vaArray.ReDim 0, -1, 0, 8
  cSQL = " 1=1 "
  If chkJatuhTempo.value = 1 Then
    cSQL = cSQL & " and b.jatuhtempo = '" & Format(dJatuhTempo.value, "yyyy-MM-dd") & "'"
  End If
  If chkMember.value = 1 Then
    cSQL = cSQL & " and a.kodeanggota = '" & cMember.Text & "'"
  End If
  Set db = objData.Browse(GetDSN, "bg b", "b.nomorpelunasanpiutang,b.reff,a.nama,b.tableid,b.jumlah,b.jatuhtempo", , , , cSQL, , Array("left join totpelunasanpiutang t on t.nomorpelunasanpiutang = b.nomorpelunasanpiutang", "left join anggota a on a.kodeanggota = t.kodeanggota"))
  If Not db.EOF Then
    Do While Not db.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = 0
      vaArray(n, 1) = GetNull(db!reff)
      vaArray(n, 2) = GetNull(db!nama)
      vaArray(n, 3) = GetNull(db!tableid)
      vaArray(n, 4) = GetNull(db!jumlah)
      vaArray(n, 5) = Format(GetNull(db!jatuhtempo), "dd-MM-yyyy")
      vaArray(n, 6) = ""
      vaArray(n, 7) = ""
      vaArray(n, 8) = GetNull(db!nomorpelunasanpiutang)
    
      If lCekPencairan(vaArray(n, 3)) Then
        vaArray.DeleteRows n
      End If
      db.MoveNext
    Loop
  End If
  
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
  TDBGrid1.Refresh

End Sub

Private Function lCekPencairan(tableid) As Boolean
Dim db As New ADODB.Recordset

lCekPencairan = False
  Set db = objData.Browse(GetDSN, "pencairanbg", , "tableid", sisAssign, tableid)
  If Not db.EOF Then
    lCekPencairan = True
  End If
End Function

Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
  TDBGrid1.Update
End Sub

Private Sub tdbgrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Cancel = True

  Select Case ColIndex
    Case 0
      Cancel = False
    Case 6
      Cancel = False
      TDBGrid1.Update
    Case 7
      Cancel = False
  End Select
End Sub

Private Sub TDBGrid1_Validate(Cancel As Boolean)
  TDBGrid1.Update
End Sub
