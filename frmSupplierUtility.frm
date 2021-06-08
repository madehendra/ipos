VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmSupplierUtility 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Utility for Supplier"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11820
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   11820
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   6315
      Left            =   -15
      Top             =   0
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   11139
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
      Begin BiSATextBoxProject.BiSATextBox cMember 
         Height          =   345
         Left            =   6255
         TabIndex        =   0
         Top             =   105
         Width           =   4875
         _ExtentX        =   8599
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
         Height          =   1845
         Left            =   60
         TabIndex        =   1
         Top             =   465
         Width           =   11715
         _ExtentX        =   20664
         _ExtentY        =   3254
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
         Columns(1).Caption=   "NAMA"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "ALAMAT"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   873
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   15790320
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2672"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2593"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=512"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=5265"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=5186"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=512"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=5741"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=5662"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=512"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
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
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=0"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=54,.parent=13,.alignment=0"
         _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=0"
         _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Named:id=33:Normal"
         _StyleDefs(45)  =   ":id=33,.parent=0"
         _StyleDefs(46)  =   "Named:id=34:Heading"
         _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(48)  =   ":id=34,.wraptext=-1"
         _StyleDefs(49)  =   "Named:id=35:Footing"
         _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(51)  =   "Named:id=36:Selected"
         _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(53)  =   "Named:id=37:Caption"
         _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(55)  =   "Named:id=38:HighlightRow"
         _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(57)  =   "Named:id=39:EvenRow"
         _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HE6E6E6&"
         _StyleDefs(59)  =   "Named:id=40:OddRow"
         _StyleDefs(60)  =   ":id=40,.parent=33"
         _StyleDefs(61)  =   "Named:id=41:RecordSelector"
         _StyleDefs(62)  =   ":id=41,.parent=34"
         _StyleDefs(63)  =   "Named:id=42:FilterBar"
         _StyleDefs(64)  =   ":id=42,.parent=33"
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
         Height          =   3660
         Left            =   60
         TabIndex        =   2
         Top             =   2625
         Width           =   11700
         _ExtentX        =   20638
         _ExtentY        =   6456
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
         Columns(1).Caption=   "NAMA"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "SATUAN"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "LAST PRICE"
         Columns(3).DataField=   ""
         Columns(3).NumberFormat=   "###,###,###,###,##0.00"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "CURRENT PRICE"
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "###,###,###,###,##0.00"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "MIN PRICE"
         Columns(5).DataField=   ""
         Columns(5).NumberFormat=   "###,###,###,###,##0.00"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "MAX PRICE"
         Columns(6).DataField=   ""
         Columns(6).NumberFormat=   "###,###,###,###,##0.00"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "TGL LAST PRICE"
         Columns(7).DataField=   ""
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   8
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   873
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   15790320
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2275"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2196"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=512"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=5768"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=5689"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=512"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=1349"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1270"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=512"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=3228"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=3149"
         Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=514"
         Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(25)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=514"
         Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(31)=   "Column(5).Width=2672"
         Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=2593"
         Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=514"
         Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(37)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=514"
         Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(43)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=516"
         Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
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
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=0"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=54,.parent=13,.alignment=0"
         _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=0"
         _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=1"
         _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=58,.parent=13,.alignment=1"
         _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(5).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(53)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.alignment=1"
         _StyleDefs(57)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
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
         _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&HE6E6E6&"
         _StyleDefs(79)  =   "Named:id=40:OddRow"
         _StyleDefs(80)  =   ":id=40,.parent=33"
         _StyleDefs(81)  =   "Named:id=41:RecordSelector"
         _StyleDefs(82)  =   ":id=41,.parent=34"
         _StyleDefs(83)  =   "Named:id=42:FilterBar"
         _StyleDefs(84)  =   ":id=42,.parent=33"
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   345
         Left            =   11145
         TabIndex        =   3
         Top             =   105
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   609
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
         Picture         =   "frmSupplierUtility.frx":0000
      End
      Begin VB.Label Label1 
         Caption         =   "Please type your supplier name here..."
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
         Left            =   3405
         TabIndex        =   5
         Top             =   210
         Width           =   3000
      End
      Begin VB.Label Label2 
         Caption         =   "Daftar Item barang yg pernah diambil sebelumnya"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   75
         TabIndex        =   4
         Top             =   2370
         Width           =   3720
      End
   End
End
Attribute VB_Name = "frmSupplierUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vasupplier As New XArrayDB
Dim vaBarang As New XArrayDB

Private Sub cmdPreview_Click()
Dim n As Integer

  vasupplier.ReDim 0, -1, 0, 3
  Set dbData = objData.Browse(GetDSN, "supplier", "kodesupplier,nama,alamat", "nama", sisContent, cMember.Text, , "nama")
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vasupplier.InsertRows vasupplier.UpperBound(1) + 1
      n = vasupplier.UpperBound(1)
      vasupplier(n, 0) = GetNull(dbData!kodesupplier)
      vasupplier(n, 1) = GetNull(dbData!nama)
      vasupplier(n, 2) = GetNull(dbData!alamat)
      dbData.MoveNext
    Loop
  End If
  Set TDBGrid1.Array = vasupplier
  TDBGrid1.ReBind
  TDBGrid1.Refresh

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then
    Unload Me
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd
  CenterForm Me
  
  TabIndex cMember, n
  TabIndex cmdPreview, n

End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim cSQL As String
Dim n As Integer

  cSQL = ""
  cSQL = cSQL & " select distinct(p.kodestock),min(p.harga) as hargamin,max(p.harga) as hargamax,p.kodesatuan,s.nama,s.hargabeli from totpembelian t"
  cSQL = cSQL & " left join pembelian p on p.nomorpembelian = t.nomorpembelian"
  cSQL = cSQL & " left join stock s on s.kodestock = p.kodestock"
  cSQL = cSQL & " where t.kodesupplier = '" & TDBGrid1.Columns(0) & "'"
  cSQL = cSQL & " group by p.kodestock"
  cSQL = cSQL & " order by s.nama"
  
  vaBarang.ReDim 0, -1, 0, 7
  Set dbData = objData.Sql(GetDSN, cSQL)
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaBarang.InsertRows vaBarang.UpperBound(1) + 1
      n = vaBarang.UpperBound(1)
      vaBarang(n, 0) = GetNull(dbData!KodeStock)
      vaBarang(n, 1) = GetNull(dbData!nama)
      vaBarang(n, 2) = GetNull(dbData!kodesatuan)
      vaBarang(n, 3) = GetLastPrice(TDBGrid1.Columns(0), vaBarang(n, 0), n)
      vaBarang(n, 4) = GetNull(dbData!hargabeli)
      vaBarang(n, 5) = GetNull(dbData!hargamin)
      vaBarang(n, 6) = GetNull(dbData!hargamax)
      dbData.MoveNext
    Loop
  End If
  Set TDBGrid2.Array = vaBarang
  TDBGrid2.ReBind
  TDBGrid2.Refresh
  
End Sub

Private Function GetLastPrice(ByVal Pembeli As String, ByVal Stock As String, ByVal nIndex As Integer) As Double
Dim cSQL As String
Dim db As New ADODB.Recordset
  
  GetLastPrice = 0
  cSQL = ""
  cSQL = cSQL & " select p.nomorpembelian,p.kodestock,s.nama,p.harga,p.tgl from totpembelian t"
  cSQL = cSQL & " left join pembelian p on p.nomorpembelian = t.nomorpembelian"
  cSQL = cSQL & " left join stock s on s.kodestock = p.kodestock"
  cSQL = cSQL & " where p.kodestock = '" & Stock & "' and t.kodesupplier = '" & Pembeli & "'"
  cSQL = cSQL & " order by tgl desc limit 0,1"
  
  Set db = objData.Sql(GetDSN, cSQL)
  If Not db.EOF Then
    vaBarang(nIndex, 7) = Format(GetNull(db!tgl), "dd/MM/yyyy")
    GetLastPrice = GetNull(db!Harga)
  End If
End Function

