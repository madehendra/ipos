VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form trPurchaseOrder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Order"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   11700
   Begin BiSAButtonProject.BiSAButton cmdGet 
      Height          =   555
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   979
      Caption         =   "GET FAKTUR PURCHASE ORDER"
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
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   6045
      Left            =   30
      TabIndex        =   0
      Top             =   750
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   10663
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
      Columns(5).NumberFormat=   "###,###,###,###,##0.00"
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
      Columns(7).NumberFormat=   "###,###,###,###,##0.00"
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
      ColumnFooters   =   -1  'True
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1.5
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
   Begin BiSATextBoxProject.BiSABrowse cAkunKas 
      Height          =   540
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   953
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   21.75
      Enabled         =   0   'False
      Appearance      =   0
      CaptionWidth    =   2000
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
   Begin BiSAButtonProject.BiSAButton cmdKeluar 
      Cancel          =   -1  'True
      Height          =   510
      Left            =   10305
      TabIndex        =   3
      Top             =   6840
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   900
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
      Picture         =   "trPurchaseOrder.frx":0000
   End
   Begin BiSAButtonProject.BiSAButton cmdSimpan 
      Height          =   510
      Left            =   8835
      TabIndex        =   4
      Top             =   6840
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   900
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
      Picture         =   "trPurchaseOrder.frx":00A6
   End
End
Attribute VB_Name = "trPurchaseOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub cmdGet_Click()
Dim n As Integer
Dim nTotal As Double
Dim nTotal2 As Double
Dim nQtyTmp As Double

  nTotal = 0
  nTotal2 = 0
  nQtyTmp = 0
  
  Set dbData = objData.Browse(GetDSN, "po p", "p.kodestock,sum(p.qty) AS qty,s.nama,s.kodesatuan,s.hargabeli,s.diskonpenjualan", "p.statusorder", sisAssign, 0, " GROUP BY p.kodestock", , Array("LEFT JOIN stock s ON s.kodestock = p.kodestock"))
  vaArray.ReDim 0, -1, 0, 7
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.EOF
      FrmPB.RunPB
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = n + 1
      vaArray(n, 1) = GetNull(dbData!KodeStock)
      vaArray(n, 2) = GetNull(dbData!nama)
      vaArray(n, 3) = GetNull(dbData!qty)
      vaArray(n, 4) = GetNull(dbData!kodesatuan)
      vaArray(n, 5) = GetNull(dbData!hargabeli)
      vaArray(n, 6) = GetNull(dbData!diskonpenjualan)
      vaArray(n, 7) = vaArray(n, 3) * (vaArray(n, 5) - (vaArray(n, 5) * vaArray(n, 6) / 100))
      nTotal = nTotal + (vaArray(n, 5) * vaArray(n, 3))
      nTotal2 = nTotal2 + vaArray(n, 7)
      nQtyTmp = nQtyTmp + vaArray(n, 3)
      dbData.MoveNext
    Loop
    FrmPB.EndPB
  End If
  TDBGrid1.Columns(3).FooterText = nQtyTmp
  TDBGrid1.Columns(5).FooterText = Format(nTotal, "###,###,###,##0.00")
  TDBGrid1.Columns(7).FooterText = Format(nTotal2, "###,###,###,##0.00")
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdSimpan_Click()
  Set dbData = objData.Browse(GetDSN, "po p", "p.id,p.kodestock,p.qty,s.nama,s.kodesatuan,s.hargabeli,s.diskonpenjualan", "p.statusorder", sisAssign, 0, , , Array("LEFT JOIN stock s ON s.kodestock = p.kodestock"))
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.EOF
      FrmPB.RunPB
      objData.Update GetDSN, "po", "id = " & GetNull(dbData!ID), Array("statusorder", "tgl", "datetime", "username"), Array(1, Date, SNow, GetRegistry(reg_UserName))
      dbData.MoveNext
    Loop
    FrmPB.EndPB
    cmdGet_Click
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  cAkunKas.Text = Year(Now) & Month(Now) & Day(Now)
  TabIndex cmdGet, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  
End Sub
