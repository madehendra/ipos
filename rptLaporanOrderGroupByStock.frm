VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form rptLaporanOrderGroupByStock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laporan Order Group By Stock"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11325
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   11325
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   5790
      Left            =   0
      Top             =   0
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   10213
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
      Begin BiSATextBoxProject.BiSABrowse cKodeBarang 
         Height          =   330
         Left            =   150
         TabIndex        =   3
         Top             =   210
         Width           =   3915
         _ExtentX        =   6906
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
         Caption         =   "Kode Barang"
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
         Height          =   4635
         Left            =   60
         TabIndex        =   2
         Top             =   1095
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   8176
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "QTY"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "KODE MEMBER"
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
         Columns(4).Caption=   "TELP"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   5
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=5"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2011"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1931"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=514"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=3704"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3625"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=512"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=6085"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=6006"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=512"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=3572"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=3493"
         Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=197120"
         Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(25)=   "Column(4).Width=1482"
         Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1402"
         Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=512"
         Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
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
         _StyleDefs(28)  =   "Splits(0).FooterStyle:id=15,.parent=3,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(29)  =   ":id=15,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(30)  =   ":id=15,.fontname=Tahoma"
         _StyleDefs(31)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(32)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(33)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=1"
         _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
         _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=0"
         _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=66,.parent=13,.alignment=0"
         _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=14"
         _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=15,.alignment=1"
         _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=17"
         _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=0"
         _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(59)  =   "Named:id=33:Normal"
         _StyleDefs(60)  =   ":id=33,.parent=0"
         _StyleDefs(61)  =   "Named:id=34:Heading"
         _StyleDefs(62)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(63)  =   ":id=34,.wraptext=-1"
         _StyleDefs(64)  =   "Named:id=35:Footing"
         _StyleDefs(65)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(66)  =   "Named:id=36:Selected"
         _StyleDefs(67)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(68)  =   "Named:id=37:Caption"
         _StyleDefs(69)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(70)  =   "Named:id=38:HighlightRow"
         _StyleDefs(71)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(72)  =   "Named:id=39:EvenRow"
         _StyleDefs(73)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(74)  =   "Named:id=40:OddRow"
         _StyleDefs(75)  =   ":id=40,.parent=33"
         _StyleDefs(76)  =   "Named:id=41:RecordSelector"
         _StyleDefs(77)  =   ":id=41,.parent=34"
         _StyleDefs(78)  =   "Named:id=42:FilterBar"
         _StyleDefs(79)  =   ":id=42,.parent=33"
      End
      Begin BiSATextBoxProject.BiSABrowse cNamaBarang 
         Height          =   330
         Left            =   150
         TabIndex        =   4
         Top             =   570
         Width           =   5850
         _ExtentX        =   10319
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
         Caption         =   "Nama Barang"
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
      Begin BiSAButtonProject.BiSAButton cmdCari 
         Height          =   435
         Left            =   6045
         TabIndex        =   5
         Top             =   465
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   767
         Caption         =   "Cari"
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
         Picture         =   "rptLaporanOrderGroupByStock.frx":0000
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   0
      Top             =   5775
      Width           =   11280
      _ExtentX        =   19897
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
         Left            =   10125
         TabIndex        =   0
         Top             =   90
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
         Picture         =   "rptLaporanOrderGroupByStock.frx":0286
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   9690
         TabIndex        =   1
         Top             =   90
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
         Picture         =   "rptLaporanOrderGroupByStock.frx":032C
      End
   End
End
Attribute VB_Name = "rptLaporanOrderGroupByStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB


Private Sub cKodeBarang_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "memberorder m", "distinct(m.kodestock),s.barcode,s.nama", "s.barcode", sisContent, cKodeBarang.Text, " and t.status = 0", , Array("LEFT JOIN stock s on s.kodestock = m.kodestock", "left join totmemberorder t on t.nomormemberorder = m.nomormemberorder"))
  If Not dbData.EOF Then
    cKodeBarang.Text = cKodeBarang.Browse(dbData)
    cNamaBarang.Text = GetNull(dbData!nama)
  End If
End Sub

Private Sub cmdCari_Click()
  GetSQL2
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
  GetSQL
End Sub

Private Sub cNamaBarang_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "memberorder m", "distinct(m.kodestock),s.barcode,s.nama", "s.nama", sisContent, cNamaBarang.Text, " and t.status = 0", , Array("LEFT JOIN stock s on s.kodestock = m.kodestock", "left join totmemberorder t on t.nomormemberorder = m.nomormemberorder"))
  If Not dbData.EOF Then
    cNamaBarang.Text = cNamaBarang.Browse(dbData)
    cNamaBarang.Text = GetNull(dbData!nama, "")
    cKodeBarang.Text = GetNull(dbData!KodeStock, "")
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  SetIcon Me.hWnd
  TabIndex cKodeBarang, n
  TabIndex cNamaBarang, n
  TabIndex cmdCari, n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub

Private Sub GetSQL()
Dim cSQL As String
Dim n As Single
  
  vaArray.ReDim 0, -1, 0, 6
  
  cSQL = "select s.barcode,s.nama,t.kodeanggota,a.nama as namaanggota,m.tgl,m.kodestock,m.qty,m.harga from totmemberorder t"
  cSQL = cSQL & " left join memberorder m on m.nomormemberorder = t.nomormemberorder"
  cSQL = cSQL & " left join stock s on s.kodestock = m.kodestock"
  cSQL = cSQL & " left join anggota a on a.kodeanggota = t.kodeanggota"
  cSQL = cSQL & " Where t.status = 0 Or t.status Is Null"
  cSQL = cSQL & " order by s.barcode,t.kodeanggota,m.tgl"
  Set dbData = objData.Sql(GetDSN, cSQL)
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!Barcode)
      vaArray(n, 1) = GetNull(dbData!nama)
      vaArray(n, 2) = GetNull(dbData!tgl)
      vaArray(n, 3) = GetNull(dbData!kodeanggota)
      vaArray(n, 4) = GetNull(dbData!namaanggota)
      vaArray(n, 5) = GetNull(dbData!qty)
      vaArray(n, 6) = GetNull(dbData!Harga)
      dbData.MoveNext
    Loop
    GetRpt
  Else
    MsgBox "Maaf, tidak ada data untuk ditampilkan"
  End If

End Sub

Private Sub GetSQL2()
Dim cSQL As String
Dim n As Single
Dim db As New ADODB.Recordset
Dim vaArr As New XArrayDB
Dim nTotal As Double

  vaArr.ReDim 0, -1, 0, 6
  nTotal = 0
  
  
  Set db = objData.Browse(GetDSN, "memberorder m", "sum(m.qty) as qty,t.kodeanggota,a.nama as namaanggota,a.alamat,a.telp", "s.kodestock", sisAssign, cKodeBarang.Text, " and t.status = 0 GROUP BY t.kodeanggota", , Array("left join stock s on s.kodestock = m.kodestock", "left join totmemberorder t on t.nomormemberorder = m.nomormemberorder", "left join anggota a on a.kodeanggota = t.kodeanggota"))
  
  If Not db.EOF Then
    Do While Not db.EOF
      vaArr.InsertRows vaArr.UpperBound(1) + 1
      n = vaArr.UpperBound(1)
      vaArr(n, 0) = GetNull(db!qty)
      vaArr(n, 1) = GetNull(db!kodeanggota)
      vaArr(n, 2) = GetNull(db!namaanggota)
      vaArr(n, 3) = GetNull(db!alamat)
      vaArr(n, 4) = GetNull(db!telp)
      nTotal = nTotal + vaArr(n, 0)
      db.MoveNext
    Loop
    TDBGrid1.Columns(0).FooterText = nTotal
    Set TDBGrid1.Array = vaArr
    TDBGrid1.ReBind
    TDBGrid1.Refresh
  Else
    MsgBox "Maaf, tidak ada data untuk ditampilkan"
  End If

End Sub

Private Sub GetRpt()
  With FrmRPT
    
    .AddPageHeader "Laporan Order Yg di Group By Item Stock", tdbHalignCenter, , , True, , 10, True, False, True, False, tdbPageHeaderSect
    .AddPageHeader aCfg(objData, msNamaPerusahaan), tdbHalignCenter, , , True, , 14, True, False, True, False, tdbPageHeaderSect
        
    .AddTableGroupHeader True, "[]", , , , 10
    .AddTableGroupHeader
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    
    
    .AddTableHeader "", , , , , , , , , , , , , , , , , , , False
    .AddTableHeader "", , , , , , , , , , , , , , , , , , , False
    .AddTableHeader "Tgl", , , , 11
    .AddTableHeader "Kode Member", , , , 15
    .AddTableHeader "Nama", , , , 25
    .AddTableHeader "Qty", , , , 10
    .AddTableHeader "Harga", , , , 15
    
     
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    
    .Refresh
    .Preview vaArray, True
  End With
End Sub

