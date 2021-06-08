VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trSupplierBalance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Supplier Balance"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10395
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   10395
   Begin VB.CheckBox Check1 
      Caption         =   "Clear all balance"
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
      Left            =   735
      TabIndex        =   0
      Top             =   510
      Width           =   2745
   End
   Begin BiSADateProject.BiSADate dTgl 
      Height          =   330
      Left            =   615
      TabIndex        =   1
      Top             =   90
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   582
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
      Caption         =   "Tgl Posting"
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
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   600
      Left            =   120
      Top             =   5730
      Width           =   10230
      _ExtentX        =   18045
      _ExtentY        =   1058
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
         Left            =   6525
         TabIndex        =   2
         Top             =   90
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   767
         Caption         =   "    &Un Posting"
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
         Picture         =   "trSupplierBalance.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   435
         Left            =   5265
         TabIndex        =   3
         Top             =   90
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   767
         Caption         =   "  &Posting"
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
         Picture         =   "trSupplierBalance.frx":028A
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   9060
         TabIndex        =   4
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
         Picture         =   "trSupplierBalance.frx":0435
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   7980
         TabIndex        =   5
         Top             =   90
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
         Picture         =   "trSupplierBalance.frx":04DB
      End
   End
   Begin TrueOleDBGrid70.TDBGrid tdbgrid1 
      Height          =   4665
      Left            =   135
      TabIndex        =   6
      Top             =   1065
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   8229
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
      Columns(1).Caption=   "SUPPLIER"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "ALAMAT"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "BALANCE"
      Columns(3).DataField=   ""
      Columns(3).NumberFormat=   "###,###,###,##0.00"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2408"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2328"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=6112"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=6033"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=512"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=5080"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=5001"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=3069"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2990"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=514"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      Appearance      =   0
      ColumnFooters   =   -1  'True
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
      _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
      _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=1"
      _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(48)  =   "Named:id=33:Normal"
      _StyleDefs(49)  =   ":id=33,.parent=0"
      _StyleDefs(50)  =   "Named:id=34:Heading"
      _StyleDefs(51)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(52)  =   ":id=34,.wraptext=-1"
      _StyleDefs(53)  =   "Named:id=35:Footing"
      _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(55)  =   "Named:id=36:Selected"
      _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(57)  =   "Named:id=37:Caption"
      _StyleDefs(58)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(59)  =   "Named:id=38:HighlightRow"
      _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(61)  =   "Named:id=39:EvenRow"
      _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HE9E9E9&"
      _StyleDefs(63)  =   "Named:id=40:OddRow"
      _StyleDefs(64)  =   ":id=40,.parent=33"
      _StyleDefs(65)  =   "Named:id=41:RecordSelector"
      _StyleDefs(66)  =   ":id=41,.parent=34"
      _StyleDefs(67)  =   "Named:id=42:FilterBar"
      _StyleDefs(68)  =   ":id=42,.parent=33"
   End
   Begin BiSATextBoxProject.BiSABrowse cRekeningLaba 
      Height          =   330
      Left            =   4710
      TabIndex        =   7
      Top             =   105
      Width           =   3540
      _ExtentX        =   6244
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
      Caption         =   "Rekening"
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
   Begin BiSATextBoxProject.BiSATextBox cNamaRekeningLaba 
      Height          =   330
      Left            =   5685
      TabIndex        =   8
      Top             =   420
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
End
Attribute VB_Name = "trSupplierBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB
Dim lSave As Boolean

Private Sub GetLoadRows()
Dim n As Integer

  vaArray.ReDim 0, -1, 0, 6
  Set dbData = objData.Browse(GetDSN, "supplier a", "a.kodesupplier,a.nama,a.alamat,m.balance", , , , , "a.kodesupplier", Array("left join supplierbalance m on m.kodesupplier = a.kodesupplier"))
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!kodesupplier)
      vaArray(n, 1) = GetNull(dbData!nama)
      vaArray(n, 2) = GetNull(dbData!alamat)
      vaArray(n, 3) = GetNull(dbData!Balance)
      vaArray(n, 4) = "SBL-" & vaArray(n, 0)
      dbData.MoveNext
    Loop
  End If
  GetTotal
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
  TDBGrid1.Refresh
End Sub

Private Sub Check1_Click()
Dim n As Integer

  If Check1.Value = 1 Then
    For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
      vaArray(n, 6) = vaArray(n, 3)
      vaArray(n, 3) = 0
    Next n
  Else
    For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
      vaArray(n, 3) = vaArray(n, 6)
      vaArray(n, 6) = 0
    Next n
  End If
  TDBGrid1.ReBind
  TDBGrid1.Refresh
  TDBGrid1.Update
End Sub

Private Sub cmdAdd_Click()
Dim n As Integer
lSave = True

  If Not GetValidDataBrowse(objData, "akun", "kodeakun", cRekeningLaba.Text) Then
    MsgBox "Kode Rekening Belum Disiis Proses Posting Tidak Bisa Dilanjutkan" & vbCrLf & "Data tidak bisa disimpan"
    cRekeningLaba.SetFocus
    Exit Sub
  End If

  objData.Start GetDSN
  FrmPB.InitPB vaArray.UpperBound(1)
  For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
    FrmPB.RunPB
    lSave = IIf(lSave, UpdKodeTr(objData, msSupplierBalance, vaArray(n, 4), Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msHutangDagang), GetCostCenterUser(objData, GetRegistry(reg_username)), "Supplier Balance as " & vaArray(n, 1), 0, vaArray(n, 3), "", SNow), False)
    lSave = IIf(lSave, UpdKodeTr(objData, msSupplierBalance, vaArray(n, 4), Format(dTgl.Value, "yyyy-MM-dd"), cRekeningLaba.Text, GetCostCenterUser(objData, GetRegistry(reg_username)), "Supplier Balance as " & vaArray(n, 1), vaArray(n, 3), 0, "", SNow), False)
    
  Next n
  FrmPB.EndPB
  
  If lSave Then
    objData.Save GetDSN
  Else
    objData.Cancel GetDSN
  End If
  MsgBox "OK, Proses Posting selesai"
End Sub

Private Sub cmdHapus_Click()
Dim n As Integer
lSave = True

  objData.Start GetDSN
  For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
    lSave = IIf(lSave, objData.Delete(GetDSN, "bukubesar", "status", sisAssign, vbTrigger.msSupplierBalance), False)
  Next n
  If lSave Then
    objData.Save GetDSN
  Else
    objData.Cancel GetDSN
  End If
  MsgBox "OK, Proses Un Posting selesai"
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdSimpan_Click()
Dim n As Integer
lSave = True

  objData.Start GetDSN
  FrmPB.InitPB vaArray.UpperBound(1)
  For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
    FrmPB.RunPB
    lSave = IIf(lSave, objData.Delete(GetDSN, "supplierbalance", "supplierbalanceid", sisAssign, vaArray(n, 4)), False)
    lSave = IIf(lSave, objData.Add(GetDSN, "supplierbalance", Array("supplierbalanceid", "kodesupplier", "balance"), Array(vaArray(n, 4), vaArray(n, 0), vaArray(n, 3))), False)
    lSave = IIf(lSave, UpdKartuHutang(objData, SisSupplierBalance, vaArray(n, 4), Format(dTgl.Value, "yyyy-MM-dd"), vaArray(n, 0), "Supplier Balance as " & vaArray(n, 1), vaArray(n, 3), SNow, GetRegistry(reg_username)), False)
    
    lSave = IIf(lSave, objData.Delete(GetDSN, "totpembelian", "nomorpembelian", sisAssign, vaArray(n, 4)), False)
    lSave = IIf(lSave, objData.Add(GetDSN, "totpembelian", _
    Array("nomorpembelian", "kodesupplier", "username", "kodecostcenter", "tgl", "jthtmp", "subtotal", "total", "hutang", "flaglunas"), _
    Array(vaArray(n, 4), vaArray(n, 0), GetRegistry(reg_username), GetCostCenterUser(objData, GetRegistry(reg_username)), Format(dTgl.Value, "yyyy-MM-dd"), Format(dTgl.Value, "yyyy-MM-dd"), vaArray(n, 3), vaArray(n, 3), vaArray(n, 3), 0)), False)
  
  
  Next n
  FrmPB.EndPB
  If lSave Then
    objData.Save GetDSN
  Else
    objData.Cancel GetDSN
  End If
  MsgBox "Data telah disimpan"
  GetLoadRows
End Sub

Private Sub cRekeningLaba_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", , "kodeakun", sisPrefix, "3", " AND jenis = 'D'")
  If Not dbData.EOF Then
    cRekeningLaba.Text = cRekeningLaba.Browse(dbData)
    cNamaRekeningLaba.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  SetIcon Me.hWnd
  
  GetLoadRows
  cRekeningLaba.Default
  cNamaRekeningLaba.Default
  TabIndex dTgl, n
  TabIndex cRekeningLaba, n
  TabIndex cmdAdd, n
  TabIndex cmdHapus, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
End Sub

Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
  TDBGrid1.Update
  GetTotal
End Sub

Private Sub tdbgrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
  If ColIndex <> 3 Then
    Cancel = True
    Exit Sub
  End If
  
  If Not IsNumeric(TDBGrid1.Columns(3).Value) Then
    Cancel = True
    Exit Sub
  End If
End Sub

Private Sub GetTotal()
Dim n As Integer
Dim nTmp As Double

  nTmp = 0
  For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
    nTmp = nTmp + vaArray(n, 3)
  Next n
  TDBGrid1.Columns(3).FooterText = Format(nTmp, "###,###,###,##0.00")
End Sub

