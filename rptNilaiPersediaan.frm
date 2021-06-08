VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptNilaiPersediaan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nilai Persediaan"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   5970
   Begin BiSAButtonProject.BiSAButton BiSAButton1 
      Height          =   435
      Left            =   3165
      TabIndex        =   13
      Top             =   1845
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   767
      Caption         =   "    Update"
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
      Picture         =   "rptNilaiPersediaan.frx":0000
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Pilih Gudang"
      Height          =   195
      Left            =   1050
      TabIndex        =   9
      Top             =   3090
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.OptionButton optKodeStock 
      Caption         =   "Barcode"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   2715
      TabIndex        =   8
      Top             =   1455
      Width           =   1395
   End
   Begin VB.OptionButton optKodeStock 
      Caption         =   "Kode Index"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1290
      TabIndex        =   7
      Top             =   1455
      Width           =   1365
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Tampilkan Seluruh Stock"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   150
      Width           =   2190
   End
   Begin BiSADateProject.BiSADate dTgl 
      Height          =   405
      Left            =   45
      TabIndex        =   4
      Top             =   945
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   714
      Value           =   "04-12-2018"
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
      Caption         =   "sd Tgl"
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
      Index           =   0
      Left            =   45
      TabIndex        =   0
      Top             =   570
      Width           =   3105
      _ExtentX        =   5477
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
   Begin BiSATextBoxProject.BiSABrowse cGolongan 
      Height          =   330
      Index           =   1
      Left            =   3150
      TabIndex        =   1
      Top             =   570
      Width           =   2685
      _ExtentX        =   4736
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
      Button          =   -1  'True
      Caption         =   "s.d"
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
   Begin BiSAButtonProject.BiSAButton cmdKeluar 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   4725
      TabIndex        =   2
      Top             =   1845
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
      Picture         =   "rptNilaiPersediaan.frx":059A
   End
   Begin BiSAButtonProject.BiSAButton cmdPreview 
      Height          =   435
      Left            =   4290
      TabIndex        =   3
      Top             =   1845
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
      Picture         =   "rptNilaiPersediaan.frx":0640
   End
   Begin BiSATextBoxProject.BiSABrowse cGudang 
      Height          =   330
      Left            =   375
      TabIndex        =   10
      Top             =   2910
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      Caption         =   "Gudang"
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
   Begin BiSATextBoxProject.BiSATextBox cNamaGudang 
      Height          =   330
      Left            =   3270
      TabIndex        =   11
      Top             =   2940
      Visible         =   0   'False
      Width           =   2445
      _ExtentX        =   4313
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
   Begin VB.Label Label3 
      Caption         =   "Pilih Gudang"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1605
      TabIndex        =   12
      Top             =   2790
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Tampilkan"
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
      Left            =   75
      TabIndex        =   6
      Top             =   1500
      Width           =   945
   End
End
Attribute VB_Name = "rptNilaiPersediaan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub BiSAButton1_Click()
Dim cSQL As String

  cSQL = "select k.kodestock,s.barcode,s.nama,(sum(k.debet*k.hp)-sum(k.kredit*k.hp))/sum(k.debet-k.kredit) as harga_pokok from kartustock k"
  cSQL = cSQL & " left join stock s on s.kodestock = k.kodestock"
  cSQL = cSQL & " GROUP BY k.kodestock"

  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    dbData.MoveFirst
    Do While Not dbData.EOF
      FrmPB.RunPB
      objData.Edit GetDSN, "stock", "kodestock = '" & GetNull(dbData!KodeStock) & "'", Array("cogs"), Array(GetNull(dbData!harga_pokok))
      dbData.MoveNext
    Loop
    FrmPB.EndPB
    MsgBox "Selesai", vbInformation
  End If
End Sub

Private Sub cGolongan_ButtonClick(Index As Integer)
  Set dbData = objData.Browse(GetDSN, "golongan", "kodegolongan,keterangan", "kodegolongan", sisContent, cGolongan(Index).Text, , "kodegolongan")
  If Not dbData.EOF Then
    cGolongan(Index).Text = cGolongan(Index).Browse(dbData)
  End If
End Sub

Private Sub cGudang_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "gudang", "kodegudang,keterangan")
  If Not dbData.EOF Then
    cGudang.Text = cGudang.Browse(dbData)
    cGudang.Text = GetNull(dbData!Kodegudang)
    cNamaGudang.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    keybd_event VK_TAB, 0, 0, 0
  End If
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
  GetRpt
End Sub

Private Sub GetRpt()
   Set vaArray = GetNilaiPersediaan(cGolongan(0).Text, cGolongan(1).Text, True, dTgl.Value, IIf(Check1.Value = 1, True, False), IIf(optKodeStock(0).Value = True, True, False), IIf(Check2.Value = 1, True, False), cGudang.Text)
   With FrmRPT

    .AddPageHeader "Nilai Persediaan Stock", tdbHalignCenter, , , True, , 12, True, True, , False, tdbPageHeaderSect
    If Check1.Value = 1 Then
      .AddPageHeader "All Category", , , 10, True, , , True
    Else
      .AddPageHeader "Golongan ", , , 10, True, , , True
      .AddPageHeader " : [ " & cGolongan(0).Text & " ] s/d [ " & cGolongan(1).Text & " ]", tdbHalignLeft, , , , , , True, , , False, , , , , , 15
    End If
    
    .AddPageHeader "Sd Tgl " & Format(dTgl.Value, "dd-MM-yyyyy"), , , 20, True, , , True
    
    If Check2.Value = 1 Then
      .AddPageHeader "Gudang " & cGudang.Text, , , 10, True, , , True
    End If
    
  
    .AddTableGroupHeader True, "[]", , , , 10, , , , , , , , , , , , , , , True
    .AddTableGroupHeader
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    
    .AddTableHeader , , , , , , , , , , , , , , , , , , , False
    .AddTableHeader , , , , , , , , , , , , , , , , , , , False
    .AddTableHeader "Kode", , , , 11
    .AddTableHeader "Nama", , , , 0
    .AddTableHeader "Satuan", , , , 6
    .AddTableHeader "Stock", , , , 7
    .AddTableHeader "Hrg Beli", , , , 12
    .AddTableHeader "Hrg Pokok", , , , 14
    .AddTableHeader "Nilai Persediaan", , , , 16
    
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody Sis_Rpt_Number, , , 8
    .AddTableBody Sis_Rpt_Number, , , 8
    .AddTableBody Sis_Rpt_Number, , , 15
    .AddTableBody Sis_Rpt_Number, , , 15
    
    .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupFooter "SubTotal", , tdbHalignRight, , , , , , , , , , , , 6
    .AddTableGroupFooter
    .AddTableGroupFooter
    .AddTableGroupFooter
    .AddTableGroupFooter
    .AddTableGroupFooter
    .AddTableGroupFooter "&Sum", Sis_Rpt_Number
    
    
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableFooter "Grand Total", , tdbHalignRight, , , , , , , , , , , , 6
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter "&Sum", Sis_Rpt_Number
    
    .Refresh
    .Preview vaArray, True
  End With
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  Check1.Value = 1
  dTgl.Value = Date
  optKodeStock(1).Value = True
  cGudang.Default
  cNamaGudang.Default
  
  TabIndex Check1, n
  TabIndex cGolongan(0), n
  TabIndex cGolongan(1), n
  TabIndex dTgl, n
  TabIndex optKodeStock(0), n
  TabIndex optKodeStock(1), n
  TabIndex Check2, n
  TabIndex cGudang, n
  
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub

Private Sub optKodeStock_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    keybd_event VK_TAB, 0, 0, 0
  End If
End Sub
