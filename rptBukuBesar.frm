VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptBukuBesar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laporan Buku Besar..."
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   9855
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1845
      Left            =   0
      Top             =   0
      Width           =   9840
      _ExtentX        =   17357
      _ExtentY        =   3254
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
         Caption         =   "Ya, Tampilkan Buku Besar konsolidasi"
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
         Left            =   2085
         TabIndex        =   7
         Top             =   1065
         Width           =   3645
      End
      Begin BiSATextBoxProject.BiSATextBox cNamaRekening 
         Height          =   330
         Left            =   4725
         TabIndex        =   0
         Top             =   210
         Width           =   5010
         _ExtentX        =   8837
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
      Begin BiSATextBoxProject.BiSABrowse cRekening 
         Height          =   330
         Left            =   285
         TabIndex        =   1
         Top             =   210
         Width           =   4440
         _ExtentX        =   7832
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
         Caption         =   "REKENING"
         CaptionWidth    =   1700
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
      Begin BiSADateProject.BiSADate dDate 
         Height          =   330
         Index           =   0
         Left            =   285
         TabIndex        =   2
         Top             =   675
         Width           =   3180
         _ExtentX        =   5609
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
         Caption         =   "ANTARA TANGGAL"
         CaptionWidth    =   1700
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
      Begin BiSADateProject.BiSADate dDate 
         Height          =   330
         Index           =   1
         Left            =   3870
         TabIndex        =   3
         Top             =   675
         Width           =   1995
         _ExtentX        =   3519
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
         Caption         =   "S.D"
         CaptionWidth    =   500
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
      Begin BiSATextBoxProject.BiSABrowse cCostCenter 
         Height          =   330
         Left            =   270
         TabIndex        =   6
         Top             =   1350
         Width           =   3255
         _ExtentX        =   5741
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
         Caption         =   "Cost Centre"
         CaptionWidth    =   1700
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
         Caption         =   "Konsolidasi"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   300
         TabIndex        =   8
         Top             =   1050
         Width           =   1350
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   630
      Left            =   0
      Top             =   1830
      Width           =   9840
      _ExtentX        =   17357
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
         Left            =   8580
         TabIndex        =   4
         Top             =   90
         Width           =   1140
         _ExtentX        =   2011
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
         Picture         =   "rptBukuBesar.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   7410
         TabIndex        =   5
         Top             =   90
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   767
         Caption         =   "     &Preview"
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
         Picture         =   "rptBukuBesar.frx":00A6
      End
   End
End
Attribute VB_Name = "rptBukuBesar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objData As New CodeSuiteLibrary.Data
Dim dbData As New ADODB.Recordset
Dim db As New ADODB.Recordset
Dim vaArray As New XArrayDB

Private Sub cCostCenter_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "costcenter", "kodecostcenter,keterangan")
  If Not dbData.EOF Then
    cCostCenter.Text = cCostCenter.Browse(dbData)
  End If
End Sub

Private Sub Check1_Click()
  If Check1.Value = 1 Then
    cCostCenter.Enabled = False
  Else
    cCostCenter.Enabled = True
  End If
End Sub

Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    keybd_event VK_TAB, 0, 0, 0
  End If
End Sub

Private Sub cmdPreview_Click()
  getSQL
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  SetIcon Me.hWnd
  dDate(0).Value = Date
  dDate(1).Value = Date
  Check1.Value = 1
  cCostCenter.Default
  If Check1.Value = 1 Then
    cCostCenter.Enabled = False
  End If
  TabIndex cRekening, n
  TabIndex dDate(0), n
  TabIndex dDate(1), n
  TabIndex Check1, n
  TabIndex cCostCenter, n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
  cRekening.BackColor = vbWhite
  cRekening.Enabled = True
  If GetRegistry(reg_UserLevel) <> 0 Then
    cRekening.Text = GetAkunKas(objData, GetRegistry(reg_Username))
    cRekening.Enabled = False
    cRekening.BackColor = vbButtonFace
  End If
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cRekening_ButtonClick()
  'Set db = objData.PICK(GetDSN, "akun", "kodeakun", cRekening, "Kodeakun,Keterangan", " AND jenis = 'D'")
  Set db = objData.Browse(GetDSN, "akun", "kodeakun,keterangan", "(kodeakun", sisContent, cRekening.Text, " or keterangan like '%" & cRekening.Text & "%') and jenis = 'D'")
  If Not db.EOF Then
    cRekening.Text = cRekening.Browse(db)
    cRekening.Text = GetNull(db!kodeakun)
    cNamaRekening.Text = GetNull(db!keterangan, "")
  End If
End Sub

'Private Sub cRekening_Validate(Cancel As Boolean)
'  If cRekening.LastKey = 13 Then
'    cRekening_ButtonClick
'  End If
'End Sub

Private Sub getSQL()
Dim cSQL As String
Dim n As Double
Dim nDebet As Double
Dim nKredit As Double
Dim cSQLCostCenter As String

  cSQLCostCenter = ""
  If Check1.Value <> 1 Then
    cSQLCostCenter = " AND kodecostcenter = '" & cCostCenter.Text & "' "
  End If
  vaArray.ReDim 0, 0, 0, 5
  nDebet = 0
  nKredit = 0
  vaArray(0, 2) = "SALDO AWAL"
'  cSQL = "Select Sum(Awal) as Awal From SaldoRekening where kodeakun = '" & cRekening.Text & "'"
'  cSQL = cSQL & " union "
  cSQL = cSQL & "Select Sum(debet-kredit) as awal From bukubesar Where tgl < '" & Format(dDate(0).Value, "yyyy-mm-dd") & "' and kodeakun = '" & cRekening.Text & "' " & cSQLCostCenter
  Set dbData = objData.SQL(GetDSN, cSQL)
  vaArray(0, 5) = 0
  If Not dbData.EOF Then
    dbData.MoveFirst
    Do While Not dbData.EOF
      vaArray(0, 5) = GetNull(vaArray(0, 5)) + GetNull((dbData!AWAL))
      dbData.MoveNext
    Loop
  End If
  
  'mengambil data mutasi
  cSQL = ""
  cSQL = "Select faktur,tgl,keterangan,debet,kredit "
  cSQL = cSQL & "From bukubesar Where tgl >= '" & Format(dDate(0).Value, "yyyy-mm-dd") & "' and tgl <= '" & Format(dDate(1).Value, "yyyy-mm-dd") & "' and kodeakun = '" & cRekening.Text & "'" & cSQLCostCenter & " order by tgl,idbukubesar,faktur"
  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    dbData.MoveFirst
    Do While Not dbData.EOF
      nDebet = 0
      nKredit = 0
      FrmPB.RunPB
      n = n + 1
      vaArray.InsertRows n
      vaArray(n, 0) = GetNull((dbData!Faktur), "")
      vaArray(n, 1) = GetNull((dbData!tgl), "")
      vaArray(n, 2) = GetNull((dbData!keterangan), "")
      vaArray(n, 3) = GetNull(dbData!debet)
      vaArray(n, 4) = GetNull(dbData!kredit)
      vaArray(n, 5) = GetNull(vaArray(n - 1, 5)) + GetNull(vaArray(n, 3)) - GetNull(vaArray(n, 4))
      nDebet = nDebet + vaArray(n, 3)
      nKredit = nKredit + vaArray(n, 4)
      dbData.MoveNext
    Loop
    FrmPB.EndPB
  End If
  Rpt
End Sub

Private Sub Rpt()
Dim cHeader As String
  If Check1.Value <> 1 Then
    Set dbData = objData.Browse(GetDSN, "costcenter", , "kodecostcenter", sisAssign, cCostCenter.Text)
    If Not dbData.EOF Then
      cHeader = GetNull(dbData!keterangan, "")
    End If
  Else
    cHeader = "KONSOLIDASI"
  End If
  
  With FrmRPT
    .AddPageHeader "BUKU BESAR " & cHeader, tdbHalignCenter, , , , , 10, True
    .AddPageHeader aCfg(objData, msNamaPerusahaan), tdbHalignCenter, , , True, , 12, True
    .AddPageHeader " ", , , , True
    .AddPageHeader " ", , , , True
    
    .AddPageHeader "REKENING", , , 15, True, , , , , True, , tdbPageHeaderSect
    .AddPageHeader " : " & " [ " & cRekening.Text & " ] " & cNamaRekening.Text
    .AddPageHeader "ANTARA TANGGAL", , , 15, True
    .AddPageHeader " : " & Format(dDate(0).Value, "dd-MM-yyyy") & " S.D " & Format(dDate(1).Value, "dd-MM-yyyy")
    
    .AddTableHeader "FAKTUR", , , , 18
    .AddTableHeader "TANGGAL", , , , 9
    .AddTableHeader "KETERANGAN"
    .AddTableHeader "DEBET", , , , 11
    .AddTableHeader "KREDIT", , , , 11
    .AddTableHeader "SALDO", , , , 13
    
'    .AddTableBody
    .AddTableBody
    .AddTableBody Sis_Rpt_dd_MM_yyyy
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
        
    .AddTableFooter "Total", , tdbHalignCenter, , , , , , , , , , , , 3
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter
        
    .Preview vaArray, True
  End With
End Sub
