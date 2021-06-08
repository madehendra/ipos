VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptMutasiKodeTransaksi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mutasi Kode Transaksi..."
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   7725
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1425
      Left            =   0
      Top             =   0
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   2514
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
      Begin VB.OptionButton Option1 
         Caption         =   "&Ya"
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
         Index           =   0
         Left            =   2115
         TabIndex        =   1
         Top             =   555
         Value           =   -1  'True
         Width           =   525
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&Tidak"
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
         Index           =   1
         Left            =   2775
         TabIndex        =   0
         Top             =   555
         Width           =   840
      End
      Begin BiSADateProject.BiSADate dDate 
         Height          =   330
         Left            =   315
         TabIndex        =   2
         Top             =   150
         Width           =   3270
         _ExtentX        =   5768
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
         Caption         =   "TANGGAL"
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
      Begin BiSATextBoxProject.BiSABrowse cJenisTransaksi 
         Height          =   330
         Left            =   315
         TabIndex        =   3
         Top             =   810
         Width           =   2730
         _ExtentX        =   4815
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
         Caption         =   "JENIS TRANSAKSI"
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
      Begin BiSATextBoxProject.BiSATextBox cNamaJenisTransaksi 
         Height          =   330
         Left            =   3045
         TabIndex        =   4
         Top             =   810
         Width           =   4095
         _ExtentX        =   7223
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
      Begin VB.Label Label1 
         Caption         =   "FILTER"
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
         Left            =   375
         TabIndex        =   5
         Top             =   540
         Width           =   1095
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   630
      Left            =   0
      Top             =   1410
      Width           =   7710
      _ExtentX        =   13600
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
         Left            =   6525
         TabIndex        =   6
         Top             =   105
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
         Picture         =   "rptMutasiKodeTransaksi.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   5355
         TabIndex        =   7
         Top             =   105
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
         Picture         =   "rptMutasiKodeTransaksi.frx":00A6
      End
   End
End
Attribute VB_Name = "rptMutasiKodeTransaksi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub cJenisTransaksi_ButtonClick()
  Set dbData = objData.PICK(GetDSN, "KodeTransaksi", "kodetransaksi", cJenisTransaksi, "Kodetransaksi,Keterangan")
  If Not dbData.EOF Then
    cNamaJenisTransaksi.Text = GetNull(dbData!Keterangan, "")
  End If
End Sub

Private Sub cJenisTransaksi_Validate(Cancel As Boolean)
  cJenisTransaksi_ButtonClick
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
  GetSQL
End Sub

Private Sub GetRpt()
    With FrmRPT
    .AddPageHeader UCase("Daftar Mutasi Harian Simpanan Sukarela"), tdbHalignCenter, , , , , 12, True, True
    .AddPageHeader UCase(aCfg(objData, msNamaPerusahaan)), tdbHalignCenter, , , True, dbArial, 10, True
    .AddPageHeader "TANGGAL  : " & Format(dDate.Value, "dd MMMM yyyy"), tdbHalignCenter, , , True, , 8, True
    .AddPageHeader " ", , , , True
    .AddPageHeader " ", , , , True
    
    'Faktur
    'Rekening
    'Nama
    'Kode
    'Keterangan
    'Debet
    'Kredit
        
    .AddTableHeader "NO", , , , 10
    .AddTableHeader "KODE", , , , 6
    .AddTableHeader "NAMA", , , , 20
    .AddTableHeader "KODE", , , , 5
    .AddTableHeader "KETERANGAN"
    .AddTableHeader "DEBET", , , , 12
    .AddTableHeader "KREDIT", , , , 12
    
    .AddTableBody
    .AddTableBody , tdbHalignCenter
    .AddTableBody
    .AddTableBody , tdbHalignCenter
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
        
    .AddTableFooter "SUB TOTAL", , tdbHalignRight, , , , , , , , , , , , 5
    .AddTableFooter ""
    .AddTableFooter ""
    .AddTableFooter ""
    .AddTableFooter ""
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    
    .Preview vaArray, True
  End With
End Sub

Private Sub GetSQL()
Dim vaJoin
Dim cWhere As String
Dim cField As String

  cField = "m.nomormutasitabungan,m.kodeanggota,r.Nama,"
  cField = cField & "m.KodeTransaksi,m.Keterangan,if(m.DK='D',m.Jumlah,0),if(m.DK='K',m.Jumlah,0)"
  cWhere = " and m.Tgl = '" & Format(dDate.Value, "yyyy-mm-dd") & "'"
  If Trim(cJenisTransaksi.Text) <> "" Then
    cWhere = cWhere & " and m.KodeTransaksi = '" & cJenisTransaksi.Text & "'"
  End If
  vaJoin = Array("Left Join anggota r On r.kodeanggota = m.kodeanggota")
                 
  Set dbData = objData.Browse(GetDSN, "MutasiTabungan m", cField, _
                              , , , " 1=1 " & cWhere, "m.kodeanggota,m.Tgl", vaJoin)
  If Not dbData.EOF Then
    vaArray.LoadRows dbData.GetRows(dbData.RecordCount)
    vaArray.QuickSort vaArray.LowerBound(1), vaArray.UpperBound(1), 0, XORDER_ASCEND, XTYPE_DEFAULT, 2, XORDER_ASCEND, XTYPE_DEFAULT
    GetRpt
  Else
    MsgBox "Data Tidak Ada,..", vbInformation, Me.Caption
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd
  CenterForm Me
  dDate.Value = Date
      
  TabIndex dDate, n
  TabIndex Option1(0), n
  TabIndex Option1(1), n
  TabIndex cJenisTransaksi, n
  TabIndex cNamaJenisTransaksi, n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
  cJenisTransaksi.Enabled = True
End Sub

Private Sub Option1_Click(Index As Integer)
  If Option1(0).Value = True Then
      cJenisTransaksi.Enabled = True
    Else
      cJenisTransaksi.Enabled = False
      cJenisTransaksi.Default
      cNamaJenisTransaksi.Default
  End If
End Sub

Private Sub Option1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    keybd_event VK_TAB, 0, 0, 0
  End If
End Sub

Private Sub Option1_LostFocus(Index As Integer)
  If Option1(0).Value = True Then
      cJenisTransaksi.Enabled = True
    Else
      cJenisTransaksi.Enabled = False
      cJenisTransaksi.Default
      cNamaJenisTransaksi.Default
  End If
End Sub
