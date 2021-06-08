VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptKartuTopUp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kartu Top Up Member"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   7575
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   2205
      Left            =   15
      Top             =   0
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   3889
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
      Begin VB.OptionButton optSaldoAwal 
         Caption         =   "Ya"
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
         Left            =   2490
         TabIndex        =   1
         Top             =   1815
         Width           =   585
      End
      Begin VB.OptionButton optSaldoAwal 
         Caption         =   "Tidak"
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
         Left            =   3105
         TabIndex        =   0
         Top             =   1815
         Width           =   735
      End
      Begin BiSATextBoxProject.BiSABrowse cNama 
         Height          =   330
         Left            =   345
         TabIndex        =   2
         Top             =   210
         Width           =   5520
         _ExtentX        =   9737
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
         Caption         =   "Anggota"
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
      Begin BiSADateProject.BiSADate dTgl 
         Height          =   330
         Index           =   0
         Left            =   330
         TabIndex        =   3
         Top             =   1290
         Width           =   3465
         _ExtentX        =   6112
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
         Caption         =   "Tanggal"
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
      Begin BiSATextBoxProject.BiSABrowse cKode 
         Height          =   330
         Left            =   345
         TabIndex        =   4
         Top             =   570
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   582
         Text            =   "123456"
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
         Caption         =   "Kode Anggota"
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
      Begin BiSADateProject.BiSADate dTgl 
         Height          =   330
         Index           =   1
         Left            =   4185
         TabIndex        =   5
         Top             =   1290
         Width           =   2010
         _ExtentX        =   3545
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
         Caption         =   "s.d"
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
      Begin BiSATextBoxProject.BiSABrowse cAlamat 
         Height          =   330
         Left            =   345
         TabIndex        =   6
         Top             =   915
         Width           =   6660
         _ExtentX        =   11748
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
         Caption         =   "Dept"
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
      Begin VB.Label Label1 
         Caption         =   "Tampilkan Saldo Awal"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   375
         TabIndex        =   7
         Top             =   1800
         Width           =   1950
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   0
      Top             =   2190
      Width           =   7545
      _ExtentX        =   13309
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
         Left            =   6360
         TabIndex        =   8
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
         Picture         =   "rptKartuTopUp.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   5925
         TabIndex        =   9
         Top             =   120
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
         Picture         =   "rptKartuTopUp.frx":00A6
      End
   End
End
Attribute VB_Name = "rptKartuTopUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vaArray As New XArrayDB
Dim objData As New CodeSuiteLibrary.Data
Dim dbData As New ADODB.Recordset

Private Sub GetData()
  Set vaArray = GetKartuTopUp(cKode.Text, dTgl(0).Value, dTgl(1).Value, IIf(optSaldoAwal(0).Value = True, True, False))
  With FrmRPT
    .AddPageHeader aCfg(objData, msNamaPerusahaan), tdbHalignCenter, , , , dbArial, 12, True
    .AddPageHeader "LAPORAN KARTU TOP UP", tdbHalignCenter, , , True
    .AddPageHeader " ", , , , True
    .AddPageHeader " ", , , , True
    
    .AddPageHeader "NAMA MEMBER", , , 15, , , , True, , True, False
    .AddPageHeader ": " & cNama.Text, , , , , , , True
    .AddPageHeader "TANGGAL", , , 15, True, , , True
    .AddPageHeader ": " & Format(dTgl(0).Value, "dd-MM-yyyy") & " s/d " & Format(dTgl(1).Value, "dd-MM-yyyy"), , , , , , , True
    .AddPageHeader " ", , , , True
    
    .AddTableHeader "TANGGAL", , , , 10, , , , , , , , , tdbMergeOnText
    .AddTableHeader "KETERANGAN", , , , , , , , , , , , , tdbMergeOnText
    .AddTableHeader "FAKTUR", , , , 12, , , , , , , , , tdbMergeOnText
    .AddTableHeader "MUTASI", , , , 13, , , , , , , , , , 2
    .AddTableHeader , , , , 13
    .AddTableHeader "SALDO", , , , 16, , , , , , , , , tdbMergeOnText
    
    .AddTableHeader "TANGGAL", , , , , , , , , , True, , , tdbMergeOnText
    .AddTableHeader "KETERANGAN", , , , , , , , , , , , , tdbMergeOnText
    .AddTableHeader "FAKTUR", , , , , , , , , , , , , tdbMergeOnText
    .AddTableHeader "DEBET"
    .AddTableHeader "KREDIT"
    .AddTableHeader "SALDO", , , , , , , , , , , , , tdbMergeOnText
    
     
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    
    
    .AddTableFooter "TOTAL", , tdbHalignRight, , , , , , , , , , , , 3
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter
    

    
    .Preview vaArray, True
  End With
End Sub

Private Sub cKode_ButtonClick()
  Set dbData = objData.PICK(GetDSN, "anggota", "kodeanggota", cKode, "kodeanggota,nama,kodedep")
  If Not dbData.EOF Then
    GetDataanggota
  End If
End Sub

Private Sub cmdPreview_Click()
  GetData
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cNama_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "anggota", "nama,kodedep,kodeanggota", "nama", sisContent, cNama.Text, " or kodeanggota like '%" & cNama.Text & "%'", "nama")
  cNama.Text = cNama.Browse(dbData)
  If Not dbData.EOF Then
    GetDataanggota
  End If
End Sub

Private Sub GetDataanggota()
  cKode.Text = GetNull(dbData!kodeanggota, "")
  cNama.Text = GetNull(dbData!nama, "")
  cAlamat.Text = GetNull(dbData!kodedep, "")
End Sub

Private Sub Form_Load()
Dim n As Single

    SetIcon Me.hWnd, "SIKD"
    CenterForm Me
    initvalue
    TabIndex cKode, n
    TabIndex cNama, n
    TabIndex dTgl(0), n
    TabIndex dTgl(1), n
    TabIndex optSaldoAwal(0), n
    TabIndex optSaldoAwal(1), n
    TabIndex cmdPreview, n
    TabIndex cmdKeluar, n
End Sub

Sub initvalue()
  cKode.Default
  cNama.Default
  dTgl(0).Value = BOM(Date)
  dTgl(1).Value = Date
  optSaldoAwal(0).Value = True
End Sub

Private Sub optSaldoAwal_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeysA vbKeyTab, True
  End If
End Sub

