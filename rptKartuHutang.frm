VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptKartuHutang 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "KARTU HUTANG"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   7575
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1830
      Left            =   15
      Top             =   0
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   3228
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
      Begin BiSATextBoxProject.BiSABrowse cNama 
         Height          =   330
         Left            =   330
         TabIndex        =   0
         Top             =   540
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
         Button          =   -1  'True
         Caption         =   "Nama"
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
         TabIndex        =   1
         Top             =   1290
         Width           =   3465
         _ExtentX        =   6112
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
         Left            =   330
         TabIndex        =   2
         Top             =   165
         Width           =   3420
         _ExtentX        =   6033
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
         MaxLength       =   6
         Button          =   -1  'True
         Caption         =   "Kode Supplier"
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
         TabIndex        =   3
         Top             =   1290
         Width           =   2010
         _ExtentX        =   3545
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
         Left            =   330
         TabIndex        =   4
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
         Caption         =   "Alamat"
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
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   0
      Top             =   1815
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
         TabIndex        =   5
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
         Picture         =   "rptKartuHutang.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   5925
         TabIndex        =   6
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
         Picture         =   "rptKartuHutang.frx":00A6
      End
   End
End
Attribute VB_Name = "rptKartuHutang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vaArray As New XArrayDB
Dim objData As New CodeSuiteLibrary.Data
Dim dbData As New ADODB.Recordset

Private Sub GetData()
  Set vaArray = Getkartuhutang(cKode.Text, dTgl(0).Value, dTgl(1).Value)
  
  With FrmRPT
    .AddPageHeader "KARTU HUTANG", tdbHalignCenter, , , , dbArial, 12, True, True
    .AddPageHeader " ", , , , True
    .AddPageHeader " ", , , , True
    .AddPageHeader " ", , , , True
    
    .AddPageHeader "NAMA supplier", , , 15, , , , True, , True, False
    .AddPageHeader ": " & cNama.Text, , , , , , , True
    .AddPageHeader "TANGGAL", , , 15, True, , , True
    .AddPageHeader ": " & Format(dTgl(0).Value, "dd-MM-yyyy") & " s/d " & Format(dTgl(1).Value, "dd-MM-yyyy"), , , , , , , True
    .AddPageHeader " ", , , , True
    
    .AddTableHeader "TANGGAL", , , , 10, , , , , , , , , tdbMergeOnText
    .AddTableHeader "KETERANGAN", , , , , , , , , , , , , tdbMergeOnText
    .AddTableHeader "FAKTUR", , , , 13, , , , , , , , , tdbMergeOnText
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
    
    .Preview vaArray, True
  End With
End Sub

Private Sub cKode_ButtonClick()
  Set dbData = objData.PICK(GetDSN, "supplier", "kodesupplier", cKode, "kodesupplier,nama,alamat")
  If Not dbData.EOF Then
    GetDatasupplier
  End If
End Sub

Private Sub cmdPreview_Click()
  GetData
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cNama_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "supplier", "nama,alamat,kodesupplier", "nama", sisContent, cNama.Text, , "nama")
  cNama.Text = cNama.Browse(dbData)
  If Not dbData.EOF Then
    GetDatasupplier
  End If
End Sub

Private Sub GetDatasupplier()
  cKode.Text = GetNull(dbData!kodesupplier, "")
  cNama.Text = GetNull(dbData!nama)
  cAlamat.Text = GetNull(dbData!alamat)
End Sub

Private Sub Form_Load()
Dim n As Single

    SetIcon Me.hWnd, "SIKD"
    CenterForm Me
    InitValue
    TabIndex cKode, n
    TabIndex cNama, n
    TabIndex dTgl(0), n
    TabIndex dTgl(1), n
    TabIndex cmdPreview, n
    TabIndex cmdKeluar, n
End Sub

Sub InitValue()
  cKode.Default
  cNama.Default
  dTgl(0).Value = BOM(Date)
  dTgl(1).Value = Date
End Sub


