VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Begin VB.Form cfgRekeningNilaiPersediaan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rekening Nilai Persediaan"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   8895
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   2655
      Left            =   0
      Top             =   0
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   4683
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
      Begin BiSATextBoxProject.BiSABrowse cRekeningPersediaan 
         Height          =   330
         Left            =   135
         TabIndex        =   0
         Top             =   345
         Width           =   4500
         _ExtentX        =   7938
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
         Caption         =   "Rekening Nilai Persediaan"
         CaptionWidth    =   2500
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekeningPersediaan 
         Height          =   330
         Left            =   4650
         TabIndex        =   1
         Top             =   345
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
      Begin BiSATextBoxProject.BiSABrowse cRekeningPenyesuaian 
         Height          =   330
         Left            =   135
         TabIndex        =   4
         Top             =   690
         Width           =   4500
         _ExtentX        =   7938
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
         Caption         =   "Rekening Penyesuaian (+)"
         CaptionWidth    =   2500
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekeningPenyesuaian 
         Height          =   330
         Left            =   4650
         TabIndex        =   5
         Top             =   690
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
      Begin BiSATextBoxProject.BiSABrowse cRekeningPersediaanKurang 
         Height          =   330
         Left            =   135
         TabIndex        =   6
         Top             =   1035
         Width           =   4500
         _ExtentX        =   7938
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
         Caption         =   "Rekening Penyesuaian (-)"
         CaptionWidth    =   2500
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekeningPersediaanKurang 
         Height          =   330
         Left            =   4650
         TabIndex        =   7
         Top             =   1035
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
      Begin BiSATextBoxProject.BiSABrowse cRekeningBiaya 
         Height          =   330
         Left            =   135
         TabIndex        =   8
         Top             =   1590
         Width           =   4500
         _ExtentX        =   7938
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
         Caption         =   "Rekening Biaya Barang"
         CaptionWidth    =   2500
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekeningBiaya 
         Height          =   330
         Left            =   4650
         TabIndex        =   9
         Top             =   1590
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
      Begin BiSATextBoxProject.BiSABrowse cRekeningHutangBiaya 
         Height          =   330
         Left            =   135
         TabIndex        =   10
         Top             =   1950
         Width           =   4500
         _ExtentX        =   7938
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
         Caption         =   "Rekening Hutang Biaya"
         CaptionWidth    =   2500
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
      Begin BiSATextBoxProject.BiSATextBox cNamaRekeningHutangBiaya 
         Height          =   330
         Left            =   4650
         TabIndex        =   11
         Top             =   1950
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
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   645
      Left            =   -15
      Top             =   2655
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   1138
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
         Left            =   7680
         TabIndex        =   2
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
         Picture         =   "cfgRekeningNilaiPersediaan.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   6600
         TabIndex        =   3
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
         Picture         =   "cfgRekeningNilaiPersediaan.frx":00A6
      End
   End
End
Attribute VB_Name = "cfgRekeningNilaiPersediaan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data


Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdSimpan_Click()
  objData.Start GetDSN
  UpdCfg msRekeningPersediaan, cRekeningPersediaan.Text, objData, cRekeningPersediaan.Caption, Me.Caption
  UpdCfg msRekeningPenyesuian, cRekeningPenyesuaian.Text, objData, cRekeningPenyesuaian.Caption, Me.Caption
  UpdCfg msRekeningPenyesuaianKurang, cRekeningPersediaanKurang.Text, objData, cRekeningPersediaanKurang.Caption, Me.Caption
  UpdCfg msRekeningBiayaBarang, cRekeningBiaya.Text, objData, cRekeningBiaya.Caption, Me.Caption
  UpdCfg msRekeningHutangBiaya, cRekeningHutangBiaya.Text, objData, cRekeningHutangBiaya.Caption, Me.Caption
  
  objData.Save GetDSN
  MsgBox "Data telah tersimpan", vbInformation
End Sub


Private Sub cRekeningBiaya_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", , "jenis", sisAssign, "D", " AND left(kodeakun,1) = '5'")
  If Not dbData.EOF Then
    cRekeningBiaya.Text = cRekeningBiaya.Browse(dbData)
    cNamaRekeningBiaya.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub cRekeningHutangBiaya_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", , "jenis", sisAssign, "D", " AND left(kodeakun,1) = '2'")
  If Not dbData.EOF Then
    cRekeningHutangBiaya.Text = cRekeningBiaya.Browse(dbData)
    cNamaRekeningHutangBiaya.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub cRekeningPenyesuaian_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", , "jenis", sisAssign, "D", " AND (left(kodeakun,1) = '4' or left(kodeakun,1)='3')")
  If Not dbData.EOF Then
    cRekeningPenyesuaian.Text = cRekeningPersediaan.Browse(dbData)
    cNamaRekeningPenyesuaian.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub cRekeningPersediaan_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", , "kodeakun", sisPrefix, "1", " AND jenis = 'D'")
  If Not dbData.EOF Then
    cRekeningPersediaan.Text = cRekeningPersediaan.Browse(dbData)
    cNamaRekeningPersediaan.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub cRekeningPersediaanKurang_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", , "jenis", sisAssign, "D", " AND left(kodeakun,1) = '5'")
  If Not dbData.EOF Then
    cRekeningPersediaanKurang.Text = cRekeningPersediaanKurang.Browse(dbData)
    cNamaRekeningPersediaanKurang.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  SetIcon Me.hwnd
  TabIndex cRekeningPersediaan, n
  TabIndex cRekeningPenyesuaian, n
  TabIndex cRekeningPersediaanKurang, n
  TabIndex cRekeningBiaya, n
  TabIndex cRekeningHutangBiaya, n
  
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  
  cRekeningPersediaan.Text = aCfg(objData, msRekeningPersediaan)
  cNamaRekeningPersediaan.Text = GetNamaRekening(cRekeningPersediaan.Text)
  cRekeningPenyesuaian.Text = aCfg(objData, msRekeningPenyesuian)
  cNamaRekeningPenyesuaian.Text = GetNamaRekening(cRekeningPenyesuaian.Text)
  cRekeningPersediaanKurang.Text = aCfg(objData, msRekeningPenyesuaianKurang)
  cNamaRekeningPersediaanKurang.Text = GetNamaRekening(cRekeningPersediaanKurang.Text)
  cRekeningBiaya.Text = aCfg(objData, msRekeningBiayaBarang)
  cNamaRekeningBiaya.Text = GetNamaRekening(cRekeningBiaya.Text)
  cRekeningHutangBiaya.Text = aCfg(objData, msRekeningHutangBiaya)
  cNamaRekeningHutangBiaya.Text = GetNamaRekening(cRekeningHutangBiaya.Text)
End Sub

Private Function GetNamaRekening(cAkun As String) As String
  GetNamaRekening = ""
  Set dbData = objData.Browse(GetDSN, "Akun", "Keterangan", "KodeAkun", sisAssign, cAkun)
  If Not dbData.EOF Then
    GetNamaRekening = GetNull(dbData!keterangan, "")
  End If
End Function



