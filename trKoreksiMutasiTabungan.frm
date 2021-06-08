VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trKoreksiMutasiTabungan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Koreksi Mutasi Simpanan Harian"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   7020
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   3600
      Left            =   0
      Top             =   1590
      Width           =   7020
      _ExtentX        =   12383
      _ExtentY        =   6350
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
      Begin BiSATextBoxProject.BiSATextBox cKeterangan 
         Height          =   330
         Left            =   450
         TabIndex        =   0
         Top             =   2805
         Width           =   6120
         _ExtentX        =   10795
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
         Caption         =   "Keterangan"
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
      Begin BiSANumberBoxProject.BiSANumberBox nJumlah 
         Height          =   330
         Left            =   450
         TabIndex        =   1
         Top             =   2040
         Width           =   3405
         _ExtentX        =   6006
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
         Caption         =   "Jumlah Mutasi"
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
      Begin BiSADateProject.BiSADate dTgl 
         Height          =   330
         Left            =   450
         TabIndex        =   2
         Top             =   1680
         Width           =   2970
         _ExtentX        =   5239
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
         Caption         =   "Tanggal Mutasi"
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
      Begin BiSATextBoxProject.BiSATextBox cNamaKodeTransaksi 
         Height          =   330
         Left            =   2865
         TabIndex        =   3
         Top             =   960
         Width           =   3780
         _ExtentX        =   6668
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
         BackColor       =   12632256
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
      Begin BiSATextBoxProject.BiSABrowse cFaktur 
         Height          =   330
         Left            =   450
         TabIndex        =   4
         Top             =   600
         Width           =   4665
         _ExtentX        =   8229
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
         Caption         =   "No Faktur"
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
      Begin BiSATextBoxProject.BiSATextBox cUser 
         Height          =   330
         Left            =   450
         TabIndex        =   5
         Top             =   2415
         Width           =   2940
         _ExtentX        =   5186
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
         BackColor       =   12632256
         Enabled         =   0   'False
         Appearance      =   0
         Caption         =   "User Name"
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
      Begin BiSATextBoxProject.BiSATextBox cFullName 
         Height          =   330
         Left            =   3405
         TabIndex        =   6
         Top             =   2415
         Width           =   3195
         _ExtentX        =   5636
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
         BackColor       =   12632256
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
      Begin BiSADateProject.BiSADate dAwal 
         Height          =   330
         Left            =   435
         TabIndex        =   14
         Top             =   225
         Width           =   2925
         _ExtentX        =   5159
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
         Caption         =   "Antara Tgl"
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
      Begin BiSADateProject.BiSADate dAkhir 
         Height          =   330
         Left            =   3585
         TabIndex        =   15
         Top             =   225
         Width           =   1995
         _ExtentX        =   3519
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
      Begin BiSATextBoxProject.BiSATextBox cDK 
         Height          =   330
         Left            =   450
         TabIndex        =   16
         Top             =   1320
         Width           =   2130
         _ExtentX        =   3757
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
         BackColor       =   12632256
         Enabled         =   0   'False
         Appearance      =   0
         Caption         =   "D/K"
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
      Begin BiSATextBoxProject.BiSABrowse cKodeTransaksi 
         Height          =   330
         Left            =   450
         TabIndex        =   18
         Top             =   960
         Width           =   2415
         _ExtentX        =   4260
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
         Caption         =   "Kode Transaksi"
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
      Begin BiSATextBoxProject.BiSABrowse cRekeningJurnal 
         Height          =   330
         Left            =   2670
         TabIndex        =   17
         Top             =   1320
         Width           =   3240
         _ExtentX        =   5715
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
         BackColor       =   12632256
         Enabled         =   0   'False
         Appearance      =   0
         Caption         =   "Rekening"
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
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1590
      Left            =   0
      Top             =   0
      Width           =   7020
      _ExtentX        =   12383
      _ExtentY        =   2805
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
         Left            =   375
         TabIndex        =   7
         Top             =   465
         Width           =   5400
         _ExtentX        =   9525
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
         BackColor       =   12632256
         Enabled         =   0   'False
         Appearance      =   0
         Caption         =   "Nama Nasabah"
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
      Begin BiSATextBoxProject.BiSABrowse cAlamat 
         Height          =   330
         Left            =   375
         TabIndex        =   8
         Top             =   825
         Width           =   5400
         _ExtentX        =   9525
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
         BackColor       =   12632256
         Enabled         =   0   'False
         Appearance      =   0
         Caption         =   "Alamat Nasabah"
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
      Begin BiSANumberBoxProject.BiSANumberBox nAkhir 
         Height          =   330
         Left            =   375
         TabIndex        =   9
         Top             =   1185
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   582
         Appearance      =   0
         Enabled         =   0   'False
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12632256
         Caption         =   "Saldo Akhir"
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
      Begin BiSATextBoxProject.BiSABrowse cNoAnggota 
         Height          =   330
         Left            =   375
         TabIndex        =   13
         Top             =   120
         Width           =   3960
         _ExtentX        =   6985
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
         Caption         =   "No Anggota"
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
      Begin VB.Label Label1 
         Caption         =   "per tgl. "
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
         Left            =   3915
         TabIndex        =   10
         Top             =   1230
         Width           =   1920
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   645
      Left            =   0
      Top             =   5175
      Width           =   7020
      _ExtentX        =   12383
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
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   4725
         TabIndex        =   11
         Top             =   105
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   767
         Caption         =   "  &Save"
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
         Picture         =   "trKoreksiMutasiTabungan.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   5775
         TabIndex        =   12
         Top             =   105
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   767
         Caption         =   "     E&xit"
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
         Picture         =   "trKoreksiMutasiTabungan.frx":012C
      End
   End
End
Attribute VB_Name = "trKoreksiMutasiTabungan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim xarray As New XArrayDB
Dim vaView As New XArrayDB
Dim cSQL As String
Dim cRekening As String

Private Sub cFaktur_ButtonClick()
Dim cWhere As String
  
  cWhere = "And m.Tgl >= '" & Format(dAwal.Value, "yyyy-MM-dd") & "'"
  cWhere = cWhere & "And m.Tgl <='" & Format(dAkhir.Value, "yyyy-MM-dd") & "'"
  cWhere = cWhere & "And m.kodeanggota = '" & cRekening & "'"
  Set dbData = objData.Browse(GetDSN, "MutasiTabungan m", "m.nomormutasitabungan,m.Tgl,m.KodeTransaksi,m.Keterangan as KeteranganMutasi,k.Keterangan as NamaKodeTransaksi,m.Jumlah,m.UserName,u.Fullname,m.dk,m.kodeakun", "m.nomormutasitabungan", sisContent, cFaktur.Text, cWhere, "m.Tgl,m.nomormutasitabungan", _
                              Array("Left join kodetransaksi k on k.Kodetransaksi = m.KodeTransaksi", _
                                    "Left Join userName u on u.username=m.username"))
  cFaktur.Text = cFaktur.Browse(dbData, Array("No Trans", "Tgl", "Kode Trans", "Keterangan Mutasi", "Nama Kode Transaksi", "Jumlah", "User Name", "Fullname"), , Array(20, 10, 5, 40))
  If Not dbData.EOF Then
    cKodeTransaksi.Text = GetNull(dbData!KodeTransaksi)
    cNamaKodeTransaksi.Text = GetNull(dbData!NamaKodeTransaksi)
    dTgl.Value = GetNull(dbData!tgl)
    cRekeningJurnal.Text = GetNull(dbData!kodeakun)
    cDK.Text = GetNull(dbData!DK)
    nJumlah.Value = GetNull(dbData!Jumlah)
    cUser.Text = GetNull(dbData!UserName)
    cFullName.Text = GetNull(dbData!FullName)
    cKeterangan.Text = GetNull(dbData!KeteranganMutasi)
  End If
End Sub

Private Sub cKodeTransaksi_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "kodetransaksi", "kodetransaksi,keterangan,kodeakun,dk")
  If Not dbData.EOF Then
    cKodeTransaksi.Text = cKodeTransaksi.Browse(dbData)
    cNamaKodeTransaksi.Text = GetNull(dbData!Keterangan)
    cDK.Text = GetNull(dbData!DK)
    cRekeningJurnal.Text = GetNull(dbData!kodeakun)
    cKeterangan.Text = GetNull(dbData!Keterangan) & " an " & cNama.Text
  End If
End Sub

Private Sub cmdEdit_Click()
Dim nDebet As Double
Dim nKredit As Double
Dim lSave As Boolean
Dim cRekeningDebet As String
Dim cRekeningKredit As String

  lSave = True
  If ValidSaving Then
    If MsgBox("Data Benar-benar di Edit/Koreksi ?", vbQuestion + vbYesNo) = vbYes Then
      lSave = IIf(lSave, objData.Edit(GetDSN, "mutasitabungan", "nomormutasitabungan = '" & cFaktur.Text & "'", Array("kodetransaksi", "kodeakun", "dk", "tgl", "jumlah", "keterangan", "datetime", "username"), Array(cKodeTransaksi.Text, cRekeningJurnal.Text, cDK.Text, Format(dTgl.Value, "yyyy-MM-dd"), nJumlah.Value, cKeterangan.Text, SNow, GetRegistry(reg_UserName))), False)
      
      lSave = IIf(lSave, DelKodeTr(objData, msSimpananHarian, cFaktur.Text), False)
      If cDK.Text = "D" Then
        cRekeningDebet = aCfg(objData, msRekeningSimpananHarian)
        cRekeningKredit = cRekeningJurnal.Text
      Else
        cRekeningDebet = cRekeningJurnal.Text
        cRekeningKredit = aCfg(objData, msRekeningSimpananHarian)
      End If
        lSave = IIf(lSave, DelKodeTr(objData, msSimpananHarian, cFaktur.Text), False)
        lSave = IIf(lSave, UpdKodeTr(objData, msSimpananHarian, cFaktur.Text, Format(dTgl.Value, "yyyy-MM-dd"), cRekeningDebet, aCfg(objData, msCostCenterSimpanPinjam), cKeterangan.Text, nJumlah.Value, 0, cDK.Text, SNow), False)
            lSave = IIf(lSave, UpdKodeTr(objData, msSimpananHarian, cFaktur.Text, Format(dTgl.Value, "yyyy-MM-dd"), cRekeningKredit, aCfg(objData, msCostCenterSimpanPinjam), cKeterangan.Text, 0, nJumlah.Value, cDK.Text, SNow), False)

      MsgBox "Data sudah diEdit/Koreksi", vbInformation
      Exit Sub
    End If
  End If
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Function ValidSaving() As Boolean
  ValidSaving = True
  
  If Not CheckData(cFaktur.Text, "Faktur harus diisi!") Then
    ValidSaving = False
    cFaktur.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cKeterangan.Text, "Keterangan haris diisi..!") Then
    ValidSaving = False
    cKeterangan.SetFocus
    Exit Function
  End If
  
  If nJumlah.Value <= 0 Then
    MsgBox "Jumlah tidak valid", vbInformation + vbOKOnly
    ValidSaving = False
    nJumlah.SetFocus
    Exit Function
  End If
  
  
End Function

Private Sub GetData()
  cNama.Text = GetNull(dbData!nama, "")
  cAlamat.Text = GetNull(dbData!alamat, "")
  nAkhir.Value = GetSaldoTabungan(objData, cRekening, "01-01-1900", Date)
End Sub

Private Sub cNoAnggota_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "anggota", "kodeanggota,nama,alamat", "kodeanggota", sisContent, cNoAnggota.Text)
  If Not dbData.EOF Then
    cNoAnggota.Text = cNoAnggota.Browse(dbData, Array("No Anggota", "Nama", "Alamat"), , Array(10, 15, 20))
    cNama.Text = GetNull(dbData!nama)
    cAlamat.Text = GetNull(dbData!alamat)
    cRekening = GetNull(dbData!kodeanggota)
    nAkhir.Value = GetSaldoTabungan(objData, GetNull(dbData!kodeanggota), "01-01-1900", Date)
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd
  CenterForm Me
  Me.Top = 0
  InitValue
  dAwal.Value = Date
  dAkhir.Value = Date
  Label1.Caption = Label1.Caption & Format(Date, "dd-MM-yyyy")
  TabIndex cNoAnggota, n
  TabIndex dAwal, n
  TabIndex dAkhir, n
  TabIndex cFaktur, n
  TabIndex cKodeTransaksi, n
  TabIndex dTgl, n
  TabIndex nJumlah, n
  TabIndex cKeterangan, n
  TabIndex cmdEdit, n
  TabIndex cmdKeluar, n
End Sub

Private Sub InitValue()
  dAwal.Value = Date
  dAkhir.Value = Date
  cNama.Default
  cAlamat.Default
  nAkhir.Value = 0
  cFaktur.Default
  cKodeTransaksi.Default
  cNamaKodeTransaksi.Default
  dTgl.Value = Date
  nJumlah.Value = 0
  cUser.Default
  cFullName.Default
  cKeterangan.Default
End Sub


