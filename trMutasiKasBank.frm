VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trMutasiKasBank 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mutasi Kas dan Bank"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   9555
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   660
      Left            =   75
      Top             =   5295
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   1164
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
         Left            =   2235
         TabIndex        =   0
         Top             =   120
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   767
         Caption         =   "    &Delete"
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
         Picture         =   "trMutasiKasBank.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   6705
         TabIndex        =   1
         Top             =   105
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
         Picture         =   "trMutasiKasBank.frx":028A
      End
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   1185
         TabIndex        =   2
         Top             =   120
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   767
         Caption         =   "  &Edit"
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
         Picture         =   "trMutasiKasBank.frx":0429
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   435
         Left            =   105
         TabIndex        =   3
         Top             =   120
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   767
         Caption         =   "  &Add"
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
         Picture         =   "trMutasiKasBank.frx":0555
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   8235
         TabIndex        =   4
         Top             =   105
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
         Picture         =   "trMutasiKasBank.frx":0700
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   7155
         TabIndex        =   5
         Top             =   105
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
         Picture         =   "trMutasiKasBank.frx":07A6
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   5145
      Left            =   75
      Top             =   90
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   9075
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      BackColor       =   -2147483633
      Begin BiSANumberBoxProject.BiSANumberBox nTotal 
         Height          =   930
         Left            =   615
         TabIndex        =   13
         Top             =   2640
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   1640
         Appearance      =   0
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   32.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         Left            =   615
         TabIndex        =   6
         Top             =   210
         Width           =   2880
         _ExtentX        =   5080
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
      Begin BiSATextBoxProject.BiSABrowse cFaktur 
         Height          =   330
         Left            =   615
         TabIndex        =   7
         Top             =   570
         Width           =   3750
         _ExtentX        =   6615
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
         Caption         =   "Nomor"
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
      Begin BiSATextBoxProject.BiSABrowse cAkunKas 
         Height          =   330
         Left            =   615
         TabIndex        =   8
         Top             =   1950
         Width           =   1365
         _ExtentX        =   2408
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
      Begin BiSATextBoxProject.BiSABrowse cNamaAkunKas 
         Height          =   330
         Left            =   1995
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1950
         Width           =   2505
         _ExtentX        =   4419
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
      Begin BiSATextBoxProject.BiSABrowse cCostCenter 
         Height          =   330
         Left            =   615
         TabIndex        =   10
         Top             =   930
         Width           =   3090
         _ExtentX        =   5450
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
      Begin BiSATextBoxProject.BiSABrowse cAkunTujuan 
         Height          =   330
         Left            =   4590
         TabIndex        =   11
         Top             =   1950
         Width           =   1440
         _ExtentX        =   2540
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
      Begin BiSATextBoxProject.BiSABrowse cNamaAkunTujuan 
         Height          =   330
         Left            =   6060
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1950
         Width           =   2505
         _ExtentX        =   4419
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
      Begin BiSATextBoxProject.BiSABrowse cKeterangan 
         Height          =   330
         Left            =   600
         TabIndex        =   14
         Top             =   4035
         Width           =   7950
         _ExtentX        =   14023
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
      Begin VB.Label Label4 
         Caption         =   "Keterangan"
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
         Left            =   600
         TabIndex        =   18
         Top             =   3735
         Width           =   1350
      End
      Begin VB.Label Label3 
         Caption         =   "Sejumlah"
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
         Left            =   615
         TabIndex        =   17
         Top             =   2370
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Ke Akun Kas"
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
         Left            =   4590
         TabIndex        =   16
         Top             =   1650
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Dari Akun Kas"
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
         Left            =   630
         TabIndex        =   15
         Top             =   1650
         Width           =   1455
      End
   End
End
Attribute VB_Name = "trMutasiKasBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nPos As SisPos
Dim lEdit As Boolean
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub GetFakturBrowse(ByVal lStat As Boolean)
  cFaktur.Button = lStat
End Sub

Private Sub cAkunKas_ButtonClick()
Set dbData = objData.Browse(GetDSN, "akun", "kodeakun,keterangan", "kodeakun", sisContent, cAkunKas.Text, " AND jenis = 'D' AND (left(kodeakun,1) = 1 OR left(kodeakun,1)=3)")
  If Not dbData.EOF Then
    cAkunKas.Text = cAkunKas.Browse(dbData, Array("Kode Akun", "Keterangan"), , Array(15, 25))
    cNamaAkunKas.Text = GetNull(dbData!keterangan)
    nTotal.value = GetSaldoKasBank(objData, cAkunKas.Text, dTgl.value)
  End If
End Sub

Private Function GetSaldoKasBank(ByVal obj As CodeSuiteLibrary.Data, ByVal koderek As String, ByVal tgl As Date) As Double
Dim db As New ADODB.Recordset

  GetSaldoKasBank = 0
  Set db = obj.Browse(GetDSN, "bukubesar", "sum(debet-kredit) as saldo", "kodeakun", sisAssign, koderek, " and tgl <= '" & Format(tgl, "yyyy-MM-dd") & "'")
  If Not db.EOF Then
    GetSaldoKasBank = GetNull(db!saldo)
  End If
End Function

Private Sub cAkunKas_Validate(Cancel As Boolean)
  nTotal.value = GetSaldoKasBank(objData, cAkunKas.Text, dTgl.value)
  GetKeterangan
End Sub

Private Sub GetKeterangan()
  cKeterangan.Text = "Setoran Tgl " & Format(dTgl.value, "ddMMyyyy")
End Sub

Private Sub cAkunTujuan_ButtonClick()
Set dbData = objData.Browse(GetDSN, "akun", "kodeakun,keterangan", "kodeakun", sisContent, cAkunTujuan.Text, " AND jenis = 'D' AND (left(kodeakun,1) = 1 OR left(kodeakun,1)=3)")
  If Not dbData.EOF Then
    cAkunTujuan.Text = cAkunTujuan.Browse(dbData, Array("Kode Akun", "Keterangan"), , Array(15, 25))
    cNamaAkunTujuan.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub cCostCenter_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "costcenter", "kodecostcenter,keterangan")
  If Not dbData.EOF Then
    cCostCenter.Text = cCostCenter.Browse(dbData, Array("Kode", "Keterangan"), , Array(15, 25))
  End If
End Sub

Private Sub cFaktur_ButtonClick()
Dim db As New ADODB.Recordset
Dim n As Integer
Dim lSave As Boolean
Dim objMenu As New CodeSuiteLibrary.Menu

  If aCfg(objData, msOtorisasiPenuh) = "Y" Then
    If GetRegistry(reg_UserLevel) <> 0 Then
      If objMenu.GetPassword("", Me, GetDSN) Then
        If objMenu.UserLevel <> 0 Then
            MsgBox "Maaf, Anda tidak diberikan wewenang untuk melakukan PENGEDITAN." & vbCrLf & _
                   "Hanya user dengan LEVEL 0 (SUPERVISOR) yg diijinkan", vbInformation, "OTORISASI not ALLOWED"
            Exit Sub
'        Else
'          MsgBox "OTORISASI DIBATALKAN", vbCritical
'          Exit Sub
        End If
      Else
        Exit Sub
      End If
    End If
  End If

  lSave = True
  vaArray.ReDim 0, -1, 0, 4

  Set dbData = objData.Browse(GetDSN, "mutasikasbank m", "m.nomormutasikasbank,m.dariakun,m.keakun,m.total,m.kodecostcenter,a.keterangan as namaakundari,m.keterangan,b.keterangan as namaakunke", "m.tgl", sisAssign, Format(dTgl.value, "yyyy-MM-dd"), , "m.nomormutasikasbank desc", Array("LEFT JOIN akun a ON a.kodeakun = m.dariakun", "LEFT JOIN akun b ON b.kodeakun = m.keakun"))
  If Not dbData.EOF Then
    cFaktur.Text = cFaktur.Browse(dbData)
    cAkunKas.Text = GetNull(dbData!dariakun)
    cNamaAkunKas.Text = GetNull(dbData!namaakundari)
    cAkunTujuan.Text = GetNull(dbData!keakun)
    cNamaAkunTujuan.Text = GetNull(dbData!namaakunke)
    cCostCenter.Text = GetNull(dbData!kodecostcenter)
    nTotal.value = GetNull(dbData!Total)
    cKeterangan.Text = GetNull(dbData!keterangan)

    If nPos = Delete Then
      If MsgBox("Data akan dihapus?", vbYesNo) = vbYes Then
        objData.Start GetDSN
        lSave = IIf(lSave, DelKodeTr(objData, msMutasiKasBank, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "mutasikasbank", "nomormutasikasbank", sisAssign, cFaktur.Text), False)
        If lSave Then
          objData.Save GetDSN
        Else
          objData.Cancel GetDSN
        End If
      End If
      GetEdit False
      initvalue
    End If
    If nPos = Edit Then
      SendKeysA vbKeyReturn, True
    End If
  End If
End Sub

Private Sub cmdAktivasi_Click()
  frmAktivasi.Action Me
End Sub

Private Sub cmdAdd_Click()
  nPos = Add
  GetEdit True
  GetFakturBrowse False
  cFaktur.Text = CreateNomorFaktur(objData, sisModulTransaksi.MutasiKasBank, "mutasikasbank", "nomormutasikasbank")
End Sub

Private Sub cmdEdit_Click()
  nPos = Edit
  GetEdit True
  GetFakturBrowse True
End Sub

Private Sub cmdHapus_Click()
  nPos = Delete
  GetEdit True
  GetFakturBrowse True
End Sub

Private Sub cmdKeluar_Click()
  If lEdit Then
    GetEdit False
    initvalue
  Else
    Unload Me
  End If
End Sub

Private Sub cmdSimpan_Click()
Dim lSave As Boolean
Dim Faktur As String
Dim n As Integer
Dim i As Integer
lSave = True

  If isValidSaving Then
    objData.Start GetDSN
    Faktur = cFaktur.Text
    lSave = IIf(lSave, objData.Update(GetDSN, "mutasikasbank", "nomormutasikasbank = '" & Faktur & "'", Array("nomormutasikasbank", "dariakun", "keakun", "total", "tgl", "datetime", "username", "kodecostcenter", "debet", "kredit", "keterangan"), Array(Faktur, cAkunKas.Text, cAkunTujuan.Text, nTotal.value, Format(dTgl.value, "yyyy-MM-dd"), SNow, GetRegistry(reg_Username), cCostCenter.Text, IIf(Left(cAkunKas.Text, 1) = 3, nTotal.value, 0), IIf(Left(cAkunKas.Text, 1) = 1, nTotal.value, 0), cKeterangan.Text)), False)
    lSave = IIf(lSave, objData.Delete(GetDSN, "bukubesar", "faktur", sisAssign, Faktur), False)
    
    lSave = IIf(lSave, UpdKodeTr(objData, msMutasiKasBank, Faktur, Format(dTgl.value, "yyyy-MM-dd"), cAkunKas.Text, cCostCenter.Text, cKeterangan.Text, 0, nTotal.value, "K", SNow), False)
    lSave = IIf(lSave, UpdKodeTr(objData, msMutasiKasBank, Faktur, Format(dTgl.value, "yyyy-MM-dd"), cAkunTujuan.Text, cCostCenter.Text, cKeterangan.Text, nTotal.value, 0, "K", SNow), False)
    
    If lSave Then
      objData.Save GetDSN
    Else
      objData.Cancel GetDSN
    End If
    
    'print struk
    trPrintMutasiKasBank.noOrder = Faktur
    
      Set dbData = objData.Browse(GetDSN, "mutasikasbank t", "t.*", "t.nomormutasikasbank", sisAssign, Faktur)
      If Not dbData.EOF Then

'        trPrintMutasiKasBank.nSubTotal = GetNull(dbData!Subtotal)
'        trPrintMutasiKasBank.nDiscount = GetNull(dbData!dp)
'        trPrintMutasiKasBank.nCash = GetNull(dbData!Tunai)
'        trPrintMutasiKasBank.nChange = GetNull(dbData!Piutang)
'        trPrintMutasiKasBank.cKodeMember = GetNull(dbData!kodeanggota)
'        trPrintMutasiKasBank.cMember = GetNull(dbData!nama)
'        trPrintMutasiKasBank.cTeleponMember = GetNull(dbData!telp)
'        trPrintMutasiKasBank.Ups = GetNull(dbData!upkepada)

        Load trPrintMutasiKasBank
        trPrintMutasiKasBank.Show vbModal
      End If
      
    initvalue
    GetEdit False
  End If
End Sub

Private Sub PrintThermal(ByVal Faktur As String)
Dim n As Double
Dim nBruto As Double
Dim nTotQty As Double
Dim nHargaArray As Double


    Open "lpt1" For Output As #1
    Print #1, Chr(27); Chr(33); Chr(4);
    Print #1, Chr(27) & Chr(97) & Chr(1)
    Print #1, "STRUK SETORAN"
    Print #1, aCfg(objData, msNamaPerusahaan)
    Print #1, aCfg(objData, msAlamatPerusahaan)
    Print #1, ""
    Select Case GetRegistry(reg_AlignmentThermal)
      Case 1 ' rata kiri
                Print #1, Chr(27) & Chr(97) & Chr(0)
      Case 2 ' rata kanan
                Print #1, Chr(27) & Chr(97) & Chr(2)
    End Select
    Print #1, "No. " & Faktur
    Print #1, Format(Now, "dd-MM-yyyy HH:MM:SS")
    Print #1, ""

'    Print #1, "KASIR: "; cCustomer.Text; ""
'    Print #1, cNamaCustomer.Text
'    Print #1, "Telp. "; cTelp
    Print #1, ""

    Print #1, Replicate("-", 27)
    Print #1, Padl("Qty", 6); Padl("Hrg Net", 11); Padl("Jml", 10)
    Print #1, Replicate("-", 27)
    For n = 0 To vaArray.UpperBound(1)
      If vaArray(n, 3) <> 0 Then
        nHargaArray = vaArray(n, 5) - (vaArray(n, 5) * vaArray(n, 6) / 100)
        nBruto = nBruto + (vaArray(n, 3) * nHargaArray)
        nTotQty = nTotQty + vaArray(n, 3)
        Print #1, vaArray(n, 2) ' vaArray(n, 1) kolom REF barang ditiadakan karena terlalu panjang
        If vaArray(n, 6) <> 0 Then
          Print #1, vaArray(n, 1) & " Rp." & Format(vaArray(n, 5), "#,##0") & " -" & vaArray(n, 6) & "%"
        End If
        Print #1, Padl(Format(vaArray(n, 3), "#,##0"), 3) & " x " & Padl(Format(nHargaArray, "#,###,##0"), 8) & " = " & Padl(Format(vaArray(n, 3) * nHargaArray, "#,###,##0"), 10)
      End If
    Next
    
    Print #1, Chr(10) ' feed kertas
    Print #1, ""
    Print #1, ""
'    Print #1, ""
    Close #1
End Sub

Private Function isValidSaving() As Boolean
isValidSaving = True
  
  If Trim(cFaktur.Text) = "" Then
    MsgBox "Nomor transaksi tidak boleh kosong"
    isValidSaving = False
    Exit Function
  End If
  
  If Trim(cAkunKas.Text) = "" Then
    MsgBox "Akun kas tidak boleh kosong"
    isValidSaving = False
    Exit Function
  End If
  
  If Trim(cAkunTujuan.Text) = "" Then
    MsgBox "Akun kas tidak boleh kosong"
    isValidSaving = False
    Exit Function
  End If
  
  If Trim(cCostCenter.Text) = "" Then
    MsgBox "Cost Center tidak boleh kosong"
    isValidSaving = False
    Exit Function
  End If
  
  If Trim(cKeterangan.Text) = "" Then
    MsgBox "Keterangan mutasi tidak boleh kosong!! Transaksi tidak bisa dilanjutkan."
    isValidSaving = False
    Exit Function
  End If
  
  If Not GetValidDataBrowse(objData, "costcenter", "kodecostcenter", cCostCenter.Text) Then
    MsgBox "Data cost center tidak benar"
    isValidSaving = False
    Exit Function
  End If
  
  If Not GetValidDataBrowse(objData, "akun", "kodeakun", cAkunKas.Text) Then
    MsgBox "Akun kas tidak ada dalam database" & vbCrLf & "data tidak bisa disimpan"
    isValidSaving = False
    Exit Function
  End If
  
  If Not GetValidDataBrowse(objData, "akun", "kodeakun", cAkunTujuan.Text) Then
    MsgBox "Akun kas tidak ada dalam database" & vbCrLf & "data tidak bisa disimpan"
    isValidSaving = False
    Exit Function
  End If
End Function

Private Sub dTgl_Validate(Cancel As Boolean)
  If Not IsInPeriod(dTgl.value) Or (dTgl.value > Date) Then
    Cancel = True
    dTgl.SetFocus
    GetEdit False
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

'  If CheckTrial(nRecordsTrial, TrialPengeluaranBiaya) = True Then
'    End
'  End If

  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  initvalue
  GetEdit False
  
  
  TabIndex dTgl, n
  TabIndex cFaktur, n
  TabIndex cCostCenter, n
  TabIndex cAkunKas, n
  TabIndex cAkunTujuan, n
  TabIndex nTotal, n
  TabIndex cKeterangan, n
    
  TabIndex cmdAdd, n
  TabIndex cmdEdit, n
  TabIndex cmdHapus, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  TabIndex cmdAktivasi, n
  cAkunKas.BackColor = vbWhite
  cAkunKas.Enabled = True
  cCostCenter.BackColor = vbWhite
  cCostCenter.Enabled = True
  If GetRegistry(reg_UserLevel) <> 0 Then
    cAkunKas.BackColor = vbButtonFace
    cAkunKas.Enabled = False
    cCostCenter.BackColor = vbButtonFace
    cCostCenter.Enabled = False
    cCostCenter.Text = GetCostCenterUser(objData, GetRegistry(reg_Username))
  End If
End Sub

Private Sub initvalue()
  
  cFaktur.Default
  dTgl.value = Date
  
'  Set dbData = objData.Browse(GetDSN, "costcenter", , "kodecostcenter", sisAssign, aCfg(objData, msCostCenterJualBeli))
'  If Not dbData.EOF Then
'    cCostCenter.Text = GetNull(dbData!kodecostcenter)
'  End If
  cCostCenter.Text = GetCostCenterUser(objData, GetRegistry(reg_Username))
  
  cAkunKas.Text = cKasTeller
  cNamaAkunKas.Text = cNamaKasTeller
  cAkunTujuan.Default
  cNamaAkunTujuan.Default
  nTotal.Default
  cKeterangan.Default
  
  'set default akun tujuan setoran
  cAkunTujuan.Text = aCfg(objData, msRekeningSetoranKas)
  Set dbData = objData.Browse(GetDSN, "akun", "kodeakun,keterangan", "kodeakun", sisAssign, cAkunTujuan.Text)
  If Not dbData.EOF Then
    cNamaAkunTujuan.Text = GetNull(dbData!keterangan)
  End If
  
  'overide if user already has kas tujuan
  Set dbData = objData.Browse(GetDSN, "akunkas a", "a.akunsetoran,k.keterangan", "a.username", sisAssign, GetRegistry(reg_Username), , , Array("left join akun k on k.kodeakun = a.akunsetoran"))
  If Not dbData.EOF Then
    cAkunTujuan.Text = GetNull(dbData!akunsetoran)
    cNamaAkunTujuan.Text = GetNull(dbData!keterangan)
  End If
  cAkunTujuan.Enabled = IIf(aCfg(objData, msKunciRekeningSetoranKas) = "Y", False, True)
  nTotal.value = GetSaldoKasBank(objData, cAkunKas.Text, dTgl.value)
End Sub

Private Sub GetEdit(lPar As Boolean)
  BiSAFrame1.Enabled = lPar
  lEdit = lPar
  initvalue
  SetButton cmdSimpan, cmdKeluar, cmdAdd, cmdEdit, cmdHapus, nPos, lPar, cmdAktivasi
  
  If lPar Then
    dTgl.SetFocus
    If nPos = Add Then
      cFaktur.Enabled = False
      cFaktur.BackColor = vbButtonFace
    Else
      cFaktur.Enabled = True
      cFaktur.BackColor = vbWindowBackground
      cFaktur.CaptionBackColor = vbButtonFace
    End If
  End If
End Sub

