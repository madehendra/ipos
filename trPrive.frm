VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trPrive 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form Prive"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   7800
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   585
      Left            =   0
      Top             =   3495
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   1032
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
         Left            =   2220
         TabIndex        =   0
         Top             =   75
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
         Picture         =   "trPrive.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   5115
         TabIndex        =   1
         Top             =   75
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
         Picture         =   "trPrive.frx":028A
      End
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   1170
         TabIndex        =   2
         Top             =   75
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
         Picture         =   "trPrive.frx":0429
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   435
         Left            =   90
         TabIndex        =   3
         Top             =   75
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
         Picture         =   "trPrive.frx":0555
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   6645
         TabIndex        =   4
         Top             =   75
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
         Picture         =   "trPrive.frx":0700
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   5565
         TabIndex        =   5
         Top             =   75
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
         Picture         =   "trPrive.frx":07A6
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   3510
      Left            =   0
      Top             =   0
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   6191
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
      Begin BiSANumberBoxProject.BiSANumberBox nTotal 
         Height          =   330
         Left            =   4500
         TabIndex        =   6
         Top             =   1905
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   582
         Appearance      =   0
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
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
         TabIndex        =   7
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
         TabIndex        =   8
         Top             =   525
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
      Begin BiSATextBoxProject.BiSABrowse cCostCenter 
         Height          =   330
         Left            =   615
         TabIndex        =   9
         Top             =   840
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
      Begin BiSATextBoxProject.BiSABrowse cAkunKas 
         Height          =   330
         Left            =   615
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1905
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
         TabIndex        =   11
         Top             =   2685
         Width           =   6465
         _ExtentX        =   11404
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
      Begin VB.Label Label2 
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
         Height          =   210
         Left            =   630
         TabIndex        =   14
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Prive dari atau ke akun Kas?"
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
         TabIndex        =   13
         Top             =   1650
         Width           =   2700
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
         Left            =   4530
         TabIndex        =   12
         Top             =   1635
         Width           =   1455
      End
   End
End
Attribute VB_Name = "trPrive"
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
Dim objMenu As New CodeSuiteLibrary.Menu

Private Sub GetFakturBrowse(ByVal lStat As Boolean)
  cFaktur.Button = lStat
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

  lSave = True
  vaArray.ReDim 0, -1, 0, 4

  Set dbData = objData.Browse(GetDSN, "prive m", "m.nomorprive,m.akunkas,m.akunprive,m.total,m.kodecostcenter,m.keterangan", "m.tgl", sisAssign, Format(dTgl.Value, "yyyy-MM-dd"))
  If Not dbData.EOF Then
    cFaktur.Text = cFaktur.Browse(dbData)
    cAkunKas.Text = GetNull(dbData!akunkas)
    nTotal.Value = GetNull(dbData!Total)
    cKeterangan.Text = GetNull(dbData!keterangan)

    If nPos = Delete Then
      If MsgBox("Data akan dihapus?", vbYesNo) = vbYes Then
        objData.Start GetDSN
        lSave = IIf(lSave, DelKodeTr(objData, msPrive, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "prive", "nomorprive", sisAssign, cFaktur.Text), False)
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
  cFaktur.Text = CreateNomorFaktur(objData, sisModulTransaksi.Prive, "prive", "nomorprive")

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

    lSave = IIf(lSave, objData.Update(GetDSN, "prive", "nomorprive = '" & Faktur & "'", Array("nomorprive", "akunkas", "akunprive", "total", "tgl", "datetime", "username", "kodecostcenter", "keterangan"), Array(Faktur, cAkunKas.Text, aCfg(objData, msRekeningPrive), nTotal.Value, Format(dTgl.Value, "yyyy-MM-dd"), SNow, GetRegistry(reg_username), cCostCenter.Text, cKeterangan.Text)), False)
    lSave = IIf(lSave, objData.Delete(GetDSN, "bukubesar", "faktur", sisAssign, Faktur), False)

    lSave = IIf(lSave, UpdKodeTr(objData, msPrive, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), cAkunKas.Text, cCostCenter.Text, cKeterangan.Text, nTotal.Value, 0, "K", SNow), False)
    lSave = IIf(lSave, UpdKodeTr(objData, msPrive, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningPrive), cCostCenter.Text, cKeterangan.Text, 0, nTotal.Value, "K", SNow), False)
        
    If lSave Then
      objData.Save GetDSN
    Else
      objData.Cancel GetDSN
    End If
        
    initvalue
    GetEdit False
  End If
End Sub

Private Function isValidSaving() As Boolean
isValidSaving = True
  
  If nTotal.Value < 0 Then
    If GetRegistry(reg_UserLevel) <> 0 Then
      If objMenu.GetPassword("", Me, GetDSN) Then
        If objMenu.UserLevel <> 0 Then
          Me.Hide
          MsgBox "Maaf, Anda tidak diberikan wewenang untuk melakukan otorisasi." & vbCrLf & _
                 "Hanya user dengan LEVEL 0 (SUPERVISOR) yg diijinkan", vbInformation, "OTORISASI not ALLOWED"
          isValidSaving = False
          Me.Show
        End If
      Else
        Unload Me
      End If
    End If
  End If
  
  If Trim(cFaktur.Text) = "" Then
    MsgBox "Nomor transaksi tidak boleh kosong"
    isValidSaving = False
    Exit Function
  End If

  If Trim(cCostCenter.Text) = "" Then
    MsgBox "Cost Center tidak boleh kosong"
    isValidSaving = False
    Exit Function
  End If
  
End Function

Private Sub dTgl_Validate(Cancel As Boolean)
  If Not IsInPeriod(dTgl.Value) Or (dTgl.Value > Date) Then
    Cancel = True
    dTgl.SetFocus
    GetEdit False
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  initvalue
  GetEdit False
  
  
  TabIndex dTgl, n
  TabIndex cFaktur, n
  TabIndex cCostCenter, n
  TabIndex cAkunKas, n
  
  
  TabIndex nTotal, n
  TabIndex cKeterangan, n
    
  TabIndex cmdAdd, n
  TabIndex cmdEdit, n
  TabIndex cmdHapus, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  TabIndex cmdAktivasi, n
  
End Sub

Private Sub initvalue()
  
  cFaktur.Default
  dTgl.Value = Date
  
  Set dbData = objData.Browse(GetDSN, "costcenter", , "kodecostcenter", sisAssign, GetCostCenterUser(objData, GetRegistry(reg_username)))
  If Not dbData.EOF Then
    cCostCenter.Text = GetNull(dbData!kodecostcenter)
  End If
  
  cAkunKas.Text = cKasTeller
  nTotal.Default
  cKeterangan.Default
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
