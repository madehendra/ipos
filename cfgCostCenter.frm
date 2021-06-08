VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Begin VB.Form cfgCostCenter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cost Centre Settings..."
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   8895
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   2175
      Left            =   0
      Top             =   0
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   3836
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
      Begin BiSATextBoxProject.BiSABrowse cCostCenterJualBeli 
         Height          =   330
         Left            =   135
         TabIndex        =   0
         Top             =   105
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
         Caption         =   "Cost Center Default"
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
      Begin BiSATextBoxProject.BiSATextBox cNamaCostCenterJualBeli 
         Height          =   330
         Left            =   4650
         TabIndex        =   1
         Top             =   105
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
      Begin BiSATextBoxProject.BiSABrowse cGudangPembelian 
         Height          =   330
         Left            =   135
         TabIndex        =   5
         Top             =   450
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
         Caption         =   "Gudang Pembelian"
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
      Begin BiSATextBoxProject.BiSATextBox cNamaGudangPembelian 
         Height          =   330
         Left            =   4650
         TabIndex        =   6
         Top             =   450
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
      Begin BiSATextBoxProject.BiSABrowse cGudangPenjualan 
         Height          =   330
         Left            =   135
         TabIndex        =   7
         Top             =   795
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
         Caption         =   "Gudang Penjualan"
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
      Begin BiSATextBoxProject.BiSATextBox cNamaGudangPenjualan 
         Height          =   330
         Left            =   4650
         TabIndex        =   8
         Top             =   795
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
      Begin BiSATextBoxProject.BiSABrowse cGudangPenyimpanan 
         Height          =   330
         Left            =   135
         TabIndex        =   9
         Top             =   1140
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
         Caption         =   "Gudang Penyimpanan"
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
      Begin BiSATextBoxProject.BiSATextBox cNamaGudangPenyimpanan 
         Height          =   330
         Left            =   4650
         TabIndex        =   10
         Top             =   1140
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
      Left            =   0
      Top             =   2160
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
         Picture         =   "cfgCostCenter.frx":0000
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
         Picture         =   "cfgCostCenter.frx":00A6
      End
      Begin VB.Label Label1 
         Caption         =   "F2 = SIMPAN"
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
         Left            =   180
         TabIndex        =   4
         Top             =   180
         Width           =   1815
      End
   End
End
Attribute VB_Name = "cfgCostCenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data

Private Sub cCostCenterJualBeli_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "costcenter", "kodecostcenter,keterangan", , , , , "kodecostcenter")
  If Not dbData.EOF Then
    cCostCenterJualBeli.Text = cCostCenterJualBeli.Browse(dbData)
    cNamaCostCenterJualBeli.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub cGudangPembelian_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "gudang")
  If Not dbData.EOF Then
    cGudangPembelian.Text = cGudangPembelian.Browse(dbData)
    cNamaGudangPembelian.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub cGudangPenjualan_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "gudang")
  If Not dbData.EOF Then
    cGudangPenjualan.Text = cGudangPenjualan.Browse(dbData)
    cNamaGudangPenjualan.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub cGudangPenyimpanan_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "gudang")
  If Not dbData.EOF Then
    cGudangPenyimpanan.Text = cGudangPenyimpanan.Browse(dbData)
    cNamaGudangPenyimpanan.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdSimpan_Click()
  GetNotifikasiAdd "Data lagi disimpan"
  objData.Start GetDSN
'  UpdCfg msCostCenterJualBeli, cCostCenterJualBeli.Text, objData, cCostCenterJualBeli.Caption, Me.Caption
  UpdCfg msGudangPembelian, cGudangPembelian.Text, objData, cGudangPembelian.Caption, Me.Caption
  UpdCfg msGudangPenjualan, cGudangPenjualan.Text, objData, cGudangPenjualan.Caption, Me.Caption
  UpdCfg msGudangPenyimpanan, cGudangPenyimpanan.Text, objData, cNamaGudangPenyimpanan.Caption, Me.Caption
  objData.Save GetDSN
  GetNotifikasiRemove
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  SetIcon Me.hWnd
  TabIndex cCostCenterJualBeli, n
  TabIndex cGudangPembelian, n
  TabIndex cGudangPenjualan, n
  TabIndex cGudangPenyimpanan, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  
  'cCostCenterJualBeli.Text = aCfg(objData, msCostCenterJualBeli)
  cNamaCostCenterJualBeli.Text = GetNamaRekening(cCostCenterJualBeli.Text)

  cGudangPembelian.Text = aCfg(objData, msGudangPembelian)
  cNamaGudangPembelian.Text = GetNamaGudang(cGudangPembelian.Text)
  cGudangPenjualan.Text = aCfg(objData, msGudangPenjualan)
  cNamaGudangPenjualan.Text = GetNamaGudang(cGudangPenjualan.Text)
  cGudangPenyimpanan.Text = aCfg(objData, msGudangPenyimpanan)
  cNamaGudangPenyimpanan.Text = GetNamaGudang(cGudangPenyimpanan.Text)
End Sub

Private Function GetNamaRekening(cAkun As String) As String
  GetNamaRekening = ""
  Set dbData = objData.Browse(GetDSN, "costcenter", "keterangan", "kodecostcenter", sisAssign, cAkun)
  If Not dbData.EOF Then
    GetNamaRekening = GetNull(dbData!keterangan, "")
  End If
End Function

Private Function GetNamaGudang(cAkun As String) As String
  GetNamaGudang = ""
  Set dbData = objData.Browse(GetDSN, "gudang", "keterangan", "kodegudang", sisAssign, cAkun)
  If Not dbData.EOF Then
    GetNamaGudang = GetNull(dbData!keterangan, "")
  End If
End Function

