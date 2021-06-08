VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Begin VB.Form trCetakStikerBarcode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cetak Stiker Barcode"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   5745
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   645
      Left            =   0
      Top             =   4260
      Width           =   5730
      _ExtentX        =   10107
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
      Begin BiSAButtonProject.BiSAButton BiSAButton1 
         Height          =   315
         Left            =   1695
         TabIndex        =   11
         Top             =   150
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         Caption         =   "Label1"
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
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   4590
         TabIndex        =   0
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
         Picture         =   "trCetakStikerBarcode.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   4155
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
         Picture         =   "trCetakStikerBarcode.frx":00A6
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   4260
      Left            =   0
      Top             =   0
      Width           =   5730
      _ExtentX        =   10107
      _ExtentY        =   7514
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
      Begin VB.OptionButton opt 
         Caption         =   "&2 Printer Biasa/Jet"
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
         Index           =   1
         Left            =   3600
         TabIndex        =   3
         Top             =   3825
         Value           =   -1  'True
         Width           =   1980
      End
      Begin VB.OptionButton opt 
         Caption         =   "&1 Printer Barcode (Sato 208)"
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
         Index           =   0
         Left            =   465
         TabIndex        =   2
         Top             =   3810
         Width           =   2865
      End
      Begin BiSANumberBoxProject.BiSANumberBox nQty 
         Height          =   330
         Left            =   285
         TabIndex        =   4
         Top             =   2295
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   582
         Appearance      =   0
         Decimals        =   0
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Qty"
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
      Begin BiSATextBoxProject.BiSATextBox cSatuan 
         Height          =   330
         Left            =   285
         TabIndex        =   5
         Top             =   1500
         Width           =   2670
         _ExtentX        =   4710
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
         Caption         =   "Satuan"
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
         Left            =   285
         TabIndex        =   6
         Top             =   870
         Width           =   2835
         _ExtentX        =   5001
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
         Caption         =   "Kode"
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
      Begin BiSATextBoxProject.BiSABrowse cNama 
         Height          =   330
         Left            =   285
         TabIndex        =   7
         Top             =   1185
         Width           =   4680
         _ExtentX        =   8255
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
         Caption         =   "Nama"
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
      Begin BiSATextBoxProject.BiSATextBox cBarcode 
         Height          =   330
         Left            =   285
         TabIndex        =   8
         Top             =   1815
         Width           =   2670
         _ExtentX        =   4710
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
         Caption         =   "Barcode"
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
      Begin BiSANumberBoxProject.BiSANumberBox nHarga 
         Height          =   330
         Left            =   330
         TabIndex        =   9
         Top             =   2700
         Width           =   3240
         _ExtentX        =   5715
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
         Caption         =   "Harga"
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
      Begin BiSANumberBoxProject.BiSANumberBox nHargaLama 
         Height          =   330
         Left            =   330
         TabIndex        =   12
         Top             =   3105
         Width           =   3240
         _ExtentX        =   5715
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
         Caption         =   "Lama"
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
         Caption         =   "PENCETAKAN STIKER BARCODE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   315
         TabIndex        =   10
         Top             =   360
         Width           =   3090
      End
   End
End
Attribute VB_Name = "trCetakStikerBarcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub BiSAButton1_Click()
  Load cfgStikerBarcode
  cfgStikerBarcode.Show
End Sub

Private Sub cKode_ButtonClick()
  If Trim(cKode.Text) <> "" Then
    Set dbData = objData.Browse(GetDSN, "stock2", , "kodestock", sisContent, cKode.Text)
    If dbData.RecordCount > 0 Then
      GetStockMemory
    End If
  End If
End Sub

Private Sub GetStockMemory()
Dim dbSaldo As New ADODB.Recordset

  cKode.Text = GetNull(dbData!KodeStock)
  cNama.Text = GetNull(dbData!nama)
  cSatuan.Text = GetNull(dbData!kodesatuan)
  cBarcode.Text = GetNull(dbData!barcode)
  nHarga.Value = GetNull(dbData!hargajual)
  nHargaLama.Value = GetNull(dbData!hargalama)

'  Set dbSaldo = objData.Browse(GetDSN, "kartustock", "sum(debet - kredit) as Saldo", "kodestock", sisAssign, cKode.Text)
'  nQty.Value = GetNull(dbSaldo!saldo)
End Sub

Private Sub cKode_Validate(Cancel As Boolean)
  cKode_ButtonClick
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
Dim n As Double
  vaArray.ReDim 0, -1, 0, 2
  For n = 0 To (nQty.Value) - 1
    vaArray.InsertRows n
    vaArray(n, 0) = Format(nHarga.Value, "###,###,###,###,#.00")
    vaArray(n, 1) = cNama.Text 'barcode(cBarcode.Text)
    vaArray(n, 2) = Format(nHargaLama.Value, "###,###,###,###,#.00") 'cKode.Text
  Next
  cfgStikerBarcode.PrintBarcode vaArray, GetOpt(opt)
End Sub

Private Sub cNama_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "stock2", , "nama", sisContent, cNama.Text)
  If dbData.RecordCount > 0 Then
    cNama.Text = cNama.Browse(dbData)
    GetStockMemory
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  TabIndex cKode, n
  TabIndex cNama, n
  TabIndex nQty, n
  TabIndex nHarga, n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub

