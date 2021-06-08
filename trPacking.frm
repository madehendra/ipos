VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trPacking 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Packing Stok - Inventory"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   8505
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   585
      Left            =   0
      Top             =   4110
      Width           =   8490
      _ExtentX        =   14975
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
         Picture         =   "trPacking.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   5820
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
         Picture         =   "trPacking.frx":028A
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
         Picture         =   "trPacking.frx":0429
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   435
         Left            =   105
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
         Picture         =   "trPacking.frx":0555
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   7350
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
         Picture         =   "trPacking.frx":0700
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   6270
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
         Picture         =   "trPacking.frx":07A6
      End
   End
   Begin BiSAFramProject.BiSAFrame BisaFrame2 
      Height          =   2985
      Left            =   0
      Top             =   1125
      Width           =   8490
      _ExtentX        =   14975
      _ExtentY        =   5265
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
         Left            =   120
         TabIndex        =   15
         Top             =   2280
         Width           =   8280
         _ExtentX        =   14605
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
      Begin BiSANumberBoxProject.BiSANumberBox nQty 
         Height          =   330
         Index           =   0
         Left            =   4950
         TabIndex        =   8
         Top             =   600
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   582
         Appearance      =   0
         MinValue        =   0
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " "
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
         Index           =   0
         Left            =   1755
         TabIndex        =   9
         Top             =   600
         Width           =   3210
         _ExtentX        =   5662
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
      Begin BiSATextBoxProject.BiSABrowse cBarcode 
         Height          =   330
         Index           =   0
         Left            =   90
         TabIndex        =   10
         Top             =   600
         Width           =   1650
         _ExtentX        =   2910
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
      Begin BiSATextBoxProject.BiSABrowse cBarcode 
         Height          =   330
         Index           =   1
         Left            =   105
         TabIndex        =   12
         Top             =   1485
         Width           =   1665
         _ExtentX        =   2937
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
         Index           =   1
         Left            =   1770
         TabIndex        =   13
         Top             =   1485
         Width           =   3180
         _ExtentX        =   5609
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
      Begin BiSANumberBoxProject.BiSANumberBox nQty 
         Height          =   330
         Index           =   1
         Left            =   4935
         TabIndex        =   14
         Top             =   1485
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   582
         Appearance      =   0
         MinValue        =   0
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " "
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
      Begin BiSATextBoxProject.BiSABrowse cUnit 
         Height          =   330
         Index           =   0
         Left            =   5805
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   600
         Width           =   1080
         _ExtentX        =   1905
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
      Begin BiSATextBoxProject.BiSABrowse cUnit 
         Height          =   330
         Index           =   1
         Left            =   5790
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1485
         Width           =   1080
         _ExtentX        =   1905
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
      Begin BiSANumberBoxProject.BiSANumberBox nHarga 
         Height          =   330
         Index           =   0
         Left            =   6900
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   600
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   582
         Appearance      =   0
         MinValue        =   0
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
         Caption         =   " "
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
         Index           =   1
         Left            =   6885
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1485
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   582
         Appearance      =   0
         MinValue        =   0
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
         Caption         =   " "
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
         Caption         =   "Harga Beli"
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
         Left            =   6900
         TabIndex        =   22
         Top             =   135
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "From Stock /Inventory"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   195
         TabIndex        =   21
         Top             =   135
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Keterangan ..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   180
         TabIndex        =   16
         Top             =   1905
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "To Stock /Inventory"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   180
         TabIndex        =   11
         Top             =   1095
         Width           =   1935
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1125
      Left            =   0
      Top             =   15
      Width           =   8490
      _ExtentX        =   14975
      _ExtentY        =   1984
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
      Begin BiSADateProject.BiSADate dTgl 
         Height          =   330
         Left            =   75
         TabIndex        =   6
         Top             =   135
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   582
         Value           =   "24-12-2008"
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
         Left            =   75
         TabIndex        =   7
         Top             =   495
         Width           =   3405
         _ExtentX        =   6006
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
   End
End
Attribute VB_Name = "trPacking"
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
Dim cKode As String
Dim cJenis  As String
Dim nSaldoStock As Double

Private Sub GetFakturBrowse(ByVal lStat As Boolean)
  cFaktur.Button = lStat
End Sub


Private Sub cBarcode_ButtonClick(Index As Integer)
  Set dbData = objData.Browse(GetDSN, "stock", "kodestock,nama,kodesatuan,hargabeli,cogs", "kodestock", sisContent, cBarcode(Index).Text, " and jenis =1")
  If Not dbData.EOF Then
    cBarcode(Index).Text = cBarcode(Index).Browse(dbData)
    cNama(Index).Text = GetNull(dbData!nama)
    cUnit(Index).Text = GetNull(dbData!kodesatuan)
    nHarga(Index).Value = IIf(GetNull(dbData!cogs) <> 0, GetNull(dbData!cogs), GetNull(dbData!hargabeli))
  End If
End Sub

Private Sub cFaktur_ButtonClick()
Dim db As New ADODB.Recordset
Dim n As Integer
Dim lSave As Boolean

lSave = True
  
  Set db = objData.Browse(GetDSN, "totpacking", "nopacking,tgl,keterangan", "nopacking", sisContent, cFaktur.Text, " and tgl = '" & Format(dTgl.Value, "yyyy-MM-dd") & "'")
  If Not db.EOF Then
    cFaktur.Text = cFaktur.Browse(db)
    'ambil nilai total
    Set db = objData.Browse(GetDSN, "totpacking", , "nopacking", sisAssign, cFaktur.Text)
    If Not db.EOF Then
      dTgl.Value = GetNull(db!tgl)
      cKeterangan.Text = GetNull(db!keterangan)
    End If
    
    'ambil nilai detail
    Set db = objData.Browse(GetDSN, "packing p", "p.nopacking,p.kodestock,p.status,s.nama,p.jumlah,s.kodestock,p.jumlah,s.kodesatuan", "p.nopacking", sisAssign, cFaktur.Text, , , Array("Left join stock s on s.kodestock = p.kodestock"))
    If Not db.EOF Then
      Do While Not db.EOF
        If GetNull(db!Status) = "-1" Then
          cBarcode(0).Text = GetNull(db!KodeStock)
          cNama(0).Text = GetNull(db!nama)
          nQty(0).Value = GetNull(db!jumlah)
          cUnit(0).Text = GetNull(db!kodesatuan)
        End If
        If GetNull(db!Status) = "1" Then
          cBarcode(1).Text = GetNull(db!KodeStock)
          cNama(1).Text = GetNull(db!nama)
          nQty(1).Value = GetNull(db!jumlah)
          cUnit(1).Text = GetNull(db!kodesatuan)
        End If
        db.MoveNext
      Loop
    End If
    
    If nPos = Delete Then
      If MsgBox("Data akan dihapus?", vbYesNo) = vbYes Then
        objData.Start GetDSN
        lSave = IIf(lSave, objData.Delete(GetDSN, "totpacking", "nopacking", sisAssign, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "packing", "nopacking", sisAssign, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "kartustock", "nomor", sisAssign, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "bukubesar", "faktur", sisAssign, cFaktur.Text), False)
        If lSave Then
          objData.Save GetDSN
        Else
          objData.Cancel GetDSN
        End If
      End If
      GetEdit False
      initvalue
    End If
  End If
End Sub

Private Sub cmdAktivasi_Click()
  frmAktivasi.Action Me
End Sub

Private Sub GetDataStock()
End Sub

Private Sub cmdAdd_Click()
  nPos = Add
  GetEdit True
  GetFakturBrowse False
  cFaktur.Text = GetNomor("totpacking", "nopacking", GetID, sisModulTransaksi.Packing)
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

Private Function validOK() As Boolean
  validOK = True
End Function

Private Function ValidSaving() As Boolean
  ValidSaving = True
End Function

Private Sub cmdSimpan_Click()
Dim lSave As Boolean
Dim Faktur As String
Dim n As Integer
Dim i As Integer
Dim nValueTunai As Double
Dim nValueKredit As Double

lSave = True
  
  If isValidSaving Then
    objData.Start GetDSN
    Faktur = cFaktur.Text
    If nPos = Add Then
      If Not GetAvailable(cFaktur.Text, "totpacking", "nopacking") Then
        Faktur = GetNomor("totpacking", "nopacking", GetID, sisModulTransaksi.Packing)
      End If
    End If
    lSave = IIf(lSave, objData.Update(GetDSN, "totpacking", "nopacking = '" & Faktur & "'", Array("nopacking", "tgl", "username", "keterangan"), Array(Faktur, Format(dTgl.Value, "yyyy-MM-dd"), GetRegistry(reg_Username), cKeterangan.Text)), False)
    lSave = IIf(lSave, objData.Delete(GetDSN, "packing", "nopacking", sisAssign, Faktur), False)
    lSave = IIf(lSave, objData.Delete(GetDSN, "kartustock", "nomor", sisAssign, Faktur), False)
    lSave = IIf(lSave, objData.Delete(GetDSN, "bukubesar", "faktur", sisAssign, Faktur), False)

    lSave = IIf(lSave, objData.Add(GetDSN, "packing", Array("nopacking", "kodestock", "jumlah", "status"), Array(Faktur, cBarcode(0).Text, nQty(0).Value, "-1")), False)
    lSave = IIf(lSave, UpdKartuStock(objData, packingBahan, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), cBarcode(0).Text, nQty(0).Value, GetHargaPokok(objData, cBarcode(0).Text), 0, "Unpack Stock To " & cNama(1).Text & " " & cKeterangan.Text, aCfg(objData, msGudangPembelian), GetHargaPokok(objData, cBarcode(0).Text)), False)
    
    
    lSave = IIf(lSave, objData.Add(GetDSN, "packing", Array("nopacking", "kodestock", "jumlah", "status"), Array(Faktur, cBarcode(1).Text, nQty(1).Value, "1")), False)
    lSave = IIf(lSave, UpdKartuStock(objData, packingHasil, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), cBarcode(1).Text, nQty(1).Value, GetHargaPokok(objData, cBarcode(0).Text) / nQty(1).Value, 0, "Build Stock From " & cNama(0).Text & " " & cKeterangan.Text, aCfg(objData, msGudangPembelian), GetHargaPokok(objData, cBarcode(0).Text) / nQty(1).Value), False)
        
    'update harga pokok baru untuk produk hasil paking
    lSave = IIf(lSave, NewUpdHargaPokok(objData, cBarcode(1).Text), False)
    
    'posting nilai persediaan
    'persediaan from = berkurang
    'persediaan to = bertambah
    
    lSave = IIf(lSave, DelKodeTr(objData, vbTrigger.msPacking, Faktur), False)
    lSave = IIf(lSave, UpdKodeTr(objData, msPacking, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), GetAkunInventory(objData, cBarcode(0).Text), GetCostCenterUser(objData, GetRegistry(reg_Username)), "Unpack Stock From " & cNama(0).Text, 0, GetHargaPokok(objData, cBarcode(0).Text) * nQty(0).Value, "", SNow), False)
    lSave = IIf(lSave, UpdKodeTr(objData, msPacking, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), GetAkunInventory(objData, cBarcode(1).Text), GetCostCenterUser(objData, GetRegistry(reg_Username)), "Build Stock To " & cNama(1).Text, GetHargaPokok(objData, cBarcode(0).Text) * nQty(0).Value, 0, "", SNow), False)

    
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
Dim dba As New ADODB.Recordset
isValidSaving = True
  
End Function

Private Sub cNama_ButtonClick(Index As Integer)
  Set dbData = objData.Browse(GetDSN, "stock", "nama,kodestock,kodesatuan,hargabeli,cogs", "nama", sisContent, cNama(Index).Text, " and jenis = 1", "nama")
  If Not dbData.EOF Then
    cNama(Index).Text = cNama(Index).Browse(dbData)
    cBarcode(Index).Text = GetNull(dbData!KodeStock)
    cUnit(Index).Text = GetNull(dbData!kodesatuan)
    nHarga(Index).Value = IIf(GetNull(dbData!cogs) <> 0, GetNull(dbData!cogs), GetNull(dbData!hargabeli))
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

'  If CheckTrial(nRecordsTrial, TrialPenjualan) = True Then
'    End
'  End If
  
  
  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  initvalue
  GetEdit False
  
  
  TabIndex dTgl, n
  TabIndex cFaktur, n
  TabIndex cBarcode(0), n
  TabIndex cNama(0), n
  TabIndex nQty(0), n
  TabIndex cBarcode(1), n
  TabIndex cNama(1), n
  TabIndex nQty(1), n
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
  cBarcode(0).Default
  cBarcode(1).Default
  cNama(0).Default
  cNama(1).Default
  cKeterangan.Default
  cUnit(0).Default
  cUnit(1).Default
  nQty(0).Default
  nQty(1).Default
  nHarga(0).Default
  nHarga(1).Default
  trPacking.Caption = "Packing stok di Gudang " & aCfg(objData, msGudangPembelian)
End Sub

Private Sub GetEdit(lPar As Boolean)
  BiSAFrame1.Enabled = lPar
  BiSAFrame2.Enabled = lPar
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
