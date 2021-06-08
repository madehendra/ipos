VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptSaldoSimpananPokok 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Saldo Simpanan Pokok"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7710
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   7710
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1170
      Left            =   0
      Top             =   0
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   2064
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
      Begin BiSANumberBoxProject.BiSANumberBox nSaldo 
         Height          =   330
         Index           =   0
         Left            =   135
         TabIndex        =   0
         Top             =   480
         Width           =   3735
         _ExtentX        =   6588
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
         Caption         =   "ANTARA SALDO"
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
      Begin BiSADateProject.BiSADate dDate 
         Height          =   330
         Left            =   135
         TabIndex        =   1
         Top             =   120
         Width           =   3165
         _ExtentX        =   5583
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
         Caption         =   "SAMPAI TANGGAL"
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
      Begin BiSANumberBoxProject.BiSANumberBox nSaldo 
         Height          =   330
         Index           =   1
         Left            =   4170
         TabIndex        =   2
         Top             =   495
         Width           =   2340
         _ExtentX        =   4128
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
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   630
      Left            =   0
      Top             =   1155
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
         Left            =   6435
         TabIndex        =   3
         Top             =   90
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
         Picture         =   "rptSaldoSimpananPokok.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   5265
         TabIndex        =   4
         Top             =   90
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
         Picture         =   "rptSaldoSimpananPokok.frx":00A6
      End
   End
End
Attribute VB_Name = "rptSaldoSimpananPokok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB
Dim lEmpty As Boolean

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
  GetSQL
End Sub

Private Sub GetSQL()
Dim cWhere As String
Dim n As Double
Dim dTanggalTutup
Dim cField As String
  
  lEmpty = False
  vaArray.Clear
  vaArray.ReDim 0, -1, 0, 4
  Set dbData = objData.Browse(GetDSN, "simpananpokok t", "t.kodeanggota,a.nama,a.kodedep,a.alamat, t.jumlah", "t.tgl", sisLTEqual, Format(dDate.Value, "yyyy-MM-dd"), , , Array("LEFT JOIN anggota a ON a.kodeanggota = t.kodeanggota"))
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!kodeanggota)
      vaArray(n, 1) = GetNull(dbData!nama)
      vaArray(n, 2) = GetNull(dbData!kodedep)
      vaArray(n, 3) = GetNull(dbData!alamat)
      vaArray(n, 4) = GetNull(dbData!Jumlah)
      dbData.MoveNext
    Loop
    Rpt
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd
  CenterForm Me
  nSaldo(0).Value = 0
  nSaldo(1).Value = 9999999999#
  dDate.Value = Date
      
  TabIndex dDate, n
  TabIndex nSaldo(0), n
  TabIndex nSaldo(1), n
  TabIndex cmdPreview, n
End Sub

Private Sub Rpt()
  With FrmRPT
    .AddPageHeader "DAFTAR SALDO SIMPANAN POKOK", tdbHalignCenter, , , , , 10, True
    .AddPageHeader UCase(aCfg(objData, msNamaPerusahaan)), tdbHalignCenter, , , True, dbArial, 12, True, False
    .AddPageHeader "Sampai dengan Tanggal : " & Format(dDate.Value, "dd MMMM yyyy"), tdbHalignCenter, , , True, , 9, False
    .AddPageHeader " ", , , , True
    .AddPageHeader " ", , , , True
    

    
    .AddTableHeader "No Anggota", , , , 10
    .AddTableHeader "Nama", , , , 25
    .AddTableHeader "Dep", , , , 6
    .AddTableHeader "Alamat"
    .AddTableHeader "Saldo Simpanan", , , , 15
    
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2
  
    
    .AddTableFooter "Grand Total", , tdbHalignRight, , , , , , , , , , , , 4
    .AddTableFooter ""
    .AddTableFooter ""
    .AddTableFooter ""
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    
    .Preview vaArray, True
  End With
End Sub




