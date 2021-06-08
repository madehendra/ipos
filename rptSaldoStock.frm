VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptSaldoStock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SALDO STOCK"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   6660
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1995
      Left            =   0
      Top             =   0
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   3519
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
      Begin BiSANumberBoxProject.BiSANumberBox nQty 
         Height          =   330
         Index           =   0
         Left            =   405
         TabIndex        =   0
         Top             =   1095
         Width           =   2445
         _ExtentX        =   4313
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
         Caption         =   "Antara"
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
         Left            =   420
         TabIndex        =   1
         Top             =   360
         Width           =   2430
         _ExtentX        =   4286
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
         Caption         =   "Sampai"
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
      Begin BiSATextBoxProject.BiSABrowse cGolongan 
         Height          =   330
         Index           =   0
         Left            =   405
         TabIndex        =   2
         Top             =   735
         Width           =   3105
         _ExtentX        =   5477
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
         Caption         =   "Golongan"
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
      Begin BiSATextBoxProject.BiSABrowse cGolongan 
         Height          =   330
         Index           =   1
         Left            =   3510
         TabIndex        =   3
         Top             =   735
         Width           =   2685
         _ExtentX        =   4736
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
         Caption         =   "s.d"
         CaptionWidth    =   0
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
         Left            =   2880
         TabIndex        =   4
         Top             =   1095
         Width           =   1890
         _ExtentX        =   3334
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
         Caption         =   "s.d"
         CaptionWidth    =   0
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
      Top             =   1980
      Width           =   6645
      _ExtentX        =   11721
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
         Left            =   5505
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
         Picture         =   "rptSaldoStock.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   5070
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
         Picture         =   "rptSaldoStock.frx":00A6
      End
      Begin MSComctlLib.ProgressBar pr 
         Height          =   375
         Left            =   90
         TabIndex        =   7
         Top             =   150
         Visible         =   0   'False
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         Min             =   1e-4
         Scrolling       =   1
      End
   End
End
Attribute VB_Name = "rptSaldoStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB
Dim vaGudang As New XArrayDB

Private Sub cGolongan_ButtonClick(Index As Integer)
  Set dbData = objData.PICK(GetDSN, "golongan", "kodegolongan", cGolongan(Index), "kodegolongan,keterangan")
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
  GetData
End Sub

Private Sub GetData()
Dim cField As String
Dim n As Double
Dim cWhere As String
Dim nCol As Double
Dim nRow As Double

  cField = "s.Golongan,g.Keterangan as NamaGolongan,s.Kode,s.Nama,s.Satuan,s.Min,s.Max"
  Set dbData = objData.Browse(GetDSN, "Gudang", "Kode", "Kode", sisGTEqual, cGudang(0).Text, " and Kode <= '" & cGudang(1).Text & "'", "Kode")
  If dbData.RecordCount > 0 Then
    vaGudang.LoadRows dbData.GetRows(dbData.RecordCount)
    dbData.MoveFirst
    Do While Not dbData.EOF
      cField = cField & ",0 as F" & Trim(dbData!Kode)
      dbData.MoveNext
    Loop
    cField = cField & ",0 as ssAkhir"
  
  
    Set dbData = objData.Browse(GetDSN, "Stock s", cField, "s.Golongan", sisGTEqual, cGolongan(0).Text, " and s.Golongan <= '" & cGolongan(1).Text & "'", "s.Golongan,s.Kode", _
                 Array("Left Join Golongan g on s.Golongan = g.Kode"))
    If dbData.RecordCount > 0 Then
      vaArray.LoadRows dbData.GetRows(dbData.RecordCount)
      
      ' Ambil Data pada table KartuStock
      cField = "Gudang,Kode,Sum(Debet-Kredit) as Mutasi"
      cWhere = "Gudang >= '" & cGudang(0).Text & "' and Gudang <= '" & cGudang(1).Text & "' "
      cWhere = cWhere & " and Tgl <= '" & Format(dTgl.Value, "yyyy-MM-dd") & "'"
      cWhere = cWhere & " Group by Gudang,Kode"
      Set dbData = objData.Browse(GetDSN, "KartuStock", cField, , , , cWhere, "Gudang,Kode")
      If dbData.RecordCount > 0 Then
        InitGauge pr, dbData.RecordCount
        Do While Not dbData.EOF
          RunGauge pr
          nRow = vaArray.Find(0, 2, GetNull(dbData!Kode))
          nCol = vaGudang.Find(0, 0, UCase(GetNull(dbData!Gudang)))
          SumArray nRow, nCol, GetNull(dbData!Mutasi), 0
          dbData.MoveNext
        Loop
        EndGauge pr
      End If
    Else
      MsgBox "Data tidak ada. Silahkan ulangi pengisian.", vbExclamation
      Exit Sub
    End If
    
    InitGauge pr, vaArray.UpperBound(1)
    n = 0
    Do While n <= vaArray.UpperBound(1)
      RunGauge pr
      If Not BetWeen(vaArray(n, vaArray.UpperBound(2)), nQty(0).Value, nQty(1).Value) Then
        vaArray.DeleteRows n
        n = n - 1
      End If
      n = n + 1
    Loop
    EndGauge pr
    
    With FrmRPT
      'Page Header
      .AddPageHeader "Laporan Saldo Stock", tdbHalignCenter, , , , , 12, True, True
      
      .AddPageHeader "Golongan", tdbHalignLeft, True, 15, True, , , True, , True, False
      .AddPageHeader ": " & cGolongan(0).Text & " s/d " & cGolongan(1).Text, , , , , , , True
      .AddPageHeader "Gudang", tdbHalignLeft, True, 15, True, , , True
      .AddPageHeader ": " & cGudang(0).Text & " s/d " & cGudang(1).Text, , , , , , , True
      .AddPageHeader "Antara Saldo", tdbHalignLeft, True, 15, True, , , True
      .AddPageHeader ": " & SisFormat(nQty(0).Value, sis_BilRpPict2) & " s/d " & SisFormat(nQty(1).Value, sis_BilRpPict2), , , , , , , True
      .AddPageHeader "Sampai Tanggal", tdbHalignLeft, True, 15, True, , , True
      .AddPageHeader ": " & Format(dTgl.Value, "dd-MM-yyyy"), , , , , , , True
      
      ' Tambah Group Header
      ' (Judul di atas Tabel jika laporan di buat Group)
      .AddTableGroupHeader True, "[]", , tdbHalignLeft, , 8
      .AddTableGroupHeader
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      
      ' Tambah Group Footer (Pembuatan Sub Total)
      .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupFooter "Sub Total", , tdbHalignRight, , , , , , , , , tdbTableFooterSect, , , 5
      .AddTableGroupFooter
      .AddTableGroupFooter
      .AddTableGroupFooter
      .AddTableGroupFooter
      .AddTableGroupFooter "&Sum", Sis_Rpt_Number2, tdbHalignRight
      
      ' Tambah Group Footer (Pembuatan Grand Total)
      .AddTableFooter , , , , , , , , , , , , , , , , , , , False
      .AddTableFooter , , , , , , , , , , , , , , , , , , , False
      .AddTableFooter "Grand Total", , tdbHalignRight, , , , , , , , , , , , 5
      .AddTableFooter
      .AddTableFooter
      .AddTableFooter
      .AddTableFooter
      .AddTableFooter "&Sum", Sis_Rpt_Number2
      
      ' Tambah Header Baris Pertama
      .AddTableHeader , , , , , , , , , , , , , , , , , , , False
      .AddTableHeader , , , , , , , , , , , , , , , , , , , False
      .AddTableHeader "Kode", , , , 7, , , , , , , , , tdbMergeOnText
      .AddTableHeader "Nama", , , , , , , , , , , , , tdbMergeOnText
      .AddTableHeader "Satuan", , , , 6, , , , , , , , , tdbMergeOnText
      .AddTableHeader "Min", , , , 4, , , , , , , , , tdbMergeOnText
      .AddTableHeader "Max", , , , 4, , , , , , , , , tdbMergeOnText
      .AddTableHeader "Gudang", , , , 7, , , , , , , , , , vaGudang.UpperBound(1) + 1
      
      For n = 1 To vaGudang.UpperBound(1)
        ' Tambah Table Group
        .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
        .AddTableGroupFooter "&Sum", Sis_Rpt_Number2, tdbHalignRight
        .AddTableFooter "&Sum", Sis_Rpt_Number2, tdbHalignRight

        .AddTableHeader "", , , , 7
      Next
      
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupFooter "&Sum", Sis_Rpt_Number2, tdbHalignRight
      .AddTableFooter "&Sum", Sis_Rpt_Number2, tdbHalignRight
      
      .AddTableHeader "Total Stock", , , , 7, , , , , , , , , tdbMergeOnText
      
      ' Tambah Header Baris 2
      .AddTableHeader , , , , , , , , , , , , , , , , , , , False
      .AddTableHeader , , , , , , , , , , , , , , , , , , , False
      .AddTableHeader "Kode", , , , 7, , , , , , True, , , tdbMergeOnText
      .AddTableHeader "Nama", , , , , , , , , , , , , tdbMergeOnText
      .AddTableHeader "Satuan", , , , 3, , , , , , , , , tdbMergeOnText
      .AddTableHeader "Min", , , , 3, , , , , , , , , tdbMergeOnText
      .AddTableHeader "Max", , , , 15, , , , , , , , , tdbMergeOnText
      
      .AddTableBody , , , , , , , , , , , , , False
      .AddTableBody , , , , , , , , , , , , , False
      .AddTableBody
      .AddTableBody
      .AddTableBody
      .AddTableBody Sis_Rpt_Number2
      .AddTableBody Sis_Rpt_Number2
      
      For n = 0 To vaGudang.UpperBound(1)
        .AddTableHeader vaGudang(n, 0), , , , 5
        
      .AddTableBody Sis_Rpt_Number2, tdbHalignRight
      Next
      .AddTableHeader "Total Stock", , , , 5, , , , , , , , , tdbMergeOnText
      .AddTableBody Sis_Rpt_Number2, tdbHalignRight
            
      .Preview vaArray, True, , True
    End With
  End If
End Sub

Private Sub SumArray(ByVal nRow As Double, ByVal nCol As Double, ByVal nDebet As Double, ByVal nKredit As Double)
Dim n As Double
Dim nSaldo As Double

  If nRow >= 0 And nCol >= 0 Then
    vaArray(nRow, nCol + 7) = vaArray(nRow, nCol + 7) + nDebet - nKredit
    For n = 7 To vaArray.UpperBound(2) - 1
      nSaldo = nSaldo + GetNull(vaArray(nRow, n))
    Next
    vaArray(nRow, vaArray.UpperBound(2)) = nSaldo
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  dTgl.Value = Date
  nQty(0).Value = -99999999
  nQty(1).Value = 99999999
  
  TabIndex dTgl, n
  TabIndex cGolongan(0), n
  TabIndex cGolongan(1), n
  TabIndex nQty(0), n
  TabIndex nQty(1), n
  TabIndex cmdPreview, n
End Sub
