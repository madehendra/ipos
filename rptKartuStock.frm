VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptKartuStock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "KARTU STOCK"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   7380
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   3435
      Left            =   0
      Top             =   0
      Width           =   7380
      _ExtentX        =   13018
      _ExtentY        =   6059
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
      Begin BiSANumberBoxProject.BiSANumberBox nHK 
         Height          =   330
         Left            =   510
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1260
         Width           =   3195
         _ExtentX        =   5636
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
         BackColor       =   -2147483633
         Caption         =   "Hrg JL"
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
      Begin VB.CheckBox chkGudang 
         Caption         =   "Seluruh Gudang"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1620
         TabIndex        =   7
         Top             =   2610
         Width           =   1485
      End
      Begin BiSADateProject.BiSADate dDate 
         Height          =   330
         Index           =   0
         Left            =   510
         TabIndex        =   0
         Top             =   2055
         Width           =   2430
         _ExtentX        =   4286
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
         Caption         =   "Mutasi"
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
         Left            =   510
         TabIndex        =   1
         Top             =   915
         Width           =   5385
         _ExtentX        =   9499
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
      Begin BiSATextBoxProject.BiSABrowse cKode 
         Height          =   330
         Left            =   510
         TabIndex        =   2
         Top             =   570
         Width           =   3555
         _ExtentX        =   6271
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
      Begin BiSADateProject.BiSADate dDate 
         Height          =   330
         Index           =   1
         Left            =   3015
         TabIndex        =   3
         Top             =   2055
         Width           =   1740
         _ExtentX        =   3069
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
      Begin BiSATextBoxProject.BiSABrowse cGudang 
         Height          =   330
         Left            =   510
         TabIndex        =   6
         Top             =   2865
         Width           =   2430
         _ExtentX        =   4286
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
         Caption         =   "Gudang"
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
         Left            =   510
         TabIndex        =   8
         Top             =   225
         Width           =   4290
         _ExtentX        =   7567
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
         GetPicture      =   1
         Button          =   -1  'True
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
      Begin BiSANumberBoxProject.BiSANumberBox nStok 
         Height          =   330
         Left            =   510
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1605
         Width           =   2475
         _ExtentX        =   4366
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
         BackColor       =   -2147483633
         Caption         =   "Stok"
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
         Caption         =   "Seluruh Gudang"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3090
         TabIndex        =   11
         Top             =   1665
         Width           =   1200
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   0
      Top             =   3420
      Width           =   7380
      _ExtentX        =   13018
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
         Left            =   6210
         TabIndex        =   4
         Top             =   135
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
         Picture         =   "rptKartuStock.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   5775
         TabIndex        =   5
         Top             =   135
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
         Picture         =   "rptKartuStock.frx":00A6
      End
   End
End
Attribute VB_Name = "rptKartuStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB


Private Sub cBarcode_ButtonClick()
Dim cSQL As String
  'Set dbData = objData.Browse(GetDSN, "stock", "barcode,kodestock,nama,kodesatuan,hargajual", "barcode", sisContent, cBarcode.Text, " AND jenis < 9", "kodestock")
  'Set dbData = objData.Browse(GetDSN, "stock", "barcode,kodestock,nama,kodesatuan,hargajual", "barcode", sisContent, cBarcode.Text, "kodestock")
  
  cSQL = "select barcode,nama,hargajual,kodesatuan,kodestock,stok from stock where barcode like '%" & cBarcode.Text & "%' or nama like '%" & cBarcode.Text & "%'"
  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.EOF Then
'    cNama.Text = cNama.Browse(dbData, Array("BARCODE", "NAMA", "JUAL", "SATUAN"), , Array(13, 35, 10, 8))

    cBarcode.Text = cBarcode.Browse(dbData, Array("BARCODE", "NAMA", "JUAL"), , Array(13, 35, 10))
    cKode.Text = GetNull(dbData!KodeStock)
    cNama.Text = GetNull(dbData!nama)
    nHK.value = GetNull(dbData!HargaJual)
    nStok.value = GetNull(dbData!stok)
  End If
End Sub



Private Sub cBarcode_LostFocus()
'  Dim cSQL As String
'    'Set dbData = objData.Browse(GetDSN, "stock", "barcode,kodestock,nama,kodesatuan,hargajual", "barcode", sisContent, cBarcode.Text, " AND jenis < 9", "kodestock")
'    'Set dbData = objData.Browse(GetDSN, "stock", "barcode,kodestock,nama,kodesatuan,hargajual", "barcode", sisContent, cBarcode.Text, "kodestock")
'
'    cSQL = "select barcode,nama,hargajual,kodesatuan,kodestock,stok from stock where barcode = '" & cBarcode.Text & "'"
'    Set dbData = objData.SQL(GetDSN, cSQL)
'    If Not dbData.EOF Then
'      cKode.Text = GetNull(dbData!KodeStock)
'      cNama.Text = GetNull(dbData!nama)
'      nHK.Value = GetNull(dbData!HargaJual)
'      nStok.Value = GetNull(dbData!stok)
'    Else
'      cBarcode_ButtonClick
'    End If
End Sub

Private Sub cGudang_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "gudang", "kodegudang,keterangan")
  If Not dbData.EOF Then
    cGudang.Text = cGudang.Browse(dbData)
  End If
End Sub

Private Sub chkGudang_Click()
  If chkGudang.value = 1 Then
    cGudang.Enabled = False
  Else
    cGudang.Enabled = True
  End If
'  MsgBox chkGudang
End Sub

Private Sub chkGudang_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    keybd_event VK_TAB, 0, 0, 0
  End If
End Sub

Private Sub cKode_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "stock", "kodestock,barcode,nama,kodesatuan,hargajual", "kodestock", sisContent, cKode.Text, " AND jenis < 9", "kodestock")
  If Not dbData.EOF Then
    cKode.Text = cKode.Browse(dbData)
    cKode.Text = GetNull(dbData!KodeStock)
    cNama.Text = GetNull(dbData!nama)
    cBarcode.Text = GetNull(dbData!barcode)
    nHK.value = GetNull(dbData!HargaJual)
  End If
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
Dim n As Double
Dim nSaldo As Double
Dim cField  As String
Dim cWhere As String
Dim cSQLGudang As String


  cSQLGudang = ""
  If chkGudang.value <> 1 Then
    cSQLGudang = " AND kodegudang = '" & cGudang.Text & "'"
  End If
   
  cField = "tgl,keterangan,nomor,debet,kredit,Sum(debet-kredit) as SaldoAwal,kodestock"
  cWhere = cWhere & cSQLGudang
  cWhere = cWhere & " AND tgl < '" & Format(dDate(0).value, "yyyy-MM-dd") & "' GROUP BY kodestock"
  cWhere = cWhere & " AND status<> '" & SisKartuStock.refund & "'"
  Set dbData = objData.Browse(GetDSN, "kartustock", cField, "kodestock", sisAssign, cKode.Text, cWhere, "kodestock,tgl,id")
  If dbData.RecordCount > 0 Then
    nSaldo = GetNull(dbData!SaldoAwal)
  End If

  cWhere = ""
  cField = "tgl,keterangan,nomor,debet,kredit"
  cWhere = cWhere & cSQLGudang
  cWhere = cWhere & " AND tgl >= '" & Format(dDate(0).value, "yyyy-MM-dd") & "' AND tgl <= '" & Format(dDate(1).value, "yyyy-MM-dd") & "'"
  cWhere = cWhere & " AND status <> '" & SisKartuStock.refund & "'"
  Set dbData = objData.Browse(GetDSN, "kartustock", cField, "kodestock", sisAssign, cKode.Text, cWhere, "kodestock,tgl,id")
               
  vaArray.ReDim 0, 0, 0, 5
  vaArray(0, 2) = "Saldo Awal"
  vaArray(0, 3) = 0
  vaArray(0, 4) = 0
  vaArray(0, 5) = nSaldo
  
  If dbData.RecordCount > 0 Then
    dbData.MoveFirst
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = Format(dbData!tgl, "dd-MM-yyyy")
      vaArray(n, 1) = GetNull(dbData!keterangan)
      vaArray(n, 2) = GetNull(dbData!nomor)
      vaArray(n, 3) = GetNull(dbData!debet)
      vaArray(n, 4) = GetNull(dbData!kredit)
      dbData.MoveNext
    Loop
    For n = 0 To vaArray.UpperBound(1)
      nSaldo = nSaldo + vaArray(n, 3) - vaArray(n, 4)
      vaArray(n, 5) = nSaldo
    Next
  End If

  With FrmRPT
   .AddPageHeader "Kartu Stock" & IIf(chkGudang.value = 1, "", " Gudang " & cGudang.Text), tdbHalignCenter, , , , , 12, True, True
   
   .AddPageHeader "Kode", , , 15, , , , , , True, False
   .AddPageHeader ": [" & cKode.Text & "]- " & cBarcode.Text & " - " & cNama.Text
   .AddPageHeader "HK", , , 15, , , , , , True, False
   .AddPageHeader ": [" & Format(nHK.value, "###,###,###,##0.00") & "]"
   
   
   .AddPageHeader "Antara Tanggal", , , 15, True
   .AddPageHeader ": " & Format(dDate(0).value, "dd-MM-yyyy") & " s.d " & Format(dDate(1).value, "dd-MM-yyyy"), , , , , , , , , , , , , , , , 5
   
   .AddTableHeader "Tanggal", , , , 9
   .AddTableHeader "Keterangan"
   
   .AddTableHeader "Nomor", , , , 17
   .AddTableHeader "Debet", , , , 8
   .AddTableHeader "Kredit", , , , 8
   .AddTableHeader "Saldo", , , , 12
   
   .AddTableBody
   .AddTableBody
   
   .AddTableBody
   .AddTableBody Sis_Rpt_Number2
   .AddTableBody Sis_Rpt_Number2
   .AddTableBody Sis_Rpt_Number2
   
   .AddTableFooter "Total", , tdbHalignRight, , , , , , , , , , , , 3
   .AddTableFooter
   
   .AddTableFooter
   .AddTableFooter "&Sum", Sis_Rpt_Number2
   .AddTableFooter "&Sum", Sis_Rpt_Number2
   .AddTableFooter
   
   .Preview vaArray
  End With
End Sub

Private Sub cNama_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "stock", "kodestock,barcode,nama,kodesatuan,hargajual", "nama", sisContent, cNama.Text, " AND jenis < 9", "kodestock")
  If Not dbData.EOF Then
    cNama.Text = cNama.Browse(dbData)
    cKode.Text = GetNull(dbData!KodeStock)
    cNama.Text = GetNull(dbData!nama)
    cBarcode.Text = GetNull(dbData!barcode)
    nHK.value = GetNull(dbData!HargaJual)
  End If
End Sub

Private Sub Form_Load()
Dim n As Single
  
  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  dDate(0).value = BOM(Date)
  dDate(1).value = Date
  cGudang.Default
  chkGudang.value = 1
  cGudang.Enabled = False
  cBarcode.Default
  nHK.Default
  
  TabIndex cBarcode, n
  TabIndex cKode, n
  TabIndex cNama, n
  
  TabIndex dDate(0), n
  TabIndex dDate(1), n
  TabIndex chkGudang, n
  TabIndex cGudang, n
  TabIndex cmdPreview, n
  If GetRegistry(reg_UserLevel) <> 0 Then
    chkGudang.value = 0
    cGudang.Text = GetGudangUser(objData, GetRegistry(reg_Username))
  End If
End Sub
