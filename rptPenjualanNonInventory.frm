VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptPenjualanNonInventory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Penjualan Non Inventory"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   6660
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   2340
      Left            =   0
      Top             =   0
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   4128
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
      Begin VB.CheckBox chkStock 
         Caption         =   "All Item"
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
         Left            =   1365
         TabIndex        =   0
         Top             =   390
         Width           =   1485
      End
      Begin BiSADateProject.BiSADate dDate 
         Height          =   330
         Index           =   0
         Left            =   255
         TabIndex        =   1
         Top             =   1305
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
         Left            =   255
         TabIndex        =   2
         Top             =   945
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
      Begin BiSATextBoxProject.BiSABrowse cKode 
         Height          =   330
         Left            =   255
         TabIndex        =   3
         Top             =   600
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
      Begin BiSADateProject.BiSADate dDate 
         Height          =   330
         Index           =   1
         Left            =   2670
         TabIndex        =   4
         Top             =   1305
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
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   0
      Top             =   2340
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
         Picture         =   "rptPenjualanNonInventory.frx":0000
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
         Picture         =   "rptPenjualanNonInventory.frx":00A6
      End
   End
End
Attribute VB_Name = "rptPenjualanNonInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
  GetData
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  TabIndex chkStock, n
  TabIndex cKode, n
  TabIndex cNama, n
  TabIndex dDate(0), n
  TabIndex dDate(1), n
  chkStock.Value = 0
  cKode.Default
  cNama.Default
  dDate(0).Value = Date
  dDate(1).Value = Date
End Sub

Private Sub GetData()
Dim cSQL As String
Dim cWhere As String
Dim n As Integer


  vaArray.ReDim 0, -1, 0, 9
  
  If chkStock.Value = 0 Then
    cWhere = " AND p.kodestock = '" & cKode.Text & "'"
  End If

  cSQL = ""
  cSQL = "select p.nomorpenjualan,p.tgl,p.kodestock,s.barcode,s.nama,p.qty,p.kodesatuan,p.harga,p.discount,p.jumlah,s.kodestock from penjualan p"
  cSQL = cSQL & " left join stock s on s.kodestock = p.kodestock"
  cSQL = cSQL & " Where s.jenis = 9"
  cSQL = cSQL & cWhere
  cSQL = cSQL & " and p.tgl >= '" & Format(dDate(0).Value, "yyyy-MM-dd") & "' and p.tgl <= '" & Format(dDate(1).Value, "yyyy-MM-dd") & "'"
  cSQL = cSQL & " order by p.tgl asc"
  Set dbData = objData.Sql(GetDSN, cSQL)
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!nomorpenjualan)
      vaArray(n, 1) = GetNull(dbData!tgl)
      vaArray(n, 2) = GetNull(dbData!KodeStock)
      vaArray(n, 3) = GetNull(dbData!Barcode)
      vaArray(n, 4) = GetNull(dbData!nama)
      vaArray(n, 5) = GetNull(dbData!qty)
      vaArray(n, 6) = GetNull(dbData!kodesatuan)
      vaArray(n, 7) = GetNull(dbData!Harga)
      vaArray(n, 8) = GetNull(dbData!Discount)
      vaArray(n, 9) = GetNull(dbData!jumlah)
      dbData.MoveNext
    Loop
    GetPrint
  Else
    MsgBox "Maaf, tidak ada data untuk ditampilkan"
  End If
End Sub

Private Sub GetPrint()
  With FrmRPT
    .AddPageHeader "Penjualan Non Inventory", tdbHalignCenter, , , , , 12, True, True
    
    .AddPageHeader "Antara Tanggal", , , 15, True
    .AddPageHeader ": " & Format(dDate(0).Value, "dd-MM-yyyy") & " s.d " & Format(dDate(1).Value, "dd-MM-yyyy"), , , , , , , , , , , , , , , , 5
        
    
    .AddTableHeader "No.Faktur", , , , 12
    .AddTableHeader "Tgl", , , , 9
    .AddTableHeader "Kode", , , , 7
    .AddTableHeader "Barcode", , , , 14
    .AddTableHeader "Nama"
    .AddTableHeader "Qty", , , , 7
    .AddTableHeader "Unit", , , , 5
    .AddTableHeader "Harga", , , , 10
    .AddTableHeader "Dsc", , , , 5
    .AddTableHeader "Jumlah", , , , 12
    
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody Sis_Rpt_Number2
    
    .AddTableFooter "Total", , tdbHalignRight, , , , , , , , , , , , 9
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    
    .Preview vaArray
  End With
End Sub
