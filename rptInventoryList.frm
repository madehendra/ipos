VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Begin VB.Form rptInventoryList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventory List"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   4065
   Begin BiSAButtonProject.BiSAButton cmdInventoryList 
      Height          =   450
      Left            =   450
      TabIndex        =   0
      Top             =   315
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   794
      Caption         =   "Print Inventory"
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
   Begin BiSAButtonProject.BiSAButton cmdNonInventoryList 
      Height          =   450
      Left            =   450
      TabIndex        =   1
      Top             =   795
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   794
      Caption         =   "Print Non Inventory"
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
   Begin BiSAButtonProject.BiSAButton cmdExportExcel 
      Height          =   450
      Left            =   450
      TabIndex        =   2
      Top             =   1290
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   794
      Caption         =   "Export To Excel"
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
End
Attribute VB_Name = "rptInventoryList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub cmdExportExcel_Click()
Dim db As New ADODB.Recordset
Dim a As New exportExcel
Dim na As Integer
Dim cSQL As String
Dim vaExport As New XArrayDB
Dim n As Single

  cSQL = "select s.kodestock,s.barcode,s.nama,g.keterangan,s.hargabeli,s.hargajual,s.kodesatuan from stock s"
  cSQL = cSQL & " LEFT JOIN golongan g on g.kodegolongan = s.kodegolongan"

  Set db = objData.SQL(GetDSN, cSQL)
  If Not db.EOF Then
    vaExport.ReDim 0, -1, 0, 7
    Do While Not db.EOF
      vaExport.InsertRows vaExport.UpperBound(1) + 1
      n = vaExport.UpperBound(1)
      vaExport(n, 0) = GetNull(db!KodeStock)
      vaExport(n, 1) = GetNull(db!barcode)
      vaExport(n, 2) = GetNull(db!nama)
      vaExport(n, 3) = GetNull(db!keterangan)
      vaExport(n, 4) = GetNull(db!hargabeli)
      vaExport(n, 5) = GetNull(db!hargajual)
      vaExport(n, 6) = GetNull(db!kodesatuan)
      vaExport(n, 7) = GetSaldoStock(objData, "", vaExport(n, 0))
      db.MoveNext
    Loop
    a.RecordSource = vaExport
    a.ExportToExcel
  Else
    MsgBox "Tidak ada data untuk di export", vbInformation
  End If
End Sub

Private Sub cmdInventoryList_Click()
  PreviewDong 1
End Sub

Private Sub cmdNonInventoryList_Click()
  PreviewDong 9
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hwnd
  CenterForm Me
  TabIndex cmdInventoryList, n
  TabIndex cmdNonInventoryList, n
End Sub

Private Sub PreviewDong(ByVal cStatus)
Dim n As Double
  
  vaArray.ReDim 0, -1, 0, 5
  Set dbData = objData.Browse(GetDSN, "stock s", "s.kodegolongan,g.keterangan as namagolongan,s.kodestock,s.nama as namastock,s.kodesatuan,s.hargajual", "s.jenis", sisAssign, cStatus, , "s.kodegolongan,s.kodestock", Array("left join golongan g on g.kodegolongan = s.kodegolongan"))
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!kodegolongan)
      vaArray(n, 1) = GetNull(dbData!namagolongan)
      vaArray(n, 2) = GetNull(dbData!KodeStock)
      vaArray(n, 3) = GetNull(dbData!namastock)
      vaArray(n, 4) = GetNull(dbData!kodesatuan)
      vaArray(n, 5) = GetNull(dbData!hargajual)
      dbData.MoveNext
    Loop
  End If
  
  With FrmRPT
    .AddPageHeader IIf(cStatus = 1, "Invetory List", "Non Inventory List"), tdbHalignCenter, , , True, dbArial, 12, True, , , False
    .AddPageHeader aCfg(objData, msNamaPerusahaan), tdbHalignCenter, , , True, , 14
    .AddPageHeader "", , , , True
    .AddPageHeader "", , , , True
    
    .AddTableGroupHeader True, "[]", , , , 10, , , , , , , , , , , , , , , True
    .AddTableGroupHeader
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    
    .AddTableHeader , , , , , , , , , , , , , , , , , , , False
    .AddTableHeader , , , , , , , , , , , , , , , , , , , False
    .AddTableHeader "KODE", , , , 7
    .AddTableHeader "NAMA"
    .AddTableHeader "SATUAN", , , , 15
    .AddTableHeader "HARGA JUAL", , , , 20
    
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    
    .Preview vaArray, True
  End With
  
End Sub
