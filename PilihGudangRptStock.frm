VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form PilihGudangRptStock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pilih Gudang..."
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   4065
   Begin BiSADateProject.BiSADate BiSADate1 
      Height          =   405
      Left            =   1800
      TabIndex        =   5
      Top             =   2370
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   714
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
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
   Begin BiSAButtonProject.BiSAButton BiSAButton2 
      Height          =   540
      Left            =   120
      TabIndex        =   4
      Top             =   2790
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   953
      Caption         =   "Cek Saldo Stok as Tgl"
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
   Begin BiSAButtonProject.BiSAButton BiSAButton1 
      Height          =   420
      Left            =   195
      TabIndex        =   3
      Top             =   900
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   741
      Caption         =   "Print Seluruh Stock - Seluruh Gudang"
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
   Begin VB.CommandButton cmdOK 
      Caption         =   "OKEY!!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   195
      TabIndex        =   0
      Top             =   525
      Width           =   3525
   End
   Begin BiSATextBoxProject.BiSABrowse cKodePoli 
      Height          =   330
      Left            =   165
      TabIndex        =   1
      Top             =   75
      Width           =   3525
      _ExtentX        =   6218
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
      CaptionWidth    =   2000
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
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Perhatian: Daftar stock berikut akan menampilkan stock/inventory yang ada nilai atau jumlah stocknya"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   1410
      Width           =   3720
   End
End
Attribute VB_Name = "PilihGudangRptStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub BiSAButton1_Click()
  GetExportToExcel
End Sub

Private Sub GetExportToExcel()
Dim cSQL As String
Dim n As Integer


'  cSQL = ""
'  cSQL = cSQL & "select s.kodegolongan,s.nama,s.barcode,sum(k.debet-k.kredit)as sal  from kartustock k"
'  cSQL = cSQL & " left join stock s on s.kodestock = k.kodestock"
'  cSQL = cSQL & " group by k.kodestock"
'  cSQL = cSQL & " order by s.kodegolongan,s.barcode,sal"
  
  cSQL = ""
'  cSQL = cSQL & " select s.kodegolongan,s.nama,s.barcode,sum(k.debet-k.kredit)as sal,s.hargabeli  from kartustock k"
'  cSQL = cSQL & " left join stock s on s.kodestock = k.kodestock "
'  cSQL = cSQL & " where s.diskonpenjualan >=20 "
'  cSQL = cSQL & " group by k.kodestock"
'  cSQL = cSQL & " order by s.kodegolongan,s.barcode,sal"
  
cSQL = " select s.kodegolongan,s.nama,s.barcode,sum(k.debet-k.kredit)as sal,s.hargabeli"
cSQL = cSQL & " from kartustock k left join stock s on s.kodestock = k.kodestock"
cSQL = cSQL & " group by k.kodestock"
cSQL = cSQL & " Having sal <> 0"
cSQL = cSQL & " order by s.kodegolongan,s.barcode,sal"

  
  
  vaArray.ReDim 0, -1, 0, 5
  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.EOF Then
'    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.EOF
'      FrmPB.RunPB
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!kodegolongan)
      vaArray(n, 1) = GetNull(dbData!nama)
      vaArray(n, 2) = GetNull(dbData!barcode)
      vaArray(n, 3) = GetNull(dbData!sal)
      vaArray(n, 4) = GetNull(dbData!hargabeli)
      vaArray(n, 5) = 0
'      vaArray(n, 5) = GetHapusBarangLama(vaArray(n, 2))
'
''      If Format(GetHapusBarangLama(vaArray(n, 2)), "yyyy-MM-dd") < "2015-1-1" Then
''        vaArray.DeleteRows n
''      Else
''        MsgBox "ini baru barang"
''
''      End If
'
'      If vaArray(n, 3) = 0 Then
'        vaArray.DeleteRows n
'      End If

      dbData.MoveNext
    Loop
  End If
  
'  If Not dbData.EOF Then
'    vaArray
'  End If
  
  Dim a As New exportExcel
  a.RecordSource = vaArray
  a.ExportToExcel
  
End Sub

Private Function GetHapusBarangLama(cBarcode As String) As Date
Dim cSQL As String
Dim dba As New ADODB.Recordset
  
GetHapusBarangLama = "2015-1-1"
cSQL = " select tgl from kartustock k"
cSQL = cSQL & " LEFT JOIN stock s on s.kodestock = k.kodestock"
cSQL = cSQL & " Where s.barcode = '" & cBarcode & "'"
cSQL = cSQL & " ORDER BY tgl DESC LIMIT 0,1"

  Set dba = objData.SQL(GetDSN, cSQL)
  If Not dba.EOF Then
    GetHapusBarangLama = GetNull(dba!tgl)
  End If
  
End Function

Private Sub BiSAButton2_Click()
Dim cSQL As String
Dim n As Integer
Dim a As New exportExcel

cSQL = "select DISTINCT(s.kodestock) as kodestock,s.barcode,s.nama,s.hargajual from pembelian p"
cSQL = cSQL & " left join stock s on s.kodestock = p.kodestock"
cSQL = cSQL & " Where p.Tgl >= '" & Format(BiSADate1.value, "yyyy-mm-dd") & "'"

  vaArray.ReDim 0, -1, 0, 3
  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.EOF
      FrmPB.RunPB
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!barcode)
      vaArray(n, 1) = GetNull(dbData!nama)
      vaArray(n, 2) = GetNull(dbData!HargaJual)
      vaArray(n, 3) = GetSaldoStock(objData, "", GetNull(dbData!KodeStock), 0)
      If vaArray(n, 3) = 0 Then vaArray.DeleteRows n
      dbData.MoveNext
    Loop
    FrmPB.EndPB
    a.RecordSource = vaArray
    a.ExportToExcel
  End If
End Sub

Private Sub cKodePoli_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "gudang")
  If Not dbData.EOF Then
    cKodePoli.Text = cKodePoli.Browse(dbData)
  End If
End Sub

Private Sub cmdOK_Click()
  rptStock.Gudang = cKodePoli.Text
  Unload Me
  Load rptStock
  rptStock.Show
  rptStock.SetFocus
End Sub

Private Sub Form_Load()
Dim n As Single
  
  SetIcon Me.hWnd
  CenterForm Me
  cKodePoli.Text = GetGudangUser(objData, GetRegistry(reg_Username))
  TabIndex cKodePoli, n
  TabIndex cmdOK, n
End Sub
