VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Begin VB.Form trPromoTools 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PROMO TOOLS"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   3975
   Begin BiSAButtonProject.BiSAButton cmdRekapRebutan 
      Height          =   480
      Left            =   315
      TabIndex        =   0
      Top             =   195
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   847
      Caption         =   "Rekapitulasi Rebutan"
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
Attribute VB_Name = "trPromoTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub cmdRekapRebutan_Click()
Dim cSQL As String
Dim n As Single
Dim a As New exportExcel
Dim nTemp As Double

  'ambil seluruh member yg order
  cSQL = ""
'  cSQL = "select DISTINCT(t.kodeanggota) as kodeanggota,a.nama from memberorder m"
'  cSQL = cSQL & " left join totmemberorder t on t.nomormemberorder = m.nomormemberorder"
'  cSQL = cSQL & " left join anggota a on a.kodeanggota = t.kodeanggota"
'  cSQL = cSQL & " where m.tgl > '2013-6-1'"
    
cSQL = cSQL & " select DISTINCT(p.kodeanggota),a.nama from promo p"
cSQL = cSQL & " left join anggota a on a.kodeanggota = p.kodeanggota"
    
  vaArray.ReDim 0, -1, 0, 5
  nTemp = 0
  Set dbData = objData.Sql(GetDSN, cSQL)
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.EOF
      FrmPB.RunPB
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!kodeanggota)
      vaArray(n, 1) = GetNull(dbData!nama)
      vaArray(n, 2) = GetTotalOrderanPromo(objData, vaArray(n, 0))
      vaArray(n, 3) = GetBarangHabisBagi(objData, vaArray(n, 0))
      vaArray(n, 4) = GetMemberTopUp(vaArray(n, 0))
      vaArray(n, 5) = GetKeaktifan(objData, vaArray(n, 0))
      dbData.MoveNext
    Loop
    FrmPB.EndPB
    vaArray.QuickSort 0, vaArray.UpperBound(1), 5, XORDER_DESCEND, XTYPE_DOUBLE, 4, XORDER_DESCEND, XTYPE_DOUBLE
    a.RecordSource = vaArray
    a.ExportToExcel
  End If
End Sub

Private Function GetTotalOrderanPromo(ByVal obj As CodeSuiteLibrary.Data, ByVal Member As String) As Double
Dim db As New ADODB.Recordset
Dim cSQL As String


  GetTotalOrderanPromo = 0
  
  cSQL = ""
  cSQL = "select sum(m.qty*s.hargabeli) as total from memberorder m"
  cSQL = cSQL & " left join totmemberorder t on t.nomormemberorder = m.nomormemberorder"
  cSQL = cSQL & " left join stock s on s.kodestock = m.kodestock"
  cSQL = cSQL & " left join anggota a on a.kodeanggota = t.kodeanggota"
  cSQL = cSQL & " where a.kodeanggota = '" & Member & "' and m.tgl >= '2013-6-1'"

  Set db = obj.Sql(GetDSN, cSQL)
  If Not db.EOF Then
    GetTotalOrderanPromo = GetNull(db!Total)
  End If
End Function

Private Function GetBarangHabisBagi(ByVal obj As CodeSuiteLibrary.Data, ByVal Member As String) As Double
Dim db As New ADODB.Recordset
Dim cSQL As String


  GetBarangHabisBagi = 0
  
  cSQL = ""
  cSQL = cSQL & " select sum(p.qty*s.hargabeli) as total from promo p"
  cSQL = cSQL & " left join stock s on s.barcode = p.barcode"
  cSQL = cSQL & " where p.tgl > '2013-6-1' and p.kodeanggota = '" & Member & "'"

  Set db = obj.Sql(GetDSN, cSQL)
  If Not db.EOF Then
    GetBarangHabisBagi = GetNull(db!Total)
  End If
  
End Function


Private Function GetKeaktifan(ByVal obj As CodeSuiteLibrary.Data, ByVal Member As String) As Double
Dim cSQL As String
Dim db As New ADODB.Recordset
  
cSQL = cSQL & " SELECT"
cSQL = cSQL & "   a.telp AS nohp,"
  cSQL = cSQL & " a.nama,"
  cSQL = cSQL & " sum(p.qty * p.harga) AS total"
cSQL = cSQL & " From"
  cSQL = cSQL & " penjualan p"
cSQL = cSQL & " LEFT JOIN totpenjualan t ON t.nomorpenjualan = p.nomorpenjualan"
cSQL = cSQL & " LEFT JOIN anggota a ON a.kodeanggota = t.kodeanggota"
cSQL = cSQL & " Where"
  cSQL = cSQL & " p.discount >= 30"
cSQL = cSQL & " AND a.telp <> ''"
cSQL = cSQL & " AND p.tgl >= '2013-1-1'"
cSQL = cSQL & " AND t.kodeanggota = '" & Member & "'"
cSQL = cSQL & " GROUP BY t.kodeanggota"
'cSQL = cSQL & " Group By"
'  cSQL = cSQL & " a.nama"
  
  Set db = obj.Sql(GetDSN, cSQL)
  If Not db.EOF Then
    GetKeaktifan = GetNull(db!Total)
  End If
  
End Function

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  
End Sub
