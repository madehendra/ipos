VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptMemberTopUp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laporan Member Top Up"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   5460
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1245
      Left            =   0
      Top             =   0
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   2196
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
      Begin BiSAButtonProject.BiSAButton BiSAButton3 
         Height          =   345
         Left            =   4380
         TabIndex        =   6
         Top             =   825
         Visible         =   0   'False
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   609
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
      Begin BiSAButtonProject.BiSAButton cmdRebutan 
         Height          =   390
         Left            =   4185
         TabIndex        =   5
         Top             =   165
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   688
         Caption         =   "rebutan"
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
      Begin BiSAButtonProject.BiSAButton BiSAButton2 
         Height          =   360
         Left            =   525
         TabIndex        =   4
         Top             =   825
         Visible         =   0   'False
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   635
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
      Begin BiSAButtonProject.BiSAButton BiSAButton1 
         Height          =   330
         Left            =   2565
         TabIndex        =   3
         Top             =   840
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   582
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
      Begin BiSADateProject.BiSADate dDate 
         Height          =   330
         Left            =   885
         TabIndex        =   0
         Top             =   345
         Width           =   3225
         _ExtentX        =   5689
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
         Caption         =   "Sampai Tgl"
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
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   0
      Top             =   1230
      Width           =   5445
      _ExtentX        =   9604
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
         Left            =   4290
         TabIndex        =   1
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
         Picture         =   "rptMemberTopUp.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   3855
         TabIndex        =   2
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
         Picture         =   "rptMemberTopUp.frx":00A6
      End
   End
End
Attribute VB_Name = "rptMemberTopUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub BiSAButton1_Click()
Dim cSQL As String
Dim n As Single
Dim a As New exportExcel

  cSQL = "select t.kodeanggota,a.nama,t.total from totpenjualan t"
  cSQL = cSQL & " left join anggota a on a.kodeanggota = t.kodeanggota"
  cSQL = cSQL & " Where t.jenis = 'P' And t.tgl = '2012-01-22' and t.flaglunas = 0"
  
  vaArray.ReDim 0, -1, 0, 2
  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!nama)
      vaArray(n, 1) = GetNull(dbData!Total)
      vaArray(n, 2) = GetMemberTopUp(GetNull(dbData!kodeanggota))
      dbData.MoveNext
    Loop
    a.RecordSource = vaArray
    a.ExportToExcel
  End If
End Sub

Private Sub BiSAButton2_Click()
Dim cSQL As String
Dim n As Integer
Dim db As New ADODB.Recordset
Dim a As New exportExcel

cSQL = cSQL & " select a.nama,a.telp,a.kodeanggota,sum(subtotal) as jumlah from totmemberorder t"
cSQL = cSQL & " LEFT JOIN anggota a on a.kodeanggota = t.kodeanggota"
cSQL = cSQL & " Where t.Tgl >= '2012-06-01'"
cSQL = cSQL & " GROUP BY a.nama ORDER BY jumlah DESC"

vaArray.ReDim 0, -1, 0, 4
Set db = objData.SQL(GetDSN, cSQL)
If Not db.EOF Then
  Do While Not db.EOF
    vaArray.InsertRows vaArray.UpperBound(1) + 1
    n = vaArray.UpperBound(1)
    vaArray(n, 0) = GetNull(db!nama)
    vaArray(n, 1) = GetNull(db!telp)
    vaArray(n, 2) = GetNull(db!jumlah)
    vaArray(n, 3) = GetSaldoTopUpMember(objData, GetNull(db!kodeanggota))
    vaArray(n, 4) = GetBerapaSudahDapatPromo(objData, GetNull(db!kodeanggota))
    
'    If vaArray(n, 3) <= 0 And vaArray(n, 4) < 0 Then
'      vaArray.DeleteRows n
'    End If

    db.MoveNext
  Loop
  a.RecordSource = vaArray
  a.ExportToExcel
End If

End Sub

Private Function GetBerapaSudahDapatPromo(ByVal obj As CodeSuiteLibrary.Data, ByVal cAnggota As String) As Double
Dim cSQL As String
Dim db As New ADODB.Recordset

  GetBerapaSudahDapatPromo = 0
  cSQL = "select p.kodeanggota,a.nama,sum(p.qty*s.hargabeli) as total from promo p"
  cSQL = cSQL & " left join stock s on s.kodestock = p.kodestock"
  cSQL = cSQL & " left join anggota a on a.kodeanggota = p.kodeanggota"
  cSQL = cSQL & " Where p.Tgl >= '2012-06-01' And p.kodekatalog = 'JUN2012' and p.kodeanggota = '" & cAnggota & "'"
  cSQL = cSQL & " GROUP BY p.kodeanggota"
  
  Set db = objData.SQL(GetDSN, cSQL)
  If Not db.EOF Then
    GetBerapaSudahDapatPromo = GetNull(db!Total)
  End If

End Function

Private Sub BiSAButton3_Click()
Dim cSQL As String
Dim db As New ADODB.Recordset
Dim n As Single
Dim a As New exportExcel

cSQL = cSQL & " select a.nama,a.kodeanggota,a.telp,sum(p.qty*s.hargabeli) as total from promo p"
cSQL = cSQL & " left join stock s on s.barcode = p.barcode"
cSQL = cSQL & " left join  anggota a on a.kodeanggota = p.kodeanggota"
cSQL = cSQL & " Where p.kodekatalog = 'JUN2012' or p.tgl >='2012-06-01'"
cSQL = cSQL & " GROUP BY a.nama"
cSQL = cSQL & " ORDER BY sum(p.qty*s.hargabeli) desc;"

vaArray.ReDim 0, -1, 0, 5
Set db = objData.SQL(GetDSN, cSQL)
If Not db.EOF Then
  FrmPB.InitPB db.RecordCount
  Do While Not db.EOF
    FrmPB.RunPB
    vaArray.InsertRows vaArray.UpperBound(1) + 1
    n = vaArray.UpperBound(1)
    vaArray(n, 0) = GetNull(db!nama)
    vaArray(n, 1) = GetNull(db!telp)
    vaArray(n, 2) = GetNull(db!Total)
    vaArray(n, 3) = GetSaldoTopUpMember2(objData, GetNull(db!kodeanggota))
    vaArray(n, 4) = vaArray(n, 2) - vaArray(n, 3)
    vaArray(n, 5) = GetBarangPromoYgDpt(objData, "JUN2012", GetNull(db!kodeanggota))
    'cek apakah barang sudah diambil atau belum
    
    db.MoveNext
  Loop
  FrmPB.EndPB
  vaArray.QuickSort 0, vaArray.UpperBound(1), 4, XORDER_DESCEND, XTYPE_DOUBLE
  a.RecordSource = vaArray
  a.ExportToExcel
  
End If

End Sub

Private Function GetBarangPromoYgDpt(ByVal obj As CodeSuiteLibrary.Data, ByVal cKodeKatalog As String, ByVal cKodeMember As String) As String
Dim db As New ADODB.Recordset
Dim cSQL As String
Dim cTmp As String

cSQL = cSQL & " select barcode,qty from promo"
cSQL = cSQL & " Where kodekatalog = '" & cKodeKatalog & "' And kodeanggota = '" & cKodeMember & "'"

  Set db = obj.SQL(GetDSN, cSQL)
  If Not db.EOF Then
    Do While Not db.EOF
      cTmp = GetNull(db!barcode) & IIf(GetNull(db!qty) > 1, "(" & GetNull(db!qty) & ")", "")
      GetBarangPromoYgDpt = cTmp & "," & GetBarangPromoYgDpt
      db.MoveNext
    Loop
  End If
End Function

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Function isSudahDiambilBelum(ByVal obj As CodeSuiteLibrary.Data, ByVal cKodeMember As String) As Boolean
Dim db As New ADODB.Recordset
Dim cSQL As String

  cSQL = cSQL & " select * from totpenjualan"
  cSQL = cSQL & " Where jenis = 'P' And Tgl >= '2012-06-01' and kodeanggota = '" & cKodeMember & "'"
  
  Set db = obj.SQL(GetDSN, cSQL)
  If Not db.EOF Then
    
    
  End If

End Function

Private Function GetBerapaDiaOrderPromo(ByVal obj As CodeSuiteLibrary.Data, ByVal cAnggota As String) As Double
Dim cSQL As String
Dim db As New ADODB.Recordset
  
  GetBerapaDiaOrderPromo = 0
  cSQL = cSQL & " select a.nama,a.telp,a.kodeanggota,sum(subtotal) as jumlah from totmemberorder t"
  cSQL = cSQL & " LEFT JOIN anggota a on a.kodeanggota = t.kodeanggota"
  cSQL = cSQL & " Where t.Tgl >= '2012-06-01' and t.kodeanggota = '" & cAnggota & "'"
  cSQL = cSQL & " GROUP BY a.nama ORDER BY jumlah DESC"
  
  Set db = objData.SQL(GetDSN, cSQL)
  If Not db.EOF Then
    GetBerapaDiaOrderPromo = GetNull(db!jumlah)
  End If

End Function

Private Sub cmdPreview_Click()
Dim cFields As String
Dim cWhere As String
Dim n As Double
Dim nCol As Double

  
  vaArray.ReDim 0, -1, 0, 7
'  cFields = "s.kodeanggota,s.nama,s.alamat,Sum(h.debet) as Debet,Sum(h.kredit) as Kredit,s.kodedep,d.keterangan as namadep"
'  cWhere = cWhere & " h.tgl <= '" & Format(dDate(0).Value, "yyyy-MM-dd") & "'"
'  cWhere = cWhere & " GROUP BY s.kodeanggota"
'  Set dbData = objData.Browse(GetDSN, "anggota s", cFields, , , , cWhere, "s.kodedep,s.kodeanggota", _
'                              Array("LEFT JOIN kartupiutang h on h.kodeanggota = s.kodeanggota", _
'                              "Left join dep d on d.kodedep = s.kodedep"))
                              
  Set dbData = objData.Browse(GetDSN, "membertopup m", "m.kodeanggota,a.nama,a.alamat,sum(debet) as debet,sum(kredit) as kredit,sum(m.debet-m.kredit) as saldo", "m.tgl", sisLTEqual, Format(dDate.Value, "yyyy-MM-dd"), " GROUP BY m.kodeanggota", , Array("left join anggota a on a.kodeanggota = m.kodeanggota"))
  
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    dbData.MoveFirst
    Do While Not dbData.EOF
      FrmPB.RunPB
      vaArray.InsertRows (vaArray.UpperBound(1)) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!kodeanggota)
      vaArray(n, 1) = GetNull(dbData!nama)
      vaArray(n, 2) = GetNull(dbData!alamat)
      vaArray(n, 3) = GetNull(dbData!debet)
      vaArray(n, 4) = GetNull(dbData!kredit)
      vaArray(n, 5) = vaArray(n, 3) - vaArray(n, 4)
      vaArray(n, 6) = GetNull(GetMemberPiutang(vaArray(n, 0)))
      vaArray(n, 7) = vaArray(n, 5) - vaArray(n, 6)
      dbData.MoveNext
    Loop
    FrmPB.EndPB
  End If
  
  n = 0
  Do While n <= vaArray.UpperBound(1)
    If vaArray(n, 7) = 0 Then
      vaArray.DeleteRows n
      n = n - 1
    End If
    n = n + 1
  Loop
  
'  vaArray.QuickSort 0, vaArray.UpperBound(1), 3, XORDER_DESCEND, XTYPE_DOUBLE
'  vaArray.QuickSort 0, vaArray.UpperBound(6), 3, XORDER_DESCEND, XTYPE_DOUBLE
  vaArray.QuickSort 0, vaArray.UpperBound(1), 5, XORDER_DESCEND, XTYPE_DOUBLE
  GetRpt
  
  If MsgBox("Apakah laporan ini akan di export ke format Excel?", vbYesNo) = vbYes Then
    Dim a As New exportExcel
    vaArray.DeleteColumns (0)
    a.RecordSource = vaArray
    a.ExportToExcel
  End If
End Sub

Private Sub GetFullOutstanding(ByVal obj As CodeSuiteLibrary.Data, ByVal member As String, ByVal dTgl As Date, ByRef nPiutangReg As Double, ByRef nSaldoTopUp As Double)
'mencari nilai saldo piutang reguler, saldo topup dari serorang member
Dim db As New ADODB.Recordset

  nPiutangReg = 0
  nSaldoTopUp = 0
  Set db = obj.Browse(GetDSN, "anggota", "kodeanggota,nama", "kodeanggota", sisAssign, member)
  If Not db.EOF Then
    'mengambil data piutang member
    nPiutangReg = GetMemberPiutang(member)
    nSaldoTopUp = GetMemberTopUp(member)
  End If
End Sub


Private Sub GetRpt()
  With FrmRPT
      .AddPageHeader "SALDO MEMBER TOP UP", tdbHalignCenter, , , True, dbArial, 11, False, False, , False
      .AddPageHeader aCfg(objData, msNamaPerusahaan, ""), tdbHalignCenter, , , True, , 12, True
      .AddPageHeader "sda Tgl : " & Format(dDate.Value, "dd-MM-yyyy"), tdbHalignCenter, , , True, , , True
      .AddPageHeader " ", , , , True
'      .AddPageHeader " ", , , , True
'      .AddPageHeader " ", , , , True
'      .AddPageHeader " ", , , , True
'      .AddPageHeader " ", , , , True
'      .AddPageHeader " ", , , , True
      
      .AddTableHeader "Kode", , , , 11
      .AddTableHeader "Nama"
      .AddTableHeader "Alamat", , , , , , , , , , , , , , , , , , , False
      .AddTableHeader "Debet", , , , 12, , , , , , , , , , , , , , , False
      .AddTableHeader "Kredit", , , , 12, , , , , , , , , , , , , , , False
      .AddTableHeader "Top Up", , , , 12
      .AddTableHeader "Outstanding", , , , 12
      .AddTableHeader "O-Top Up", , , , 12

      .AddTableBody
      .AddTableBody
      .AddTableBody , , , , , , , , , , , , , False
      .AddTableBody Sis_Rpt_Number2, , , , , , , , , , , , , False
      .AddTableBody Sis_Rpt_Number2, , , , , , , , , , , , , False
      .AddTableBody Sis_Rpt_Number2
      .AddTableBody Sis_Rpt_Number2
      .AddTableBody Sis_Rpt_Number2
           
      .AddTableFooter "TOTAL", , tdbHalignRight, , , , , , , , , , , , 2
      .AddTableFooter
      .AddTableFooter , , , , , , , , , , , , , , , , , , , False
      .AddTableFooter "&Sum", Sis_Rpt_Number2, , , , , , , , , , , , , , , , , , False
      .AddTableFooter "&Sum", Sis_Rpt_Number2, , , , , , , , , , , , , , , , , , False
      .AddTableFooter "&Sum", Sis_Rpt_Number2
      .AddTableFooter "&Sum", Sis_Rpt_Number2
      .AddTableFooter "&Sum", Sis_Rpt_Number2
              
    .Preview vaArray, True
  End With
End Sub

Private Sub cmdRebutan_Click()
Dim cSQL As String
Dim db As New ADODB.Recordset
Dim n As Single
Dim a As New exportExcel

  cSQL = cSQL & " SELECT s.barcode,s.nama,s.hargajual,h.qty as adv,k.qty as kuota,a.kodeanggota,a.nama as namamember,m.qty as qtyorder from orderpromo_rebutan h"
  cSQL = cSQL & " left join stock s on s.barcode = h.barcode"
  cSQL = cSQL & " left join orderpromo_kuota k on k.barcode = s.barcode"
  cSQL = cSQL & " left join memberorder m on m.kodestock = s.kodestock"
  cSQL = cSQL & " left join totmemberorder t on t.nomormemberorder = m.nomormemberorder"
  cSQL = cSQL & " left join anggota a on a.kodeanggota = t.kodeanggota"
  cSQL = cSQL & " Where k.qty Is Not Null And m.Tgl >= '2012-06-01'"
  cSQL = cSQL & " ORDER BY s.barcode"
  
  vaArray.ReDim 0, -1, 0, 10
  Set db = objData.SQL(GetDSN, cSQL)
  If Not db.EOF Then
    FrmPB.InitPB db.RecordCount
    Do While Not db.EOF
      FrmPB.RunPB
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(db!barcode)
      vaArray(n, 1) = GetNull(db!nama)
      vaArray(n, 2) = GetNull(db!HargaJual)
      vaArray(n, 3) = GetNull(db!adv)
      vaArray(n, 4) = GetNull(db!kuota)
      vaArray(n, 5) = GetNull(db!kodeanggota)
      vaArray(n, 6) = GetNull(db!namamember)
      vaArray(n, 7) = GetNull(db!qtyorder)
      vaArray(n, 8) = GetSaldoTopUpMember2(objData, GetNull(db!kodeanggota))
      vaArray(n, 9) = GetBerapaSudahDapatPromo(objData, GetNull(db!kodeanggota))
      vaArray(n, 10) = GetBerapaDiaOrderPromo(objData, GetNull(db!kodeanggota))
      db.MoveNext
    Loop
    FrmPB.EndPB
    a.RecordSource = vaArray
    a.ExportToExcel
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  dDate.Value = Date
  TabIndex dDate, n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub




