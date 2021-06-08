VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Begin VB.Form cfgMarkUP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mark Up Harga Jual"
   ClientHeight    =   5844
   ClientLeft      =   48
   ClientTop       =   348
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5844
   ScaleWidth      =   9780
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   648
      Left            =   120
      Top             =   5136
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   1143
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
         Left            =   8220
         TabIndex        =   0
         Top             =   90
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   762
         Caption         =   "     &Exit"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         BackColor       =   -2147483633
         Picture         =   "cfgMarkUP.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   7140
         TabIndex        =   1
         Top             =   90
         Width           =   1065
         _ExtentX        =   1884
         _ExtentY        =   762
         Caption         =   "    &Save"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         BackColor       =   -2147483633
         Picture         =   "cfgMarkUP.frx":00A6
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   4920
      Left            =   120
      Top             =   210
      Width           =   9510
      _ExtentX        =   16785
      _ExtentY        =   8678
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      Begin BiSAButtonProject.BiSAButton BiSAButton9 
         Height          =   405
         Left            =   5505
         TabIndex        =   12
         Top             =   1650
         Width           =   3615
         _ExtentX        =   6371
         _ExtentY        =   720
         Caption         =   "Label1"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
      End
      Begin BiSAButtonProject.BiSAButton BiSAButton1 
         Height          =   390
         Left            =   525
         TabIndex        =   2
         Top             =   1080
         Width           =   3690
         _ExtentX        =   6519
         _ExtentY        =   699
         Caption         =   "Update HP pada Tabel Stock"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
      End
      Begin BiSANumberBoxProject.BiSANumberBox nMarkup 
         Height          =   330
         Left            =   585
         TabIndex        =   3
         Top             =   660
         Width           =   2895
         _ExtentX        =   5101
         _ExtentY        =   593
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Mark UP Harga %"
         CaptionWidth    =   1700
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin BiSAButtonProject.BiSAButton BiSAButton2 
         Height          =   420
         Left            =   525
         TabIndex        =   4
         Top             =   1500
         Width           =   3675
         _ExtentX        =   6477
         _ExtentY        =   741
         Caption         =   "Upate hp in table kartu stock dan bukubesar"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
      End
      Begin BiSAButtonProject.BiSAButton BiSAButton3 
         Height          =   390
         Left            =   525
         TabIndex        =   5
         Top             =   2460
         Width           =   3690
         _ExtentX        =   6519
         _ExtentY        =   699
         Caption         =   "update tothp in kartustock"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
      End
      Begin BiSAButtonProject.BiSAButton BiSAButton4 
         Height          =   390
         Left            =   525
         TabIndex        =   7
         Top             =   2925
         Width           =   3690
         _ExtentX        =   6519
         _ExtentY        =   699
         Caption         =   "update stok qty di tabel master stok"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
      End
      Begin BiSAButtonProject.BiSAButton BiSAButton5 
         Height          =   390
         Left            =   510
         TabIndex        =   8
         Top             =   3360
         Width           =   3690
         _ExtentX        =   6519
         _ExtentY        =   699
         Caption         =   "Update oustanding di table Anggota"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
      End
      Begin BiSAButtonProject.BiSAButton BiSAButton6 
         Height          =   390
         Left            =   495
         TabIndex        =   9
         Top             =   3780
         Width           =   3690
         _ExtentX        =   6519
         _ExtentY        =   699
         Caption         =   "Update oustanding di table Supplier"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
      End
      Begin BiSAButtonProject.BiSAButton BiSAButton7 
         Height          =   390
         Left            =   480
         TabIndex        =   10
         Top             =   4200
         Width           =   3690
         _ExtentX        =   6519
         _ExtentY        =   699
         Caption         =   "Posting Stok Awal  ke Akunting"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
      End
      Begin BiSAButtonProject.BiSAButton BiSAButton8 
         Height          =   390
         Left            =   5460
         TabIndex        =   11
         Top             =   1170
         Width           =   3690
         _ExtentX        =   6519
         _ExtentY        =   699
         Caption         =   "Posting Saldo Awal Piutang"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
      End
      Begin VB.Label Label2 
         Caption         =   "Harga khusus penjualan ke Non Anggota/Member"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   645
         TabIndex        =   6
         Top             =   135
         Width           =   2910
      End
   End
End
Attribute VB_Name = "cfgMarkUP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data

Private Sub BiSAButton1_Click()
Dim cSQL As String

  cSQL = "select kodestock,nama,hargabeli from stock"
  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.EOF
      FrmPB.RunPB
      objData.Edit GetDSN, "stock", "kodestock = '" & GetNull(dbData!KodeStock) & "'", Array("cogs"), Array(GetUpdateHPPStock(GetNull(dbData!KodeStock)))
      dbData.MoveNext
    Loop
    FrmPB.EndPB
  End If
End Sub

Private Function GetUpdateHPPStock(cKdStock As String) As Double
Dim db As New ADODB.Recordset
Dim cSQL As String
Dim obj As New CodeSuiteLibrary.Data

  cSQL = "select s.kodestock,s.nama,s.hargabeli,sum(k.debet) as totqty,sum(k.debet*k.harga)as tothp from kartustock k"
  cSQL = cSQL & " LEFT JOIN stock s on s.kodestock = k.kodestock"
  cSQL = cSQL & " Where s.KodeStock = '" & cKdStock & "' and status = '10'"
  cSQL = cSQL & " GROUP BY s.kodestock"
  
  Set db = obj.SQL(GetDSN, cSQL)
  If Not db.EOF Then
    If GetNull(db!totqty) = 0 Then
      GetUpdateHPPStock = GetNull(db!hargabeli)
    Else
      GetUpdateHPPStock = GetNull(db!tothp) / GetNull(db!totqty)
    End If
  End If
  
End Function

Private Function GetHargaCOGS(cKdStock As String) As Double
Dim db As New ADODB.Recordset
Dim cSQL As String
Dim obj As New CodeSuiteLibrary.Data

  GetHargaCOGS = 0
  cSQL = "select cogs,hargabeli from stock where kodestock = '" & cKdStock & "'"
  Set db = obj.SQL(GetDSN, cSQL)
  If Not db.EOF Then
    GetHargaCOGS = GetNull(db!cogs)
    If GetNull(db!cogs) <= 0 Then
      GetHargaCOGS = GetNull(db!hargabeli)
    End If
  End If
End Function

Private Sub BiSAButton2_Click()
Dim cSQL As String

  cSQL = "select k.id,k.nomor,k.kodestock,k.hp,k.qty,s.nama from kartustock k left join stock s on s.kodestock = k.kodestock"
  cSQL = cSQL & " where status = '60'"
  
  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.EOF
      FrmPB.RunPB
      Dim cCogs As Double
      Dim totCogs As Double
      cCogs = GetHargaCOGS(GetNull(dbData!KodeStock))
      totCogs = GetHargaCOGS(GetNull(dbData!KodeStock)) * GetNull(dbData!qty)
      'objData.Edit GetDSN, "kartustock", "nomor = '" & GetNull(dbData!nomor) & "' and kodestock = '" & GetNull(dbData!KodeStock) & "'", Array("hp", "tothp"), Array(cCogs, totCogs)
      objData.Edit GetDSN, "kartustock", "id = '" & GetNull(dbData!Id) & "'", Array("hp", "tothp"), Array(cCogs, totCogs)
      objData.Edit GetDSN, "bukubesar", "faktur = '" & GetNull(dbData!nomor) & "' and status ='4' and kodeakun = '5.200' and keterangan = 'COGS Penjualan an " & GetNull(dbData!nama) & "'", Array("debet", "kodestock"), Array(totCogs, GetNull(dbData!KodeStock))
      objData.Edit GetDSN, "bukubesar", "faktur = '" & GetNull(dbData!nomor) & "' and status ='4' and kodeakun = '1.400' and keterangan = 'COGS Penjualan an " & GetNull(dbData!nama) & "'", Array("kredit", "kodestock"), Array(totCogs, GetNull(dbData!KodeStock))
      dbData.MoveNext
    Loop
    FrmPB.EndPB
  End If
End Sub

Private Sub BiSAButton3_Click()
Dim cSQL As String

  cSQL = "select id,qty,hp from kartustock"
  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.EOF
      FrmPB.RunPB
      objData.Edit GetDSN, "kartustock", "id = '" & GetNull(dbData!Id) & "'", Array("tothp"), Array(GetNull(dbData!qty) * GetNull(dbData!hp))
      dbData.MoveNext
    Loop
    FrmPB.EndPB
  End If
End Sub

Private Sub BiSAButton4_Click()
Dim cSQL As String

  cSQL = "select kodestock,nama,hargabeli from stock"
  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.EOF
      FrmPB.RunPB
      objData.Edit GetDSN, "stock", "kodestock = '" & GetNull(dbData!KodeStock) & "'", Array("stok"), Array(GetSaldoStock(objData, "", GetNull(dbData!KodeStock)))
      dbData.MoveNext
    Loop
    FrmPB.EndPB
  End If
End Sub

Private Sub BiSAButton5_Click()
Dim cSQL As String

  cSQL = "SELECT kodeanggota,sum(debet-kredit)  as saldo from kartupiutang"
  cSQL = cSQL & " GROUP BY kodeanggota"
  cSQL = cSQL & " Having Sum(debet - kredit) > 0"
  cSQL = cSQL & " ORDER BY kodeanggota"
  
  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.EOF
      FrmPB.RunPB
      objData.Edit GetDSN, "anggota", "kodeanggota = '" & GetNull(dbData!kodeanggota) & "'", Array("outstanding"), Array(GetNull(dbData!saldo))
      dbData.MoveNext
    Loop
    FrmPB.EndPB
  End If
End Sub

Private Sub BiSAButton6_Click()
Dim cSQL As String

  cSQL = "SELECT kodesupplier,sum(debet-kredit)  as saldo from kartuhutang"
  cSQL = cSQL & " GROUP BY kodesupplier"
  cSQL = cSQL & " Having Sum(debet - kredit) > 0"
  cSQL = cSQL & " ORDER BY kodesupplier"
  
  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.EOF
      FrmPB.RunPB
      objData.Edit GetDSN, "supplier", "kodesupplier = '" & GetNull(dbData!kodesupplier) & "'", Array("outstanding"), Array(GetNull(dbData!saldo))
      dbData.MoveNext
    Loop
    FrmPB.EndPB
  End If
End Sub

Private Sub BiSAButton7_Click()
Dim cSQL As String

  cSQL = "select * from stock where jenis = 1 and outsource = 'T' and asbiaya = 2 and statusnonaktif = 0"
  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.EOF
      FrmPB.RunPB
      GetPostingSaldoAwalStock objData, GetNull(dbData!KodeStock), GetNull(dbData!hargabeli), GetNull(dbData!SaldoAwal), GetNull(dbData!nama)
      dbData.MoveNext
    Loop
    FrmPB.EndPB
    MsgBox "Proses Selesai"
  End If
End Sub

Private Function GetPostingSaldoAwalStock(ByVal objData As CodeSuiteLibrary.Data, ByVal cKodeStk As String, ByVal nHargaBeli As Double, ByVal nQty As Double, Optional ByVal cNamaBarang As String = "") As Boolean
Dim cSQL As String
Dim Faktur As String
Dim dTglNow As Date
Dim cKet As String
Dim cGudangAwal As String
Dim lSave As Boolean

  objData.Start GetDSN
  lSave = True
  GetPostingSaldoAwalStock = True
  dTglNow = SNow
  Faktur = "AWALSTK-" & cKodeStk
  cKet = "Saldo Awal Stock SKU-" & cKodeStk & " " & cNamaBarang
  cGudangAwal = aCfg(objData, msGudangPembelian)
  
  lSave = IIf(lSave, objData.Delete(GetDSN, "kartustock", "nomor", sisAssign, Faktur), False)
  'lSave = IIf(lSave, objData.Edit(GetDSN, "stock", "kodestock = '" & cKodeStk & "'", Array("stok"), Array(1)), False)
  lSave = IIf(lSave, UpdKartuStock(objData, SisKartuStock.SaldoAwal, Faktur, Format(dTglNow, "yyyy-MM-dd"), GetNull(dbData!KodeStock), nQty, nHargaBeli, 0, cKet, cGudangAwal, nHargaBeli), False)
  lSave = IIf(lSave, DelKodeTr(objData, msSaldoAwalStock, Faktur), False)
  
  
  '[D]ebet]
  'Rekening : Persediaan
  lSave = IIf(lSave, UpdKodeTr(objData, msSaldoAwalStock, Faktur, Format(dTglNow, "yyyy-MM-dd"), GetAkunInventory(objData, cKodeStk), GetCostCenterUser(objData, GetRegistry(reg_Username)), cKet, nHargaBeli * nQty, 0, "", SNow, cKodeStk), False)
        '[K]redit
        'Rekening : Stock Awal
        lSave = IIf(lSave, UpdKodeTr(objData, msSaldoAwalStock, Faktur, Format(dTglNow, "yyyy-MM-dd"), "3.300.11", GetCostCenterUser(objData, GetRegistry(reg_Username)), cKet, 0, nHargaBeli * nQty, "", SNow, cKodeStk), False)
  
  GetPostingSaldoAwalStock = lSave
  If lSave = True Then
    'simpan
    objData.Save GetDSN
  Else
    objData.Cancel GetDSN
    MsgBox "ERR: " & cKet
  End If
End Function

Private Sub BiSAButton8_Click()
Dim cSQL As String
Dim cKodeFaktur As String
Dim dTgl As Date

  cKodeFaktur = InputBox("Masukkan KodeFaktur yg masu diposting", "Input KodeFaktur")
  dTgl = InputBox("Masukkan tgl :", "Input Tanggal")
  
  cSQL = "select * from saldoawalpiutang t where t.statusposting='0' and t.kodefaktur='" & cKodeFaktur & "'"
  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.EOF Then
    objData.Delete GetDSN, "kartupiutang", "nomorkartupiutang", sisAssign, cKodeFaktur
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.EOF
      FrmPB.RunPB
      'simpan di kartupiutang
      UpdKartuHutang objData, SisMemberBalance, cKodeFaktur, dTgl, GetNull(dbData!kodeanggota), GetNull(dbData!keterangan), GetNull(dbData!Total), SNow, GetRegistry(reg_Username), False, GetNull(dbData!KodeGroupSales)
      'update status kalau sudah di posting
      objData.Edit GetDSN, "saldoawalpiutang", "id='" & GetNull(dbData!Id) & "'", Array("statusposting"), Array(sisFlag.Posting)
      dbData.MoveNext
    Loop
    FrmPB.EndPB
    MsgBox "Proses Selesai"
  Else
    MsgBox "Maaf data tidak ada", vbExclamation
  End If
End Sub

Private Sub BiSAButton9_Click()
Dim cSQL As String
Dim cKodeFaktur As String
Dim dTgl As Date

  cKodeFaktur = InputBox("Masukkan KodeFaktur yg mau dihapus", "Input KodeFaktur")
  dTgl = InputBox("Masukkan tgl :", "Input Tanggal")
  cSQL = "select * from saldoawalpiutang t where t.statusposting='1' and t.kodefaktur='" & cKodeFaktur & "'"
  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.EOF
      FrmPB.RunPB
      'delete di kartupiutang
      objData.Delete GetDSN, "kartupiutang ", "nomorkartupiutang", sisAssign, cKodeFaktur
      'update status kalau sudah di delete
      objData.Edit GetDSN, "saldoawalpiutang", "id='" & GetNull(dbData!Id) & "'", Array("statusposting"), Array(sisFlag.Nul)
      dbData.MoveNext
    Loop
    FrmPB.EndPB
    MsgBox "Proses Selesai"
  Else
    MsgBox "Maaf. Data tidak ada", vbCritical
  End If
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdSimpan_Click()
  objData.Start GetDSN
  UpdCfg msMarkUpHargaJual, nMarkUp.Value, objData, nMarkUp.Caption, Me.Caption
  objData.Save GetDSN
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  SetIcon Me.hWnd
  TabIndex nMarkUp, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  nMarkUp.Value = aCfg(objData, msMarkUpHargaJual)
End Sub
