VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Begin VB.Form rptRekapBarangPromo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rekap Barang Promo"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   7560
   Begin VB.OptionButton optProses 
      Caption         =   "Belum Diproses"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   2385
      TabIndex        =   8
      Top             =   1635
      Width           =   1500
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   2280
      Left            =   15
      Top             =   0
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   4022
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
      Begin VB.OptionButton optProses 
         Caption         =   "Semuanya"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   3975
         TabIndex        =   9
         Top             =   1650
         Width           =   1155
      End
      Begin VB.OptionButton optProses 
         Caption         =   "Sudah Diproses"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   735
         TabIndex        =   7
         Top             =   1635
         Width           =   1500
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Pilih Member"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2700
         TabIndex        =   0
         Top             =   840
         Width           =   1305
      End
      Begin BiSATextBoxProject.BiSABrowse cNamaCustomer 
         Height          =   330
         Left            =   4350
         TabIndex        =   1
         Top             =   1140
         Width           =   2670
         _ExtentX        =   4710
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
      Begin BiSATextBoxProject.BiSABrowse cCustomer 
         Height          =   330
         Left            =   570
         TabIndex        =   2
         Top             =   1140
         Width           =   3765
         _ExtentX        =   6641
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
         Button          =   -1  'True
         Caption         =   "Member"
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
      Begin BiSATextBoxProject.BiSABrowse cKatalog 
         Height          =   330
         Left            =   555
         TabIndex        =   5
         Top             =   375
         Width           =   3795
         _ExtentX        =   6694
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
         Button          =   -1  'True
         Caption         =   "Katalog"
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
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4425
         TabIndex        =   6
         Top             =   390
         Width           =   2370
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   0
      Top             =   2265
      Width           =   7545
      _ExtentX        =   13309
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
         Left            =   6360
         TabIndex        =   3
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
         Picture         =   "rptRekapBarangPromo.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   5925
         TabIndex        =   4
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
         Picture         =   "rptRekapBarangPromo.frx":00A6
      End
   End
End
Attribute VB_Name = "rptRekapBarangPromo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub cCustomer_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "anggota", "kodeanggota,nama,alamat", "kodeanggota", sisContent, cCustomer.Text, , "kodeanggota")
  If Not dbData.EOF Then
    cCustomer.Text = cCustomer.Browse(dbData)
    cNamaCustomer.Text = GetNull(dbData!nama)
  End If
End Sub

Private Sub cKatalog_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "katalog", "kodekatalog,keterangan", "kodekatalog", sisContent, cKatalog.Text, , "kodekatalog")
  If Not dbData.EOF Then
    cKatalog.Text = cKatalog.Browse(dbData)
    Label1.Caption = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
 ' GetSQL
  GetSQL2
  
End Sub

Private Sub cNamaCustomer_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "anggota", "nama,kodeanggota,alamat", "nama", sisContent, cNamaCustomer.Text, , "nama")
  If Not dbData.EOF Then
    cNamaCustomer.Text = cNamaCustomer.Browse(dbData)
    cCustomer.Text = GetNull(dbData!kodeanggota)
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  SetIcon Me.hWnd
  
  TabIndex cKatalog, n
  TabIndex cCustomer, n
  TabIndex cNamaCustomer, n
  
  cKatalog.Default
  Label1.Caption = ""
  Check1.Value = 0
  cCustomer.Default
  cNamaCustomer.Default
  optProses(2).Value = True
  
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub

Private Sub GetSQL()
Dim cSQL As String
Dim n As Single
  
  vaArray.ReDim 0, -1, 0, 8
  
  cSQL = ""
  cSQL = "select m.kodeanggota,a.nama as namaanggota,m.tgl,m.barcode,s.barcode,s.nama,m.qty,s.hargajual,s.diskonpenjualan from totkatalogpromo t"
  cSQL = cSQL & " left join promo m on m.nomorpromo = t.nomorpromo"
  cSQL = cSQL & " left join stock s on s.kodestock = m.kodestock"
  cSQL = cSQL & " left join anggota a on a.kodeanggota = m.kodeanggota"
  cSQL = cSQL & " Where t.kodekatalog = '" & cKatalog.Text & "'"
  If Check1.Value = 1 Then
    cSQL = cSQL & " and (a.kodeanggota = '" & cCustomer.Text & "')"
  End If

'  cSQL = cSQL & " and t.kodekatalog = '" & cKatalog.Text & "'"
'  cSQL = cSQL & " and (m.tgl >='" & Format(dTgl(0).Value, "yyyy-MM-dd") & "' and m.tgl <='" & Format(dTgl(1).Value, "yyyy-MM-dd") & "') "
  
  cSQL = cSQL & " order by a.kodeanggota,m.kodestock,m.tgl"
  Set dbData = objData.Sql(GetDSN, cSQL)
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!kodeanggota)
      vaArray(n, 1) = GetNull(dbData!namaanggota)
      vaArray(n, 2) = GetNull(dbData!barcode)
      vaArray(n, 3) = Format(GetNull(dbData!Tgl), "dd MM yyyy")
      vaArray(n, 4) = GetNull(dbData!nama)
      vaArray(n, 5) = GetNull(dbData!qty)
      vaArray(n, 6) = 0 'GetNull(dbData!diskonpenjualan)
      vaArray(n, 7) = GetNull(dbData!hargajual)
      vaArray(n, 8) = vaArray(n, 5) * (vaArray(n, 7) - (vaArray(n, 7) * vaArray(n, 6) / 100)) 'vaArray(n, 5) * vaArray(n, 7)
      
      dbData.MoveNext
    Loop
    GetRpt
  Else
    MsgBox "Maaf, tidak ada data untuk ditampilkan"
  End If

End Sub

Private Sub GetSQL2()
Dim cSQL As String
Dim n As Single
  
  vaArray.ReDim 0, -1, 0, 7
  
'  cSQL = ""
'  cSQL = "select m.kodeanggota,a.nama as namaanggota,m.tgl,m.barcode,s.barcode,s.nama,m.qty,s.hargajual,s.diskonpenjualan from totkatalogpromo t"
'  cSQL = cSQL & " left join promo m on m.nomorpromo = t.nomorpromo"
'  cSQL = cSQL & " left join stock s on s.kodestock = m.kodestock"
'  cSQL = cSQL & " left join anggota a on a.kodeanggota = m.kodeanggota"
'  cSQL = cSQL & " Where t.kodekatalog = '" & cKatalog.Text & "'"
'
'  cSQL = cSQL & " order by a.kodeanggota,m.kodestock,m.tgl"
'
cSQL = ""
cSQL = cSQL & " select f.*,k.nmbrg as barang from finalsort f"
cSQL = cSQL & " left join katalogpromo k on k.kdbrg = f.ref "
'cSQL = cSQL & " LEFT JOIN member  m on m.nama = f.nama"
'cSQL = cSQL & " where f.`group` = '" & cCustomer.Text & "'"
cSQL = cSQL & " ORDER BY F.`group`,f.nama"

  Set dbData = objData.Sql(GetDSN, cSQL)
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!telp) '
      vaArray(n, 1) = GetNull(dbData!nama) '
      vaArray(n, 2) = GetNull(dbData!ref)
      vaArray(n, 3) = GetNull(dbData!barang)
      vaArray(n, 4) = GetNull(dbData!qty)
      vaArray(n, 5) = GetNull(dbData!Harga)
      vaArray(n, 6) = vaArray(n, 4) * vaArray(n, 5)
      vaArray(n, 7) = GetNull(dbData!Group)
      dbData.MoveNext
    Loop
    GetRpt
  Else
    MsgBox "Maaf, tidak ada data untuk ditampilkan"
  End If

End Sub

Private Sub GetRpt()
  With FrmRPT
    
    .AddPageHeader "Rekap Advance Promo", tdbHalignCenter, , , True, , 10, True, False, True, False, tdbPageHeaderSect
    .AddPageHeader aCfg(objData, msNamaPerusahaan), tdbHalignCenter, , , True, , 14, True, False, True, False, tdbPageHeaderSect
        
    .AddTableGroupHeader True, , , , , 11
    .AddTableGroupHeader , , , , , 50
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    
    .AddTableHeader "", , , , , , , , , , , , , , , , , , , False
    .AddTableHeader "", , , , , , , , , , , , , , , , , , , False
    .AddTableHeader "REF", , , , 11
    .AddTableHeader "BARANG"
    .AddTableHeader "QTY", , , , 5
    .AddTableHeader "HARGA", , , , 12
    .AddTableHeader "JUMLAH", , , , 11
    .AddTableHeader "GROUP", , , , 15
    
     
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody
    .AddTableBody
    .AddTableBody Sis_Rpt_Number
    .AddTableBody Sis_Rpt_Number
    .AddTableBody Sis_Rpt_Number
    .AddTableBody
    
    .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupFooter "SubTotal", , tdbHalignRight, , , , , , , , , , , , 2
    .AddTableGroupFooter
    .AddTableGroupFooter "&Sum", Sis_Rpt_Number
    .AddTableGroupFooter
    .AddTableGroupFooter "&Sum", Sis_Rpt_Number
    .AddTableGroupFooter
    
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableFooter "Total", , tdbHalignRight, , , , , , , , , , , , 2
    .AddTableFooter
    .AddTableFooter "&Sum", Sis_Rpt_Number
    .AddTableFooter
    .AddTableFooter "&Sum", Sis_Rpt_Number
    .AddTableFooter ""
    
    .Refresh
    .Preview vaArray, True
  End With
End Sub

