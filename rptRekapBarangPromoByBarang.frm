VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Begin VB.Form rptRekapBarangPromoByBarang 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rekap Barang Promo Group By Barang"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   7545
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
      Begin VB.CheckBox Check1 
         Caption         =   "Pilih Barang"
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
      Begin BiSATextBoxProject.BiSABrowse cNamaBarang 
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
      Begin BiSATextBoxProject.BiSABrowse cBarcode 
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
         Caption         =   "Barang"
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
         TabIndex        =   3
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
         TabIndex        =   4
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
         Picture         =   "rptRekapBarangPromoByBarang.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   5925
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
         Picture         =   "rptRekapBarangPromoByBarang.frx":00A6
      End
   End
End
Attribute VB_Name = "rptRekapBarangPromoByBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub cBarcode_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "stock", "barcode,nama", "barcode", sisContent, cBarcode.Text, , "barcode")
  If Not dbData.EOF Then
    cBarcode.Text = cBarcode.Browse(dbData)
    cNamaBarang.Text = GetNull(dbData!nama)
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
  GetSQL
End Sub

Private Sub cNamaBarang_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "stock", "nama,barcode", "nama", sisContent, cNamaBarang.Text, , "nama")
  If Not dbData.EOF Then
    cNamaBarang.Text = cNamaBarang.Browse(dbData)
    cBarcode.Text = GetNull(dbData!Barcode)
  End If
End Sub



Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  SetIcon Me.hWnd
  
  TabIndex cKatalog, n
  TabIndex cBarcode, n
  TabIndex cNamaBarang, n
  
  cKatalog.Default
  Label1.Caption = ""
  Check1.Value = 0
  cBarcode.Default
  cNamaBarang.Default
  
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub

Private Sub GetSQL()
Dim cSQL As String
Dim n As Single
  
  vaArray.ReDim 0, -1, 0, 5
  
  cSQL = ""
  cSQL = "select p.barcode,s.nama as namabarang,a.nama as namaanggota,p.qty,p.nomorpromo,p.tgl from promo p"
  cSQL = cSQL & " left join stock s on s.kodestock = p.kodestock"
  cSQL = cSQL & " left join anggota a on a.kodeanggota = p.kodeanggota"
  cSQL = cSQL & " Where p.kodekatalog = '" & cKatalog.Text & "'"
  If Check1.Value = 1 Then
    cSQL = cSQL & " and (p.barcode = '" & cBarcode.Text & "')"
  End If

  
  cSQL = cSQL & " order by p.barcode,s.nama"
  Set dbData = objData.Sql(GetDSN, cSQL)
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!Barcode)
      vaArray(n, 1) = GetNull(dbData!namabarang)
      vaArray(n, 2) = GetNull(dbData!namaanggota)
      vaArray(n, 3) = GetNull(dbData!qty)
      vaArray(n, 4) = GetNull(dbData!nomorpromo)
      vaArray(n, 5) = Format(GetNull(dbData!tgl), "dd/MM/yyyy")
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
        
    .AddTableGroupHeader True, "[]", , , , 15
    .AddTableGroupHeader
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    
    
    .AddTableHeader "", , , , , , , , , , , , , , , , , , , False
    .AddTableHeader "", , , , , , , , , , , , , , , , , , , False
    .AddTableHeader "Member", , , , 35
    .AddTableHeader "qty", , , , 6
    .AddTableHeader "nomor", , , , 14
    .AddTableHeader "tgl", , , , 10
     
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody
    
    
    .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupFooter
    .AddTableGroupFooter "&Sum", Sis_Rpt_Number
    .AddTableGroupFooter
    .AddTableGroupFooter
    
    
    
'    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
'    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
'    .AddTableFooter "Total", , tdbHalignRight, , , , , , , , , , , , 6
'    .AddTableFooter
'    .AddTableFooter
'    .AddTableFooter
'    .AddTableFooter
'    .AddTableFooter
'    .AddTableFooter "&Sum", Sis_Rpt_Number
'
    .Refresh
    .Preview vaArray, True
  End With
End Sub


