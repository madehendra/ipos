VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptLaporanOrderByMember 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laporan Order by Member"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   7575
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1905
      Left            =   15
      Top             =   0
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   3360
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
         TabIndex        =   6
         Top             =   840
         Width           =   1305
      End
      Begin BiSADateProject.BiSADate dTgl 
         Height          =   330
         Index           =   0
         Left            =   555
         TabIndex        =   0
         Top             =   390
         Width           =   3465
         _ExtentX        =   6112
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
         Caption         =   "Tanggal"
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
      Begin BiSADateProject.BiSADate dTgl 
         Height          =   330
         Index           =   1
         Left            =   4110
         TabIndex        =   1
         Top             =   390
         Width           =   2025
         _ExtentX        =   3572
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
         CaptionWidth    =   500
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
      Begin BiSATextBoxProject.BiSABrowse cNamaCustomer 
         Height          =   330
         Left            =   4350
         TabIndex        =   4
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
         TabIndex        =   5
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
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   0
      Top             =   1890
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
         TabIndex        =   2
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
         Picture         =   "rptLaporanOrderByMamber.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   5925
         TabIndex        =   3
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
         Picture         =   "rptLaporanOrderByMamber.frx":00A6
      End
   End
End
Attribute VB_Name = "rptLaporanOrderByMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub ccustomer_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "anggota", "kodeanggota,nama,alamat", "kodeanggota", sisContent, cCustomer.Text, , "kodeanggota")
  If Not dbData.EOF Then
    cCustomer.Text = cCustomer.Browse(dbData)
    cNamaCustomer.Text = GetNull(dbData!nama)
  End If
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
  GetSQL
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
  
  TabIndex dTgl(0), n
  TabIndex dTgl(1), n
  TabIndex cCustomer, n
  TabIndex cNamaCustomer, n
  
  Check1.Value = 0
  cCustomer.Default
  cNamaCustomer.Default
  
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub

Private Sub GetSQL()
Dim cSQL As String
Dim n As Single
  
  vaArray.ReDim 0, -1, 0, 8
  
  cSQL = "select t.kodeanggota,a.nama as namaanggota,m.tgl,m.kodestock,s.barcode,s.nama,m.qty,m.harga,m.discount from totmemberorder t"
  cSQL = cSQL & " left join memberorder m on m.nomormemberorder = t.nomormemberorder"
  cSQL = cSQL & " left join stock s on s.kodestock = m.kodestock"
  cSQL = cSQL & " left join anggota a on a.kodeanggota = t.kodeanggota"
  cSQL = cSQL & " Where (t.status = 0 Or t.status Is Null)"
  If Check1.Value = 1 Then
    cSQL = cSQL & " and (a.kodeanggota = '" & cCustomer.Text & "')"
  End If
  
  cSQL = cSQL & " and (m.tgl >='" & Format(dTgl(0).Value, "yyyy-MM-dd") & "' and m.tgl <='" & Format(dTgl(1).Value, "yyyy-MM-dd") & "') "
  
  cSQL = cSQL & " order by a.kodeanggota,m.kodestock,m.tgl"
  Set dbData = objData.Sql(GetDSN, cSQL)
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!kodeanggota)
      vaArray(n, 1) = GetNull(dbData!namaanggota)
      vaArray(n, 2) = GetNull(dbData!Barcode)
      vaArray(n, 3) = Format(GetNull(dbData!tgl), "dd MM yyyy")
      vaArray(n, 4) = GetNull(dbData!nama)
      vaArray(n, 5) = GetNull(dbData!qty)
      vaArray(n, 6) = GetNull(dbData!Discount)
      vaArray(n, 7) = GetNull(dbData!Harga)
      vaArray(n, 8) = vaArray(n, 5) * (vaArray(n, 7) - (vaArray(n, 7) * vaArray(n, 6) / 100)) 'vaArray(n, 5) * vaArray(n, 7)
      
      dbData.MoveNext
    Loop
    GetRpt
  Else
    MsgBox "Maaf, tidak ada data untuk ditampilkan"
  End If

End Sub
Private Sub GetRpt()
  With FrmRPT
    
    .AddPageHeader "Laporan Order Yg di Group By Member", tdbHalignCenter, , , True, , 10, True, False, True, False, tdbPageHeaderSect
    .AddPageHeader aCfg(objData, msNamaPerusahaan), tdbHalignCenter, , , True, , 14, True, False, True, False, tdbPageHeaderSect
        
    .AddTableGroupHeader True, "[]", , , , 15
    .AddTableGroupHeader
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    
    
    .AddTableHeader "", , , , , , , , , , , , , , , , , , , False
    .AddTableHeader "", , , , , , , , , , , , , , , , , , , False
    .AddTableHeader "Kode", , , , 11
    .AddTableHeader "Tgl", , , , 11
    .AddTableHeader "Nama"
    .AddTableHeader "Qty", , , , 5
    .AddTableHeader "Dsc", , , , 5
    .AddTableHeader "Harga", , , , 13
    .AddTableHeader "Total", , , , 15
    
     
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody Sis_Rpt_Number
    .AddTableBody Sis_Rpt_Number
    .AddTableBody Sis_Rpt_Number
    .AddTableBody Sis_Rpt_Number
    
    .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupFooter "SubTotal", , tdbHalignRight, , , , , , , , , , , , 6
    .AddTableGroupFooter
    .AddTableGroupFooter
    .AddTableGroupFooter
    .AddTableGroupFooter
    .AddTableGroupFooter
    .AddTableGroupFooter "&Sum", Sis_Rpt_Number2
    
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableFooter "Total", , tdbHalignRight, , , , , , , , , , , , 6
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter "&Sum", Sis_Rpt_Number
    
    .Refresh
    .Preview vaArray, True
  End With
End Sub
