VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptDetailItemPenjualanPerAnggota 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detail Item Penjualan Per Anggota"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7545
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   7545
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   2295
      Left            =   15
      Top             =   0
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   4048
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
      Begin BiSAFramProject.BiSAFrame BiSAFrame2 
         Height          =   465
         Left            =   2415
         Top             =   1635
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   820
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
         Begin VB.OptionButton optOpsi 
            Caption         =   "Semua"
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
            Index           =   2
            Left            =   1890
            TabIndex        =   10
            Top             =   120
            Width           =   810
         End
         Begin VB.OptionButton optOpsi 
            Caption         =   "Kredit"
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
            Index           =   1
            Left            =   1020
            TabIndex        =   9
            Top             =   120
            Width           =   810
         End
         Begin VB.OptionButton optOpsi 
            Caption         =   "Tunai"
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
            Index           =   0
            Left            =   165
            TabIndex        =   8
            Top             =   120
            Width           =   810
         End
         Begin VB.Label Label1 
            Caption         =   "* Print All Only"
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
            Left            =   2850
            TabIndex        =   11
            Top             =   120
            Width           =   1320
         End
      End
      Begin BiSATextBoxProject.BiSABrowse cNama 
         Height          =   330
         Left            =   330
         TabIndex        =   0
         Top             =   540
         Width           =   5520
         _ExtentX        =   9737
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
         Index           =   0
         Left            =   330
         TabIndex        =   1
         Top             =   1290
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
      Begin BiSATextBoxProject.BiSABrowse cKode 
         Height          =   330
         Left            =   330
         TabIndex        =   2
         Top             =   165
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   582
         Text            =   "123456"
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
         Caption         =   "Kode Anggota"
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
         Left            =   4185
         TabIndex        =   3
         Top             =   1290
         Width           =   2010
         _ExtentX        =   3545
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
      Begin BiSATextBoxProject.BiSABrowse cAlamat 
         Height          =   330
         Left            =   330
         TabIndex        =   4
         Top             =   915
         Width           =   5505
         _ExtentX        =   9710
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
         Caption         =   "Dept"
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
      Top             =   2280
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
         Picture         =   "rptDetailItemPenjualanPerAnggota.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   3270
         TabIndex        =   6
         Top             =   120
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   767
         Caption         =   "Selected Print"
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
         Picture         =   "rptDetailItemPenjualanPerAnggota.frx":00A6
      End
      Begin BiSAButtonProject.BiSAButton cmdPrintAll 
         Height          =   435
         Left            =   165
         TabIndex        =   7
         Top             =   120
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   767
         Caption         =   "Print All"
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
         Picture         =   "rptDetailItemPenjualanPerAnggota.frx":032C
      End
   End
End
Attribute VB_Name = "rptDetailItemPenjualanPerAnggota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vaArray As New XArrayDB
Dim objData As New CodeSuiteLibrary.Data
Dim dbData As New ADODB.Recordset

Private Sub GetData()
Dim cSQL As String
Dim n As Integer
  
  vaArray.ReDim 0, -1, 0, 9
  cSQL = cSQL & " select t.kodeanggota,p.tgl,p.nomorpenjualan,p.kodestock,s.nama,p.harga,p.qty,p.kodesatuan,t.tunai,t.piutang from penjualan p"
  cSQL = cSQL & " left join totpenjualan t on t.nomorpenjualan = p.nomorpenjualan"
  cSQL = cSQL & " left join stock s on s.kodestock = p.kodestock"
  cSQL = cSQL & " where t.kodeanggota = '" & cKode.Text & "'  AND p.tgl >= '" & Format(dTgl(0).Value, "yyyy-MM-dd") & "' AND p.tgl <= '" & Format(dTgl(1).Value, "yyyy-MM-dd") & "' and p.statuslunas = '0'"
  cSQL = cSQL & " order by p.tgl"
  
  Set dbData = objData.Sql(GetDSN, cSQL)
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!kodeanggota)
      vaArray(n, 1) = Format(GetNull(dbData!tgl), "dd-MM-yyyy")
      vaArray(n, 2) = GetNull(dbData!nomorpenjualan)
      vaArray(n, 3) = GetNull(dbData!KodeStock)
      vaArray(n, 4) = GetNull(dbData!nama)
      vaArray(n, 5) = GetNull(dbData!Harga)
      vaArray(n, 6) = GetNull(dbData!qty)
      vaArray(n, 7) = GetNull(dbData!kodesatuan)
      If GetNull(dbData!Tunai) = 0 Then
        vaArray(n, 8) = 0
        vaArray(n, 9) = vaArray(n, 6) * vaArray(n, 5)
      Else
        vaArray(n, 9) = 0
        vaArray(n, 8) = vaArray(n, 6) * vaArray(n, 5)
      End If
      dbData.MoveNext
    Loop
  
    With FrmRPT
      .AddPageHeader "Detail Item penjualan", tdbHalignCenter, , , , , 10, True
      .AddPageHeader cNama.Text & " Dept. " & cAlamat.Text, tdbHalignCenter, , , True
      .AddPageHeader "Tgl " & Format(dTgl(0).Value, "dd-MM-yyyy") & " - " & Format(dTgl(1).Value, "dd-MM-yyyy"), tdbHalignCenter, , , True
      .AddPageHeader "", tdbHalignCenter, , , True
      .AddPageHeader "", tdbHalignCenter, , , True
      
      .AddTableHeader "", , , , , , , , , , , , , , , , , , , False
      .AddTableHeader "Tgl ", , , , 10
      .AddTableHeader "Inv", , , , 12
      .AddTableHeader "Kode", , , , 7
      .AddTableHeader "Nama"
      .AddTableHeader "Harga", , , , 11
      .AddTableHeader "Qty", , , , 10
      .AddTableHeader "Satuan", , , , 6
      .AddTableHeader "Cash", , , , 13
      .AddTableHeader "Charge", , , , 13
      
       
      .AddTableBody , , , , , , , , , , , , , False
      .AddTableBody
      .AddTableBody
      .AddTableBody
      .AddTableBody
      .AddTableBody Sis_Rpt_Number2
      .AddTableBody Sis_Rpt_Number
      .AddTableBody
      .AddTableBody Sis_Rpt_Number2
      .AddTableBody Sis_Rpt_Number2
      
      
      .AddTableFooter , , , , , , , , , , , , , , , , , , , False
      .AddTableFooter "Total", , tdbHalignRight, , , , , , , , , , , , 5
      .AddTableFooter
      .AddTableFooter
      .AddTableFooter
      .AddTableFooter
      .AddTableFooter "&Sum", Sis_Rpt_Number2
      .AddTableFooter
      .AddTableFooter "&Sum", Sis_Rpt_Number2
      .AddTableFooter "&Sum", Sis_Rpt_Number2
      
      .Preview vaArray, , True
    End With
    
  End If
End Sub

Private Sub GetDataAll()
Dim cSQL As String
Dim n As Integer
  
  vaArray.ReDim 0, -1, 0, 7
  
  cSQL = ""
  cSQL = cSQL & " select t.kodeanggota,a.nama as namaanggota,a.kodedep,p.tgl,p.kodestock,s.nama,p.harga,sum(p.qty) as qty,p.kodesatuan,sum(p.qty*p.tunai) as tunai,sum(p.qty*p.piutang) as piutang"
  cSQL = cSQL & " from penjualan p"
  cSQL = cSQL & " left join totpenjualan t on t.nomorpenjualan = p.nomorpenjualan"
  cSQL = cSQL & " left join stock s on s.kodestock = p.kodestock"
  cSQL = cSQL & " left join anggota a on a.kodeanggota = t.kodeanggota"
  cSQL = cSQL & " where p.tgl >= '" & Format(dTgl(0).Value, "yyyy-MM-dd") & "' AND p.tgl <= '" & Format(dTgl(1).Value, "yyyy-MM-dd") & "' and p.statuslunas = '0'"
  cSQL = cSQL & " group by t.kodeanggota,p.tgl,p.kodestock"
  cSQL = cSQL & " order by a.kodedep,t.kodeanggota,p.tgl"
  
  
  Set dbData = objData.Sql(GetDSN, cSQL)
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!kodeanggota)
      vaArray(n, 1) = GetNull(dbData!namaanggota) & " Dept. " & GetNull(dbData!kodedep)
      vaArray(n, 2) = Format(GetNull(dbData!tgl), "dd-MM-yyyy")
      vaArray(n, 3) = GetNull(dbData!KodeStock)
      vaArray(n, 4) = GetNull(dbData!nama)
      vaArray(n, 5) = GetNull(dbData!qty)
      vaArray(n, 6) = GetNull(dbData!Tunai)
      vaArray(n, 7) = GetNull(dbData!Piutang)
      If optOpsi(0).Value = True Then
        If vaArray(n, 6) = 0 Then
          vaArray.DeleteRows n
        End If
      ElseIf optOpsi(1).Value = True Then
        If vaArray(n, 7) = 0 Then
          vaArray.DeleteRows n
        End If
      End If
      dbData.MoveNext
    Loop
  
    With FrmRPT
      .AddPageHeader "Detail Item penjualan", tdbHalignCenter, , , , , 10, True
      .AddPageHeader "Tgl " & Format(dTgl(0).Value, "dd-MM-yyyy") & " - " & Format(dTgl(1).Value, "dd-MM-yyyy"), tdbHalignCenter, , , True, , , True
      .AddPageHeader aCfg(objData, msNamaPerusahaan), tdbHalignCenter, , , True, , 14, True
      .AddPageHeader "", tdbHalignCenter, , , True
      .AddPageHeader "", tdbHalignCenter, , , True
      
      
      .AddTableGroupHeader True, "[]", , , , 10
      .AddTableGroupHeader
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
        
      .AddTableHeader "", , , , , , , , , , , , , , , , , , , False
      .AddTableHeader "", , , , , , , , , , , , , , , , , , , False
      .AddTableHeader "Tgl ", , , , 10
      .AddTableHeader "Kode", , , , 7
      .AddTableHeader "Nama"
      .AddTableHeader "Qty", , , , 10
      .AddTableHeader "Cash", , , , 13
      .AddTableHeader "Charge", , , , 13
       
      .AddTableBody , , , , , , , , , , , , , False
      .AddTableBody , , , , , , , , , , , , , False
      .AddTableBody
      .AddTableBody
      .AddTableBody
      .AddTableBody Sis_Rpt_Number2
      .AddTableBody Sis_Rpt_Number2
      .AddTableBody Sis_Rpt_Number2
            
      
      .AddTableFooter , , , , , , , , , , , , , , , , , , , False
      .AddTableFooter , , , , , , , , , , , , , , , , , , , False
      .AddTableFooter "Total", , tdbHalignRight, , , , , , , , , , , , 3
      .AddTableFooter
      .AddTableFooter
      .AddTableFooter "&Sum", Sis_Rpt_Number2
      .AddTableFooter "&Sum", Sis_Rpt_Number2
      .AddTableFooter "&Sum", Sis_Rpt_Number2
      
      
      .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupFooter "SubTotal", , tdbHalignRight, , , , , , , , , , , , 3
      .AddTableGroupFooter
      .AddTableGroupFooter
      .AddTableGroupFooter "&Sum", Sis_Rpt_Number2
      .AddTableGroupFooter "&Sum", Sis_Rpt_Number2
      .AddTableGroupFooter "&Sum", Sis_Rpt_Number2
      
      .Preview vaArray, , True
    End With
    
  End If
End Sub

Private Sub cKode_ButtonClick()
  Set dbData = objData.PICK(GetDSN, "anggota", "kodeanggota", cKode, "kodeanggota,nama,kodedep")
  If Not dbData.EOF Then
    GetDataanggota
  End If
End Sub

Private Sub cmdPreview_Click()
  GetData
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPrintAll_Click()
  GetDataAll
End Sub

Private Sub cNama_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "anggota", "nama,kodedep,kodeanggota", "nama", sisContent, cNama.Text, , "nama")
  cNama.Text = cNama.Browse(dbData)
  If Not dbData.EOF Then
    GetDataanggota
  End If
End Sub

Private Sub GetDataanggota()
  cKode.Text = GetNull(dbData!kodeanggota, "")
  cNama.Text = GetNull(dbData!nama, "")
  cAlamat.Text = GetNull(dbData!kodedep, "")
End Sub

Private Sub Form_Load()
Dim n As Single

    SetIcon Me.hWnd, "SIKD"
    CenterForm Me
    InitValue
    
    TabIndex cKode, n
    TabIndex cNama, n
    TabIndex dTgl(0), n
    TabIndex dTgl(1), n
    TabIndex optOpsi(0), n
    TabIndex optOpsi(1), n
    TabIndex optOpsi(2), n
    TabIndex cmdPrintAll, n
    TabIndex cmdPreview, n
    TabIndex cmdKeluar, n
End Sub

Sub InitValue()
  cKode.Default
  cNama.Default
  dTgl(0).Value = BOM(Date)
  dTgl(1).Value = Date
  optOpsi(1).Value = True
End Sub

Private Sub optOpsi_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeysA vbKeyTab, True
  End If
End Sub
