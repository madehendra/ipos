VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptDaySalesConsignmentDetil 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detil Days Sales Consignment Detil"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   7530
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1830
      Left            =   15
      Top             =   0
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   3228
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
      Begin BiSATextBoxProject.BiSABrowse cNama 
         Height          =   330
         Left            =   330
         TabIndex        =   0
         Top             =   780
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
         Top             =   1155
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
         Top             =   405
         Width           =   3420
         _ExtentX        =   6033
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
         MaxLength       =   6
         Appearance      =   0
         Button          =   -1  'True
         Caption         =   "Kode Supplier"
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
         Top             =   1155
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
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   0
      Top             =   1815
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
         TabIndex        =   4
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
         Picture         =   "rptDaySalesConsignmentDetil.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   4380
         TabIndex        =   5
         Top             =   120
         Width           =   1965
         _ExtentX        =   3466
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
         Picture         =   "rptDaySalesConsignmentDetil.frx":00A6
      End
   End
End
Attribute VB_Name = "rptDaySalesConsignmentDetil"
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
  
  vaArray.ReDim 0, -1, 0, 7
  cSQL = cSQL & " select s.kodesupplier,p.tgl,p.kodestock,s.nama,p.qty,p.harga,p.tunai,p.piutang from penjualan p"
  cSQL = cSQL & " left join stock s on s.kodestock = p.kodestock"
  cSQL = cSQL & " Where s.jenis = 9 and s.kodesupplier = '" & cKode.Text & "'"
  cSQL = cSQL & " and p.tgl >= '" & Format(dTgl(0).Value, "yyyy-MM-dd") & "' and p.tgl <= '" & Format(dTgl(1).Value, "yyyy-MM-dd") & "'"
  cSQL = cSQL & " order by p.nomorpenjualan"
  
  Set dbData = objData.Sql(GetDSN, cSQL)
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!kodesupplier)
      vaArray(n, 1) = Format(GetNull(dbData!tgl), "dd-MM-yyyy")
      vaArray(n, 2) = GetNull(dbData!KodeStock)
      vaArray(n, 3) = GetNull(dbData!nama)
      vaArray(n, 4) = GetNull(dbData!qty)
      vaArray(n, 5) = GetNull(dbData!Tunai)
      vaArray(n, 6) = GetNull(dbData!Piutang)
      vaArray(n, 7) = (vaArray(n, 4) * vaArray(n, 5)) + (vaArray(n, 4) * vaArray(n, 6))
      dbData.MoveNext
    Loop
  
    With FrmRPT
      .AddPageHeader "Day Sales Consignment", tdbHalignCenter, , , , , 10, True
      .AddPageHeader cNama.Text, tdbHalignCenter, , , True, , 10, True
      .AddPageHeader "Tgl " & dTgl(0).Value & " - " & dTgl(1).Value, tdbHalignCenter, , , True
      .AddPageHeader "", tdbHalignCenter, , , True
      .AddPageHeader "", tdbHalignCenter, , , True
      
      .AddTableHeader "", , , , , , , , , , , , , , , , , , , False
      .AddTableHeader "Tgl ", , , , 10
      .AddTableHeader "Kode", , , , 7
      .AddTableHeader "Nama"
      .AddTableHeader "Qty", , , , 6
      .AddTableHeader "Cash", , , , 12
      .AddTableHeader "Credit", , , , 12
      .AddTableHeader "Jumlah", , , , 14
      
       
      .AddTableBody , , , , , , , , , , , , , False
      .AddTableBody
      .AddTableBody
      .AddTableBody
      .AddTableBody Sis_Rpt_Number2
      .AddTableBody Sis_Rpt_Number2
      .AddTableBody Sis_Rpt_Number2
      .AddTableBody Sis_Rpt_Number2
      
      
      .AddTableFooter , , , , , , , , , , , , , , , , , , , False
      .AddTableFooter "Total", , tdbHalignRight, , , , , , , , , , , , 3
      .AddTableFooter
      .AddTableFooter
      .AddTableFooter
      .AddTableFooter "&Sum", Sis_Rpt_Number2
      .AddTableFooter "&Sum", Sis_Rpt_Number2
      .AddTableFooter "&Sum", Sis_Rpt_Number2
      
      .Preview vaArray, , True
    End With
  End If
End Sub



Private Sub cKode_ButtonClick()
  Set dbData = objData.PICK(GetDSN, "supplier", "kodesupplier", cKode, "kodesupplier,nama")
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

Private Sub cNama_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "supplier", "nama,kodesupplier", "nama", sisContent, cNama.Text, , "nama")
  cNama.Text = cNama.Browse(dbData)
  If Not dbData.EOF Then
    GetDataanggota
  End If
End Sub

Private Sub GetDataanggota()
  cKode.Text = GetNull(dbData!kodesupplier, "")
  cNama.Text = GetNull(dbData!nama, "")
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
    TabIndex cmdPreview, n
    TabIndex cmdKeluar, n
End Sub

Sub InitValue()
  cKode.Default
  cNama.Default
  dTgl(0).Value = BOM(Date)
  dTgl(1).Value = Date
End Sub


