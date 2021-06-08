VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptPenjualanKonsinyasi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laporan Penjualan Konsinyasi"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7005
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   7005
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   0
      Top             =   2805
      Width           =   6975
      _ExtentX        =   12303
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
         Left            =   5790
         TabIndex        =   1
         Top             =   90
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
         Picture         =   "rptPenjualanKonsinyasi.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   5355
         TabIndex        =   0
         Top             =   90
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
         Picture         =   "rptPenjualanKonsinyasi.frx":00A6
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   2730
      Left            =   0
      Top             =   75
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   4815
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
      Begin VB.OptionButton Option1 
         Caption         =   "H. Jual"
         Height          =   285
         Index           =   1
         Left            =   3135
         TabIndex        =   8
         Top             =   1485
         Width           =   930
      End
      Begin VB.OptionButton Option1 
         Caption         =   "H. Beli"
         Height          =   285
         Index           =   0
         Left            =   2250
         TabIndex        =   7
         Top             =   1485
         Width           =   930
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   195
         Left            =   2325
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   915
         Width           =   240
      End
      Begin BiSATextBoxProject.BiSABrowse cSupplier 
         Height          =   330
         Left            =   2565
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   915
         Width           =   1725
         _ExtentX        =   3043
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
      Begin BiSADateProject.BiSADate dDate 
         Height          =   330
         Index           =   0
         Left            =   225
         TabIndex        =   4
         Top             =   510
         Width           =   3600
         _ExtentX        =   6350
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
         Caption         =   "ANTARA TANGGAL"
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
      Begin BiSADateProject.BiSADate dDate 
         Height          =   330
         Index           =   1
         Left            =   3855
         TabIndex        =   5
         Top             =   510
         Width           =   1965
         _ExtentX        =   3466
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
         Caption         =   "S.D"
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
      Begin VB.Label Label1 
         Caption         =   "Supplier"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   270
         TabIndex        =   6
         Top             =   885
         Width           =   1575
      End
   End
End
Attribute VB_Name = "rptPenjualanKonsinyasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
  GetRpt
End Sub

Private Sub cSupplier_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "supplier", "kodesupplier,nama,alamat")
  If Not dbData.EOF Then
    cSupplier.Text = cSupplier.Browse(dbData)
  End If
End Sub

Private Sub Form_Load()
Dim n As Single
  
  CenterForm Me
  SetIcon Me.hWnd
  dDate(0).value = Date
  dDate(1).value = Date
  
  Option1(1).value = True
  
  TabIndex dDate(0), n
  TabIndex dDate(1), n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub

Private Sub GetRpt()
Dim n As Integer
Dim cSQL As String
Dim cWhere As String

  cWhere = ""
  cWhere = cWhere & " and p.tgl >= '" & Format(dDate(0).value, "yyyy-MM-dd") & "' and p.tgl <= '" & Format(dDate(1).value, "yyyy-MM-dd") & "'"
  If Check1.value = 1 Then
    cWhere = cWhere & " and s.kodesupplier = '" & cSupplier.Text & "'"
  End If
  
  cSQL = "select s.kodesupplier,sp.nama as supplier,p.tgl,p.nomorpenjualan,p.kodestock,s.nama,p.qty,p.harga,s.hargabeli,p.kodesatuan"
  cSQL = cSQL & " from penjualan p"
  cSQL = cSQL & " left join stock s on s.kodestock = p.kodestock"
  cSQL = cSQL & " left join supplier sp on sp.kodesupplier = s.kodesupplier"
  cSQL = cSQL & " where s.konsi = '1'" & cWhere
  cSQL = cSQL & " order by s.kodesupplier,p.tgl"
  

  
  vaArray.ReDim 0, -1, 0, 9
  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      
      vaArray(n, 0) = GetNull(dbData!kodesupplier)
      vaArray(n, 1) = GetNull(dbData!supplier)
      vaArray(n, 2) = Format(GetNull(dbData!tgl), "dd/MM/yy")
      vaArray(n, 3) = GetNull(dbData!nomorpenjualan)
      vaArray(n, 4) = GetNull(dbData!KodeStock)
      vaArray(n, 5) = GetNull(dbData!nama)
      vaArray(n, 6) = GetNull(dbData!qty)
      If Option1(1).value = True Then
        vaArray(n, 7) = GetNull(dbData!Harga)
      Else
        vaArray(n, 7) = GetNull(dbData!hargabeli)
      End If
      vaArray(n, 8) = GetNull(dbData!kodesatuan)
      vaArray(n, 9) = vaArray(n, 6) * vaArray(n, 7)
      dbData.MoveNext
    Loop
    
    With FrmRPT
      
      .AddPageHeader "Laporan Penjualan Konsinyasi", tdbHalignCenter, , , True, , 14, True, True, True, False, tdbPageHeaderSect, , , , , 3
      .AddPageHeader "Periode : " & Format(dDate(0).value, "dd-mm-yyyy") & " s.d " & Format(dDate(1).value, "dd-mm-yyyy"), tdbHalignCenter, , , True, , 10, True, True
  
      
      .AddTableGroupHeader True, "[]", , , , 10
      .AddTableGroupHeader
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      
      .AddTableHeader "", , , , , , , , , , , , , , , , , , , False
      .AddTableHeader "", , , , , , , , , , , , , , , , , , , False
      .AddTableHeader "Tgl", , , , 7
      .AddTableHeader "Nomor", , , , 13
      .AddTableHeader "Kode", , , , 7
      .AddTableHeader "Nama"
      .AddTableHeader "Qty", , , , 6
      If Option1(1).value = True Then
        .AddTableHeader "Harga", , , , 10
      Else
        .AddTableHeader "H.Beli", , , , 10
      End If
      .AddTableHeader "Satuan", , , , 6
      .AddTableHeader "Total", , , , 12
      
       
      .AddTableBody , , , , , , , , , , , , , False
      .AddTableBody , , , , , , , , , , , , , False
      .AddTableBody
      .AddTableBody
      .AddTableBody
      .AddTableBody
      .AddTableBody Sis_Rpt_Number, tdbHalignRight
      .AddTableBody Sis_Rpt_Number2, tdbHalignRight
      .AddTableBody
      .AddTableBody Sis_Rpt_Number2, tdbHalignRight
      
      .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupFooter "SubTotal", , tdbHalignRight, , , , , , , , , , , , 7
      .AddTableGroupFooter
      .AddTableGroupFooter
      .AddTableGroupFooter
      .AddTableGroupFooter
      .AddTableGroupFooter
      .AddTableGroupFooter
      .AddTableGroupFooter "&Sum", Sis_Rpt_Number2
      
      
      .AddTableFooter , , , , , , , , , , , , , , , , , , , False
      .AddTableFooter , , , , , , , , , , , , , , , , , , , False
      .AddTableFooter "TOTAL", , tdbHalignRight, , , , , , , , , , , , 7
      .AddTableFooter
      .AddTableFooter
      .AddTableFooter
      .AddTableFooter
      .AddTableFooter
      .AddTableFooter
      .AddTableFooter "&Sum", Sis_Rpt_Number2
      
      .Refresh
      .Preview vaArray, True
    End With
  Else
    MsgBox "Maaf tidak ada data", vbExclamation
  End If
    
End Sub

