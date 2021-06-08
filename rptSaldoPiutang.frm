VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptSaldoPiutang 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SALDO PIUTANG"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   345
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
      Begin BiSADateProject.BiSADate dDate 
         Height          =   330
         Index           =   0
         Left            =   345
         TabIndex        =   0
         Top             =   300
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
      Begin BiSADateProject.BiSADate dDate 
         Height          =   330
         Index           =   1
         Left            =   3645
         TabIndex        =   3
         Top             =   300
         Visible         =   0   'False
         Width           =   1470
         _ExtentX        =   2593
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
      Begin BiSATextBoxProject.BiSABrowse cGroupSales 
         Height          =   330
         Left            =   930
         TabIndex        =   5
         Top             =   690
         Width           =   2850
         _ExtentX        =   5027
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
         Caption         =   "Kategori"
         CaptionWidth    =   1100
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
      Begin BiSAButtonProject.BiSAButton BiSAButton1 
         Height          =   435
         Left            =   3345
         TabIndex        =   4
         Top             =   105
         Width           =   480
         _ExtentX        =   847
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
         BackColor       =   -2147483633
         Picture         =   "rptSaldoPiutang.frx":0000
      End
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
         Picture         =   "rptSaldoPiutang.frx":059A
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
         Picture         =   "rptSaldoPiutang.frx":0640
      End
   End
End
Attribute VB_Name = "rptSaldoPiutang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub BiSAButton1_Click()
  Load trSentSMS
  trSentSMS.Show vbModal
End Sub

Private Sub cGroupSales_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "groupsales", "kode,keterangan")
  If Not dbData.EOF Then
    If dbData.RecordCount > 1 Then
      cGroupSales.Text = cGroupSales.Browse(dbData)
    End If
    cGroupSales.Text = GetNull(dbData!Kode)
  Else
    cGroupSales.Default
  End If
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
Dim cFields As String
Dim cWhere As String
Dim n As Double
Dim nCol As Double


  
  cWhere = ""
  vaArray.ReDim 0, -1, 0, 4
  cFields = "s.kodeanggota,s.nama,s.alamat,Sum(h.debet) as Debet,Sum(h.kredit) as Kredit,s.kodedep,d.keterangan as namadep"
'  cWhere = cWhere & " h.tgl >= '" & Format(dDate(0).Value, "yyyy-MM-dd") & "' and h.tgl <= '" & Format(dDate(1).Value, "yyyy-MM-dd") & "'"
  cWhere = cWhere & " h.tgl <= '" & Format(dDate(0).Value, "yyyy-MM-dd") & "'"
  If Trim(cGroupSales.Text) <> "" Then
    cWhere = cWhere & " and h.groupsales='" & cGroupSales.Text & "'"
  End If
  cWhere = cWhere & " GROUP BY s.kodeanggota"
  Set dbData = objData.Browse(GetDSN, "anggota s", cFields, , , , cWhere, "s.kodedep", _
                              Array("LEFT JOIN kartupiutang h on h.kodeanggota = s.kodeanggota", _
                              "Left join dep d on d.kodedep = s.kodedep"))
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    dbData.MoveFirst
    Do While Not dbData.EOF
      FrmPB.RunPB
      vaArray.InsertRows (vaArray.UpperBound(1)) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!kodedep)
      vaArray(n, 1) = GetNull(dbData!namadep)
      vaArray(n, 2) = GetNull(dbData!kodeanggota)
      vaArray(n, 3) = GetNull(dbData!nama)
      vaArray(n, 4) = GetNull(dbData!debet) - GetNull(dbData!kredit)
'      vaArray(n, 5) = GetMemberTopUp(objData, vaArray(n, 2))
'      vaArray(n, 6) = vaArray(n, 5) - vaArray(n, 4)
      dbData.MoveNext
    Loop
    FrmPB.EndPB
  End If
  
  n = 0
  Do While n <= vaArray.UpperBound(1)
    If vaArray(n, 4) = 0 Then
      vaArray.DeleteRows n
      n = n - 1
    End If
    n = n + 1
  Loop
  
  vaArray.QuickSort 0, vaArray.UpperBound(1), 0, XORDER_DESCEND, XTYPE_STRING, 4, XORDER_DESCEND, XTYPE_DOUBLE
  
  
  
  GetRpt
  
  If MsgBox("Apakah laporan ini akan di export ke format Excel?", vbYesNo) = vbYes Then
    Dim a As New exportExcel
    vaArray.DeleteColumns (0)
    a.RecordSource = vaArray
    a.ExportToExcel
  End If
  
End Sub

Function GetMemberTopUp(ByVal obj As CodeSuiteLibrary.Data, cMemberKode As String) As Double
Dim dba As New ADODB.Recordset

  GetMemberTopUp = 0
  Set dba = obj.Browse(GetDSN, "membertopup", "sum(debet-kredit)as saldo", "kodeanggota", sisAssign, cMemberKode)
  If Not dba.EOF Then
    GetMemberTopUp = GetNull(dba!saldo)
  End If

End Function



Private Sub GetRpt()
'  vaArray.QuickSort 0, vaArray.UpperBound(1), 0, XORDER_ASCEND, XTYPE_DEFAULT, 2, XORDER_ASCEND, XTYPE_DEFAULT
  With FrmRPT
      .AddPageHeader "SALDO PIUTANG", tdbHalignCenter, , , True, dbArial, 14, True, True, , False
      .AddPageHeader aCfg(objData, msNamaPerusahaan, ""), tdbHalignCenter, , , True, , 12, True
      .AddPageHeader "Sampai Dengan Tanggal : " & Format(dDate(0).Value, "dd-MM-yyyy"), tdbHalignCenter, , , True, , , True
      .AddPageHeader cGroupSales.Text, tdbHalignCenter, , , True, , , True
      .AddPageHeader " ", , , , True

      
      .AddTableGroupHeader True, "[]", , , , 10 ',True, , , , 10, , , , , , , , , , , , , , , , True
      .AddTableGroupHeader
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
'      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
'      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      
      .AddTableHeader "", , , , , , , , , , , , , , , , , , , False
      .AddTableHeader "", , , , , , , , , , , , , , , , , , , False
      .AddTableHeader "KODE", , , , 15
      .AddTableHeader "NAMA"
      .AddTableHeader "TAGIHAN", , , , 16
'      .AddTableHeader "Saldo T.Up", , , , 12
'      .AddTableHeader "S. Akhir", , , , 15

      .AddTableBody , , , , , , , , , , , , , False
      .AddTableBody , , , , , , , , , , , , , False
      .AddTableBody
      .AddTableBody
      .AddTableBody Sis_Rpt_Number2
'      .AddTableBody Sis_Rpt_Number2
'      .AddTableBody Sis_Rpt_Number2
           
      .AddTableFooter , , , , , , , , , , , , , , , , , , , False
      .AddTableFooter , , , , , , , , , , , , , , , , , , , False
      .AddTableFooter "GRAND TOTAL", , tdbHalignRight, , , , , , , , , , , , 2
      .AddTableFooter
      .AddTableFooter "&Sum", Sis_Rpt_Number2
'      .AddTableFooter "&Sum", Sis_Rpt_Number2
'      .AddTableFooter "&Sum", Sis_Rpt_Number2
          
    
      .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupFooter "SubTotal", , tdbHalignRight, , , , , , , , , , , , 2
      .AddTableGroupFooter
      .AddTableGroupFooter "&Sum", Sis_Rpt_Number2
'      .AddTableGroupFooter "&Sum", Sis_Rpt_Number2
'      .AddTableGroupFooter "&Sum", Sis_Rpt_Number2
    
    .Preview vaArray, True
  End With
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  dDate(0).Value = Date
  TabIndex dDate(0), n
  TabIndex dDate(1), n
  TabIndex cGroupSales, n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub


