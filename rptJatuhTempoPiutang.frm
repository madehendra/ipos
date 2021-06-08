VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptJatuhTempoPiutang 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laporan Jatuh Tempo Piutang"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   5445
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
      Begin VB.CheckBox Check1 
         Caption         =   "Sekarang"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2160
         TabIndex        =   4
         Top             =   750
         Width           =   1170
      End
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
         Left            =   3630
         TabIndex        =   1
         Top             =   300
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
         TabIndex        =   2
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
         Picture         =   "rptJatuhTempoPiutang.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   3855
         TabIndex        =   3
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
         Picture         =   "rptJatuhTempoPiutang.frx":00A6
      End
   End
End
Attribute VB_Name = "rptJatuhTempoPiutang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub GetSQL()
Dim cSQL As String
Dim n As Integer

  cSQL = ""
  
  cSQL = cSQL & " select t.kodeanggota,a.nama,t.jthtmp,t.nomorpenjualan,t.piutang from totpenjualan t"
  cSQL = cSQL & " left join anggota a on a.kodeanggota = t.kodeanggota"
  If Check1.Value = 1 Then
    cSQL = cSQL & " where t.jthtmp <= '" & Format(Date, "yyyy-MM-dd") & "'"
  Else
    cSQL = cSQL & " where t.jthtmp >= '" & Format(dDate(0).Value, "yyyy-MM-dd") & "' and t.jthtmp <= '" & Format(dDate(1).Value, "yyyy-MM-dd") & "'"
  End If
  cSQL = cSQL & " order by t.kodeanggota,a.nama,t.jthtmp"
  
  vaArray.ReDim 0, -1, 0, 6
  Set dbData = objData.Sql(GetDSN, cSQL)
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!kodeanggota)
      vaArray(n, 1) = GetNull(dbData!nama) & " " & Format(GetSaldoPiutang(objData, vaArray(n, 0)), "###,###,##0.00")
      vaArray(n, 2) = Format(GetNull(dbData!jthtmp), "dd/MM/yyyy") & " (" & DateDiff("d", Date, Format(GetNull(dbData!jthtmp), "yyyy-MM-dd")) & ")"
      vaArray(n, 3) = GetNull(dbData!nomorpenjualan)
      vaArray(n, 4) = GetNull(dbData!Piutang)
      vaArray(n, 5) = GetLunasPiutang(objData, vaArray(n, 3))
      vaArray(n, 6) = vaArray(n, 4) - vaArray(n, 5)
      If vaArray(n, 6) <= 0 Then
        vaArray.DeleteRows n
      End If
      dbData.MoveNext
    Loop
    GetRpt
  End If
End Sub

Private Function GetSaldoPiutang(ByVal obj As CodeSuiteLibrary.Data, ByVal cKodeAnggota As String) As Double
Dim db As New ADODB.Recordset
  GetSaldoPiutang = 0
  Set db = obj.Browse(GetDSN, "kartupiutang", "sum(debet-kredit) as saldopiutang", "kodeanggota", sisAssign, cKodeAnggota, " and tgl <= '" & Format(Date, "yyyy-MM-dd") & "'")
  If Not db.EOF Then
    GetSaldoPiutang = GetNull(db!saldopiutang)
  End If
End Function

Private Function GetLunasPiutang(ByVal obj As CodeSuiteLibrary.Data, ByVal cNomorPenjualan) As Double
Dim db As New ADODB.Recordset

  GetLunasPiutang = 0
  Set db = obj.Browse(GetDSN, "pelunasanpiutang", "sum(pelunasan) as totalpelunasan", "nomorpenjualan", sisAssign, cNomorPenjualan)
  If Not db.EOF Then
    GetLunasPiutang = GetNull(db!totalpelunasan)
  End If
End Function

Private Sub GetRpt()
  With FrmRPT
      .AddPageHeader "Laporan Jatuh Tempo Piutang", tdbHalignCenter, , , True, dbArial, 10, True, False, , False
      .AddPageHeader aCfg(objData, msNamaPerusahaan, ""), tdbHalignCenter, , , True, , 12, True
      .AddPageHeader "Jatuh Tempo Tgl : " & Format(dDate(0).Value, "dd-MM-yyyy") & " sd " & Format(dDate(1).Value, "dd-MM-yyyy"), tdbHalignCenter, , , True, , , True
      .AddPageHeader " ", , , , True
      .AddPageHeader " ", , , , True
      
      .AddTableGroupHeader True, "[]", , , , 10
      .AddTableGroupHeader
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      
      
      .AddTableHeader "", , , , , , , , , , , , , , , , , , , False
      .AddTableHeader "", , , , , , , , , , , , , , , , , , , False
      .AddTableHeader "Jatuh Tempo", , , , 17
      .AddTableHeader "Nomor Penjualan"
      .AddTableHeader "Piutang", , , , 17
      .AddTableHeader "Lunas", , , , 17
      .AddTableHeader "Sisa", , , , 17


      .AddTableBody , , , , , , , , , , , , , False
      .AddTableBody , , , , , , , , , , , , , False
      .AddTableBody
      .AddTableBody
      .AddTableBody Sis_Rpt_Number2
      .AddTableBody Sis_Rpt_Number2
      .AddTableBody Sis_Rpt_Number2
           
           
      .AddTableFooter , , , , , , , , , , , , , , , , , , , False
      .AddTableFooter , , , , , , , , , , , , , , , , , , , False
      .AddTableFooter "GRAND TOTAL", , tdbHalignRight, , , , , , , , , , , , 1
      .AddTableFooter
      .AddTableFooter "&Sum", Sis_Rpt_Number2
      .AddTableFooter "&Sum", Sis_Rpt_Number2
      .AddTableFooter "&Sum", Sis_Rpt_Number2
    
    
      .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupFooter
      .AddTableGroupFooter
      .AddTableGroupFooter "&Sum", Sis_Rpt_Number2
      .AddTableGroupFooter "&Sum", Sis_Rpt_Number2
      .AddTableGroupFooter "&Sum", Sis_Rpt_Number2
    
    .Preview vaArray, True
  End With
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
  GetSQL
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd
  CenterForm Me
  TabIndex dDate(0), n
  TabIndex dDate(1), n
  TabIndex Check1, n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
  
End Sub
