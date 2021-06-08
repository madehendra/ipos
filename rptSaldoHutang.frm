VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptSaldoHutang 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SALDO HUTANG"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   7965
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1245
      Left            =   -15
      Top             =   0
      Width           =   7980
      _ExtentX        =   14076
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
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   -15
      Top             =   1230
      Width           =   7980
      _ExtentX        =   14076
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
         Left            =   6780
         TabIndex        =   1
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
         Picture         =   "rptSaldoHutang.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   6345
         TabIndex        =   2
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
         Picture         =   "rptSaldoHutang.frx":00A6
      End
   End
End
Attribute VB_Name = "rptSaldoHutang"
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
Dim cFields As String
Dim cWhere As String
Dim n As Double
Dim nCol As Double

  
  vaArray.ReDim 0, -1, 0, 4
  cFields = "s.kodesupplier,s.kota,s.nama,s.alamat,Sum(h.debet) as Debet,Sum(h.kredit) as Kredit"
  cWhere = cWhere & " h.tgl <= '" & Format(dDate.Value, "yyyy-MM-dd") & "'"
  cWhere = cWhere & " GROUP BY s.kodesupplier"
  Set dbData = objData.Browse(GetDSN, "supplier s", cFields, , , , cWhere, "s.kodesupplier", _
                              Array("LEFT JOIN kartuhutang h on h.kodesupplier = s.kodesupplier"))
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    dbData.MoveFirst
    Do While Not dbData.EOF
      FrmPB.RunPB
      vaArray.InsertRows (vaArray.UpperBound(1)) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!kodesupplier)
      vaArray(n, 1) = GetNull(dbData!nama)
      vaArray(n, 2) = GetNull(dbData!alamat)
      vaArray(n, 3) = GetNull(dbData!kota)
      vaArray(n, 4) = GetNull(dbData!debet) - GetNull(dbData!kredit)
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
  
  GetRpt
End Sub

Private Sub GetRpt()
  vaArray.QuickSort 0, vaArray.UpperBound(1), 0, XORDER_ASCEND, XTYPE_DEFAULT, 2, XORDER_ASCEND, XTYPE_DEFAULT
  With FrmRPT
      .AddPageHeader "SALDO HUTANG", tdbHalignCenter, , , True, dbArial, 14, True, True, , False
      .AddPageHeader aCfg(objData, msNamaPerusahaan, ""), tdbHalignCenter, , , True, , 12, True
      .AddPageHeader "Sampai Tanggal : " & Format(dDate.Value, "dd-MM-yyyy"), tdbHalignCenter, , , True, , , True
      .AddPageHeader " ", , , , True
      .AddPageHeader " ", , , , True
      
      
      .AddTableHeader "KODE", , , , 7
      .AddTableHeader "NAMA", , , , 25
      .AddTableHeader "ALAMAT"
      .AddTableHeader "KOTA", , , , 18
      .AddTableHeader "HUTANG", , , , 17

      .AddTableBody
      .AddTableBody
      .AddTableBody
      .AddTableBody
      .AddTableBody Sis_Rpt_Number2
           
     .AddTableFooter "GRAND TOTAL", , tdbHalignRight, , , , , , , , , , , , 4
     .AddTableFooter
     .AddTableFooter
     .AddTableFooter
     .AddTableFooter "&Sum", Sis_Rpt_Number2
          
    .Preview vaArray, True
  End With
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
