VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptMutasiStock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mutasi Stock..."
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   6675
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1710
      Left            =   0
      Top             =   15
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   3016
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
      Begin VB.CheckBox chkGudang 
         Caption         =   "Seluruh Gudang"
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
         Left            =   1395
         TabIndex        =   0
         Top             =   765
         Width           =   1485
      End
      Begin BiSADateProject.BiSADate dDate 
         Height          =   330
         Index           =   0
         Left            =   270
         TabIndex        =   1
         Top             =   300
         Width           =   2430
         _ExtentX        =   4286
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
         Caption         =   "Mutasi"
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
         Left            =   2685
         TabIndex        =   2
         Top             =   300
         Width           =   1740
         _ExtentX        =   3069
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
         CaptionWidth    =   0
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
      Begin BiSATextBoxProject.BiSABrowse cGudang 
         Height          =   330
         Left            =   270
         TabIndex        =   3
         Top             =   1005
         Width           =   2430
         _ExtentX        =   4286
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
         Caption         =   "Gudang"
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
      Top             =   1710
      Width           =   6645
      _ExtentX        =   11721
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
         Left            =   5505
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
         Picture         =   "rptMutasiStock.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   5070
         TabIndex        =   5
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
         Picture         =   "rptMutasiStock.frx":00A6
      End
   End
End
Attribute VB_Name = "rptMutasiStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB


Private Sub cGudang_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "gudang", "kodegudang,keterangan")
  If Not dbData.EOF Then
    cGudang.Text = cGudang.Browse(dbData)
  End If
End Sub

Private Sub chkGudang_Click()
  If chkGudang.Value = 1 Then
    cGudang.Enabled = False
  Else
    cGudang.Enabled = True
  End If
End Sub

Private Sub chkGudang_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    keybd_event VK_TAB, 0, 0, 0
  End If
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
Dim n As Double
Dim nSaldo As Double
Dim cField  As String
Dim cWhere As String
Dim cSQLGudang As String

  vaArray.ReDim 0, -1, 0, 4
  cSQLGudang = ""
  If chkGudang.Value <> 1 Then
    cSQLGudang = " AND k.kodegudang = '" & cGudang.Text & "'"
  End If
  
  cField = "s.nama,sa.keterangan as satuan,k.tgl,k.keterangan,k.nomor,Sum(k.debet) as InStock, sum(k.kredit) as OutStock,k.kodestock"
  cWhere = cWhere & cSQLGudang
  cWhere = cWhere & " AND s.jenis = 1"
  cWhere = cWhere & " AND k.tgl <= '" & Format(dDate(1).Value, "yyyy-MM-dd") & "' AND k.tgl >= '" & Format(dDate(0).Value, "yyyy-MM-dd") & "' GROUP BY k.kodestock"
  Set dbData = objData.Browse(GetDSN, "kartustock k", cField, , , , "1=1" & cWhere, "kodestock,tgl,id", Array(" LEFT JOIN stock s ON s.kodestock = k.kodestock", "LEFT JOIN satuan sa on sa.kodesatuan = s.kodesatuan"))
  If dbData.RecordCount > 0 Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!KodeStock)
      vaArray(n, 1) = GetNull(dbData!nama)
      vaArray(n, 2) = GetNull(dbData!InStock)
      vaArray(n, 3) = GetNull(dbData!OutStock)
      vaArray(n, 4) = GetNull(dbData!Satuan)
      If vaArray(n, 2) = 0 And vaArray(n, 3) = 0 Then
        vaArray.DeleteRows n
      End If
      dbData.MoveNext
    Loop
  End If
              
  With FrmRPT
   .AddPageHeader "Mutasi Stock", tdbHalignCenter, , , , , 12, True, True
   .AddPageHeader "", , , , , , , , , True
   .AddPageHeader "", , , , , , , , , True
   .AddPageHeader "ANTARA TGL", , , 15, True
   .AddPageHeader ": " & Format(dDate(0).Value, "dd/MM/yyyy") & " s.d " & Format(dDate(1).Value, "dd/MM/yyyy"), , , , , , , , , , , , , , , , 5
   If chkGudang.Value <> 1 Then
    Set dbData = objData.Browse(GetDSN, "gudang", , "kodegudang", sisAssign, cGudang.Text)
    If Not dbData.EOF Then
      .AddPageHeader "DI " & GetNull(dbData!Keterangan), , , , True
    End If
   End If
   
   .AddTableHeader "SKU", , , , 12
   .AddTableHeader "Stock"
   .AddTableHeader "In", , , , 13
   .AddTableHeader "Out", , , , 13
   .AddTableHeader "Satuan", , , , 8
   
   .AddTableBody
   .AddTableBody
   .AddTableBody Sis_Rpt_Number2
   .AddTableBody Sis_Rpt_Number2
   .AddTableBody
   
   .AddTableFooter "Total", , tdbHalignRight, , , , , , , , , , , , 2
   .AddTableFooter
   .AddTableFooter "&Sum", Sis_Rpt_Number2
   .AddTableFooter "&Sum", Sis_Rpt_Number2
   .AddTableFooter
   
   .Preview vaArray
  End With
End Sub

Private Sub Form_Load()
Dim n As Single
  
  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  dDate(0).Value = BOM(Date)
  dDate(1).Value = Date
  cGudang.Default
  chkGudang.Value = 1
  cGudang.Enabled = False
  TabIndex dDate(0), n
  TabIndex dDate(1), n
  TabIndex chkGudang, n
  TabIndex cGudang, n
  TabIndex cmdPreview, n
End Sub
