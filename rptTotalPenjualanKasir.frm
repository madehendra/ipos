VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptTotalPenjualanKasir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TOTAL PENJUALAN KASIR"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   7020
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   945
      Left            =   0
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   1667
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
         Left            =   225
         TabIndex        =   0
         Top             =   300
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   582
         Value           =   "15-11-2003"
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
         TabIndex        =   1
         Top             =   300
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   582
         Value           =   "15-11-2003"
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
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   645
      Left            =   0
      Top             =   930
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   1138
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
         Left            =   5685
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
         Picture         =   "rptTotalPenjualanKasir.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   5250
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
         Picture         =   "rptTotalPenjualanKasir.frx":00A6
      End
   End
End
Attribute VB_Name = "rptTotalPenjualanKasir"
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
  GetData
End Sub

Private Sub GetData()
Dim n As Double
Dim cWhere As String
Dim cFields As String

  vaArray.ReDim 0, -1, 0, 2
  cWhere = " AND tgl <= '" & Format(dDate(1).Value, "yyyy-MM-dd") & "'"
  cFields = "tgl,nomorkasir,total"
  Set dbData = objData.Browse(GetDSN, "totkasir", cFields, "tgl", sisGTEqual, Format(dDate(0).Value, "yyyy-MM-dd"), cWhere, "nomorkasir")
  If Not dbData.EOF Then
    dbData.MoveFirst
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.EOF
      FrmPB.RunPB
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = (dbData!nomorkasir)
      vaArray(n, 1) = Format(dbData!tgl, "dd-MM-yyyy")
      vaArray(n, 2) = (dbData!Total)
      dbData.MoveNext
    Loop
    FrmPB.EndPB
    GetRpt
  End If
End Sub

Private Sub GetRpt()
  With FrmRPT
    
    .AddPageHeader "LAPORAN PENJUALAN KASIR", tdbHalignCenter, , , True, , 14
    .AddPageHeader "Antara Tanggal : " & Format(dDate(0).Value, "dd-mm-yyyy") & " s.d " & Format(dDate(1).Value, "dd-mm-yyyy"), tdbHalignCenter, , , True, , 10
    .AddPageHeader " ", , , , True
    .AddPageHeader " ", , , , True
    
    .AddTableHeader "NOMOR", , , , 20, , , , , , , , , , , False, 7
    .AddTableHeader "TANGGAL", , , , 20
    .AddTableHeader "TOTAL", , , , 20
    
    .AddTableBody
    .AddTableBody , tdbHalignCenter
    .AddTableBody Sis_Rpt_Number2
    
    .AddTableFooter "TOTAL", , tdbHalignRight, , , , , , , , , , , , 2
    .AddTableFooter
    .AddTableFooter "&sum", Sis_Rpt_Number2, tdbHalignRight
    
    .Refresh
    .Preview vaArray, True
  End With
End Sub


Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  dDate(0).Value = Date
  dDate(1).Value = Date
  
  TabIndex dDate(0), n
  TabIndex dDate(1), n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub


