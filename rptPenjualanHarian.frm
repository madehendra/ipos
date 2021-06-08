VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptPenjualanHarian 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PENJUALAN HARIAN..."
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   6990
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1125
      Left            =   15
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   1984
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
      Height          =   630
      Left            =   0
      Top             =   1110
      Width           =   6990
      _ExtentX        =   12330
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
         Left            =   5805
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
         Picture         =   "rptPenjualanHarian.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   5355
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
         Picture         =   "rptPenjualanHarian.frx":00A6
      End
   End
End
Attribute VB_Name = "rptPenjualanHarian"
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

Private Sub Form_Load()
Dim n As Single
  
  CenterForm Me
  SetIcon Me.hWnd
  'Initvalue
  dDate(0).Value = Date
  dDate(1).Value = Date
  
  TabIndex dDate(0), n
  TabIndex dDate(1), n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub

Private Sub GetRpt()
Dim n As Integer

  vaArray.ReDim 0, -1, 0, 3
  Set dbData = objData.Browse(GetDSN, "totpenjualan", "tgl,sum(tunai) as tunai, sum(piutang) as bon,sum(total) as total", "tgl", sisGTEqual, Format(dDate(0).Value, "yyyy-MM-dd"), " AND tgl <= '" & Format(dDate(1).Value, "yyyy-MM-dd") & "' GROUP BY tgl")
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = Format(GetNull(dbData!tgl), "dd/MM/yy")
      vaArray(n, 1) = GetNull(dbData!Tunai)
      vaArray(n, 2) = GetNull(dbData!Bon)
      vaArray(n, 3) = GetNull(dbData!Total)
      dbData.MoveNext
    Loop
  End If

  
  With FrmRPT
    .AddPageHeader "LAPORAN PENJUALAN HARIAN", tdbHalignCenter, , , , dbArial, 12, True
    .AddPageHeader aCfg(objData, msNamaPerusahaan, ""), tdbHalignCenter, , , True, dbArial, 12, True
    .AddPageHeader "Antara Tanggal : " & Format(dDate(0).Value, "dd/MM/yy") & " S/D " & Format(dDate(1).Value, "dd/MM/yy"), tdbHalignCenter, , , True, , 10
    .AddPageHeader " ", , , , True
    .AddPageHeader " ", , , , True
    
    .AddTableHeader "Tgl", , , , 15
    .AddTableHeader "Tunai", , , , 20
    .AddTableHeader "Bon", , , , 20
    .AddTableHeader "Total", , , , 20
    
    'isi Laporan
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    
    .AddTableFooter "TOTAL", , tdbHalignRight
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    
    .Preview vaArray, True
  End With
End Sub
