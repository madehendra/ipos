VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptOmzetSalesman 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Omzet Salesman"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   6990
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   930
      Left            =   0
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   1640
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
         Left            =   390
         TabIndex        =   0
         Top             =   300
         Width           =   3495
         _ExtentX        =   6165
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
         Caption         =   "ANTARA  TANGGAL"
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
         Left            =   4005
         TabIndex        =   1
         Top             =   315
         Width           =   2010
         _ExtentX        =   3545
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
      Height          =   645
      Left            =   0
      Top             =   915
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
         Left            =   5775
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
         Picture         =   "rptOmzetSalesman.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   5340
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
         Picture         =   "rptOmzetSalesman.frx":00A6
      End
   End
End
Attribute VB_Name = "rptOmzetSalesman"
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
Dim cSQL As String
Dim n As Integer

  cSQL = cSQL & " select t.kodesalesman,s.nama,sum(t.total) as omzet, sum(t.piutang) as piutang,sum(t.tunai) as tunai from totpenjualan t"
  cSQL = cSQL & " left join salesman s on s.kodesalesman = t.kodesalesman"
  cSQL = cSQL & " where t.tgl >= '" & Format(dDate(0).Value, "yyyy-MM-dd") & "' and t.tgl <= '" & Format(dDate(1).Value, "yyyy-MM-dd") & "'"
  cSQL = cSQL & " group by t.kodesalesman;"
  
  vaArray.ReDim 0, -1, 0, 4
  Set dbData = objData.Sql(GetDSN, cSQL)
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!kodesalesman)
      vaArray(n, 1) = GetNull(dbData!nama)
      vaArray(n, 2) = GetNull(dbData!omzet)
      vaArray(n, 3) = GetNull(dbData!Piutang)
      vaArray(n, 4) = GetNull(dbData!Tunai)
      dbData.MoveNext
    Loop
    With FrmRPT
      .AddPageHeader "Omzet Salesman", tdbHalignCenter, , , True, dbArial, 12, True, , , False
      .AddPageHeader "Tgl " & Format(dDate(0).Value, "dd-MM-yyyy") & " sd " & Format(dDate(1).Value, "dd-MM-yyyy"), tdbHalignCenter, , , True, , 14
      .AddPageHeader aCfg(objData, msNamaPerusahaan), tdbHalignCenter, , , True, , 14
      .AddPageHeader "", , , , True
      .AddPageHeader "", , , , True
      
      .AddTableHeader "KODE", , , , 7
      .AddTableHeader "NAMA"
      .AddTableHeader "OMZET", , , , 20
      .AddTableHeader "CREDIT", , , , 20
      .AddTableHeader "TUNAI", , , , 20
      
      .AddTableBody
      .AddTableBody
      .AddTableBody Sis_Rpt_Number2, tdbHalignRight
      .AddTableBody Sis_Rpt_Number2, tdbHalignRight
      .AddTableBody Sis_Rpt_Number2, tdbHalignRight
            
      .AddTableFooter "Total", , tdbHalignRight, , , , , , , , , , , , 2
      .AddTableFooter
      .AddTableFooter "&Sum", Sis_Rpt_Number2
      .AddTableFooter "&Sum", Sis_Rpt_Number2
      .AddTableFooter "&Sum", Sis_Rpt_Number2
            
      .Preview vaArray, True
    End With

  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd
  CenterForm Me
  dDate(0).Value = BOM(Date)
  dDate(1).Value = EOM(Date)
  TabIndex dDate(0), n
  TabIndex dDate(1), n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub
