VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptGrossProfit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gross Profit"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   7545
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1740
      Left            =   0
      Top             =   0
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   3069
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
         Caption         =   "Tampilkan Semua"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2430
         TabIndex        =   7
         Top             =   1155
         Width           =   2250
      End
      Begin BiSADateProject.BiSADate dTgl 
         Height          =   330
         Index           =   0
         Left            =   330
         TabIndex        =   0
         Top             =   345
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
      Begin BiSADateProject.BiSADate dTgl 
         Height          =   330
         Index           =   1
         Left            =   4185
         TabIndex        =   1
         Top             =   345
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
      Begin BiSATextBoxProject.BiSABrowse cNamaKasir 
         Height          =   330
         Left            =   3855
         TabIndex        =   2
         Top             =   765
         Width           =   2325
         _ExtentX        =   4101
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
      Begin BiSATextBoxProject.BiSATextBox cKodeKasir 
         Height          =   330
         Left            =   2430
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   765
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         Text            =   "12345678901234567890"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Verdana"
         BackColor       =   -2147483633
         Enabled         =   0   'False
         MaxLength       =   20
         Appearance      =   0
         CaptionWidth    =   1300
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
         Caption         =   "Kasir"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   435
         TabIndex        =   6
         Top             =   795
         Width           =   1005
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   0
      Top             =   1740
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
         Picture         =   "rptGrossProfit.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   5925
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
         Picture         =   "rptGrossProfit.frx":00A6
      End
   End
End
Attribute VB_Name = "rptGrossProfit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vaArray As New XArrayDB
Dim objData As New CodeSuiteLibrary.Data
Dim dbData As New ADODB.Recordset
Dim cUser As String

Private Sub GetData1()
Dim cSQL As String
Dim n As Integer

  
  vaArray.ReDim 0, -1, 0, 5
  
  cSQL = "select s.kodestock,s.nama,s.hargabeli,sum(k.qty) as tqty,sum(k.qty*(k.harga-(k.harga*k.disc/100))) as hjual,sum(k.qty*k.hp) as hpokok,sum(k.qty*(k.harga-(k.harga*k.disc/100)))- sum(k.qty*k.hp) as profit from kartustock k"
  cSQL = cSQL & " LEFT JOIN stock s on s.kodestock = k.kodestock"
  cSQL = cSQL & " where k.`status` = '60' and k.tgl >= '" & Format(dTgl(0).value, "yyyy-MM-dd") & "' and k.tgl <= '" & Format(dTgl(1).value, "yyyy-MM-dd") & "'"
  
  If Check1.value <> 1 Then
    cSQL = cSQL & " and k.username = '" & cKodeKasir.Text & "'"
  End If
  
  cSQL = cSQL & " GROUP BY s.kodestock"
  cSQL = cSQL & " ORDER BY profit desc"

  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = n + 1
      vaArray(n, 1) = GetNull(dbData!nama)
      vaArray(n, 2) = GetNull(dbData!tqty)
      vaArray(n, 3) = GetNull(dbData!hjual)
      vaArray(n, 4) = GetNull(dbData!hpokok)
      vaArray(n, 5) = GetNull(dbData!profit)
      dbData.MoveNext
    Loop
  
    With FrmRPT
      .AddPageHeader "Gross Profit", tdbHalignCenter, , , , , 10, True
      .AddPageHeader "Tgl " & Format(dTgl(0).value, "dd-MM-yyyy") & " - " & Format(dTgl(1).value, "dd-MM-yyyy"), tdbHalignCenter, , , True, , , True
      .AddPageHeader aCfg(objData, msNamaPerusahaan), tdbHalignCenter, , , True, , 14, True
      .AddPageHeader "", tdbHalignCenter, , , True
      .AddPageHeader "", tdbHalignCenter, , , True
      
      .AddTableHeader "No", , , , 5
      .AddTableHeader "Nama"
      .AddTableHeader "Tot Qty", , , , 8
      .AddTableHeader "Tot Sales", , , , 15
      .AddTableHeader "Tot H Pokok", , , , 15
      .AddTableHeader "Profit", , , , 17
            
      .AddTableBody
      .AddTableBody
      .AddTableBody Sis_Rpt_Number
      .AddTableBody Sis_Rpt_Number2
      .AddTableBody Sis_Rpt_Number2
      .AddTableBody Sis_Rpt_Number2
            
      .AddTableFooter "Total", , tdbHalignRight, , , , , , , , , , , , 3
      .AddTableFooter
      .AddTableFooter
      .AddTableFooter "&Sum", Sis_Rpt_Number2
      .AddTableFooter "&Sum", Sis_Rpt_Number2
      .AddTableFooter "&Sum", Sis_Rpt_Number2
      
      .Preview vaArray, , False
    End With
    
  End If
End Sub

Private Sub cmdPreview_Click()
  GetData1
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cNamaKasir_ButtonClick()
Dim db As New ADODB.Recordset

  Set db = objData.Browse(GetDSN, "username", "username,fullname", "(username", sisContent, cNamaKasir.Text, " or fullname like '%" & cNamaKasir.Text & "%')")
  If Not db.EOF Then
    cKodeKasir.Text = cNamaKasir.Browse(db)
    cKodeKasir.Text = GetNull(db!UserName)
    cNamaKasir.Text = GetNull(db!FullName)
    'cNamaRekening.Text = GetNull(db!keterangan, "")
  End If
End Sub

Private Sub dTgl_Validate(Index As Integer, Cancel As Boolean)
  If dTgl(1).value < dTgl(0).value Then
    dTgl(0).value = dTgl(1).value
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

    cNamaKasir.Enabled = True
    initvalue
    If GetRegistry(reg_UserLevel) <> 0 Then
      cNamaKasir.Enabled = False
      cKodeKasir.Text = GetRegistry(reg_Username)
      Check1.value = 0
      Check1.Enabled = False
    End If
    
    SetIcon Me.hWnd, "SIKD"
    CenterForm Me
    
    TabIndex dTgl(0), n
    TabIndex dTgl(1), n
    TabIndex cNamaKasir, n
    TabIndex cmdPreview, n
    TabIndex cmdKeluar, n
End Sub

Sub initvalue()
  dTgl(0).value = BOM(Date)
  dTgl(1).value = Date
  cKodeKasir.Text = ""
  cNamaKasir.Text = ""
  Check1.value = 1
  Check1.Enabled = True
End Sub


