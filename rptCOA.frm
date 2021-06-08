VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Begin VB.Form rptCOA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "COA..."
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7215
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   7215
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   720
      Left            =   0
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   1270
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
      Begin BiSATextBoxProject.BiSABrowse cRekening 
         Height          =   330
         Index           =   0
         Left            =   255
         TabIndex        =   0
         Top             =   180
         Width           =   3990
         _ExtentX        =   7038
         _ExtentY        =   582
         Text            =   "123"
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
         Caption         =   "ANTARA REKENING"
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
      Begin BiSATextBoxProject.BiSABrowse cRekening 
         Height          =   330
         Index           =   1
         Left            =   4305
         TabIndex        =   1
         Top             =   180
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   582
         Text            =   "123"
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
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   630
      Left            =   0
      Top             =   705
      Width           =   7215
      _ExtentX        =   12726
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
         Left            =   5970
         TabIndex        =   2
         Top             =   90
         Width           =   1140
         _ExtentX        =   2011
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
         Picture         =   "rptCOA.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   4800
         TabIndex        =   3
         Top             =   90
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   767
         Caption         =   "     &Preview"
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
         Picture         =   "rptCOA.frx":00A6
      End
   End
End
Attribute VB_Name = "rptCOA"
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
  getSQL
End Sub

Private Sub getSQL()
Dim cField As String
Dim n As Integer

  vaArray.ReDim 0, -1, 0, 2
  cField = "kodeakun,ConCat(Space(Length(kodeakun)-5),Keterangan) as Keterangan,Jenis"
  Set dbData = objData.Browse(GetDSN, "akun", cField, "kodeakun", sisGTEqual, cRekening(0).Text, " and kodeakun <= '" & cRekening(1).Text & "'", "kodeakun")
  If Not dbData.EOF Then
    dbData.MoveFirst
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!kodeakun, "")
      vaArray(n, 1) = GetNull(dbData!keterangan, "")
      vaArray(n, 2) = GetNull(dbData!jenis, "")
      dbData.MoveNext
    Loop
    GetRpt
  End If
End Sub

Private Sub GetRpt()
  With FrmRPT
    .AddPageHeader "Chart Of Account", tdbHalignCenter, , , , dbArial, 10, True
    .AddPageHeader aCfg(objData, msNamaPerusahaan), tdbHalignCenter, , , True, dbArial, 12, True
    .AddPageHeader "", , , , True
    .AddPageHeader "", , , , True
    
    .AddTableHeader "Kode", , , , 15
    .AddTableHeader "Account"
    .AddTableHeader "Jenis", , , , 5
    
    .AddTableBody
    .AddTableBody
    .AddTableBody , tdbHalignCenter
    
    .Preview vaArray, True
  End With
End Sub

Private Sub cRekening_ButtonClick(Index As Integer)
  Set dbData = objData.Browse(GetDSN, "akun", "kodeakun,Keterangan", "kodeakun", sisContent, cRekening(Index).Text)
  cRekening(Index).Text = cRekening(Index).Browse(dbData)
End Sub

Private Sub cRekening_Validate(Index As Integer, Cancel As Boolean)
  cRekening_ButtonClick (Index)
End Sub

Private Sub Form_Load()
Dim n As Single

  GetMinMax "akun", cRekening
  CenterForm Me
  SetIcon Me.hWnd, "SIKD"
  
  TabIndex cRekening(0), n
  TabIndex cRekening(1), n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub

