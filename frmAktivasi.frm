VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Begin VB.Form frmAktivasi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "USER LEVEL"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2775
   ControlBox      =   0   'False
   Icon            =   "frmAktivasi.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   2775
   ShowInTaskbar   =   0   'False
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   630
      Left            =   60
      Top             =   1350
      Width           =   2670
      _ExtentX        =   4710
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
      BorderStyle     =   4
      BackColor       =   -2147483633
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   1365
         TabIndex        =   6
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
         Picture         =   "frmAktivasi.frx":030A
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   285
         TabIndex        =   7
         Top             =   90
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   767
         Caption         =   "    &Save"
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
         Picture         =   "frmAktivasi.frx":03B0
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1335
      Left            =   60
      Top             =   15
      Width           =   2670
      _ExtentX        =   4710
      _ExtentY        =   2355
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   4
      BackColor       =   -2147483633
      Begin VB.CheckBox optAdd 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1245
         TabIndex        =   3
         Top             =   480
         Width           =   945
      End
      Begin VB.CheckBox optEdit 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1245
         TabIndex        =   2
         Top             =   765
         Width           =   945
      End
      Begin VB.CheckBox optDel 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1245
         TabIndex        =   1
         Top             =   1035
         Width           =   945
      End
      Begin BiSANumberBoxProject.BiSANumberBox nLevel 
         Height          =   315
         Left            =   1230
         TabIndex        =   0
         Top             =   75
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   556
         Decimals        =   0
         Separator       =   ""
         MaxValue        =   999
         MinValue        =   0
         xxxx            =   999
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " "
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
      Begin VB.Label Label2 
         Caption         =   "Level"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   195
         TabIndex        =   5
         Top             =   105
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "Permission"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   195
         TabIndex        =   4
         Top             =   465
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmAktivasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objdata As New CodeSuiteLibrary.Data
Dim dbData As New ADODB.Recordset
Dim cFormName As String
Dim objMenu As New CodeSuiteLibrary.Menu

Sub Action(ByVal obj As Form)
  cFormName = obj.Name
  Me.Show vbModal
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdSimpan_Click()
  objdata.Update GetDSN, "FormLevel", "Nama = '" & cFormName & "' and UserLevel = " & nLevel.Value, Array("Nama", "UserLevel", "Status"), Array(cFormName, nLevel.Value, GetStatus())
  Unload Me
End Sub

Private Function GetStatus() As String
  GetStatus = Abs(optAdd.Value) & Abs(optEdit.Value) & Abs(optDel.Value)
End Function

Private Sub Form_Activate()
  'otorisasi hanya jika user level tidak sama dengan 0 atau root
  If GetRegistry(reg_UserLevel) <> 0 Then
    If objMenu.GetPassword("", Me, GetDSN) Then
      If objMenu.UserLevel <> 0 Then
        Me.Hide
        MsgBox "Maaf, Anda tidak diberikan wewenang untuk melakukan otorisasi." & vbCrLf & _
               "Hanya user dengan LEVEL 0 (SUPERVISOR) yg diijinkan", vbInformation, "OTORISASI not ALLOWED"
      End If
    Else
      Unload Me
    End If
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

    CenterForm Me
    TabIndex nLevel, n
    TabIndex optAdd, n
    TabIndex optEdit, n
    TabIndex optDel, n
    TabIndex cmdSimpan, n
    TabIndex cmdKeluar, n
End Sub

Private Sub nLevel_LostFocus()
Dim cStatus As String
  cStatus = GetFormLevel(cFormName, nLevel.Value)
  
  optAdd.Value = Val(Left(cStatus, 1))
  optEdit.Value = Val(Mid(cStatus, 2, 1))
  optDel.Value = Val(Mid(cStatus, 3, 1))
End Sub

Private Sub optAdd_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    keybd_event VK_TAB, 0, 0, 0
  End If
End Sub
