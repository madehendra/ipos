VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Begin VB.Form frmtest 
   Caption         =   "Form1"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin BiSAButtonProject.BiSAButton BiSAButton5 
      Height          =   540
      Left            =   2715
      TabIndex        =   4
      Top             =   675
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   953
      Caption         =   "5"
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
   End
   Begin BiSAButtonProject.BiSAButton BiSAButton4 
      Height          =   450
      Left            =   2880
      TabIndex        =   3
      Top             =   2400
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   794
      Caption         =   "Label1"
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
   End
   Begin BiSAButtonProject.BiSAButton BiSAButton3 
      Height          =   390
      Left            =   330
      TabIndex        =   2
      Top             =   2235
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   688
      Caption         =   "3"
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
   End
   Begin BiSAButtonProject.BiSAButton BiSAButton2 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   1620
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   503
      Caption         =   "Label1"
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
   End
   Begin BiSAButtonProject.BiSAButton BiSAButton1 
      Height          =   900
      Left            =   135
      TabIndex        =   0
      Top             =   465
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   1588
      Caption         =   "Label1"
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
   End
End
Attribute VB_Name = "frmtest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub BiSAButton1_Click()
  Set dbData = objData.Browse(GetDSN, "gol")
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      objData.Edit GetDSN, "gol", "kodegolongan = '" & GetNull(dbData!kodegolongan) & "'", Array("kodegolongan"), Array(Padl(GetNull(dbData!kodegolongan), 3, "0"))
      dbData.MoveNext
    Loop
  End If
End Sub

Private Sub BiSAButton2_Click()
Set dbData = objData.Browse(GetDSN, "satuan")
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      objData.Edit GetDSN, "stock_old", "kodesatuan = '" & GetNull(dbData!keterangan) & "'", Array("kodesatuan"), Array(GetNull(dbData!kodesatuan))
      dbData.MoveNext
    Loop
  End If
  MsgBox "OK"
End Sub

Private Sub BiSAButton3_Click()
Set dbData = objData.Browse(GetDSN, "stock", "distinct(kodesatuan) as kodesatuan")
If Not dbData.EOF Then
  Do While Not dbData.EOF
    objData.Add GetDSN, "satuan", Array("kodesatuan", "keterangan"), Array(GetNull(dbData!kodesatuan), GetNull(dbData!kodesatuan))
    dbData.MoveNext
  Loop
  MsgBox "OK"
End If
End Sub

Private Sub BiSAButton4_Click()
Set dbData = objData.Browse(GetDSN, "golongan")
If Not dbData.EOF Then
  Do While Not dbData.EOF
    objData.Edit GetDSN, "golongan", "kodegolongan = '" & GetNull(dbData!kodegolongan) & "'", Array("kodegolongan"), Array(Padl(GetNull(dbData!kodegolongan), 3, "0"))
    dbData.MoveNext
  Loop
  MsgBox "OK"
End If
End Sub

Private Sub BiSAButton5_Click()
Set dbData = objData.Browse(GetDSN, "stock")
If Not dbData.EOF Then
  FrmPB.InitPB dbData.RecordCount
  Do While Not dbData.EOF
    FrmPB.RunPB
    UpdKartuStock objData, SaldoAwal, "XAWAL", Date, GetNull(dbData!KodeStock), GetNull(dbData!saldostock), GetNull(dbData!hargabeli), 0, "saldo awal", "GD"
    dbData.MoveNext
  Loop
  FrmPB.EndPB
  MsgBox "ok"
End If
End Sub
