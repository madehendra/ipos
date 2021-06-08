VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Begin VB.Form frmUpdateData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   6405
   Begin BiSAButtonProject.BiSAButton BiSAButton3 
      Height          =   630
      Left            =   2985
      TabIndex        =   2
      Top             =   2235
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   1111
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
   Begin BiSAButtonProject.BiSAButton BiSAButton2 
      Height          =   405
      Left            =   735
      TabIndex        =   1
      Top             =   2085
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   714
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
      Height          =   315
      Left            =   465
      TabIndex        =   0
      Top             =   120
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
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
Attribute VB_Name = "frmUpdateData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub BiSAButton1_Click()
  Set dbData = objData.Browse(GetDSN, "penjualan")
  FrmPB.InitPB dbData.RecordCount
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      FrmPB.RunPB
      objData.Edit GetDSN, "penjualan", "kodestock = '" & dbData!KodeStock & "'", Array("hb"), Array(GetHargaBeli(objData, dbData!KodeStock))
      dbData.MoveNext
    Loop
  End If
  FrmPB.EndPB
  MsgBox "Selesai"
End Sub

Private Sub BiSAButton2_Click()
Dim cSQL As String

  cSQL = "select p.nomorpenjualan,p.tgl,p.kodestock,p.qty,p.harga,t.tunai,t.piutang from penjualan p"
  cSQL = cSQL & " left join totpenjualan t on t.nomorpenjualan = p.nomorpenjualan"
  
  Set dbData = objData.Sql(GetDSN, cSQL)
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.EOF
      FrmPB.RunPB
      If GetNull(dbData!Tunai) <> 0 Then
        objData.Edit GetDSN, "penjualan", "nomorpenjualan = '" & GetNull(dbData!nomorpenjualan) & "' and kodestock = '" & GetNull(dbData!KodeStock) & "'", Array("tunai", "piutang"), Array(dbData!Harga, 0)
      Else
        objData.Edit GetDSN, "penjualan", "nomorpenjualan = '" & GetNull(dbData!nomorpenjualan) & "' and kodestock = '" & GetNull(dbData!KodeStock) & "'", Array("piutang", "tunai"), Array(dbData!Harga, 0)
      End If
      dbData.MoveNext
    Loop
    FrmPB.EndPB
  End If
End Sub

Private Sub BiSAButton3_Click()
  Set dbData = objData.Browse(GetDSN, "pelunasanpiutang")
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.EOF
      FrmPB.RunPB
      objData.Edit GetDSN, "penjualan", "nomorpenjualan = '" & GetNull(dbData!nomorpenjualan) & "'", Array("statuslunas"), Array(1)
      dbData.MoveNext
    Loop
    FrmPB.EndPB
  End If
End Sub
