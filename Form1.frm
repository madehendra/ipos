VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2805
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   3120
   LinkTopic       =   "Form1"
   ScaleHeight     =   2805
   ScaleWidth      =   3120
   StartUpPosition =   3  'Windows Default
   Begin BiSAButtonProject.BiSAButton BiSAButton1 
      Height          =   450
      Left            =   1050
      TabIndex        =   0
      Top             =   810
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   794
      Caption         =   "OK"
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub BiSAButton1_Click()
Dim n As Integer

  n = 1
  BiSAButton1.Enabled = False
  BiSAButton1.Caption = "Processing"
'  Set dbData = objData.Browse(GetDSN, "customers1", , , , , , "cusname")
'  If Not dbData.EOF Then
'    Do While Not dbData.EOF
'      objData.Add GetDSN, "anggota", Array("kodeanggota", "kodeakun", "nama", "status", "kodedep"), Array(Padl(Trim(str(n)), 6, "0"), "1.200.01", dbData!cusname, "A", GetNull(dbData!cusdept))
'      n = n + 1
'      dbData.MoveNext
'    Loop
'  End If
'
'  Set dbData = objData.Browse(GetDSN, "inventory", , , , , , "nama")
'  If Not dbData.EOF Then
'    Do While Not dbData.EOF
'      objData.Add GetDSN, "stock", Array("kodesatuan", "kodegolongan", "nama", "hargabeli", "hargajual", "cogs"), Array("PCS", "100", GetNull(dbData!nama), GetNull(dbData!beli), GetNull(dbData!jual), GetNull(dbData!beli))
'      n = n + 1
'      dbData.MoveNext
'    Loop
'  End If
  
  Set dbData = objData.Browse(GetDSN, "sheet1")
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      objData.Add GetDSN, "anggota", Array("kodeanggota", "kodeakun", "nama", "status", "kodedep", "nopeg"), Array(GetNull(dbData!Kode, ""), "1.200.01", GetNull(dbData!nama, ""), "A", GetNull(dbData!dep), GetNull(dbData!nopeg))
      n = n + 1
      dbData.MoveNext
    Loop
  End If

  
  BiSAButton1.Enabled = True
  BiSAButton1.Caption = "OK"
End Sub
