VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3360
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   5550
   LinkTopic       =   "Form2"
   ScaleHeight     =   3360
   ScaleWidth      =   5550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Insert ke table dep"
      Height          =   495
      Left            =   705
      TabIndex        =   1
      Top             =   1185
      Width           =   1470
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Insert ke table anggota"
      Height          =   495
      Left            =   660
      TabIndex        =   0
      Top             =   615
      Width           =   1470
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data

Private Sub Command1_Click()
Set dbData = objData.Browse(GetDSN, "customers1")
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      objData.Edit GetDSN, "anggota", "kodeanggota = '" & GetNull(dbData!cusid) & "'", Array("kodedep"), Array(GetNull(dbData!cusdept))
      dbData.MoveNext
    Loop
  End If
  MsgBox "Done"
End Sub

Private Sub Command2_Click()
Set dbData = objData.Browse(GetDSN, "customers1", "distinct(cusdept) as dep")
If Not dbData.EOF Then
  Do While Not dbData.EOF
    objData.Update GetDSN, "dep", "kodedep = '" & GetNull(dbData!dep) & "'", Array("kodedep"), Array(GetNull(dbData!dep))
    dbData.MoveNext
  Loop
  MsgBox "Done"
End If
End Sub
