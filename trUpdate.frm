VERSION 5.00
Begin VB.Form trUpdate 
   Caption         =   "Form1"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Golongan"
      Height          =   465
      Left            =   285
      TabIndex        =   2
      Top             =   1620
      Width           =   915
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Supplier"
      Height          =   465
      Left            =   285
      TabIndex        =   1
      Top             =   960
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Stock"
      Height          =   465
      Left            =   300
      TabIndex        =   0
      Top             =   285
      Width           =   915
   End
End
Attribute VB_Name = "trUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub Command1_Click()
Dim vaField, vaValue
Dim lSave As Boolean

  lSave = True
  Set dbData = objData.Browse(GetDSN, "stock_lama", "nama,golongan,satuan,min,max,hb,hj")
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaField = Array("nama", "barcode", "kodegolongan", _
                      "kodesatuan", _
                      "hargabeli", "hargajual")
      vaValue = Array(dbData!nama, "", dbData!golongan, _
                      dbData!Satuan, _
                      dbData!hb, dbData!hj)
      lSave = IIf(lSave, objData.Add(GetDSN, "stock", vaField, vaValue), False)
      If lSave = True Then
        lSave = objData.Save(GetDSN)
      Else
        lSave = objData.Cancel(GetDSN)
      End If
      dbData.MoveNext
    Loop
    MsgBox "selesai"
  End If
End Sub

Private Sub Command2_Click()
Dim vaField, vaValue
Dim lSave As Boolean

  lSave = True
  Set dbData = objData.Browse(GetDSN, "supplier_lama", "kode,nama,alamat,kota")
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaField = Array("kodesupplier", "nama", "alamat", _
                      "kota")
      vaValue = Array(dbData!kode, dbData!nama, dbData!alamat, _
                      dbData!kota)
      lSave = IIf(lSave, objData.Add(GetDSN, "supplier", vaField, vaValue), False)
      If lSave = True Then
        lSave = objData.Save(GetDSN)
      Else
        lSave = objData.Cancel(GetDSN)
      End If
      dbData.MoveNext
    Loop
    MsgBox "selesai"
  End If
End Sub

Private Sub Command3_Click()
Dim vaField, vaValue
Dim lSave As Boolean

  lSave = True
  Set dbData = objData.Browse(GetDSN, "golongan_lama", "kode,keterangan")
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaField = Array("kodegolongan", "keterangan")
      vaValue = Array(dbData!kode, dbData!Keterangan)
      lSave = IIf(lSave, objData.Add(GetDSN, "golongan", vaField, vaValue), False)
      If lSave = True Then
        lSave = objData.Save(GetDSN)
      Else
        lSave = objData.Cancel(GetDSN)
      End If
      dbData.MoveNext
    Loop
    MsgBox "selesai"
  End If
End Sub
