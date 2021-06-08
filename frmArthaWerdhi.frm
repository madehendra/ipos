VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Begin VB.Form frmArthaWerdhi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Artha Werdhi"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4890
   StartUpPosition =   3  'Windows Default
   Begin BiSAButtonProject.BiSAButton BiSAButton2 
      Height          =   600
      Left            =   795
      TabIndex        =   1
      Top             =   990
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   1058
      Caption         =   "Hapus Brg yg 0 di kartu stock"
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
      Height          =   360
      Left            =   795
      TabIndex        =   0
      Top             =   555
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   635
      Caption         =   "Posting Kasir"
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
Attribute VB_Name = "frmArthaWerdhi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub BiSAButton1_Click()
Dim db As New ADODB.Recordset
Dim lSave As Boolean
  
  lSave = True
  objData.Start GetDSN
  
  Set dbData = objData.Browse(GetDSN, "kasir k", "t.tgl,k.nomorkasir,k.kodestock,k.jumlah,k.hargabeli,k.qty", , , , , , Array("LEFT JOIN totkasir t on t.nomorkasir = k.nomorkasir"))
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
'    lSave = IIf(lSave, objData.Delete(GetDSN, "bukubesar", "status", sisAssign, vbTrigger.msPenjualanKasir, " and kas = 'N' and (left(kodeakun,1) = '5' or left(kodeakun,1) = '1') and faktur = '" & GetNull(dbData!nomorkasir) & "'"), False)
    lSave = IIf(lSave, objData.Delete(GetDSN, "bukubesar", "status", sisAssign, vbTrigger.msPenjualanKasir, " and kas = 'N' and (left(kodeakun,1) = '5' or left(kodeakun,1) = '1')"), False)
    Do While Not dbData.EOF
      FrmPB.RunPB
      Set db = objData.Browse(GetDSN, "stock", "kodestock,asbiaya", "kodestock", sisAssign, GetNull(dbData!KodeStock))
      If Not db.EOF Then
        If GetNull(db!asbiaya) <> "1" Then
        
        'HP (5)
          'persediaan (1)
        
        lSave = IIf(lSave, UpdKodeTr(objData, msPenjualanKasir, GetNull(dbData!nomorkasir), Format(GetNull(dbData!tgl), "yyyy-MM-dd"), aCfg(objData, msRekeningCOGS), aCfg(objData, msCostCenterJualBeli), "Harga Pokok Penjualan Kasir No " & GetNull(dbData!nomorkasir), GetNull(dbData!qty) * GetNull(dbData!hargabeli), 0, "N", SNow), False)
            lSave = IIf(lSave, UpdKodeTr(objData, msPenjualanKasir, GetNull(dbData!nomorkasir), Format(GetNull(dbData!tgl), "yyyy-MM-dd"), GetAkunInventory(objData, GetNull(dbData!KodeStock)), aCfg(objData, msCostCenterJualBeli), "Harga Pokok Penjualan Kasir No " & GetNull(dbData!nomorkasir), 0, GetNull(dbData!qty) * GetNull(dbData!hargabeli), "N"), False)
            
        End If
      End If
      dbData.MoveNext
    Loop
    FrmPB.EndPB
  End If
  If lSave = True Then
    objData.Save GetDSN
  Else
    objData.Cancel GetDSN
  End If
End Sub

Private Sub BiSAButton2_Click()
Dim cSQL As String

cSQL = "select s.kodestock,s.nama from stock s"
cSQL = cSQL & " left join kartustock k on k.kodestock = s.kodestock"
cSQL = cSQL & " Where nama Is Null"

  Set dbData = objData.Sql(GetDSN, cSQL)
  FrmPB.InitPB dbData.RecordCount
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      FrmPB.RunPB
      objData.Delete GetDSN, "stock", "kodestock", sisAssign, GetNull(dbData!KodeStock)
      dbData.MoveNext
    Loop
  End If
  FrmPB.EndPB
  
cSQL = "select k.kodestock,s.nama from kartustock k"
cSQL = cSQL & " left join stock s on s.kodestock = k.kodestock"
cSQL = cSQL & " Where nama Is Null"

  Set dbData = objData.Sql(GetDSN, cSQL)
  FrmPB.InitPB dbData.RecordCount
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      FrmPB.RunPB
      objData.Delete GetDSN, "kartustock", "kodestock", sisAssign, GetNull(dbData!KodeStock)
      dbData.MoveNext
    Loop
  End If
  FrmPB.EndPB
  
  
cSQL = "select k.kodestock,s.nama,s.kodestock as stocksingade from stock s"
cSQL = cSQL & " left join kartustock k on k.kodestock = s.kodestock"
cSQL = cSQL & " Where k.KodeStock Is Null"

  Set dbData = objData.Sql(GetDSN, cSQL)
  FrmPB.InitPB dbData.RecordCount
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      FrmPB.RunPB
      objData.Delete GetDSN, "stock", "kodestock", sisAssign, GetNull(dbData!stocksingade)
      dbData.MoveNext
    Loop
  End If
  FrmPB.EndPB
  
End Sub
