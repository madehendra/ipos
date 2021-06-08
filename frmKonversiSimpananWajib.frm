VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Begin VB.Form frmKonversiSimpananWajib 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Konversi Simpanan Wajib"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4815
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   4815
   Begin BiSAButtonProject.BiSAButton BiSAButton5 
      Height          =   510
      Left            =   1905
      TabIndex        =   6
      Top             =   2265
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   900
      Caption         =   "konvers2"
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
      Height          =   480
      Left            =   3210
      TabIndex        =   5
      Top             =   1485
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   847
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
      Height          =   420
      Left            =   750
      TabIndex        =   4
      Top             =   1455
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   741
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
      Height          =   420
      Left            =   2505
      TabIndex        =   3
      Top             =   855
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   741
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
      Height          =   405
      Left            =   195
      TabIndex        =   2
      Top             =   885
      Width           =   2085
      _ExtentX        =   3678
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
   Begin BiSAButtonProject.BiSAButton cmd1 
      Height          =   480
      Left            =   195
      TabIndex        =   0
      Top             =   270
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   847
      Caption         =   "Konv 1"
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
   Begin BiSAButtonProject.BiSAButton cmd2 
      Height          =   480
      Left            =   2190
      TabIndex        =   1
      Top             =   270
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   847
      Caption         =   "Konv2"
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
Attribute VB_Name = "frmKonversiSimpananWajib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data

Private Sub Konversi(ByVal cNo As String)
Dim lSave As Boolean
Dim Faktur As String
Dim nValue As Double
Dim nBulan As Double
Dim n As Integer
Dim nTmp
Dim dInit As Date
Dim nSimpananWajib As Double

'Mengkonversi simpanan wajib bulan 1
'Pilih hanya anggota yg sudah memiliki kode
  lSave = True
  Set dbData = objData.Browse(GetDSN, "sheet2", , "kode", sisDifference, "", " AND kode IS NOT NULL")
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.EOF
      FrmPB.RunPB "Tunggu sebentar..."
      'nomorsimpananwajib
      'kodeanggota
      'username
      'tgl
      'datetime
      'jumlah
      
      'nomorsimpananwajib
      'tahun
      'bulan
      'jumlah
        
        If cNo = "1" Then
          nValue = GetNull(dbData!sepuluh)
          dInit = "2006-08-01"
          nSimpananWajib = 10000
        ElseIf cNo = "2" Then
          nValue = GetNull(dbData!duapuluh)
          dInit = "2006-11-01"
          nSimpananWajib = 20000
        End If
        nBulan = 8
        For n = 1 To 3
          nTmp = 1
          Do While nValue > 0
            Faktur = GetNomor("totsimpananwajib", "nomorsimpananwajib", GetID, SimpananWajib)
            objData.Start GetDSN
            lSave = IIf(lSave, objData.Add(GetDSN, "totsimpananwajib", Array("nomorsimpananwajib", "kodeanggota", "username", "tgl", "datetime", "jumlah"), Array(Faktur, GetNull(dbData!Kode), GetRegistry(reg_UserName), Format(dInit, "yyyy-MM-dd"), SNow, nSimpananWajib)), False)
            lSave = IIf(lSave, objData.Add(GetDSN, "simpananwajib", Array("nomorsimpananwajib", "tahun", "bulan", "jumlah"), Array(Faktur, Year(dInit), Month(dInit), nSimpananWajib)), False)
            If lSave Then
              objData.Save GetDSN
            Else
              objData.Cancel GetDSN
            End If
            nValue = nValue - nSimpananWajib
            nTmp = nTmp + 1
            dInit = DateAdd("M", 1, dInit)
          Loop
        Next n
      dbData.MoveNext
    Loop
  End If
  FrmPB.EndPB
  MsgBox "Selesai"
End Sub

Private Sub Konvers(ByVal cNo As String)
Dim lSave As Boolean
Dim Faktur As String
Dim nValue As Double
Dim nBulan As Double
Dim n As Integer
Dim nTmp
Dim dInit As Date
Dim nSimpananWajib As Double

'Mengkonversi simpanan wajib bulan 1
'Pilih hanya anggota yg sudah memiliki kode
  lSave = True
  Set dbData = objData.Browse(GetDSN, "sheet1", , "kode", sisDifference, "", " AND kode IS NOT NULL")
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.EOF
      FrmPB.RunPB "Tunggu sebentar..."
      'nomorsimpananwajib
      'kodeanggota
      'username
      'tgl
      'datetime
      'jumlah
      
      'nomorsimpananwajib
      'tahun
      'bulan
      'jumlah
        
        If cNo = "1" Then
          nValue = GetNull(dbData!sepuluh)
          dInit = "2006-08-01"
          nSimpananWajib = 10000
        ElseIf cNo = "2" Then
          nValue = GetNull(dbData!wajib07)
          dInit = "2007-01-01"
          nSimpananWajib = 20000
        End If
        nBulan = 8
        For n = 1 To 3
          nTmp = 1
          Do While nValue > 0
            Faktur = GetNomor("totsimpananwajib", "nomorsimpananwajib", GetID, SimpananWajib)
            objData.Start GetDSN
            If nValue < nSimpananWajib Then
              nSimpananWajib = nValue
            End If
            lSave = IIf(lSave, objData.Add(GetDSN, "totsimpananwajib", Array("nomorsimpananwajib", "kodeanggota", "username", "tgl", "datetime", "jumlah"), Array(Faktur, GetNull(dbData!Kode), GetRegistry(reg_UserName), Format(dInit, "yyyy-MM-dd"), SNow, nSimpananWajib)), False)
            lSave = IIf(lSave, objData.Add(GetDSN, "simpananwajib", Array("nomorsimpananwajib", "tahun", "bulan", "jumlah"), Array(Faktur, Year(dInit), Month(dInit), nSimpananWajib)), False)
            If lSave Then
              objData.Save GetDSN
            Else
              objData.Cancel GetDSN
            End If
            nValue = nValue - nSimpananWajib
            nTmp = nTmp + 1
            dInit = DateAdd("M", 1, dInit)
            nSimpananWajib = 20000
          Loop
        Next n
      dbData.MoveNext
    Loop
  End If
  FrmPB.EndPB
  MsgBox "Selesai"
End Sub

Private Sub Konvers2()
Dim lSave As Boolean
Dim Faktur As String
Dim nValue As Double
Dim nBulan As Double
Dim n As Integer
Dim nTmp
Dim dInit As Date
Dim nSimpananWajib As Double

'Mengkonversi simpanan wajib bulan 1
'Pilih hanya anggota yg sudah memiliki kode
  lSave = True
  Set dbData = objData.Browse(GetDSN, "sheet1", , "kode", sisDifference, "", " AND kode IS NOT NULL")
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.EOF
      FrmPB.RunPB "Tunggu sebentar..."
      'nomorsimpananwajib
      'kodeanggota
      'username
      'tgl
      'datetime
      'jumlah
      
      'nomorsimpananwajib
      'tahun
      'bulan
      'jumlah
        
        dInit = "2006-12-31"
        nSimpananWajib = GetNull(dbData!wajib06)
        Faktur = GetNomor("totsimpananwajib", "nomorsimpananwajib", GetID, SimpananWajib)
        objData.Start GetDSN
        lSave = IIf(lSave, objData.Add(GetDSN, "totsimpananwajib", Array("nomorsimpananwajib", "kodeanggota", "username", "tgl", "datetime", "jumlah"), Array(Faktur, GetNull(dbData!Kode), GetRegistry(reg_UserName), Format(dInit, "yyyy-MM-dd"), SNow, nSimpananWajib)), False)
        lSave = IIf(lSave, objData.Add(GetDSN, "simpananwajib", Array("nomorsimpananwajib", "tahun", "bulan", "jumlah", "dk", "kodeanggota", "tgl"), Array(Faktur, Year(dInit), Month(dInit), nSimpananWajib, "K", GetNull(dbData!Kode), dInit)), False)
        If lSave Then
          objData.Save GetDSN
        Else
          objData.Cancel GetDSN
        End If
      dbData.MoveNext
    Loop
  End If
  FrmPB.EndPB
  MsgBox "Selesai"
End Sub

Private Sub Command1_Click()
  MsgBox "Nama saya Purnama"
End Sub

Private Sub BiSAButton1_Click()
  Set dbData = objData.Browse(GetDSN, "totsimpananwajib")
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.EOF
      FrmPB.RunPB
      objData.Edit GetDSN, "simpananwajib", "nomorsimpananwajib = '" & GetNull(dbData!nomorsimpananwajib) & "'", Array("kodeanggota"), Array(GetNull(dbData!KodeAnggota))
      dbData.MoveNext
    Loop
    FrmPB.EndPB
  End If
End Sub

Private Sub BiSAButton2_Click()
  Set dbData = objData.Browse(GetDSN, "totsimpananwajib")
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.EOF
      FrmPB.RunPB
      objData.Edit GetDSN, "simpananwajib", "nomorsimpananwajib = '" & GetNull(dbData!nomorsimpananwajib) & "'", Array("tgl"), Array(GetNull(dbData!Tgl))
      dbData.MoveNext
    Loop
    FrmPB.EndPB
  End If
End Sub

Private Sub BiSAButton3_Click()
  Load Form1
  Form1.Show
  Form1.SetFocus
End Sub

Private Sub BiSAButton4_Click()
  Konvers 2
End Sub

Private Sub BiSAButton5_Click()
  Konvers2
End Sub

Private Sub cmd1_Click()
  Konversi "1"
End Sub

Private Sub cmd2_Click()
  Konversi "2"
End Sub

Private Sub Form_Load()
  CenterForm Me
  SetIcon Me.hWnd
End Sub
