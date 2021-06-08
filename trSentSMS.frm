VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trSentSMS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SENT SMS"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   3840
   Begin BiSAButtonProject.BiSAButton BiSAButton2 
      Height          =   390
      Left            =   765
      TabIndex        =   10
      Top             =   1200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   688
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
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   450
      Left            =   3645
      TabIndex        =   9
      Top             =   3810
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   840
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   4890
      Width           =   3810
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   675
      Left            =   765
      TabIndex        =   7
      Top             =   3720
      Width           =   2475
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   750
      Left            =   675
      TabIndex        =   6
      Top             =   2925
      Width           =   2190
   End
   Begin BiSAButtonProject.BiSAButton BiSAButton1 
      Height          =   615
      Left            =   4155
      TabIndex        =   5
      Top             =   1590
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   1085
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
   Begin VB.OptionButton optJenis 
      Caption         =   "&2 Promo"
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
      Index           =   1
      Left            =   660
      TabIndex        =   4
      Top             =   2100
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.OptionButton optJenis 
      Caption         =   "&1 Reguler"
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
      Index           =   0
      Left            =   660
      TabIndex        =   3
      Top             =   1845
      Visible         =   0   'False
      Width           =   1185
   End
   Begin BiSADateProject.BiSADate dTgl 
      Height          =   360
      Left            =   780
      TabIndex        =   1
      Top             =   765
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   635
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
   Begin BiSAButtonProject.BiSAButton cmdOK 
      Height          =   600
      Left            =   2130
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   1058
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Tampilkan yang belum mengambil barang dari tanggal : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   780
      TabIndex        =   2
      Top             =   210
      Width           =   2700
   End
End
Attribute VB_Name = "trSentSMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Dim nPos As SisPos
Dim lEdit As Boolean
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB
Dim objMenu As New CodeSuiteLibrary.Menu

Private Sub BiSAButton1_Click()
  Load trSMSMemberBarangPromo
  trSMSMemberBarangPromo.Show
End Sub

Private Sub BiSAButton2_Click()
  GetCSV
End Sub

Private Sub cmdOK_Click()
Dim cSQL As String
Dim n As Single
Dim a As New exportExcel
Dim cJenisOPT As String

'If GetOpt(optJenis) = 1 Then
'  cJenisOPT = "R"
'Else
'  cJenisOPT = "P"
'End If
cSQL = "select t.nomorpenjualan,t.tgl,a.nama,a.telp,t.total,a.kodeanggota,t.jenis FROM totpenjualan t"
cSQL = cSQL & " LEFT JOIN anggota a on a.kodeanggota = t.kodeanggota"
cSQL = cSQL & " Where t.tgl >= '" & Format(dTgl.Value, "yyyy-MM-dd") & "' And flaglunas <> 1 AND jenis = '" & cJenisOPT & "' order by t.tgl "

  vaArray.ReDim 0, 1, 0, 12
  
  vaArray(0, 0) = aCfg(objData, msNamaPerusahaan) & " " & aCfg(objData, msAlamatPerusahaan)
  
  vaArray(1, 0) = "NOHP"
  vaArray(1, 1) = "OPERATOR"
  vaArray(1, 2) = "TGL"
  vaArray(1, 3) = "NAMA"
  vaArray(1, 4) = "NOMOR"
  vaArray(1, 5) = "NO"
  vaArray(1, 6) = "TOTAL"
  vaArray(1, 7) = "BARANG"
  vaArray(1, 8) = "HARI"
  vaArray(1, 9) = "HPTOKO"
  vaArray(1, 10) = "TOPUP"
  vaArray(1, 11) = "JENIS"
  vaArray(1, 12) = "SMS"

  vaArray.DefaultColumnType(1) = XTYPE_STRING
  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.EOF Then
  
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      
      vaArray(n, 0) = "'" & GetNull(dbData!telp)
      vaArray(n, 1) = "" & GetSelectOperatorHP(GetNull(dbData!telp))
      vaArray(n, 2) = Format(GetNull(dbData!tgl), "dd.MM.yyyy")
      vaArray(n, 3) = UCase(GetNull(dbData!nama))
      vaArray(n, 4) = ""
      vaArray(n, 5) = "'" & GetNull(dbData!nomorpenjualan)
      vaArray(n, 6) = Format(GetNull(dbData!Total), "##,###,###,##.00")
      vaArray(n, 7) = GetItemBarang(objData, GetNull(dbData!nomorpenjualan))
      vaArray(n, 8) = DateDiff("d", GetNull(dbData!tgl), Date)
      vaArray(n, 9) = "'" & aCfg(objData, msTelepon)
      vaArray(n, 10) = "" & Format(GetSaldoTopUpMember(objData, GetNull(dbData!kodeanggota)), "##,###,##.00")
      vaArray(n, 11) = GetNull(dbData!jenis)
      vaArray(n, 12) = "" & "Bonjour " & vaArray(n, 3) & ", barang SOPHIE yg diorder: " & vaArray(n, 7) & " = " & Format(vaArray(n, 6), "###,###,##0") & IIf(vaArray(n, 8) > 3, " sudah datang " & calculateAge(Format(GetNull(dbData!tgl), "yyyy-MM-dd"), Format(Date, "yyyy-MM-dd")) & " yang lalu tgl " & Format(GetNull(dbData!tgl), "dd/MM/yy") & ", silahkan diambil ya.", " sudah datang, silahkan diambil ya. ") & " MAU ORDER/INFO HUB. " & aCfg(objData, msTelepon)
      dbData.MoveNext
      
    Loop
    a.RecordSource = vaArray
    a.ExportToExcel
  End If
  

Dim strLine As String
Dim fso As New FileSystemObject
Dim fsoStream As TextStream

Set fsoStream = fso.CreateTextFile("C:\Users\Windows User\Desktop\Sample.csv", True)

'prepare the first row
strLine = "Like This is what I have tried" 'this must be written in the first column
'write the first row
fsoStream.WriteLine strLine

'prepare the second row
strLine = "But this is a sample only"      'this must be written in the first column
strLine = strLine & "," & "Sample 1"       'this must be written in the second column
'write the seconde row
fsoStream.WriteLine strLine

'prepare the third row
strLine = ""                               'an empty first column
strLine = strLine & "," & "Sample 1"       'this must be written in the second column
strLine = strLine & "," & "Sample 2"       'this must be written in the third column
'write the third row
fsoStream.WriteLine strLine

'prepare the fourth row
strLine = ""                               'an empty first column
strLine = strLine & ","                    'an empty second column
strLine = strLine & "," & "Sample 2"       'this must be written in the third column
'write the fourth row
fsoStream.WriteLine strLine

fsoStream.Close
Set fsoStream = Nothing
Set fso = Nothing
End Sub

Private Sub GetCSV()
Dim cSQL As String
Dim n As Single
Dim a As New exportExcel
Dim cJenisOPT As String
Dim strLine As String
Dim fso As New FileSystemObject
Dim fsoStream As TextStream

  CommonDialog1.Filter = "Excel File (*.csv)|*.csv"
  CommonDialog1.ShowSave

  Set fsoStream = fso.CreateTextFile(CommonDialog1.FileName, True)
'  If GetOpt(optJenis) = 1 Then
'    cJenisOPT = "R"
'  Else
'    cJenisOPT = "P"
'  End If
  cSQL = "select t.nomorpenjualan,t.tgl,a.nama,a.telp,t.total,a.kodeanggota,t.jenis FROM totpenjualan t"
  cSQL = cSQL & " LEFT JOIN anggota a on a.kodeanggota = t.kodeanggota"
  cSQL = cSQL & " Where t.tgl >= '" & Format(dTgl.Value, "yyyy-MM-dd") & "' And flaglunas <> 1 order by t.tgl "

  vaArray.ReDim 0, 1, 0, 12
  
  vaArray(0, 0) = aCfg(objData, msNamaPerusahaan) & " " & aCfg(objData, msAlamatPerusahaan)
  
  vaArray(1, 0) = "NOHP"
  vaArray(1, 1) = "OPERATOR"
  vaArray(1, 2) = "TGL"
  vaArray(1, 3) = "NAMA"
  vaArray(1, 4) = "NOMOR"
  vaArray(1, 5) = "NO"
  vaArray(1, 6) = "TOTAL"
  vaArray(1, 7) = "BARANG"
  vaArray(1, 8) = "HARI"
  vaArray(1, 9) = "HPTOKO"
  vaArray(1, 10) = "TOPUP"
  vaArray(1, 11) = "JENIS"
  vaArray(1, 12) = "SMS"
  fsoStream.WriteLine vaArray(1, 0) & "," & vaArray(1, 1) & "," & vaArray(1, 2) & "," & vaArray(1, 3) & "," & vaArray(1, 4) & "," & vaArray(1, 5) & "," & vaArray(1, 6) & "," & vaArray(1, 7) & "," & vaArray(1, 8) & "," & vaArray(1, 9) & "," & vaArray(1, 10) & "," & vaArray(1, 11) & "," & vaArray(1, 12)

  vaArray.DefaultColumnType(1) = XTYPE_STRING
  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.EOF Then
  
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      
      vaArray(n, 0) = "=" & Chr(34) & GetNull(dbData!telp) & Chr(34)
      vaArray(n, 1) = "" '& GetSelectOperatorHP(GetNull(dbData!telp))
      vaArray(n, 2) = Format(GetNull(dbData!tgl), "dd.MM.yyyy")
      vaArray(n, 3) = Replace(UCase(GetNull(dbData!nama)), ",", "")
      vaArray(n, 4) = ""
      vaArray(n, 5) = GetNull(dbData!nomorpenjualan)
      vaArray(n, 6) = GetNull(dbData!Total) 'Format(GetNull(dbData!Total), "##,###,###,##.00")
      vaArray(n, 7) = GetItemBarang(objData, GetNull(dbData!nomorpenjualan))
      vaArray(n, 8) = "=" & DateDiff("d", GetNull(dbData!tgl), Date)
      vaArray(n, 9) = "'" & aCfg(objData, msTelepon)
      vaArray(n, 10) = "" '& Format(GetSaldoTopUpMember(objData, GetNull(dbData!kodeanggota)), "##,###,##.00")
      vaArray(n, 11) = GetNull(dbData!jenis)
      vaArray(n, 12) = "" '& "Bonjour " & vaArray(n, 3) & ", barang SOPHIE yg diorder: " & vaArray(n, 7) & " = " & Format(vaArray(n, 6), "###,###,##0") & IIf(vaArray(n, 8) > 3, " sudah datang " & calculateAge(Format(GetNull(dbData!tgl), "yyyy-MM-dd"), Format(Date, "yyyy-MM-dd")) & " yang lalu tgl " & Format(GetNull(dbData!tgl), "dd/MM/yy") & ", silahkan diambil ya.", " sudah datang, silahkan diambil ya. ") & " MAU ORDER/INFO HUB. " & aCfg(objData, msTelepon)
      
      fsoStream.WriteLine vaArray(n, 0) & "," & vaArray(n, 1) & "," & vaArray(n, 2) & "," & vaArray(n, 3) & "," & vaArray(n, 4) & "," & vaArray(n, 5) & "," & vaArray(n, 6) & "," & vaArray(n, 7) & "," & vaArray(n, 8) & "," & vaArray(n, 9) & "," & vaArray(n, 10) & "," & vaArray(n, 11) & "," & vaArray(n, 12)
'      fsoStream.WriteLine vaArray(n, 0) & vbTab & vaArray(n, 1) & vbTab & vaArray(n, 2) & vbTab & vaArray(n, 3) & vbTab & vaArray(n, 4) & vbTab & vaArray(n, 5) & vbTab & vaArray(n, 6) & vbTab & vaArray(n, 7) & vbTab & vaArray(n, 8) & vbTab & vaArray(n, 9) & vbTab & vaArray(n, 10) & vbTab & vaArray(n, 11) & vbTab & vaArray(n, 12)

      dbData.MoveNext
    Loop
  End If
  fsoStream.Close
  Set fsoStream = Nothing
  Set fso = Nothing
  
End Sub

Private Function isPromo(ByVal obj As CodeSuiteLibrary.Data, ByVal cFkt As String) As Boolean
Dim db As New ADODB.Recordset

  isPromo = True
  Set db = obj.Browse(GetDSN, "penjualan", , "nomorpenjualan", sisAssign, cFkt)
  If Not db.EOF Then
    Do While Not db.EOF
      If GetNull(db!Discount) <> 0 Then
        isPromo = False
        Exit Function
      End If
      db.MoveNext
    Loop
  End If
End Function

Private Function GetItemBarang(ByVal obj As CodeSuiteLibrary.Data, ByVal cNo As String) As String
Dim db As New ADODB.Recordset
Dim cSQL2 As String

  cSQL2 = "select s.barcode,p.qty from penjualan p"
  cSQL2 = cSQL2 & " left join stock s on s.kodestock = p.kodestock"
  cSQL2 = cSQL2 & " Where p.nomorpenjualan = '" & cNo & "'"
  
  Set db = obj.SQL(GetDSN, cSQL2)
  If Not db.EOF Then
    Do While Not db.EOF
      GetItemBarang = GetItemBarang & GetNull(db!barcode) & IIf(GetNull(db!qty) > 1, "(" & GetNull(db!qty) & ")", "") & " "
      db.MoveNext
    Loop
  End If
End Function

Private Sub Command1_Click()
  MsgBox GetProcessorName
End Sub

Private Sub Command2_Click()
  Text1.Text = genNumber(GetBIOSName)
End Sub

Private Sub Command3_Click()
 If authKey(Text1.Text, GetBIOSName) Then
  MsgBox "KEY OK"
 Else
  MsgBox "INVALID KEY"
 End If
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  SetIcon Me.hWnd
  TabIndex dTgl, n
  TabIndex cmdOK, n
  dTgl.Value = Now
End Sub
