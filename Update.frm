VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Begin VB.Form Update 
   Caption         =   "Form1"
   ClientHeight    =   4455
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   6390
   StartUpPosition =   3  'Windows Default
   Begin BiSAButtonProject.BiSAButton BiSAButton3 
      Height          =   405
      Left            =   3105
      TabIndex        =   2
      Top             =   3180
      Width           =   1455
      _ExtentX        =   2566
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
   Begin BiSAButtonProject.BiSAButton BiSAButton2 
      Height          =   825
      Left            =   915
      TabIndex        =   1
      Top             =   2790
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   1455
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
      Height          =   570
      Left            =   1605
      TabIndex        =   0
      Top             =   1140
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   1005
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
Attribute VB_Name = "Update"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objData As New CodeSuiteLibrary.Data
Dim dbData As New ADODB.Recordset


'Public Function MBSerialNumber() As String
''RETRIEVES SERIAL NUMBER OF MOTHERBOARD
''IF THERE IS MORE THAN ONE MOTHERBOARD, THE SERIAL
''NUMBERS WILL BE DELIMITED BY COMMAS
'
''YOU MUST HAVE WMI INSTALLED AND A REFERENCE TO
''Microsoft WMI Scripting Library IS REQUIRED
'
'Dim objs As Object
'Dim obj As Object
'Dim WMI As Object
'Dim sAns As String
'
'Set WMI = GetObject("WinMgmts:")
'Set objs = WMI.InstancesOf("Win32_BaseBoard")
'For Each obj In objs
'sAns = sAns & obj.SerialNumber
'If sAns < objs.Count Then sAns = sAns & ","
'Next
'MBSerialNumber = sAns
'End Function

Private Sub BiSAButton1_Click()
  Set dbData = objData.Browse(GetDSN, "penjualan", "distinct(nomorpenjualan) as nomorpenjualan,discount", "tgl", sisAssign, "2010-05-08", " and discount = 10")
  If Not dbData.EOF Then
    objData.Start GetDSN
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.EOF
      FrmPB.RunPB
      objData.Edit GetDSN, "totpenjualan", "nomorpenjualan = '" & GetNull(dbData!nomorpenjualan) & "'", Array("kodegudang"), Array("PR")
      objData.Edit GetDSN, "kartustock", "nomor = '" & GetNull(dbData!nomorpenjualan) & "'", Array("kodegudang"), Array("PR")
      objData.Edit GetDSN, "penjualan", "nomorpenjualan = '" & GetNull(dbData!nomorpenjualan) & "'", Array("kodegudang"), Array("PR")
      dbData.MoveNext
    Loop
    FrmPB.EndPB
    objData.Save GetDSN
  End If
End Sub

Private Sub BiSAButton2_Click()
'  MsgBox MBSerialNumber
'  MsgBox Encrypt(MBSerialNumber)
'  MsgBox Decrypt(Encrypt(MBSerialNumber))
End Sub
'
''Author: Cyrus - Biohazard ®
'
'Private Sub Command1_Click(Index As Integer)
'On Error Resume Next
'Dim a_cls As New clsEncryptDecrypt
'    Select Case Index
'        Case 0
'            txtOut = a_cls.EncDecryptData(Trim(txtInput), False)
'        Case 1
'            txtOut = a_cls.EncDecryptData(Trim(txtInput), True)
'        Case 2
'            End
'    End Select
'End Sub
'
'
'
'Private strEncrypted$
'
'Public Property Get Encrypted() As String
'    Encrypted = strEncrypted$
'End Property
'
'Public Function EncDecryptData(ByVal strval$, ByVal blndec As Boolean, Optional lngEncDecVal& = 3) As String
'Dim strOutput$, _
'    PWDArr(), _
'    inttnochar&, i&
'
'    ReDim PWDArr(Len(strval$))
'
'    If blndec = False Then
'        'Encrypt
'        For i& = 1 To Len(strval$)
'            PWDArr(i&) = Asc(Mid(strval$, i&, 1))
'            PWDArr(i&) = PWDArr(i&) Xor lngEncDecVal&
'            PWDArr(i&) = Chr(PWDArr(i&) + 10)
'        Next
'        strOutput$ = Join(PWDArr, vbNullString)
'        strEncrypted$ = strOutput$
'    Else
'        'Decrypt
'        For i& = 1 To Len(strval$)
'            PWDArr(i&) = Asc(Mid(strval$, i&, 1))
'        Next
'        strval$ = ""
'
'        For i& = LBound(PWDArr) + 1 To UBound(PWDArr)
'            PWDArr(i&) = Chr(PWDArr(i&) - 10)
'            PWDArr(i&) = Asc(PWDArr(i&)) Xor lngEncDecVal&
'            strval$ = strval$ + Chr(PWDArr(i&))
'        Next
'        strOutput$ = strval$
'    End If
'
'    EncDecryptData = strOutput$
'
'End Function
'
'Public Function Distribute(ByVal strval$, ByVal intDestribute%, ByVal strSeparator$) As String
'Dim a, b, c, d, e, f, g, h, i, i2
'Dim separator_arr(), strsptr$
'Dim X As StdFont
'
'
'    ReDim separator_arr(Len(strSeparator$))
'
'    For i2 = 1 To Len(strSeparator$)
'        separator_arr(i2) = Mid$(strSeparator$, 1, 1)
'        strSeparator$ = Mid$(strSeparator$, 2)
'    Next
'
'    On Error GoTo Attached_Separator
'    a = Len(Trim(strval$))
'    If a = 0 Then Exit Function
'
'    h = Mid(CStr(Len(strval$) / intDestribute%), 1, InStr(1, CStr(Len(strval$) / intDestribute%), ".") - 1)
'    b = CLng(h)
'    c = Len(strval$) Mod intDestribute%
'    e = vbNullString
'    f = vbNullString
'
'    On Error GoTo Attached_Other
'    For i = 1 To b
'        d = Len(e) + 1
'        e = e + Mid(strval$, d, intDestribute%)
'        g = g + Mid(e, d, intDestribute%)
'        f = f + Mid(e, d, intDestribute%) & separator_arr(i)
'    Next
'
'    f = f + Mid(strval$, (Len(g) - intDestribute%) + 1)
'    Distribute = f
'    Exit Function
'
'Attached_Other:
'
'    If Err = 9 Then
'        f = f + Mid(strval$, (Len(g) - intDestribute%) + 1)
'    End If
'    Distribute = f
'    Exit Function
'
'Attached_Separator:
'    Distribute = Join(separator_arr, vbNullString)
'End Function
'
'Public Function ClassTimer(vsec%, vmin%, vhr%) As Boolean
'Static tsec%, tmin%, thr%
'Static tsec2%, tmin2%, thr2%
'
'
'    If tsec2% = vsec% And tmin2% = vmin% And thr2% = vhr% Then
'        ClassTimer = True
'        Exit Function
'    End If
'    tsec% = tsec% + 1
'    If tsec% = vsec% Then tsec2% = tsec%
'
'    If tsec% = 60 Then
'        If tmin% = vmin% Then tmin2% = tmin%
'        tmin% = tmin% + 1
'        tsec% = 0
'        If tmin% = 60 Then
'            If thr% = vhr% Then thr2% = thr%
'            thr% = thr% + 1
'            tmin% = 0
'        End If
'    End If
'
'End Function

Public Function Encrypt(ByVal icText As String) As String
 Dim icLen As Integer
 Dim icNewText As String
 Dim icChar As String
 Dim i As Integer
 
 
 icChar = ""
    icLen = Len(icText)
    For i = 1 To icLen
        icChar = Mid(icText, i, 1)
        icChar = Chr(Asc(icChar) - 1) + Chr(25)
        icNewText = icNewText + icChar
    Next
    Encrypt = icNewText
End Function

Public Function Decrypt(ByVal icText As String) As String
 Dim icLen As Integer
 Dim icNewText As String
 Dim icChar As String
 Dim i As Integer
 
 icChar = ""
    icLen = Len(icText)
    For i = 1 To icLen
        icChar = Mid(icText, i, 1)
        icChar = Chr(Asc(icChar) + 1) - Chr(25)
        icNewText = icNewText + icChar
    Next
    Decrypt = icNewText
End Function


Private Function getPlusMinus(chrr) As Boolean ' <<< This function retunrs either true or false
chrr = UCase(chrr)                             '     depending on if a charachter is more than
                                               '     halfway through the alphabet or not...
If Asc(chrr) - 65 < 12 Then
    getPlusMinus = True
Else
    getPlusMinus = False
End If
End Function

'Public Function genNumber(appName)
'Dim appVal As Long
'Dim genVal As Long
'Dim tmpVar As String
'Dim i As Integer
'Dim seedMod As Integer
'
'For i = 1 To Len(appName) - 0
'    appVal = appVal + Val(Asc(Mid$(appName, i, 1))) ' <<< Counts the value of each ascii chr
'Next                                                '     in the app name
'seedMod = Int((Day(Date) & Month(Date) & Year(Date) & Hour(time) & Minute(time) & Second(time)) ^ 0.2)
'
''For i = 0 To Int(seedMod + Minute(time) & Second(time)) ' <<< Vb's random num generator is not
''    Rnd                                                 '     very random so i will make it more
''Next                                                    '     random
'
'tmpVar = ""
'For i = 1 To 20                                   ' <<< Randomly create the 1st 4 parts of the code
'    If Rnd < 0.5 Then                             ' <<< 1 in two chance of a letter or a number
'        tmpVar = tmpVar & Chr(Int(Rnd * 25) + 65)
'    Else
'        tmpVar = tmpVar & Int(Rnd * 9)
'    End If
'
'    If Int(i / 5) = i / 5 And i <> 25 Then    ' <<< Add a ' - ' every 5 charachters
'        tmpVar = tmpVar & " - "
'    End If
'Next
'
'For i = 1 To Len(tmpVar) - 0                              ' <<< Creates a number based on the
'    If i < Len(appName) Then                              '     first sections. Adds or takes
'        If getPlusMinus(Mid(appName, i, 1)) = False Then  '     depending on various things
'            genVal = genVal + Val(Asc(Mid$(tmpVar, i, 1))) '    Makes it mathematicaly harder
'        Else                                              '     to re-order the code.
'            genVal = genVal - Val(Asc(Mid$(tmpVar, i, 1)))
'        End If
'    Else
'        If Int(i / 2) = i / 2 Then
'            genVal = genVal - Val(Asc(Mid$(tmpVar, i, 1)))
'        Else
'            genVal = genVal + Val(Asc(Mid$(tmpVar, i, 1)))
'        End If
'    End If
'Next
'If genVal < 0 Then genVal = 0 - genVal      ' <<< If the number is less than 0 then make it
'                                            '     positive
'
'tmpVar = tmpVar & Mid((genVal * appVal) & "JSDEU", 1, 5) ' <<< Last part of the code is the
'                                                         '     'value' of the first part of
'                                                         '     the code times the 'value'
'                                                         '     of the program name, limited
'                                                         '     to 5 charachters. "JSDEU" is
'                                                         '     to make sure the result is
'                                                         '     atleast 5 chars.
'
'genNumber = UCase(tmpVar)    ' <<< Returns the new key
'End Function
'
'
'Public Function authKey(key, appName) As Boolean
'authKey = False
'On Error GoTo err
'
'Dim splt() As String
'Dim appVal As Long
'Dim genVal As Long
'Dim tempVar As String
'Dim i As Integer
'key = UCase(key)
'
'For i = 1 To Len(appName) - 0
'    appVal = appVal + Val(Asc(Mid$(appName, i, 1)))
'Next
'
'splt = Split(key, " - ")
'splt(4) = ""
'
'tempVar = Join(splt, " - ")
'
'For i = 1 To Len(tempVar) - 0
'    If i < Len(appName) Then
'        If getPlusMinus(Mid(appName, i, 1)) = False Then
'            genVal = genVal + Val(Asc(Mid$(tempVar, i, 1)))
'        Else
'            genVal = genVal - Val(Asc(Mid$(tempVar, i, 1)))
'        End If
'    Else
'        If Int(i / 2) = i / 2 Then
'            genVal = genVal - Val(Asc(Mid$(tempVar, i, 1)))
'        Else
'            genVal = genVal + Val(Asc(Mid$(tempVar, i, 1)))
'        End If
'    End If
'Next
'If genVal < 0 Then genVal = 0 - genVal
'
'splt = Split(key, " - ")
'
'If genVal = Val(splt(4)) / appVal Then
'    authKey = True
'Else
'    authKey = False
'End If
'
'
'Debug.Print Mid((appVal * genVal) & "JSDEU", 1, 5)
'Debug.Print splt(4)
'
'If Mid((appVal * genVal) & "JSDEU", 1, 5) = splt(4) Then
'    authKey = True
'Else
'    authKey = False
'End If
'
'err:
'
'End Function

'DEMO SUB PROCEDURE
Private Sub Command1_Click()
'    Dim myAppName As String
'    myAppName = Text1.Text
'    mySN = genNumber(myAppName)
'    Label1.Caption = mySN
End Sub

Private Sub BiSAButton3_Click()
'<<   - genNumber(nameOfProgram)                                         >>
'<<   - authKey(Key, nameOfProgram)
    Dim myAppName As String
    Dim mysn As String
    
'    myAppName = MBSerialNumber
'    mysn = genNumber(MBSerialNumber)
'    auth
'     Label1.Caption = mySN

MsgBox mysn & vbCrLf & authKey(mysn, myAppName)

End Sub
