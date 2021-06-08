VERSION 5.00
Begin VB.Form trReg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form Registrasi"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   6675
   Begin VB.TextBox Text2 
      Height          =   660
      Left            =   270
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   1860
      Width           =   6285
   End
   Begin VB.TextBox Text1 
      Height          =   525
      Left            =   255
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   600
      Width           =   6270
   End
   Begin VB.CommandButton cmdEncrypt 
      Caption         =   "Command2"
      Height          =   465
      Left            =   2595
      TabIndex        =   3
      Top             =   2940
      Width           =   1740
   End
   Begin VB.CommandButton cmdDecrypt 
      Caption         =   "Command2"
      Height          =   450
      Left            =   240
      TabIndex        =   2
      Top             =   2955
      Width           =   1860
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   525
      Left            =   4980
      TabIndex        =   1
      Top             =   2955
      Width           =   1320
   End
   Begin VB.TextBox Text3 
      Height          =   345
      Left            =   300
      TabIndex        =   0
      Text            =   "Text3"
      Top             =   1230
      Width           =   6165
   End
End
Attribute VB_Name = "trReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DeviceFound() As Variant
Dim DeviceList() As Variant
Dim DeviCecount As Integer
Dim ramas As Variant
Dim ramotipas As Variant
Dim PelesInt() As Variant
Dim PelesTipas() As Variant

Dim isClient As Boolean
Dim isClienta As Boolean
Dim strUserName As String
Dim strPassword As String
Dim klientoID As Integer
Dim webUserName As String
Dim webPassword As String

Dim oDeviceType() As Variant
Dim oDeviceCaption() As Variant
Dim oDeviceParam() As Variant
Dim oDeviceInterf() As Variant

Dim eilute As Integer
Dim isHardware As Boolean

Dim cNamaKomputer As String


Dim User_Text_Len               As Double

Dim I_For                       As Double
Dim Char_Val                    As String

Dim Encrypt_Char                As String
Dim Decrypt_Char                As String
    

Private Sub Command1_Click()
  GetScan
End Sub

Private Sub GetScan()
        
    cNamaKomputer = ""
    Set objLocator = CreateObject("WbemScripting.SWbemLocator")
    Set objService = objLocator.ConnectServer("", "", "", "")
    Set objWshNet = CreateObject("Wscript.Network")
    cNamaKomputer = objWshNet.ComputerName
    
    eilute = 0
    ReDim Preserve DeviceList(40)
    ReDim Preserve DeviceFound(40)
    DeviceListLen = 2
    DeviceList = Array("Win32_DiskDrive", _
                "Win32_Processor")
    For i = 0 To DeviceListLen - 1
        Set objDeviceSet = objService.InstancesOf(DeviceList(i))
        If objDeviceSet.Count <> 0 Then
            DeviceFound(DeviCecount) = DeviceList(i)
            DeviCecount = DeviCecount + 1
            Call GetSndDevInfo(objService, DeviceList(i))
        End If
    Next
    MsgBox cNamaKomputer
    Text1.Text = UCase(Trim(Replace(cNamaKomputer, " ", "")))
End Sub


Private Sub GetSndDevInfo(objService, strWBEMClass)
  
    On Error Resume Next
    
    ReDim Preserve oDeviceType(100)
    ReDim Preserve oDeviceCaption(100)
    ReDim Preserve oDeviceParam(100)
    ReDim Preserve oDeviceInterf(100)
    

    Set objDeviceSet = objService.InstancesOf(strWBEMClass)
    'MsgBox strWBEMClass
    If objDeviceSet.Count <> 0 Then
        For Each Device In objDeviceSet
        
    Select Case strWBEMClass
' DISK----------------------------------
        Case "Win32_DiskDrive"
            List1.AddItem Device.Description & vbTab & Device.Caption & vbTab & Device.Size & vbTab & Device.InterfaceType
            oDeviceType(eilute) = Device.Description
            oDeviceCaption(eilute) = Device.Caption
            oDeviceParam(eilute) = Device.Size
            oDeviceInterf(eilute) = Device.InterfaceType
            cNamaKomputer = cNamaKomputer & ";" & oDeviceCaption(eilute) & ";" & oDeviceParam(eilute)
            eilute = eilute + 1
' PROCESSOR-------------------------------------
        Case "Win32_Processor"
            List1.AddItem Device.Role & vbTab & vbTab & Device.Name & vbTab & Device.CurrentClockSpeed & vbTab & ""
            oDeviceType(eilute) = Device.Role
            oDeviceCaption(eilute) = Device.Name
            oDeviceParam(eilute) = Device.CurrentClockSpeed
            oDeviceInterf(eilute) = ""
            cNamaKomputer = cNamaKomputer & ";" & oDeviceCaption(eilute)
            eilute = eilute + 1
            
    End Select
        Next
    End If
End Sub

Private Sub CmdDecrypt_Click()

    
'    Dim I_Oct                       As Double
'    Dim Len_Char_Val                As Integer
'    Dim Oct_Val                     As Integer
'    Dim B_Oct                       As Integer
'    Dim Temp                        As Byte
'
'    User_Text = "": User_Text = StrReverse(Text1.Text)
'    Decrypt_Char = ""
'    If User_Text = "" Then Exit Sub
'    User_Text_Len = Len(User_Text)
'
'    For I_For = 1 To User_Text_Len
'        Char_Val = Asc(Mid(User_Text, I_For, 1))
'        '----------------------------------------------------
'            If Char_Val <= 245 Then Char_Val = Char_Val - 10
'            Len_Char_Val = Len(Char_Val)
'        '----------------------------------------------------
'        Oct_Val = 0
'        For I_Oct = 0 To Len_Char_Val - 1
'            B_Oct = 8 ^ I_Oct
'            If I_Oct = 0 Then Temp = B_Oct Else Temp = I_Oct + 1
'            Oct_Val = CInt(Oct_Val) + CInt((B_Oct) * (Mid(StrReverse(Char_Val), Temp, 1)))
'        Next I_Oct
'        Decrypt_Char = Decrypt_Char & Chr(Oct_Val)
'    Next I_For
'
'    Text3.Text = Decrypt_Char
'    cmdEncrypt.Enabled = True
'    cmdDecrypt.Enabled = False
    
Text3.Text = Cryptt(Text1.Text, False)
End Sub

Private Sub CmdEncrypt_Click()
    
'    Dim Check_Val                   As Byte
'
'    Check_Val = 0
'    User_Text = "": User_Text = Text3.Text
'    Encrypt_Char = ""
'    If User_Text = "" Then Exit Sub
'    User_Text_Len = Len(User_Text)
'
'    For I_For = 1 To User_Text_Len
'        Char_Val = Asc(Mid(User_Text, I_For, 1))
'        '--------------------------------------------
'            Check_Val = 0: Check_Val = Oct(Char_Val)
'        '--------------------------------------------
'        If Check_Val >= 245 Then
'            Encrypt_Char = Encrypt_Char & Chr(Check_Val)
'        Else
'            Check_Val = Check_Val + 10
'            Encrypt_Char = Encrypt_Char & Chr(Check_Val)
'        End If
'    Next I_For
'
'    Text2.Text = StrReverse(Encrypt_Char)
'    cmdEncrypt.Enabled = False
'    cmdDecrypt.Enabled = True
'

Text2.Text = Cryptt(Text3.Text, True)
If Text1.Text = Left(Text2.Text, Len(Text2.Text) - 1) Then
  MsgBox "Benar"
End If
End Sub

Function Zip(ByVal cPassword As String, Optional lTile As Boolean = True) As String
Dim n As Single
Dim cRetval As String
Dim cZip As String
Dim cPad As String

  If lTile Then
    cPassword = Trim(cPassword)
    cPassword = IIf(cPassword = "", " ", cPassword)
    Do While Len(cPad) <= 20
      cPad = cPad & cPassword
    Loop
    cPassword = Left(cPad, 20)
  End If

  For n = 1 To Len(cPassword)
    cRetval = cRetval & Trim(str(Asc(Mid(cPassword, n, 1)) + (n * 59)))
  Next

  For n = 1 To Len(cRetval) Step 2
    cZip = cZip & Chr(Val(Mid(cRetval, n, 2)) + 65)
  Next
  Zip = cZip
'  Zip = cPassword
End Function

Function UnZip(ByVal cPassword As String) As String
Dim n As Single
Dim cRetval As String
Dim cUnZip As String
Dim i As Single
Dim c As String

  For n = 1 To Len(cPassword)
    c = Trim(str(Asc(Mid(cPassword, n, 1)) - 65))
    cUnZip = cUnZip & IIf(Len(c) = 1, "0", "") & c
  Next

  For n = 1 To Len(cUnZip) Step 4
    i = i + 1
    cRetval = cRetval & Chr(Val(Mid(cUnZip, n, 4)) - (i * 59))
  Next

  UnZip = cRetval
'  UnZip = cPassword
End Function


Public Function Cryptt(ByVal StrString As String, Decrypt As Boolean)
Dim intloop As Integer
Dim intloop2 As Integer
Dim Rand As Integer
Dim intRand As Integer
Dim start As Integer
Dim CS As String
Dim CS2
Dim CS3

On Error Resume Next

CS = ""
CS2 = StrString
CS3 = ""
Randomize

If Decrypt = True Then

  If Len(CS) Mod 2 = 0 Then

    For intloop2 = 1 To 3
        CS = ""
        
        For intloop = 1 To Len(CS2) Step 6
             CS = CS & Chr(Mid(CS2, intloop, 2))
        Next intloop
        CS2 = CS
    Next intloop2
     CS = CS2

     CS2 = ""

    For intloop = 1 To Len(CS) / 2 Step 1
        CS2 = CS2 & Mid(CS, intloop, 1)
    Next intloop

    For intloop = Len(CS) / 2 + 1 To Len(CS) Step 1
        CS3 = CS3 & Mid(CS, intloop, 1)
    Next intloop

        CS = ""
    
    For intloop = 1 To Len(CS3) + Len(CS2) Step 1
        CS = CS & Mid(CS3, intloop, 1) & Mid(CS2, intloop, 1)
    Next intloop
    
        CS2 = ""
    
    For intloop = Len(CS) To 1 Step -1
        CS2 = CS2 & Mid(CS, intloop, 1)
    Next intloop

        CS = ""

    For intloop = 1 To Len(CS2) / 2 Step 1
        CS = CS & Mid(CS2, intloop, 1)
    Next intloop

        CS3 = ""
    
    For intloop = Len(CS) To 1 Step -1
        CS3 = CS3 & Mid(CS, intloop, 1)
    Next intloop

        CS = ""
    
    For intloop = Len(CS2) / 2 + 1 To Len(CS2) Step 1
        CS = CS & Mid(CS2, intloop, 1)
    Next intloop

    CS2 = CS
    CS = ""

    For intloop = 1 To Len(CS2) Step 1
        CS = CS & Mid(CS2, intloop, 1) & Mid(CS3, intloop, 1)
    Next intloop

    CS2 = CS
    CS = ""
    CS3 = ""
    
Else
  
    For intloop2 = 1 To 3
        CS = ""
        For intloop = 1 To Len(CS2) Step 6
             CS = CS & Chr(Mid(CS2, intloop, 2))
        Next intloop
        CS2 = CS
    Next intloop2
    MsgBox CS
     CS = CS2

     CS2 = ""
     
  
    For intloop = 1 To Len(CS) / 2 Step 1
        CS2 = CS2 & Mid(CS, intloop, 1)
    Next intloop

    For intloop = Len(CS) / 2 + 1 To Len(CS) Step 1
        CS3 = CS3 & Mid(CS, intloop, 1)
    Next intloop

        CS = ""
    
    For intloop = 1 To Len(CS3) + Len(CS2) Step 1
        CS = CS & Mid(CS3, intloop, 1) & Mid(CS2, intloop, 1)
    Next intloop
    
        CS2 = ""
    
    For intloop = Len(CS) To 1 Step -1
        CS2 = CS2 & Mid(CS, intloop, 1)
    Next intloop

        CS = ""

    For intloop = 1 To Len(CS2) / 2 Step 1
        CS = CS & Mid(CS2, intloop, 1)
    Next intloop

        CS3 = ""
    
    For intloop = Len(CS) To 1 Step -1
        CS3 = CS3 & Mid(CS, intloop, 1)
    Next intloop

        CS = ""
    
    For intloop = Len(CS2) / 2 + 1 To Len(CS2) Step 1
        CS = CS & Mid(CS2, intloop, 1)
    Next intloop

    CS2 = CS
    CS = ""

    For intloop = 1 To Len(CS2) Step 1
        CS = CS & Mid(CS2, intloop, 1) & Mid(CS3, intloop, 1)
    Next intloop

    CS2 = CS
    CS = ""
    CS3 = ""
  
  
  End If
    
ElseIf Decrypt = False Then
        
        CS = ""
        
    If Len(CS2) Mod 2 = 0 Then
        
        
        For intloop = 2 To Len(CS2) Step 2
            CS = CS & Mid(CS2, intloop, 1)
        Next intloop

        For intloop = Len(CS) To 1 Step -1
            CS3 = CS3 & Mid(CS, intloop, 1)
        Next intloop

            CS = CS3
            CS3 = ""
            
        For intloop = 1 To Len(CS2) - 1 Step 2
        Next intloop

            CS2 = CS
            CS = ""
            
        For intloop = Len(CS2) To 1 Step -1
        Next intloop

            CS2 = ""
            
        For intloop = 2 To Len(CS) Step 2
            CS2 = CS2 & Mid(CS, intloop, 1)
        Next intloop

            
        For intloop = 1 To Len(CS) - 1 Step 2
            CS2 = CS2 & Mid(CS, intloop, 1)
        Next intloop
        
        
        For intloop2 = 1 To 3
        
                CS = ""
                
            For intloop = 1 To Len(CS2) Step 1
                Rand = Rand - Rand
                    For intRand = 1 To 4 Step 1
                        Rand = Rand & Int((9 - 1 + 1) * Rnd + 1)
                    Next intRand
                
                
                CS = CS & Asc(Mid(UCase(CS2), intloop, 1)) & Rand
            Next intloop
            CS2 = CS
        
        Next intloop2
        
        CS = ""
    Else
        
        CS2 = CS2 & Mid(CS2, Len(CS2), 1)
        MsgBox CS2
        
        For intloop = 2 To Len(CS2) Step 2
            CS = CS & Mid(CS2, intloop, 1)
        Next intloop

        For intloop = Len(CS) To 1 Step -1
            CS3 = CS3 & Mid(CS, intloop, 1)
        Next intloop

            CS = CS3
            CS3 = ""
            
        For intloop = 1 To Len(CS2) - 1 Step 2
            CS = CS & Mid(CS2, intloop, 1)
        Next intloop

            CS2 = CS
            CS = ""
            
        For intloop = Len(CS2) To 1 Step -1
            CS = CS & Mid(CS2, intloop, 1)
        Next intloop

            CS2 = ""
            
        For intloop = 2 To Len(CS) Step 2
            CS2 = CS2 & Mid(CS, intloop, 1)
        Next intloop

            
        For intloop = 1 To Len(CS) - 1 Step 2
            CS2 = CS2 & Mid(CS, intloop, 1)
        Next intloop
        
        CS = ""
        
        For intloop2 = 1 To 3
                CS = ""
                
                
            For intloop = 1 To Len(CS2) Step 1
                Rand = Rand - Rand
                    For intRand = 1 To 4 Step 1
                        Rand = Rand & Int((9 - 1 + 1) * Rnd + 1)
                    Next intRand
                
                CS = CS & Asc(Mid(UCase(CS2), intloop, 1)) & Rand
            Next intloop
            CS2 = CS
        
        Next intloop2
        
    End If
    
End If

Cryptt = CS2
End Function
