Attribute VB_Name = "reg"
Public cNamaKomputer
Dim DeviceFound() As Variant
Dim DeviceList() As Variant
Dim DeviCecount As Integer

Dim oDeviceType() As Variant
Dim oDeviceCaption() As Variant
Dim oDeviceParam() As Variant
Dim oDeviceInterf() As Variant

Function EncDec(inData As Variant, Optional inPW As Variant = "") As Variant
     On Error Resume Next
     Dim arrSBox(0 To 255) As Integer
     Dim arrPW(0 To 255) As Integer
     Dim Bi As Integer, Bj As Integer
     Dim mKey As Integer
     Dim i As Integer, j As Integer
     Dim x As Integer, y As Integer
     Dim mCode As Byte, mCodeSeries As Variant
    
     EncDec = ""
     If Trim(inData) = "" Then
         Exit Function
     End If
    
     If inPW <> "" Then
         j = 1
         For i = 0 To 255
             arrPW(i) = Asc(Mid$(inPW, j, 1))
             j = j + 1
             If j > Len(inPW) Then
                  j = 1
             End If
         Next i
     Else
         For i = 0 To 255
             arrPW(i) = 0
         Next i
     End If
     
       ' Reseed arrSBox()
     For i = 0 To 255
         arrSBox(i) = i
     Next i
     
     j = 0
     For i = 0 To 255
         j = (arrSBox(i) + arrPW(i)) Mod 256
           ' Swap
         x = arrSBox(i)
         arrSBox(i) = arrSBox(j)
         arrSBox(j) = x
     Next i
     
     mCodeSeries = ""
     Bi = 0: Bj = 0
     For i = 1 To Len(inData)
         Bi = (Bi + 1) Mod 256
         Bj = (Bj + arrSBox(Bi)) Mod 256
           ' Swap
         x = arrSBox(Bi)
         arrSBox(Bi) = arrSBox(Bj)
         arrSBox(Bj) = x
            'Generate a key
         mKey = arrSBox((arrSBox(Bi) + arrSBox(Bj)) Mod 256)
            'xor the key
         mCode = Asc(Mid$(inData, i, 1)) Xor mKey
         mCodeSeries = mCodeSeries & Chr(mCode)
     Next i
     EncDec = mCodeSeries
End Function

Sub GetScan()

    cNamaKomputer = ""
    Set objLocator = CreateObject("WbemScripting.SWbemLocator")
    Set objService = objLocator.ConnectServer("", "", "", "")
    Set objWshNet = CreateObject("Wscript.Network")
    cNamaKomputer = objWshNet.ComputerName
    
    eilute = 0
    ReDim Preserve DeviceList(40)
    ReDim Preserve DeviceFound(40)
    DeviceListLen = 2
    DeviceList = Array("Win32_Processor", "Win32_Processor")
    For i = 0 To DeviceListLen - 1
        Set objDeviceSet = objService.InstancesOf(DeviceList(i))
        If objDeviceSet.Count <> 0 Then
            DeviceFound(DeviCecount) = DeviceList(i)
            DeviCecount = DeviCecount + 1
            Call GetSndDevInfo(objService, DeviceList(i))
        End If
    Next
    cNamaKomputer = UCase(Trim(Replace(cNamaKomputer, " ", "")))
End Sub


Sub GetSndDevInfo(objService, strWBEMClass)
  
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


