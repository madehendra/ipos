Attribute VB_Name = "FuncDek2Text"
Option Explicit

Function Dec2Text(ByVal nDec As Double) As String
Dim cRetval As String
  Select Case nDec
    Case Is <= 11
      cRetval = cRetval & Satuan(nDec)
    Case Is <= 99
      cRetval = cRetval & Puluhan(nDec)
    Case Is <= 999
      cRetval = cRetval & Ratusan(nDec)
    Case Is <= 999999
      cRetval = cRetval & Ribuan(nDec)
    Case Is <= 999999999
      cRetval = cRetval & Jutaan(nDec)
    Case Is <= 999999999999#
      cRetval = cRetval & Milyard(nDec)
    Case Is <= 999999999999999#
      cRetval = cRetval & Trilyon(nDec)
  End Select
  Dec2Text = cRetval
End Function

Private Function Satuan(nDec As Double) As String
Dim vaSatuan
  If nDec > 0 Then
    vaSatuan = Array("Satu", "Dua", "Tiga", "Empat", "Lima", _
                     "Enam", "Tujuh", "Delapan", "Sembilan", _
                     "Sepuluh", "Sebelas")
    Satuan = vaSatuan(nDec - 1) & " "
  End If
End Function

Private Function Puluhan(nDec As Double) As String
Dim cRetval As String, cDecimal As String
  cDecimal = LTrim(str(Round(nDec, 0)))
  If nDec > 0 Then
    Select Case nDec
      Case Is <= 11
        cRetval = cRetval + Satuan(nDec)
      Case Is <= 19
        cRetval = cRetval + Satuan(Mid(cDecimal, 2, 1)) & "Belas "
      Case Is <= 99
        cRetval = cRetval & Satuan(Left(cDecimal, 1)) & "Puluh "
        cRetval = cRetval & Satuan(Mid(cDecimal, 2, 1))
    End Select
  End If
  Puluhan = cRetval
End Function

Private Function Ratusan(nDec As Double) As String
Dim cRetval As String, cDecimal As String

  cDecimal = LTrim(str(Round(nDec, 0)))
  If nDec > 0 Then
    Select Case nDec
      Case Is <= 99
        cRetval = cRetval & Puluhan(nDec)
      Case Is <= 199
        cRetval = cRetval & "Seratus " & Puluhan(Mid(cDecimal, 2))
      Case Is <= 999
        cRetval = cRetval & Satuan(Left(cDecimal, 1)) & "Ratus "
        cRetval = cRetval & Puluhan(Mid(cDecimal, 2))
    End Select
  End If
  Ratusan = cRetval
End Function

Private Function Ribuan(nDec As Double) As String
Dim cRetval As String, cDecimal As String
  
  cDecimal = LTrim(str(Round(nDec, 0)))
  If nDec > 0 Then
    Select Case nDec
      Case Is <= 999
        cRetval = cRetval & Ratusan(nDec)
      Case Is <= 1999
        cRetval = cRetval & "Seribu " & Ratusan(Mid(cDecimal, 2))
      Case Is <= 999999
        cDecimal = Padl(cDecimal, 6, "0")
        cRetval = cRetval & Ratusan(Left(cDecimal, 3)) & "Ribu "
        cRetval = cRetval & Ratusan(Mid(cDecimal, 4))
    End Select
  End If
  Ribuan = cRetval
End Function

Private Function Jutaan(nDec As Double) As String
Dim cRetval As String, cDecimal As String
  
  cDecimal = LTrim(str(Round(nDec, 0)))
  If nDec > 0 Then
    Select Case nDec
      Case Is <= 999999
        cRetval = cRetval & Ribuan(nDec)
      Case Is <= 999999999
        cDecimal = Padl(cDecimal, 9, "0")
        cRetval = cRetval & Ratusan(Left(cDecimal, 3)) & "Juta "
        cRetval = cRetval & Ribuan(Mid(cDecimal, 4))
    End Select
  End If
  Jutaan = cRetval
End Function

Private Function Milyard(nDec As Double) As String
Dim cRetval As String, cDecimal As String
  
  cDecimal = LTrim(str(Round(nDec, 0)))
  If nDec > 0 Then
    Select Case nDec
      Case Is <= 999999999
        cRetval = cRetval & Jutaan(nDec)
      Case Is <= 999999999999#
        cDecimal = Padl(cDecimal, 12, "0")
        cRetval = cRetval & Ratusan(Left(cDecimal, 3)) & "Milyard "
        cRetval = cRetval & Jutaan(Mid(cDecimal, 4))
    End Select
  End If
  Milyard = cRetval
End Function

Private Function Trilyon(nDec As Double) As String
Dim cRetval As String, cDecimal As String
  
  cDecimal = LTrim(str(Round(nDec, 0)))
  If nDec > 0 Then
    Select Case nDec
      Case Is <= 999999999999#
        cRetval = cRetval & Milyard(nDec)
      Case Is <= 999999999999999#
        cDecimal = Padl(cDecimal, 15, "0")
        cRetval = cRetval & Ratusan(Left(cDecimal, 3)) & "Trilyon "
        cRetval = cRetval & Milyard(Mid(cDecimal, 4))
    End Select
  End If
  Trilyon = cRetval
End Function
