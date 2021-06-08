Attribute VB_Name = "modSimpananHarian"
Public Function GetSaldoTabungan(ByVal obj As CodeSuiteLibrary.Data, ByVal kodeanggota As String, ByVal dAwal As Date, ByVal dAkhir As Date) As Double
Dim cTgl As String
Dim nJumlah As Double
Dim cWhere As String
Dim dbSaldo As New ADODB.Recordset
  
  cWhere = " Tgl <= '" & Format(dAkhir, "yyyy-MM-dd") & "' and KodeAnggota = '" & kodeanggota & "'"
  nJumlah = 0
  Set dbSaldo = obj.Browse(GetDSN, "MutasiTabungan", "DK,Jumlah", , , , cWhere)
  If Not dbSaldo.EOF Then
    Do While Not dbSaldo.EOF
      nJumlah = nJumlah + IIf(dbSaldo!DK = "K", dbSaldo!Jumlah, -dbSaldo!Jumlah)
      dbSaldo.MoveNext
    Loop
  End If
  GetSaldoTabungan = nJumlah
End Function
