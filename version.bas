Attribute VB_Name = "version"
Option Explicit

Dim objData As New CodeSuiteLibrary.Data

Function CheckVersion()
Dim nVersion As Double
Dim nOldVersion As Double

  On Error Resume Next
  nVersion = 4
  nOldVersion = aCfg(objData, msVersion)
  
  If nOldVersion = nVersion Then Exit Function
  
'  If nOldVersion < 1 Then
'    'tambah Field
'    objData.Start GetDSN
'    objData.Sql GetDSN, "ALTER TABLE `anggota` ADD COLUMN `dd`  double NULL AFTER `nlevel`, ADD COLUMN `diskon`  double NULL AFTER `dd`"
'    objData.Save GetDSN
'  End If
  
  If nOldVersion < 4 Then
    objData.Start GetDSN
      objData.Sql GetDSN, "delete from periode"
      objData.Sql GetDSN, "INSERT INTO `periode` VALUES ('1', '0001', '2009-1-12', '2010-11-30', 'Setup Awal Akuntansi')"
      objData.Sql GetDSN, "INSERT INTO `periode` VALUES ('0', '0002', '2010-12-1', '2010-12-31', 'Desember 2010')"
    objData.Save GetDSN
  End If

  nOldVersion = nVersion
  UpdCfg msVersion, nVersion, objData, "Versi Database", "Versi Database"
  
End Function

