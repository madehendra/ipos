VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "exportExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim rs As New XArrayDB

Public Property Let RecordSource(va As XArrayDB)
  Set rs = va
End Property

Public Sub ExportToExcel(Optional ByVal lPassword As Boolean = False, Optional ByVal cPasswordWk As String = "", Optional ByVal lPreview As Boolean = True, Optional lSave As Boolean = False, Optional ByVal lNama As String)
Dim App As Excel.Application
'Dim Wb As Excel.Workbook
'Dim Wk As Excel.Worksheet

'Dim App As Object
Dim Wb As Object
Dim Wk As Object
Dim xlsheet As Object
Dim Baris As Long, Kolom As Integer



  If lPassword = True And Trim(cPasswordWk) = "" Then
    cPasswordWk = "sophiejelek"
  End If
  
  Set App = CreateObject("excel.Application")
 'Set App = New Excel.Application
  Set App = Excel.Application
  Set Wb = Excel.Workbooks.Add
  Set Wk = Wb.Worksheets(1)
  'Set xlsheet = Wb.Sheets.Item(1)
  
'  For n = 1 To 4
'    Wk.Cells(1, 1) = "Barcode"
'    Wk.Cells(1, 2) = "Nama"
'    Wk.Cells(1, 3) = "In"
'    Wk.Cells(1, 4) = "Out"
'    Wk.Cells(1, 5) = "Satuan"
'  Next n
  
  For Baris = rs.LowerBound(1) To rs.UpperBound(1)
  FrmPB.InitPB rs.UpperBound(1)
    For Kolom = 0 To rs.UpperBound(2)
      FrmPB.RunPB
    'Wk.Cells(Baris + 2, Kolom + 1) = rs(Baris, Kolom)
    Wk.Cells(Baris + 1, Kolom + 1) = rs(Baris, Kolom)
    
    Next
  Next
  FrmPB.EndPB
  
'  If lSave = True Then
'    Wk.SaveAs lNama, ".xls"
'  End If
  
  'memunculkan aplikasi excel
  
  If lPassword = True Then
'    Wb.Protect cPasswordWk
'    For Each Wk In ActiveWorkbook.Worksheets
'      Wk.Protect cPasswordWk
'    Next Wk
  End If
  
  Set Wk = Nothing
  Set Wb = Nothing
  
    
  'App.ActiveWorkbook.Save ("d:\myfilename.xls")
  App.ActiveWorkbook.SaveAs lNama, , , , False
  Set App = Nothing
  'App.Visible = lPreview
  
'  If lNama <> "" Then
'    Wb.SaveAs "d:\" & lNama, ".xls"
'  End If
  'memutus sambungan dengan excel
  

  
  'Sub UnrotectAll()
  '****Dim sh As Worksheet
  '****Dim myPassword As String
  '****myPassword = "password"
  '****
  '****For Each sh In ActiveWorkbook.Worksheets
  '********sh.Unprotect Password:=myPassword
  '****Next sh
  '****
  'End Sub

'worksheets("SheetName").protect userinterfaceonly: = True
'  For Each sh In ActiveWorkbook.Worksheets
'    sh.EnableSelection = xlUnlockedCells
'    sh.Protect
'  Next sh
  MsgBox "File Excel Berhasil Dibuat", vbInformation, "Export to Excel"
End Sub

