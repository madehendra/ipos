VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Begin VB.Form frmExcel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Katalog"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   6645
   Begin BiSAButtonProject.BiSAButton BiSAButton7 
      Height          =   615
      Left            =   1275
      TabIndex        =   7
      Top             =   2595
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1085
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
   Begin BiSAButtonProject.BiSAButton BiSAButton6 
      Height          =   615
      Left            =   1275
      TabIndex        =   6
      Top             =   1950
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1085
      Caption         =   "Open"
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
      Height          =   495
      Left            =   1290
      TabIndex        =   5
      Top             =   4815
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   873
      Caption         =   "PROSES NEW DBF"
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
      Height          =   495
      Left            =   1290
      TabIndex        =   4
      Top             =   4260
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   873
      Caption         =   "OPEN NEW DBF"
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
      Left            =   5400
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin BiSAButtonProject.BiSAButton BiSAButton2 
      Height          =   525
      Left            =   1245
      TabIndex        =   1
      Top             =   165
      Width           =   3630
      _ExtentX        =   6403
      _ExtentY        =   926
      Caption         =   "Buka File DBF"
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
      Height          =   510
      Left            =   1245
      TabIndex        =   0
      Top             =   720
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   900
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
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   5400
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog3 
      Left            =   5400
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
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
      Left            =   1320
      TabIndex        =   11
      Top             =   3960
      Width           =   4695
   End
   Begin VB.Label Label5 
      Caption         =   "Step 1."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   10
      Top             =   2055
      Width           =   1140
   End
   Begin VB.Label Label4 
      Caption         =   "Step 2."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   9
      Top             =   2670
      Width           =   990
   End
   Begin VB.Label Label3 
      Caption         =   "Buka File Daily Report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2280
      TabIndex        =   8
      Top             =   1665
      Width           =   1755
   End
   Begin VB.Label Label2 
      Caption         =   "Step 2."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   3
      Top             =   780
      Width           =   990
   End
   Begin VB.Label Label1 
      Caption         =   "Step 1."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   2
      Top             =   135
      Width           =   1140
   End
End
Attribute VB_Name = "frmExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Excel As Object 'Excel.Application
Dim ExcelWBk As Object 'Excel.Workbook
Dim ExcelWS As Object 'Excel.Worksheet
Dim oSheet As Object
Dim objData As New CodeSuiteLibrary.Data
Dim dbData As New ADODB.Recordset
Dim vaArray As New XArrayDB

Private Sub StartExcel()
  On Error GoTo err:
  Set Excel = GetObject(, "Excel.Application")
  Exit Sub
err:
  Set Excel = CreateObject("Excel.Application")
End Sub

Private Sub CloseWorkSheet()
  ExcelWBk.Close
  Excel.Quit
End Sub

Private Sub FinishExcel()
  'Jangan lupa, selalu bersihkan memory saat mengakhiri
  If Not ExcelWS Is Nothing Then Set ExcelWS = Nothing
  If Not ExcelWBk Is Nothing Then Set ExcelWBk = Nothing
  If Not Excel Is Nothing Then Set Excel = Nothing
End Sub

Private Sub BiSAButton1_Click()
Dim lSave As Boolean
Dim vaField, vaValue
Dim i, j As Integer

  StartExcel
  lSave = True
  
  objData.Start GetDSN
    
  Excel.Workbooks.Close
  Set ExcelWBk = Excel.Workbooks.Open(CommonDialog1.FileName)
  Set ExcelWS = ExcelWBk.Worksheets(1)
  
  MsgBox "Mohon bersabar, tunggu sampai indikator menunjukkan selesai"
  
  'MsgBox "Yakin akan menambahkan " & ExcelWS.Cells.SpecialCells(xlCellTypeLastCell).Row & " Item stock baru??"
  'Loop berikut untuk membaca nilai setiap baris
  'mulai dari baris pertama sampai ketiga
  'dan setiap baris terdiri dari 2 kolom
  
  FrmPB.InitPB 2000
  Dim cNama, cBarcode, cKodeGolongan, cKodeSatuan
  Dim nHargaBeli, nHargaJual, nDiskonPenjualan
  Dim cJenis, cBiaya
  Dim nBV

  For i = 2 To ExcelWS.Cells.SpecialCells(xlCellTypeLastCell).Row
    FrmPB.RunPB
   For j = 1 To 8
    With ExcelWS
      cNama = .Cells(i, 1).Value
      cBarcode = .Cells(i, 2).Value
      cKodeGolongan = .Cells(i, 8).Value
      cKodeSatuan = "PCS"
      nHargaBeli = .Cells(i, 3).Value
      nHargaJual = .Cells(i, 3).Value
      nDiskonPenjualan = .Cells(i, 5).Value
      nBV = .Cells(i, 4).Value
      cJenis = 1
      cBiaya = 1
      '####################
      'FORMAT UPDATE STOCK
      '####################
      '1. NAMA
      '2. BARCODE
      '3. HARGA BELI
      '4. - (CV)
      '5. DISKON PENJUALAN
      '6. -
      '7. -
      '8. KODE GOLONGAN
      '####################
    End With
   Next j
   
    If Trim(cNama) = "" Then
      Exit For
    End If
    vaField = Array("nama", "barcode", "kodegolongan", _
                    "kodesatuan", _
                    "hargabeli", "hargajual", "jenis", "asbiaya", "diskonpenjualan", "datetime", "bv")
                    
    vaValue = Array(StrConv(cNama, vbProperCase), cBarcode, cKodeGolongan, _
                     "PCS", _
                     nHargaBeli, nHargaJual, 1, 2, nDiskonPenjualan, SNow, nBV)
                     
    lSave = IIf(lSave, objData.Update(GetDSN, "stock", "barcode = '" & cBarcode & "'", vaField, vaValue), False)
  Next i
  
  If lSave Then
    objData.Save GetDSN
  Else
    objData.Cancel GetDSN
  End If
  
  FrmPB.EndPB
  MsgBox "Proses Update Stock Selesai, sekarang melakukan update golongan stock, silahkan tunggu sebentar"
  lSave = True
  objData.Start GetDSN
  Set dbData = objData.Browse(GetDSN, "stock", "distinct(kodegolongan) as kodegolongan")
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.EOF
      FrmPB.RunPB
      lSave = IIf(lSave, objData.Update(GetDSN, "golongan", "kodegolongan = '" & GetNull(dbData!kodegolongan) & "'", Array("kodegolongan", "keterangan"), Array(GetNull(dbData!kodegolongan), GetNull(dbData!kodegolongan))), False)
      dbData.MoveNext
    Loop
    FrmPB.EndPB
  End If

  If lSave Then
    objData.Save GetDSN
    MsgBox "Selesai ... !"
  Else
    MsgBox "Maaf, terjadi error/kesalahn dalam memasukkan database. Silahkan hubungi developer program ini"
    objData.Cancel GetDSN
  End If
  CloseWorkSheet
  FinishExcel
End Sub

Private Sub BiSAButton2_Click()
  CommonDialog1.Filter = "Excel File (*.xls)|*.xls| Sophie File (*.dbf)|*.dbf|"
  CommonDialog1.ShowOpen
  If CommonDialog1.FileName <> Trim("") Then
    BiSAButton1.Enabled = True
  End If
End Sub

Private Sub BiSAButton3_Click()
  Label6.Caption = ""
  CommonDialog2.Filter = "Excel File (*.xls)|*.xls| Sophie File (*.dbf)|*.dbf|"
  CommonDialog2.ShowOpen
  Label6.Caption = CommonDialog2.FileTitle
  If CommonDialog2.FileName <> Trim("") Then
    BiSAButton3.Enabled = True
    BiSAButton4.Enabled = True
  End If
End Sub

Private Sub BiSAButton4_Click()
Dim lSave As Boolean
Dim vaField, vaValue
Dim i, j As Integer

  StartExcel
  lSave = True
  
  objData.Start GetDSN
    
  Excel.Workbooks.Close
  Set ExcelWBk = Excel.Workbooks.Open(CommonDialog2.FileName)
  Set ExcelWS = ExcelWBk.Worksheets(1)
  
  MsgBox "Mohon bersabar, tunggu sampai indikator menunjukkan selesai"
  
  'MsgBox "Yakin akan menambahkan " & ExcelWS.Cells.SpecialCells(xlCellTypeLastCell).Row & " Item stock baru??"
  'Loop berikut untuk membaca nilai setiap baris
  'mulai dari baris pertama sampai ketiga
  'dan setiap baris terdiri dari 2 kolom
  
  FrmPB.InitPB 2000
  Dim cNMBRG, cKDBRG, nHRGKON, nCV, nDISCMBR

  For i = 2 To ExcelWS.Cells.SpecialCells(xlCellTypeLastCell).Row
    FrmPB.RunPB
   For j = 1 To 8
    With ExcelWS
      cNMBRG = .Cells(i, 1).Value
      cKDBRG = .Cells(i, 2).Value
      nHRGKON = .Cells(i, 3).Value
      nCV = .Cells(i, 4).Value
      nDISCMBR = .Cells(i, 5).Value
      '####################
      'FORMAT UPDATE STOCK
      '####################
      '1. NMBRG
      '2. KDBRG
      '3. HRGKON
      '4. CV
      '5. DISCMBR
      '####################
    End With
   Next j
   
'    If Trim(cNama) = "" Then
'      Exit For
'    End If
    vaField = Array("NMBRG", "KDBRG", "HRGKON", _
                    "CV", _
                    "DISCMBR")
                    
    vaValue = Array(cNMBRG, cKDBRG, nHRGKON, _
                     nCV, _
                     nDISCMBR)
                     
    lSave = IIf(lSave, objData.Update(GetDSN, "newdbf", "KDBRG = '" & cKDBRG & "'", vaField, vaValue), False)
  Next i
  
  If lSave Then
    objData.Save GetDSN
  Else
    objData.Cancel GetDSN
  End If
  
  FrmPB.EndPB
  If lSave Then
    objData.Save GetDSN
    MsgBox "Selesai ... !"
  Else
    MsgBox "Maaf, terjadi error/kesalahn dalam memasukkan database. Silahkan hubungi developer program ini"
    objData.Cancel GetDSN
  End If
  CloseWorkSheet
  FinishExcel
End Sub






Private Sub BiSAButton6_Click()
  CommonDialog3.Filter = "Excel File (*.xls)|*.xls| Sophie File (*.dbf)|*.dbf|"
  CommonDialog3.ShowOpen
End Sub

Private Sub BiSAButton7_Click()
Dim lSave As Boolean
Dim vaField, vaValue
Dim i, j As Integer

  StartExcel
  lSave = True
  
  objData.Start GetDSN
    
  Excel.Workbooks.Close
  Set ExcelWBk = Excel.Workbooks.Open(CommonDialog3.FileName)
  Set ExcelWS = ExcelWBk.Worksheets(1)
  Set oSheet = ExcelWBk.Sheets.Item(1)
  
  MsgBox "Mohon bersabar, tunggu sampai indikator menunjukkan selesai"
  
  'MsgBox "Yakin akan menambahkan " & ExcelWS.Cells.SpecialCells(xlCellTypeLastCell).Row & " Item stock baru??"
  'Loop berikut untuk membaca nilai setiap baris
  'mulai dari baris pertama sampai ketiga
  'dan setiap baris terdiri dari 2 kolom
  
  FrmPB.InitPB oSheet.UsedRange.Rows.Count
  MsgBox oSheet.UsedRange.Rows.Count
  MsgBox ExcelWS.Cells.SpecialCells(xlCellTypeLastCell).Row
  
  Dim cNama, cBarcode, cKodeGolongan, cKodeSatuan
  Dim nHargaBeli, nHargaJual, nDiskonPenjualan
  Dim cJenis, cBiaya

  For i = 2 To ExcelWS.Cells.SpecialCells(xlCellTypeLastCell).Row
    FrmPB.RunPB
   For j = 1 To 8
    With ExcelWS
      cNama = .Cells(i, 15).Value
      cBarcode = .Cells(i, 14).Value
      cKodeGolongan = .Cells(i, 17).Value
      cKodeSatuan = "PCS"
      nHargaBeli = .Cells(i, 19).Value
      nHargaJual = .Cells(i, 19).Value
      If .Cells(i, 21).Value <> 0 Then
        nDiskonPenjualan = 100 - ((.Cells(i, 22).Value / .Cells(i, 21).Value) * 100) - 3.5
        If nDiskonPenjualan > 30 Then
          nDiskonPenjualan = 30
        End If
      Else
        nDiskonPenjualan = 0
      End If
      '100-((V2/U2)*100)-3.5
      cJenis = 1
      cBiaya = 1
      '####################
      'FORMAT UPDATE STOCK
      '####################
      '1. NAMA
      '2. BARCODE
      '3. HARGA BELI
      '4. -
      '5. DISKON PENJUALAN
      '6. -
      '7. -
      '8. KODE GOLONGAN
      '####################
    End With
   Next j
   
    If Trim(cNama) = "" Then
      Exit For
    End If
    vaField = Array("nama", "barcode", "kodegolongan", _
                    "kodesatuan", _
                    "hargabeli", "hargajual", "jenis", "asbiaya", "diskonpenjualan", "datetime")
                    
    vaValue = Array(StrConv(cNama, vbProperCase), cBarcode, cKodeGolongan, _
                     "PCS", _
                     nHargaBeli, nHargaJual, 1, 2, nDiskonPenjualan, SNow)
                     
    lSave = IIf(lSave, objData.Update(GetDSN, "stock", "barcode = '" & cBarcode & "'", vaField, vaValue), False)
  Next i
  
  If lSave Then
    objData.Save GetDSN
  Else
    objData.Cancel GetDSN
  End If
  
  FrmPB.EndPB
  MsgBox "Proses Update Stock Selesai, sekarang melakukan update golongan stock, silahkan tunggu sebentar"
  lSave = True
  objData.Start GetDSN
  Set dbData = objData.Browse(GetDSN, "stock", "distinct(kodegolongan) as kodegolongan")
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.EOF
      FrmPB.RunPB
      lSave = IIf(lSave, objData.Update(GetDSN, "golongan", "kodegolongan = '" & GetNull(dbData!kodegolongan) & "'", Array("kodegolongan", "keterangan"), Array(GetNull(dbData!kodegolongan), GetNull(dbData!kodegolongan))), False)
      dbData.MoveNext
    Loop
    FrmPB.EndPB
  End If

  If lSave Then
    objData.Save GetDSN
    MsgBox "Selesai ... !"
  Else
    MsgBox "Maaf, terjadi error/kesalahn dalam memasukkan database. Silahkan hubungi developer program ini"
    objData.Cancel GetDSN
  End If
  CloseWorkSheet
  FinishExcel
End Sub

Private Sub Form_Load()
  CenterForm Me
  SetIcon Me.hWnd
  BiSAButton1.Enabled = False
  BiSAButton4.Enabled = False
  Label6.Caption = ""
End Sub
