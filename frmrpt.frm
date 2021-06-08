VERSION 5.00
Object = "{0D6235E7-DBA2-11D1-B5DF-0060976089D0}#1.0#0"; "tdbr6.ocx"
Begin VB.Form FrmRPT 
   Caption         =   "Form2"
   ClientHeight    =   555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1665
   LinkTopic       =   "Form2"
   ScaleHeight     =   555
   ScaleWidth      =   1665
   StartUpPosition =   3  'Windows Default
   Begin TrueDBReports60Ctl.TDBReports Rpt 
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   1005
      Caption         =   "TDBReports1"
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
      ErrorMsgCaption =   ""
      Filtered        =   0   'False
      DataMode        =   1
      DataMember      =   ""
      LinkSequence    =   1
      LinkOrder       =   0
      NameSubstitute  =   ""
      ConnectionString=   ""
      ConnectStringType=   1
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      CursorLocation  =   2
      ConnectionTimeout=   15
      CommandTimeout  =   30
      RecordSource    =   ""
      CursorType      =   3
      CommandType     =   8
      MaxRecords      =   0
      LinkType        =   0
      Master          =   ""
      CallDataRead    =   0   'False
      ConvertNullToEmpty=   -1  'True
      DesignConnection=   -1  'True
      DesignTimeout   =   5
      UnitsOfMeasurement=   4
      Vedit_ShowGrid  =   -1  'True
      Vedit_SnapToGrid=   0   'False
      Vedit_GridUnitWidth=   2.822
      Vedit_GridUnitHeight=   2.822
      Vedit_ShowCellExpressions=   -1  'True
      Norm_rect_left  =   0
      Norm_rect_top   =   0
      Norm_rect_right =   0
      Norm_rect_bottom=   0
      Virgin          =   0   'False
      Profiles.Count  =   1
      Profiles(0).Name=   "PROFILE_0"
      Profiles(0).Active=   -1  'True
      Profiles(0).PreviewNoMinimize=   -1  'True
      Profiles(0).PreviewNoMaximize=   -1  'True
      Profiles(0).PreviewNoResize=   -1  'True
      Profiles(0).PreviewMaximized=   -1  'True
      Profiles(0).PrinterMarginLeft=   10
      Profiles(0).PrinterMarginTop=   10
      Profiles(0).PrinterMarginRight=   10
      Profiles(0).PrinterMarginBottom=   10
      Profiles(0).PrinterMargins_set=   -1  'True
      Profiles(0).PrinterPaperUserSize_set=   -1  'True
   End
End
Attribute VB_Name = "FrmRPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vaFontName
Dim vaHeader As New XArrayDB
Dim vaFooter As New XArrayDB
Dim vaTableHeader As New XArrayDB
Dim vaTableBody As New XArrayDB
Dim vaTableFooter As New XArrayDB
Dim vaTableGroupHeader As New XArrayDB
Dim vaTableGroupFooter As New XArrayDB
Dim vaFieldName As New XArrayDB
Dim vaCallExpresion As New XArrayDB
Dim vaFormat

Dim lNextHeader As Boolean
Dim lNextFooter As Boolean
Dim lNextTableHeader As Boolean
Dim lNextTableBody As Boolean
Dim lNextTableFooter As Boolean
Dim lNextTableGroupHeader As Boolean
Dim lNextTableGroupFooter As Boolean
Dim lNextFieldName As Boolean
Dim lNextCallExpresion As Boolean

Dim nGroupKey As Double
Dim nRec As Double
Dim nOldRec As Double
Dim cOldLeft As String

Public Enum SisFontName
  dbArial = 0
  dbTimesNewRoman = 1
End Enum

Public Enum SisNegatifSymbol
  ssBracket = 0
  ssMinus = 1
End Enum

Public Enum SisRptFormat
  Sis_Rpt_None = 0
  Sis_Rpt_Number = 1
  Sis_Rpt_Number2 = 2
  Sis_Rpt_dd_MM_yyyy = 3
  Sis_Rpt_MM_dd_yyyy = 4
  Sis_Rpt_yyyy_MM_dd = 5
End Enum

Public Enum SisCellBorder
  DB_SINGLE = 0
  DB_DOUBLE = 1
  db_Quart = 2
  db_None = 3
End Enum

Sub Preview(vaArray As XArrayDB, Optional lRecNumber As Boolean = False, Optional lShowPagesNumber As Boolean = True, Optional lLanscape As Boolean = False)
  On Error GoTo PreviewError
  If vaArray.UpperBound(1) >= -1 Then
    nRec = 0
    cOldLeft = ""
    
    Set Rpt.Array = vaArray
    DefineField vaArray
    AddLine
    AddStyle
    
    ' Masukkan Header ke laporan
    SetHeader lShowPagesNumber
    SetTableGroupHeader vaArray, lRecNumber
    SetTableHeader lRecNumber, lLanscape
    SetBody vaArray, lRecNumber
    SetTableGroupFooter lRecNumber
    SetTableFooter lRecNumber
    SetFooter lShowPagesNumber
    
    With Rpt
      .Profiles(0).PrinterLandscape_set = True
      .Profiles(0).PrinterLandscape = lLanscape

      .Refresh
      .PrintPreview
    End With
    ResetValue
  Else
    MsgBox "Data Tidak Ada...!", vbExclamation
  End If
  ResetValue
  Unload Me
  Exit Sub
PreviewError:
  MsgBox "Data Tidak Ada ....!", vbExclamation
  ResetValue
  Unload Me
End Sub

Private Sub ResetValue()
  lNextHeader = False
  lNextFooter = False
  lNextTableHeader = False
  lNextTableBody = False
  lNextTableFooter = False
  lNextTableGroupHeader = False
  lNextTableGroupFooter = False
  lNextFieldName = False
  nGroupKey = -1
End Sub

' Untuk mengambil Nama Field
Private Function GetFieldPosition(ByVal n As Double) As String
Dim cFieldName As String
  If n >= 0 Then
    cFieldName = "F" & n
    If lNextFieldName Then
      If n <= vaFieldName.UpperBound(2) Then
        If Trim(vaFieldName(n, 0)) <> "" Then
          cFieldName = vaFieldName(n, 0)
        End If
      End If
    End If
  End If
  GetFieldPosition = cFieldName
End Function

Private Sub DefineField(vaArray As XArrayDB)
Dim n As Double
Dim cFieldName As String
  With Rpt
    For n = 0 To vaArray.UpperBound(2)
      cFieldName = GetFieldPosition(n)
      
      .Fields.Add (n)
      .Fields(n).Name = cFieldName
      .Fields(n).DisplayName = cFieldName
    Next
  End With
End Sub

Sub AddFields(Optional ByVal cFieldName As String = "")
Dim n As Double
  If Not lNextFooter Then
    vaFooter.ReDim 0, -1, 0, 0
  End If
  lNextFooter = True
  vaFooter.InsertRows vaFooter.UpperBound(1) + 1
  n = vaFooter.UpperBound(1)
  
  vaFooter(n, 0) = cFieldName
End Sub

Sub AddPageFooter(ByVal cCaption, Optional nAlignment As HorzAlignEnum = tdbHalignGeneral, Optional lWidthInpercent As Boolean = True, _
                  Optional nWidth As Single = 0, Optional lNewLine As Boolean = False, _
                  Optional nFontName As SisFontName = dbArial, Optional nFontSize As Single = 8, _
                  Optional lFontBold As Boolean = False, Optional lFontUnderLine As Boolean = False, _
                  Optional lNewSection As Boolean = False, Optional lShowPerPage As Boolean = True, _
                  Optional nSectionType As SectionTypeEnum = tdbPageFooterSect, Optional cCondition As String = "", _
                  Optional lAutoHight As Boolean = True, Optional nCellHeight As Double = 0, _
                  Optional nSpasingBefore As Double = 0, Optional nSpasingAfter As Double = 0)
Dim n As Double
  InitFont
  If Not lNextFooter Then
    vaFooter.ReDim 0, -1, 0, 16
  End If
  lNextFooter = True
  vaFooter.InsertRows vaFooter.UpperBound(1) + 1
  n = vaFooter.UpperBound(1)
  
  vaFooter(n, 0) = cCaption
  vaFooter(n, 1) = nAlignment
  vaFooter(n, 2) = lWidthInpercent
  vaFooter(n, 3) = nWidth
  vaFooter(n, 4) = lNewLine
  vaFooter(n, 5) = vaFontName(nFontName)
  vaFooter(n, 6) = nFontSize
  vaFooter(n, 7) = lFontBold
  vaFooter(n, 8) = lFontUnderLine
  vaFooter(n, 9) = lNewSection
  vaFooter(n, 10) = lShowPerPage
  vaFooter(n, 11) = nSectionType
  vaFooter(n, 12) = cCondition
  vaFooter(n, 13) = lAutoHight
  vaFooter(n, 14) = nCellHeight
  vaFooter(n, 15) = nSpasingBefore
  vaFooter(n, 16) = nSpasingAfter
End Sub

Sub AddPageHeader(ByVal cCaption, Optional nAlignment As HorzAlignEnum = tdbHalignGeneral, Optional lWidthInpercent As Boolean = True, _
              Optional nWidth As Single = 0, Optional lNewLine As Boolean = False, _
              Optional nFontName As SisFontName = dbArial, Optional nFontSize As Single = 8, _
              Optional lFontBold As Boolean = False, Optional lFontUnderLine As Boolean = False, _
              Optional lNewSection As Boolean = False, Optional lShowPerPage As Boolean = True, _
              Optional nSectionType As SectionTypeEnum = tdbPageHeaderSect, Optional cCondition As String = "", _
              Optional lAutoHight As Boolean = True, Optional nCellHeight As Double = 0, _
              Optional nSpasingBefore As Double = 0, Optional nSpasingAfter As Double = 0, _
              Optional nVerticalAlign As VertAlignEnum = tdbValignCenter, _
              Optional nBorderHT As SisCellBorder = db_None, _
              Optional nBorderHI As SisCellBorder = db_None, _
              Optional nBorderHB As SisCellBorder = db_None, _
              Optional nBorderVL As SisCellBorder = db_None, _
              Optional nBorderVI As SisCellBorder = db_None, _
              Optional nBorderVR As SisCellBorder = db_None)
Dim n As Double
  InitFont
  If Not lNextHeader Then
    vaHeader.ReDim 0, -1, 0, 23
  End If
  lNextHeader = True
  vaHeader.InsertRows vaHeader.UpperBound(1) + 1
  n = vaHeader.UpperBound(1)
  
  vaHeader(n, 0) = cCaption
  vaHeader(n, 1) = nAlignment
  vaHeader(n, 2) = lWidthInpercent
  vaHeader(n, 3) = nWidth
  vaHeader(n, 4) = lNewLine
  vaHeader(n, 5) = vaFontName(nFontName)
  vaHeader(n, 6) = nFontSize
  vaHeader(n, 7) = lFontBold
  vaHeader(n, 8) = lFontUnderLine
  vaHeader(n, 9) = lNewSection
  vaHeader(n, 10) = lShowPerPage
  vaHeader(n, 11) = nSectionType
  vaHeader(n, 12) = cCondition
  vaHeader(n, 13) = lAutoHight
  vaHeader(n, 14) = nCellHeight
  vaHeader(n, 15) = nSpasingBefore
  vaHeader(n, 16) = nSpasingAfter
  vaHeader(n, 17) = nVerticalAlign
  vaHeader(n, 18) = nBorderHT
  vaHeader(n, 19) = nBorderHI
  vaHeader(n, 20) = nBorderHB
  vaHeader(n, 21) = nBorderVL
  vaHeader(n, 22) = nBorderVI
  vaHeader(n, 23) = nBorderVR
End Sub

Sub AddTableHeader(Optional ByVal cCaption As String = "", Optional nFieldFormat As SisRptFormat = Sis_Rpt_None, Optional nAlignment As HorzAlignEnum = tdbHalignCenter, Optional lWidthInpercent As Boolean = True, _
                   Optional nWidth As Single = 0, Optional lNewLine As Boolean = False, _
                   Optional nFontName As SisFontName = dbArial, Optional nFontSize As Single = 8, _
                   Optional lFontBold As Boolean = True, Optional lFontUnderLine As Boolean = False, _
                   Optional lNewSection As Boolean = False, _
                   Optional nSectionType As SectionTypeEnum = tdbTableHeaderSect, Optional cCondition As String = "", _
                   Optional nCelMerge As MergeCondEnum = tdbMergeNone, Optional nCellSpan As Double = 1, _
                   Optional lAutoHight As Boolean = True, Optional nCellHeight As Double = 0, _
                   Optional nSpasingBefore As Double = 0, Optional nSpasingAfter As Double = 0, _
                   Optional lVisible As Boolean = True, _
                   Optional nVerticalAlign As VertAlignEnum = tdbValignCenter, _
                   Optional nBorderHT As SisCellBorder = DB_DOUBLE, _
                   Optional nBorderHI As SisCellBorder = DB_DOUBLE, _
                   Optional nBorderHB As SisCellBorder = DB_DOUBLE, _
                   Optional nBorderVL As SisCellBorder = DB_SINGLE, _
                   Optional nBorderVI As SisCellBorder = DB_SINGLE, _
                   Optional nBorderVR As SisCellBorder = DB_SINGLE)
Dim n As Double

  If Not lNextTableHeader Then
    vaTableHeader.ReDim 0, -1, 0, 26
  End If
  lNextTableHeader = True
  vaTableHeader.InsertRows vaTableHeader.UpperBound(1) + 1
  n = vaTableHeader.UpperBound(1)
  
  vaTableHeader(n, 0) = cCaption
  vaTableHeader(n, 1) = nAlignment
  vaTableHeader(n, 2) = lWidthInpercent
  vaTableHeader(n, 3) = nWidth
  vaTableHeader(n, 4) = lNewLine
  vaTableHeader(n, 5) = vaFontName(nFontName)
  vaTableHeader(n, 6) = nFontSize
  vaTableHeader(n, 7) = lFontBold
  vaTableHeader(n, 8) = lFontUnderLine
  vaTableHeader(n, 9) = lNewSection
  vaTableHeader(n, 10) = nSectionType
  vaTableHeader(n, 11) = cCondition
  vaTableHeader(n, 12) = nCelMerge
  vaTableHeader(n, 13) = nCellSpan
  vaTableHeader(n, 14) = vaFormat(nFieldFormat)
  vaTableHeader(n, 15) = lAutoHight
  vaTableHeader(n, 16) = nCellHeight
  vaTableHeader(n, 17) = nSpasingBefore
  vaTableHeader(n, 18) = nSpasingAfter
  vaTableHeader(n, 19) = lVisible
  vaTableHeader(n, 20) = nVerticalAlign
  vaTableHeader(n, 21) = nBorderHT
  vaTableHeader(n, 22) = nBorderHI
  vaTableHeader(n, 23) = nBorderHB
  vaTableHeader(n, 24) = nBorderVL
  vaTableHeader(n, 25) = nBorderVI
  vaTableHeader(n, 26) = nBorderVR

End Sub

Sub AddTableBody(Optional nFieldFormat As SisRptFormat = Sis_Rpt_None, Optional nAlignment As HorzAlignEnum = tdbHalignGeneral, Optional lWidthInpercent As Boolean = True, _
                 Optional nWidth As Single = 0, Optional lNewLine As Boolean = False, _
                 Optional nFontName As SisFontName = dbArial, Optional nFontSize As Single = 8, _
                 Optional lFontBold As Boolean = False, Optional lFontUnderLine As Boolean = False, _
                 Optional nSectionType As SectionTypeEnum = tdbTableBodySect, Optional cCondition As String = "", _
                 Optional nCelMerge As MergeCondEnum = tdbMergeNone, Optional nCellSpan As Double = 1, _
                 Optional lVisible As Boolean = True, Optional nVerticalAlign As VertAlignEnum = tdbValignCenter, _
                 Optional lSuppressIfZero As Boolean = True, Optional nNegatifSimbol As SisNegatifSymbol = ssBracket, _
                 Optional nBorderHT As SisCellBorder = db_Quart, _
                 Optional nBorderHI As SisCellBorder = db_Quart, _
                 Optional nBorderHB As SisCellBorder = DB_DOUBLE, _
                 Optional nBorderVL As SisCellBorder = DB_SINGLE, _
                 Optional nBorderVI As SisCellBorder = DB_SINGLE, _
                 Optional nBorderVR As SisCellBorder = DB_SINGLE)
Dim n As Double
  
  If Not lNextTableBody Then
    vaTableBody.ReDim 0, -1, 0, 22
  End If
  lNextTableBody = True
  vaTableBody.InsertRows vaTableBody.UpperBound(1) + 1
  n = vaTableBody.UpperBound(1)
  
  vaTableBody(n, 0) = vaFormat(nFieldFormat)
  vaTableBody(n, 1) = GetAlignment(nAlignment, nFieldFormat)
  vaTableBody(n, 2) = lWidthInpercent
  vaTableBody(n, 3) = nWidth
  vaTableBody(n, 4) = lNewLine
  vaTableBody(n, 5) = vaFontName(nFontName)
  vaTableBody(n, 6) = nFontSize
  vaTableBody(n, 7) = lFontBold
  vaTableBody(n, 8) = lFontUnderLine
  vaTableBody(n, 9) = nSectionType
  vaTableBody(n, 10) = cCondition
  vaTableBody(n, 11) = nCelMerge
  vaTableBody(n, 12) = nCellSpan
  vaTableBody(n, 13) = lVisible
  vaTableBody(n, 14) = nVerticalAlign
  vaTableBody(n, 15) = lSuppressIfZero
  vaTableBody(n, 16) = nNegatifSimbol
  vaTableBody(n, 17) = nBorderHT
  vaTableBody(n, 18) = nBorderHI
  vaTableBody(n, 19) = nBorderHB
  vaTableBody(n, 20) = nBorderVL
  vaTableBody(n, 21) = nBorderVI
  vaTableBody(n, 22) = nBorderVR
End Sub

Private Function GetAlignment(ByVal nAlignment As HorzAlignEnum, ByVal nFieldFormat As SisRptFormat) As HorzAlignEnum
  If nAlignment = tdbHalignGeneral Then
    If nFieldFormat = Sis_Rpt_Number Or nFieldFormat = Sis_Rpt_Number2 Then
      GetAlignment = tdbHalignRight
    End If
  Else
    GetAlignment = nAlignment
  End If
End Function

Sub AddTableFooter(Optional ByVal cCaption As String = "", Optional nFieldFormat As SisRptFormat = Sis_Rpt_None, Optional nAlignment As HorzAlignEnum = tdbHalignGeneral, Optional lWidthInpercent As Boolean = True, _
                   Optional nWidth As Single = 0, Optional lNewLine As Boolean = False, _
                   Optional nFontName As SisFontName = dbArial, Optional nFontSize As Single = 8, _
                   Optional lFontBold As Boolean = True, Optional lFontUnderLine As Boolean = False, _
                   Optional lNewSection As Boolean = False, _
                   Optional nSectionType As SectionTypeEnum = tdbTableFooterSect, Optional cCondition As String = "", _
                   Optional nCelMerge As MergeCondEnum = tdbMergeNone, Optional nCellSpan As Double = 1, _
                   Optional lAutoHight As Boolean = True, Optional nCellHeight As Double = 0, _
                   Optional nSpasingBefore As Double = 0, Optional nSpasingAfter As Double = 0, _
                   Optional lVisible As Boolean = True, Optional nVerticalAlign As VertAlignEnum = tdbValignCenter, _
                   Optional lSuppressIfZero As Boolean = True, Optional nNegatifSimbol As SisNegatifSymbol = ssBracket)
Dim n As Double

  If Not lNextTableFooter Then
    vaTableFooter.ReDim 0, -1, 0, 22
  End If
  lNextTableFooter = True
  vaTableFooter.InsertRows vaTableFooter.UpperBound(1) + 1
  n = vaTableFooter.UpperBound(1)
  
  vaTableFooter(n, 0) = cCaption
  vaTableFooter(n, 1) = GetAlignment(nAlignment, nFieldFormat)
  vaTableFooter(n, 2) = lWidthInpercent
  vaTableFooter(n, 3) = nWidth
  vaTableFooter(n, 4) = lNewLine
  vaTableFooter(n, 5) = vaFontName(nFontName)
  vaTableFooter(n, 6) = nFontSize
  vaTableFooter(n, 7) = lFontBold
  vaTableFooter(n, 8) = lFontUnderLine
  vaTableFooter(n, 9) = lNewSection
  vaTableFooter(n, 10) = nSectionType
  vaTableFooter(n, 11) = "ISLastRec() " & cCondition
  vaTableFooter(n, 12) = nCelMerge
  vaTableFooter(n, 13) = nCellSpan
  vaTableFooter(n, 14) = vaFormat(nFieldFormat)
  vaTableFooter(n, 15) = lAutoHight
  vaTableFooter(n, 16) = nCellHeight
  vaTableFooter(n, 17) = nSpasingBefore
  vaTableFooter(n, 18) = nSpasingAfter
  vaTableFooter(n, 19) = lVisible
  vaTableFooter(n, 20) = nVerticalAlign
  vaTableFooter(n, 21) = lSuppressIfZero
  vaTableFooter(n, 22) = nNegatifSimbol
End Sub

Sub AddTableGroupFooter(Optional ByVal cCaption As String = "", Optional nFieldFormat As SisRptFormat = Sis_Rpt_None, Optional nAlignment As HorzAlignEnum = tdbHalignGeneral, Optional lWidthInpercent As Boolean = True, _
                        Optional nWidth As Single = 0, Optional lNewLine As Boolean = False, _
                        Optional nFontName As SisFontName = dbArial, Optional nFontSize As Single = 8, _
                        Optional lFontBold As Boolean = True, Optional lFontUnderLine As Boolean = False, _
                        Optional lNewSection As Boolean = False, _
                        Optional nSectionType As SectionTypeEnum = tdbTableFooterSect, Optional cCondition As String = "", _
                        Optional nCelMerge As MergeCondEnum = tdbMergeNone, Optional nCellSpan As Double = 1, _
                        Optional lAutoHight As Boolean = True, Optional nCellHeight As Double = 0, _
                        Optional nSpasingBefore As Double = 0, Optional nSpasingAfter As Double = 0, _
                        Optional lVisible As Boolean = True, Optional nVerticalAlign As VertAlignEnum = tdbValignCenter, _
                        Optional lSuppressIfZero As Boolean = True, Optional nNegatifSimbol As SisNegatifSymbol = ssBracket)
Dim n As Double

  If Not lNextTableGroupFooter Then
    vaTableGroupFooter.ReDim 0, -1, 0, 22
  End If
  lNextTableGroupFooter = True
  vaTableGroupFooter.InsertRows vaTableGroupFooter.UpperBound(1) + 1
  n = vaTableGroupFooter.UpperBound(1)
  
  vaTableGroupFooter(n, 0) = cCaption
  vaTableGroupFooter(n, 1) = GetAlignment(nAlignment, nFieldFormat)
  vaTableGroupFooter(n, 2) = lWidthInpercent
  vaTableGroupFooter(n, 3) = nWidth
  vaTableGroupFooter(n, 4) = lNewLine
  vaTableGroupFooter(n, 5) = vaFontName(nFontName)
  vaTableGroupFooter(n, 6) = nFontSize
  vaTableGroupFooter(n, 7) = lFontBold
  vaTableGroupFooter(n, 8) = lFontUnderLine
  vaTableGroupFooter(n, 9) = lNewSection
  vaTableGroupFooter(n, 10) = nSectionType
  vaTableGroupFooter(n, 11) = cCondition
  vaTableGroupFooter(n, 12) = nCelMerge
  vaTableGroupFooter(n, 13) = nCellSpan
  vaTableGroupFooter(n, 14) = vaFormat(nFieldFormat)
  vaTableGroupFooter(n, 15) = lAutoHight
  vaTableGroupFooter(n, 16) = nCellHeight
  vaTableGroupFooter(n, 17) = nSpasingBefore
  vaTableGroupFooter(n, 18) = nSpasingAfter
  vaTableGroupFooter(n, 19) = lVisible
  vaTableGroupFooter(n, 20) = nVerticalAlign
  vaTableGroupFooter(n, 21) = lSuppressIfZero
  vaTableGroupFooter(n, 22) = nNegatifSimbol
End Sub

Sub AddTableGroupHeader(Optional lGroupKey As Boolean = False, _
                        Optional ByVal cFieldSeparator As String = "", Optional nFieldFormat As SisRptFormat = Sis_Rpt_None, _
                        Optional nAlignment As HorzAlignEnum = tdbHalignLeft, Optional lWidthInpercent As Boolean = True, _
                        Optional nWidth As Single = 0, Optional lNewLine As Boolean = False, _
                        Optional nFontName As SisFontName = dbArial, Optional nFontSize As Single = 8, _
                        Optional lFontBold As Boolean = True, Optional lFontUnderLine As Boolean = False, _
                        Optional lNewSection As Boolean = False, _
                        Optional nSectionType As SectionTypeEnum = tdbTableHeaderSect, Optional cCondition As String = "", _
                        Optional nCelMerge As MergeCondEnum = tdbMergeNone, Optional nCellSpan As Double = 1, _
                        Optional lAutoHight As Boolean = True, Optional nCellHeight As Double = 0, _
                        Optional nSpasingBefore As Double = 0, Optional nSpasingAfter As Double = 0, _
                        Optional lVisible As Boolean = True)
Dim n As Double

  If Not lNextTableGroupHeader Then
    vaTableGroupHeader.ReDim 0, -1, 0, 21
  End If
  lNextTableGroupHeader = True
  vaTableGroupHeader.InsertRows vaTableGroupHeader.UpperBound(1) + 1
  n = vaTableGroupHeader.UpperBound(1)
  
  vaTableGroupHeader(n, 0) = lGroupKey
  vaTableGroupHeader(n, 1) = cFieldSeparator
  vaTableGroupHeader(n, 2) = vaFormat(nFieldFormat)
  vaTableGroupHeader(n, 3) = nAlignment
  vaTableGroupHeader(n, 4) = lWidthInpercent
  vaTableGroupHeader(n, 5) = nWidth
  vaTableGroupHeader(n, 6) = lNewLine
  vaTableGroupHeader(n, 7) = vaFontName(nFontName)
  vaTableGroupHeader(n, 8) = nFontSize
  vaTableGroupHeader(n, 9) = lFontBold
  vaTableGroupHeader(n, 10) = lFontUnderLine
  vaTableGroupHeader(n, 11) = lNewSection
  vaTableGroupHeader(n, 12) = nSectionType
  vaTableGroupHeader(n, 13) = cCondition
  vaTableGroupHeader(n, 14) = nCelMerge
  vaTableGroupHeader(n, 15) = nCellSpan
  vaTableGroupHeader(n, 16) = vaFormat(nFieldFormat)
  vaTableGroupHeader(n, 17) = lAutoHight
  vaTableGroupHeader(n, 18) = nCellHeight
  vaTableGroupHeader(n, 19) = nSpasingBefore
  vaTableGroupHeader(n, 20) = nSpasingAfter
  vaTableGroupHeader(n, 21) = lVisible
  
  If lGroupKey Then
    nGroupKey = n
  End If
End Sub

Private Sub SetPrivateStyle(nSection, nCell, nVerticalAlign, nBorderHT, _
                      nBorderHI, nBorderHB, nBorderVL, nBorderVI, nBorderVR)

  With Rpt.Sections(nSection).Cells(nCell).Style
    .TextVAlign_own = True
    .TextVAlign = nVerticalAlign
    
    .BorderHT_own = True
    .BorderHT = Rpt.Lines(nBorderHT).Name
    
    .BorderHI_own = True
    .BorderHI = Rpt.Lines(nBorderHI).Name
    
    .BorderHB_own = True
    .BorderHB = Rpt.Lines(nBorderHB).Name
    
    .BorderVL_own = True
    .BorderVL = Rpt.Lines(nBorderVL).Name
    
    .BorderVI_own = True
    .BorderVI = Rpt.Lines(nBorderVI).Name
    
    .BorderVR_own = True
    .BorderVR = Rpt.Lines(nBorderVR).Name
  End With
End Sub

Private Sub SetHeader(ByVal lShowPage As Boolean)
Dim n As Double
Dim i As Double
Dim lFirst As Boolean
Dim a As Double

  lFirst = True
  With Rpt

    If lShowPage Then
      .Sections.Add (0)
      .Sections(0).Style = "db_Base"
      .Sections(0).Type = tdbPageHeaderSect
      
      .Sections(0).Cells.Add (0)
      .Sections(0).Cells(0).PrivateStyle = True
      .Sections(0).Cells(0).Style.ParentName = "db_Base"
      .Sections(0).Cells(0).Style.TextAlign_own = True
      .Sections(0).Cells(0).Style.TextAlign = tdbHalignLeft
      .Sections(0).Cells(0).Exp = "'Page : ' & PageNo()"
      
'      .Sections(0).Cells.Add (1)
'      .Sections(0).Cells(1).PrivateStyle = True
'      .Sections(0).Cells(1).Style.ParentName = "db_Base"
'      .Sections(0).Cells(1).Style.TextAlign_own = True
'      .Sections(0).Cells(1).Style.TextAlign = tdbHalignLeft
'      .Sections(0).Cells(1).NewLine = True
'      .Sections(0).Cells(1).Exp = "'" & Format(Now, "dd-MM-yyyy HH:MM:SS") & " " & GetRegistry(reg_UserName) & "'"
    End If
    
    For n = 0 To vaHeader.UpperBound(1)
      If vaHeader(n, 9) Or lFirst Then
        i = .Sections.Count
        .Sections.Add (i)
        .Sections(i).Type = vaHeader(n, 11)
        .Sections(i).Style = "db_Base"
        .Sections(i).SpacingAfter = vaHeader(n, 16)
        .Sections(i).SpacingBefore = vaHeader(n, 15)
        
        If Not vaHeader(n, 10) Then
          .Sections(i).Condition = "PageNo() = 1 "
        End If
        
        .Sections(i).Condition = .Sections(i).Condition & vaHeader(n, 12)
        a = 0
      End If
      lFirst = False
      
      .Sections(i).Cells.Add (a)
      .Sections(i).Cells(a).PrivateStyle = True
      SetPrivateStyle i, a, vaHeader(n, 17), vaHeader(n, 18), _
                      vaHeader(n, 19), vaHeader(n, 20), _
                      vaHeader(n, 21), vaHeader(n, 22), vaHeader(n, 23)
                            
      .Sections(i).Cells(a).Style.TextVAlign_own = True
      .Sections(i).Cells(a).Style.TextVAlign = vaHeader(n, 17)
      
      .Sections(i).Cells(a).AutoHeight = vaHeader(n, 13)
      .Sections(i).Cells(a).Height = vaHeader(n, 14)
      .Sections(i).Cells(a).Exp = "'" & vaHeader(n, 0) & "'"
      .Sections(i).Cells(a).WidthInPercent = vaHeader(n, 2)
      .Sections(i).Cells(a).Width = vaHeader(n, 3)
      .Sections(i).Cells(a).NewLine = vaHeader(n, 4)
      
      .Sections(i).Cells(a).Style.TextAlign_own = True
      .Sections(i).Cells(a).Style.TextAlign = vaHeader(n, 1)
      
      .Sections(i).Cells(a).Style.Font_Name_own = True
      .Sections(i).Cells(a).Style.Font_Name = vaHeader(n, 5)
      
      .Sections(i).Cells(a).Style.Font_Size_own = True
      .Sections(i).Cells(a).Style.Font_Size = vaHeader(n, 6)
      
      .Sections(i).Cells(a).Style.Font_Bold_own = True
      .Sections(i).Cells(a).Style.Font_Bold = vaHeader(n, 7)
      
      .Sections(i).Cells(a).Style.Font_Underline_own = True
      .Sections(i).Cells(a).Style.Font_Underline = vaHeader(n, 8)
      
      a = a + 1
    Next
  End With
  
End Sub

Private Sub SetFooter(Optional ByVal lShowPage As Boolean = False)
Dim n As Double
Dim i As Double
Dim lFirst As Boolean
Dim a As Double

  If lNextFooter Then
    lFirst = True
    With Rpt
      For n = 0 To vaFooter.UpperBound(1)
        If vaFooter(n, 9) Or lFirst Then
          i = .Sections.Count
          .Sections.Add (i)
          .Sections(i).Type = vaFooter(n, 11)
          .Sections(i).Style = "db_Base"
          .Sections(i).SpacingAfter = vaFooter(n, 16)
          .Sections(i).SpacingBefore = vaFooter(n, 15)
          
          If Not vaFooter(n, 10) Then
            .Sections(i).Condition = "ISLastRec() "
          End If
          
          .Sections(i).Condition = .Sections(i).Condition & vaFooter(n, 12)
          a = 0
        End If
        lFirst = False
        
        .Sections(i).Cells.Add (a)
        
        .Sections(i).Cells(a).AutoHeight = vaFooter(n, 13)
        .Sections(i).Cells(a).Height = vaFooter(n, 14)
        .Sections(i).Cells(a).PrivateStyle = True
        .Sections(i).Cells(a).Exp = "'" & vaFooter(n, 0) & "'"
        .Sections(i).Cells(a).WidthInPercent = vaFooter(n, 2)
        .Sections(i).Cells(a).Width = vaFooter(n, 3)
        .Sections(i).Cells(a).NewLine = vaFooter(n, 4)
        
        .Sections(i).Cells(a).Style.TextAlign_own = True
        .Sections(i).Cells(a).Style.TextAlign = vaFooter(n, 1)
        
        .Sections(i).Cells(a).Style.Font_Name_own = True
        .Sections(i).Cells(a).Style.Font_Name = vaFooter(n, 5)
        
        .Sections(i).Cells(a).Style.Font_Size_own = True
        .Sections(i).Cells(a).Style.Font_Size = vaFooter(n, 6)
        
        .Sections(i).Cells(a).Style.Font_Bold_own = True
        .Sections(i).Cells(a).Style.Font_Bold = vaFooter(n, 7)
        
        .Sections(i).Cells(a).Style.Font_Underline_own = True
        .Sections(i).Cells(a).Style.Font_Underline = vaFooter(n, 8)
        
        a = a + 1
      Next
    End With
  End If
End Sub

Private Sub SetBody(vaArray As XArrayDB, lRecNumber As Boolean)
Dim n As Double
Dim i As Double
Dim lFirst As Boolean
Dim a As Double

  If lNextTableBody Then
    lFirst = True
    With Rpt
      For n = 0 To vaArray.UpperBound(2)
        If ISVisible(vaTableBody, n, 13) Then
          If lFirst Then
            i = .Sections.Count
            .Sections.Add (i)
            .Sections(i).Type = tdbTableBodySect
            .Sections(i).Tabulator = "db_Header"
            .Sections(i).Style = "db_TableBody"
            
            If lRecNumber Then
              .Sections(i).Cells.Add 0
              '.Sections(i).Cells(0).Exp = "Sum(1," & IIf(Trim(GetFieldPosition(nGroupKey)) = "", "False", "WillChange(" & GetFieldPosition(nGroupKey) & ")") & ")"
              .Sections(i).Cells(0).CallExpression = True
              .Sections(i).Cells(0).Exp = "Sum(1,False) & '~' & " & IIf(nGroupKey >= 0, "WillChange(" & GetFieldPosition(nGroupKey) & ")", "False")
              a = 1
            Else
              a = 0
            End If
          End If
          lFirst = False
          
          .Sections(i).Cells.Add (a)
          .Sections(i).Cells(a).Exp = GetFieldPosition(n)
          .Sections(i).Cells(a).PrivateStyle = True
          .Sections(i).Cells(a).Style.ParentName = "db_TableBody"
          
          SetPrivateStyle i, a, vaTableBody(n, 14), vaTableBody(n, 17), _
                          vaTableBody(n, 18), vaTableBody(n, 19), _
                          vaTableBody(n, 20), vaTableBody(n, 21), _
                          vaTableBody(n, 22)
          
          If vaTableBody.UpperBound(1) >= n Then
            AddCallExpresion i, a, vaTableBody(n, 15), vaTableBody(n, 16), vaTableBody(n, 0)
            
'            .Sections(i).Cells(a).Format = vaTableBody(n, 0)
            .Sections(i).Cells(a).CallExpression = True
            
            .Sections(i).Cells(a).Style.TextAlign_own = True
            .Sections(i).Cells(a).Style.TextAlign = vaTableBody(n, 1)
            
    
            .Sections(i).Cells(a).Style.Font_Name_own = True
            .Sections(i).Cells(a).Style.Font_Name = vaTableBody(n, 5)
    
            .Sections(i).Cells(a).Style.Font_Size_own = True
            .Sections(i).Cells(a).Style.Font_Size = vaTableBody(n, 6)
    
            .Sections(i).Cells(a).Style.Font_Bold_own = True
            .Sections(i).Cells(a).Style.Font_Bold = vaTableBody(n, 7)
    
            .Sections(i).Cells(a).Style.Font_Underline_own = True
            .Sections(i).Cells(a).Style.Font_Underline = vaTableBody(n, 8)
    
            .Sections(i).Cells(a).CellSpan = vaTableBody(n, 12)
            .Sections(i).Cells(a).Merge = vaTableBody(n, 11)
            
          End If
          a = a + 1
        End If
      Next
    End With
  End If
End Sub

Private Function ISVisible(va As XArrayDB, ByVal n As Double, nVisiblePos) As Boolean
  ISVisible = True
  If va.UpperBound(1) >= n Then
    ISVisible = va(n, nVisiblePos)
  End If
End Function

Private Sub InsertRow(va As XArrayDB, Optional nCol As Double = 0)
Dim n As Double
  va.InsertRows nCol
  For n = 0 To va.UpperBound(2)
    va(nCol, n) = va(nCol + 1, n)
  Next
End Sub

Private Sub SetTableHeader(lRecNumber As Boolean, lLanscape As Boolean)
Dim n As Double
Dim i As Double
Dim lFirst As Boolean
Dim a As Double
Dim lFirstHeader As Boolean

  If lNextTableHeader Then
    lFirst = True
    lFirstHeader = True
    With Rpt
      Do While n <= vaTableHeader.UpperBound(1)
        If ISVisible(vaTableHeader, n, 19) Then
          If vaTableHeader(n, 9) Or lFirst Then
            ' Jika ada tambahan Record Number pertamakali tambah Nomor
            If lRecNumber Then
              InsertRow vaTableHeader, n
              vaTableHeader(n, 0) = "No."
              vaTableHeader(n, 3) = 0
              vaTableHeader(n, 13) = 1
              vaTableHeader(n, 19) = True
              vaTableHeader(n, 12) = MergeCondEnum.tdbMergeOnText
              vaTableHeader(n, 9) = vaTableHeader(n + 1, 9)
              vaTableHeader(n + 1, 9) = False
              vaTableHeader(n, 3) = IIf(lLanscape, 3, 5)
            End If

            i = .Sections.Count
            .Sections.Add (i)
            If lFirstHeader Then
              .Sections(i).Name = "db_Header"
            Else
              .Sections(i).Tabulator = "db_Header"
            End If
            lFirstHeader = False
            .Sections(i).Type = vaTableHeader(n, 10)
            .Sections(i).Style = "db_TableHeader"
            .Sections(i).SpacingBefore = vaTableHeader(n, 17)
            .Sections(i).SpacingAfter = vaTableHeader(n, 18)
            
            .Sections(i).Condition = vaTableHeader(n, 11)
            a = 0
          End If
          lFirst = False
          
          .Sections(i).Cells.Add (a)
          
          .Sections(i).Cells(a).PrivateStyle = True
          .Sections(i).Cells(a).Style.ParentName = "db_TableHeader"
          .Sections(i).Cells(a).Exp = "'" & vaTableHeader(n, 0) & "'"
          .Sections(i).Cells(a).Format = vaTableHeader(n, 14)
          
          SetPrivateStyle i, a, vaTableHeader(n, 20), vaTableHeader(n, 21), _
                          vaTableHeader(n, 22), vaTableHeader(n, 23), _
                          vaTableHeader(n, 24), vaTableHeader(n, 25), _
                          vaTableHeader(n, 26)
          
          .Sections(i).Cells(a).WidthInPercent = vaTableHeader(n, 2)
          .Sections(i).Cells(a).Width = vaTableHeader(n, 3)
          .Sections(i).Cells(a).NewLine = vaTableHeader(n, 4)
          
          .Sections(i).Cells(a).AutoHeight = vaTableHeader(n, 15)
          .Sections(i).Cells(a).Height = vaTableHeader(n, 16)
          
          .Sections(i).Cells(a).Style.TextAlign_own = True
          .Sections(i).Cells(a).Style.TextAlign = vaTableHeader(n, 1)
  
          .Sections(i).Cells(a).Style.Font_Name_own = True
          .Sections(i).Cells(a).Style.Font_Name = vaTableHeader(n, 5)
  
          .Sections(i).Cells(a).Style.Font_Size_own = True
          .Sections(i).Cells(a).Style.Font_Size = vaTableHeader(n, 6)
  
          .Sections(i).Cells(a).Style.Font_Bold_own = True
          .Sections(i).Cells(a).Style.Font_Bold = vaTableHeader(n, 7)
  
          .Sections(i).Cells(a).Style.Font_Underline_own = True
          .Sections(i).Cells(a).Style.Font_Underline = vaTableHeader(n, 8)
  
          .Sections(i).Cells(a).CellSpan = vaTableHeader(n, 13)
          .Sections(i).Cells(a).Merge = vaTableHeader(n, 12)
          a = a + 1
        End If
        
        n = n + 1
      Loop
    End With
  End If
End Sub

Private Sub SetTableGroupHeader(ByVal vaArray As XArrayDB, lRecNumber As Boolean)
Dim n As Double
Dim i As Double
Dim lFirst As Boolean
Dim a As Double
Dim cSeparator As String

  If lNextTableGroupHeader Then
    lFirst = True
    With Rpt
      For n = 0 To vaArray.UpperBound(2)
        If ISVisible(vaTableGroupHeader, n, 21) Then
          If lFirst Then
            ' Tambahkan Section Kosong Untuk Group Header
            ' Hal ini berguna kalau kita ingin memberi group header
            ' dan tidak ada group footer supaya Group berfungsi
            i = .Sections.Count
            .Sections.Add (i)
            .Sections(i).Condition = "HasChanged(" & GetFieldPosition(nGroupKey) & ")"
            .Sections(i).SpacingBefore = 3

            i = .Sections.Count
            .Sections.Add (i)
            .Sections(i).Condition = "HasChanged(" & GetFieldPosition(nGroupKey) & ")"
            .Sections(i).Style = "db_Base"
            .Sections(i).Type = vaTableGroupHeader(n, 12)
            
            a = 0
          End If
          lFirst = False
          
          cSeparator = IIf(vaTableGroupHeader.UpperBound(1) >= n, vaTableGroupHeader(n, 1), "")
          .Sections(i).Cells.Add (a)
          .Sections(i).Cells(a).Exp = "'" & Left(cSeparator, 1) & "' & " & GetFieldPosition(n) & " & '" & Right(cSeparator, 1) & "'"
          .Sections(i).Cells(a).PrivateStyle = True
          .Sections(i).Cells(a).NewLine = vaTableGroupHeader(n, 6)
          
          If vaTableGroupHeader.UpperBound(1) >= n Then
            .Sections(i).Cells(a).WidthInPercent = vaTableGroupHeader(n, 4)
            .Sections(i).Cells(a).Width = vaTableGroupHeader(n, 5)
            
            .Sections(i).Cells(a).Format = vaTableGroupHeader(n, 2)
            
            .Sections(i).Cells(a).Style.TextAlign_own = True
            .Sections(i).Cells(a).Style.TextAlign = vaTableGroupHeader(n, 3)
    
            .Sections(i).Cells(a).Style.Font_Name_own = True
            .Sections(i).Cells(a).Style.Font_Name = vaTableGroupHeader(n, 7)
    
            .Sections(i).Cells(a).Style.Font_Size_own = True
            .Sections(i).Cells(a).Style.Font_Size = vaTableGroupHeader(n, 8)
    
            .Sections(i).Cells(a).Style.Font_Bold_own = True
            .Sections(i).Cells(a).Style.Font_Bold = vaTableGroupHeader(n, 9)
    
            .Sections(i).Cells(a).Style.Font_Underline_own = True
            .Sections(i).Cells(a).Style.Font_Underline = vaTableGroupHeader(n, 10)
    
            .Sections(i).Cells(a).CellSpan = vaTableGroupHeader(n, 15)
            .Sections(i).Cells(a).Merge = vaTableGroupHeader(n, 14)
            
            .Sections(i).Cells(a).Style.MarginTop_own = True
            .Sections(i).Cells(a).Style.MarginTop = tdbLineThickness_12
            
            .Sections(i).Cells(a).Style.MarginBottom_own = True
            .Sections(i).Cells(a).Style.MarginBottom = tdbLineThickness_12
          End If
          a = a + 1
        End If
      Next
    End With
  End If
End Sub

Private Sub SetTableFooter(lRecNumber As Boolean)
Dim n As Double
Dim i As Double
Dim lFirst As Boolean
Dim a As Double

  If lNextTableFooter Then
    lFirst = True
    With Rpt
      n = 0
      Do While n <= vaTableFooter.UpperBound(1)
        If ISVisible(vaTableFooter, n, 19) Then
          If vaTableFooter(n, 9) Or lFirst Then
            If lRecNumber Then
              InsertRow vaTableFooter, n
              If vaTableFooter(n, 13) > 1 Then
                vaTableFooter(n, 0) = vaTableFooter(n + 1, 0)
                vaTableFooter(n, 13) = vaTableFooter(n + 1, 13) + 1
              Else
                vaTableFooter(n, 0) = ""
              End If
              vaTableFooter(n, 19) = True
            End If
            
            i = .Sections.Count
            If lFirst Then
              .Sections.Add (i)
              .Sections(i).Condition = vaTableFooter(n, 11)
              i = i + 1
            End If
            
            .Sections.Add (i)
            .Sections(i).Tabulator = "db_Header"
            .Sections(i).Type = vaTableFooter(n, 10)
            .Sections(i).Style = "db_TableHeader"
            .Sections(i).SpacingBefore = vaTableFooter(n, 17)
            .Sections(i).SpacingAfter = vaTableFooter(n, 18)
            
            .Sections(i).Condition = vaTableFooter(n, 11)
            a = 0
          End If
          lFirst = False
          
          .Sections(i).Cells.Add (a)
          AddCallExpresion i, a, vaTableFooter(n, 21), vaTableFooter(n, 22), vaTableFooter(n, 14)
          
          .Sections(i).Cells(a).PrivateStyle = True
          .Sections(i).Cells(a).Style.ParentName = "db_TableHeader"
          .Sections(i).Cells(a).Exp = GetFooterExp(vaTableFooter(n, 0), n, lRecNumber)
          .Sections(i).Cells(a).CallExpression = True
          
          .Sections(i).Cells(a).WidthInPercent = vaTableFooter(n, 2)
          .Sections(i).Cells(a).Width = vaTableFooter(n, 3)
          .Sections(i).Cells(a).NewLine = vaTableFooter(n, 4)
          
          .Sections(i).Cells(a).AutoHeight = vaTableFooter(n, 15)
          .Sections(i).Cells(a).Height = vaTableFooter(n, 16)
          
          .Sections(i).Cells(a).Style.TextAlign_own = True
          .Sections(i).Cells(a).Style.TextAlign = vaTableFooter(n, 1)
  
          .Sections(i).Cells(a).Style.Font_Name_own = True
          .Sections(i).Cells(a).Style.Font_Name = vaTableFooter(n, 5)
  
          .Sections(i).Cells(a).Style.Font_Size_own = True
          .Sections(i).Cells(a).Style.Font_Size = vaTableFooter(n, 6)
  
          .Sections(i).Cells(a).Style.Font_Bold_own = True
          .Sections(i).Cells(a).Style.Font_Bold = vaTableFooter(n, 7)
  
          .Sections(i).Cells(a).Style.Font_Underline_own = True
          
          
          .Sections(i).Cells(a).Style.Font_Underline = vaTableFooter(n, 8)
  
          .Sections(i).Cells(a).CellSpan = vaTableFooter(n, 13)
          .Sections(i).Cells(a).Merge = vaTableFooter(n, 12)
          a = a + 1
        End If
        n = n + 1
      Loop
    End With
  End If
End Sub

Private Sub SetTableGroupFooter(lRecNumber As Boolean)
Dim n As Double
Dim i As Double
Dim lFirst As Boolean
Dim a As Double

  If lNextTableGroupFooter Then
    lFirst = True
    With Rpt
      n = 0
      Do While n <= vaTableGroupFooter.UpperBound(1)
        If ISVisible(vaTableGroupFooter, n, 19) Then
          If vaTableGroupFooter(n, 9) Or lFirst Then
            If lRecNumber Then
              InsertRow vaTableGroupFooter, n
              If vaTableGroupFooter(n, 13) > 1 Then
                vaTableGroupFooter(n, 0) = vaTableGroupFooter(n + 1, 0)
                vaTableGroupFooter(n, 13) = vaTableGroupFooter(n + 1, 13) + 1
              Else
                vaTableGroupFooter(n, 0) = ""
              End If
              vaTableGroupFooter(n, 19) = True
              vaTableGroupFooter(n, 9) = vaTableGroupFooter(n + 1, 9)
              vaTableGroupFooter(n + 1, 9) = False
            End If

            i = .Sections.Count
            If lFirst Then
              .Sections.Add (i)
              .Sections(i).Condition = "WillChange(" & GetFieldPosition(nGroupKey) & ")"
              i = i + 1
            End If
            
            .Sections.Add (i)
            .Sections(i).Tabulator = "db_Header"
            .Sections(i).Type = vaTableGroupFooter(n, 10)
            .Sections(i).Style = "db_TableHeader"
            .Sections(i).SpacingBefore = vaTableGroupFooter(n, 17)
            .Sections(i).SpacingAfter = vaTableGroupFooter(n, 18)
            
            .Sections(i).Condition = "WillChange(" & GetFieldPosition(nGroupKey) & ") " & vaTableGroupFooter(n, 11)
            a = 0
          End If
          lFirst = False
          
          .Sections(i).Cells.Add (a)
          AddCallExpresion i, a, vaTableGroupFooter(n, 21), vaTableGroupFooter(n, 22), vaTableGroupFooter(n, 14)
                      
          .Sections(i).Cells(a).PrivateStyle = True
          .Sections(i).Cells(a).Style.ParentName = "db_TableHeader"
          .Sections(i).Cells(a).Exp = GetFooterExp(vaTableGroupFooter(n, 0), n, lRecNumber, True)
          
          .Sections(i).Cells(a).CallExpression = True
          '.Sections(i).Cells(a).Format = vaTableGroupFooter(n, 14)
          
          .Sections(i).Cells(a).WidthInPercent = vaTableGroupFooter(n, 2)
          .Sections(i).Cells(a).Width = vaTableGroupFooter(n, 3)
          .Sections(i).Cells(a).NewLine = vaTableGroupFooter(n, 4)
          
          .Sections(i).Cells(a).AutoHeight = vaTableGroupFooter(n, 15)
          .Sections(i).Cells(a).Height = vaTableGroupFooter(n, 16)
          
          .Sections(i).Cells(a).Style.TextAlign_own = True
          .Sections(i).Cells(a).Style.TextAlign = vaTableGroupFooter(n, 1)
  
          .Sections(i).Cells(a).Style.Font_Name_own = True
          .Sections(i).Cells(a).Style.Font_Name = vaTableGroupFooter(n, 5)
  
          .Sections(i).Cells(a).Style.Font_Size_own = True
          .Sections(i).Cells(a).Style.Font_Size = vaTableGroupFooter(n, 6)
  
          .Sections(i).Cells(a).Style.Font_Bold_own = True
          .Sections(i).Cells(a).Style.Font_Bold = vaTableGroupFooter(n, 7)
  
          .Sections(i).Cells(a).Style.Font_Underline_own = True
          .Sections(i).Cells(a).Style.Font_Underline = vaTableGroupFooter(n, 8)
  
          .Sections(i).Cells(a).CellSpan = vaTableGroupFooter(n, 13)
          .Sections(i).Cells(a).Merge = vaTableGroupFooter(n, 12)
          a = a + 1
        End If
        n = n + 1
      Loop
    End With
  End If
End Sub

Function GetFieldName(ByVal cFieldExp As String) As String
Dim nStart As Double
Dim nEnd As Double

  nStart = InStr(1, cFieldExp, "&")
  If nStart > 0 Then
    nEnd = InStr(nStart + 1, cFieldExp, "&")
  End If
  If nStart > 0 Or nEnd > 0 Then
    GetFieldName = Mid(cFieldExp, nStart + 1, IIf(nEnd = 0, Len(cFieldExp) - nStart + 1, nEnd))
  End If
End Function

Private Function GetFooterExp(ByVal cFieldExp As String, ByVal nCol, ByVal lRecNumber As Boolean, Optional ByVal lGroupFooter As Boolean = False) As String
Dim n As Double
Dim cF As String
Dim cRemark As String
  ' Jika menampilkan Record number berarti colom - 1
  nCol = nCol - IIf(lRecNumber, 1, 0)
  n = InStr(1, cFieldExp, "&")
  If n > 0 Then
    cF = GetFieldName(cFieldExp)
    Select Case UCase(cF)
      Case "SUM"
        ' Jika Group footer maka perhitungan menggunakan Reset
        If lGroupFooter Then
          GetFooterExp = "Sum(" & GetFieldPosition(nCol) & ")"
        Else
          GetFooterExp = "Sum(" & GetFieldPosition(nCol) & ",False)"
        End If
      Case "FIELD"
        GetFooterExp = GetFieldName(nCol)
    End Select
  Else
    GetFooterExp = "'" & cFieldExp & "'"
  End If
End Function

Private Sub AddStyle()
Dim n As Single
  With Rpt
    .Styles.Add (n)
    .Styles(n).Name = "db_Base"
    .Styles(n).HasBorders = True
    
    .Styles(n).TextVAlign_own = True
    .Styles(n).TextVAlign = tdbValignCenter
    
    .Styles(n).TextWrap_own = True
    .Styles(n).TextWrap = False
    
    .Styles(n).Font_Name_own = True
    .Styles(n).Font_Name = "Arial"
    
    .Styles(n).Font_Size_own = True
    .Styles(n).Font_Size = 8
    
    .Styles(n).MarginBottom_own = True
    .Styles(n).MarginBottom = tdbLineThickness_1
    
    .Styles(n).MarginTop_own = True
    .Styles(n).MarginTop = tdbLineThickness_1
    
    ' Tabah Style Untuk Header
    n = 1
    .Styles.Add (n)
    .Styles(n).Name = "db_TableHeader"
    .Styles(n).ParentName = "db_Base"
    
    .Styles(n).TextWrap_own = True
    .Styles(n).TextWrap = True
    
    .Styles(n).BorderHB_own = True
    .Styles(n).BorderHB = "db_Double"
    
    .Styles(n).BorderHI_own = True
    .Styles(n).BorderHI = "db_Double"
    
    .Styles(n).BorderHT_own = True
    .Styles(n).BorderHT = "db_Double"
    
    .Styles(n).BorderVI_own = True
    .Styles(n).BorderVI = "db_Single"
    
    .Styles(n).BorderVL_own = True
    .Styles(n).BorderVL = "db_Single"
    
    .Styles(n).BorderVR_own = True
    .Styles(n).BorderVR = "db_Single"
   
    .Styles(n).MarginBottom_own = True
    .Styles(n).MarginBottom = tdbLineThickness_2_14
    
    .Styles(n).MarginTop_own = True
    .Styles(n).MarginTop = tdbLineThickness_2_14

    ' Tabah Style Untuk Body
    n = 1
    .Styles.Add (n)
    .Styles(n).Name = "db_TableBody"
    .Styles(n).ParentName = "db_Base"
    
    .Styles(n).BorderHB_own = True
    .Styles(n).BorderHB = "db_Double"
    
    .Styles(n).BorderHI_own = True
    .Styles(n).BorderHI = "db_Quart"
    
    .Styles(n).BorderHT_own = True
    .Styles(n).BorderHT = "db_Quart"
    
    .Styles(n).BorderVI_own = True
    .Styles(n).BorderVI = "db_Single"
    
    .Styles(n).BorderVL_own = True
    .Styles(n).BorderVL = "db_Single"
    
    .Styles(n).BorderVR_own = True
    .Styles(n).BorderVR = "db_Single"
   
    .Styles(n).MarginBottom_own = True
    .Styles(n).MarginBottom = tdbLineThickness_2_14
    
    .Styles(n).MarginTop_own = True
    .Styles(n).MarginTop = tdbLineThickness_2_14

  End With
End Sub

Private Sub AddLine()
  With Rpt
    .Lines.Add (0)
    .Lines(0).Name = "db_Single"
    .Lines(0).Thickness = tdbLineThickness_1
    
    .Lines.Add (1)
    .Lines(1).Name = "db_Double"
    .Lines(1).Thickness = tdbLineThickness_1_12
    
    .Lines.Add (2)
    .Lines(2).Name = "db_Quart"
    .Lines(2).Thickness = tdbLineThickness_14
    
    .Lines.Add (3)
    .Lines(3).Name = "db_None"
    .Lines(3).Thickness = tdbLineThicknessNone
  End With
End Sub

Private Sub InitFont()
  vaFontName = Array("Arial", "Times New Roman")
  
  vaFormat = Array("", "###,###,###,###,###,##0", _
                 "###,###,###,###,###,##0.00", _
                 "dd-MM-yyyy", "MM-dd-YYYY", _
                 "yyyy-MM-dd")
End Sub

Private Sub Form_Initialize()
  nGroupKey = -1
End Sub

Private Sub Rpt_CellExpression(ByVal Section As Integer, ByVal Cell As Integer, Value As Variant)
Dim n As Double
Dim cStatus As String
Dim cLeft As String
Dim nFound As Double
  
  nFound = vaCallExpresion.Find(0, 0, GetSection(Section, Cell))
  If nFound >= 0 Then
    If vaCallExpresion(nFound, 1) = True Then
      If Value = 0 Then
        Value = ""
      Else
        If vaCallExpresion(n, 2) = SisNegatifSymbol.ssBracket And Value < 0 Then
          Value = "(" & Format(-Value, vaCallExpresion(nFound, 3)) & ")"
        ElseIf vaCallExpresion(nFound, 3) <> "" Then
          Value = Format(Value, vaCallExpresion(nFound, 3))
        End If
      End If
    End If
  Else
    n = InStr(1, Value, "~")
    If n <> 0 Then
      cStatus = Mid(Value, n + 1)
      cLeft = RTrim(Left(Value, n - 1))
      If cLeft <> cOldLeft Then
        nRec = nRec + 1
      End If

      cOldLeft = cLeft
      Value = IIf(nRec = 0, nOldRec, nRec)
      nOldRec = nRec
      If cStatus = "True" Then
        nRec = 0
      End If
    End If
  End If
End Sub

Private Sub AddCallExpresion(ByVal nSection As Double, ByVal nCell As Double, ByVal lSuppressIfZero As Boolean, _
                             ByVal nNegatifSimbol, ByVal cFormat As String)
Dim n As Double
  If Not lNextCallExpresion Then
    lNextCallExpresion = True
    vaCallExpresion.ReDim 0, -1, 0, 3
  End If

  ' Tambah pada vaCallExpresion Sifatnya
  vaCallExpresion.InsertRows vaCallExpresion.UpperBound(1) + 1
  n = vaCallExpresion.UpperBound(1)
  vaCallExpresion(n, 0) = GetSection(nSection, nCell)
  vaCallExpresion(n, 1) = lSuppressIfZero
  vaCallExpresion(n, 2) = nNegatifSimbol
  vaCallExpresion(n, 3) = cFormat
End Sub

Private Function GetSection(ByVal nSection As Double, ByVal nCell As Double)
  GetSection = Trim(Format(nSection, "##########")) & "~~" & Trim(Format(nCell, "##########"))
End Function
