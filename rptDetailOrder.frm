VERSION 5.00
Object = "{0D6235E7-DBA2-11D1-B5DF-0060976089D0}#1.0#0"; "tdbr6.ocx"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptDetailOrder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detail Order"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   7020
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   3555
      Left            =   15
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   6271
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   195
         Left            =   2310
         TabIndex        =   5
         Top             =   915
         Width           =   240
      End
      Begin BiSAFramProject.BiSAFrame BiSAFrame4 
         Height          =   465
         Left            =   2265
         Top             =   1245
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   820
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         Begin VB.OptionButton optKodeStock 
            Caption         =   "Kode Index"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   120
            TabIndex        =   1
            Top             =   105
            Width           =   1230
         End
         Begin VB.OptionButton optKodeStock 
            Caption         =   "Barcode"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   1380
            TabIndex        =   0
            Top             =   105
            Width           =   1035
         End
      End
      Begin BiSAFramProject.BiSAFrame BiSAFrame2 
         Height          =   510
         Left            =   2265
         Top             =   1695
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   900
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         Begin VB.OptionButton optTunai 
            Caption         =   "Semua"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   1905
            TabIndex        =   4
            Top             =   150
            Width           =   975
         End
         Begin VB.OptionButton optTunai 
            Caption         =   "Kredit"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   1020
            TabIndex        =   3
            Top             =   150
            Width           =   780
         End
         Begin VB.OptionButton optTunai 
            Caption         =   "Tunai"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   105
            TabIndex        =   2
            Top             =   150
            Width           =   840
         End
      End
      Begin BiSATextBoxProject.BiSABrowse cCustomer 
         Height          =   330
         Left            =   2535
         TabIndex        =   6
         Top             =   915
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   582
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Verdana"
         Button          =   -1  'True
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
      Begin BiSADateProject.BiSADate dDate 
         Height          =   330
         Index           =   0
         Left            =   225
         TabIndex        =   7
         Top             =   510
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   582
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
         Caption         =   "ANTARA TANGGAL"
         CaptionWidth    =   2000
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
      Begin BiSADateProject.BiSADate dDate 
         Height          =   330
         Index           =   1
         Left            =   3855
         TabIndex        =   8
         Top             =   510
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   582
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
         Caption         =   "S.D"
         CaptionWidth    =   500
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
      Begin TrueDBReports60Ctl.TDBReports tdb 
         Height          =   570
         Left            =   5610
         TabIndex        =   9
         Top             =   -120
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   1005
         Caption         =   "Penjualan"
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
         CursorLocation  =   3
         ConnectionTimeout=   15
         CommandTimeout  =   30
         RecordSource    =   ""
         CursorType      =   1
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
         Parameters.Count=   10
         Parameters(0).Name=   "TGL1"
         Parameters(1).Name=   "TGL2"
         Parameters(2).Name=   "TOTALJUMLAH"
         Parameters(2).Type=   5
         Parameters(3).Name=   "TOTALQTY"
         Parameters(3).Type=   5
         Parameters(4).Name=   "TOTALDISCOUNT1"
         Parameters(4).Type=   5
         Parameters(5).Name=   "TOTALDISCOUNT2"
         Parameters(5).Type=   5
         Parameters(6).Name=   "TOTALPAJAK"
         Parameters(6).Type=   5
         Parameters(7).Name=   "GRANDTOTAL"
         Parameters(7).Type=   5
         Parameters(8).Name=   "TUNAIREF"
         Parameters(9).Name=   "PIUTANGREF"
         Fields.Count    =   17
         Fields(0).Name  =   "faktur"
         Fields(0).DisplayName=   "faktur"
         Fields(1).Name  =   "supplier"
         Fields(1).DisplayName=   "supplier"
         Fields(2).Name  =   "tgl"
         Fields(2).DisplayName=   "tgl"
         Fields(2).Type  =   7
         Fields(3).Name  =   "kode"
         Fields(3).DisplayName=   "kode"
         Fields(4).Name  =   "namabarang"
         Fields(4).DisplayName=   "namabarang"
         Fields(5).Name  =   "qty"
         Fields(5).DisplayName=   "qty"
         Fields(5).Type  =   5
         Fields(6).Name  =   "satuan"
         Fields(6).DisplayName=   "satuan"
         Fields(7).Name  =   "harga"
         Fields(7).DisplayName=   "harga"
         Fields(7).Type  =   5
         Fields(8).Name  =   "jumlah"
         Fields(8).DisplayName=   "jumlah"
         Fields(8).Type  =   5
         Fields(9).Name  =   "subtotal"
         Fields(9).DisplayName=   "subtotal"
         Fields(9).Type  =   5
         Fields(10).Name =   "discount"
         Fields(10).DisplayName=   "discount"
         Fields(10).Type =   5
         Fields(11).Name =   "discount2"
         Fields(11).DisplayName=   "discount2"
         Fields(11).Type =   5
         Fields(12).Name =   "pajak"
         Fields(12).DisplayName=   "pajak"
         Fields(12).Type =   5
         Fields(13).Name =   "total"
         Fields(13).DisplayName=   "total"
         Fields(13).Type =   5
         Fields(14).Name =   "disc"
         Fields(14).DisplayName=   "disc"
         Fields(14).Type =   5
         Fields(15).Name =   "tunai"
         Fields(15).DisplayName=   "tunai"
         Fields(16).Name =   "piutang"
         Fields(16).DisplayName=   "piutang"
         Sections.Count  =   6
         Sections(0).Name=   "SECTION_1"
         Sections(0).Type=   1
         Sections(0).Cells.Count=   6
         Sections(0).Cells(0).Name=   "CELL_0"
         Sections(0).Cells(0).Exp=   """Hal : "" & PageNo()"
         Sections(0).Cells(0).PrivateStyle=   -1  'True
         Sections(0).Cells(0).Style.Name=   "<private>"
         Sections(0).Cells(0).Style.ParentName=   "<null>"
         Sections(0).Cells(0).Style.Font_Name=   "Times New Roman"
         Sections(0).Cells(0).Style.Font_Size=   10
         Sections(0).Cells(0).Style.Font_Bold=   0   'False
         Sections(0).Cells(0).Style.Font_Italic=   0   'False
         Sections(0).Cells(0).Style.Font_Underline=   0   'False
         Sections(0).Cells(0).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(0).Style.Font_Charset=   1
         Sections(0).Cells(0).Style.TextAlign=   2
         Sections(0).Cells(0).Style.TextVAlign=   0
         Sections(0).Cells(0).Style.TextWrap=   -1  'True
         Sections(0).Cells(0).Style.ForeColor=   0
         Sections(0).Cells(0).Style.BackColor=   16777215
         Sections(0).Cells(0).Style.NoFill=   -1  'True
         Sections(0).Cells(0).Style.BackPicFile=   ""
         Sections(0).Cells(0).Style.ForePicFile=   ""
         Sections(0).Cells(0).Style.BackPicVertPlacement=   0
         Sections(0).Cells(0).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(0).Style.ForePicPlacement=   0
         Sections(0).Cells(0).Style.ForePicDrawMode=   0
         Sections(0).Cells(0).Style.MarginLeft=   6
         Sections(0).Cells(0).Style.MarginTop=   6
         Sections(0).Cells(0).Style.MarginRight=   6
         Sections(0).Cells(0).Style.MarginBottom=   6
         Sections(0).Cells(0).Style.HasBorders=   -1  'True
         Sections(0).Cells(0).Style.BorderHT=   ""
         Sections(0).Cells(0).Style.BorderHI=   ""
         Sections(0).Cells(0).Style.BorderHB=   ""
         Sections(0).Cells(0).Style.BorderVL=   ""
         Sections(0).Cells(0).Style.BorderVI=   ""
         Sections(0).Cells(0).Style.BorderVR=   ""
         Sections(0).Cells(0).Style.NoClipping=   0   'False
         Sections(0).Cells(0).Style.RTF=   0   'False
         Sections(0).Cells(0).Style.fprops=   1
         Sections(0).Cells(1).Name=   "CELL_1"
         Sections(0).Cells(1).Exp=   """ """
         Sections(0).Cells(1).NewLine=   -1  'True
         Sections(0).Cells(2).Name=   "CELL_2"
         Sections(0).Cells(2).Exp=   """LAPORAN DETAIL ORDER"""
         Sections(0).Cells(2).NewLine=   -1  'True
         Sections(0).Cells(2).PrivateStyle=   -1  'True
         Sections(0).Cells(2).Style.Name=   "<private>"
         Sections(0).Cells(2).Style.ParentName=   "<null>"
         Sections(0).Cells(2).Style.Font_Name=   "Verdana"
         Sections(0).Cells(2).Style.Font_Size=   12
         Sections(0).Cells(2).Style.Font_Bold=   -1  'True
         Sections(0).Cells(2).Style.Font_Italic=   0   'False
         Sections(0).Cells(2).Style.Font_Underline=   0   'False
         Sections(0).Cells(2).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(2).Style.Font_Charset=   1
         Sections(0).Cells(2).Style.TextAlign=   1
         Sections(0).Cells(2).Style.TextVAlign=   0
         Sections(0).Cells(2).Style.TextWrap=   -1  'True
         Sections(0).Cells(2).Style.ForeColor=   0
         Sections(0).Cells(2).Style.BackColor=   16777215
         Sections(0).Cells(2).Style.NoFill=   -1  'True
         Sections(0).Cells(2).Style.BackPicFile=   ""
         Sections(0).Cells(2).Style.ForePicFile=   ""
         Sections(0).Cells(2).Style.BackPicVertPlacement=   0
         Sections(0).Cells(2).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(2).Style.ForePicPlacement=   0
         Sections(0).Cells(2).Style.ForePicDrawMode=   0
         Sections(0).Cells(2).Style.MarginLeft=   6
         Sections(0).Cells(2).Style.MarginTop=   6
         Sections(0).Cells(2).Style.MarginRight=   6
         Sections(0).Cells(2).Style.MarginBottom=   6
         Sections(0).Cells(2).Style.HasBorders=   -1  'True
         Sections(0).Cells(2).Style.BorderHT=   ""
         Sections(0).Cells(2).Style.BorderHI=   ""
         Sections(0).Cells(2).Style.BorderHB=   ""
         Sections(0).Cells(2).Style.BorderVL=   ""
         Sections(0).Cells(2).Style.BorderVI=   ""
         Sections(0).Cells(2).Style.BorderVR=   ""
         Sections(0).Cells(2).Style.NoClipping=   0   'False
         Sections(0).Cells(2).Style.RTF=   0   'False
         Sections(0).Cells(2).Style.fprops=   23068673
         Sections(0).Cells(3).Name=   "CELL_3"
         Sections(0).Cells(3).Exp=   """Antara Tanggal : "" & TGL1 & "" S.D "" & TGL2"
         Sections(0).Cells(3).NewLine=   -1  'True
         Sections(0).Cells(3).PrivateStyle=   -1  'True
         Sections(0).Cells(3).Format=   "dd-Mm-yyyy"
         Sections(0).Cells(3).Style.Name=   "<private>"
         Sections(0).Cells(3).Style.ParentName=   "<null>"
         Sections(0).Cells(3).Style.Font_Name=   "Verdana"
         Sections(0).Cells(3).Style.Font_Size=   9.75
         Sections(0).Cells(3).Style.Font_Bold=   -1  'True
         Sections(0).Cells(3).Style.Font_Italic=   0   'False
         Sections(0).Cells(3).Style.Font_Underline=   0   'False
         Sections(0).Cells(3).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(3).Style.Font_Charset=   1
         Sections(0).Cells(3).Style.TextAlign=   1
         Sections(0).Cells(3).Style.TextVAlign=   0
         Sections(0).Cells(3).Style.TextWrap=   -1  'True
         Sections(0).Cells(3).Style.ForeColor=   0
         Sections(0).Cells(3).Style.BackColor=   16777215
         Sections(0).Cells(3).Style.NoFill=   -1  'True
         Sections(0).Cells(3).Style.BackPicFile=   ""
         Sections(0).Cells(3).Style.ForePicFile=   ""
         Sections(0).Cells(3).Style.BackPicVertPlacement=   0
         Sections(0).Cells(3).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(3).Style.ForePicPlacement=   0
         Sections(0).Cells(3).Style.ForePicDrawMode=   0
         Sections(0).Cells(3).Style.MarginLeft=   6
         Sections(0).Cells(3).Style.MarginTop=   6
         Sections(0).Cells(3).Style.MarginRight=   6
         Sections(0).Cells(3).Style.MarginBottom=   6
         Sections(0).Cells(3).Style.HasBorders=   -1  'True
         Sections(0).Cells(3).Style.BorderHT=   ""
         Sections(0).Cells(3).Style.BorderHI=   ""
         Sections(0).Cells(3).Style.BorderHB=   ""
         Sections(0).Cells(3).Style.BorderVL=   ""
         Sections(0).Cells(3).Style.BorderVI=   ""
         Sections(0).Cells(3).Style.BorderVR=   ""
         Sections(0).Cells(3).Style.NoClipping=   0   'False
         Sections(0).Cells(3).Style.RTF=   0   'False
         Sections(0).Cells(3).Style.fprops=   23068673
         Sections(0).Cells(4).Name=   "CELL_4"
         Sections(0).Cells(4).Exp=   """ """
         Sections(0).Cells(4).NewLine=   -1  'True
         Sections(0).Cells(5).Name=   "CELL_5"
         Sections(0).Cells(5).Exp=   """ """
         Sections(0).Cells(5).NewLine=   -1  'True
         Sections(1).Name=   "SECTION_5"
         Sections(1).Condition=   "HasChanged(faktur)"
         Sections(1).StyleExp=   "'tdb_Base'"
         Sections(1).KeepWithNext=   2
         Sections(1).Cells.Count=   6
         Sections(1).Cells(0).Name=   "CELL_0"
         Sections(1).Cells(0).Exp=   """NO. ORDER"""
         Sections(1).Cells(0).Width=   20
         Sections(1).Cells(0).Height=   4
         Sections(1).Cells(0).AutoHeight=   0   'False
         Sections(1).Cells(0).PrivateStyle=   -1  'True
         Sections(1).Cells(0).Style.Name=   "<private>"
         Sections(1).Cells(0).Style.ParentName=   "tdb_Base"
         Sections(1).Cells(0).Style.Font_Name=   "Arial"
         Sections(1).Cells(0).Style.Font_Size=   8.25
         Sections(1).Cells(0).Style.Font_Bold=   -1  'True
         Sections(1).Cells(0).Style.Font_Italic=   0   'False
         Sections(1).Cells(0).Style.Font_Underline=   0   'False
         Sections(1).Cells(0).Style.Font_Strikeout=   0   'False
         Sections(1).Cells(0).Style.Font_Charset=   0
         Sections(1).Cells(0).Style.TextAlign=   0
         Sections(1).Cells(0).Style.TextVAlign=   1
         Sections(1).Cells(0).Style.TextWrap=   -1  'True
         Sections(1).Cells(0).Style.ForeColor=   0
         Sections(1).Cells(0).Style.BackColor=   16777215
         Sections(1).Cells(0).Style.NoFill=   -1  'True
         Sections(1).Cells(0).Style.BackPicFile=   ""
         Sections(1).Cells(0).Style.ForePicFile=   ""
         Sections(1).Cells(0).Style.BackPicVertPlacement=   0
         Sections(1).Cells(0).Style.BackPicHorzPlacement=   0
         Sections(1).Cells(0).Style.ForePicPlacement=   0
         Sections(1).Cells(0).Style.ForePicDrawMode=   0
         Sections(1).Cells(0).Style.MarginLeft=   6
         Sections(1).Cells(0).Style.MarginTop=   6
         Sections(1).Cells(0).Style.MarginRight=   6
         Sections(1).Cells(0).Style.MarginBottom=   6
         Sections(1).Cells(0).Style.HasBorders=   -1  'True
         Sections(1).Cells(0).Style.BorderHT=   ""
         Sections(1).Cells(0).Style.BorderHI=   ""
         Sections(1).Cells(0).Style.BorderHB=   ""
         Sections(1).Cells(0).Style.BorderVL=   ""
         Sections(1).Cells(0).Style.BorderVI=   ""
         Sections(1).Cells(0).Style.BorderVR=   ""
         Sections(1).Cells(0).Style.NoClipping=   -1  'True
         Sections(1).Cells(0).Style.RTF=   0   'False
         Sections(1).Cells(0).Style.fprops=   16777216
         Sections(1).Cells(1).Name=   "CELL_1"
         Sections(1).Cells(1).Exp=   """: "" & faktur"
         Sections(1).Cells(1).Height=   4
         Sections(1).Cells(1).AutoHeight=   0   'False
         Sections(1).Cells(1).PrivateStyle=   -1  'True
         Sections(1).Cells(1).Style.Name=   "<private>"
         Sections(1).Cells(1).Style.ParentName=   "tdb_Base"
         Sections(1).Cells(1).Style.Font_Name=   "Arial"
         Sections(1).Cells(1).Style.Font_Size=   8.25
         Sections(1).Cells(1).Style.Font_Bold=   -1  'True
         Sections(1).Cells(1).Style.Font_Italic=   0   'False
         Sections(1).Cells(1).Style.Font_Underline=   0   'False
         Sections(1).Cells(1).Style.Font_Strikeout=   0   'False
         Sections(1).Cells(1).Style.Font_Charset=   0
         Sections(1).Cells(1).Style.TextAlign=   0
         Sections(1).Cells(1).Style.TextVAlign=   1
         Sections(1).Cells(1).Style.TextWrap=   -1  'True
         Sections(1).Cells(1).Style.ForeColor=   0
         Sections(1).Cells(1).Style.BackColor=   16777215
         Sections(1).Cells(1).Style.NoFill=   -1  'True
         Sections(1).Cells(1).Style.BackPicFile=   ""
         Sections(1).Cells(1).Style.ForePicFile=   ""
         Sections(1).Cells(1).Style.BackPicVertPlacement=   0
         Sections(1).Cells(1).Style.BackPicHorzPlacement=   0
         Sections(1).Cells(1).Style.ForePicPlacement=   0
         Sections(1).Cells(1).Style.ForePicDrawMode=   0
         Sections(1).Cells(1).Style.MarginLeft=   6
         Sections(1).Cells(1).Style.MarginTop=   6
         Sections(1).Cells(1).Style.MarginRight=   6
         Sections(1).Cells(1).Style.MarginBottom=   6
         Sections(1).Cells(1).Style.HasBorders=   -1  'True
         Sections(1).Cells(1).Style.BorderHT=   ""
         Sections(1).Cells(1).Style.BorderHI=   ""
         Sections(1).Cells(1).Style.BorderHB=   ""
         Sections(1).Cells(1).Style.BorderVL=   ""
         Sections(1).Cells(1).Style.BorderVI=   ""
         Sections(1).Cells(1).Style.BorderVR=   ""
         Sections(1).Cells(1).Style.NoClipping=   -1  'True
         Sections(1).Cells(1).Style.RTF=   0   'False
         Sections(1).Cells(1).Style.fprops=   16777216
         Sections(1).Cells(2).Name=   "CELL_3"
         Sections(1).Cells(2).Exp=   """CUSTOMER"""
         Sections(1).Cells(2).NewLine=   -1  'True
         Sections(1).Cells(2).Width=   20
         Sections(1).Cells(2).Height=   4
         Sections(1).Cells(2).AutoHeight=   0   'False
         Sections(1).Cells(2).PrivateStyle=   -1  'True
         Sections(1).Cells(2).Style.Name=   "<private>"
         Sections(1).Cells(2).Style.ParentName=   "tdb_Base"
         Sections(1).Cells(2).Style.Font_Name=   "Arial"
         Sections(1).Cells(2).Style.Font_Size=   8.25
         Sections(1).Cells(2).Style.Font_Bold=   -1  'True
         Sections(1).Cells(2).Style.Font_Italic=   0   'False
         Sections(1).Cells(2).Style.Font_Underline=   0   'False
         Sections(1).Cells(2).Style.Font_Strikeout=   0   'False
         Sections(1).Cells(2).Style.Font_Charset=   0
         Sections(1).Cells(2).Style.TextAlign=   0
         Sections(1).Cells(2).Style.TextVAlign=   1
         Sections(1).Cells(2).Style.TextWrap=   -1  'True
         Sections(1).Cells(2).Style.ForeColor=   0
         Sections(1).Cells(2).Style.BackColor=   16777215
         Sections(1).Cells(2).Style.NoFill=   -1  'True
         Sections(1).Cells(2).Style.BackPicFile=   ""
         Sections(1).Cells(2).Style.ForePicFile=   ""
         Sections(1).Cells(2).Style.BackPicVertPlacement=   0
         Sections(1).Cells(2).Style.BackPicHorzPlacement=   0
         Sections(1).Cells(2).Style.ForePicPlacement=   0
         Sections(1).Cells(2).Style.ForePicDrawMode=   0
         Sections(1).Cells(2).Style.MarginLeft=   6
         Sections(1).Cells(2).Style.MarginTop=   6
         Sections(1).Cells(2).Style.MarginRight=   6
         Sections(1).Cells(2).Style.MarginBottom=   6
         Sections(1).Cells(2).Style.HasBorders=   -1  'True
         Sections(1).Cells(2).Style.BorderHT=   ""
         Sections(1).Cells(2).Style.BorderHI=   ""
         Sections(1).Cells(2).Style.BorderHB=   ""
         Sections(1).Cells(2).Style.BorderVL=   ""
         Sections(1).Cells(2).Style.BorderVI=   ""
         Sections(1).Cells(2).Style.BorderVR=   ""
         Sections(1).Cells(2).Style.NoClipping=   -1  'True
         Sections(1).Cells(2).Style.RTF=   0   'False
         Sections(1).Cells(2).Style.fprops=   16777216
         Sections(1).Cells(3).Name=   "CELL_4"
         Sections(1).Cells(3).Exp=   """: "" & supplier"
         Sections(1).Cells(3).Height=   4
         Sections(1).Cells(3).AutoHeight=   0   'False
         Sections(1).Cells(3).PrivateStyle=   -1  'True
         Sections(1).Cells(3).Style.Name=   "<private>"
         Sections(1).Cells(3).Style.ParentName=   "tdb_Base"
         Sections(1).Cells(3).Style.Font_Name=   "Arial"
         Sections(1).Cells(3).Style.Font_Size=   8.25
         Sections(1).Cells(3).Style.Font_Bold=   -1  'True
         Sections(1).Cells(3).Style.Font_Italic=   0   'False
         Sections(1).Cells(3).Style.Font_Underline=   0   'False
         Sections(1).Cells(3).Style.Font_Strikeout=   0   'False
         Sections(1).Cells(3).Style.Font_Charset=   0
         Sections(1).Cells(3).Style.TextAlign=   0
         Sections(1).Cells(3).Style.TextVAlign=   1
         Sections(1).Cells(3).Style.TextWrap=   -1  'True
         Sections(1).Cells(3).Style.ForeColor=   0
         Sections(1).Cells(3).Style.BackColor=   16777215
         Sections(1).Cells(3).Style.NoFill=   -1  'True
         Sections(1).Cells(3).Style.BackPicFile=   ""
         Sections(1).Cells(3).Style.ForePicFile=   ""
         Sections(1).Cells(3).Style.BackPicVertPlacement=   0
         Sections(1).Cells(3).Style.BackPicHorzPlacement=   0
         Sections(1).Cells(3).Style.ForePicPlacement=   0
         Sections(1).Cells(3).Style.ForePicDrawMode=   0
         Sections(1).Cells(3).Style.MarginLeft=   6
         Sections(1).Cells(3).Style.MarginTop=   6
         Sections(1).Cells(3).Style.MarginRight=   6
         Sections(1).Cells(3).Style.MarginBottom=   6
         Sections(1).Cells(3).Style.HasBorders=   -1  'True
         Sections(1).Cells(3).Style.BorderHT=   ""
         Sections(1).Cells(3).Style.BorderHI=   ""
         Sections(1).Cells(3).Style.BorderHB=   ""
         Sections(1).Cells(3).Style.BorderVL=   ""
         Sections(1).Cells(3).Style.BorderVI=   ""
         Sections(1).Cells(3).Style.BorderVR=   ""
         Sections(1).Cells(3).Style.NoClipping=   -1  'True
         Sections(1).Cells(3).Style.RTF=   0   'False
         Sections(1).Cells(3).Style.fprops=   16777216
         Sections(1).Cells(4).Name=   "CELL_6"
         Sections(1).Cells(4).Exp=   """TANGGAL TRANSAKSI"""
         Sections(1).Cells(4).NewLine=   -1  'True
         Sections(1).Cells(4).Width=   20
         Sections(1).Cells(4).Height=   4
         Sections(1).Cells(4).AutoHeight=   0   'False
         Sections(1).Cells(4).PrivateStyle=   -1  'True
         Sections(1).Cells(4).Style.Name=   "<private>"
         Sections(1).Cells(4).Style.ParentName=   "tdb_Base"
         Sections(1).Cells(4).Style.Font_Name=   "Arial"
         Sections(1).Cells(4).Style.Font_Size=   8.25
         Sections(1).Cells(4).Style.Font_Bold=   -1  'True
         Sections(1).Cells(4).Style.Font_Italic=   0   'False
         Sections(1).Cells(4).Style.Font_Underline=   0   'False
         Sections(1).Cells(4).Style.Font_Strikeout=   0   'False
         Sections(1).Cells(4).Style.Font_Charset=   0
         Sections(1).Cells(4).Style.TextAlign=   0
         Sections(1).Cells(4).Style.TextVAlign=   1
         Sections(1).Cells(4).Style.TextWrap=   -1  'True
         Sections(1).Cells(4).Style.ForeColor=   0
         Sections(1).Cells(4).Style.BackColor=   16777215
         Sections(1).Cells(4).Style.NoFill=   -1  'True
         Sections(1).Cells(4).Style.BackPicFile=   ""
         Sections(1).Cells(4).Style.ForePicFile=   ""
         Sections(1).Cells(4).Style.BackPicVertPlacement=   0
         Sections(1).Cells(4).Style.BackPicHorzPlacement=   0
         Sections(1).Cells(4).Style.ForePicPlacement=   0
         Sections(1).Cells(4).Style.ForePicDrawMode=   0
         Sections(1).Cells(4).Style.MarginLeft=   6
         Sections(1).Cells(4).Style.MarginTop=   6
         Sections(1).Cells(4).Style.MarginRight=   6
         Sections(1).Cells(4).Style.MarginBottom=   6
         Sections(1).Cells(4).Style.HasBorders=   -1  'True
         Sections(1).Cells(4).Style.BorderHT=   ""
         Sections(1).Cells(4).Style.BorderHI=   ""
         Sections(1).Cells(4).Style.BorderHB=   ""
         Sections(1).Cells(4).Style.BorderVL=   ""
         Sections(1).Cells(4).Style.BorderVI=   ""
         Sections(1).Cells(4).Style.BorderVR=   ""
         Sections(1).Cells(4).Style.NoClipping=   -1  'True
         Sections(1).Cells(4).Style.RTF=   0   'False
         Sections(1).Cells(4).Style.fprops=   16777216
         Sections(1).Cells(5).Name=   "CELL_7"
         Sections(1).Cells(5).Exp=   """: "" & tgl"
         Sections(1).Cells(5).Height=   4
         Sections(1).Cells(5).AutoHeight=   0   'False
         Sections(1).Cells(5).PrivateStyle=   -1  'True
         Sections(1).Cells(5).Style.Name=   "<private>"
         Sections(1).Cells(5).Style.ParentName=   "tdb_Base"
         Sections(1).Cells(5).Style.Font_Name=   "Arial"
         Sections(1).Cells(5).Style.Font_Size=   8.25
         Sections(1).Cells(5).Style.Font_Bold=   -1  'True
         Sections(1).Cells(5).Style.Font_Italic=   0   'False
         Sections(1).Cells(5).Style.Font_Underline=   0   'False
         Sections(1).Cells(5).Style.Font_Strikeout=   0   'False
         Sections(1).Cells(5).Style.Font_Charset=   0
         Sections(1).Cells(5).Style.TextAlign=   0
         Sections(1).Cells(5).Style.TextVAlign=   1
         Sections(1).Cells(5).Style.TextWrap=   -1  'True
         Sections(1).Cells(5).Style.ForeColor=   0
         Sections(1).Cells(5).Style.BackColor=   16777215
         Sections(1).Cells(5).Style.NoFill=   -1  'True
         Sections(1).Cells(5).Style.BackPicFile=   ""
         Sections(1).Cells(5).Style.ForePicFile=   ""
         Sections(1).Cells(5).Style.BackPicVertPlacement=   0
         Sections(1).Cells(5).Style.BackPicHorzPlacement=   0
         Sections(1).Cells(5).Style.ForePicPlacement=   0
         Sections(1).Cells(5).Style.ForePicDrawMode=   0
         Sections(1).Cells(5).Style.MarginLeft=   6
         Sections(1).Cells(5).Style.MarginTop=   6
         Sections(1).Cells(5).Style.MarginRight=   6
         Sections(1).Cells(5).Style.MarginBottom=   6
         Sections(1).Cells(5).Style.HasBorders=   -1  'True
         Sections(1).Cells(5).Style.BorderHT=   ""
         Sections(1).Cells(5).Style.BorderHI=   ""
         Sections(1).Cells(5).Style.BorderHB=   ""
         Sections(1).Cells(5).Style.BorderVL=   ""
         Sections(1).Cells(5).Style.BorderVI=   ""
         Sections(1).Cells(5).Style.BorderVR=   ""
         Sections(1).Cells(5).Style.NoClipping=   -1  'True
         Sections(1).Cells(5).Style.RTF=   0   'False
         Sections(1).Cells(5).Style.fprops=   16777216
         Sections(2).Name=   "DetailHeader"
         Sections(2).Type=   3
         Sections(2).StyleExp=   "tdb_TableHeader"
         Sections(2).Tabulator=   "Detail"
         Sections(2).Cells.Count=   7
         Sections(2).Cells(0).Name=   "Nomor"
         Sections(2).Cells(0).Exp=   """No."""
         Sections(2).Cells(1).Name=   "Kode"
         Sections(2).Cells(1).Exp=   """KODE """
         Sections(2).Cells(2).Name=   "JudulBuku"
         Sections(2).Cells(2).Exp=   """NAMA STOCK"""
         Sections(2).Cells(3).Name=   "Jumlah"
         Sections(2).Cells(3).Exp=   """QTY"""
         Sections(2).Cells(4).Name=   "Harga"
         Sections(2).Cells(4).Exp=   """HARGA"""
         Sections(2).Cells(5).Name=   "Discount"
         Sections(2).Cells(5).Exp=   """DISC%"""
         Sections(2).Cells(6).Name=   "Qty"
         Sections(2).Cells(6).Exp=   """JUMLAH"""
         Sections(3).Name=   "Detail"
         Sections(3).Type=   4
         Sections(3).StyleExp=   "'tdb_TableOddRow'"
         Sections(3).Cells.Count=   7
         Sections(3).Cells(0).Name=   "No"
         Sections(3).Cells(0).Exp=   "Sum(1,WillChange(faktur))"
         Sections(3).Cells(0).Width=   4
         Sections(3).Cells(1).Name=   "Kode"
         Sections(3).Cells(1).Exp=   "kode"
         Sections(3).Cells(1).Width=   13
         Sections(3).Cells(1).PrivateStyle=   -1  'True
         Sections(3).Cells(1).Style.Name=   "<private>"
         Sections(3).Cells(1).Style.ParentName=   "tdb_TableOddRow"
         Sections(3).Cells(1).Style.Font_Name=   "Arial"
         Sections(3).Cells(1).Style.Font_Size=   8.25
         Sections(3).Cells(1).Style.Font_Bold=   0   'False
         Sections(3).Cells(1).Style.Font_Italic=   0   'False
         Sections(3).Cells(1).Style.Font_Underline=   0   'False
         Sections(3).Cells(1).Style.Font_Strikeout=   0   'False
         Sections(3).Cells(1).Style.Font_Charset=   0
         Sections(3).Cells(1).Style.TextAlign=   1
         Sections(3).Cells(1).Style.TextVAlign=   0
         Sections(3).Cells(1).Style.TextWrap=   -1  'True
         Sections(3).Cells(1).Style.ForeColor=   0
         Sections(3).Cells(1).Style.BackColor=   16777215
         Sections(3).Cells(1).Style.NoFill=   -1  'True
         Sections(3).Cells(1).Style.BackPicFile=   ""
         Sections(3).Cells(1).Style.ForePicFile=   ""
         Sections(3).Cells(1).Style.BackPicVertPlacement=   0
         Sections(3).Cells(1).Style.BackPicHorzPlacement=   0
         Sections(3).Cells(1).Style.ForePicPlacement=   0
         Sections(3).Cells(1).Style.ForePicDrawMode=   0
         Sections(3).Cells(1).Style.MarginLeft=   6
         Sections(3).Cells(1).Style.MarginTop=   6
         Sections(3).Cells(1).Style.MarginRight=   6
         Sections(3).Cells(1).Style.MarginBottom=   6
         Sections(3).Cells(1).Style.HasBorders=   -1  'True
         Sections(3).Cells(1).Style.BorderHT=   "Quarter"
         Sections(3).Cells(1).Style.BorderHI=   "Quarter"
         Sections(3).Cells(1).Style.BorderHB=   "Double"
         Sections(3).Cells(1).Style.BorderVL=   "Single"
         Sections(3).Cells(1).Style.BorderVI=   "Single"
         Sections(3).Cells(1).Style.BorderVR=   "Single"
         Sections(3).Cells(1).Style.NoClipping=   -1  'True
         Sections(3).Cells(1).Style.RTF=   0   'False
         Sections(3).Cells(1).Style.fprops=   1
         Sections(3).Cells(2).Name=   "JudulBuku"
         Sections(3).Cells(2).Exp=   "namabarang"
         Sections(3).Cells(2).Width=   44
         Sections(3).Cells(3).Name=   "QTy"
         Sections(3).Cells(3).Exp=   "qty & "" "" & satuan"
         Sections(3).Cells(3).Width=   13
         Sections(3).Cells(4).Name=   "Harga"
         Sections(3).Cells(4).Exp=   "harga"
         Sections(3).Cells(4).Width=   15
         Sections(3).Cells(4).Format=   "###,###,##0.00"
         Sections(3).Cells(5).Name=   "Discount"
         Sections(3).Cells(5).Exp=   "disc"
         Sections(3).Cells(5).Width=   6
         Sections(3).Cells(6).Name=   "Total"
         Sections(3).Cells(6).Exp=   "jumlah"
         Sections(3).Cells(6).Width=   15
         Sections(3).Cells(6).Format=   "###,###,##0.00"
         Sections(4).Name=   "SECTION_2"
         Sections(4).Type=   5
         Sections(4).StyleExp=   "'Tdb_TableFooter'"
         Sections(4).Cells.Count=   18
         Sections(4).Cells(0).Name=   "SubTotal"
         Sections(4).Cells(0).Exp=   """Sub Total"""
         Sections(4).Cells(0).NewLine=   -1  'True
         Sections(4).Cells(0).Width=   80
         Sections(4).Cells(0).Height=   4
         Sections(4).Cells(0).AutoHeight=   0   'False
         Sections(4).Cells(1).Name=   "CELL_1"
         Sections(4).Cells(1).Exp=   """: Rp"""
         Sections(4).Cells(1).Width=   5
         Sections(4).Cells(1).Height=   4
         Sections(4).Cells(1).AutoHeight=   0   'False
         Sections(4).Cells(2).Name=   "CELL_2"
         Sections(4).Cells(2).Exp=   "subtotal"
         Sections(4).Cells(2).Height=   4
         Sections(4).Cells(2).AutoHeight=   0   'False
         Sections(4).Cells(2).Format=   "###,###,##0.00"
         Sections(4).Cells(3).Name=   "Discount1"
         Sections(4).Cells(3).Exp=   """Discount"""
         Sections(4).Cells(3).NewLine=   -1  'True
         Sections(4).Cells(3).Width=   80
         Sections(4).Cells(3).Height=   4
         Sections(4).Cells(3).AutoHeight=   0   'False
         Sections(4).Cells(4).Name=   "CELL_4"
         Sections(4).Cells(4).Exp=   """: Rp"""
         Sections(4).Cells(4).Width=   5
         Sections(4).Cells(4).Height=   4
         Sections(4).Cells(4).AutoHeight=   0   'False
         Sections(4).Cells(5).Name=   "CELL_5"
         Sections(4).Cells(5).Exp=   "discount"
         Sections(4).Cells(5).Height=   4
         Sections(4).Cells(5).AutoHeight=   0   'False
         Sections(4).Cells(5).Format=   "###,###,##0.00"
         Sections(4).Cells(6).Name=   "Pajak"
         Sections(4).Cells(6).Exp=   """DP"""
         Sections(4).Cells(6).NewLine=   -1  'True
         Sections(4).Cells(6).Width=   80
         Sections(4).Cells(6).Height=   4
         Sections(4).Cells(6).AutoHeight=   0   'False
         Sections(4).Cells(7).Name=   "CELL_10"
         Sections(4).Cells(7).Exp=   """: Rp"""
         Sections(4).Cells(7).Width=   5
         Sections(4).Cells(8).Name=   "CELL_11"
         Sections(4).Cells(8).Exp=   "pajak"
         Sections(4).Cells(8).Width=   15
         Sections(4).Cells(8).Height=   4
         Sections(4).Cells(8).AutoHeight=   0   'False
         Sections(4).Cells(8).Format=   "###,###,##0.00"
         Sections(4).Cells(9).Name=   "GrandTotal"
         Sections(4).Cells(9).Exp=   """Grand Total"""
         Sections(4).Cells(9).NewLine=   -1  'True
         Sections(4).Cells(9).Width=   80
         Sections(4).Cells(9).Height=   4
         Sections(4).Cells(9).AutoHeight=   0   'False
         Sections(4).Cells(10).Name=   "CELL_13"
         Sections(4).Cells(10).Exp=   """: Rp"""
         Sections(4).Cells(10).StyleExp=   "'Tdb_FooterGarisBawah'"
         Sections(4).Cells(10).Width=   5
         Sections(4).Cells(10).Height=   4
         Sections(4).Cells(10).AutoHeight=   0   'False
         Sections(4).Cells(11).Name=   "CELL_14"
         Sections(4).Cells(11).Exp=   "total"
         Sections(4).Cells(11).StyleExp=   "'Tdb_FooterGarisBawah'"
         Sections(4).Cells(11).Height=   4
         Sections(4).Cells(11).AutoHeight=   0   'False
         Sections(4).Cells(11).Format=   "###,###,##0.00"
         Sections(4).Cells(12).Name=   "CELL_12"
         Sections(4).Cells(12).Exp=   """Tunai"""
         Sections(4).Cells(12).NewLine=   -1  'True
         Sections(4).Cells(12).Width=   80
         Sections(4).Cells(12).Height=   4
         Sections(4).Cells(12).AutoHeight=   0   'False
         Sections(4).Cells(13).Name=   "CELL_15"
         Sections(4).Cells(13).Exp=   """: Rp"""
         Sections(4).Cells(13).Width=   5
         Sections(4).Cells(13).Height=   4
         Sections(4).Cells(13).AutoHeight=   0   'False
         Sections(4).Cells(14).Name=   "CELL_16"
         Sections(4).Cells(14).Exp=   "tunai"
         Sections(4).Cells(14).Height=   4
         Sections(4).Cells(14).AutoHeight=   0   'False
         Sections(4).Cells(14).Format=   "###,###,##0.00"
         Sections(4).Cells(15).Name=   "CELL_17"
         Sections(4).Cells(15).Exp=   """Piutang"""
         Sections(4).Cells(15).NewLine=   -1  'True
         Sections(4).Cells(15).Width=   80
         Sections(4).Cells(15).Height=   4
         Sections(4).Cells(15).AutoHeight=   0   'False
         Sections(4).Cells(16).Name=   "CELL_18"
         Sections(4).Cells(16).Exp=   """: Rp"""
         Sections(4).Cells(16).Width=   5
         Sections(4).Cells(16).Height=   4
         Sections(4).Cells(16).AutoHeight=   0   'False
         Sections(4).Cells(17).Name=   "CELL_19"
         Sections(4).Cells(17).Exp=   "piutang"
         Sections(4).Cells(17).Height=   4
         Sections(4).Cells(17).AutoHeight=   0   'False
         Sections(4).Cells(17).Format=   "###,###,##0.00"
         Sections(5).Name=   "SECTION_6"
         Sections(5).Condition=   "IsLastRec()"
         Sections(5).StyleExp=   "'total'"
         Sections(5).AutoHeight=   0   'False
         Sections(5).Height=   5
         Sections(5).Cells.Count=   23
         Sections(5).Cells(0).Name=   "CELL_0"
         Sections(5).Cells(0).Exp=   """ """
         Sections(5).Cells(0).NewLine=   -1  'True
         Sections(5).Cells(1).Name=   "CELL_1"
         Sections(5).Cells(1).Exp=   """JUMLAH KESELURUHAN :"""
         Sections(5).Cells(1).NewLine=   -1  'True
         Sections(5).Cells(2).Name=   "CELL_2"
         Sections(5).Cells(2).Exp=   """JUMLAH BARANG (QTY)"""
         Sections(5).Cells(2).NewLine=   -1  'True
         Sections(5).Cells(2).Width=   35
         Sections(5).Cells(2).PrivateStyle=   -1  'True
         Sections(5).Cells(2).Style.Name=   "<private>"
         Sections(5).Cells(2).Style.ParentName=   "total"
         Sections(5).Cells(2).Style.Font_Name=   "MS Sans Serif"
         Sections(5).Cells(2).Style.Font_Size=   8.25
         Sections(5).Cells(2).Style.Font_Bold=   -1  'True
         Sections(5).Cells(2).Style.Font_Italic=   0   'False
         Sections(5).Cells(2).Style.Font_Underline=   0   'False
         Sections(5).Cells(2).Style.Font_Strikeout=   0   'False
         Sections(5).Cells(2).Style.Font_Charset=   0
         Sections(5).Cells(2).Style.TextAlign=   0
         Sections(5).Cells(2).Style.TextVAlign=   1
         Sections(5).Cells(2).Style.TextWrap=   -1  'True
         Sections(5).Cells(2).Style.ForeColor=   0
         Sections(5).Cells(2).Style.BackColor=   16777215
         Sections(5).Cells(2).Style.NoFill=   -1  'True
         Sections(5).Cells(2).Style.BackPicFile=   ""
         Sections(5).Cells(2).Style.ForePicFile=   ""
         Sections(5).Cells(2).Style.BackPicVertPlacement=   0
         Sections(5).Cells(2).Style.BackPicHorzPlacement=   0
         Sections(5).Cells(2).Style.ForePicPlacement=   0
         Sections(5).Cells(2).Style.ForePicDrawMode=   0
         Sections(5).Cells(2).Style.MarginLeft=   6
         Sections(5).Cells(2).Style.MarginTop=   6
         Sections(5).Cells(2).Style.MarginRight=   6
         Sections(5).Cells(2).Style.MarginBottom=   6
         Sections(5).Cells(2).Style.HasBorders=   -1  'True
         Sections(5).Cells(2).Style.BorderHT=   ""
         Sections(5).Cells(2).Style.BorderHI=   ""
         Sections(5).Cells(2).Style.BorderHB=   ""
         Sections(5).Cells(2).Style.BorderVL=   ""
         Sections(5).Cells(2).Style.BorderVI=   ""
         Sections(5).Cells(2).Style.BorderVR=   ""
         Sections(5).Cells(2).Style.NoClipping=   -1  'True
         Sections(5).Cells(2).Style.RTF=   0   'False
         Sections(5).Cells(2).Style.fprops=   0
         Sections(5).Cells(3).Name=   "CELL_3"
         Sections(5).Cells(3).Exp=   """  : """
         Sections(5).Cells(3).Width=   3
         Sections(5).Cells(4).Name=   "CELL_4"
         Sections(5).Cells(4).Exp=   "TOTALQTY"
         Sections(5).Cells(4).Width=   20
         Sections(5).Cells(4).PrivateStyle=   -1  'True
         Sections(5).Cells(4).Format=   "###,###,##0"
         Sections(5).Cells(4).Style.Name=   "<private>"
         Sections(5).Cells(4).Style.ParentName=   "total"
         Sections(5).Cells(4).Style.Font_Name=   "MS Sans Serif"
         Sections(5).Cells(4).Style.Font_Size=   8.25
         Sections(5).Cells(4).Style.Font_Bold=   -1  'True
         Sections(5).Cells(4).Style.Font_Italic=   0   'False
         Sections(5).Cells(4).Style.Font_Underline=   0   'False
         Sections(5).Cells(4).Style.Font_Strikeout=   0   'False
         Sections(5).Cells(4).Style.Font_Charset=   0
         Sections(5).Cells(4).Style.TextAlign=   2
         Sections(5).Cells(4).Style.TextVAlign=   1
         Sections(5).Cells(4).Style.TextWrap=   -1  'True
         Sections(5).Cells(4).Style.ForeColor=   0
         Sections(5).Cells(4).Style.BackColor=   16777215
         Sections(5).Cells(4).Style.NoFill=   -1  'True
         Sections(5).Cells(4).Style.BackPicFile=   ""
         Sections(5).Cells(4).Style.ForePicFile=   ""
         Sections(5).Cells(4).Style.BackPicVertPlacement=   0
         Sections(5).Cells(4).Style.BackPicHorzPlacement=   0
         Sections(5).Cells(4).Style.ForePicPlacement=   0
         Sections(5).Cells(4).Style.ForePicDrawMode=   0
         Sections(5).Cells(4).Style.MarginLeft=   6
         Sections(5).Cells(4).Style.MarginTop=   6
         Sections(5).Cells(4).Style.MarginRight=   6
         Sections(5).Cells(4).Style.MarginBottom=   6
         Sections(5).Cells(4).Style.HasBorders=   -1  'True
         Sections(5).Cells(4).Style.BorderHT=   ""
         Sections(5).Cells(4).Style.BorderHI=   ""
         Sections(5).Cells(4).Style.BorderHB=   ""
         Sections(5).Cells(4).Style.BorderVL=   ""
         Sections(5).Cells(4).Style.BorderVI=   ""
         Sections(5).Cells(4).Style.BorderVR=   ""
         Sections(5).Cells(4).Style.NoClipping=   -1  'True
         Sections(5).Cells(4).Style.RTF=   0   'False
         Sections(5).Cells(4).Style.fprops=   1
         Sections(5).Cells(5).Name=   "CELL_5"
         Sections(5).Cells(5).Exp=   """SUB TOTAL ORDER (Rp)"""
         Sections(5).Cells(5).NewLine=   -1  'True
         Sections(5).Cells(5).Width=   35
         Sections(5).Cells(6).Name=   "CELL_6"
         Sections(5).Cells(6).Exp=   """ : """
         Sections(5).Cells(6).Width=   3
         Sections(5).Cells(7).Name=   "CELL_7"
         Sections(5).Cells(7).Exp=   "TOTALJUMLAH"
         Sections(5).Cells(7).Width=   20
         Sections(5).Cells(7).PrivateStyle=   -1  'True
         Sections(5).Cells(7).Format=   "###,###,###,###,##0.00"
         Sections(5).Cells(7).Style.Name=   "<private>"
         Sections(5).Cells(7).Style.ParentName=   "total"
         Sections(5).Cells(7).Style.Font_Name=   "MS Sans Serif"
         Sections(5).Cells(7).Style.Font_Size=   8.25
         Sections(5).Cells(7).Style.Font_Bold=   -1  'True
         Sections(5).Cells(7).Style.Font_Italic=   0   'False
         Sections(5).Cells(7).Style.Font_Underline=   0   'False
         Sections(5).Cells(7).Style.Font_Strikeout=   0   'False
         Sections(5).Cells(7).Style.Font_Charset=   0
         Sections(5).Cells(7).Style.TextAlign=   2
         Sections(5).Cells(7).Style.TextVAlign=   1
         Sections(5).Cells(7).Style.TextWrap=   -1  'True
         Sections(5).Cells(7).Style.ForeColor=   0
         Sections(5).Cells(7).Style.BackColor=   16777215
         Sections(5).Cells(7).Style.NoFill=   -1  'True
         Sections(5).Cells(7).Style.BackPicFile=   ""
         Sections(5).Cells(7).Style.ForePicFile=   ""
         Sections(5).Cells(7).Style.BackPicVertPlacement=   0
         Sections(5).Cells(7).Style.BackPicHorzPlacement=   0
         Sections(5).Cells(7).Style.ForePicPlacement=   0
         Sections(5).Cells(7).Style.ForePicDrawMode=   0
         Sections(5).Cells(7).Style.MarginLeft=   6
         Sections(5).Cells(7).Style.MarginTop=   6
         Sections(5).Cells(7).Style.MarginRight=   6
         Sections(5).Cells(7).Style.MarginBottom=   6
         Sections(5).Cells(7).Style.HasBorders=   -1  'True
         Sections(5).Cells(7).Style.BorderHT=   ""
         Sections(5).Cells(7).Style.BorderHI=   ""
         Sections(5).Cells(7).Style.BorderHB=   ""
         Sections(5).Cells(7).Style.BorderVL=   ""
         Sections(5).Cells(7).Style.BorderVI=   ""
         Sections(5).Cells(7).Style.BorderVR=   ""
         Sections(5).Cells(7).Style.NoClipping=   -1  'True
         Sections(5).Cells(7).Style.RTF=   0   'False
         Sections(5).Cells(7).Style.fprops=   1
         Sections(5).Cells(8).Name=   "CELL_8"
         Sections(5).Cells(8).Exp=   """TOTAL DISCOUNT"""
         Sections(5).Cells(8).NewLine=   -1  'True
         Sections(5).Cells(8).Width=   35
         Sections(5).Cells(9).Name=   "CELL_9"
         Sections(5).Cells(9).Exp=   """ : """
         Sections(5).Cells(9).Width=   3
         Sections(5).Cells(10).Name=   "CELL_10"
         Sections(5).Cells(10).Exp=   "TOTALDISCOUNT1"
         Sections(5).Cells(10).Width=   20
         Sections(5).Cells(10).PrivateStyle=   -1  'True
         Sections(5).Cells(10).Format=   "###,###,###,###,##0.00"
         Sections(5).Cells(10).Style.Name=   "<private>"
         Sections(5).Cells(10).Style.ParentName=   "total"
         Sections(5).Cells(10).Style.Font_Name=   "MS Sans Serif"
         Sections(5).Cells(10).Style.Font_Size=   8.25
         Sections(5).Cells(10).Style.Font_Bold=   -1  'True
         Sections(5).Cells(10).Style.Font_Italic=   0   'False
         Sections(5).Cells(10).Style.Font_Underline=   0   'False
         Sections(5).Cells(10).Style.Font_Strikeout=   0   'False
         Sections(5).Cells(10).Style.Font_Charset=   0
         Sections(5).Cells(10).Style.TextAlign=   2
         Sections(5).Cells(10).Style.TextVAlign=   1
         Sections(5).Cells(10).Style.TextWrap=   -1  'True
         Sections(5).Cells(10).Style.ForeColor=   0
         Sections(5).Cells(10).Style.BackColor=   16777215
         Sections(5).Cells(10).Style.NoFill=   -1  'True
         Sections(5).Cells(10).Style.BackPicFile=   ""
         Sections(5).Cells(10).Style.ForePicFile=   ""
         Sections(5).Cells(10).Style.BackPicVertPlacement=   0
         Sections(5).Cells(10).Style.BackPicHorzPlacement=   0
         Sections(5).Cells(10).Style.ForePicPlacement=   0
         Sections(5).Cells(10).Style.ForePicDrawMode=   0
         Sections(5).Cells(10).Style.MarginLeft=   6
         Sections(5).Cells(10).Style.MarginTop=   6
         Sections(5).Cells(10).Style.MarginRight=   6
         Sections(5).Cells(10).Style.MarginBottom=   6
         Sections(5).Cells(10).Style.HasBorders=   -1  'True
         Sections(5).Cells(10).Style.BorderHT=   ""
         Sections(5).Cells(10).Style.BorderHI=   ""
         Sections(5).Cells(10).Style.BorderHB=   ""
         Sections(5).Cells(10).Style.BorderVL=   ""
         Sections(5).Cells(10).Style.BorderVI=   ""
         Sections(5).Cells(10).Style.BorderVR=   ""
         Sections(5).Cells(10).Style.NoClipping=   -1  'True
         Sections(5).Cells(10).Style.RTF=   0   'False
         Sections(5).Cells(10).Style.fprops=   1
         Sections(5).Cells(11).Name=   "CELL_14"
         Sections(5).Cells(11).Exp=   """TOTAL PAJAK"""
         Sections(5).Cells(11).NewLine=   -1  'True
         Sections(5).Cells(11).Width=   35
         Sections(5).Cells(12).Name=   "CELL_15"
         Sections(5).Cells(12).Exp=   """ : """
         Sections(5).Cells(12).Width=   3
         Sections(5).Cells(13).Name=   "CELL_16"
         Sections(5).Cells(13).Exp=   "TOTALPAJAK"
         Sections(5).Cells(13).Width=   20
         Sections(5).Cells(13).PrivateStyle=   -1  'True
         Sections(5).Cells(13).Format=   "###,###,###,###,##0.00"
         Sections(5).Cells(13).Style.Name=   "<private>"
         Sections(5).Cells(13).Style.ParentName=   "total"
         Sections(5).Cells(13).Style.Font_Name=   "MS Sans Serif"
         Sections(5).Cells(13).Style.Font_Size=   8.25
         Sections(5).Cells(13).Style.Font_Bold=   -1  'True
         Sections(5).Cells(13).Style.Font_Italic=   0   'False
         Sections(5).Cells(13).Style.Font_Underline=   0   'False
         Sections(5).Cells(13).Style.Font_Strikeout=   0   'False
         Sections(5).Cells(13).Style.Font_Charset=   0
         Sections(5).Cells(13).Style.TextAlign=   2
         Sections(5).Cells(13).Style.TextVAlign=   1
         Sections(5).Cells(13).Style.TextWrap=   -1  'True
         Sections(5).Cells(13).Style.ForeColor=   0
         Sections(5).Cells(13).Style.BackColor=   16777215
         Sections(5).Cells(13).Style.NoFill=   -1  'True
         Sections(5).Cells(13).Style.BackPicFile=   ""
         Sections(5).Cells(13).Style.ForePicFile=   ""
         Sections(5).Cells(13).Style.BackPicVertPlacement=   0
         Sections(5).Cells(13).Style.BackPicHorzPlacement=   0
         Sections(5).Cells(13).Style.ForePicPlacement=   0
         Sections(5).Cells(13).Style.ForePicDrawMode=   0
         Sections(5).Cells(13).Style.MarginLeft=   6
         Sections(5).Cells(13).Style.MarginTop=   6
         Sections(5).Cells(13).Style.MarginRight=   6
         Sections(5).Cells(13).Style.MarginBottom=   6
         Sections(5).Cells(13).Style.HasBorders=   -1  'True
         Sections(5).Cells(13).Style.BorderHT=   ""
         Sections(5).Cells(13).Style.BorderHI=   ""
         Sections(5).Cells(13).Style.BorderHB=   ""
         Sections(5).Cells(13).Style.BorderVL=   ""
         Sections(5).Cells(13).Style.BorderVI=   ""
         Sections(5).Cells(13).Style.BorderVR=   ""
         Sections(5).Cells(13).Style.NoClipping=   -1  'True
         Sections(5).Cells(13).Style.RTF=   0   'False
         Sections(5).Cells(13).Style.fprops=   1
         Sections(5).Cells(14).Name=   "CELL_17"
         Sections(5).Cells(14).Exp=   """TOTAL PENJUALAN"""
         Sections(5).Cells(14).NewLine=   -1  'True
         Sections(5).Cells(14).Width=   35
         Sections(5).Cells(15).Name=   "CELL_18"
         Sections(5).Cells(15).Exp=   """ : """
         Sections(5).Cells(15).Width=   3
         Sections(5).Cells(16).Name=   "CELL_19"
         Sections(5).Cells(16).Exp=   "GRANDTOTAL"
         Sections(5).Cells(16).Width=   20
         Sections(5).Cells(16).PrivateStyle=   -1  'True
         Sections(5).Cells(16).Format=   "###,###,###,###,##0.00"
         Sections(5).Cells(16).Style.Name=   "<private>"
         Sections(5).Cells(16).Style.ParentName=   "total"
         Sections(5).Cells(16).Style.Font_Name=   "MS Sans Serif"
         Sections(5).Cells(16).Style.Font_Size=   8.25
         Sections(5).Cells(16).Style.Font_Bold=   -1  'True
         Sections(5).Cells(16).Style.Font_Italic=   0   'False
         Sections(5).Cells(16).Style.Font_Underline=   0   'False
         Sections(5).Cells(16).Style.Font_Strikeout=   0   'False
         Sections(5).Cells(16).Style.Font_Charset=   0
         Sections(5).Cells(16).Style.TextAlign=   2
         Sections(5).Cells(16).Style.TextVAlign=   1
         Sections(5).Cells(16).Style.TextWrap=   -1  'True
         Sections(5).Cells(16).Style.ForeColor=   0
         Sections(5).Cells(16).Style.BackColor=   16777215
         Sections(5).Cells(16).Style.NoFill=   -1  'True
         Sections(5).Cells(16).Style.BackPicFile=   ""
         Sections(5).Cells(16).Style.ForePicFile=   ""
         Sections(5).Cells(16).Style.BackPicVertPlacement=   0
         Sections(5).Cells(16).Style.BackPicHorzPlacement=   0
         Sections(5).Cells(16).Style.ForePicPlacement=   0
         Sections(5).Cells(16).Style.ForePicDrawMode=   0
         Sections(5).Cells(16).Style.MarginLeft=   6
         Sections(5).Cells(16).Style.MarginTop=   6
         Sections(5).Cells(16).Style.MarginRight=   6
         Sections(5).Cells(16).Style.MarginBottom=   6
         Sections(5).Cells(16).Style.HasBorders=   -1  'True
         Sections(5).Cells(16).Style.BorderHT=   ""
         Sections(5).Cells(16).Style.BorderHI=   ""
         Sections(5).Cells(16).Style.BorderHB=   ""
         Sections(5).Cells(16).Style.BorderVL=   ""
         Sections(5).Cells(16).Style.BorderVI=   ""
         Sections(5).Cells(16).Style.BorderVR=   ""
         Sections(5).Cells(16).Style.NoClipping=   -1  'True
         Sections(5).Cells(16).Style.RTF=   0   'False
         Sections(5).Cells(16).Style.fprops=   1
         Sections(5).Cells(17).Name=   "CELL_20"
         Sections(5).Cells(17).Exp=   """TOTAL TUNAI"""
         Sections(5).Cells(17).NewLine=   -1  'True
         Sections(5).Cells(17).Width=   35
         Sections(5).Cells(18).Name=   "CELL_21"
         Sections(5).Cells(18).Exp=   """ : """
         Sections(5).Cells(18).Width=   3
         Sections(5).Cells(19).Name=   "CELL_22"
         Sections(5).Cells(19).Exp=   "TUNAIREF"
         Sections(5).Cells(19).Width=   20
         Sections(5).Cells(19).PrivateStyle=   -1  'True
         Sections(5).Cells(19).Format=   "###,###,###,###,##0.00"
         Sections(5).Cells(19).Style.Name=   ""
         Sections(5).Cells(19).Style.ParentName=   "total"
         Sections(5).Cells(19).Style.Font_Name=   "MS Sans Serif"
         Sections(5).Cells(19).Style.Font_Size=   8.25
         Sections(5).Cells(19).Style.Font_Bold=   -1  'True
         Sections(5).Cells(19).Style.Font_Italic=   0   'False
         Sections(5).Cells(19).Style.Font_Underline=   0   'False
         Sections(5).Cells(19).Style.Font_Strikeout=   0   'False
         Sections(5).Cells(19).Style.Font_Charset=   0
         Sections(5).Cells(19).Style.TextAlign=   2
         Sections(5).Cells(19).Style.TextVAlign=   1
         Sections(5).Cells(19).Style.TextWrap=   -1  'True
         Sections(5).Cells(19).Style.ForeColor=   0
         Sections(5).Cells(19).Style.BackColor=   16777215
         Sections(5).Cells(19).Style.NoFill=   -1  'True
         Sections(5).Cells(19).Style.BackPicFile=   ""
         Sections(5).Cells(19).Style.ForePicFile=   ""
         Sections(5).Cells(19).Style.BackPicVertPlacement=   0
         Sections(5).Cells(19).Style.BackPicHorzPlacement=   0
         Sections(5).Cells(19).Style.ForePicPlacement=   0
         Sections(5).Cells(19).Style.ForePicDrawMode=   0
         Sections(5).Cells(19).Style.MarginLeft=   6
         Sections(5).Cells(19).Style.MarginTop=   6
         Sections(5).Cells(19).Style.MarginRight=   6
         Sections(5).Cells(19).Style.MarginBottom=   6
         Sections(5).Cells(19).Style.HasBorders=   -1  'True
         Sections(5).Cells(19).Style.BorderHT=   ""
         Sections(5).Cells(19).Style.BorderHI=   ""
         Sections(5).Cells(19).Style.BorderHB=   ""
         Sections(5).Cells(19).Style.BorderVL=   ""
         Sections(5).Cells(19).Style.BorderVI=   ""
         Sections(5).Cells(19).Style.BorderVR=   ""
         Sections(5).Cells(19).Style.NoClipping=   -1  'True
         Sections(5).Cells(19).Style.RTF=   0   'False
         Sections(5).Cells(19).Style.fprops=   1
         Sections(5).Cells(20).Name=   "CELL_23"
         Sections(5).Cells(20).Exp=   """TOTAL PIUTANG"""
         Sections(5).Cells(20).NewLine=   -1  'True
         Sections(5).Cells(20).Width=   35
         Sections(5).Cells(21).Name=   "CELL_24"
         Sections(5).Cells(21).Exp=   """ : """
         Sections(5).Cells(21).Width=   3
         Sections(5).Cells(22).Name=   "CELL_25"
         Sections(5).Cells(22).Exp=   "PIUTANGREF"
         Sections(5).Cells(22).Width=   20
         Sections(5).Cells(22).PrivateStyle=   -1  'True
         Sections(5).Cells(22).Format=   "###,###,###,###,##0.00"
         Sections(5).Cells(22).Style.Name=   ""
         Sections(5).Cells(22).Style.ParentName=   "total"
         Sections(5).Cells(22).Style.Font_Name=   "MS Sans Serif"
         Sections(5).Cells(22).Style.Font_Size=   8.25
         Sections(5).Cells(22).Style.Font_Bold=   -1  'True
         Sections(5).Cells(22).Style.Font_Italic=   0   'False
         Sections(5).Cells(22).Style.Font_Underline=   0   'False
         Sections(5).Cells(22).Style.Font_Strikeout=   0   'False
         Sections(5).Cells(22).Style.Font_Charset=   0
         Sections(5).Cells(22).Style.TextAlign=   2
         Sections(5).Cells(22).Style.TextVAlign=   1
         Sections(5).Cells(22).Style.TextWrap=   -1  'True
         Sections(5).Cells(22).Style.ForeColor=   0
         Sections(5).Cells(22).Style.BackColor=   16777215
         Sections(5).Cells(22).Style.NoFill=   -1  'True
         Sections(5).Cells(22).Style.BackPicFile=   ""
         Sections(5).Cells(22).Style.ForePicFile=   ""
         Sections(5).Cells(22).Style.BackPicVertPlacement=   0
         Sections(5).Cells(22).Style.BackPicHorzPlacement=   0
         Sections(5).Cells(22).Style.ForePicPlacement=   0
         Sections(5).Cells(22).Style.ForePicDrawMode=   0
         Sections(5).Cells(22).Style.MarginLeft=   6
         Sections(5).Cells(22).Style.MarginTop=   6
         Sections(5).Cells(22).Style.MarginRight=   6
         Sections(5).Cells(22).Style.MarginBottom=   6
         Sections(5).Cells(22).Style.HasBorders=   -1  'True
         Sections(5).Cells(22).Style.BorderHT=   ""
         Sections(5).Cells(22).Style.BorderHI=   ""
         Sections(5).Cells(22).Style.BorderHB=   ""
         Sections(5).Cells(22).Style.BorderVL=   ""
         Sections(5).Cells(22).Style.BorderVI=   ""
         Sections(5).Cells(22).Style.BorderVR=   ""
         Sections(5).Cells(22).Style.NoClipping=   -1  'True
         Sections(5).Cells(22).Style.RTF=   0   'False
         Sections(5).Cells(22).Style.fprops=   1
         Styles.Count    =   7
         Styles(0).Name  =   "tdb_Base"
         Styles(0).ParentName=   ""
         Styles(0).Font_Name=   "Arial"
         Styles(0).Font_Size=   8.25
         Styles(0).Font_Charset=   0
         Styles(0).TextAlign=   0
         Styles(0).TextVAlign=   1
         Styles(0).NoClipping=   -1  'True
         Styles(1).Name  =   "Tdb_FooterGarisBawah"
         Styles(1).ParentName=   "tdb_Base"
         Styles(1).Font_Name=   "Arial"
         Styles(1).Font_Size=   8.25
         Styles(1).Font_Charset=   0
         Styles(1).TextAlign=   2
         Styles(1).TextVAlign=   1
         Styles(1).BorderHT=   "Double"
         Styles(1).NoClipping=   -1  'True
         Styles(1).fprops=   163841
         Styles(2).Name  =   "tdb_PageHeader"
         Styles(2).ParentName=   "tdb_Base"
         Styles(2).Font_Name=   "Arial"
         Styles(2).Font_Size=   8.25
         Styles(2).Font_Charset=   0
         Styles(2).TextAlign=   2
         Styles(2).TextVAlign=   1
         Styles(2).NoClipping=   -1  'True
         Styles(2).fprops=   1
         Styles(3).Name  =   "tdb_TableOddRow"
         Styles(3).ParentName=   "tdb_Base"
         Styles(3).Font_Name=   "Arial"
         Styles(3).Font_Size=   8.25
         Styles(3).Font_Charset=   0
         Styles(3).BorderHT=   "Quarter"
         Styles(3).BorderHI=   "Quarter"
         Styles(3).BorderHB=   "Double"
         Styles(3).BorderVL=   "Single"
         Styles(3).BorderVI=   "Single"
         Styles(3).BorderVR=   "Single"
         Styles(3).NoClipping=   -1  'True
         Styles(4).Name  =   "tdb_TableHeader"
         Styles(4).ParentName=   "tdb_Base"
         Styles(4).Font_Name=   "Arial"
         Styles(4).Font_Size=   8.25
         Styles(4).Font_Bold=   -1  'True
         Styles(4).Font_Charset=   0
         Styles(4).TextAlign=   1
         Styles(4).TextVAlign=   1
         Styles(4).ForeColor=   4194304
         Styles(4).NoFill=   0   'False
         Styles(4).BorderHT=   "Double"
         Styles(4).BorderHI=   "Double"
         Styles(4).BorderHB=   "Double"
         Styles(4).BorderVL=   "Single"
         Styles(4).BorderVI=   "Single"
         Styles(4).BorderVR=   "Single"
         Styles(4).NoClipping=   -1  'True
         Styles(5).Name  =   "Tdb_TableFooter"
         Styles(5).ParentName=   "tdb_Base"
         Styles(5).Font_Name=   "Arial"
         Styles(5).Font_Size=   8.25
         Styles(5).Font_Charset=   0
         Styles(5).TextAlign=   2
         Styles(5).TextVAlign=   1
         Styles(5).NoClipping=   -1  'True
         Styles(5).fprops=   3
         Styles(6).Name  =   "total"
         Styles(6).ParentName=   "tdb_Base"
         Styles(6).Font_Name=   "MS Sans Serif"
         Styles(6).Font_Size=   8.25
         Styles(6).Font_Bold=   -1  'True
         Styles(6).Font_Charset=   0
         Styles(6).TextAlign=   0
         Styles(6).TextVAlign=   1
         Styles(6).NoClipping=   -1  'True
         Styles(6).fprops=   23068672
         Lines.Count     =   3
         Lines(0).Name   =   "Single"
         Lines(0).Thickness=   4
         Lines(1).Name   =   "Double"
         Lines(1).Thickness=   5
         Lines(2).Name   =   "Quarter"
         Lines(2).Thickness=   1
         Lines(2).Color  =   8421504
         Profiles.Count  =   1
         Profiles(0).Name=   "PROFILE_0"
         Profiles(0).Active=   -1  'True
         Profiles(0).PreviewNoMinimize=   -1  'True
         Profiles(0).PreviewNoMaximize=   -1  'True
         Profiles(0).PreviewNoResize=   -1  'True
         Profiles(0).PreviewMaximized=   -1  'True
         Profiles(0).PreviewNoSaveLoad=   -1  'True
         Profiles(0).PrinterMarginLeft=   10
         Profiles(0).PrinterMarginTop=   10
         Profiles(0).PrinterMarginRight=   10
         Profiles(0).PrinterMarginBottom=   10
         Profiles(0).PrinterMargins_set=   -1  'True
         Profiles(0).PrinterPaperUserSize_set=   -1  'True
      End
      Begin BiSAFramProject.BiSAFrame BiSAFrame5 
         Height          =   510
         Left            =   2265
         Top             =   2190
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   900
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         Begin VB.OptionButton optOutstanding 
            Caption         =   "All"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   1410
            TabIndex        =   15
            Top             =   120
            Width           =   600
         End
         Begin VB.OptionButton optOutstanding 
            Caption         =   "Outstanding"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   105
            Width           =   1230
         End
      End
      Begin BiSAFramProject.BiSAFrame BiSAFrame6 
         Height          =   510
         Left            =   2265
         Top             =   2700
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   900
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         Begin VB.OptionButton optJenisOrder 
            Caption         =   "All"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   2010
            TabIndex        =   18
            Top             =   120
            Width           =   945
         End
         Begin VB.OptionButton optJenisOrder 
            Caption         =   "Promo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   1080
            TabIndex        =   17
            Top             =   120
            Width           =   945
         End
         Begin VB.OptionButton optJenisOrder 
            Caption         =   "Reguler"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   16
            Top             =   120
            Width           =   945
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   270
         TabIndex        =   11
         Top             =   885
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Tampilkan"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   285
         TabIndex        =   10
         Top             =   1365
         Width           =   1065
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   0
      Top             =   3525
      Width           =   6990
      _ExtentX        =   12330
      _ExtentY        =   1111
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   5805
         TabIndex        =   12
         Top             =   120
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   767
         Caption         =   "     &Exit"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         BackColor       =   -2147483633
         Picture         =   "rptDetailOrder.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   5355
         TabIndex        =   13
         Top             =   120
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   767
         Caption         =   ""
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         BackColor       =   -2147483633
         Picture         =   "rptDetailOrder.frx":00A6
      End
   End
End
Attribute VB_Name = "rptDetailOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    keybd_event VK_TAB, 0, 0, 0
  End If
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
  GetData
End Sub

Private Sub ccustomer_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "anggota", "kodeanggota,nama,alamat", "nama", sisContent, cCustomer.Text)
  If Not dbData.EOF Then
    cCustomer.Text = cCustomer.Browse(dbData)
    cCustomer.Text = GetNull(dbData!kodeanggota)
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  dDate(0).Value = BOM(Date)
  optKodeStock(1).Value = True
  optTunai(2).Value = True
  optOutstanding(0).Value = True
  optJenisOrder(0).Value = True

  TabIndex dDate(0), n
  TabIndex dDate(1), n
  TabIndex Check1, n
  TabIndex cCustomer, n
  TabIndex optKodeStock(0), n
  TabIndex optKodeStock(1), n
  
  TabIndex optTunai(0), n
  TabIndex optTunai(1), n
  TabIndex optTunai(2), n
  
  TabIndex optOutstanding(0), n
  TabIndex optOutstanding(1), n
  TabIndex optJenisOrder(0), n
  TabIndex optJenisOrder(1), n
  TabIndex optJenisOrder(2), n

  
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub

Private Sub GetData()
Dim cSQL As String
Dim cFilter As String
Dim n As Integer
Dim nJumlah As Double
Dim nQty As Double
Dim nDiscount1 As Double
Dim nDiscount2 As Double
Dim nPajak As Double
Dim nTotal As Double
Dim Tunai As Double
Dim Piutang As Double
Dim dp As Double

  nJumlah = 0
  nQty = 0
  nDiscount1 = 0
  nDiscount2 = 0
  nPajak = 0
  nTotal = 0

  
  vaArray.ReDim 0, -1, 0, 16
  cFilter = ""
  If Check1.Value = 1 Then
    cFilter = " AND t.kodeanggota = '" & cCustomer.Text & "'"
  End If
  cSQL = "SELECT p.nomormemberorder,p.kodesatuan,s.barcode,p.kodestock,p.qty,p.harga,p.discount as disc,p.jumlah,t.tunai,t.piutang,t.dp,"
  cSQL = cSQL & " t.tgl,t.subtotal,t.discount,t.pajak,t.total,"
  cSQL = cSQL & " t.kodeanggota,s.nama as namabarang,r.nama as namaanggota"
  cSQL = cSQL & " From memberorder p"
  cSQL = cSQL & " LEFT JOIN totmemberorder t on t.nomormemberorder = p.nomormemberorder"
  cSQL = cSQL & " LEFT JOIN stock s on s.kodestock = p.kodestock"
  cSQL = cSQL & " LEFT JOIN anggota r on r.kodeanggota = t.kodeanggota"
  cSQL = cSQL & " WHERE p.tgl >='" & Format(dDate(0).Value, "yyyy-MM-dd") & "'"
  cSQL = cSQL & " AND p.tgl <='" & Format(dDate(1).Value, "yyyy-MM-dd") & "'" & cFilter
  If optOutstanding(0).Value = True Then
    cSQL = cSQL & " AND t.status = 0 "
  End If
  If optJenisOrder(0).Value = True Then
    cSQL = cSQL & " AND t.jenisorder = 'R'"
  End If
  If optJenisOrder(1).Value = True Then
    cSQL = cSQL & " AND t.jenisorder = 'P'"
  End If
  cSQL = cSQL & " ORDER BY p.nomormemberorder"
  
  Set dbData = objData.Sql(GetDSN, cSQL)
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    dbData.MoveFirst
    Do While Not dbData.EOF
      FrmPB.RunPB
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      
      vaArray(n, 0) = (dbData!nomormemberorder)
      vaArray(n, 1) = (dbData!namaanggota)
      vaArray(n, 2) = Format(dbData!tgl, "dd-MM-yyyy")
      vaArray(n, 3) = IIf(optKodeStock(0).Value = True, (dbData!KodeStock), (dbData!Barcode))
      vaArray(n, 4) = (dbData!Namabarang)
      vaArray(n, 5) = (dbData!qty)
      vaArray(n, 6) = (dbData!kodesatuan)
      vaArray(n, 7) = (dbData!Harga)
      vaArray(n, 8) = (dbData!jumlah)
      vaArray(n, 9) = (dbData!Subtotal)
      vaArray(n, 10) = (dbData!Discount)
      vaArray(n, 11) = 0
      vaArray(n, 12) = (dbData!dp)
      vaArray(n, 13) = (dbData!Total)
      vaArray(n, 14) = (dbData!Disc)
      
      vaArray(n, 15) = (dbData!Tunai)
      vaArray(n, 16) = (dbData!Piutang)
      
      nQty = nQty + vaArray(n, 5)
      
      If optTunai(0).Value = True Then
        If vaArray(n, 16) > 0 Then
          vaArray.DeleteRows n
        End If
      End If
      
      If optTunai(1).Value = True Then
        If vaArray(n, 16) = 0 Then
          vaArray.DeleteRows n
        End If
      End If
      
      dbData.MoveNext
    Loop
    FrmPB.EndPB
    
    GetSUM nJumlah, nDiscount1, nDiscount2, nPajak, nTotal, Tunai, Piutang, dp
        
    tdb.Parameters("TGL1").ValueExpression = "'" & Format(dDate(0).Value, "dd-MM-yyyy") & "'"
    tdb.Parameters("TGL2").ValueExpression = "'" & Format(dDate(1).Value, "dd-MM-yyyy") & "'"
    
    tdb.Parameters("TOTALJUMLAH").ValueExpression = nJumlah
    tdb.Parameters("TOTALQTY").ValueExpression = nQty
    tdb.Parameters("TOTALDISCOUNT1").ValueExpression = nDiscount1
    tdb.Parameters("TOTALPAJAK").ValueExpression = nPajak
    tdb.Parameters("GRANDTOTAL").ValueExpression = nTotal
    tdb.Parameters("TUNAIREF").ValueExpression = Tunai
    tdb.Parameters("PIUTANGREF").ValueExpression = Piutang
    
    
    Set tdb.Array = vaArray
    tdb.Refresh
    tdb.PrintPreview
  Else
    MsgBox "Data tidak ada...", vbInformation
    Exit Sub
  End If
End Sub

Private Sub GetSUM(ByRef nSubTotal As Double, ByRef nDisc As Double, _
                    ByRef nDisc1 As Double, ByRef nPajak As Double, _
                    ByRef nTotal As Double, ByRef Tunai As Double, ByRef Piutang As Double, ByRef nDP As Double)
  
Dim cSQL As String
Dim cWhere As String
  
  nSubTotal = 0
  nDisc = 0
  nDisc1 = 0
  nPajak = 0
  nTotal = 0
  
  cWhere = ""
  
  If optTunai(0).Value = True Then
    cWhere = " AND piutang = 0"
  End If
  
  If optTunai(1).Value = True Then
    cWhere = " AND piutang <> 0"
  End If
  
  cSQL = "SELECT SUM(subtotal) as SubTotal,SUM(discount) as Disc, SUM(pajak) as Pajak,SUM(total) as Total,sum(tunai) as tunai,sum(piutang) as piutang,sum(dp) as dp "
  cSQL = cSQL & " FROM totmemberorder"
  cSQL = cSQL & " WHERE Tgl >='" & Format(dDate(0).Value, "yyyy-MM-dd") & "'"
  cSQL = cSQL & " AND Tgl <='" & Format(dDate(1).Value, "yyyy-MM-dd") & "'"
  cSQL = cSQL & cWhere
  If Check1.Value = 1 Then
    cSQL = cSQL & " AND kodeanggota = '" & cCustomer.Text & "'"
  End If
  Set dbData = objData.Sql(GetDSN, cSQL)
  If Not dbData.EOF Then
    nSubTotal = GetNull(dbData!Subtotal)
    nDisc = GetNull(dbData!Disc)
    nDisc1 = 0
    nPajak = GetNull(dbData!PAJAK)
    nTotal = GetNull(dbData!Total)
    Tunai = GetNull(dbData!Tunai)
    Piutang = GetNull(dbData!Piutang)
    nDP = GetNull(dbData!dp)
  End If
End Sub

Private Sub optJenisOrder_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    keybd_event VK_TAB, 0, 0, 0
  End If
End Sub

Private Sub optKodeStock_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    keybd_event VK_TAB, 0, 0, 0
  End If
End Sub

Private Sub optOutstanding_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    keybd_event VK_TAB, 0, 0, 0
  End If
End Sub

Private Sub optTunai_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    keybd_event VK_TAB, 0, 0, 0
  End If
End Sub


