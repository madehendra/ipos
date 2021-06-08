VERSION 5.00
Begin VB.Form cfgAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   9990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Registrasi...."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6180
      TabIndex        =   3
      Top             =   5850
      Width           =   1605
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Your System Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7800
      TabIndex        =   1
      Top             =   5835
      Width           =   2175
   End
   Begin VB.Image Image3 
      Height          =   1080
      Left            =   4050
      Picture         =   "cfgAbout.frx":0000
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   1080
      Left            =   2100
      Picture         =   "cfgAbout.frx":0C2D
      Top             =   5040
      Width           =   1905
   End
   Begin VB.Image Image1 
      Height          =   1080
      Left            =   195
      Picture         =   "cfgAbout.frx":16EA
      Top             =   5040
      Width           =   1860
   End
   Begin VB.Label Label2 
      Caption         =   "Powered By..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   225
      TabIndex        =   2
      Top             =   4800
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4560
      Left            =   120
      TabIndex        =   0
      Top             =   195
      Width           =   9405
   End
End
Attribute VB_Name = "cfgAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cText As String


Private Sub Command1_Click()
  AboutOS
End Sub

Private Sub Form_Load()
 SetIcon Me.hWnd, "SIKD"
 CenterForm Me
 AboutMe
End Sub


Private Sub AboutMe()
Dim cString As String
  
    cString = "Perhatian: " & vbCrLf & vbCrLf
    cString = cString & "Program ini bukan merupakan program freeware atau program bebas atau program open source" & vbCrLf & vbCrLf
    cString = cString & "Program ini sudah terdaftar dan dilindungi oleh Undang Undang Hak Intelektual di Direktorat Hukum dan Kehakiman Republik Indonesia" & vbCrLf & vbCrLf
    cString = cString & "Tidak diperkenankan untuk mereproduksi, menyebarluaskan atau menyalin (copy) program ini tanpa izin" & vbCrLf & vbCrLf
    cString = cString & "Segala bentuk pelanggaran terhadap Software ini dapat dituntut dipengadilan di seluruh wilayah Republik Indonesia" & vbCrLf
    cString = cString & "dengan ancaman denda sebesar besarnya Rp 500.000.000,00 (Lima Ratus Juta Rupiah) atau hukuman kurungan selama lamanya 10 tahun penjara" & vbCrLf & vbCrLf
    
    cString = cString & App.ProductName & " v" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf
    cString = cString & App.LegalCopyright & vbCrLf
    cString = cString & App.CompanyName & vbCrLf & vbCrLf
    cString = cString & "Need Support?" & vbCrLf
    cString = cString & "Pande Software Indonesia" & vbCrLf
    cString = cString & "http://www.codesuite.uni.cc" & vbCrLf
    cString = cString & "support@codesuite.uni.cc" & vbCrLf
    cString = cString & "e-mail: made.hendra@yahoo.co.id"
    
    Label1.Caption = cString
End Sub

Private Sub AboutOS()
    Call ShellAbout(Me.hWnd, _
          App.Title & " System Info Window#OS Information:", _
          vbCrLf & _
          "NOT REGISTERED", _
          Me.Icon)
End Sub

