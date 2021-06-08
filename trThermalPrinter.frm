VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Begin VB.Form trThermalPrinter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "THREMAL PRINTER TEST PRINT"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin BiSAButtonProject.BiSAButton BiSAButton1 
      Height          =   510
      Left            =   1305
      TabIndex        =   0
      Top             =   2040
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   900
      Caption         =   "Label1"
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
End
Attribute VB_Name = "trThermalPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

