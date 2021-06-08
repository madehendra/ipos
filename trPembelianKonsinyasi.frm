VERSION 5.00
Begin VB.Form trPembelianKonsinyasi 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10860
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   20265
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10860
   ScaleWidth      =   20265
End
Attribute VB_Name = "trPembelianKonsinyasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BiSAButton1_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  
  trPembelianKonsinyasi.Width = 1357 * Screen.TwipsPerPixelX
  trPembelianKonsinyasi.Height = 635 * Screen.TwipsPerPixelY
  CenterForm Me
  MsgBox Me.Width & " x " & Me.Height
  
  Frame1.Left = (trPembelianKonsinyasi.ScaleWidth - Frame1.Width) * 0.5
  
End Sub
