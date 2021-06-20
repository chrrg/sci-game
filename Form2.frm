VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function AppPath() As String 'ÎÞ\
AppPath = App.Path
If Right(AppPath, 1) = "\" Then AppPath = Left(AppPath, Len(AppPath) - 1)
End Function
Private Sub Form_Load()
Me.Picture = LoadPicture(AppPath & "\pic\b.bmp")
End Sub
