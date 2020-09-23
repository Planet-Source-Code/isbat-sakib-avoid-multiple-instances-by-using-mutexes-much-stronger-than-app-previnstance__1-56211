VERSION 5.00
Begin VB.Form frmDemo 
   Caption         =   "Demo Test Form"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    MutexCleanUp       'This sub should be called at the end of the program, though
                       'as much as I understand, it might not cause much trouble if
                       'you don't call it. However, it is recommended that you do call
                       'this sub as it is somewhat a rule of mutex manipulation.

End Sub
