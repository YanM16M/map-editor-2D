VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editeur de Map"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   587
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   932
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Unload(Cancel As Integer)
    Map_Changed = True
    If Map_Changed Then
        If MsgBox("Etes-vous sûr de vouloir quitter sans sauvegarder ?", vbYesNo, "Editeur de Map") = vbYes Then
            DestroyGame
        Else
            Cancel = True
            Exit Sub
        End If
    End If
    
    DestroyGame
End Sub

Private Sub Form_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
    If GettingMap Then Exit Sub
    HandleMouseDown button
End Sub

Private Sub Form_MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
    If GettingMap Then Exit Sub
    HandleMouseUp button
End Sub

Private Sub Form_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
    If GettingMap Then Exit Sub
    HandleMouseMove CLng(X), CLng(Y), button
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If GettingMap Then Exit Sub
    HandleKeyDown KeyCode
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If GettingMap Then Exit Sub
    HandleKeyUp KeyCode
End Sub
