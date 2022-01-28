VERSION 5.00
Begin VB.Form frmMapReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Liste des Maps"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   3810
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Report"
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.CommandButton cmdSetMap 
         Caption         =   "Changez de Map"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   2520
         Width           =   3375
      End
      Begin VB.ListBox lstMap 
         Height          =   2205
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmMapReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSetMap_Click()
Dim Index As Long

Index = lstMap.ListIndex + 1

If Map_Changed Then
    If MsgBox("Es-tu sûr de vouloir changer de map sans sauvegarder ?", vbYesNo, "Editeur de Map") = vbNo Then
        Exit Sub
    End If
End If
Call SetCurrentMap(Index)
    
End Sub
