Attribute VB_Name = "modInput"
Option Explicit

' keyboard input
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Sub HandleKeyDown(ByVal KeyCode As Long)

If GettingMap Then Exit Sub

'###################################
'## Déplacement sur la map
'###################################

Call ProcessMovementCamera(KeyCode)

'###################################
'## Raccourci dirblock
'###################################

If InMapEditor And frmEditor_Map.chkDirBlock.Value Then
    Dim TryDirs As Byte
    TryDirs = GetTryDir
    If TryDirs <> 255 Then
        setDirBlock EditorTile(frmEditor_Map.scrlTileSet, EditorTileX, EditorTileY).DirBlock, CByte(TryDirs), Not isDirBlocked(EditorTile(frmEditor_Map.scrlTileSet, EditorTileX, EditorTileY).DirBlock, CByte(TryDirs))
    End If
End If

End Sub

Public Sub HandleMouseMove(ByVal X As Long, ByVal Y As Long, ByVal button As Long)

If GettingMap Then Exit Sub

' Set the global cursor position
GlobalX = X
GlobalY = Y
GlobalX_Map = (TileView.Left * 32) + Camera.Left
GlobalY_Map = GlobalY + (TileView.Top * 32) + Camera.Top

' Handle the events
CurX = TileView.Left + ((X + Camera.Left) \ 32)
CurY = TileView.Top + ((Y + Camera.Top) \ 32)

If InMapEditor Then
    If button = vbLeftButton Or button = vbRightButton Then
        Call MapEditorMouseDown(button, X, Y)
    End If
    frmEditor_Map.lblX.Caption = "CurX : " & CurX
    frmEditor_Map.lblY.Caption = "CurY : " & CurY
End If
    
End Sub

Public Sub HandleMouseDown(ByVal button As Long)
    If GettingMap Then Exit Sub
    Call MapEditorMouseDown(button, GlobalX, GlobalY, False)
End Sub

Public Sub HandleMouseUp(ByVal button As Long)
Dim I As Long
    If GettingMap Then Exit Sub
    For I = 1 To MAX_BYTE
        LastClick(I) = False
    Next

End Sub

Public Sub HandleKeyUp(ByVal KeyCode As Long)

Call ProcessMovementCamera(KeyCode)

End Sub
