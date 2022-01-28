Attribute VB_Name = "modGameLogic"
Option Explicit

Public Sub MainLoop()
Dim Tick As Long, TickFPS As Long, FPS As Long, FrameTime As Long
Dim tmr500 As Long
Dim FogTmr As Long, TilesTimer As Long


Do While True
    Tick = TimeGetTime                            ' Set the inital tick
    ElapsedTime = Tick - FrameTime                 ' Set the time difference for time-based movement
    FrameTime = Tick                               ' Set the time second loop time to the first.
    
    ' fog scrolling
    If Map.Fog > 0 Then
        If FogTmr < Tick Then
            If Map.FogSpeed > 0 Then
                ' move
                fogOffsetX = fogOffsetX - 1
                fogOffsetY = fogOffsetY - 1
                ' reset
                If fogOffsetX < -Tex_Fog(Map.Fog).Width Then fogOffsetX = 0
                If fogOffsetY < -Tex_Fog(Map.Fog).Height Then fogOffsetY = 0
                FogTmr = Tick + 255 - Map.FogSpeed
            End If
        End If
    End If
    
    'Animated Tiles sans autotile
    If Tick > TilesTimer Then
        If CurTiles = 0 Then
            CurTiles = 1
        Else
            CurTiles = 0
        End If
        TilesTimer = Tick + 250
    End If
    
    If tmr500 < Tick Then
        ' animate waterfalls
        Select Case waterfallFrame
            Case 0
                waterfallFrame = 1
            Case 1
                waterfallFrame = 2
            Case 2
                waterfallFrame = 0
        End Select
        
        ' animate autotiles
        Select Case autoTileFrame
            Case 0
                autoTileFrame = 1
            Case 1
                autoTileFrame = 2
            Case 2
                autoTileFrame = 0
        End Select
        
        tmr500 = Tick + 500
    End If

    '###################################
    '### RENDER GRAPHICS
    '###################################
    Call Render_Graphics
    
    DoEvents
    
    ' Lock fps
    Do While TimeGetTime < Tick + 15
        DoEvents
        Sleep 1
    Loop
    
    ' Calculate fps
    If TickFPS < Tick Then
        GameFPS = FPS
        TickFPS = Tick + 1000
        FPS = 0
    Else
        FPS = FPS + 1
    End If
    
Loop

Call DestroyGame

End Sub

Public Function TwipsToPixels(ByVal twip_val As Long, ByVal XorY As Byte) As Long

    If XorY = 0 Then
        TwipsToPixels = twip_val / Screen.TwipsPerPixelX
    ElseIf XorY = 1 Then
        TwipsToPixels = twip_val / Screen.TwipsPerPixelY
    End If

End Function

Public Function PixelsToTwips(ByVal pixel_val As Long, ByVal XorY As Byte) As Long

    If XorY = 0 Then
        PixelsToTwips = pixel_val * Screen.TwipsPerPixelX
    ElseIf XorY = 1 Then
        PixelsToTwips = pixel_val * Screen.TwipsPerPixelY
    End If

End Function

' BitWise Operators for directional blocking
Public Sub setDirBlock(ByRef blockvar As Byte, ByRef Dir As Byte, ByVal Block As Boolean)

    If Block Then
        blockvar = blockvar Or (2 ^ Dir)
    Else
        blockvar = blockvar And Not (2 ^ Dir)
    End If
    
End Sub

Public Function isDirBlocked(ByRef blockvar As Byte, ByRef Dir As Byte) As Boolean

    If Not blockvar And (2 ^ Dir) Then
        isDirBlocked = False
    Else
        isDirBlocked = True
    End If
    
End Function
Function GetTryDir() As Byte
    
GetTryDir = MAX_BYTE


'dir_up_left
If GetAsyncKeyState(vbKeyNumpad7) < 0 Then
    GetTryDir = DIR_UP_LEFT
    Exit Function
End If

'dir_up_right
If GetAsyncKeyState(vbKeyNumpad9) < 0 Then
    GetTryDir = DIR_UP_RIGHT
    Exit Function
End If

'dir_down_left
If GetAsyncKeyState(vbKeyNumpad1) < 0 Then
    GetTryDir = DIR_DOWN_LEFT
    Exit Function
End If

'dir_down_right
If GetAsyncKeyState(vbKeyNumpad3) < 0 Then
    GetTryDir = DIR_DOWN_RIGHT
    Exit Function
End If

'dir_up
If GetAsyncKeyState(vbKeyNumpad8) < 0 Then
    GetTryDir = DIR_UP
    Exit Function
End If

'dir_down
If GetAsyncKeyState(vbKeyNumpad2) < 0 Then
    GetTryDir = DIR_DOWN
    Exit Function
End If

'dir_left
If GetAsyncKeyState(vbKeyNumpad4) < 0 Then
    GetTryDir = DIR_LEFT
    Exit Function
End If

'dir_right
If GetAsyncKeyState(vbKeyNumpad6) < 0 Then
    GetTryDir = DIR_RIGHT
    Exit Function
End If


End Function

Function GetOpposite(ByVal Dir As Byte) As Byte

Select Case Dir
    Case DIR_UP
        GetOpposite = DIR_DOWN
    Case DIR_LEFT
        GetOpposite = DIR_RIGHT
    Case DIR_DOWN
        GetOpposite = DIR_UP
    Case DIR_RIGHT
        GetOpposite = DIR_LEFT
    Case DIR_UP_LEFT
        GetOpposite = DIR_DOWN_RIGHT
    Case DIR_UP_RIGHT
        GetOpposite = DIR_DOWN_LEFT
    Case DIR_DOWN_LEFT
        GetOpposite = DIR_UP_RIGHT
    Case DIR_DOWN_RIGHT
        GetOpposite = DIR_UP_LEFT
End Select

End Function

Public Sub ProcessMovementCamera(ByVal KeyCode As Long)

'###################################
'## Déplacement sur la map
'###################################

' Déplacement à droite
If KeyCode = vbKeyRight Then
    If (CurrentX + 1 <= (Map.MaxX - ((ScreenWidth / 32) - 1))) Then
        CurrentX = CurrentX + 1
    End If
End If

' Déplacement à gauche
If KeyCode = vbKeyLeft Then
    If (CurrentX - 1 >= 0) Then
        CurrentX = CurrentX - 1
    End If
End If

' Déplacement vers le haut
If KeyCode = vbKeyUp Then
    If (CurrentY - 1 >= 0) Then
        CurrentY = CurrentY - 1
    End If
End If

' Déplacement vers le bas
If KeyCode = vbKeyDown Then
    If (CurrentY + 1 <= (Map.MaxY - ((ScreenHeight / 32) - 1))) Then
        CurrentY = CurrentY + 1
    End If
End If
End Sub



Public Sub TakeScreenshot()
Dim Surface As Direct3DSurface8, srcPalette As PALETTEENTRY, tmpSurface As Direct3DSurface8, oldSurface As Direct3DSurface8
Dim sRect As RECT
Dim tmpDesc As D3DSURFACE_DESC
Dim X As Integer, Y As Integer, i As Integer
Dim tmpX As Byte, tmpY As Byte

If Not CurrentX = 0 Or Not CurrentY = 0 Then
    With TileView
        .Top = 0
        .bottom = Map.MaxY
        .Left = 0
        .Right = Map.MaxX
    End With
End If

Set Surface = Direct3D_Device.CreateRenderTarget((Map.MaxX + 1) * 32, (Map.MaxY + 1) * 32, D3DFMT_A8R8G8B8, D3DMULTISAMPLE_NONE, 0)

Direct3D_Device.SetRenderTarget Surface, tmpSurface, ByVal 0

With sRect
    .Top = 0
    .bottom = (Map.MaxY + 1) * 32
    .Left = 0
    .Right = (Map.MaxX + 1) * 32
End With

Call Direct3D_Device.Clear(0, ByVal 0, D3DCLEAR_TARGET, White, 1#, 0)
Direct3D_Device.BeginScene
    
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            DrawMapTile X, Y
            DrawMapFringeTile X, Y
        Next
    Next
    
Direct3D_Device.EndScene
Direct3D_Device.SetRenderTarget Direct3D_Device.GetBackBuffer(0, D3DBACKBUFFER_TYPE_MONO), oldSurface, ByVal 0

Do While FileExist("\screenshot" & i & ".bmp")
    i = i + 1
Loop

Direct3DX.SaveSurfaceToFile App.path & "\screenshot" & i & ".bmp", D3DXIFF_BMP, Surface, srcPalette, sRect
Call MsgBox("Screenshot done!", vbOKOnly, "Editeur de Map")
End Sub


