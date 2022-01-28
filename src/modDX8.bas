Attribute VB_Name = "modDX8"
Option Explicit
' **********************
' ** Renders graphics **
' **********************
' DirectX8 Object
Private DirectX8 As DirectX8 'The master DirectX object.
Private Direct3D As Direct3D8 'Controls all things 3D.
Public Direct3D_Device As Direct3DDevice8 'Represents the hardware rendering.
Public Direct3DX As D3DX8

Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long

'The 2D (Transformed and Lit) vertex format.
Public Const FVF_TLVERTEX As Long = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE

'The 2D (Transformed and Lit) vertex format type.
Public Type TLVERTEX
    X As Single
    Y As Single
    Z As Single
    RHW As Single
    Color As Long
    TU As Single
    TV As Single
End Type

Private Vertex_List(3) As TLVERTEX '4 vertices will make a square.

'Some color depth constants to help make the DX constants more readable.
Private Const COLOR_DEPTH_16_BIT As Long = D3DFMT_R5G6B5
Private Const COLOR_DEPTH_24_BIT As Long = D3DFMT_A8R8G8B8
Private Const COLOR_DEPTH_32_BIT As Long = D3DFMT_X8R8G8B8

Public RenderingMode As Long

Private Direct3D_Window As D3DPRESENT_PARAMETERS 'Backbuffer and viewport description.
Private Display_Mode As D3DDISPLAYMODE

Public ScreenWidth As Long
Public ScreenHeight As Long

'Graphic Textures

' Tableaux
Public Tex_Tileset() As DX8TextureRec
Public Tex_Fog() As DX8TextureRec
Public Tex_Misc() As DX8TextureRec

' Number of graphic files
Public NumTilesets As Long
Public NumFogs As Long
Public NumMiscs As Long

Public Type DX8TextureRec
    Texture As Long
    Width As Long
    Height As Long
    filepath As String
    TexWidth As Long
    TexHeight As Long
    ImageData() As Byte
    MaxAnim As Byte
    HasData As Boolean
End Type

Public Type GlobalTextureRec
    Texture As Direct3DTexture8
    TexWidth As Long
    TexHeight As Long
    Loaded As Boolean
    UnloadTimer As Long
End Type

'MAP
Public ScreenX As Integer, ScreenY As Integer
Public TileWidth As Long, TileHeight As Long

Public gTexture() As GlobalTextureRec
Public NumTextures As Long

' ********************
' ** Initialization **
' ********************
Public Function InitDX8() As Boolean

    Set DirectX8 = New DirectX8 'Creates the DirectX object.
    Set Direct3D = DirectX8.Direct3DCreate() 'Creates the Direct3D object using the DirectX object.
    Set Direct3DX = New D3DX8
    
    frmMain.Width = PixelsToTwips(ScreenWidth, 0)
    frmMain.Height = PixelsToTwips(ScreenHeight, 1)
    
    'set resolution
    TileWidth = (ScreenWidth / 32) - 1
    TileHeight = (ScreenHeight / 32) - 2
    ScreenX = (TileWidth) * 32
    ScreenY = (TileHeight) * 32
   
    
    Direct3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Display_Mode 'Use the current display mode that you
                                                                    'are already on. Incase you are confused, I'm
                                                                    'talking about your current screen resolution. ;)
    Direct3D_Window.Windowed = True 'The app will be in windowed mode.
    
    Direct3D_Window.SwapEffect = D3DSWAPEFFECT_DISCARD 'Refresh when the monitor does.
    Direct3D_Window.BackBufferFormat = Display_Mode.Format 'Sets the format that was retrieved into the backbuffer.
    'Creates the rendering device with some useful info, along with the info
    'DispMode.Format = D3DFMT_X8R8G8B8
    Direct3D_Window.SwapEffect = D3DSWAPEFFECT_COPY
    Direct3D_Window.BackBufferCount = 1 '1 backbuffer only
    Direct3D_Window.BackBufferWidth = ScreenWidth ' frmMain.ScaleWidth 'Match the backbuffer width with the display width
    Direct3D_Window.BackBufferHeight = ScreenHeight 'frmMain.Scaleheight 'Match the backbuffer height with the display height
    Direct3D_Window.hDeviceWindow = frmMain.hWnd 'Use frmMain as the device window.
    
    'we've already setup for Direct3D_Window.
    If TryCreateDirectX8Device = False Then
        MsgBox "Unable to initialize DirectX8. You may be missing dx8vb.dll or have incompatible hardware to use DirectX8."
        DestroyGame
    End If

    With Direct3D_Device
        .SetVertexShader D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE
    
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        .SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        .SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
        .SetRenderState D3DRS_ZENABLE, False
        .SetRenderState D3DRS_ZWRITEENABLE, False
        
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    
        .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_POINT
        .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_POINT
        .SetTextureStageState 0, D3DTSS_MIPFILTER, D3DTEXF_NONE
    End With
    
    ' Initialise the surfaces
    LoadTextures
    
    ' We're done
    InitDX8 = True
End Function

Function TryCreateDirectX8Device() As Boolean
Dim I As Long

On Error GoTo nexti

    For I = 1 To 4
        Select Case I
            Case 1
                Set Direct3D_Device = Direct3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, Direct3D_Window)
                TryCreateDirectX8Device = True
                Exit Function
            'Case 2
            '    Set Direct3D_Device = Direct3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.hwnd, D3DCREATE_MIXED_VERTEXPROCESSING, Direct3D_Window)
            '    TryCreateDirectX8Device = True
            '    Exit Function
            Case 2
                Set Direct3D_Device = Direct3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, Direct3D_Window)
                TryCreateDirectX8Device = True
                Exit Function
            Case 3
                TryCreateDirectX8Device = False
                Exit Function
        End Select
nexti:
    Next

End Function

Public Sub DestroyDX8()
    UnloadTextures
    Set Direct3DX = Nothing
    Set Direct3D_Device = Nothing
    Set Direct3D = Nothing
    Set DirectX8 = Nothing
End Sub


Function GetNearestPOT(Value As Long) As Long
Dim I As Long
    Do While 2 ^ I < Value
        I = I + 1
    Loop
    GetNearestPOT = 2 ^ I
End Function
Public Sub SetTexture(ByRef TextureRec As DX8TextureRec, Optional isCharacter As Boolean = False)

If TextureRec.Texture > NumTextures Then TextureRec.Texture = NumTextures
If TextureRec.Texture < 0 Then TextureRec.Texture = 0

If Not TextureRec.Texture = 0 Then
    If Not gTexture(TextureRec.Texture).Loaded Then
        Call LoadTexture(TextureRec)
        If isCharacter Then
            TextureRec.MaxAnim = TextureRec.Width / (TextureRec.Height / 4)
        End If
    End If
End If

End Sub
Public Sub LoadTexture(ByRef TextureRec As DX8TextureRec)
Dim SourceBitmap As cGDIpImage, ConvertedBitmap As cGDIpImage, GDIGraphics As cGDIpRenderer, GDIToken As cGDIpToken, I As Long
Dim newWidth As Long, newHeight As Long, ImageData() As Byte, fn As Long
    
    Dim newFileName As String
    
    newFileName = TextureRec.filepath
    newFileName = Left(newFileName, Len(newFileName) - 4)
    
    Encryption_RC4_DecryptFile newFileName & ".fight", TextureRec.filepath, "Freebox"

    If TextureRec.HasData = False Then
        Set GDIToken = New cGDIpToken
        If GDIToken.Token = 0& Then MsgBox "GDI+ failed to load, exiting game!": DestroyGame
        Set SourceBitmap = New cGDIpImage
        Call SourceBitmap.LoadPicture_FileName(TextureRec.filepath, GDIToken)
        
        TextureRec.Width = SourceBitmap.Width
        TextureRec.Height = SourceBitmap.Height
        
        newWidth = GetNearestPOT(TextureRec.Width)
        newHeight = GetNearestPOT(TextureRec.Height)
        If newWidth <> SourceBitmap.Width Or newHeight <> SourceBitmap.Height Then
            Set ConvertedBitmap = New cGDIpImage
            Set GDIGraphics = New cGDIpRenderer
            I = GDIGraphics.CreateGraphicsFromImageClass(SourceBitmap)
            Call ConvertedBitmap.LoadPicture_FromNothing(newHeight, newWidth, I, GDIToken) 'I HAVE NO IDEA why this is backwards but it works.
            Call GDIGraphics.DestroyHGraphics(I)
            I = GDIGraphics.CreateGraphicsFromImageClass(ConvertedBitmap)
            Call GDIGraphics.AttachTokenClass(GDIToken)
            Call GDIGraphics.RenderImageClassToHGraphics(SourceBitmap, I)
            Call ConvertedBitmap.SaveAsPNG(ImageData)
            GDIGraphics.DestroyHGraphics (I)
            TextureRec.ImageData = ImageData
            Set ConvertedBitmap = Nothing
            Set GDIGraphics = Nothing
            Set SourceBitmap = Nothing
        Else
            Call SourceBitmap.SaveAsPNG(ImageData)
            TextureRec.ImageData = ImageData
            Set SourceBitmap = Nothing
        End If
    Else
        ImageData = TextureRec.ImageData
    End If
    
    
    Set gTexture(TextureRec.Texture).Texture = Direct3DX.CreateTextureFromFileInMemoryEx(Direct3D_Device, _
                                                    ImageData(0), _
                                                    UBound(ImageData) + 1, _
                                                    newWidth, _
                                                    newHeight, _
                                                    D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, ByVal (0), ByVal 0, ByVal 0)
    
    gTexture(TextureRec.Texture).TexWidth = newWidth
    gTexture(TextureRec.Texture).TexHeight = newHeight
    gTexture(TextureRec.Texture).Loaded = True
    gTexture(TextureRec.Texture).UnloadTimer = TimeGetTime
    
    Kill TextureRec.filepath
End Sub

Public Sub LoadTextures()

Tex_Tileset = CheckFolder(NumTilesets, "tilesets\")
Tex_Fog = CheckFolder(NumFogs, "fogs\")
Tex_Misc = CheckFolder(NumMiscs, "misc\")

ReDim Preserve gTexture(NumTextures)

EngineInitFontTextures

End Sub

Public Sub UnloadTexture(ByVal Index As Long)

If Index < 1 Or Index > NumTextures Then Exit Sub

Set gTexture(Index).Texture = Nothing
ZeroMemory ByVal VarPtr(gTexture(Index)), LenB(gTexture(Index))
gTexture(Index).UnloadTimer = 0
gTexture(Index).Loaded = False

End Sub
Public Sub UnloadTextures(Optional ByVal Complete As Boolean = False)
Dim I As Long
    
    ' If debug mode, handle error then exit out
    On Error Resume Next
    
    If Complete = False Then
        For I = 1 To NumTextures
            If gTexture(I).UnloadTimer > TimeGetTime + 150000 Then
                Set gTexture(I).Texture = Nothing
                ZeroMemory ByVal VarPtr(gTexture(I)), LenB(gTexture(I))
                gTexture(I).UnloadTimer = 0
                gTexture(I).Loaded = False
            End If
        Next
    Else
    
        For I = 1 To NumTextures
            Set gTexture(I).Texture = Nothing
            ZeroMemory ByVal VarPtr(gTexture(I)), LenB(gTexture(I))
        Next
        
        UnloadFontTextures
        
        ReDim gTexture(1)
    
        
        For I = 1 To NumTilesets
            Tex_Tileset(I).Texture = 0
        Next
        
        For I = 1 To NumMiscs
            Tex_Misc(I).Texture = 0
        Next
        
        For I = 1 To NumFogs
            Tex_Fog(I).Texture = 0
        Next
        
    End If

End Sub

'################################
'## Drawing
'################################

'This function will make it much easier to setup the vertices with the info it needs.
Private Function Create_TLVertex(X As Single, Y As Single, Z As Single, RHW As Single, Color As Long, Specular As Long, TU As Single, TV As Single) As TLVERTEX

    Create_TLVertex.X = X
    Create_TLVertex.Y = Y
    Create_TLVertex.Z = Z
    Create_TLVertex.RHW = RHW
    Create_TLVertex.Color = Color
    'Create_TLVertex.Specular = Specular
    Create_TLVertex.TU = TU
    Create_TLVertex.TV = TV
    
End Function

Public Sub RenderTexture(ByRef TextureRec As DX8TextureRec, ByVal dX As Single, ByVal dY As Single, ByVal sx As Single, ByVal sy As Single, ByVal dWidth As Single, ByVal dHeight As Single, ByVal sWidth As Single, ByVal sHeight As Single, Optional Color As Long = -1, Optional ByVal Degrees As Single = 0)
    Dim TextureNum As Long
    Dim textureWidth As Long, textureHeight As Long, sourceX As Single, sourceY As Single, sourceWidth As Single, sourceHeight As Single
    Dim RadAngle As Single 'The angle in Radians
    Dim CenterX As Single
    Dim CenterY As Single
    Dim NewX As Single
    Dim NewY As Single
    Dim SinRad As Single
    Dim CosRad As Single
    Dim I As Long
    
    SetTexture TextureRec
    
    TextureNum = TextureRec.Texture
    
    textureWidth = gTexture(TextureNum).TexWidth
    textureHeight = gTexture(TextureNum).TexHeight
    
    If sy + sHeight > textureHeight Then Exit Sub
    If sx + sWidth > textureWidth Then Exit Sub
    If sx < 0 Then Exit Sub
    If sy < 0 Then Exit Sub

    sx = sx - 0.5
    sy = sy - 0.5
    dY = dY - 0.5
    dX = dX - 0.5
    sWidth = sWidth
    sHeight = sHeight
    dWidth = dWidth
    dHeight = dHeight
    sourceX = (sx / textureWidth)
    sourceY = (sy / textureHeight)
    sourceWidth = ((sx + sWidth) / textureWidth)
    sourceHeight = ((sy + sHeight) / textureHeight)
    
    Vertex_List(0) = Create_TLVertex(dX, dY, 0, 1, Color, 0, sourceX + 0.000003, sourceY + 0.000003)
    Vertex_List(1) = Create_TLVertex(dX + dWidth, dY, 0, 1, Color, 0, sourceWidth + 0.000003, sourceY + 0.000003)
    Vertex_List(2) = Create_TLVertex(dX, dY + dHeight, 0, 1, Color, 0, sourceX + 0.000003, sourceHeight + 0.000003)
    Vertex_List(3) = Create_TLVertex(dX + dWidth, dY + dHeight, 0, 1, Color, 0, sourceWidth + 0.000003, sourceHeight + 0.000003)
    
    'Check if a rotation is required
    If Degrees <> 0 And Degrees <> 360 Then

        'Converts the angle to rotate by into radians
        RadAngle = Degrees * DegreeToRadian

        'Set the CenterX and CenterY values
        CenterX = dX + (dWidth * 0.5)
        CenterY = dY + (dHeight * 0.5)

        'Pre-calculate the cosine and sine of the radiant
        SinRad = Sin(RadAngle)
        CosRad = Cos(RadAngle)

        'Loops through the passed vertex buffer
        For I = 0 To 3

            'Calculates the new X and Y co-ordinates of the vertices for the given angle around the center co-ordinates
            NewX = CenterX + (Vertex_List(I).X - CenterX) * CosRad - (Vertex_List(I).Y - CenterY) * SinRad
            NewY = CenterY + (Vertex_List(I).Y - CenterY) * CosRad + (Vertex_List(I).X - CenterX) * SinRad

            'Applies the new co-ordinates to the buffer
            Vertex_List(I).X = NewX
            Vertex_List(I).Y = NewY
        Next
    End If
    
    Call Direct3D_Device.SetTexture(0, gTexture(TextureNum).Texture)
    Direct3D_Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex_List(0), Len(Vertex_List(0))
End Sub

Public Sub Render_Graphics()
Dim I As Long, X As Long, Y As Long
Dim rec As RECT
Dim rec_pos As RECT, srcRect As D3DRECT

    ' If debug mode, handle error then exit out
    On Error GoTo ErrorHandler
   
    'Check for device lost.
    If Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST Or Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then HandleDeviceLost: Exit Sub
    If frmMain.WindowState = vbMinimized Or GettingMap Then Exit Sub
    
    'update the viewpoint
    Call UpdateCamera
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorARGB(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    
    ' on dessine les nuages
    If Map.FogOpacity = 0 Then
        DrawFog
    End If
            
    ' blit lower tiles
    If NumTilesets > 0 Then
        For X = TileView.Left To TileView.Right
            For Y = TileView.Top To TileView.bottom
                If IsValidMapPoint(X, Y) Then
                
                    Call DrawMapTile(X, Y)
                    
                End If
            Next
        Next
    End If
    
    ' blit out upper tiles
    If NumTilesets > 0 Then
        For X = TileView.Left To TileView.Right
            For Y = TileView.Top To TileView.bottom
                If IsValidMapPoint(X, Y) Then
                
                    Call DrawMapFringeTile(X, Y) 'on dessine les fringes
                    
                End If
            Next
        Next
    End If
    
    
    If Map.FogOpacity > 0 Then DrawFog
    DrawTint
    
    ' blit out a square at mouse cursor
    If InMapEditor Then
        If frmEditor_Map.optBlock.Value = True Then
            For X = TileView.Left To TileView.Right
                For Y = TileView.Top To TileView.bottom
                    If IsValidMapPoint(X, Y) Then
                        Call DrawDirection(X, Y)
                    End If
                Next
            Next
        End If
        Call DrawTileOutline
        Call DrawTileApercu
    End If
    
    Call DrawMapAttributes
    
    ' draw cursor, player X and Y locations
    If BloC Then
        RenderText Font_Default, Trim$("Gx: " & GlobalX & " Gy: " & GlobalY), ScreenWidth - 120, 70, Red
        RenderText Font_Default, Trim$("cur x: " & CurX & " y: " & CurY), ScreenWidth - 120, 84, Yellow
        RenderText Font_Default, Trim$(" (map #" & CurrentMap & ")"), ScreenWidth - 120, 112, Yellow
    End If
  
    ' Get rec
    With rec
        .Top = Camera.Top
        .bottom = .Top + ScreenY
        .Left = Camera.Left
        .Right = .Left + ScreenX
    End With
        
    ' rec_pos
    With rec_pos
        .bottom = ScreenY
        .Right = ScreenX
    End With
        
    With srcRect
        .X1 = 0
        .X2 = frmMain.ScaleWidth
        .Y1 = 0
        .Y2 = frmMain.ScaleHeight
    End With
    
    Direct3D_Device.EndScene
        
    If Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST Or Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        HandleDeviceLost
        Exit Sub
    Else
        Direct3D_Device.Present srcRect, ByVal 0, 0, ByVal 0
        DrawGDI
    End If

    ' Error handler
    Exit Sub
ErrorHandler:
    If Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST Or Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        HandleDeviceLost
        Exit Sub
    Else
        MsgBox "Unrecoverable DX8 error." & Err.Number
        DestroyGame
    End If
End Sub

Sub HandleDeviceLost()
'Do a loop while device is lost
   Do While Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST
       Exit Sub
   Loop
   
   UnloadTextures True
   
   'Reset the device
   Direct3D_Device.Reset Direct3D_Window
   
   DirectX_ReInit
    
   LoadTextures
   
End Sub

Public Function DirectX_ReInit() As Boolean
Dim Width As Integer, Height As Integer
    On Error GoTo Error_Handler
    
    frmMain.Width = PixelsToTwips(ScreenWidth, 0)
    frmMain.Height = PixelsToTwips(ScreenHeight, 1)
    
     'set resolution
    TileWidth = (ScreenWidth / 32) - 1
    TileHeight = (ScreenHeight / 32) - 2
    ScreenX = (TileWidth) * 32
    ScreenY = (TileHeight) * 32

    Direct3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Display_Mode 'Use the current display mode that you
                                                                    'are already on. Incase you are confused, I'm
                                                                    'talking about your current screen resolution. ;)
        
    Direct3D_Window.Windowed = True 'The app will be in windowed mode.

    Direct3D_Window.SwapEffect = D3DSWAPEFFECT_COPY 'Refresh when the monitor does.
    Direct3D_Window.BackBufferFormat = Display_Mode.Format 'Sets the format that was retrieved into the backbuffer.
    'Creates the rendering device with some useful info, along with the info
    'we've already setup for Direct3D_Window.
    'Creates the rendering device with some useful info, along with the info
    Direct3D_Window.BackBufferCount = 1 '1 backbuffer only
    Direct3D_Window.BackBufferWidth = Width ' frmMain.ScaleWidth 'Match the backbuffer width with the display width
    Direct3D_Window.BackBufferHeight = Height 'frmMain.Scaleheight 'Match the backbuffer height with the display height
    Direct3D_Window.hDeviceWindow = frmMain.hWnd 'Use frmMain as the device window.
    
    With Direct3D_Device
        .SetVertexShader D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE
    
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        .SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        .SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
        .SetRenderState D3DRS_ZENABLE, False
        .SetRenderState D3DRS_ZWRITEENABLE, False
        
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    
        .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_POINT
        .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_POINT
        .SetTextureStageState 0, D3DTSS_MIPFILTER, D3DTEXF_NONE
    End With
    
    
    
    DirectX_ReInit = True

    Exit Function
    
Error_Handler:
    MsgBox "An error occured while initializing DirectX", vbCritical
    
    DestroyGame
    
    DirectX_ReInit = False
End Function

Public Sub UpdateCamera()
 Dim offsetX As Long, offsetY As Long, startX As Long, startY As Long, EndX As Long, EndY As Long
    

    TileWidth = (ScreenWidth / 32) - 1
    TileHeight = (ScreenHeight / 32) - 1
    startX = CurrentX
    startY = CurrentY
    
    If startX > Map.MaxX - TileWidth Then
        startX = Map.MaxX - TileWidth
    End If
    
    If startY > Map.MaxY - TileHeight Then
        startY = Map.MaxY - TileHeight
    End If
    
    EndX = startX + (TileWidth + 1) + 1
    EndY = startY + (TileHeight + 1) + 1

    With TileView
        .Top = startY
        .bottom = EndY
        .Left = startX
        .Right = EndX
    End With

    With Camera
        .Top = offsetY
        .bottom = .Top + ScreenY
        .Left = offsetX
        .Right = .Left + ScreenX
    End With
    
End Sub

Public Function ConvertMapX(ByVal X As Long) As Long

    ConvertMapX = X - (TileView.Left * 32) - Camera.Left

End Function

Public Function ConvertMapY(ByVal Y As Long) As Long

    ConvertMapY = Y - (TileView.Top * 32) - Camera.Top
    
End Function

Public Function InViewPort(ByVal X As Long, ByVal Y As Long) As Boolean

    InViewPort = False

    If X < TileView.Left Then Exit Function
    If Y < TileView.Top Then Exit Function
    If X > TileView.Right Then Exit Function
    If Y > TileView.bottom Then Exit Function
    InViewPort = True

End Function

Public Function IsValidMapPoint(ByVal X As Long, ByVal Y As Long) As Boolean

    IsValidMapPoint = False

    If X < 0 Then Exit Function
    If Y < 0 Then Exit Function
    If X > Map.MaxX Then Exit Function
    If Y > Map.MaxY Then Exit Function
    IsValidMapPoint = True

End Function


'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'   All of this code is for auto tiles and the math behind generating them.
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Public Sub placeAutotile(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long, ByVal tileQuarter As Byte, ByVal autoTileLetter As String)
    With Autotile(X, Y).layer(layerNum).QuarterTile(tileQuarter)
        Select Case autoTileLetter
            Case "a"
                .X = autoInner(1).X
                .Y = autoInner(1).Y
            Case "b"
                .X = autoInner(2).X
                .Y = autoInner(2).Y
            Case "c"
                .X = autoInner(3).X
                .Y = autoInner(3).Y
            Case "d"
                .X = autoInner(4).X
                .Y = autoInner(4).Y
            Case "e"
                .X = autoNW(1).X
                .Y = autoNW(1).Y
            Case "f"
                .X = autoNW(2).X
                .Y = autoNW(2).Y
            Case "g"
                .X = autoNW(3).X
                .Y = autoNW(3).Y
            Case "h"
                .X = autoNW(4).X
                .Y = autoNW(4).Y
            Case "i"
                .X = autoNE(1).X
                .Y = autoNE(1).Y
            Case "j"
                .X = autoNE(2).X
                .Y = autoNE(2).Y
            Case "k"
                .X = autoNE(3).X
                .Y = autoNE(3).Y
            Case "l"
                .X = autoNE(4).X
                .Y = autoNE(4).Y
            Case "m"
                .X = autoSW(1).X
                .Y = autoSW(1).Y
            Case "n"
                .X = autoSW(2).X
                .Y = autoSW(2).Y
            Case "o"
                .X = autoSW(3).X
                .Y = autoSW(3).Y
            Case "p"
                .X = autoSW(4).X
                .Y = autoSW(4).Y
            Case "q"
                .X = autoSE(1).X
                .Y = autoSE(1).Y
            Case "r"
                .X = autoSE(2).X
                .Y = autoSE(2).Y
            Case "s"
                .X = autoSE(3).X
                .Y = autoSE(3).Y
            Case "t"
                .X = autoSE(4).X
                .Y = autoSE(4).Y
        End Select
    End With
End Sub

Public Sub initAutotiles()
Dim X As Long, Y As Long, layerNum As Long
    ' Procedure used to cache autotile positions. All positioning is
    ' independant from the tileset. Calculations are convoluted and annoying.
    ' Maths is not my strong point. Luckily we're caching them so it's a one-off
    ' thing when the map is originally loaded. As such optimisation isn't an issue.
    
    ' For simplicity's sake we cache all subtile SOURCE positions in to an array.
    ' We also give letters to each subtile for easy rendering tweaks. ;]
    
    ' First, we need to re-size the array
    ReDim Autotile(0 To Map.MaxX, 0 To Map.MaxY)
    
    ' Inner tiles (Top right subtile region)
    ' NW - a
    autoInner(1).X = 32
    autoInner(1).Y = 0
    
    ' NE - b
    autoInner(2).X = 48
    autoInner(2).Y = 0
    
    ' SW - c
    autoInner(3).X = 32
    autoInner(3).Y = 16
    
    ' SE - d
    autoInner(4).X = 48
    autoInner(4).Y = 16
    
    ' Outer Tiles - NW (bottom subtile region)
    ' NW - e
    autoNW(1).X = 0
    autoNW(1).Y = 32
    
    ' NE - f
    autoNW(2).X = 16
    autoNW(2).Y = 32
    
    ' SW - g
    autoNW(3).X = 0
    autoNW(3).Y = 48
    
    ' SE - h
    autoNW(4).X = 16
    autoNW(4).Y = 48
    
    ' Outer Tiles - NE (bottom subtile region)
    ' NW - i
    autoNE(1).X = 32
    autoNE(1).Y = 32
    
    ' NE - g
    autoNE(2).X = 48
    autoNE(2).Y = 32
    
    ' SW - k
    autoNE(3).X = 32
    autoNE(3).Y = 48
    
    ' SE - l
    autoNE(4).X = 48
    autoNE(4).Y = 48
    
    ' Outer Tiles - SW (bottom subtile region)
    ' NW - m
    autoSW(1).X = 0
    autoSW(1).Y = 64
    
    ' NE - n
    autoSW(2).X = 16
    autoSW(2).Y = 64
    
    ' SW - o
    autoSW(3).X = 0
    autoSW(3).Y = 80
    
    ' SE - p
    autoSW(4).X = 16
    autoSW(4).Y = 80
    
    ' Outer Tiles - SE (bottom subtile region)
    ' NW - q
    autoSE(1).X = 32
    autoSE(1).Y = 64
    
    ' NE - r
    autoSE(2).X = 48
    autoSE(2).Y = 64
    
    ' SW - s
    autoSE(3).X = 32
    autoSE(3).Y = 80
    
    ' SE - t
    autoSE(4).X = 48
    autoSE(4).Y = 80
    
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            For layerNum = 1 To MapLayer.Layer_Count - 1
                ' calculate the subtile positions and place them
                CalculateAutotile X, Y, layerNum
                ' cache the rendering state of the tiles and set them
                CacheRenderState X, Y, layerNum
            Next
        Next
    Next
End Sub

Public Sub CacheRenderState(ByVal X As Long, ByVal Y As Long, ByVal layerNum As Long)
Dim quarterNum As Long

    ' exit out early
    If X < 0 Or X > Map.MaxX Or Y < 0 Or Y > Map.MaxY Then Exit Sub

    With Map.Tile(X, Y)
        ' check if the tile can be rendered
        If .layer(layerNum).Tileset < 1 Or .layer(layerNum).Tileset > NumTilesets Then
            Autotile(X, Y).layer(layerNum).RenderState = RENDER_STATE_NONE
            Exit Sub
        End If
        
        ' check if it needs to be rendered as an autotile
        If .Autotile(layerNum) = AUTOTILE_NONE Then
            ' default to... default
            Autotile(X, Y).layer(layerNum).RenderState = RENDER_STATE_NORMAL
        ElseIf .Autotile(layerNum) = AUTOTILE_FAKE Then
            Autotile(X, Y).layer(layerNum).RenderState = RENDER_STATE_ANIMATE
        Else
            Autotile(X, Y).layer(layerNum).RenderState = RENDER_STATE_AUTOTILE
            ' cache tileset positioning
            For quarterNum = 1 To 4
                Autotile(X, Y).layer(layerNum).srcX(quarterNum) = (Map.Tile(X, Y).layer(layerNum).X * 32) + Autotile(X, Y).layer(layerNum).QuarterTile(quarterNum).X
                Autotile(X, Y).layer(layerNum).srcY(quarterNum) = (Map.Tile(X, Y).layer(layerNum).Y * 32) + Autotile(X, Y).layer(layerNum).QuarterTile(quarterNum).Y
            Next
        End If
    End With

End Sub

Public Sub CalculateAutotile(ByVal X As Long, ByVal Y As Long, ByVal layerNum As Long)
    ' Right, so we've split the tile block in to an easy to remember
    ' collection of letters. We now need to do the calculations to find
    ' out which little lettered block needs to be rendered. We do this
    ' by reading the surrounding tiles to check for matches.
    
    ' First we check to make sure an autotile situation is actually there.
    ' Then we calculate exactly which situation has arisen.
    ' The situations are "inner", "outer", "horizontal", "vertical" and "fill".
    
    ' Exit out if we don't have an auatotile
    If Map.Tile(X, Y).Autotile(layerNum) = 0 Then Exit Sub
    
    ' Okay, we have autotiling but which one?
    Select Case Map.Tile(X, Y).Autotile(layerNum)
    
        ' Normal or animated - same difference
        Case AUTOTILE_NORMAL, AUTOTILE_ANIM
            ' North West Quarter
            CalculateNW_Normal layerNum, X, Y
            
            ' North East Quarter
            CalculateNE_Normal layerNum, X, Y
            
            ' South West Quarter
            CalculateSW_Normal layerNum, X, Y
            
            ' South East Quarter
            CalculateSE_Normal layerNum, X, Y
            
        ' Cliff
        Case AUTOTILE_CLIFF
            ' North West Quarter
            CalculateNW_Cliff layerNum, X, Y
            
            ' North East Quarter
            CalculateNE_Cliff layerNum, X, Y
            
            ' South West Quarter
            CalculateSW_Cliff layerNum, X, Y
            
            ' South East Quarter
            CalculateSE_Cliff layerNum, X, Y
            
        ' Waterfalls
        Case AUTOTILE_WATERFALL
            ' North West Quarter
            CalculateNW_Waterfall layerNum, X, Y
            
            ' North East Quarter
            CalculateNE_Waterfall layerNum, X, Y
            
            ' South West Quarter
            CalculateSW_Waterfall layerNum, X, Y
            
            ' South East Quarter
            CalculateSE_Waterfall layerNum, X, Y
        
        ' Anything else
        Case Else
            ' Don't need to render anything... it's fake or not an autotile
    End Select
End Sub

' Normal autotiling
Public Sub CalculateNW_Normal(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North West
    If checkTileMatch(layerNum, X, Y, X - 1, Y - 1) Then tmpTile(1) = True
    
    ' North
    If checkTileMatch(layerNum, X, Y, X, Y - 1) Then tmpTile(2) = True
    
    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(2) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If Not tmpTile(2) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(2) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If Not tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 1, "e"
        Case AUTO_OUTER
            placeAutotile layerNum, X, Y, 1, "a"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 1, "i"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 1, "m"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 1, "q"
    End Select
End Sub

Public Sub CalculateNE_Normal(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North
    If checkTileMatch(layerNum, X, Y, X, Y - 1) Then tmpTile(1) = True
    
    ' North East
    If checkTileMatch(layerNum, X, Y, X + 1, Y - 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 2, "j"
        Case AUTO_OUTER
            placeAutotile layerNum, X, Y, 2, "b"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 2, "f"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 2, "r"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 2, "n"
    End Select
End Sub

Public Sub CalculateSW_Normal(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile(1) = True
    
    ' South West
    If checkTileMatch(layerNum, X, Y, X - 1, Y + 1) Then tmpTile(2) = True
    
    ' South
    If checkTileMatch(layerNum, X, Y, X, Y + 1) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 3, "o"
        Case AUTO_OUTER
            placeAutotile layerNum, X, Y, 3, "c"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 3, "s"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 3, "g"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 3, "k"
    End Select
End Sub

Public Sub CalculateSE_Normal(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' South
    If checkTileMatch(layerNum, X, Y, X, Y + 1) Then tmpTile(1) = True
    
    ' South East
    If checkTileMatch(layerNum, X, Y, X + 1, Y + 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 4, "t"
        Case AUTO_OUTER
            placeAutotile layerNum, X, Y, 4, "d"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 4, "p"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 4, "l"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 4, "h"
    End Select
End Sub

' Waterfall autotiling
Public Sub CalculateNW_Waterfall(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile As Boolean
    
    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, X, Y, 1, "i"
    Else
        ' Edge
        placeAutotile layerNum, X, Y, 1, "e"
    End If
End Sub

Public Sub CalculateNE_Waterfall(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile As Boolean
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, X, Y, 2, "f"
    Else
        ' Edge
        placeAutotile layerNum, X, Y, 2, "j"
    End If
End Sub

Public Sub CalculateSW_Waterfall(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile As Boolean
    
    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, X, Y, 3, "k"
    Else
        ' Edge
        placeAutotile layerNum, X, Y, 3, "g"
    End If
End Sub

Public Sub CalculateSE_Waterfall(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile As Boolean
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, X, Y, 4, "h"
    Else
        ' Edge
        placeAutotile layerNum, X, Y, 4, "l"
    End If
End Sub

' Cliff autotiling
Public Sub CalculateNW_Cliff(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North West
    If checkTileMatch(layerNum, X, Y, X - 1, Y - 1) Then tmpTile(1) = True
    
    ' North
    If checkTileMatch(layerNum, X, Y, X, Y - 1) Then tmpTile(2) = True
    
    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation - Horizontal
    If Not tmpTile(2) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(2) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(2) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 1, "e"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 1, "i"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 1, "m"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 1, "q"
    End Select
End Sub

Public Sub CalculateNE_Cliff(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North
    If checkTileMatch(layerNum, X, Y, X, Y - 1) Then tmpTile(1) = True
    
    ' North East
    If checkTileMatch(layerNum, X, Y, X + 1, Y - 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation - Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 2, "j"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 2, "f"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 2, "r"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 2, "n"
    End Select
End Sub

Public Sub CalculateSW_Cliff(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile(1) = True
    
    ' South West
    If checkTileMatch(layerNum, X, Y, X - 1, Y + 1) Then tmpTile(2) = True
    
    ' South
    If checkTileMatch(layerNum, X, Y, X, Y + 1) Then tmpTile(3) = True
    
    ' Calculate Situation - Horizontal
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 3, "o"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 3, "s"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 3, "g"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 3, "k"
    End Select
End Sub

Public Sub CalculateSE_Cliff(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' South
    If checkTileMatch(layerNum, X, Y, X, Y + 1) Then tmpTile(1) = True
    
    ' South East
    If checkTileMatch(layerNum, X, Y, X + 1, Y + 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation -  Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 4, "t"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 4, "p"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 4, "l"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 4, "h"
    End Select
End Sub

Public Function checkTileMatch(ByVal layerNum As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Boolean
    ' we'll exit out early if true
    checkTileMatch = True
    
    ' if it's off the map then set it as autotile and exit out early
    If X2 < 0 Or X2 > Map.MaxX Or Y2 < 0 Or Y2 > Map.MaxY Then
        checkTileMatch = True
        Exit Function
    End If
    
    ' fakes ALWAYS return true
    If Map.Tile(X2, Y2).Autotile(layerNum) = AUTOTILE_FAKE Then
        checkTileMatch = True
        Exit Function
    End If
    
    ' check neighbour is an autotile
    If Map.Tile(X2, Y2).Autotile(layerNum) = 0 Then
        checkTileMatch = False
        Exit Function
    End If
    
    ' check we're a matching
    If Map.Tile(X1, Y1).layer(layerNum).Tileset <> Map.Tile(X2, Y2).layer(layerNum).Tileset Then
        checkTileMatch = False
        Exit Function
    End If
    
    ' check tiles match
    If Map.Tile(X1, Y1).layer(layerNum).X <> Map.Tile(X2, Y2).layer(layerNum).X Then
        checkTileMatch = False
        Exit Function
    End If
        
    If Map.Tile(X1, Y1).layer(layerNum).Y <> Map.Tile(X2, Y2).layer(layerNum).Y Then
        checkTileMatch = False
        Exit Function
    End If
End Function

Public Sub DrawAutoTile(ByVal layerNum As Long, ByVal destX As Long, ByVal destY As Long, ByVal quarterNum As Long, ByVal X As Long, ByVal Y As Long)
Dim yOffset As Integer, xOffset As Integer

    ' calculate the offset
    Select Case Map.Tile(X, Y).Autotile(layerNum)
        Case AUTOTILE_WATERFALL
            yOffset = (waterfallFrame - 1) * 32
        Case AUTOTILE_ANIM
            xOffset = autoTileFrame * 64
        Case AUTOTILE_CLIFF
            yOffset = -32
    End Select
    
    ' Draw the quarter
    RenderTexture Tex_Tileset(Map.Tile(X, Y).layer(layerNum).Tileset), destX, destY, Autotile(X, Y).layer(layerNum).srcX(quarterNum) + xOffset, Autotile(X, Y).layer(layerNum).srcY(quarterNum) + yOffset, 16, 16, 16, 16, -1
End Sub

Public Sub DrawFog()
Dim fogNum As Byte, Color As Long, X As Long, Y As Long, RenderState As Long

    fogNum = Map.Fog
    
    If fogNum < 1 Or fogNum > NumFogs Then Exit Sub
    Color = D3DColorRGBA(255, 255, 255, 255 - Map.FogOpacity)

    RenderState = 0
    ' render state
    Select Case RenderState
        Case 1 ' Additive
            Direct3D_Device.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
            Direct3D_Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
        Case 2 ' Subtractive
            Direct3D_Device.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_SUBTRACT
            Direct3D_Device.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ZERO
            Direct3D_Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCCOLOR
    End Select
    
    SetTexture Tex_Fog(fogNum)
    
    For X = 0 To ((Map.MaxX * 32) / Tex_Fog(fogNum).Width) + 1
        For Y = 0 To ((Map.MaxY * 32) / Tex_Fog(fogNum).Height) + 1
            RenderTexture Tex_Fog(fogNum), ConvertMapX((X * Tex_Fog(fogNum).Width) + fogOffsetX), ConvertMapY((Y * Tex_Fog(fogNum).Height) + fogOffsetY), 0, 0, 256, 256, 256, 256, Color
        Next
    Next
    
    ' reset render state
    If RenderState > 0 Then
        Direct3D_Device.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        Direct3D_Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        Direct3D_Device.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
    End If
End Sub

Public Sub DrawMapTile(ByVal X As Long, ByVal Y As Long)
Dim I As Byte
Dim TileView As Boolean

With Map.Tile(X, Y)
    
    For I = MapLayer.Ground To MapLayer.Mask2
        If TileView And Not frmEditor_Map.optLayer(I) Then GoTo nex:
        
        If Autotile(X, Y).layer(I).RenderState = RENDER_STATE_NORMAL Then
            
            ' Draw normally
            RenderTexture Tex_Tileset(.layer(I).Tileset), ConvertMapX(X * 32), ConvertMapY(Y * 32), .layer(I).X * 32, .layer(I).Y * 32, 32, 32, 32, 32, -1
        
        ElseIf Autotile(X, Y).layer(I).RenderState = RENDER_STATE_ANIMATE Then
        
            'draw animate
            If CurTiles = 0 Then
                RenderTexture Tex_Tileset(.layer(I).Tileset), ConvertMapX(X * 32), ConvertMapY(Y * 32), .layer(I).X * 32, .layer(I).Y * 32, 32, 32, 32, 32, -1
            Else
                RenderTexture Tex_Tileset(.layer(I).Tileset), ConvertMapX(X * 32), ConvertMapY(Y * 32), (.layer(I).X + 1) * 32, .layer(I).Y * 32, 32, 32, 32, 32
            End If
            
        ElseIf Autotile(X, Y).layer(I).RenderState = RENDER_STATE_AUTOTILE Then
            ' Draw autotiles
            DrawAutoTile I, ConvertMapX(X * 32), ConvertMapY(Y * 32), 1, X, Y
            DrawAutoTile I, ConvertMapX((X * 32) + 16), ConvertMapY(Y * 32), 2, X, Y
            DrawAutoTile I, ConvertMapX(X * 32), ConvertMapY((Y * 32) + 16), 3, X, Y
            DrawAutoTile I, ConvertMapX((X * 32) + 16), ConvertMapY((Y * 32) + 16), 4, X, Y
        End If
nex:
    Next
End With

End Sub

Public Sub DrawMapFringeTile(ByVal X As Long, ByVal Y As Long)
Dim I As Byte
Dim TileView As Boolean

With Map.Tile(X, Y)

    For I = MapLayer.Fringe To MapLayer.Fringe2
        If TileView And Not frmEditor_Map.optLayer(I) Then GoTo nex:
        If Autotile(X, Y).layer(I).RenderState = RENDER_STATE_NORMAL Then
            ' Draw normally
            RenderTexture Tex_Tileset(.layer(I).Tileset), ConvertMapX(X * 32), ConvertMapY(Y * 32), .layer(I).X * 32, .layer(I).Y * 32, 32, 32, 32, 32, -1
        
        ElseIf Autotile(X, Y).layer(I).RenderState = RENDER_STATE_ANIMATE Then
        
            'draw animate
            If CurTiles = 0 Then
                RenderTexture Tex_Tileset(.layer(I).Tileset), ConvertMapX(X * 32), ConvertMapY(Y * 32), .layer(I).X * 32, .layer(I).Y * 32, 32, 32, 32, 32, -1
            Else
                RenderTexture Tex_Tileset(.layer(I).Tileset), ConvertMapX(X * 32), ConvertMapY(Y * 32), (.layer(I).X + 1) * 32, .layer(I).Y * 32, 32, 32, 32, 32
            End If
        
        ElseIf Autotile(X, Y).layer(I).RenderState = RENDER_STATE_AUTOTILE Then
            ' Draw autotiles
            DrawAutoTile I, ConvertMapX(X * 32), ConvertMapY(Y * 32), 1, X, Y
            DrawAutoTile I, ConvertMapX((X * 32) + 16), ConvertMapY(Y * 32), 2, X, Y
            DrawAutoTile I, ConvertMapX(X * 32), ConvertMapY((Y * 32) + 16), 3, X, Y
            DrawAutoTile I, ConvertMapX((X * 32) + 16), ConvertMapY((Y * 32) + 16), 4, X, Y
        End If
nex:
    Next
End With


End Sub


Public Sub DrawOverlay(Alpha As Byte, Red As Byte, Green As Byte, Blue As Byte)
    RenderTexture Tex_Misc(1), 0, 0, 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight, 32, 32, D3DColorARGB(Alpha, Red, Green, Blue)
End Sub

Public Sub DrawTint()
Dim Color As Long, Alpha As Byte, Blue As Byte

    Alpha = Map.Alpha
    Blue = Map.Blue
    Color = D3DColorRGBA(Map.Red, Map.Green, Blue, Alpha)
    
    If Map.IsDead Then ' Map obsucre
        Color = D3DColorRGBA(20, 20, 20, 125)
    End If
    
    RenderTexture Tex_Misc(1), 0, 0, 0, 0, ScreenWidth, ScreenHeight, 32, 32, Color
    
    If Map.IsDead Then
        Dim color2 As Long
        
        color2 = D3DColorRGBA(255, 255, 255, 250)
        
        RenderTexture Tex_Misc(2), ConvertMapX(CurrentX * 32) + 16 - 1300, ConvertMapY(CurrentY * 32) - 800, 0, 0, 2600, 1625, 2600, 1625, color2
    End If
    
End Sub


Public Sub DrawDirection(ByVal X As Long, ByVal Y As Long, Optional ByVal Tileset As Integer = 0)
Dim rec As RECT, Block As Byte
Dim I As Byte

    ' render grid
    rec.Top = 24
    rec.Left = 0
    rec.Right = rec.Left + 32
    rec.bottom = rec.Top + 32
    RenderTexture Tex_Misc(3), ConvertMapX(X * 32), ConvertMapY(Y * 32), rec.Left, rec.Top, rec.Right - rec.Left, rec.bottom - rec.Top, rec.Right - rec.Left, rec.bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255)
    
    If Tileset > 0 Then
        Block = EditorTile(Tileset, X, Y).DirBlock
    Else
        Block = Map.Tile(X, Y).DirBlock
    End If
     
    ' render dir blobs
    For I = 0 To 7
        If I < 5 And I > 0 Then
            rec.Left = (I - 1) * 8
        Else
            rec.Left = 3 * 8
        End If
        rec.Right = rec.Left + 8
        ' find out whether render blocked or not
        If Not isDirBlocked(Block, CByte(I)) Then
            rec.Top = 8
        Else
            rec.Top = 16
        End If
        rec.bottom = rec.Top + 8
        'render!
        RenderTexture Tex_Misc(3), ConvertMapX(X * 32) + DirArrowX(I), ConvertMapY(Y * 32) + DirArrowY(I), rec.Left, rec.Top, rec.Right - rec.Left, rec.bottom - rec.Top, rec.Right - rec.Left, rec.bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255)
    Next
    
End Sub

Public Sub DrawTileOutline()
Dim rec As RECT
    
If frmEditor_Map.optBlock.Value Then Exit Sub

With rec
    .Top = 0
    .bottom = .Top + 32
    .Left = 0
    .Right = .Left + 32
End With

RenderTexture Tex_Misc(5), ConvertMapX(CurX * 32), ConvertMapY(CurY * 32), rec.Left, rec.Top, rec.Right - rec.Left, rec.bottom - rec.Top, rec.Right - rec.Left, rec.bottom - rec.Top, D3DColorRGBA(100, 255, 255, 255)

End Sub

' ******************
' ** Game Editors **
' ******************
Public Sub EditorMap_DrawTileset()
Dim Height As Long, srcRect As D3DRECT, destRect As D3DRECT
Dim Width As Long
Dim Tileset As Long, X As Integer, Y As Integer
Dim sRect As RECT
Dim drect As RECT, scrlX As Long, scrlY As Long
    
    ' find tileset number
    Tileset = frmEditor_Map.scrlTileSet.Value
    
    ' exit out if doesn't exist
    If Tileset < 0 Or Tileset > NumTilesets Then Exit Sub
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    
    scrlX = frmEditor_Map.scrlPictureX.Value * 32
    scrlY = frmEditor_Map.scrlPictureY.Value * 32
    
    Height = Tex_Tileset(Tileset).Height - scrlY
    Width = Tex_Tileset(Tileset).Width - scrlX
    
    sRect.Left = frmEditor_Map.scrlPictureX.Value * 32
    sRect.Top = frmEditor_Map.scrlPictureY.Value * 32
    sRect.Right = sRect.Left + Width
    sRect.bottom = sRect.Top + Height
    
    drect.Top = 0
    drect.bottom = Height
    drect.Left = 0
    drect.Right = Width
    
    For X = 0 To Width / 32
        For Y = 0 To Height / 32
            RenderTexture Tex_Misc(4), X * 32, Y * 32, 0, 0, 32, 32, 32, 32
        Next
    Next
    
    RenderTextureByRects Tex_Tileset(Tileset), sRect, drect
    

    For X = 0 To Width / 32 + (64)
        For Y = 0 To Height / 32 + 64
            If EditorTile(Tileset, X, Y).Type = TILE_TYPE_BLOCKED Then
                RenderText Font_Default, "B", ((X - sRect.Left / 32) * 32) + 10, ((Y - sRect.Top / 32) * 32) + 10, Red
            End If
            '
        Next
    Next
    
    If frmEditor_Map.chkDirBlock.Value Then
        For X = 0 To Tex_Tileset(Tileset).Width / 32
            For Y = 0 To Tex_Tileset(Tileset).Height / 32
                DrawDirTileset Tileset, (X - sRect.Left / 32) * 32, (Y - sRect.Top / 32) * 32, EditorTile(Tileset, X, Y).DirBlock
            Next
        Next
    End If
    
    ' change selected shape for autotiles
    If frmEditor_Map.scrlAutotile.Value > 0 Then
        Select Case frmEditor_Map.scrlAutotile.Value
            Case 1 ' autotile
                EditorTileWidth = 2
                EditorTileHeight = 3
            Case 2 ' fake autotile
                EditorTileWidth = 2
                EditorTileHeight = 1
            Case 3 ' animated
                EditorTileWidth = 6
                EditorTileHeight = 3
            Case 4 ' cliff
                EditorTileWidth = 2
                EditorTileHeight = 2
            Case 5 ' waterfall
                EditorTileWidth = 2
                EditorTileHeight = 3
        End Select
    End If
    
    With destRect
        .X1 = (EditorTileX * 32) - sRect.Left
        .X2 = (EditorTileWidth * 32) + .X1
        .Y1 = (EditorTileY * 32) - sRect.Top
        .Y2 = (EditorTileHeight * 32) + .Y1
    End With
    
    DrawSelectionBox destRect

    With srcRect
        .X1 = 0
        .Y2 = Height
    End With
                    
    With destRect
        .X1 = 0
        .X2 = frmEditor_Map.picBack.ScaleWidth
        .Y1 = 0
        .Y2 = frmEditor_Map.picBack.ScaleHeight
    End With
    
    'Now render the selection tiles and we are done!
        
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Map.picBack.hWnd, ByVal (0)
    
End Sub

Sub DrawSelectionBox(drect As D3DRECT)
Dim Width As Long, Height As Long, X As Long, Y As Long
    Width = drect.X2 - drect.X1
    Height = drect.Y2 - drect.Y1
    X = drect.X1
    Y = drect.Y1
    If Width > 6 And Height > 6 Then
        'Draw Box 32 by 32 at graphicselx and graphicsely
        RenderTexture Tex_Misc(5), X, Y, 1, 1, 2, 2, 2, 2, -1 'top left corner
        RenderTexture Tex_Misc(5), X + 2, Y, 3, 1, Width - 4, 2, 32 - 6, 2, -1 'top line
        RenderTexture Tex_Misc(5), X + 2 + (Width - 4), Y, 29, 1, 2, 2, 2, 2, -1 'top right corner
        RenderTexture Tex_Misc(5), X, Y + 2, 1, 3, 2, Height - 4, 2, 32 - 6, -1 'Left Line
        RenderTexture Tex_Misc(5), X + 2 + (Width - 4), Y + 2, 32 - 3, 3, 2, Height - 4, 2, 32 - 6, -1 'right line
        RenderTexture Tex_Misc(5), X, Y + 2 + (Height - 4), 1, 32 - 3, 2, 2, 2, 2, -1 'bottom left corner
        RenderTexture Tex_Misc(5), X + 2 + (Width - 4), Y + 2 + (Height - 4), 32 - 3, 32 - 3, 2, 2, 2, 2, -1 'bottom right corner
        RenderTexture Tex_Misc(5), X + 2, Y + 2 + (Height - 4), 3, 32 - 3, Width - 4, 2, 32 - 6, 2, -1 'bottom line
    End If
End Sub

Public Sub DrawGDI()

'Cycle Through in-game stuff before cycling through editors
 
If frmEditor_Map.Visible Then
    EditorMap_DrawTileset
End If

End Sub

Public Sub RenderTextureByRects(TextureRec As DX8TextureRec, sRect As RECT, drect As RECT, Optional Colour As Long = -1)

    RenderTexture TextureRec, drect.Left, drect.Top, sRect.Left, sRect.Top, drect.Right - drect.Left, drect.bottom - drect.Top, sRect.Right - sRect.Left, sRect.bottom - sRect.Top, Colour

End Sub

Sub DrawDirTileset(ByVal Tileset As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Block As Byte)
Dim rec As RECT
Dim I As Byte

With rec
    .Top = 24
    .Left = 0
    .Right = .Left + 32
    .bottom = .Top + 32
End With

'grille
RenderTexture Tex_Misc(3), X, Y, rec.Left, rec.Top, rec.Right - rec.Left, rec.bottom - rec.Top, rec.Right - rec.Left, rec.bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255)

For I = 0 To 7
     If I < 5 And I > 0 Then
        rec.Left = (I - 1) * 8
    Else
        rec.Left = 3 * 8
    End If
    rec.Right = rec.Left + 8
    ' find out whether render blocked or not
    If Not isDirBlocked(Block, CByte(I)) Then
        rec.Top = 8
    Else
        rec.Top = 16
    End If
    rec.bottom = rec.Top + 8
    'render!
    RenderTexture Tex_Misc(3), X + DirArrowX(I), Y + DirArrowY(I), rec.Left, rec.Top, rec.Right - rec.Left, rec.bottom - rec.Top, rec.Right - rec.Left, rec.bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255)
Next
End Sub

Public Sub DrawTileApercu()

If frmEditor_Map.chkApercu.Value = 0 Then Exit Sub

Dim X As Integer, Y As Integer

For X = 1 To EditorTileWidth
    For Y = 1 To EditorTileHeight
        RenderTexture Tex_Tileset(frmEditor_Map.scrlTileSet.Value), ConvertMapX(CurX * 32), ConvertMapY(CurY * 32), EditorTileX * 32, EditorTileY * 32, EditorTileWidth * 32, EditorTileHeight * 32, EditorTileWidth * 32, EditorTileHeight * 32, -1
    Next
Next

End Sub
