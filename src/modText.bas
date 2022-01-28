Attribute VB_Name = "modText"
Option Explicit

' Stuffs
Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type CharVA
    Vertex(0 To 3) As TLVERTEX
End Type

Public Type VFH
    BitmapWidth As Long
    BitmapHeight As Long
    CellWidth As Long
    CellHeight As Long
    BaseCharOffset As Byte
    CharWidth(0 To 255) As Byte
    CharVA(0 To 255) As CharVA
End Type

Public Type CustomFont
    HeaderInfo As VFH
    Texture As DX8TextureRec
    RowPitch As Integer
    RowFactor As Single
    ColFactor As Single
    CharHeight As Byte
End Type


Public Font_Default As CustomFont
Public Font_Georgia As CustomFont


' Chat Buffer
Public ChatVA() As TLVERTEX
Public ChatVAS() As TLVERTEX

Public Const ChatTextBufferSize As Integer = 200
Public ChatBufferChunk As Single
'Text buffer

Public Type ChatTextBuffer
    text As String
    Color As Long
End Type

'Chat vertex buffer information
Public Const ColourChar As String * 1 = "½"
Public ChatArrayUbound As Long
Public ChatVB As Direct3DVertexBuffer8
Public ChatVBS As Direct3DVertexBuffer8
Public ChatTextBuffer(1 To ChatTextBufferSize) As ChatTextBuffer

Public Const FVF_SIZE As Long = 28

Public Sub RenderText(ByRef UseFont As CustomFont, ByVal text As String, ByVal X As Long, ByVal Y As Long, ByVal Color As Long, Optional ByVal Alpha As Long = 0, Optional Shadow As Boolean = True, Optional ByVal Dubble As Boolean = True)
Dim TempVA(0 To 3)  As TLVERTEX
Dim TempVAS(0 To 3) As TLVERTEX
Dim TempStr() As String
Dim Count As Integer
Dim Ascii() As Byte
Dim Row As Integer
Dim u As Single
Dim v As Single
Dim i As Long
Dim j As Long
Dim KeyPhrase As Byte
Dim TempColor As Long
Dim ResetColor As Byte
Dim srcRect As RECT
Dim v2 As D3DVECTOR2
Dim v3 As D3DVECTOR2
Dim yOffset As Single
Dim counter As Byte

retry:
    
    If counter = 1 Then
        Y = Y + 1
    End If

    ' set the color
    Alpha = 255 - Alpha
    Color = dx8Colour(Color, Alpha)
    
    'Check for valid text to render
    If LenB(text) = 0 Then Exit Sub
    
    'Get the text into arrays (split by vbCrLf)
    TempStr = Split(text, vbCrLf)
    
    'Set the temp color (or else the first character has no color)
    TempColor = Color
    
    'Set the texture
    Direct3D_Device.SetTexture 0, gTexture(UseFont.Texture.Texture).Texture
    'CurrentTexture = -1
    
    'Loop through each line if there are line breaks (vbCrLf)
    For i = 0 To UBound(TempStr)
        If Len(TempStr(i)) > 0 Then
            yOffset = i * UseFont.CharHeight
            Count = 0
            'Convert the characters to the ascii value
            Ascii() = StrConv(TempStr(i), vbFromUnicode)
            
            'Loop through the characters
            For j = 1 To Len(TempStr(i))
                'Copy from the cached vertex array to the temp vertex array
                Call CopyMemory(TempVA(0), UseFont.HeaderInfo.CharVA(Ascii(j - 1)).Vertex(0), FVF_SIZE * 4)
                
                'Set up the verticies
                TempVA(0).X = X + Count
                TempVA(0).Y = Y + yOffset
                TempVA(1).X = TempVA(1).X + X + Count
                TempVA(1).Y = TempVA(0).Y
                TempVA(2).X = TempVA(0).X
                TempVA(2).Y = TempVA(2).Y + TempVA(0).Y
                TempVA(3).X = TempVA(1).X
                TempVA(3).Y = TempVA(2).Y
                
                'Set the colors
                TempVA(0).Color = TempColor
                TempVA(1).Color = TempColor
                TempVA(2).Color = TempColor
                TempVA(3).Color = TempColor
                
                'Draw the verticies
                Call Direct3D_Device.DrawPrimitiveUP(D3DPT_TRIANGLESTRIP, 2, TempVA(0), Len(TempVA(0)))
                
                'Shift over the the position to render the next character
                Count = Count + UseFont.HeaderInfo.CharWidth(Ascii(j - 1))
                
                'Check to reset the color
                If ResetColor Then
                    ResetColor = 0
                    TempColor = Color
                End If
            Next j
        End If
    Next i
    
    If Dubble Then
        If counter = 0 Then
            counter = 1
            GoTo retry
        End If
    End If
    
End Sub

Sub EngineInitFontTextures()
    ' FONT DEFAULT
    NumTextures = NumTextures + 1
    ReDim Preserve gTexture(NumTextures)
    Font_Default.Texture.Texture = NumTextures
    Font_Default.Texture.filepath = App.path & FONT_PATH & "texdefault.png"
    LoadTexture Font_Default.Texture
    
    ' Georgia
    NumTextures = NumTextures + 1
    ReDim Preserve gTexture(NumTextures)
    Font_Georgia.Texture.Texture = NumTextures
    Font_Georgia.Texture.filepath = App.path & FONT_PATH & "georgia.png"
    LoadTexture Font_Georgia.Texture
    
End Sub

Sub UnloadFontTextures()
    UnloadFont Font_Default
    UnloadFont Font_Georgia
End Sub
Sub UnloadFont(Font As CustomFont)
    Font.Texture.Texture = 0
End Sub


Sub LoadFontHeader(ByRef theFont As CustomFont, ByVal FileName As String)
Dim FileNum As Byte
Dim LoopChar As Long
Dim Row As Single
Dim u As Single
Dim v As Single


    'Load the header information
    FileNum = FreeFile
    Open App.path & FONT_PATH & FileName For Binary As #FileNum
        Get #FileNum, , theFont.HeaderInfo
    Close #FileNum
    
    'Calculate some common values
    theFont.CharHeight = theFont.HeaderInfo.CellHeight - 4
    theFont.RowPitch = theFont.HeaderInfo.BitmapWidth \ theFont.HeaderInfo.CellWidth
    theFont.ColFactor = theFont.HeaderInfo.CellWidth / theFont.HeaderInfo.BitmapWidth
    theFont.RowFactor = theFont.HeaderInfo.CellHeight / theFont.HeaderInfo.BitmapHeight
    
    'Cache the verticies used to draw the character (only requires setting the color and adding to the X/Y values)
    For LoopChar = 0 To 255
        'tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
        Row = (LoopChar - theFont.HeaderInfo.BaseCharOffset) \ theFont.RowPitch
        u = ((LoopChar - theFont.HeaderInfo.BaseCharOffset) - (Row * theFont.RowPitch)) * theFont.ColFactor
        v = Row * theFont.RowFactor
        
        'Set the verticies
        With theFont.HeaderInfo.CharVA(LoopChar)
            .Vertex(0).Color = D3DColorARGB(255, 0, 0, 0)   'Black is the most common color
            .Vertex(0).RHW = 1
            .Vertex(0).TU = u
            .Vertex(0).TV = v
            .Vertex(0).X = 0
            .Vertex(0).Y = 0
            .Vertex(0).Z = 0
            .Vertex(1).Color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(1).RHW = 1
            .Vertex(1).TU = u + theFont.ColFactor
            .Vertex(1).TV = v
            .Vertex(1).X = theFont.HeaderInfo.CellWidth
            .Vertex(1).Y = 0
            .Vertex(1).Z = 0
            .Vertex(2).Color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(2).RHW = 1
            .Vertex(2).TU = u
            .Vertex(2).TV = v + theFont.RowFactor
            .Vertex(2).X = 0
            .Vertex(2).Y = theFont.HeaderInfo.CellHeight
            .Vertex(2).Z = 0
            .Vertex(3).Color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(3).RHW = 1
            .Vertex(3).TU = u + theFont.ColFactor
            .Vertex(3).TV = v + theFont.RowFactor
            .Vertex(3).X = theFont.HeaderInfo.CellWidth
            .Vertex(3).Y = theFont.HeaderInfo.CellHeight
            .Vertex(3).Z = 0
        End With
    Next LoopChar
End Sub

Sub EngineInitFontSettings()
    LoadFontHeader Font_Default, "texdefault.dat"
    LoadFontHeader Font_Georgia, "georgia.dat"
End Sub
Public Function dx8Colour(ByVal colourNum As Long, ByVal Alpha As Long) As Long
    Select Case colourNum
        Case 0 ' Black
            dx8Colour = D3DColorARGB(Alpha, 0, 0, 0)
        Case 1 ' Blue
            dx8Colour = D3DColorARGB(Alpha, 16, 104, 237)
        Case 2 ' Green
            dx8Colour = D3DColorARGB(Alpha, 119, 188, 84)
        Case 3 ' Cyan
            dx8Colour = D3DColorARGB(Alpha, 16, 224, 237)
        Case 4 ' Red
            dx8Colour = D3DColorARGB(Alpha, 201, 0, 0)
        Case 5 ' Magenta
            dx8Colour = D3DColorARGB(Alpha, 255, 0, 255)
        Case 6 ' Brown
            dx8Colour = D3DColorARGB(Alpha, 175, 149, 92)
        Case 7 ' Grey
            dx8Colour = D3DColorARGB(Alpha, 192, 192, 192)
        Case 8 ' DarkGrey
            dx8Colour = D3DColorARGB(Alpha, 128, 128, 128)
        Case 9 ' BrightBlue
            dx8Colour = D3DColorARGB(Alpha, 126, 182, 240)
        Case 10 ' BrightGreen
            dx8Colour = D3DColorARGB(Alpha, 126, 240, 137)
        Case 11 ' BrightCyan
            dx8Colour = D3DColorARGB(Alpha, 157, 242, 242)
        Case 12 ' BrightRed
            dx8Colour = D3DColorARGB(Alpha, 255, 0, 0)
        Case 13 ' Pink
            dx8Colour = D3DColorARGB(Alpha, 255, 118, 221)
        Case 14 ' Yellow
            dx8Colour = D3DColorARGB(Alpha, 255, 255, 0)
        Case 15 ' White
            dx8Colour = D3DColorARGB(Alpha, 255, 255, 255)
        Case 16 ' dark brown
            dx8Colour = D3DColorARGB(Alpha, 98, 84, 52)
        Case 17 'Orange
            dx8Colour = D3DColorARGB(Alpha, 255, 96, 0)
        Case 18 'Pinky
            dx8Colour = D3DColorARGB(Alpha, 160, 98, 255)
    End Select
    
End Function

Public Function EngineGetTextWidth(ByRef UseFont As CustomFont, ByVal text As String) As Integer
Dim LoopI As Integer

    'Make sure we have text
    If LenB(text) = 0 Then Exit Function
    
    'Loop through the text
    For LoopI = 1 To Len(text)
        EngineGetTextWidth = EngineGetTextWidth + UseFont.HeaderInfo.CharWidth(Asc(Mid$(text, LoopI, 1)))
    Next LoopI

End Function

Public Function DrawMapAttributes()
Dim X As Integer, Y As Integer, tx As Integer, ty As Integer

If frmEditor_Map.optAttribs.Value Then
    For X = TileView.Left To TileView.Right
        For Y = TileView.Top To TileView.bottom
            If IsValidMapPoint(X, Y) Then
                With Map.Tile(X, Y)
                    tx = ((ConvertMapX(X * 32)) - 4) + (32 * 0.5)
                    ty = ((ConvertMapY(Y * 32)) - 7) + (32 * 0.5)
                    Select Case .Type
                        Case TILE_TYPE_BLOCKED
                            RenderText Font_Default, "B", tx, ty, BrightRed, 0
                        Case TILE_TYPE_WARP
                            RenderText Font_Default, "W", tx, ty, BrightBlue, 0
                        Case TILE_TYPE_ITEM
                            RenderText Font_Default, "I", tx, ty, White, 0
                        Case TILE_TYPE_NPCAVOID
                            RenderText Font_Default, "N", tx, ty, White, 0
                        Case TILE_TYPE_KEY
                            RenderText Font_Default, "K", tx, ty, White, 0
                        Case TILE_TYPE_KEYOPEN
                            RenderText Font_Default, "O", tx, ty, White, 0
                        Case TILE_TYPE_RESOURCE
                            RenderText Font_Default, "B", tx, ty, Green, 0
                        Case TILE_TYPE_DOOR
                            RenderText Font_Default, "D", tx, ty, Brown, 0
                        Case TILE_TYPE_NPCSPAWN
                            RenderText Font_Default, "S" & Map.Tile(X, Y).data1, tx, ty, Yellow, 0
                        Case TILE_TYPE_SHOP
                            RenderText Font_Default, "S", tx, ty, BrightBlue, 0
                        Case TILE_TYPE_BANK
                            RenderText Font_Default, "B", tx, ty, Blue, 0
                        Case TILE_TYPE_HEAL
                            RenderText Font_Default, "H", tx, ty, BrightGreen, 0
                        Case TILE_TYPE_TRAP
                            RenderText Font_Default, "T", tx, ty, BrightRed, 0
                        Case TILE_TYPE_SLIDE
                            RenderText Font_Default, "S", tx, ty, BrightCyan, 0
                        Case TILE_TYPE_SOUND
                            RenderText Font_Default, "S", tx, ty, Orange, 0
                        Case TILE_TYPE_LIGHT
                            RenderText Font_Default, "L", tx, ty, Yellow
                        Case TILE_TYPE_CRAFT
                            RenderText Font_Default, "CR", tx, ty, Orange
                        Case TILE_TYPE_NOFIGHT
                            RenderText Font_Default, "NF", tx, ty, BrightBlue
                        Case TILE_TYPE_VOL
                            RenderText Font_Default, "V", tx, ty, BrightBlue
                        Case TILE_TYPE_VOL2
                            RenderText Font_Default, "P", tx, ty, Brown
                        Case TILE_TYPE_DB
                            RenderText Font_Default, "DB", tx, ty, Orange
                        Case TILE_TYPE_LABEL
                            RenderText Font_Default, Map.Tile(X, Y).Data4, tx, ty, Red
                        Case TILE_TYPE_OBSCUR
                            RenderText Font_Default, "OB", tx, ty, Red
                        Case TILE_TYPE_LAMPE
                            RenderText Font_Default, "LP", tx, ty, Yellow
                    End Select
                End With
            End If
        Next
    Next
End If

End Function
