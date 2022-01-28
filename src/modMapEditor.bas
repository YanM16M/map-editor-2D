Attribute VB_Name = "modMapEditor"
Option Explicit
' ////////////////
' // Map Editor //
' ////////////////
Public Sub MapEditorInit()
Dim I As Long
Dim smusic() As String
    
    ' set the width
    frmEditor_Map.Width = 8500
    
    ' we're in the map editor
    InMapEditor = True
    
    ' show the form
    frmEditor_Map.Visible = True
    
    ' set the scrolly bars
    frmEditor_Map.scrlTileSet.Max = NumTilesets
    frmEditor_Map.fraTileSet.Caption = "Tileset: " & 1
    frmEditor_Map.scrlTileSet.Value = 1
    
    ' set the scrollbars
    frmEditor_Map.scrlPictureY.Max = (Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Height \ 32) - (frmEditor_Map.picBack.Height \ 32)
    frmEditor_Map.scrlPictureX.Max = (Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Width \ 32) - (frmEditor_Map.picBack.Width \ 32)
    MapEditorTileScroll

End Sub

Public Sub MapEditorProperties()
Dim X As Long
Dim Y As Long
Dim I As Long
    

    With frmEditor_MapProperties
        .txtName.text = Trim$(Map.name)
        

        
        ' rest of it
        .txtUp.text = CStr(Map.Up)
        .txtDown.text = CStr(Map.Down)
        .txtLeft.text = CStr(Map.Left)
        .txtRight.text = CStr(Map.Right)
        .cmbMoral.ListIndex = Map.Moral
        .txtBootMap.text = CStr(Map.BootMap)
        .txtBootX.text = CStr(Map.BootX)
        .txtBootY.text = CStr(Map.BootY)
        

        
        .ScrlFog.Value = Map.Fog
        .ScrlFogSpeed.Value = Map.FogSpeed
        .scrlFogOpacity.Value = Map.FogOpacity
        
        .ScrlR.Value = Map.Red
        .ScrlG.Value = Map.Green
        .ScrlB.Value = Map.Blue
        .scrlA.Value = Map.Alpha
        
         If Map.CanTp = True Then
        .chkTP.Value = 1
        Else
        .chkTP.Value = 0
        End If

        ' show the current map
        .lblMap.Caption = "Current map: " & CurrentMap
        .txtMaxX.text = Map.MaxX
        .txtMaxY.text = Map.MaxY
        .cmbDayNight.ListIndex = Map.DayNight
        '.cmbPanorama.ListIndex = Map.Panorama
        .chkTP.Value = Map.CanTp
        .chkDead = Map.IsDead
        .scrlGravity = Map.Gravity
    End With

End Sub

Public Sub MapEditorSetTile(ByVal X As Long, ByVal Y As Long, ByVal CurLayer As Long, Optional ByVal multitile As Boolean = False, Optional ByVal theAutotile As Byte = 0)
Dim X2 As Long, Y2 As Long, I As Integer, Verif As Boolean
    
    Map_Changed = True
    
    If theAutotile > 0 Then
        With Map.Tile(X, Y)
            ' set layer
            .layer(CurLayer).X = EditorTileX
            .layer(CurLayer).Y = EditorTileY
            .layer(CurLayer).Tileset = frmEditor_Map.scrlTileSet.Value
            .Autotile(CurLayer) = theAutotile
            CacheRenderState X, Y, CurLayer
        End With
        ' do a re-init so we can see our changes
        If theAutotile <> AUTOTILE_FAKE Then initAutotiles
        Exit Sub
    End If



    If Not multitile Then ' single
        With Map.Tile(X, Y)
            For I = 1 To MAX_BYTE
                If Not Verif And LastClick(I) = False Then
                    LastX(I) = 0
                    LastY(I) = 0
                    LastTileX(I) = 0
                    LastTileY(I) = 0
                    LastTileset(I) = 0
                    LastCurlayer(I) = 0
                    LastClick(I) = False
                    If I = MAX_BYTE Then
                        I = 1
                        Verif = True
                    End If
                    If Not Verif Then GoTo nextLoop
                End If
                
                If LastX(I) = X And LastY(I) = Y Then Exit For
                If LastClick(I) = False Then
                    LastX(I) = X
                    LastY(I) = Y
                    LastTileX(I) = .layer(CurLayer).X
                    LastTileY(I) = .layer(CurLayer).Y
                    LastTileset(I) = .layer(CurLayer).Tileset
                    LastCurlayer(I) = CurLayer
                    LastClick(I) = True
                    Exit For
                End If
nextLoop:
            Next
            ' set layer
            .layer(CurLayer).X = EditorTileX
            .layer(CurLayer).Y = EditorTileY
            .layer(CurLayer).Tileset = frmEditor_Map.scrlTileSet.Value
            

            .Autotile(CurLayer) = 0
                
            CacheRenderState X, Y, CurLayer
        End With
    Else ' multitile
        Y2 = 0 ' starting tile for y axis
        For Y = CurY To CurY + EditorTileHeight - 1
            X2 = 0 ' re-set x count every y loop
            For X = CurX To CurX + EditorTileWidth - 1
                If X >= 0 And X <= Map.MaxX Then
                    If Y >= 0 And Y <= Map.MaxY Then
                        With Map.Tile(X, Y)
                            For I = 1 To MAX_BYTE
                                If Not Verif And LastClick(I) = False Then
                                    LastX(I) = 0
                                    LastY(I) = 0
                                    LastTileX(I) = 0
                                    LastTileY(I) = 0
                                    LastTileset(I) = 0
                                    LastCurlayer(I) = 0
                                    LastClick(I) = False
                                    If I = MAX_BYTE Then
                                        I = 1
                                        Verif = True
                                    End If
                                    If Not Verif Then GoTo NextLoo
                                End If
                                If LastX(I) = X And LastY(I) = Y Then Exit For
                                If LastClick(I) = False Then
                                    LastX(I) = X
                                    LastY(I) = Y
                                    LastTileX(I) = .layer(CurLayer).X
                                    LastTileY(I) = .layer(CurLayer).Y
                                    LastTileset(I) = .layer(CurLayer).Tileset
                                    LastCurlayer(I) = CurLayer
                                    LastClick(I) = True
                                    Exit For
                                End If
NextLoo:
                            Next
                            
                            .layer(CurLayer).X = EditorTileX + X2
                            .layer(CurLayer).Y = EditorTileY + Y2
                            .layer(CurLayer).Tileset = frmEditor_Map.scrlTileSet.Value
                            

                            .Autotile(CurLayer) = 0
                            
                            CacheRenderState X, Y, CurLayer
                        End With
                    End If
                End If
                X2 = X2 + 1
            Next
            Y2 = Y2 + 1
        Next
    End If

End Sub

Public Sub MapEditorMouseDown(ByVal button As Integer, ByVal X As Long, ByVal Y As Long, Optional ByVal movedMouse As Boolean = True)
Dim I As Long
Dim CurLayer As Long
Dim tmpDir As Byte, Verif As Boolean
    
    ' find which layer we're on
    For I = 1 To MapLayer.Layer_Count - 1
        If frmEditor_Map.optLayer(I).Value Then
            CurLayer = I
            Exit For
        End If
    Next

    If Not isInBounds Then Exit Sub
    If button = vbLeftButton Then
        If frmEditor_Map.optLayers.Value Then
            ' no autotiling
            If EditorTileWidth = 1 And EditorTileHeight = 1 Then 'single tile
                MapEditorSetTile CurX, CurY, CurLayer, , frmEditor_Map.scrlAutotile.Value
            Else ' multi tile!
                If frmEditor_Map.scrlAutotile.Value = 0 Then
                    MapEditorSetTile CurX, CurY, CurLayer, True
                Else
                    MapEditorSetTile CurX, CurY, CurLayer, , frmEditor_Map.scrlAutotile.Value
                End If
            End If
        ElseIf frmEditor_Map.optAttribs.Value Then
            With Map.Tile(CurX, CurY)
                ' blocked tile
                If frmEditor_Map.optBlocked.Value Then .Type = TILE_TYPE_BLOCKED
                ' warp tile
                If frmEditor_Map.optWarp.Value Then
                    .Type = TILE_TYPE_WARP
                    .data1 = EditorWarpMap
                    .Data2 = EditorWarpX
                    .Data3 = EditorWarpY
                    .Data4 = EditorWarpTp
                End If
                ' npc avoid
                If frmEditor_Map.optNpcAvoid.Value Then
                    .Type = TILE_TYPE_NPCAVOID
                    .data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' key
                If frmEditor_Map.optKey.Value Then
                    .Type = TILE_TYPE_KEY
                    .data1 = KeyEditorNum
                    .Data2 = KeyEditorTake
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' key open
                If frmEditor_Map.optKeyOpen.Value Then
                    .Type = TILE_TYPE_KEYOPEN
                    .data1 = KeyOpenEditorX
                    .Data2 = KeyOpenEditorY
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' resource
                If frmEditor_Map.optResource.Value Then
                    .Type = TILE_TYPE_RESOURCE
                    .data1 = ResourceEditorNum
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' bank
                If frmEditor_Map.optBank.Value Then
                    .Type = TILE_TYPE_BANK
                    .data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' heal
                If frmEditor_Map.optHeal.Value Then
                    .Type = TILE_TYPE_HEAL
                    .data1 = MapEditorHealType
                    .Data2 = MapEditorHealAmount
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' trap
                If frmEditor_Map.optTrap.Value Then
                    .Type = TILE_TYPE_TRAP
                    .data1 = MapEditorHealAmount
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' slide
                If frmEditor_Map.optSlide.Value Then
                    .Type = TILE_TYPE_SLIDE
                    .data1 = MapEditorSlideDir
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' sound
                If frmEditor_Map.optSound.Value Then
                    .Type = TILE_TYPE_SOUND
                    .data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = MapEditorSound
                End If
                ' Light
                If frmEditor_Map.optLight.Value Then
                    .Type = TILE_TYPE_LIGHT
                    .data1 = MapEditorLightA
                    .Data2 = MapEditorLightR
                    .Data3 = MapEditorLightG
                    .Data4 = MapEditorLightB
                End If
                ' Craft
                'If frmEditor_Map.optCraft.value Then
                    '.Type = TILE_TYPE_CRAFT
                    '.data1 = 0
                    '.Data2 = 0
                    '.Data3 = 0
                    '.Data4 = vbNullString
                'End If
                'Nofight
                If frmEditor_Map.optNoFight.Value Then
                    .Type = TILE_TYPE_NOFIGHT
                    .data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = vbNullString
                End If
                'Vol
                If frmEditor_Map.optVol.Value Then
                    .Type = TILE_TYPE_VOL
                End If
                'vol 2 pour se poser
                If frmEditor_Map.optVol2.Value Then
                    .Type = TILE_TYPE_VOL2
                End If
                
                'Dragon Ball
                If frmEditor_Map.optDB.Value Then
                    .Type = TILE_TYPE_DB
                    .data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = vbNullString
                End If
                
                'Label
                If frmEditor_Map.optLabel.Value Then
                    .Type = TILE_TYPE_LABEL
                    .data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = LabelMap
                End If
                
                'obscur
                If frmEditor_Map.optObscur.Value Then
                    .Type = TILE_TYPE_OBSCUR
                    .data1 = EditorObscurAlpha
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = vbNullString
                End If
                
                'lampe
                If frmEditor_Map.optLampe.Value Then
                    .Type = TILE_TYPE_LAMPE
                    .data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = vbNullString
                End If
                
            End With
        ElseIf frmEditor_Map.optBlock.Value Then
            If movedMouse Then Exit Sub
            ' find what tile it is
            X = X - ((X \ 32) * 32)
            Y = Y - ((Y \ 32) * 32)
            ' see if it hits an arrow
            For I = 0 To 7
                If X >= DirArrowX(I) And X <= DirArrowX(I) + 8 Then
                    If Y >= DirArrowY(I) And Y <= DirArrowY(I) + 8 Then
                        ' flip the value.
                        setDirBlock Map.Tile(CurX, CurY).DirBlock, CByte(I), Not isDirBlocked(Map.Tile(CurX, CurY).DirBlock, CByte(I))
                        Exit Sub
                    End If
                End If
            Next
        End If
    End If

    If button = vbRightButton Then
        If frmEditor_Map.optLayers.Value Then
            With Map.Tile(CurX, CurY)
                For I = 1 To MAX_BYTE
                    If Not Verif And LastClick(I) = False Then
                    LastX(I) = 0
                    LastY(I) = 0
                    LastTileX(I) = 0
                    LastTileY(I) = 0
                    LastTileset(I) = 0
                    LastCurlayer(I) = 0
                    LastClick(I) = False
                    If I = MAX_BYTE Then
                        I = 1
                        Verif = True
                    End If
                    If Not Verif Then GoTo nextLoop
                End If
                
                    If LastX(I) = CurX And LastY(I) = CurY Then Exit For
                    If LastClick(I) = False Then
                        LastX(I) = CurX
                        LastY(I) = CurY
                        LastTileX(I) = .layer(CurLayer).X
                        LastTileY(I) = .layer(CurLayer).Y
                        LastTileset(I) = .layer(CurLayer).Tileset
                        LastCurlayer(I) = CurLayer
                        LastClick(I) = True
                        Exit For
                    End If
nextLoop:
                Next
                
                ' clear layer
                .layer(CurLayer).X = 0
                .layer(CurLayer).Y = 0
                .layer(CurLayer).Tileset = 0
                If .Autotile(CurLayer) > 0 Then
                    .Autotile(CurLayer) = 0
                    ' do a re-init so we can see our changes
                    initAutotiles
                End If
                CacheRenderState X, Y, CurLayer
            End With
        ElseIf frmEditor_Map.optAttribs.Value Then
            With Map.Tile(CurX, CurY)
                ' clear attribute
                .Type = 0
                .data1 = 0
                .Data2 = 0
                .Data3 = 0
            End With

        End If
    End If

    'CacheResources
    
End Sub

Public Sub MapEditorChooseTile(button As Integer, X As Single, Y As Single)

    If button = vbLeftButton Then
        EditorTileWidth = 1
        EditorTileHeight = 1
        
        EditorTileX = X \ 32
        EditorTileY = Y \ 32
    End If
    
End Sub

Public Sub MapEditorDrag(button As Integer, X As Single, Y As Single)

    If button = vbLeftButton Then
        ' convert the pixel number to tile number
        X = (X \ 32) + 1
        Y = (Y \ 32) + 1
        ' check it's not out of bounds
        If X < 0 Then X = 0
        If X > Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Width / 32 Then X = Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Width / 32
        If Y < 0 Then Y = 0
        If Y > Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Height / 32 Then Y = Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Height / 32
        ' find out what to set the width + height of map editor to
        If X > EditorTileX Then ' drag right
            EditorTileWidth = X - EditorTileX
        Else ' drag left
            ' TO DO
        End If
        If Y > EditorTileY Then ' drag down
            EditorTileHeight = Y - EditorTileY
        Else ' drag up
            ' TO DO
        End If
    End If

End Sub

Public Sub MapEditorTileScroll()

    ' horizontal scrolling
    If Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Width < frmEditor_Map.picBack.Width Then
        frmEditor_Map.scrlPictureX.Enabled = False
    Else
        frmEditor_Map.scrlPictureX.Enabled = True
    End If
    
    ' vertical scrolling
    If Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Height < frmEditor_Map.picBack.Height Then
        frmEditor_Map.scrlPictureY.Enabled = False
    Else
        frmEditor_Map.scrlPictureY.Enabled = True
    End If

End Sub

Public Sub MapEditorSend(Optional ByVal Cls As Boolean = True)

    Call SaveMap(CurrentMap)
    
End Sub


Public Sub MapEditorClearLayer()
Dim I As Long
Dim X As Long
Dim Y As Long
Dim CurLayer As Long

    ' find which layer we're on
    For I = 1 To MapLayer.Layer_Count - 1
        If frmEditor_Map.optLayer(I).Value Then
            CurLayer = I
            Exit For
        End If
    Next
    
    If CurLayer = 0 Then Exit Sub
    
    If frmEditor_Map.optBlock.Value Then
        If MsgBox("Es-tu sûr de vouloir supprimer les bloques directions?", vbYesNo, "Editeur de Map") = vbYes Then
            For X = 0 To Map.MaxX
                For Y = 0 To Map.MaxY
                    Map.Tile(X, Y).DirBlock = 0
                Next
            Next
        End If
    Else
        ' ask to clear layer
        If MsgBox("Es-tu sûr de vouloir supprimer cette couche ?", vbYesNo, "Editeur de Map") = vbYes Then
            For X = 0 To Map.MaxX
                For Y = 0 To Map.MaxY
                    Map.Tile(X, Y).layer(CurLayer).X = 0
                    Map.Tile(X, Y).layer(CurLayer).Y = 0
                    Map.Tile(X, Y).layer(CurLayer).Tileset = 0
                    CacheRenderState X, Y, CurLayer
                Next
            Next
            
            initAutotiles
        End If
    End If

End Sub

Public Sub MapEditorFillLayer()
Dim I As Long
Dim X As Long
Dim Y As Long
Dim CurLayer As Long

    ' find which layer we're on
    For I = 1 To MapLayer.Layer_Count - 1
        If frmEditor_Map.optLayer(I).Value Then
            CurLayer = I
            Exit For
        End If
    Next

    If MsgBox("Es-tu sur de vouloir remplir la map?", vbYesNo, "Editeur de Map") = vbYes Then
        For X = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                Map.Tile(X, Y).layer(CurLayer).X = EditorTileX
                Map.Tile(X, Y).layer(CurLayer).Y = EditorTileY
                Map.Tile(X, Y).layer(CurLayer).Tileset = frmEditor_Map.scrlTileSet.Value
                Map.Tile(X, Y).Autotile(CurLayer) = frmEditor_Map.scrlAutotile.Value
                CacheRenderState X, Y, CurLayer
            Next
        Next
        
        ' now cache the positions
        initAutotiles
    End If
    
End Sub

Public Sub MapEditorClearAttribs()
Dim X As Long
Dim Y As Long

    If MsgBox("Are you sure you wish to clear the attributes on this map?", vbYesNo, "Editeur de Map") = vbYes Then

        For X = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                Map.Tile(X, Y).Type = 0
            Next
        Next

    End If

End Sub


Public Sub MapEditorPlaceRandomTile(ByVal X As Long, Y As Long)
Dim I As Long
Dim CurLayer As Long

    ' find which layer we're on
    For I = 1 To MapLayer.Layer_Count - 1
        If frmEditor_Map.optLayer(I).Value Then
            CurLayer = I
            Exit For
        End If
    Next

    If Not isInBounds Then Exit Sub

    If frmEditor_Map.optLayers.Value Then
        If EditorTileWidth = 1 And EditorTileHeight = 1 Then 'single tile
            MapEditorSetTile X, Y, CurLayer, , frmEditor_Map.scrlAutotile.Value
        Else ' multi tile!
            If frmEditor_Map.scrlAutotile.Value = 0 Then
                MapEditorSetTile X, Y, CurLayer, True
            Else
                MapEditorSetTile X, Y, CurLayer, , frmEditor_Map.scrlAutotile.Value
            End If
        End If
    End If

    'CacheResources
    
End Sub

Public Sub ClearAttributeDialogue()
   
    frmEditor_Map.fraNpcSpawn.Visible = False
    frmEditor_Map.fraResource.Visible = False
    frmEditor_Map.fraMapItem.Visible = False
    frmEditor_Map.fraMapKey.Visible = False
    frmEditor_Map.fraKeyOpen.Visible = False
    frmEditor_Map.fraMapWarp.Visible = False
    frmEditor_Map.fraShop.Visible = False
    frmEditor_Map.fraSoundEffect.Visible = False
    frmEditor_Map.fraLight.Visible = False
    frmEditor_Map.frmLabel.Visible = False
    

End Sub

Public Function isInBounds()
 
    If (CurX >= 0) Then
        If (CurX <= Map.MaxX) Then
            If (CurY >= 0) Then
                If (CurY <= Map.MaxY) Then
                    isInBounds = True
                End If
            End If
        End If
    End If

End Function

Public Sub MapReportInit()
Dim I As Long

With frmMapReport
    .lstMap.Clear
    For I = 1 To MAX_MAPS
        .lstMap.AddItem I & " - " & "Map#" & I
    Next
    
    
    .Visible = True
End With

End Sub
