Attribute VB_Name = "modGeneral"
Option Explicit

'Key-dependant
Private m_KeyS As String
'***** RC4 *****
Private m_sBoxRC4(0 To 255) As Integer

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Used for the 64-bit timer
Private GetSystemTimeOffset As Currency
Private Declare Sub GetSystemTime Lib "kernel32.dll" Alias "GetSystemTimeAsFileTime" (ByRef lpSystemTimeAsFileTime As Currency)
Public Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long

' API Declares
'For Clear functions
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
' For Copy functions
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Sub Main()
Dim I As Byte

' Set the high-resolution timer
timeBeginPeriod 1

' This must be called before any timeGetTime calls because it states what the values of timeGetTime will be
InitTimeGetTime

'######################
'#### Folder
'######################

ChkDir App.path & "\data files\", "graphics"
ChkDir App.path & "\data files\", "maps"

ChkDir App.path & "\data files\graphics\", "fonts"
ChkDir App.path & "\data files\graphics\", "fogs"
ChkDir App.path & "\data files\graphics\", "tilesets"
ChkDir App.path & "\data files\graphics\", "misc"

' set values for directional blocking arrows
DirArrowX(0) = 12 ' up
DirArrowY(0) = 0
DirArrowX(1) = 12 ' down
DirArrowY(1) = 23
DirArrowX(2) = 0 ' left
DirArrowY(2) = 12
DirArrowX(3) = 23 ' right
DirArrowY(3) = 12

DirArrowX(4) = 0 ' upleft
DirArrowY(4) = 0
DirArrowX(5) = 23 ' upright
DirArrowY(5) = 0
DirArrowX(6) = 0 ' downleft
DirArrowY(6) = 23
DirArrowX(7) = 23 ' downright
DirArrowY(7) = 23

'######################
'#### Caption
'######################
frmMain.Caption = "Editeur de Map"

'######################
'#### DX8
'######################
ScreenWidth = 800
ScreenHeight = 732
Call InitDX8

'######################
'#### Maps
'######################
CurrentMap = 1
CurrentX = 0
CurrentY = 0
Call CheckMaps
Call LoadMap(1)
Call LoadBloques

'######################
'#### Main
'######################
BloC = True
frmMain.Show
Call MapEditorInit

' Start Main Loop
Call MainLoop
End Sub

Public Sub DestroyGame()
    DestroyDX8
    DestroyForm
    End
End Sub

Public Sub DestroyForm()
    Unload frmEditor_Map
    Unload frmEditor_MapProperties
    Unload frmMapReport
End Sub

Public Function CheckFolder(ByRef Count As Long, ByVal path As String) As DX8TextureRec()
Dim Texture() As DX8TextureRec
Dim I As Integer

I = 1
Count = 1
ReDim Texture(1)

While FileExist(App.path & GFX_PATH & path & I & ".fight", True)
    ReDim Preserve Texture(Count)
    NumTextures = NumTextures + 1
    ReDim Preserve gTexture(NumTextures)
    Texture(Count).filepath = App.path & GFX_PATH & path & I & GFX_EXT
    Texture(Count).Texture = NumTextures
    Call LoadTexture(Texture(Count))
    Count = Count + 1
    I = I + 1
Wend

Count = Count - 1
CheckFolder = Texture

End Function

Public Sub Encryption_RC4_SetKey(New_Value As String)
Dim a As Long
Dim B As Long
Dim Temp As Byte
Dim Key() As Byte
Dim KeyLen As Long

    'Do nothing if the key is buffered
    If (m_KeyS = New_Value) Then Exit Sub

    'Set the new key
    m_KeyS = New_Value

    'Save the password in a byte array
    Key() = StrConv(m_KeyS, vbFromUnicode)
    KeyLen = Len(m_KeyS)

    'Initialize s-boxes
    For a = 0 To 255
        m_sBoxRC4(a) = a
    Next a
    For a = 0 To 255
        B = (B + m_sBoxRC4(a) + Key(a Mod KeyLen)) Mod 256
        Temp = m_sBoxRC4(a)
        m_sBoxRC4(a) = m_sBoxRC4(B)
        m_sBoxRC4(B) = Temp
    Next

End Sub
Public Sub Encryption_RC4_EncryptByte(ByteArray() As Byte, Optional Key As String)
Dim I As Long
Dim j As Long
Dim Temp As Byte
Dim offset As Long
Dim OrigLen As Long
Dim sBox(0 To 255) As Integer

    'Set the new key (optional)
    If (Len(Key) > 0) Then Encryption_RC4_SetKey Key

    'Create a local copy of the sboxes, this
    'is much more elegant than recreating
    'before encrypting/decrypting anything
    Call CopyMemory(sBox(0), m_sBoxRC4(0), 512)

    'Get the size of the source array
    OrigLen = UBound(ByteArray) + 1

    'Encrypt the data
    For offset = 0 To (OrigLen - 1)
        I = (I + 1) Mod 256
        j = (j + sBox(I)) Mod 256
        Temp = sBox(I)
        sBox(I) = sBox(j)
        sBox(j) = Temp
        ByteArray(offset) = ByteArray(offset) Xor (sBox((sBox(I) + sBox(j)) Mod 256))
    Next

End Sub

Public Sub Encryption_RC4_DecryptByte(ByteArray() As Byte, Optional Key As String)

    Call Encryption_RC4_EncryptByte(ByteArray(), Key)

End Sub
Public Sub Encryption_RC4_DecryptFile(SourceFile As String, DestFile As String, Optional Key As String)
Dim Filenr As Integer
Dim ByteArray() As Byte

    'Make sure the source file do exist
    If (Not Encryption_Misc_FileExist(SourceFile)) Then
        Call Err.Raise(vbObjectError, , "Sources Files don't exist.Error Decryptage.")
        Exit Sub
    End If

    'Open the source file and read the content
    'into a bytearray to decrypt
    Filenr = FreeFile
    Open SourceFile For Binary Access Read As #Filenr
    ReDim ByteArray(0 To LOF(Filenr) - 1)
    Get #Filenr, , ByteArray()
    Close #Filenr

    'Decrypt the bytearray
    Call Encryption_RC4_DecryptByte(ByteArray(), Key)

    'If the destination file already exist we need
    'to delete it since opening it for binary use
    'will preserve it if it already exist
    If (Encryption_Misc_FileExist(DestFile)) Then Kill DestFile

    'Store the decrypted data in the destination file
    Filenr = FreeFile
    Open DestFile For Binary Access Write As #Filenr
    Put #Filenr, , ByteArray()
    Close #Filenr

End Sub
Private Function Encryption_Misc_FileExist(FileName As String) As Boolean

    On Error GoTo NotExist

    Call FileLen(FileName)
    Encryption_Misc_FileExist = True
    
    On Error GoTo 0

NotExist:

End Function

Public Sub InitTimeGetTime()
'*****************************************************************
' Gets the offset time for the timer so we can start at 0 instead of
' the returned system time, allowing us to not have a time roll-over until
' the program is running for 25 days
'*****************************************************************
    ' Get the initial time
    GetSystemTime GetSystemTimeOffset
End Sub

Public Function TimeGetTime() As Long
'*****************************************************************
' Grabs the time from the 64-bit system timer and returns it in 32-bit
' after calculating it with the offset - allows us to have the
' "no roll-over" advantage of 64-bit timers with the RAM usage of 32-bit
' though we limit things slightly, so the rollover still happens, but after 25 days
'*****************************************************************
Dim CurrentTime As Currency
    ' Grab the current time (we have to pass a variable ByRef instead of a function return like the other timers)
    GetSystemTime CurrentTime
    
    ' Calculate the difference between the 64-bit times, return as a 32-bit time
    TimeGetTime = CurrentTime - GetSystemTimeOffset
End Function


Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)

    If LCase$(Dir(tDir & tName, vbDirectory)) <> tName Then Call MkDir(tDir & tName)
    
End Sub

Public Function FileExist(ByVal FileName As String, Optional RAW As Boolean = False) As Boolean

    If Not RAW Then
        If LenB(Dir(App.path & "\" & FileName)) > 0 Then
            FileExist = True
        End If

    Else

        If LenB(Dir(FileName)) > 0 Then
            FileExist = True
        End If
    End If

End Function

Sub LoadBloques()
Dim FileName As String, F As Integer

FileName = App.path & "\Data Files\Bloques.bin"
F = FreeFile

If Not FileExist(FileName, True) Then Exit Sub

Open FileName For Binary As #F
    Get #F, , EditorTile
Close

End Sub

Sub ClearMap()

    Call ZeroMemory(ByVal VarPtr(Map), LenB(Map))
    Map.name = vbNullString
    Map.MaxX = 32
    Map.MaxY = 32
    ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)
    initAutotiles
    
End Sub

Public Sub SaveMap(ByVal mapnum As Long)
Dim FileName As String
Dim F As Long
Dim X As Long
Dim Y As Long, I As Long, Z As Long, w As Long
    
    FileName = App.path & MAP_PATH & "map" & mapnum & MAP_EXT
    
    F = FreeFile
    Open FileName For Binary As #F
    Put #F, , Map.name
    Put #F, , Map.Music
    Put #F, , Map.BGS
    Put #F, , Map.Revision
    Put #F, , Map.Moral
    Put #F, , Map.Up
    Put #F, , Map.Down
    Put #F, , Map.Left
    Put #F, , Map.Right
    Put #F, , Map.BootMap
    Put #F, , Map.BootX
    Put #F, , Map.BootY
    
    Put #F, , Map.Weather
    Put #F, , Map.WeatherIntensity
    
    Put #F, , Map.Fog
    Put #F, , Map.FogSpeed
    Put #F, , Map.FogOpacity
    
    Put #F, , Map.Red
    Put #F, , Map.Green
    Put #F, , Map.Blue
    Put #F, , Map.Alpha
            
    Put #F, , Map.MaxX
    Put #F, , Map.MaxY
    
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            Put #F, , Map.Tile(X, Y)
        Next

        DoEvents
    Next

    For X = 1 To MAX_MAP_NPCS
        Put #F, , Map.NPC(X)
        Put #F, , Map.NpcSpawnType(X)
    Next
    
    Put #F, , Map.DayNight
    Put #F, , Map.Panorama
    Put #F, , Map.CanTp
    Put #F, , Map.IsDead
    Put #F, , Map.Gravity
    Put #F, , Map.Variables
    Put #F, , Map.VariablesCount
    
    Close #F
    
    Map_Changed = False
    
End Sub

Public Sub LoadMap(ByVal mapnum As Long)
Dim FileName As String
Dim F As Long
Dim X As Long
Dim Y As Long, I As Long, Z As Long, w As Long, p As Long

    GettingMap = True

    FileName = App.path & MAP_PATH & "map" & mapnum & MAP_EXT
    ClearMap
    F = FreeFile
    Open FileName For Binary As #F
    Get #F, , Map.name
    Get #F, , Map.Music
    Get #F, , Map.BGS
    Get #F, , Map.Revision
    Get #F, , Map.Moral
    Get #F, , Map.Up
    Get #F, , Map.Down
    Get #F, , Map.Left
    Get #F, , Map.Right
    Get #F, , Map.BootMap
    Get #F, , Map.BootX
    Get #F, , Map.BootY
    
    Get #F, , Map.Weather
    Get #F, , Map.WeatherIntensity
        
    Get #F, , Map.Fog
    Get #F, , Map.FogSpeed
    Get #F, , Map.FogOpacity
        
    Get #F, , Map.Red
    Get #F, , Map.Green
    Get #F, , Map.Blue
    Get #F, , Map.Alpha
    
    Get #F, , Map.MaxX
    Get #F, , Map.MaxY
    ' have to set the tile()
    ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)

    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            Get #F, , Map.Tile(X, Y)
        Next
    Next

    For X = 1 To MAX_MAP_NPCS
        Get #F, , Map.NPC(X)
        Get #F, , Map.NpcSpawnType(X)
    Next
    
    Get #F, , Map.DayNight
    Get #F, , Map.Panorama
    Get #F, , Map.CanTp
    Get #F, , Map.IsDead
    Get #F, , Map.Gravity
    Get #F, , Map.Variables
    Get #F, , Map.VariablesCount

    Close #F
    
    initAutotiles
    GettingMap = False

End Sub


Sub CheckMaps()
    Dim I As Long
    Dim AlreadyDid As Boolean

    For I = 1 To MAX_MAPS

        If Not FileExist("\data files\maps\map" & I & ".dat") Then
            If Not AlreadyDid Then
                Map.MaxX = 64
                Map.MaxY = 64
                ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)
            End If
            Call SaveMap(I)
        End If

    Next

End Sub

Sub SetCurrentMap(ByVal mapnum As Long)

    GettingMap = True
    
    CurrentMap = mapnum
    CurrentX = 0
    CurrentY = 0
    
    Call LoadMap(mapnum)
    Map_Changed = False
    
End Sub
