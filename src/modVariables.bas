Attribute VB_Name = "modVariables"
Option Explicit

Public BloC As Boolean
Public Map_Changed As Boolean
Public GettingMap As Boolean

' FPS and Time-based movement vars
Public ElapsedTime As Long
Public GameFPS As Long


Public Const MAX_BYTE = 255

'Camera
Public Camera As RECT
Public TileView As RECT

' Cursor location
Public GlobalX As Long
Public GlobalY As Long
Public GlobalX_Map As Long
Public GlobalY_Map As Long
Public CurX As Long
Public CurY As Long

' Location Fake Player
Public CurrentMap As Byte
Public CurrentX As Long
Public CurrentY As Long

' Fog
Public fogOffsetX As Long
Public fogOffsetY As Long

'Utilisé pour l'animation des tiles
Public CurTiles As Byte

' for directional blocking
Public DirArrowX(0 To 7) As Byte
Public DirArrowY(0 To 7) As Byte

' Game editors
Public Editor As Byte
Public EditorIndex As Long
Public AnimEditorFrame(0 To 1) As Byte
Public AnimEditorTimer(0 To 1) As Long

' Used to check if in editor or not and variables for use in editor
Public InMapEditor As Boolean
Public EditorTileX As Long
Public EditorTileY As Long
Public EditorTileWidth As Long
Public EditorTileHeight As Long

'resource
Public ResourceTileX As Integer
Public ResourceTileY As Integer
Public ResourceTileset As Integer
Public currentIndexResource As Byte

'obscru
Public EditorObscurAlpha As Byte

'Warp Attribute
Public EditorWarpMap As Long
Public EditorWarpX As Long
Public EditorWarpY As Long
Public EditorWarpTp As Byte

'Label Attribute
Public LabelMap As String

'Attributes
Public SpawnNpcNum As Long
Public SpawnNpcDir As Byte
Public EditorShop As Long

'return
Public LastX(1 To MAX_BYTE) As Byte
Public LastY(1 To MAX_BYTE) As Byte
Public LastTileX(1 To MAX_BYTE) As Byte
Public LastTileY(1 To MAX_BYTE) As Byte
Public LastTileset(1 To MAX_BYTE) As Byte
Public LastCurlayer(1 To MAX_BYTE) As Byte
Public LastClick(1 To MAX_BYTE) As Boolean

' Used for map item editor
Public ItemEditorNum As Long
Public ItemEditorValue As Long

' Used for map key editor
Public KeyEditorNum As Long
Public KeyEditorTake As Long

' Used for map key open editor
Public KeyOpenEditorX As Long
Public KeyOpenEditorY As Long

' Map Resources
Public ResourceEditorNum As Long

' Used for map editor heal & trap & slide tiles
Public MapEditorHealType As Long
Public MapEditorHealAmount As Long
Public MapEditorSlideDir As Long
Public MapEditorSound As String
Public MapEditorLightA As Long
Public MapEditorLightR As Long
Public MapEditorLightG As Long
Public MapEditorLightB As Long


'#######################################
'#### CONSTANTATES
'#######################################

'Fonts path
Public Const FONT_PATH As String = "\data files\graphics\fonts\"

' String constants
Public Const NAME_LENGTH As Byte = 20
Public Const MUSIC_LENGTH As Byte = 40
Public Const ACCOUNT_LENGTH As Byte = 12

' Data Maps Constantes
Public Const MAX_MAPS As Byte = 30
Public Const MAX_MAP_NPCS As Byte = 30
' Map Path and variables
Public Const MAP_PATH As String = "\Data Files\maps\"
Public Const MAP_EXT As String = ".dat"

' Gfx Path and variables
Public Const GFX_PATH As String = "\Data Files\graphics\"
Public Const GFX_EXT As String = ".png"

' ****** PI ******
Public Const DegreeToRadian As Single = 0.0174532919296  'Pi / 180
Public Const RadianToDegree As Single = 57.2958300962816 '180 / Pi

Public Const MAP_MORAL_NONE As Byte = 0
Public Const MAP_MORAL_SAFE As Byte = 1
Public Const MAP_MORAL_PARTY_MAP As Byte = 2
Public Const MAP_MORAL_SAFE_EQUIPEMENT As Byte = 3

' Tile consants
Public Const TILE_TYPE_WALKABLE As Byte = 0
Public Const TILE_TYPE_BLOCKED As Byte = 1
Public Const TILE_TYPE_WARP As Byte = 2
Public Const TILE_TYPE_ITEM As Byte = 3
Public Const TILE_TYPE_NPCAVOID As Byte = 4
Public Const TILE_TYPE_KEY As Byte = 5
Public Const TILE_TYPE_KEYOPEN As Byte = 6
Public Const TILE_TYPE_RESOURCE As Byte = 7
Public Const TILE_TYPE_DOOR As Byte = 8
Public Const TILE_TYPE_NPCSPAWN As Byte = 9
Public Const TILE_TYPE_SHOP As Byte = 10
Public Const TILE_TYPE_BANK As Byte = 11
Public Const TILE_TYPE_HEAL As Byte = 12
Public Const TILE_TYPE_TRAP As Byte = 13
Public Const TILE_TYPE_SLIDE As Byte = 14
Public Const TILE_TYPE_SOUND As Byte = 15
Public Const TILE_TYPE_LIGHT As Byte = 16
Public Const TILE_TYPE_CRAFT As Byte = 17
Public Const TILE_TYPE_NOFIGHT As Byte = 18
Public Const TILE_TYPE_VOL As Byte = 19
Public Const TILE_TYPE_VOL2 As Byte = 20
Public Const TILE_TYPE_DB As Byte = 21
Public Const TILE_TYPE_LABEL As Byte = 22
Public Const TILE_TYPE_OBSCUR As Byte = 23
Public Const TILE_TYPE_LAMPE As Byte = 24

' text color constants
Public Const Black As Byte = 0
Public Const Blue As Byte = 1
Public Const Green As Byte = 2
Public Const Cyan As Byte = 3
Public Const Red As Byte = 4
Public Const Magenta As Byte = 5
Public Const Brown As Byte = 6
Public Const Grey As Byte = 7
Public Const DarkGrey As Byte = 8
Public Const BrightBlue As Byte = 9
Public Const BrightGreen As Byte = 10
Public Const BrightCyan As Byte = 11
Public Const BrightRed As Byte = 12
Public Const Pink As Byte = 13
Public Const Yellow As Byte = 14
Public Const White As Byte = 15
Public Const DarkBrown As Byte = 16
Public Const Orange As Byte = 17
Public Const Pinky As Byte = 18

' Autotiles
Public Const AUTO_INNER As Byte = 1
Public Const AUTO_OUTER As Byte = 2
Public Const AUTO_HORIZONTAL As Byte = 3
Public Const AUTO_VERTICAL As Byte = 4
Public Const AUTO_FILL As Byte = 5

' Autotile types
Public Const AUTOTILE_NONE As Byte = 0
Public Const AUTOTILE_NORMAL As Byte = 1
Public Const AUTOTILE_FAKE As Byte = 2
Public Const AUTOTILE_ANIM As Byte = 3
Public Const AUTOTILE_CLIFF As Byte = 4
Public Const AUTOTILE_WATERFALL As Byte = 5

' Rendering
Public Const RENDER_STATE_NONE As Byte = 0
Public Const RENDER_STATE_NORMAL As Byte = 1
Public Const RENDER_STATE_AUTOTILE As Byte = 2
Public Const RENDER_STATE_ANIMATE As Byte = 3

' Direction constants
Public Const DIR_UP As Byte = 0
Public Const DIR_DOWN As Byte = 1
Public Const DIR_LEFT As Byte = 2
Public Const DIR_RIGHT As Byte = 3
Public Const DIR_UP_LEFT As Byte = 4
Public Const DIR_UP_RIGHT As Byte = 5
Public Const DIR_DOWN_LEFT As Byte = 6
Public Const DIR_DOWN_RIGHT As Byte = 7
