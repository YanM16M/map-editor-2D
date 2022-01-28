VERSION 5.00
Begin VB.Form frmEditor_Map 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Editor"
   ClientHeight    =   9615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14595
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   641
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   973
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdScreenshot 
      Caption         =   "Screenshot"
      Height          =   375
      Left            =   3600
      TabIndex        =   124
      Top             =   9120
      Width           =   1695
   End
   Begin VB.OptionButton optLayer 
      Caption         =   "Fringe 2"
      Height          =   255
      Index           =   5
      Left            =   6720
      TabIndex        =   123
      Top             =   1320
      Width           =   1215
   End
   Begin VB.OptionButton optLayer 
      Caption         =   "Fringe"
      Height          =   255
      Index           =   4
      Left            =   6720
      TabIndex        =   122
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdMapReport 
      Caption         =   "Liste des Maps"
      Height          =   375
      Left            =   1920
      TabIndex        =   121
      Top             =   9120
      Width           =   1575
   End
   Begin VB.CheckBox chkDirBlock 
      Caption         =   "DirBlock"
      Height          =   180
      Left            =   6600
      TabIndex        =   102
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton cmdSaveBloque 
      Caption         =   "SaveBloque"
      Height          =   255
      Left            =   6600
      TabIndex        =   101
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton cmdAutoBloque 
      Caption         =   "Auto Bloque"
      Height          =   390
      Left            =   6600
      TabIndex        =   100
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CheckBox chkPreBloque 
      Caption         =   "PreBloque"
      Height          =   255
      Left            =   6600
      TabIndex        =   99
      Top             =   6240
      Width           =   1455
   End
   Begin VB.PictureBox picAttributes 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7215
      Left            =   8160
      ScaleHeight     =   7215
      ScaleWidth      =   6375
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   6375
      Begin VB.Frame fraMapKey 
         Caption         =   "Map Key"
         Height          =   1815
         Left            =   1080
         TabIndex        =   32
         Top             =   2400
         Visible         =   0   'False
         Width           =   3375
         Begin VB.ComboBox lstMapKey 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   120
            TabIndex        =   119
            Text            =   "Combo1"
            Top             =   600
            Width           =   2535
         End
         Begin VB.PictureBox picMapKey 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   2760
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   36
            Top             =   600
            Width           =   480
         End
         Begin VB.CommandButton cmdMapKey 
            Caption         =   "Accept"
            Height          =   375
            Left            =   1080
            TabIndex        =   35
            Top             =   1320
            Width           =   1215
         End
         Begin VB.CheckBox chkMapKey 
            Caption         =   "Take key away upon use."
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   960
            Value           =   1  'Checked
            Width           =   2535
         End
         Begin VB.Label lblMapKey 
            Caption         =   "Item: None"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame frmObscur 
         Caption         =   "Obscur"
         Height          =   1455
         Left            =   840
         TabIndex        =   114
         Top             =   2520
         Visible         =   0   'False
         Width           =   4335
         Begin VB.CommandButton cmdOAccept 
            Caption         =   "Accept"
            Height          =   375
            Left            =   480
            TabIndex        =   117
            Top             =   960
            Width           =   3255
         End
         Begin VB.HScrollBar scrlOValue 
            Height          =   255
            Left            =   120
            Max             =   100
            TabIndex        =   115
            Top             =   480
            Width           =   3975
         End
         Begin VB.Label lblOValue 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Obscurité : 0%"
            Height          =   180
            Left            =   120
            TabIndex        =   116
            Top             =   240
            Width           =   1170
         End
      End
      Begin VB.Frame fraMapItem 
         Caption         =   "Map Item"
         Height          =   3975
         Left            =   1800
         TabIndex        =   28
         Top             =   2160
         Visible         =   0   'False
         Width           =   3375
         Begin VB.ListBox lstItem 
            Height          =   2220
            Left            =   120
            TabIndex        =   111
            Top             =   360
            Width           =   3015
         End
         Begin VB.CommandButton cmdMapItem 
            Caption         =   "Accept"
            Height          =   375
            Left            =   1080
            TabIndex        =   31
            Top             =   3480
            Width           =   1215
         End
         Begin VB.HScrollBar scrlMapItemValue 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   30
            Top             =   3120
            Value           =   1
            Width           =   2535
         End
         Begin VB.PictureBox picMapItem 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   2760
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   29
            Top             =   2880
            Width           =   480
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            Caption         =   "Value : 1"
            Height          =   180
            Left            =   240
            TabIndex        =   112
            Top             =   2880
            Width           =   690
         End
      End
      Begin VB.Frame fraNpcSpawn 
         Caption         =   "Npc Spawn"
         Height          =   3735
         Left            =   2040
         TabIndex        =   23
         Top             =   2040
         Visible         =   0   'False
         Width           =   3375
         Begin VB.ListBox lstNpc 
            Height          =   2220
            Left            =   240
            TabIndex        =   27
            Top             =   360
            Width           =   2895
         End
         Begin VB.HScrollBar scrlNpcDir 
            Height          =   255
            Left            =   120
            Max             =   3
            TabIndex        =   25
            Top             =   3000
            Width           =   2895
         End
         Begin VB.CommandButton cmdNpcSpawn 
            Caption         =   "Okay"
            Height          =   375
            Left            =   960
            TabIndex        =   24
            Top             =   3240
            Width           =   1455
         End
         Begin VB.Label lblNpcDir 
            Caption         =   "Direction: Up"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   2640
            Width           =   2535
         End
      End
      Begin VB.Frame frmLabel 
         Caption         =   "Label"
         Height          =   1095
         Left            =   960
         TabIndex        =   108
         Top             =   4920
         Width           =   4095
         Begin VB.CommandButton cmdLabel 
            Caption         =   "Okay"
            Height          =   255
            Left            =   1200
            TabIndex        =   110
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox txtLabel 
            Height          =   270
            Left            =   120
            TabIndex        =   109
            Top             =   360
            Width           =   3855
         End
      End
      Begin VB.Frame fraSlide 
         Caption         =   "Slide"
         Height          =   1455
         Left            =   1800
         TabIndex        =   68
         Top             =   2640
         Visible         =   0   'False
         Width           =   3375
         Begin VB.ComboBox cmbSlide 
            Height          =   300
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   70
            Top             =   360
            Width           =   2895
         End
         Begin VB.CommandButton cmdSlide 
            Caption         =   "Okay"
            Height          =   375
            Left            =   960
            TabIndex        =   69
            Top             =   840
            Width           =   1455
         End
      End
      Begin VB.Frame fraMapWarp 
         Caption         =   "Map Warp"
         Height          =   2895
         Left            =   120
         TabIndex        =   43
         Top             =   120
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CheckBox chkTP 
            Caption         =   "Animation de téléportation?"
            Height          =   735
            Left            =   1440
            TabIndex        =   97
            Top             =   2040
            Width           =   1815
         End
         Begin VB.CommandButton cmdMapWarp 
            Caption         =   "Accept"
            Height          =   375
            Left            =   120
            TabIndex        =   50
            Top             =   2040
            Width           =   1215
         End
         Begin VB.HScrollBar scrlMapWarpY 
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   1680
            Width           =   3135
         End
         Begin VB.HScrollBar scrlMapWarpX 
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   1080
            Width           =   3135
         End
         Begin VB.HScrollBar scrlMapWarp 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   45
            Top             =   480
            Value           =   1
            Width           =   3135
         End
         Begin VB.Label lblMapWarpY 
            Caption         =   "Y: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   1440
            Width           =   3135
         End
         Begin VB.Label lblMapWarpX 
            Caption         =   "X: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   840
            Width           =   3135
         End
         Begin VB.Label lblMapWarp 
            Caption         =   "Map: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame fraShop 
         Caption         =   "Shop"
         Height          =   1335
         Left            =   120
         TabIndex        =   51
         Top             =   3240
         Visible         =   0   'False
         Width           =   3135
         Begin VB.CommandButton cmdShop 
            Caption         =   "Accept"
            Height          =   375
            Left            =   960
            TabIndex        =   53
            Top             =   720
            Width           =   1215
         End
         Begin VB.ComboBox cmbShop 
            Height          =   300
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.Frame fraKeyOpen 
         Caption         =   "Key Open"
         Height          =   2055
         Left            =   240
         TabIndex        =   37
         Top             =   240
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton cmdKeyOpen 
            Caption         =   "Accept"
            Height          =   375
            Left            =   1080
            TabIndex        =   42
            Top             =   1440
            Width           =   1215
         End
         Begin VB.HScrollBar scrlKeyY 
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   1080
            Width           =   3015
         End
         Begin VB.HScrollBar scrlKeyX 
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   480
            Width           =   3015
         End
         Begin VB.Label lblKeyY 
            Caption         =   "Y: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   840
            Width           =   3015
         End
         Begin VB.Label lblKeyX 
            Caption         =   "X: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame fraLight 
         Caption         =   "Light"
         Height          =   2175
         Left            =   960
         TabIndex        =   76
         Top             =   1800
         Visible         =   0   'False
         Width           =   4215
         Begin VB.HScrollBar scrlB 
            Height          =   255
            Left            =   1560
            Max             =   255
            Min             =   1
            TabIndex        =   82
            Top             =   1320
            Value           =   1
            Width           =   1095
         End
         Begin VB.HScrollBar scrlG 
            Height          =   255
            Left            =   1560
            Max             =   255
            Min             =   1
            TabIndex        =   81
            Top             =   960
            Value           =   1
            Width           =   1095
         End
         Begin VB.PictureBox picLight 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1320
            Left            =   2760
            ScaleHeight     =   88
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   88
            TabIndex        =   80
            Top             =   240
            Width           =   1320
         End
         Begin VB.HScrollBar scrlA 
            Height          =   255
            Left            =   1560
            Max             =   255
            Min             =   1
            TabIndex        =   79
            Top             =   240
            Value           =   1
            Width           =   1095
         End
         Begin VB.HScrollBar scrlR 
            Height          =   255
            Left            =   1560
            Max             =   255
            Min             =   1
            TabIndex        =   78
            Top             =   600
            Value           =   1
            Width           =   1095
         End
         Begin VB.CommandButton cmdLight 
            Caption         =   "Accept"
            Height          =   375
            Left            =   1560
            TabIndex        =   77
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label lblB 
            Caption         =   "Blue: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   86
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label lblG 
            Caption         =   "Green: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   85
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lblR 
            Caption         =   "Red: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   84
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label lblA 
            Caption         =   "Alpha: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   83
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame fraSoundEffect 
         Caption         =   "Sound Effect"
         Height          =   1455
         Left            =   1800
         TabIndex        =   71
         Top             =   2040
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton cmdSoundEffectOk 
            Caption         =   "Okay"
            Height          =   375
            Left            =   960
            TabIndex        =   73
            Top             =   840
            Width           =   1455
         End
         Begin VB.ComboBox cmbSoundEffect 
            Height          =   300
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   72
            Top             =   360
            Width           =   2895
         End
      End
      Begin VB.Frame fraTrap 
         Caption         =   "Trap"
         Height          =   1575
         Left            =   1800
         TabIndex        =   64
         Top             =   2520
         Visible         =   0   'False
         Width           =   3375
         Begin VB.HScrollBar scrlTrap 
            Height          =   255
            Left            =   240
            Max             =   10000
            TabIndex        =   66
            Top             =   600
            Width           =   2895
         End
         Begin VB.CommandButton cmdTrap 
            Caption         =   "Okay"
            Height          =   375
            Left            =   960
            TabIndex        =   65
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label lblTrap 
            Caption         =   "Amount: 0"
            Height          =   255
            Left            =   240
            TabIndex        =   67
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.Frame fraHeal 
         Caption         =   "Heal"
         Height          =   1815
         Left            =   1800
         TabIndex        =   59
         Top             =   2400
         Visible         =   0   'False
         Width           =   3375
         Begin VB.ComboBox cmbHeal 
            Height          =   300
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   63
            Top             =   240
            Width           =   2895
         End
         Begin VB.CommandButton cmdHeal 
            Caption         =   "Okay"
            Height          =   375
            Left            =   960
            TabIndex        =   61
            Top             =   1200
            Width           =   1455
         End
         Begin VB.HScrollBar scrlHeal 
            Height          =   255
            Left            =   240
            Max             =   10000
            TabIndex        =   60
            Top             =   840
            Width           =   2895
         End
         Begin VB.Label lblHeal 
            Caption         =   "Amount: 0"
            Height          =   255
            Left            =   240
            TabIndex        =   62
            Top             =   600
            Width           =   2535
         End
      End
      Begin VB.Frame fraResource 
         Caption         =   "Object"
         Height          =   1695
         Left            =   1800
         TabIndex        =   19
         Top             =   2520
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton cmdResourceOk 
            Caption         =   "Okay"
            Height          =   375
            Left            =   960
            TabIndex        =   22
            Top             =   1080
            Width           =   1455
         End
         Begin VB.HScrollBar scrlResource 
            Height          =   255
            Left            =   240
            Max             =   100
            Min             =   1
            TabIndex        =   21
            Top             =   600
            Value           =   1
            Width           =   2895
         End
         Begin VB.Label lblResource 
            Caption         =   "Object:"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   360
            Width           =   2535
         End
      End
   End
   Begin VB.CommandButton cmdProperties 
      Caption         =   "Properties"
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   9120
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "Type"
      Height          =   1335
      Left            =   6600
      TabIndex        =   14
      Top             =   7680
      Width           =   1455
      Begin VB.OptionButton optBlock 
         Alignment       =   1  'Right Justify
         Caption         =   "Block"
         Height          =   255
         Left            =   480
         TabIndex        =   54
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton optAttribs 
         Alignment       =   1  'Right Justify
         Caption         =   "Attributes"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton optLayers 
         Alignment       =   1  'Right Justify
         Caption         =   "Layers"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.HScrollBar scrlPictureX 
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   7440
      Width           =   5895
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7200
      Left            =   120
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   399
      TabIndex        =   12
      Top             =   120
      Width           =   5985
   End
   Begin VB.VScrollBar scrlPictureY 
      Height          =   6855
      Left            =   6240
      Max             =   255
      TabIndex        =   11
      Top             =   240
      Width           =   255
   End
   Begin VB.Frame fraTileSet 
      Caption         =   "Tileset: 0"
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   7800
      Width           =   6495
      Begin VB.TextBox txtTile 
         Height          =   270
         Left            =   5520
         TabIndex        =   95
         Top             =   480
         Width           =   855
      End
      Begin VB.HScrollBar scrlTileSet 
         Height          =   255
         Left            =   120
         Max             =   10
         Min             =   1
         TabIndex        =   1
         Top             =   480
         Value           =   1
         Width           =   5295
      End
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Save"
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      Top             =   9120
      Width           =   1095
   End
   Begin VB.Frame fraLayers 
      Caption         =   "Layers"
      Height          =   6135
      Left            =   6600
      TabIndex        =   87
      Top             =   120
      Width           =   1455
      Begin VB.CheckBox chkApercu 
         Caption         =   "Apercu ?"
         Height          =   255
         Left            =   120
         TabIndex        =   120
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Retour"
         Height          =   375
         Left            =   120
         TabIndex        =   98
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CheckBox chkShow 
         Caption         =   "TileView"
         Height          =   255
         Left            =   120
         TabIndex        =   96
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Efface(toute la couche)"
         Height          =   375
         Left            =   120
         TabIndex        =   93
         Top             =   5160
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Mask2"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   92
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Ground"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   91
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Mask"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   90
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdFill 
         Caption         =   "Remplir"
         Height          =   390
         Left            =   120
         TabIndex        =   89
         Top             =   5640
         Width           =   1215
      End
      Begin VB.HScrollBar scrlAutotile 
         Height          =   255
         Left            =   120
         Max             =   5
         TabIndex        =   88
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label lblY 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CurY :"
         Height          =   180
         Left            =   120
         TabIndex        =   126
         Top             =   4080
         Width           =   480
      End
      Begin VB.Label lblX 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CurX :"
         Height          =   180
         Left            =   120
         TabIndex        =   125
         Top             =   3840
         Width           =   480
      End
      Begin VB.Label lblAutotile 
         Alignment       =   2  'Center
         Caption         =   "Normal"
         Height          =   255
         Left            =   120
         TabIndex        =   94
         Top             =   2160
         Width           =   1215
      End
   End
   Begin VB.Frame fraAttribs 
      Caption         =   "Attributes"
      Height          =   6015
      Left            =   6600
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
      Begin VB.OptionButton optLampe 
         Caption         =   "Lampe"
         Height          =   255
         Left            =   120
         TabIndex        =   118
         Top             =   5280
         Width           =   1215
      End
      Begin VB.OptionButton optObscur 
         Caption         =   "Obscur"
         Height          =   255
         Left            =   120
         TabIndex        =   113
         Top             =   5040
         Width           =   1215
      End
      Begin VB.OptionButton optLabel 
         Caption         =   "Label"
         Height          =   255
         Left            =   120
         TabIndex        =   107
         Top             =   4080
         Width           =   1215
      End
      Begin VB.OptionButton optDB 
         Caption         =   "DragonBall"
         Height          =   255
         Left            =   120
         TabIndex        =   106
         Top             =   1920
         Width           =   1215
      End
      Begin VB.OptionButton optVol2 
         Caption         =   "Vol : Sol"
         Height          =   255
         Left            =   120
         TabIndex        =   105
         Top             =   4800
         Width           =   1215
      End
      Begin VB.OptionButton optVol 
         Caption         =   "Vole : Air"
         Height          =   255
         Left            =   120
         TabIndex        =   104
         Top             =   4560
         Width           =   1095
      End
      Begin VB.OptionButton optNoFight 
         Caption         =   "NoFight"
         Height          =   255
         Left            =   120
         TabIndex        =   103
         Top             =   4320
         Width           =   1095
      End
      Begin VB.OptionButton optLight 
         Caption         =   "Lumière"
         Height          =   270
         Left            =   120
         TabIndex        =   75
         Top             =   3840
         Width           =   1215
      End
      Begin VB.OptionButton optSound 
         Caption         =   "Son"
         Height          =   270
         Left            =   120
         TabIndex        =   74
         Top             =   3600
         Width           =   1215
      End
      Begin VB.OptionButton optSlide 
         Caption         =   "Glisse"
         Height          =   270
         Left            =   120
         TabIndex        =   58
         Top             =   3360
         Width           =   1215
      End
      Begin VB.OptionButton optTrap 
         Caption         =   "Piège"
         Height          =   270
         Left            =   120
         TabIndex        =   57
         Top             =   3120
         Width           =   1215
      End
      Begin VB.OptionButton optHeal 
         Caption         =   "Soin"
         Height          =   270
         Left            =   120
         TabIndex        =   56
         Top             =   2880
         Width           =   1215
      End
      Begin VB.OptionButton optBank 
         Caption         =   "Banque"
         Height          =   270
         Left            =   120
         TabIndex        =   55
         Top             =   2640
         Width           =   1215
      End
      Begin VB.OptionButton optResource 
         Caption         =   "Resource"
         Height          =   240
         Left            =   120
         TabIndex        =   17
         Top             =   1680
         Width           =   1215
      End
      Begin VB.OptionButton optKeyOpen 
         Caption         =   "Key Open"
         Height          =   240
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton optBlocked 
         Caption         =   "Bloque"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optWarp 
         Caption         =   "Téléporte"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear2 
         Caption         =   "Clear"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   5640
         Width           =   1215
      End
      Begin VB.OptionButton optNpcAvoid 
         Caption         =   "Npc Avoid"
         Height          =   270
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton optKey 
         Caption         =   "Clé"
         Height          =   270
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmEditor_Map"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private currentTiles As Integer

Private Sub chkBloqueDirection_Click()
    picAttributes.Visible = True
End Sub

Private Sub cmdAutoBloque_Click()
Dim X As Integer, Y As Integer, I As Integer, Z As Byte
Dim Opposite As Byte

If MsgBox("Placer les bloques automatiquement ?", vbYesNo) = vbNo Then
    Exit Sub
End If

'clear
For X = 0 To Map.MaxX
    For Y = 0 To Map.MaxY
        Map.Tile(X, Y).DirBlock = 0
    Next
Next

For I = 1 To MapLayer.Layer_Count - 1
    For Y = 0 To Map.MaxY
        For X = 0 To Map.MaxX
            If Map.Tile(X, Y).layer(I).Tileset > 0 Then
                'Make tile_type_blocked
                If Map.Tile(X, Y).Type = TILE_TYPE_WALKABLE Then
                    If EditorTile(Map.Tile(X, Y).layer(I).Tileset, Map.Tile(X, Y).layer(I).X, Map.Tile(X, Y).layer(I).Y).Type = TILE_TYPE_BLOCKED Then
                        Map.Tile(X, Y).Type = TILE_TYPE_BLOCKED
                    End If
                End If
                'dirblock
                If EditorTile(Map.Tile(X, Y).layer(I).Tileset, Map.Tile(X, Y).layer(I).X, Map.Tile(X, Y).layer(I).Y).DirBlock > 0 Then
                    If Map.Tile(X, Y).DirBlock < 1 Then
                        Map.Tile(X, Y).DirBlock = EditorTile(Map.Tile(X, Y).layer(I).Tileset, Map.Tile(X, Y).layer(I).X, Map.Tile(X, Y).layer(I).Y).DirBlock
                    Else 'si deja un dirblock "additionne les 2"
                        For Z = 0 To 7
                            If isDirBlocked(EditorTile(Map.Tile(X, Y).layer(I).Tileset, Map.Tile(X, Y).layer(I).X, Map.Tile(X, Y).layer(I).Y).DirBlock, CByte(Z)) Then
                                setDirBlock Map.Tile(X, Y).DirBlock, CByte(Z), True
                            End If
                        Next
                    End If
                End If
            End If
        Next
    Next
Next

'Make opposite dirblock
For I = 1 To MapLayer.Layer_Count - 1
    For Y = 0 To Map.MaxY
        For X = 0 To Map.MaxX
            If Map.Tile(X, Y).layer(I).Tileset > 0 Then
                If EditorTile(Map.Tile(X, Y).layer(I).Tileset, Map.Tile(X, Y).layer(I).X, Map.Tile(X, Y).layer(I).Y).DirBlock > 0 Then
                    For Z = 0 To 7
                        If isDirBlocked(EditorTile(Map.Tile(X, Y).layer(I).Tileset, Map.Tile(X, Y).layer(I).X, Map.Tile(X, Y).layer(I).Y).DirBlock, CByte(Z)) Then
                            Opposite = GetOpposite(Z)
                            Select Case Z '
                                Case DIR_UP
                                    If Y - 1 >= 0 Then
                                        setDirBlock Map.Tile(X, Y - 1).DirBlock, CByte(Opposite), True
                                    End If
                                Case DIR_LEFT
                                    If X - 1 >= 0 Then
                                        setDirBlock Map.Tile(X - 1, Y).DirBlock, CByte(Opposite), True
                                    End If
                                Case DIR_DOWN
                                    If Y + 1 <= Map.MaxY Then
                                        setDirBlock Map.Tile(X, Y + 1).DirBlock, CByte(Opposite), True
                                    End If
                                Case DIR_RIGHT
                                    If X + 1 <= Map.MaxX Then
                                        setDirBlock Map.Tile(X + 1, Y).DirBlock, CByte(Opposite), True
                                    End If
                                Case DIR_UP_LEFT
                                    If Y - 1 >= 0 And X - 1 >= 0 Then
                                        setDirBlock Map.Tile(X - 1, Y - 1).DirBlock, CByte(Opposite), True
                                    End If
                                Case DIR_UP_RIGHT
                                    If Y - 1 >= 0 And X + 1 <= Map.MaxX Then
                                        setDirBlock Map.Tile(X + 1, Y - 1).DirBlock, CByte(Opposite), True
                                    End If
                                Case DIR_DOWN_LEFT
                                    If Y + 1 <= Map.MaxY And X - 1 >= 0 Then
                                        setDirBlock Map.Tile(X - 1, Y + 1).DirBlock, CByte(Opposite), True
                                    End If
                                Case DIR_DOWN_RIGHT
                                    If Y + 1 <= Map.MaxY And X + 1 <= Map.MaxX Then
                                        setDirBlock Map.Tile(X + 1, Y + 1).DirBlock, CByte(Opposite), True
                                    End If
                            End Select
                        End If
                    Next
                End If
            End If
        Next
    Next
Next

End Sub

Private Sub cmdHeal_Click()

    MapEditorHealType = cmbHeal.ListIndex + 1
    MapEditorHealAmount = scrlHeal.Value
    picAttributes.Visible = False
    fraHeal.Visible = False
    
End Sub

Private Sub cmdKeyOpen_Click()

    KeyOpenEditorX = scrlKeyX.Value
    KeyOpenEditorY = scrlKeyY.Value
    picAttributes.Visible = False
    fraKeyOpen.Visible = False
    

End Sub

Private Sub cmdLabel_Click()
    LabelMap = Trim$(txtLabel.text)
    picAttributes.Visible = False
    frmLabel.Visible = False
End Sub

Private Sub cmdLight_Click()

    picAttributes.Visible = False
    fraLight.Visible = False
    
End Sub
Private Sub cmdMapItem_Click()

    ItemEditorNum = lstItem.ListIndex
    ItemEditorValue = scrlMapItemValue.Value
    picAttributes.Visible = False
    fraMapItem.Visible = False
    

End Sub

Private Sub cmdMapKey_Click()

    KeyEditorNum = lstMapKey.ListIndex
    KeyEditorTake = chkMapKey.Value
    picAttributes.Visible = False
    fraMapKey.Visible = False

End Sub

Private Sub cmdMapReport_Click()
    MapReportInit
End Sub

Private Sub cmdMapWarp_Click()

    EditorWarpMap = scrlMapWarp.Value
    EditorWarpX = scrlMapWarpX.Value
    EditorWarpY = scrlMapWarpY.Value
    EditorWarpTp = chkTP.Value
    picAttributes.Visible = False
    fraMapWarp.Visible = False

End Sub


Private Sub cmdNpcSpawn_Click()

    SpawnNpcNum = lstNpc.ListIndex + 1
    SpawnNpcDir = scrlNpcDir.Value
    picAttributes.Visible = False
    fraNpcSpawn.Visible = False
    

End Sub



Private Sub cmdOAccept_Click()
    EditorObscurAlpha = scrlOValue.Value
    picAttributes.Visible = False
    frmObscur.Visible = False
End Sub


Private Sub cmdResourceOk_Click()

    ResourceEditorNum = scrlResource.Value
    picAttributes.Visible = False
    fraResource.Visible = False

End Sub

Private Sub cmdSaveBloque_Click()
Dim FileName As String, F As Integer

FileName = App.path & "\Data Files\bloques.bin"
F = FreeFile

If MsgBox("Tu es sûr de vouloir faire ça?", vbYesNo) = vbYes Then
    Open FileName For Binary As #F
        Put #F, , EditorTile
    Close
End If
        
End Sub

Private Sub cmdScreenshot_Click()
    Call TakeScreenshot
End Sub

Private Sub cmdShop_Click()

    EditorShop = cmbShop.ListIndex
    picAttributes.Visible = False
    fraShop.Visible = False
    

End Sub

Private Sub cmdSlide_Click()

    MapEditorSlideDir = cmbSlide.ListIndex
    picAttributes.Visible = False
    fraSlide.Visible = False
    

End Sub



Private Sub cmdTrap_Click()

    MapEditorHealAmount = scrlTrap.Value
    picAttributes.Visible = False
    fraTrap.Visible = False
    

End Sub

Private Sub Command1_Click()
   Call SaveMap(CurrentMap)
End Sub


Private Sub Command2_Click()
Dim I As Integer

For I = 1 To MAX_BYTE
    If LastCurlayer(I) > 0 Then
        With Map.Tile(LastX(I), LastY(I))
            .layer(LastCurlayer(I)).X = LastTileX(I)
            .layer(LastCurlayer(I)).Y = LastTileY(I)
            .layer(LastCurlayer(I)).Tileset = LastTileset(I)
            .Autotile(LastCurlayer(I)) = 0
            CacheRenderState LastX(I), LastY(I), LastCurlayer(I)
            
            LastCurlayer(I) = 0
            LastClick(I) = False
        End With
    End If
Next

                
End Sub

Private Sub Form_Load()
Dim I As Long

    ' move the entire attributes box on screen
    picAttributes.Left = 8
    picAttributes.Top = 8


End Sub


Private Sub optDoor_Click()
    ' If debug mode, handle error then exit out

    ClearAttributeDialogue
    picAttributes.Visible = True
    fraMapWarp.Visible = True
    
    scrlMapWarp.Max = MAX_MAPS
    scrlMapWarp.Value = 1
    scrlMapWarpX.Max = MAX_BYTE
    scrlMapWarpY.Max = MAX_BYTE
    scrlMapWarpX.Value = 0
    scrlMapWarpY.Value = 0
    

End Sub

Private Sub optHeal_Click()

    ClearAttributeDialogue
    picAttributes.Visible = True
    fraHeal.Visible = True

End Sub

Private Sub optLabel_Click()
    
    ClearAttributeDialogue
    picAttributes.Visible = True
    frmLabel.Visible = True
    
    txtLabel.text = vbNullString
    
End Sub

Private Sub optLayer_Click(Index As Integer)
Dim I As Byte

For I = 1 To MapLayer.Layer_Count - 1
    If I <> Index Then
        optLayer(I).Value = False
    End If
Next
    
End Sub

Private Sub optLayers_Click()

    If optLayers.Value Then
        fraLayers.Visible = True
        fraAttribs.Visible = False
    End If

End Sub

Private Sub optAttribs_Click()

    If optAttribs.Value Then
        fraLayers.Visible = False
        fraAttribs.Visible = True
    End If

End Sub

Private Sub optLight_Click()

        ClearAttributeDialogue
        picAttributes.Visible = True
        fraLight.Visible = True

End Sub

Private Sub optObscur_Click()

    ClearAttributeDialogue
    picAttributes.Visible = True
    frmObscur.Visible = True
        
End Sub

Private Sub optResource_Click()

    ClearAttributeDialogue
    picAttributes.Visible = True
    fraResource.Visible = True
    

End Sub

Private Sub optShop_Click()

    ClearAttributeDialogue
    picAttributes.Visible = True
    fraShop.Visible = True
    
 
End Sub

Private Sub optSlide_Click()

    ClearAttributeDialogue
    picAttributes.Visible = True
    fraSlide.Visible = True
    
 
End Sub

Private Sub optSound_Click()

    ClearAttributeDialogue
    picAttributes.Visible = True
    fraSoundEffect.Visible = True
    

End Sub

Private Sub optTrap_Click()

    ClearAttributeDialogue
    picAttributes.Visible = True
    fraTrap.Visible = True
 
End Sub

Private Sub cmdSend_Click()
    Call SaveMap(CurrentMap)
    Call MsgBox("Map sauvegardé", vbOKOnly, "Editeur de Map")
End Sub


Private Sub cmdProperties_Click()

    Load frmEditor_MapProperties
    MapEditorProperties
    frmEditor_MapProperties.Show vbModal
    

End Sub

Private Sub optWarp_Click()

    ClearAttributeDialogue
    picAttributes.Visible = True
    fraMapWarp.Visible = True
    
    scrlMapWarp.Max = MAX_MAPS
    scrlMapWarp.Value = 1
    scrlMapWarpX.Max = MAX_BYTE
    scrlMapWarpY.Max = MAX_BYTE
    scrlMapWarpX.Value = 0
    scrlMapWarpY.Value = 0

End Sub




Private Sub optKeyOpen_Click()

    ClearAttributeDialogue
    fraKeyOpen.Visible = True
    picAttributes.Visible = True
    
    scrlKeyX.Max = Map.MaxX
    scrlKeyY.Max = Map.MaxY
    scrlKeyX.Value = 0
    scrlKeyY.Value = 0
    
 
End Sub

Private Sub cmdFill_Click()

    MapEditorFillLayer
    

End Sub

Private Sub cmdClear_Click()

    Call MapEditorClearLayer
 
End Sub

Private Sub cmdClear2_Click()

    Call MapEditorClearAttribs
    

End Sub

Private Sub picBack_KeyDown(KeyCode As Integer, Shift As Integer)
    If chkDirBlock.Value Then
        HandleKeyDown KeyCode
    End If
End Sub

Private Sub picBack_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)

    X = X + (frmEditor_Map.scrlPictureX.Value * 32)
    Y = Y + (frmEditor_Map.scrlPictureY.Value * 32)
    Call MapEditorChooseTile(button, X, Y)
    
    If frmEditor_Map.chkPreBloque.Value Then
        If button = vbLeftButton Then
            EditorTile(frmEditor_Map.scrlTileSet, EditorTileX, EditorTileY).Type = TILE_TYPE_BLOCKED
        ElseIf button = vbRightButton Then
            EditorTile(frmEditor_Map.scrlTileSet, EditorTileX, EditorTileY).Type = TILE_TYPE_WALKABLE
        End If
    End If
        
End Sub

Private Sub picBack_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)

    X = X + (frmEditor_Map.scrlPictureX.Value * 32)
    Y = Y + (frmEditor_Map.scrlPictureY.Value * 32)
    Call MapEditorDrag(button, X, Y)
    
 
End Sub

Private Sub scrlA_Change()

    lblA.Caption = "Alpha: " & scrlA.Value
    MapEditorLightA = scrlA.Value

End Sub

Private Sub scrlOValue_Change()
    lblOValue.Caption = "Obscurité : " & scrlOValue.Value & "%"
End Sub

Private Sub scrlR_Change()

    lblR.Caption = "Red: " & ScrlR.Value
    MapEditorLightR = ScrlR.Value
    

End Sub

Private Sub scrlG_Change()
 
    lblG.Caption = "Green: " & ScrlG.Value
    MapEditorLightG = ScrlG.Value
    

End Sub

Private Sub scrlB_Change()

    lblB.Caption = "Blue: " & ScrlB.Value
    MapEditorLightB = ScrlB.Value
    

End Sub


Private Sub scrlAutotile_Change()
    Select Case scrlAutotile.Value
        Case 0 ' normal
            lblAutotile.Caption = "Normal"
        Case 1 ' autotile
            lblAutotile.Caption = "Autotile (VX)"
        Case 2 ' fake autotile
            lblAutotile.Caption = "Animation (VX)"
        Case 3 ' animated
            lblAutotile.Caption = "Animated (VX)"
        Case 4 ' cliff
            lblAutotile.Caption = "Cliff (VX)"
        Case 5 ' waterfall
            lblAutotile.Caption = "Waterfall (VX)"
        Case 6 ' autotile
            lblAutotile.Caption = "Autotile (XP)"
        Case 7 ' fake autotile
            lblAutotile.Caption = "Fake (XP)"
        Case 8 ' animated
            lblAutotile.Caption = "Animated (XP)"
        Case 9 ' cliff
            lblAutotile.Caption = "Cliff (XP)"
        Case 10 ' waterfall
            lblAutotile.Caption = "Waterfall (XP)"
    End Select
End Sub


Private Sub scrlHeal_Change()

    lblHeal.Caption = "Amount: " & scrlHeal.Value
    

End Sub

Private Sub scrlKeyX_Change()

    lblKeyX.Caption = "X: " & scrlKeyX.Value

End Sub

Private Sub scrlKeyX_Scroll()

    scrlKeyX_Change
 
End Sub

Private Sub scrlKeyY_Change()

    lblKeyY.Caption = "Y: " & scrlKeyY.Value
    
 
End Sub

Private Sub scrlKeyY_Scroll()

    scrlKeyY_Change
    

End Sub

Private Sub scrlTrap_Change()
 
    lblTrap.Caption = "Amount: " & scrlTrap.Value
  
End Sub




Private Sub scrlMapWarp_Change()
  
    lblMapWarp.Caption = "Map: " & scrlMapWarp.Value
    
 
End Sub

Private Sub scrlMapWarp_Scroll()
 
    scrlMapWarp_Change
    
   
End Sub

Private Sub scrlMapWarpX_Change()
 
    lblMapWarpX.Caption = "X: " & scrlMapWarpX.Value
    
End Sub

Private Sub scrlMapWarpX_Scroll()

    scrlMapWarpX_Change
    
 
End Sub

Private Sub scrlMapWarpY_Change()
  
    lblMapWarpY.Caption = "Y: " & scrlMapWarpY.Value
    
    
End Sub

Private Sub scrlMapWarpY_Scroll()

    scrlMapWarpY_Change
    

    Exit Sub
End Sub

Private Sub scrlNpcDir_Change()
 
    Select Case scrlNpcDir.Value
        Case DIR_DOWN
            lblNpcDir = "Direction: Down"
        Case DIR_UP
            lblNpcDir = "Direction: Up"
        Case DIR_LEFT
            lblNpcDir = "Direction: Left"
        Case DIR_RIGHT
            lblNpcDir = "Direction: Right"
    End Select
    
 
End Sub

Private Sub scrlNpcDir_Scroll()

    scrlNpcDir_Change
    
 
End Sub




Private Sub scrlPictureX_Change()
 
    Call MapEditorTileScroll
    

End Sub

Private Sub scrlPictureY_Change()
 
    Call MapEditorTileScroll
 
End Sub

Private Sub scrlPictureX_Scroll()

    scrlPictureY_Change
  
End Sub

Private Sub scrlPictureY_Scroll()

    scrlPictureY_Change
 
End Sub

Private Sub scrlTileset_Change()

    fraTileSet.Caption = "Tileset: " & scrlTileSet.Value
    
    frmEditor_Map.scrlPictureY.Max = (Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Height \ 32) - (frmEditor_Map.picBack.Height \ 32)
    frmEditor_Map.scrlPictureX.Max = (Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Width \ 32) - (frmEditor_Map.picBack.Width \ 32)
    
    MapEditorTileScroll
    
    EditorTileX = 0
    EditorTileY = 0
    EditorTileWidth = 1
    EditorTileHeight = 1
    

End Sub

Private Sub scrlTileSet_Scroll()

    scrlTileset_Change
    
    
End Sub

Private Sub txtTile_Change()
    If IsNumeric(txtTile.text) Then
        If txtTile.text > NumTilesets Then txtTile.text = NumTilesets
        If txtTile.text <= 0 Then txtTile.text = 1
        scrlTileSet.Value = txtTile.text
        scrlTileSet_Scroll
    Else
        txtTile.text = 1
    End If
End Sub
