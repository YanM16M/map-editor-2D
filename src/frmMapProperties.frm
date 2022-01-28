VERSION 5.00
Begin VB.Form frmEditor_MapProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Properties"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   425
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   439
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.HScrollBar scrlGravity 
      Height          =   255
      Left            =   120
      Max             =   100
      TabIndex        =   47
      Top             =   6000
      Width           =   2055
   End
   Begin VB.Frame Frame7 
      Caption         =   "Fog"
      Height          =   2415
      Left            =   4440
      TabIndex        =   34
      Top             =   1560
      Width           =   2055
      Begin VB.HScrollBar scrlFogOpacity 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   39
         Top             =   1620
         Width           =   1575
      End
      Begin VB.HScrollBar ScrlFog 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   37
         Top             =   480
         Width           =   1575
      End
      Begin VB.HScrollBar ScrlFogSpeed 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   35
         Top             =   1050
         Width           =   1575
      End
      Begin VB.Label lblFogOpacity 
         Caption         =   "Fog Opacity: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   1380
         Width           =   1815
      End
      Begin VB.Label lblFog 
         Caption         =   "Fog: None"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblFogSpeed 
         Caption         =   "Fog Speed: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   810
         Width           =   1815
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Map Overlay"
      Height          =   2775
      Left            =   2280
      TabIndex        =   25
      Top             =   1560
      Width           =   2055
      Begin VB.CheckBox chkDead 
         Caption         =   "Map obscure?"
         Height          =   195
         Left            =   120
         TabIndex        =   46
         Top             =   2400
         Width           =   1815
      End
      Begin VB.CheckBox chkTP 
         Caption         =   "TP interdit?"
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   2040
         Width           =   1815
      End
      Begin VB.ComboBox cmbPanorama 
         Height          =   315
         Left            =   120
         TabIndex        =   43
         Text            =   "cmbPanorama"
         Top             =   1560
         Width           =   1815
      End
      Begin VB.HScrollBar scrlA 
         Height          =   255
         Left            =   1080
         Max             =   255
         TabIndex        =   32
         Top             =   960
         Width           =   855
      End
      Begin VB.HScrollBar ScrlR 
         Height          =   255
         Left            =   1080
         Max             =   255
         TabIndex        =   28
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar ScrlG 
         Height          =   255
         Left            =   1080
         Max             =   255
         TabIndex        =   27
         Top             =   480
         Width           =   855
      End
      Begin VB.HScrollBar ScrlB 
         Height          =   255
         Left            =   1080
         Max             =   255
         TabIndex        =   26
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblPanorama 
         Caption         =   "Panorama:"
         Height          =   255
         Left            =   480
         TabIndex        =   44
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblA 
         Caption         =   "Opacity: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblR 
         Caption         =   "Red: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblG 
         Caption         =   "Green: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblB 
         Caption         =   "Blue: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Frame frmMaxSizes 
      Caption         =   "Max Sizes"
      Height          =   975
      Left            =   120
      TabIndex        =   20
      Top             =   4680
      Width           =   2055
      Begin VB.TextBox txtMaxX 
         Height          =   285
         Left            =   1080
         TabIndex        =   22
         Text            =   "0"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtMaxY 
         Height          =   285
         Left            =   1080
         TabIndex        =   21
         Text            =   "0"
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Max X:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   270
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Max Y:"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   630
         Width           =   585
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Map Links"
      Height          =   1455
      Left            =   120
      TabIndex        =   14
      Top             =   480
      Width           =   2055
      Begin VB.TextBox txtUp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   720
         TabIndex        =   18
         Text            =   "0"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtDown 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   720
         TabIndex        =   17
         Text            =   "0"
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtRight 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1320
         TabIndex        =   16
         Text            =   "0"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   120
         TabIndex        =   15
         Text            =   "0"
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblMap 
         BackStyle       =   0  'Transparent
         Caption         =   "Current map: 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame fraMapSettings 
      Caption         =   "Map Settings"
      Height          =   1095
      Left            =   2280
      TabIndex        =   11
      Top             =   360
      Width           =   4215
      Begin VB.ComboBox cmbDayNight 
         Height          =   315
         ItemData        =   "frmMapProperties.frx":0000
         Left            =   960
         List            =   "frmMapProperties.frx":000D
         TabIndex        =   41
         Text            =   "cmbDayNight"
         Top             =   600
         Width           =   3135
      End
      Begin VB.ComboBox cmbMoral 
         Height          =   315
         ItemData        =   "frmMapProperties.frx":0033
         Left            =   960
         List            =   "frmMapProperties.frx":0043
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label4 
         Caption         =   "Day/Night:"
         Height          =   255
         Left            =   30
         TabIndex        =   42
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Moral:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Boot Settings"
      Height          =   2535
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   2055
      Begin VB.Frame Frame8 
         Caption         =   "Requirment"
         Enabled         =   0   'False
         Height          =   975
         Left            =   120
         TabIndex        =   49
         Top             =   1440
         Width           =   1815
         Begin VB.ComboBox cmbVar 
            Height          =   315
            Left            =   120
            TabIndex        =   51
            Text            =   "cmbVariable"
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txtVars 
            Height          =   285
            Left            =   120
            TabIndex        =   50
            Top             =   600
            Width           =   1575
         End
      End
      Begin VB.TextBox txtBootMap 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Text            =   "0"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtBootX 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Text            =   "0"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtBootY 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Text            =   "0"
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Boot Map:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   870
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Boot X:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   645
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Boot Y:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label lblGravity 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gravité : 0g"
      Height          =   195
      Left            =   120
      TabIndex        =   48
      Top             =   5760
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frmEditor_MapProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkDead_Click()
     Map.IsDead = chkDead.Value
End Sub


Private Sub chkTP_Click()

    Map.CanTp = chkTP.Value


End Sub



Private Sub cmdOk_Click()
    Dim I As Long
    Dim sTemp As Long
    Dim X As Long, X2 As Long
    Dim Y As Long, Y2 As Long
    Dim tempArr() As TileRec

    If Not IsNumeric(txtMaxX.text) Then txtMaxX.text = Map.MaxX
    If Val(txtMaxX.text) > MAX_BYTE Then txtMaxX.text = MAX_BYTE
    If Not IsNumeric(txtMaxY.text) Then txtMaxY.text = Map.MaxY
    If Val(txtMaxY.text) > MAX_BYTE Then txtMaxY.text = MAX_BYTE
    
    If Val(txtMaxX.text) < 59 Then
        txtMaxX.text = 59
    End If
    
    If Val(txtMaxY.text) < 35 Then
        txtMaxY.text = 35
    End If
    
    
    With Map
        .name = Trim$(txtName.text)
        .Music = vbNullString
        .BGS = vbNullString

        .Up = Val(txtUp.text)
        .Down = Val(txtDown.text)
        .Left = Val(txtLeft.text)
        .Right = Val(txtRight.text)
        .Moral = cmbMoral.ListIndex
        .BootMap = Val(txtBootMap.text)
        .BootX = Val(txtBootX.text)
        .BootY = Val(txtBootY.text)
        
        .Weather = 0
        .WeatherIntensity = 0
        
        .Fog = ScrlFog.Value
        .FogSpeed = ScrlFogSpeed.Value
        .FogOpacity = scrlFogOpacity.Value
        
        .Red = ScrlR.Value
        .Green = ScrlG.Value
        .Blue = ScrlB.Value
        .Alpha = scrlA.Value
        
        .DayNight = cmbDayNight.ListIndex
        '.Panorama = cmbPanorama.ListIndex
        .CanTp = chkTP.Value
        .IsDead = chkDead.Value
        .Gravity = scrlGravity.Value
        
        ' set the data before changing it
        tempArr = Map.Tile
        X2 = Map.MaxX
        Y2 = Map.MaxY
        ' change the data
        .MaxX = Val(txtMaxX.text)
        .MaxY = Val(txtMaxY.text)
        ReDim Map.Tile(0 To .MaxX, 0 To .MaxY)

        If X2 > .MaxX Then X2 = .MaxX
        If Y2 > .MaxY Then Y2 = .MaxY

        For X = 0 To X2
            For Y = 0 To Y2
                .Tile(X, Y) = tempArr(X, Y)
            Next
        Next
        
        .Variables = cmbVar.ListIndex
        If IsNumeric(txtVars.text) Then
            .VariablesCount = txtVars.text
        Else
            txtVars.text = "0"
            .VariablesCount = txtVars.text
        End If

        'ClearTempTile
    End With
    
    initAutotiles
    Unload frmEditor_MapProperties
    
   
End Sub

Private Sub cmdCancel_Click()
 
    Unload frmEditor_MapProperties

End Sub


Private Sub scrlA_Change()

    lblA.Caption = "Opacity: " & scrlA.Value

End Sub

Private Sub scrlB_Change()

    lblB.Caption = "Blue: " & ScrlB.Value
    

End Sub

Private Sub ScrlFog_Change()

    If ScrlFog.Value = 0 Then
        lblFog.Caption = "None."
    Else
        lblFog.Caption = "Fog: " & ScrlFog.Value
    End If


End Sub

Private Sub scrlFogOpacity_Change()

    lblFogOpacity.Caption = "Fog Opacity: " & scrlFogOpacity.Value
    

End Sub

Private Sub ScrlFogSpeed_Change()

    lblFogSpeed.Caption = "Fog Speed: " & ScrlFogSpeed.Value

End Sub

Private Sub scrlG_Change()

    lblG.Caption = "Green: " & ScrlG.Value
    
End Sub

Private Sub scrlGravity_Change()
    lblGravity.Caption = "Gravity : " & scrlGravity.Value & "g"
End Sub

Private Sub scrlR_Change()

    lblR.Caption = "Red: " & ScrlR.Value
    
End Sub

