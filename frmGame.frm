VERSION 5.00
Begin VB.Form frmGame 
   BorderStyle     =   0  'None
   Caption         =   "SuperGT"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9600
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmGame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      DrawWidth       =   5
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7500
      Left            =   9675
      ScaleHeight     =   496
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   346
      TabIndex        =   0
      Top             =   -7500
      Width           =   5250
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  _____________________________________
' /                                     \
' |           SUPER GT RACING           |
' |     Copyright (c) 2003 Marshall     |
' |-------------------------------------|
' |      Car model from 3DCAFE.COM      |
' \_____________________________________/
'

Private Sub Form_Load()
    'get settings
    ReadConfig
    'setup DirectX
    SetupDX
    'build world
    SetupWorld
    'go to the main loop
    GameLoop
    'reset all buffers
    CleanUp
End Sub

Private Sub GameLoop()
    StartGameTime = DirectX.TickCount
    Delay = 5
    PlaySound EngineBuffer, False, True
    Do
        DelayGame
        If QuitFlag Then Exit Do
        
        EngineBuffer.SetFrequency MoveSpeed * 6000 + 10000
        DoEvents
        CheckKeyboard
        MoveCar
        CheckJoystick
        CheckCamera
        CheckPos
        
        Viewport.Clear D3DRMCLEAR_ALL
        Viewport.Render SceneFrame
        Device.Update 'render scene!
        'UpdateText
     Loop
End Sub

Private Sub CheckKeyboard()
    DInputDevice.GetDeviceStateKeyboard Keyboard
    
    If Keyboard.Key(DIK_LEFT) <> 0 Then
        If RotateSpeed > -0.02 Then RotateSpeed = RotateSpeed - 0.005
    End If
    If Keyboard.Key(DIK_RIGHT) <> 0 Then
        If RotateSpeed < 0.02 Then RotateSpeed = RotateSpeed + 0.005
    End If
        
    If Keyboard.Key(DIK_UP) <> 0 Then If MoveSpeed < 5 Then MoveSpeed = MoveSpeed + 0.05: Braking = False
    If Keyboard.Key(DIK_DOWN) <> 0 Then MoveSpeed = MoveSpeed - 0.04
    
    'Check for camera buttons
    If Keyboard.Key(DIK_INSERT) <> 0 Then CameraView = CHASE
    If Keyboard.Key(DIK_DELETE) <> 0 Then CameraView = INSIDE
    If Keyboard.Key(DIK_HOME) <> 0 Then CameraView = SKY
    If Keyboard.Key(DIK_END) <> 0 Then CameraView = FREE
    
    If Keyboard.Key(DIK_ESCAPE) <> 0 Then CleanUp
End Sub

Private Sub CheckJoystick()
    If UseJoystick = False Then Exit Sub
    Dim Xpos
  
    Call JoyStick_GetPos(JOYSTICKID1)
  
    If JsInfo.dwButtons = 2 Then MoveSpeed = MoveSpeed - 0.04
    If JsInfo.dwButtons = 1 Then If MoveSpeed < 5 Then MoveSpeed = MoveSpeed + 0.05: Braking = False
    Xpos = JsInfo.dwXpos
    Xpos = Xpos - 32768
    Xpos = Xpos / 32768
    Xpos = Round(Xpos, 1)
    
    If Xpos < -0.5 Then If RotateSpeed > -0.02 Then RotateSpeed = RotateSpeed - 0.005
    If Xpos > 0.5 Then If RotateSpeed < 0.02 Then RotateSpeed = RotateSpeed + 0.005
    
End Sub

Private Sub MoveCar()
    Dim CarPos As D3DVECTOR
     
    CarFrame.GetPosition Nothing, CarPos
    
    If CarPos.x > 302 Then CarPos.x = 302
    If CarPos.x < -314 Then CarPos.x = -314
    If CarPos.z > 460 Then CarPos.z = 460
    If CarPos.z < -470 Then CarPos.z = -470
    
    CarFrame.SetPosition Nothing, CarPos.x, 0, CarPos.z
    
    If MoveSpeed >= 0.01 Then MoveSpeed = MoveSpeed - 0.01 'simulate deceleration
    If MoveSpeed < 0.04 Then MoveSpeed = 0   'keep car from rolling backwards
    
    If RotateSpeed > -0.003 And RotateSpeed < 0.003 Then RotateSpeed = 0 'prevent drift
    If RotateSpeed >= 0.003 Then RotateSpeed = RotateSpeed - 0.003
    If RotateSpeed <= -0.003 Then RotateSpeed = RotateSpeed + 0.003

    If MoveSpeed <> 0 Then CarFrame.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, RotateSpeed
    'If MoveSpeed = -0.04 Then MoveSpeed = 0
    
    If Braking Then
        CarFrame.SetPosition CarFrame, 0, 0, -MoveSpeed
    Else
        If picMask.Point(CarPos.x / 2 + 350 / 2, CarPos.z / 2 + 500 / 2) > RGB(1, 1, 1) Then
            CarFrame.SetPosition CarFrame, 0, 0, MoveSpeed
        Else
            CarFrame.SetPosition CarFrame, 0, 0, MoveSpeed / 2
        End If
    End If
End Sub
Private Sub CheckCamera()
    If CameraView = CHASE Then CameraFrame.SetPosition CarFrame, 0, 20, -50
    If CameraView = SKY Then CameraFrame.SetPosition CarFrame, 0, 200, 0
    CameraFrame.LookAt CarFrame, Nothing, D3DRMCONSTRAIN_Z
    If CameraView = INSIDE Then
        CameraFrame.SetPosition CarFrame, 0, 1, -10
        CameraFrame.LookAt CarFrame, Nothing, D3DRMCONSTRAIN_Z
        CameraFrame.SetPosition CarFrame, -1, 7.5, 0
    End If
End Sub

Private Sub CheckPos()
    Dim CarPos As D3DVECTOR
     
    CarFrame.GetPosition Nothing, CarPos
    CheckPointDelay = CheckPointDelay - 1
    
    If CheckPointDelay < 1 Then
        For I = 0 To 3
            If CarPos.x > CheckPointPos(I).x - 22 And CarPos.x < CheckPointPos(I).x + 22 Then
            If CarPos.z > CheckPointPos(I).z - 22 And CarPos.z < CheckPointPos(I).z + 22 Then
                CheckPointVisited(I) = True
                If I = 0 Then 'if this is the finish line
                    For o = 0 To 3
                        If CheckPointVisited(o) = True Then k = k + 1
                    Next o
                    If k = 4 Then 'if all checkpoints have been visited
                                  'if k <> 4 then a checkpoint has been missed
                        PlaySound CheckBuffer, False, False
                        LapsToGo = LapsToGo - 1
                        If LapsToGo = 0 Then CleanUp 'you completed all laps
                        LapsCompleted = LapsCompleted + 1
                        CheckPointDelay = 200
                        For y = 0 To 3
                            CheckPointVisited(y) = False
                        Next y
                    End If
                End If
            End If
            End If
        Next
    End If
End Sub

Private Sub UpdateText()
    ForeColor = RGB(128, 128, 128)
    CurrentX = 25
    CurrentY = 15
    Print "Laps Left: " & LapsToGo
    ForeColor = RGB(255, 255, 255)
    CurrentX = 24
    CurrentY = 14
    Print "Laps Left: " & LapsToGo
    CurrentX = ScreenX - 94
    CurrentY = 15
    ForeColor = RGB(128, 128, 128)
    Print "Time: " & Round((DirectX.TickCount - StartGameTime) / 1000, 0) & " sec"
    CurrentX = ScreenX - 95
    CurrentY = 14
    ForeColor = RGB(255, 255, 255)
    Print "Time: " & Round((DirectX.TickCount - StartGameTime) / 1000, 0) & " sec"
End Sub

Private Sub DelayGame()
    StartTick = DirectX.TickCount
    NowTime = DirectX.TickCount
    Do Until NowTime - LastTick > Delay
        DoEvents
        NowTime = DirectX.TickCount
    Loop
    LastTick = NowTime
End Sub

Private Sub SetupWorld()
    Set CarFrame = D3DRM.CreateFrame(SceneFrame)
    Set CarMesh = D3DRM.CreateMeshBuilder
    CarMesh.LoadFromFile "car.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    CarFrame.AddVisual CarMesh
    CarFrame.SetPosition Nothing, 222, 0, 309 'place car at starting pos
    CarFrame.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 9.42
    CarMesh.ScaleMesh 0.2, 0.2, 0.2           'make car smaller
    CameraView = CHASE     'set camera view
    
    Set GroundFrame = D3DRM.CreateFrame(SceneFrame)
    Set GroundMesh = D3DRM.CreateMeshBuilder
    Call MakeWall(GroundMesh, -350, 0, 500, 350, 0, 500, 350, 0, -500, -350, 0, -500, 1, 1, 0, 0, 0, "ground.bmp")
    GroundFrame.AddVisual GroundMesh
    
    Set FinishFrame = D3DRM.CreateFrame(SceneFrame)
    Set FinishMesh = D3DRM.CreateMeshBuilder
    FinishMesh.LoadFromFile "finish.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    FinishMesh.ScaleMesh 1.2, 1.2, 1.2
    FinishFrame.AddVisual FinishMesh
    FinishFrame.SetPosition Nothing, 203.6, 0, -155
    
    Set MountainFrame = D3DRM.CreateFrame(SceneFrame)
    Set MountainMesh = D3DRM.CreateMeshBuilder
    MountainMesh.LoadFromFile "mountain.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    MountainFrame.AddVisual MountainMesh
    MountainMesh.ScaleMesh 35, 35, 33
    MountainFrame.SetPosition GroundFrame, -100, -20, -50
    Set MountainTexture = D3DRM.LoadTexture("grass.bmp")
    MountainMesh.SetTexture MountainTexture
    
    Set TreeMesh = D3DRM.CreateMeshBuilder
    TreeMesh.LoadFromFile "tree.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    Set TreeTexture = D3DRM.LoadTexture("tree.bmp")
    TreeTexture.SetDecalTransparency D_TRUE
    TreeTexture.SetDecalTransparentColor RGB(255, 255, 255)
    TreeMesh.SetTexture TreeTexture
    TreeMesh.ScaleMesh 1.5, 1.2, 1.5 'make the tree shorter
    
    For I = 0 To 5
        Set TreeFrame(I) = D3DRM.CreateFrame(SceneFrame)
        TreeFrame(I).AddVisual TreeMesh
    Next
    
    'place the trees
    TreeFrame(0).SetPosition Nothing, 1.1, 0, 78
    TreeFrame(1).SetPosition Nothing, -127, 0, 250
    TreeFrame(2).SetPosition Nothing, -132, 0, -113
    TreeFrame(3).SetPosition Nothing, 144, 0, -312
    TreeFrame(4).SetPosition Nothing, 70, 0, 287
    TreeFrame(5).SetPosition Nothing, -184, 0, -374
    
    Picture = LoadPicture("grndmask.bmp")
    picMask.PaintPicture Picture, 0, 0, 350, 500
    Picture = Nothing
    
    CheckPointPos(0).x = 203
    CheckPointPos(0).z = -155
    CheckPointPos(1).x = -110
    CheckPointPos(1).z = -228
    CheckPointPos(2).x = -90
    CheckPointPos(2).z = 170
    CheckPointPos(3).x = 36
    CheckPointPos(3).z = 400
    LapsToGo = TotalLaps
    Caption = "SuperGT"  'if this isn't here, there is no border!
End Sub

Private Sub ReadConfig()
    On Error GoTo NoFile 'if file exists, get setting
                         'else create file with setting
    Open "supergt.cfg" For Input As #1
        Input #1, TmpF
        Input #1, TmpL
        Input #1, TmpX
        Input #1, TmpY
        Input #1, TmpJ
        Input #1, TmpB
        Input #1, TmpR
    Close #1
    If TmpF = "TRUE" Then Fullscreen = True
    If TmpJ = "TRUE" Then UseJoystick = True
    If TmpB = "TRUE" Then UseFiltering = True
    If TmpR = "TRUE" Then UseFlat = True
    TotalLaps = Val(TmpL)
    ScreenX = Val(TmpX)
    ScreenY = Val(TmpY)
    Exit Sub
    
NoFile:
    Close #1
    Open "supergt.cfg" For Output As #1
        Print #1, "FALSE"
        Print #1, "3"
        Print #1, "640"
        Print #1, "480"
        Print #1, "FALSE"
        Print #1, "TRUE"
        Print #1, "FALSE"
    Close #1
    UseFiltering = True
    TotalLaps = 3
    ScreenX = 640
    ScreenY = 480
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'exit the loop
    QuitFlag = True
End Sub

Private Sub CleanUp()
    'clean up our program
    Set CarFrame = Nothing
    Set CarMesh = Nothing
    
    Set FinishFrame = Nothing
    Set FinishMesh = Nothing

    For t = 0 To 5
        Set TreeFrame(t) = Nothing
    Next
    
    Set TreeMesh = Nothing
    Set TreeTexture = Nothing

    Set GroundFrame = Nothing
    Set GroundMesh = Nothing
    Set GroundTexture = Nothing

    Set MountainFrame = Nothing
    Set MountainMesh = Nothing
    Set MountainTexture = Nothing

    Set Light = Nothing
    Set LightFrame = Nothing
    
    ShowTaskbar True
    Unload Me
    End
End Sub
