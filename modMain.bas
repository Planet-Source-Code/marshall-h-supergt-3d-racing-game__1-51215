Attribute VB_Name = "modMain"
'  _________________________________
' /                                 \
' |            modMain.bas          |
' \_________________________________/
'

'DirectX
Public DirectX As New DirectX7
Public CameraFrame As Direct3DRMFrame3
Public SceneFrame As Direct3DRMFrame3
Public DirectDraw As DirectDraw4
Public Clipper As DirectDrawClipper
Public D3DRM As Direct3DRM3
Public Device As Direct3DRMDevice3
Public Viewport As Direct3DRMViewport2

Public StartGameTime As Long, NowTime As Long, Delay As Integer
Public StartTick As Long, LastTick As Long

'Lights
Public Light As Direct3DRMLight
Public LightFrame As Direct3DRMFrame3

'General Declarations
Public QuitFlag As Boolean
Public UseJoystick As Boolean

Public Fullscreen As Boolean
Public ScreenX As Long
Public ScreenY As Long
Public UseFiltering As Boolean
Public UseFlat As Boolean

Public CameraView As Integer
Public Const CHASE = 0
Public Const INSIDE = 1
Public Const SKY = 2
Public Const FREE = 3

Public MoveSpeed
Public Braking As Boolean
Public RotateSpeed


'API Declarations for game timer
'Public Declare Function GetTickCount Lib "kernel32" () As Long
'Public tTime As Long
'Public Const ms_Delay = 1

'Functions to hide taskbar
Declare Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Const SWP_HIDEWINDOW = &H80
Const SWP_SHOWWINDOW = &H40

'Checkpoints
Public CheckPointPos(3) As D3DVECTOR
Public CheckPointDelay
Public CheckPointVisited(3) As Boolean

Public LapsCompleted As Integer
Public LapsToGo As Integer
Public TotalLaps As Integer

'Objects
Public CarFrame As Direct3DRMFrame3
Public CarMesh As Direct3DRMMeshBuilder3

Public FinishFrame As Direct3DRMFrame3
Public FinishMesh As Direct3DRMMeshBuilder3

Public TreeFrame(5) As Direct3DRMFrame3
Public TreeMesh As Direct3DRMMeshBuilder3
Public TreeTexture As Direct3DRMTexture3

Public GroundFrame As Direct3DRMFrame3
Public GroundMesh As Direct3DRMMeshBuilder3
Public GroundTexture As Direct3DRMTexture3

Public MountainFrame As Direct3DRMFrame3
Public MountainMesh As Direct3DRMMeshBuilder3
Public MountainTexture As Direct3DRMTexture3

Public Sub SetupDX()
    ChDir App.Path
    
    Set DirectDraw = DirectX.DirectDraw4Create("")
    Set Clipper = DirectDraw.CreateClipper(0)
    If Not Fullscreen Then frmGame.BorderStyle = 1
    frmGame.Width = ScreenX * Screen.TwipsPerPixelX
    frmGame.Height = ScreenY * Screen.TwipsPerPixelY
    If Fullscreen Then ShowTaskbar False
    If Fullscreen Then DirectDraw.SetDisplayMode ScreenX, ScreenY, 16, 0, DDSDM_DEFAULT
    Clipper.SetHWnd frmGame.Hwnd
    
    ' init d3drm and main frames
    Set D3DRM = DirectX.Direct3DRMCreate()
    Set SceneFrame = D3DRM.CreateFrame(Nothing)
    Set CameraFrame = D3DRM.CreateFrame(SceneFrame)
    
    SceneFrame.SetSceneBackgroundRGB 0.5, 0.9, 1     'make background blue
           
    'Set device = D3DRM.CreateDeviceFromClipper(clipper, "IID_IDirect3DRGBDevice", frmGame.ScaleWidth, frmGame.ScaleHeight)
    Set Device = D3DRM.CreateDeviceFromClipper(Clipper, "IID_IDirect3DHALDevice", ScreenX, ScreenY)
    
    Device.SetQuality D3DRMFILL_SOLID + D3DRMLIGHT_ON + D3DRMSHADE_GOURAUD
    Device.SetDither D_TRUE
    Device.SetShades 1

    Set Viewport = D3DRM.CreateViewport(Device, CameraFrame, 0, 0, ScreenX, ScreenY)
    Viewport.SetBack 900
                  
    If UseFiltering Then Device.SetTextureQuality D3DRMTEXTURE_LINEAR  'smooth out textures
    If UseFlat Then Device.SetQuality D3DRMRENDER_FLAT
    
    'do the rest
    SetupDInput
    If UseJoystick Then Call JoyStick_Init(JOYSTICKID1)
    SetupDSound
    SetupLights
    frmGame.Show

End Sub

Private Sub SetupLights()
    Set Light = D3DRM.CreateLightRGB(D3DRMLIGHT_DIRECTIONAL, 1, 1, 1)
    Set LightFrame = D3DRM.CreateFrame(SceneFrame)
    LightFrame.SetPosition Nothing, 0, 0, 0
    Light.SetRange 1000!
    LightFrame.AddLight Light
    
    SceneFrame.AddLight D3DRM.CreateLightRGB(D3DRMLIGHT_AMBIENT, 1, 1, 1)
End Sub

Public Sub MakeWall(mesh As Direct3DRMMeshBuilder3, X1 As Single, Y1 As Single, z1 As Single, x2 As Single, y2 As Single, z2 As Single, x3 As Single, y3 As Single, z3 As Single, x4 As Single, y4 As Single, z4 As Single, TileX As Single, TileY As Single, r As Single, g As Single, b As Single, texfile As String)
    ' local variables
    Dim Face As Direct3DRMFace2
    ' create face
    Dim texture As Direct3DRMTexture3
    Set Face = D3DRM.CreateFace()
    ' add vertexs
    Face.AddVertex X1, Y1, z1
    Face.AddVertex x2, y2, z2
    Face.AddVertex x3, y3, z3
    Face.AddVertex x4, y4, z4
    Face.AddVertex x4, y4, z4
    Face.AddVertex x3, y3, z3
    Face.AddVertex x2, y2, z2
    Face.AddVertex X1, Y1, z1
    
    ' get type of Faceile
    If texfile = "" Then
    '    ' set colors
        Face.SetColorRGB r, g, b
    Else
        ' create textuere
        Set tex = D3DRM.LoadTexture(texfile)
        ' set u and v values
        Face.SetTextureCoordinates 0, 0, TileY
        Face.SetTextureCoordinates 1, 0, 0
        Face.SetTextureCoordinates 2, TileX, 0
        Face.SetTextureCoordinates 3, TileX, TileY
        Face.SetTextureCoordinates 4, TileX, TileY
        Face.SetTextureCoordinates 5, TileX, 0
        Face.SetTextureCoordinates 6, 0, 0
        Face.SetTextureCoordinates 7, 0, TileY
        ' set the texture
         Face.SetTexture tex
    End If
    ' add face to mesh
    mesh.AddFace Face
End Sub
Public Sub ShowTaskbar(bShow As Boolean)
    Dim Thwnd As Long
    Thwnd = FindWindow("Shell_traywnd", "")
    If bShow Then
        Call SetWindowPos(Thwnd, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
    Else
        Call SetWindowPos(Thwnd, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
    End If
End Sub
