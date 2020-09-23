VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   7290
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'JohnaDX7 engine version 1.04
'all johnaDX7 classes was developed by Johna
'
'if someone want to use them or modify them
' please email me before and send me the new version
'
'   for explanation or remarks
'    johna.pop@caramail.com
'
'   VOTE FOR IT IF YOU LIKE THIS
'
'
'With that engine there are
'  -Sound Engine for unlimited buffers in 3d world
'  -Particle engine for generating many effects
'  -SkyBox
'  -DomeSky
'  -Great OOBB colision detection
'  -Xfile loader with great colision detection multi BoundingBox
'
' ect....
'=================================================




'main Engine Class
Dim DX7 As New johna_DX7

'for landscape
Dim Land As New cJohna_landScape
Dim Land_Map As DirectDrawSurface7

'for the sky
Dim Sky As New cJohna_Sphere


'sound system
Dim SOUND As New cJohna_Sound
Dim PLAYER_STEP As Integer

Private Sub Form_Load()
 Me.Show
 Me.Refresh
 
 'DX7.INIT_ShowDialog Me.hwnd
 DX7.Initialize_Windowed Me.hwnd
 Call CreateOBJECT
 Call Game_LOOP
End Sub


Sub FreeALL()
DX7.FreeDX Me.hwnd
SOUND.Engine_Free
End
End Sub






Sub CreateOBJECT()
'sound
SOUND.INIT_SoundENGINE Me.hwnd
SOUND.DMusic_OpenMidi App.path + "\data\Z64gerud.mid"
'play with loop status
SOUND.DMusic_PlayMidi True
'Load a wav sample
PLAYER_STEP = SOUND.Load_Wav(App.path + "\data\GrassyStep3.wav", True)



'making the land
Land.INITIALIZE DX7
Land.LoadTerrain App.path + "\data\big_valley.jpg", Johna_NORMAL, 1, 10, 10, App.path + "\data\rock1_goop_base.gif"
Set Land_Map = DX7.CreateTextureEX(App.path + "\data\rock1_goop_base.JPG", 256, 256)

'set the land scale
Land.Set_Scale 20, 8, 20
Land.Set_Position -500, 0, -500
'load water texture and set water scale
Land.Load_Watertexture App.path + "\data\water.bmp"
Land.Water_scal = Vector(40, 10, 40)
Land.Set_Position -500, 0, -500
Land.SetWaterEnable 1
Land.SetWaterEffect 1, , 1
Land.SetWaterTextureScale 50
Land.SetWaterAltitude 10


'making the sky
Sky.InitSPhere DX7, App.path + "\data\sky.bmp", NORMAL_QUALITY, 400, 400, 400
Sky.Set_Scale 40, 40, 40

'respecify viewfrustrum for maximum visibility
DX7.Set_ViewFrustrum 1, 50000

'set fog level
DX7.SetFog True, 1 / 5000, &HB0E0FF, D3DFOG_EXP2

End Sub










'==========================
'
'Main game loop
'
'========================
Sub Game_LOOP()

Dim ANG

Do
  
  DoEvents
  If DX7.GetKEY(Johna_KEY_ESCAPE) Then GoTo END_it
  'checkeys
  Call Me.KEY_check
  ANG = ANG + 0.05
  If ANG > 360 Then ANG = 0
  'DX7.Camera_SetXRotation uu
  
  DX7.Clear_3D
  DX7.D3D_DEV.BeginScene
   
    'draw the land
    Call RenderLand
 
  DX7.D3D_DEV.EndScene
  
  
  
  
  
  
  
  'print useful info
  Call PrintInfo(True)
  
  DX7.FLIPP Me.hwnd
Loop

END_it:

Call FreeALL
End Sub









Sub RenderLand()
 Dim VC As D3DVECTOR, Y
 
 'check player altitude
 VC = DX7.GET_CameraEYE
 
 Y = Land.Get_Altitude_EX(VC)
 VC.Y = Y + 50
 DX7.SET_cameraEYE VC


'uncomment the 2 next commented lines for multitexturing
'DX7.MULTITEXTURE_FX_Dark_Mapping 1, Land_Map
Land.RenderAll 0
'DX7.MULTITEXTURING_OFF

Land.Render_Water

'render sky
Sky.Render DX7
End Sub





Sub KEY_check()

'update player ears
SOUND.Set_ListenerPosition DX7.GET_CameraEYE
SOUND.Set_BufferPosition PLAYER_STEP, DX7.GET_CameraEYE



'Move forward
If DX7.GetKEY(Johna_KEY_UP) Then
  DX7.Camera_Move_Foward 4
  'play stepsound
  SOUND.Play_BUF PLAYER_STEP
End If
  

If DX7.GetKEY(Johna_KEY_RCONTROL) Then
  DX7.Camera_Move_Foward 12 / 2
 'play stepsound
  SOUND.Play_BUF PLAYER_STEP
End If

  
'Move backWard
If DX7.GetKEY(Johna_KEY_DOWN) Then
  DX7.Camera_Move_Backward 4 / 2
 'play stepsound
  SOUND.Play_BUF PLAYER_STEP
End If


'Turn Left
If DX7.GetKEY(Johna_KEY_LEFT) Then _
  DX7.Camera_Turn_Left 2 / 100

'Turn right
If DX7.GetKEY(Johna_KEY_RIGHT) Then _
  DX7.Camera_Turn_Right 2 / 100
  


If DX7.GetKEY(Johna_KEY_NUMPAD8) Then _
  DX7.Camera_Turn_UP 0.2

If DX7.GetKEY(Johna_KEY_NUMPAD2) Then _
  DX7.Camera_Turn_UP -0.2
  
  
If DX7.GetKEY(Johna_KEY_NUMPAD7) Then _
  DX7.Camera_Roll_Left 0.5

If DX7.GetKEY(Johna_KEY_NUMPAD9) Then _
  DX7.Camera_Roll_Right 0.5


If DX7.GetKEY(Johna_KEY_ADD) Then _
  DX7.Camera_Elevator_UP 5
  
If DX7.GetKEY(Johna_KEY_SUBTRACT) Then _
  DX7.Camera_Elevator_DOWN 5

End Sub









Sub PrintInfo(Visible As Boolean)
DX7.BackBuffer.DrawText 10, 50, "FPS=" + Str(DX7.FramesPerSec), False


End Sub
