VERSION 5.00
Begin VB.Form form3 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10065
   ClientLeft      =   1905
   ClientTop       =   2175
   ClientWidth     =   12180
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "main.frx":1CCA
   MousePointer    =   99  'Custom
   ScaleHeight     =   671
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   812
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer4 
      Interval        =   1000
      Left            =   600
      Top             =   4200
   End
   Begin VB.Timer Timer3 
      Interval        =   600
      Left            =   600
      Top             =   960
   End
   Begin VB.Timer Timer2 
      Interval        =   6000
      Left            =   960
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Interval        =   6000
      Left            =   360
      Top             =   360
   End
End
Attribute VB_Name = "form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=================================ÈÓã Çááå ÇáÑÍãä ÇáÑÍíã
'=================================
'##############################
'#Written by : jaber al_ani   #
'#                            #
'#     jaberani@yahoo.com     #
'##############################
'==================================
'===================================================================
'feel free to change or edit code or whatever U want :-)           =
'======================================================25/8/2006====
'=Press N for night vision


Public Sub Form_Load()
'=======================================================Ýí ÇáÈÏÇíå äåÏã ÇáÝæÑãíä
Unload Form1
Unload Form2
'====================================================== ÇÚÏÇÏÇÊ
setupdirectx
SetupMatrices
'=====================================================ÇáÊÍßã
startshow = 1
faceshow = 0
startgame = 0
roundnum = 1
'=====================================================ÇáãæÓíÞå
Set mouseoversound = New MusicSOUND
mouseoversound.create "mouseon"

Set moseclick = New MusicSOUND
moseclick.create "click"

Set shotsound = New MusicSOUND
shotsound.create "shot"

Set gunmovesound = New MusicSOUND
gunmovesound.create "move"

Set facemusic = New MusicMP3
facemusic.create "1"
'======================================================ÇáÚÑÖ
Set startlogo = New logo
'======================================================ÇáæÇÌåå
Set gameface = New face
'======================================================ÏÇíÑßÊ ÇäÈÊ
Set inp = New Directinput
'=====================================================ÇááÚÈå
Set mission = New mission
'======================================================ÇáãÑßÈå
Set helicopter = New heli

Set ships = New ship
'========================================================ßÊÇÈå ãÄÞÊå
Set ttext = New text
ttext.create 20, vbWhite, "kkk", 1
'========================================================ÇáÈæÕáå
Set comp = New compass
comp.create
'========================================================ÇáÓáÇÍ
Set gun = New weapon
gun.create "1"
'======================================================== ÇáãÍÑß
Set game = New gameengine
'=======================================================ÇáÇäÝÌÇÑ
Set explosion = New explo
'========================================================
''''''''''''''''''''
'''''''''''''''''''
''''''''''''''''''''
Do
DoEvents
fps
moving = 0
game.getangel

getcurpos                 'äÇÎÐ ÇÍÏÇËíÇÊ ÇáãÔíÑå
inp.CheckKey                      'äÇÎÐ ÇáãÏÎáÇÊ
directview                'ÊæÌíå ÇáäÙÑ ááßÇãÑå
RenderAll                            'ÑäÏÑ ááß

Loop

Timer5.Enabled = True

End Sub
Private Sub setupdirectx()

Set D3D = DX.Direct3DCreate()
D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode


Get #1, 1, temptext
If Val(temptext) = Val("1024 X 768") Then D3DPP.BackBufferHeight = 768: D3DPP.BackBufferWidth = 1024
If Val(temptext) = Val("800  X 600") Then D3DPP.BackBufferHeight = 600: D3DPP.BackBufferWidth = 800
     
temptext = ""
Get #1, 2, temptext
If Val(temptext) = Val("16") Then D3DPP.AutoDepthStencilFormat = D3DFMT_D16
If Val(temptext) = Val("24") Then D3DPP.AutoDepthStencilFormat = D3DFMT_D24X8
If Val(temptext) = Val("32") Then D3DPP.AutoDepthStencilFormat = D3DFMT_D32

    With D3DPP
        .Windowed = 0
        .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
        .BackBufferFormat = DispMode.Format
        .hDeviceWindow = form3.hWnd
        .BackBufferCount = 1
        .EnableAutoDepthStencil = 1
    End With
   
Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, form3.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DPP)
       
D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CW
D3DDevice.SetRenderState D3DRS_ZENABLE, 1
    
        D3DDevice.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
        D3DDevice.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
        D3DDevice.SetTextureStageState 0, D3DTSS_COLORARG2, D3DTA_DIFFUSE
        D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_DISABLE
        
temptext = ""
Get #1, 3, temptext

If temptext = "ãÊæÓØå" Then
D3DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
D3DDevice.SetTextureStageState 0, D3DTSS_MIPFILTER, D3DTEXF_LINEAR
End If

If temptext = "ãäÎÝÙå" Then
D3DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_POINT
D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_POINT
D3DDevice.SetTextureStageState 0, D3DTSS_MIPFILTER, D3DTEXF_POINT
End If

If temptext = "ÚÇáíÜå" Then
D3DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC
D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
D3DDevice.SetTextureStageState 0, D3DTSS_MIPFILTER, D3DTEXF_ANISOTROPIC
End If

Close #1

End Sub
Private Sub SetupMatrices()

EyePos = MakeVector(3600, 140, 1300)
EyeDir = MakeVector(-2.1, 0, 0)

D3DXMatrixPerspectiveFovLH matProj, pi / 2, 1, 1, 1000000
D3DDevice.SetTransform D3DTS_PROJECTION, matProj
    
D3DXMATH_MATRIX.D3DXMatrixIdentity matworld
D3DXMATH_MATRIX.D3DXMatrixIdentity tempmat
End Sub

Private Sub directview()
D3DXMatrixLookAtLH matView, MakeVector(EyePos.X, EyePos.Y, EyePos.Z), MakeVector(EyePos.X + EyeDir.X, EyePos.Y + EyeDir.Y, EyePos.Z + EyeDir.Z), MakeVector(0, 1, 0)
D3DDevice.SetTransform D3DTS_VIEW, matView
End Sub

Private Sub RenderAll()
D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, vbBlack, 1#, 0
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
D3DDevice.SetVertexShader D3DFVF_VERTEX
D3DDevice.SetTransform D3DTS_WORLD, matworld

   
D3DDevice.BeginScene  '***************&&&&&************

If startgame = 1 Then ships.render: mission.render:  gun.render: helicopter.render: comp.render: ships.renderface: gun.renderface: game.renderhealth: SetCursorPos 100, 100 ': ShowCursor 0
ttext.rendertwo Str(10 - bullnum), 60, 5, 400, 50, &HFFFF00FF, 40
'ttext.rendertwo Str(FrameRate), 60, 5, 400, 50, &HFFFF00FF, 40
explosion.render
D3DDevice.EndScene    '***************&&&&&************

If faceshow = 1 Then gameface.render
If startshow = 1 Then startlogo.render
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
End Sub
Private Sub getcurpos()
GetCursorPos curpos
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If startgame = 1 And time >= 3 Then
If enablegun = 1 Then
If gunmode = 0 And bullnum <= 11 Then gun.bullshot: shotsound.PLAYnow
If gunmode = 1 And lazershot = 0 Then lazershot = 1: gun.lazershott
End If
End If

End Sub
Private Sub Timer1_Timer()
yesrot = 1                        ' æÞÊ ÇáãÍÏÏ áÝÑ Çááæßæ
End Sub
Private Sub Timer2_Timer()
tempval = -0.01                ' ÇáÞíãå ÇáÊí ÊÚãá Óßíá ááæßæ
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If startgame = 1 And time >= 3 Then
    RotateCamera (100 - X) * -3.14159275180032 / 180 / 10, (100 - Y) * -3.14159275180032 / 180 / 10
    If X <> 100 Or Y <> 100 Then
        SetCursorPos 100, 100
    End If
    End If
End Sub
Private Sub Timer3_Timer()
time = time + 1
If time > 220 Then Timer3.Enabled = False

If time > 100 Then sun = 0: rise = 1
If time > 200 Then sun = 0: rise = 0: night = 1: Timer3.Enabled = False
End Sub
Private Sub fps()
''''''''''''''''''''fps
LastUpdated = GetTickCount()
    FrameCount = FrameCount + 1
    If (GetTickCount() - FrameRateLastCheck >= 1000) Then
        'one second has passed
        FrameRate = FrameCount
        FrameCount = 0
        FrameRateLastCheck = GetTickCount()
        Debug.Print FrameRate & "fps"
    End If
'''''''''''''''''''
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'============ ááÑÄíå Ãááíáíå
If KeyCode = vbKeyN Then
nkey = nkey + 1
If nkey = 1 Then scensematerial = setupmaterials(0, 10, 0): enablefog = 0: D3DDevice.SetRenderState D3DRS_FOGENABLE, 0
If nkey = 2 Then nkey = 0: scensematerial = currentmaterial: enablefog = 1
End If
'==============ááÓáÇÍ
If KeyCode = vbKey1 Then
gun.create "1"
End If

If KeyCode = vbKey2 Then
gun.create "2"
End If

End Sub
'==================================
'==================================
Private Sub Timer4_Timer()
smallshipattackflag = 1
End Sub

