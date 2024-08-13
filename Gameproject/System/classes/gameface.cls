VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "logo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''ÈÓã Çááå ÇáÑÍãä ÇáÑÍíã
''''''''''''''''''''''''''''''''
'===================================== Logo
'##############################
'#Written by : jaber al_ani   #
'#Date       :                #
'#Time       :                #
'##############################

'mashreq
Private mashreqmesh        As D3DXSprite
Private mashreqtex         As Direct3DTexture8
Private cdscale            As Single
Private rotation           As Single

'raed CD
Private cdmesh             As D3DXSprite
Private cdtex              As Direct3DTexture8
Private scaleval           As Single


Private clock              As Timer
Private clock2             As Timer
Private playsound          As Byte

Private Sub Class_Initialize()
Set clock = form3.Timer1
Set clock2 = form3.Timer2
clock.Enabled = True
clock2.Enabled = False

Set mashreqmesh = D3DX.CreateSprite(D3DDevice)
Set mashreqtex = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\Data\texture\face\raed.jpg", 800, 300, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_DEFAULT, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, vbRed, ByVal 0, ByVal 0)


Set cdmesh = D3DX.CreateSprite(D3DDevice)
Set cdtex = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\Data\texture\face\cd.jpg", 256, 256, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_DEFAULT, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, &HFF000000, ByVal 0, ByVal 0)

scaleval = 0.1
tempval = 0.01
If D3DPP.BackBufferWidth = 1024 Then cdscale = 2
If D3DPP.BackBufferWidth = 800 Then cdscale = 1.3

End Sub

Public Function render()
D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, vbBlack, 1#, 0
'###################################Mashreq
If yesrot = 1 Then
cdscale = cdscale - 0.01
If cdscale < 0 Then cdscale = 0
rotation = rotation + 0.005
If rotation > 1 Then rotation = 0
End If
If D3DPP.BackBufferWidth = 1024 Then cdmesh.Draw cdtex, ByVal 0, MV(cdscale, cdscale), MV(256, 256), -Cos(rotation) * 9.5 * 2.98, MV(270, 300 / 4), &HFFFFFFFF
If D3DPP.BackBufferWidth = 800 Then cdmesh.Draw cdtex, ByVal 0, MV(cdscale, cdscale), MV(256 / 1.5, 256 / 1.5), -Cos(rotation) * 9.5 * 2.98, MV(230, 80), &HFFFFFFFF
'####################################Raed
If cdscale = 0 Then   'ÚäÏãÇ íßæä ÍÌã ÇáÞÑÕ íÓÇæí ÕÝÑ

'äÔÛá ãæÓíÞå
If playsound = 0 Then Set facemusic = Nothing: Set facemusic = New MusicMP3: facemusic.create "2": facemusic.switchsound 1: playsound = 1

scaleval = scaleval + tempval
If scaleval > 1 Then scaleval = 1: clock2.Enabled = True 'ÚäÏãÇ Êßæä ÇáßÊÇÈå ÇßÈÑ ãÇ íãßä äÞáÈ Þíãå ÇáÓßíá
If scaleval < 0 Then startshow = 0: Set startlogo = Nothing: faceshow = 1 'ÚäÏãÇ ÊÕÈÍ Þíãå ÇáÓßíá ÕÝÑ äåÏã ÇáÇæÈÌßÊ

If D3DPP.BackBufferWidth = 1024 Then mashreqmesh.Draw mashreqtex, ByVal 0, MV(scaleval, scaleval), MV(800, 300), 0, MV(8, 300 / 4), &HFFFFFFFF
If D3DPP.BackBufferWidth = 800 Then mashreqmesh.Draw mashreqtex, ByVal 0, MV(scaleval / 1.5, scaleval / 1.5), MV(800, 300), 0, MV(70, 300 / 4), &HFFFFFFFF
End If
'######################################

End Function
