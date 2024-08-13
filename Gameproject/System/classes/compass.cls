VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "compass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''ÈÓã Çááå ÇáÑÍãä ÇáÑÍíã
''''''''''''''''''''''''''''''''
'===================================== Compass
'##############################
'#Written by : jaber al_ani   #
'#Date       : 28- 6 -2006    #
'#Time       : 1 :29 PM       #
'##############################

Private compass             As D3DXSprite
Private compasstex          As Direct3DTexture8

Private compass2            As D3DXSprite
Private compasstex2         As Direct3DTexture8
Public Function create()
'###############################################compass
Set compass = D3DX.CreateSprite(D3DDevice)
Set compasstex = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\Data\texture\compass\Compass1.bmp", 64, 64, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_DEFAULT, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, &HFF000000, ByVal 0, ByVal 0)

Set compass2 = D3DX.CreateSprite(D3DDevice)
Set compasstex2 = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\Data\texture\compass\Compass2.bmp", 64, 64, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_DEFAULT, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, &HFF000000, ByVal 0, ByVal 0)
'###############################################
End Function

Public Function render()
'########################################################compass
If D3DPP.BackBufferWidth = 1024 Then compass.Draw compasstex, ByVal 0, MV(1, 1), MV(64, 64), 0, MV(950, 20), &HFFFFFFFF   '&H5FFF0000 / 1.5
If D3DPP.BackBufferWidth = 800 Then compass.Draw compasstex, ByVal 0, MV(0.8, 0.8), MV(64, 64), 0, MV(750, 10), &HFFFFFFFF  '&H5FFF0000 / 1.5
'''''' for compass2
If D3DPP.BackBufferWidth = 1024 Then compass2.Draw compasstex2, ByVal 0, MV(1, 1), MV(64 / 2, 64 / 2), realangel, MV(950, 20), &HFFFFFFFF    '&H5FFF0000 / 1.5
If D3DPP.BackBufferWidth = 800 Then compass2.Draw compasstex2, ByVal 0, MV(1, 1), MV(64 / 2, 64 / 2), realangel, MV(745, 3), &HFFFFFFFF   '&H5FFF0000 / 1.5
'########################################################
End Function
