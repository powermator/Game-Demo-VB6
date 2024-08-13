VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "gameengine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private tanangel            As Single
Private tanangel2           As Single

'------------------------- health
Private health            As D3DXSprite
Private healthpic          As Direct3DTexture8

Private healthbar         As D3DXSprite
Private healthbarpic          As Direct3DTexture8
'-----------------------------------------------------------

Public Function getangel()
On Error Resume Next

tanangel = Abs(EyeDir.X) / Abs(EyeDir.Z)           ' find the tangent
realangel = Atn(tanangel) * 180 / 3.14285714285714 'find the angel in degree

' specify region to find the base angel
If EyeDir.Z + EyePos.Z > 1300 And EyeDir.X + EyePos.X < 3600 Then realangel = realangel
If EyeDir.Z + EyePos.Z < 1300 And EyeDir.X + EyePos.X < 3600 Then realangel = 180 - realangel
If EyeDir.Z + EyePos.Z < 1300 And EyeDir.X + EyePos.X > 3600 Then realangel = 180 + realangel
If EyeDir.Z + EyePos.Z > 1300 And EyeDir.X + EyePos.X > 3600 Then realangel = 360 - realangel

realangel = realangel * 3.14285714285714 / 180     'find the angel in radian

'=========================== for angel 2
tanangel2 = Abs(EyeDir.Y) ' find the tangent
realangel2 = Atn(tanangel2) * 180 / 3.14285714285714 'find the angel in degree

' specify region to find the base angel
If EyeDir.Y + EyePos.Y > 140 Then realangel2 = realangel2
If EyeDir.Y + EyePos.Y < 140 Then realangel2 = -realangel2

realangel2 = realangel2 * 3.14285714285714 / 180     'find the angel in radian

End Function
Private Sub Class_Initialize()

Set health = D3DX.CreateSprite(D3DDevice)
Set healthbar = D3DX.CreateSprite(D3DDevice)

Set healthpic = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\Data\texture\face\health.bmp", 32, 413, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_DEFAULT, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, &HFF000000, ByVal 0, ByVal 0)
Set healthbarpic = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\Data\texture\face\healthbar.bmp", 413, 32, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_DEFAULT, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, &HFF000000, ByVal 0, ByVal 0)

End Sub

Public Sub renderhealth()

If D3DPP.BackBufferWidth = 1024 Then health.Draw healthpic, ByVal 0, MV(1, 1), MV(0, 0), 0, MV(0, 240), &HFFFFFFFF

End Sub
