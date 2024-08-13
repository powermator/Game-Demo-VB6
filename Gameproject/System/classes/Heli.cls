VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Heli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''ÈÓã Çááå ÇáÑÍãä ÇáÑÍíã
''''''''''''''''''''''''''''''''
'===================================== DirectInput8
'##############################
'#Written by : jaber al_ani   #
'#Date       : 10-Jan-2006    #
'#Time       : 10.48 PM       #
'##############################


'======================= friend ship
Private helimesh             As D3DXMesh
Private helimatrix           As D3DMATRIX
Private helixpos             As Integer
Private heliypos             As Integer
Private helizpos             As Integer
Private landing              As Byte

Private texture          As Direct3DTexture8
Private texture2         As Direct3DTexture8
Private texture3         As Direct3DTexture8

Private helisound       As MusicSOUND


Private heli             As D3DXSprite
Private helipic          As Direct3DTexture8
'============================================

Private Sub Class_Initialize()

Set helisound = New MusicSOUND
helisound.create "helicopter"

Set heli = D3DX.CreateSprite(D3DDevice)
Set helipic = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\Data\texture\ships\heli.bmp", 150, 60, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_DEFAULT, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, &HFF000000, ByVal 0, ByVal 0)

'{{{{{{{{{{{{{{{{{{{{ heli
'{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{
Set texture = D3DX.CreateTextureFromFile(D3DDevice, App.Path + "\Data\texture\ships\y.jpg")
Set texture2 = D3DX.CreateTextureFromFile(D3DDevice, App.Path + "\Data\texture\ships\x.jpg")
Set texture3 = D3DX.CreateTextureFromFile(D3DDevice, App.Path + "\Data\texture\ships\z.jpg")

Set helimesh = D3DX.LoadMeshFromX(App.Path & "\Data\xfiles\ship\heli.x", D3DXMESH_MANAGED, D3DDevice, Nothing, Nothing, 0)
D3DXMATH_MATRIX.D3DXMatrixIdentity helimatrix
helixpos = 3550
heliypos = 160
helizpos = 1500

D3DXMATH_MATRIX.D3DXMatrixScaling tempmat, 3, 3, 3
D3DXMATH_MATRIX.D3DXMatrixMultiply helimatrix, helimatrix, tempmat

With helimatrix
.m41 = helixpos
.m42 = heliypos
.m43 = helizpos
End With
'}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}
'}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}

End Sub
Public Function render()

If D3DPP.BackBufferWidth = 1024 Then heli.Draw helipic, ByVal 0, MV(0.5, 0.5), MV(0, 0), 0, MV(0, 100), &HFFFFFFFF       '&H5FFF0000 / 1.5
'{{{{{{{{{{{{{{{{{{{{ Heli{{{{{{{{{{{{{{{{
'{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{

D3DDevice.SetMaterial scensematerial

If helicomeflag = 2 Then
helisound.PLAY
If landing = 0 And helizpos > 1300 Then helizpos = helizpos - 1
'áßí ÊÊæÞÝ
If landing = 0 And helizpos < 1300 Then helizpos = 1300
'áßí ÊåÈØ
If landing = 0 And helizpos = 1300 Then heliypos = heliypos - 1
If landing = 0 And heliypos <= 135 Then heliypos = 135: landing = 1: enablegun = 1: gunmode = 1: bullnum = bullnum - 10: lazershot = 0


If landing = 1 Then heliypos = heliypos + 1
If landing = 1 And heliypos >= 160 Then heliypos = 160
If landing = 1 And heliypos = 160 Then helizpos = helizpos - 1

With helimatrix
.m41 = helixpos
.m42 = heliypos
.m43 = helizpos
End With

D3DDevice.SetTransform D3DTS_WORLD, helimatrix
For i = 0 To 8
D3DDevice.SetTexture 0, texture
If i = 3 Or i = 4 Then D3DDevice.SetTexture 0, texture2
If i = 8 Or i = 6 Then D3DDevice.SetTexture 0, texture3
helimesh.DrawSubset i
Next i

'ÇáÇÚÇÏå æ ÇáÊÕÝíÑ
If landing = 1 And helizpos < 600 Then
helixpos = 3550
heliypos = 160
helizpos = 1500
landing = 0
helicomeflag = 0
End If

End If
'}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}
'}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}
End Function
