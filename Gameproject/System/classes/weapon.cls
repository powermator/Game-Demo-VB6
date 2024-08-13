VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "weapon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''ÈÓã Çááå ÇáÑÍãä ÇáÑÍíã
''''''''''''''''''''''''''''''''
'===================================== weapon
'##############################
'#Written by : jaber al_ani   #
'#Date       : 28- 6 -2006    #
'#Time       :       PM       #
'##############################
'=====================================

'----------------------------------------------------- ÇáÓáÇÍ
Private mesh           As D3DXMesh
Private matrix         As D3DMATRIX
Private texture        As Direct3DTexture8

Private xpos           As Integer
Private ypos           As Integer
Private zpos           As Integer

Private tempmatrix     As D3DMATRIX
Private tempmatrix2    As D3DMATRIX

Private lasttempangel  As Single
Private lasttempangel2 As Single

Private weaponname     As Byte
'----------------------------------------------------ÇáÇØáÇÞÇÊ
Private bullmesh           As D3DXMesh
Private bullmatrix(10)     As D3DMATRIX
Private bullpos(10)        As D3DVECTOR
Private bulletangel(10)    As Single
Private bulletangel2(10)   As Single
Private living(10)         As Byte
Private speed              As Byte
Private bullmaterial       As D3DMATERIAL8
'--------------------------------------------------- ÇááíÒÑ
Private lazermesh         As D3DXMesh
Private lazermatrix       As D3DMATRIX
Private lazerscale        As Single
Private lazersound        As MusicSOUND
'-------------------------------------------------ÇáäíÔÇä
Private nesh             As D3DXSprite
Private neshtex1         As Direct3DTexture8
Private neshtex2         As Direct3DTexture8
Private neshtex3         As Direct3DTexture8
'-------------------------------------------------ÕæÑå ÇáØáÞå
Private bull             As D3DXSprite
Private bulltex         As Direct3DTexture8
'---------------------------------------------concret
Private concmat        As D3DMATRIX
Private concmesh       As D3DXMesh
Private conctex        As Direct3DTexture8
'----------------------------------------------

Private Sub Class_Initialize()
  '0= it mean bulet gun '1= it mean laser gun
enablegun = 1
gunmode = 1
End Sub
Public Function create(name As String)
weaponname = Val(name)
'################################# ÇáäíÔÇä
Set nesh = D3DX.CreateSprite(D3DDevice)
Set neshtex1 = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\Data\cursur\" + Format(name) + ".bmp", 56, 56, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_DEFAULT, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, &HFF000000, ByVal 0, ByVal 0)
'################################

'################################  ÇáÓáÇÍ
xpos = 3600
ypos = 125
zpos = 1300
Set mesh = D3DX.LoadMeshFromX(App.Path & "\Data\xfiles\weapon\" + Format(name) + ".x", D3DXMESH_MANAGED, D3DDevice, Nothing, Nothing, 0)
Set texture = D3DX.CreateTextureFromFile(D3DDevice, App.Path + "\Data\texture\weapon\" + Format(name) + ".jpg")
D3DXMATH_MATRIX.D3DXMatrixIdentity matrix

'################################

'################################ ÇáÇØáÇÞÇÊ
speed = 20
Set bullmesh = D3DX.LoadMeshFromX(App.Path & "\Data\xfiles\bullet\" + Format(name) + ".x", D3DXMESH_MANAGED, D3DDevice, Nothing, Nothing, 0)

For i = 0 To 10
D3DXMATH_MATRIX.D3DXMatrixIdentity bullmatrix(i)
D3DXMATH_MATRIX.D3DXMatrixScaling bullmatrix(i), 0.1, 0.1, 0.1
living(i) = 0
bullpos(i).X = 3600
bullpos(i).Y = 125
bullpos(i).Z = 1300
Next i
'#################################

'############################### ÇááíÒÑ
D3DXMATH_MATRIX.D3DXMatrixIdentity lazermatrix
With lazermatrix
.m41 = 3600
.m42 = 125
.m43 = 1300
End With
Set lazermesh = D3DX.LoadMeshFromX(App.Path & "\Data\xfiles\bullet\" + Format(name) + ".x", D3DXMESH_MANAGED, D3DDevice, Nothing, Nothing, 0)

Set lazersound = New MusicSOUND
lazersound.create "laser"
'###############################



'#################################ÕæÑå ÇáØáÞå
Set bull = D3DX.CreateSprite(D3DDevice)
Set bulltex = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\Data\texture\bullet\bull.bmp", 50, 50, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_DEFAULT, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, &HFFFFFFFF, ByVal 0, ByVal 0)
'#################################

'################################ÇáÕÈå
Set conctex = D3DX.CreateTextureFromFile(D3DDevice, App.Path + "\Data\texture\ground\brick.jpg")
Set concmesh = D3DX.LoadMeshFromX(App.Path & "\Data\xfiles\box.x", D3DXMESH_MANAGED, D3DDevice, Nothing, Nothing, 0)
D3DXMATH_MATRIX.D3DXMatrixIdentity concmat
D3DXMATH_MATRIX.D3DXMatrixScaling concmat, 20, 1, 20
'D3DXMATH_MATRIX.D3DXMatrixMultiply matrix(i), matrix(i), tempmat
With concmat
.m41 = 3600
.m42 = 110
.m43 = 1295
End With

'################################
End Function
Public Function render()
If moving = 0 Then gunmovesound.stopplay
If moving = 1 Then gunmovesound.PLAY

D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
'################################# set weapon matrix to zero
D3DXMATH_MATRIX.D3DXMatrixRotationY tempmatrix, lasttempangel
D3DXMATH_MATRIX.D3DXMatrixRotationX tempmatrix2, lasttempangel2
D3DXMATH_MATRIX.D3DXMatrixMultiply matrix, tempmatrix2, tempmatrix
'################################# rotate the weapon matrix
D3DXMATH_MATRIX.D3DXMatrixRotationY tempmatrix, -realangel
D3DXMATH_MATRIX.D3DXMatrixRotationX tempmatrix2, -realangel2
D3DXMATH_MATRIX.D3DXMatrixMultiply matrix, tempmatrix2, tempmatrix
'#################################
lasttempangel = -realangel
lasttempangel2 = -realangel2

'################################# draw weapon
If weaponname = 2 Then
D3DXMATH_MATRIX.D3DXMatrixScaling tempmat, 5, 5, 5
D3DXMATH_MATRIX.D3DXMatrixMultiply matrix, matrix, tempmat
End If

With matrix
.m41 = EyePos.X
.m42 = ypos
.m43 = EyePos.Z
End With

D3DDevice.SetTexture 0, texture
D3DDevice.SetTransform D3DTS_WORLD, matrix
For i = 0 To 10
mesh.DrawSubset i
Next i
'#################################

''''''''''''''''''''''''''''''''''''''''''''''concret Box render
D3DDevice.SetTexture 0, conctex
D3DDevice.SetTransform D3DTS_WORLD, concmat
concmesh.DrawSubset 0
''''''''''''''''''''''''''''''''''''''''''''''


'################################# draw bullet
If gunmode = 0 Then
updatebulletpos
For i = 1 To 10

With bullmatrix(i)
.m41 = bullpos(i).X
.m42 = bullpos(i).Y
.m43 = bullpos(i).Z
End With
D3DDevice.SetTexture 0, Nothing
D3DDevice.SetMaterial bullmaterial
D3DDevice.SetTransform D3DTS_WORLD, bullmatrix(i)
bullmesh.DrawSubset 0

Next i
End If

'############################### ÇááíÒÑ
If gunmode = 1 And lazershot = 1 Then
bullmaterial = setupmaterials(Rnd(100), Rnd(100), 255)
lazerscale = lazerscale + 2
If lazerscale > 50 Then lazerscale = 0.1: lazershot = 0
D3DXMATH_MATRIX.D3DXMatrixRotationY tempmatrix, lasttempangel
D3DXMATH_MATRIX.D3DXMatrixRotationX tempmatrix2, lasttempangel2
D3DXMATH_MATRIX.D3DXMatrixMultiply lazermatrix, tempmatrix2, tempmatrix

D3DXMATH_MATRIX.D3DXMatrixRotationY tempmatrix, -realangel
D3DXMATH_MATRIX.D3DXMatrixRotationX tempmatrix2, -realangel2
D3DXMATH_MATRIX.D3DXMatrixMultiply lazermatrix, tempmatrix2, tempmatrix

D3DXMATH_MATRIX.D3DXMatrixScaling tempmatrix, 0.1, 0.1, lazerscale
D3DXMATH_MATRIX.D3DXMatrixMultiply lazermatrix, tempmatrix, lazermatrix

With lazermatrix
.m41 = 3600
.m42 = 125
.m43 = 1300
End With
D3DDevice.SetTexture 0, Nothing
D3DDevice.SetMaterial bullmaterial
D3DDevice.SetTransform D3DTS_WORLD, lazermatrix
bullmesh.DrawSubset 0

lazersound.PLAY

End If
'###############################

End Function
Public Function updatebulletpos()

For i = 1 To 10

If living(i) = 1 Then

bullmaterial = setupmaterials(Rnd(100), Rnd(100), Rnd(10))

bullpos(i).X = bullpos(i).X - Sin(bulletangel(i)) * speed
bullpos(i).Y = bullpos(i).Y + Sin(bulletangel2(i)) * speed
bullpos(i).Z = bullpos(i).Z + Cos(bulletangel(i)) * speed

End If

Next i

End Function
Public Function bullshot()

bullnum = bullnum + 1
If bullnum = 10 Then  '++++====

enablegun = 0 'bullnum = 0:
For i = 1 To 10

bullpos(i).X = 3600
bullpos(i).Y = 120
bullpos(i).Z = 1300
living(i) = 0

Next i

End If                 '++++====


bulletangel(bullnum) = realangel
bulletangel2(bullnum) = realangel2
living(bullnum) = 1

End Function
Public Function lazershott()
bullnum = bullnum + 1

If bullnum = 11 Then gunmode = 0: bullnum = 0

End Function
Public Function renderface()
'###############################ÕæÑå ÇáØáÞå
If D3DPP.BackBufferWidth = 1024 Then bull.Draw bulltex, ByVal 0, MV(1, 1), MV(64, 64), 0, MV(0, 0), &HFFFFFFFF
If D3DPP.BackBufferWidth = 800 Then bull.Draw bulltex, ByVal 0, MV(0.8, 0.8), MV(64, 64), 0, MV(0, 0), &HFFFFFFFF
'###############################

'#########################################################weapon nesh
If D3DPP.BackBufferWidth = 1024 Then nesh.Draw neshtex1, ByVal 0, MV(1, 1), MV(56, 56), 0, MV(1024 / 2 - 56 / 2, 768 / 2 - 56 / 2), &HFFFFFFFF '&H5FFF0000 / 1.5
If D3DPP.BackBufferWidth = 800 Then nesh.Draw neshtex1, ByVal 0, MV(0.8, 0.8), MV(56, 56), 0, MV(800 / 2 - 56 * 0.8 / 2, 600 / 2 - 56 * 0.8 / 2), &HFFFFFFFF '&H5FFF0000 / 1.5
'#########################################################
End Function
