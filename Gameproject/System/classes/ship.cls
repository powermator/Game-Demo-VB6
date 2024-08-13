VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ship"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================================ÈÓã Çááå ÇáÑÍãä ÇáÑÍíã
'=================================
'##############################
'#Written by : jaber al_ani   #
'#                            #
'#     jaberani@yahoo.com     #
'##############################
'==================================
'==================================

'=======================  ships
Private mesh             As D3DXMesh
Private matrix(5)       As D3DMATRIX
Private texture          As Direct3DTexture8
Private texture2         As Direct3DTexture8
Private texture3         As Direct3DTexture8


Private ypos(5)           As Integer
Private zpos(5)           As Integer

Private heightvalue      As Single
Private aircombat        As MusicSOUND


Private smallship           As D3DXSprite
Private smallshippic        As Direct3DTexture8
'===========================================

'===============small ships
Private smallmesh             As D3DXMesh
Private smallmatrix(5)       As D3DMATRIX
Private smallxpos(5)           As Integer
Private smallypos(5)           As Integer
Private smallzpos(5)           As Integer
Private smallrandomangel(5)       As Integer
'================================

Private Sub Class_Initialize()

'((ÇáÕæÊ
Set aircombat = New MusicSOUND
aircombat.create "Aircombt"
'))
'**********************************-ääÔÆå ships
'**********************************
Set mesh = D3DX.LoadMeshFromX(App.Path & "\Data\xfiles\ship\" + Format(roundnum) + ".x", D3DXMESH_MANAGED, D3DDevice, Nothing, Nothing, 0)

Set texture = D3DX.CreateTextureFromFile(D3DDevice, App.Path + "\Data\texture\ships\" + Format(roundnum) + ".jpg")

Set texture2 = D3DX.CreateTextureFromFile(D3DDevice, App.Path + "\Data\texture\ships\x.jpg")
Set texture3 = D3DX.CreateTextureFromFile(D3DDevice, App.Path + "\Data\texture\ships\z.jpg")

For I = 1 To 5 ']]]]]]]]]]]]]]]]]
D3DXMATH_MATRIX.D3DXMatrixIdentity matrix(I)
D3DXMATH_MATRIX.D3DXMatrixRotationY tempmat, 270 * 3.14285714285714 / 180
D3DXMATH_MATRIX.D3DXMatrixMultiply matrix(I), matrix(I), tempmat

If roundnum = 2 Then
D3DXMATH_MATRIX.D3DXMatrixScaling tempmat, 2, 2, 2
D3DXMATH_MATRIX.D3DXMatrixMultiply matrix(I), matrix(I), tempmat
End If

If roundnum = 3 Then
D3DXMATH_MATRIX.D3DXMatrixScaling tempmat, 9, 9, 9
D3DXMATH_MATRIX.D3DXMatrixMultiply matrix(I), matrix(I), tempmat

D3DXMATH_MATRIX.D3DXMatrixRotationX tempmat, 90 * 3.14285714285714 / 180
D3DXMATH_MATRIX.D3DXMatrixMultiply matrix(I), matrix(I), tempmat

D3DXMATH_MATRIX.D3DXMatrixRotationZ tempmat, 270 * 3.14285714285714 / 180
D3DXMATH_MATRIX.D3DXMatrixMultiply matrix(I), matrix(I), tempmat

End If

If roundnum = 5 Then
'D3DXMATH_MATRIX.D3DXMatrixRotationZ tempmat, 270 * 3.14285714285714 / 180
'D3DXMATH_MATRIX.D3DXMatrixMultiply matrix(i), matrix(i), tempmat
End If

If roundnum = 6 Then
D3DXMATH_MATRIX.D3DXMatrixRotationY tempmat, 90 * 3.14285714285714 / 180
D3DXMATH_MATRIX.D3DXMatrixMultiply matrix(I), matrix(I), tempmat

D3DXMATH_MATRIX.D3DXMatrixScaling tempmat, 6, 6, 6
D3DXMATH_MATRIX.D3DXMatrixMultiply matrix(I), matrix(I), tempmat


End If

'------------------------äÍÏÏ ÇáãßÇä
xpos(I) = 1300 + Sin(I) * 10
ypos(I) = 250 - Sin(I) * 6
zpos(I) = 1180 + I * 45

'dddddd small ships
smallxpos(I) = 3600 + Sin(I) * 10
smallypos(I) = 250 - Sin(I) * 6
smallzpos(I) = 1300 + I * 45
'dddddd




Next I         ']]]]]]]]]]]]]]]]]]
'*************************************
'*************************************

'========================= picture
Set smallship = D3DX.CreateSprite(D3DDevice)
Set smallshippic = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\Data\texture\ships\superpic\" + Format(roundnum) + ".bmp", 100, 100, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_DEFAULT, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, &HFF000000, ByVal 0, ByVal 0)
'=========================


'===============small ships
Set smallmesh = D3DX.LoadMeshFromX(App.Path & "\Data\xfiles\ship\small.x", D3DXMESH_MANAGED, D3DDevice, Nothing, Nothing, 0)
For I = 1 To 5
D3DXMATH_MATRIX.D3DXMatrixIdentity smallmatrix(I)
Next I


For I = 1 To 5
smallrandomangel(I) = Rnd(I) * 100
With smallmatrix(I)
.m41 = xpos(I) + 1
.m42 = ypos(I) + Sin(heightvalue) * 10
.m43 = zpos(I)
End With
Next I
'================================

End Sub
Public Function render()

D3DDevice.SetMaterial scensematerial
D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW

heightvalue = heightvalue + 0.01
If heightvalue > 180 Then heightvalue = 0

'TTTTTTTTTTTTTTTTTTTTTTT   Small ships attack
'TTTTTTTTTTTTTTTTTTTTTTT
If smallshipattackflag = 1 Then
For I = 1 To 5 Step 2
xpos(I) = xpos(I) + 3
With matrix(I)
.m41 = xpos(I) + 1
.m42 = ypos(I) + Sin(heightvalue) * 10
.m43 = zpos(I)
End With
Next I

For I = 2 To 5 Step 2
xpos(I) = xpos(I) + 3
With matrix(I)
.m41 = xpos(I) + 1
.m42 = ypos(I) - Sin(heightvalue) * 10
.m43 = zpos(I)
End With
Next I
End If

'[[[[[[[[[[[[[[[[[[[ ÇáÇÚÇÏå Çáì äÝÓ ÇáãßÇä
If xpos(1) >= 5000 Then
For I = 1 To 5
xpos(I) = 1200 + Sin(I) * 10
ypos(I) = 250 - Sin(I) * 6
zpos(I) = 1180 + I * 45
Next I
aircombat.stopplay
smallshipattackflag = 0

helicomeflag = helicomeflag + 1  'ÊãÑ ÇáãÑÇßÈ ãÑÊíä áßí ÊãÑ ÇáãÑæÍíå ãÑå æÇÍÏå

End If
'[[[[[[[[[[[[[[[[[[[[

'-----------ÇáÕæÊ
If xpos(1) >= 2000 Then aircombat.PLAY
'----------

'Render ============
If roundnum = 1 Then          '             1
For I = 1 To 5
D3DDevice.SetTransform D3DTS_WORLD, matrix(I)
D3DDevice.SetTexture 0, texture
mesh.DrawSubset 0
Next I
End If

If roundnum = 4 Or roundnum = 2 Then '     2  or  4
For I = 1 To 5
D3DDevice.SetTransform D3DTS_WORLD, matrix(I)
For j = 0 To 4
D3DDevice.SetTexture 0, texture
If j = 1 Then D3DDevice.SetTexture 0, texture2
If j = 2 Then D3DDevice.SetTexture 0, texture3
mesh.DrawSubset j
Next j
Next I
End If


If roundnum = 3 Then          '             3
For I = 1 To 5
D3DDevice.SetTransform D3DTS_WORLD, matrix(I)
For j = 0 To 5
D3DDevice.SetTexture 0, texture
mesh.DrawSubset j
Next j
Next I
End If

If roundnum = 5 Then          '             5
For I = 1 To 5
D3DDevice.SetTransform D3DTS_WORLD, matrix(I)
D3DDevice.SetTexture 0, texture2
For j = 0 To 5
D3DDevice.SetTexture 0, texture
If j >= 0 Then D3DDevice.SetTexture 0, texture3
If j >= 1 Then D3DDevice.SetTexture 0, texture
mesh.DrawSubset j
Next j
Next I
End If


If roundnum = 6 Then          '             6
For I = 1 To 5
D3DDevice.SetTransform D3DTS_WORLD, matrix(I)
D3DDevice.SetTexture 0, texture
For j = 0 To 5
mesh.DrawSubset j
Next j
Next I
End If
'TTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTT
'TTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTT


'ooooooooooo small ships
'OOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOO
For I = 1 To 5
'smallxpos(I) = smallxpos(I) + Sin(smallrandomangel(I))  '* 10
'smallypos(I) = smallypos(I) '+ Sin(heightvalue) '* 20
'smallzpos(I) = smallzpos(I) '+ Cos(smallrandomangel(I)) '* 10

'smallzpos(I) = Sqr(1000000 - smallxpos(I) * smallxpos(I))


With smallmatrix(I)
.m41 = smallxpos(I) '+ Sin(heightvalue) * 20
.m42 = smallypos(I) '+ Sin(heightvalue) * 20
.m43 = smallzpos(I) '+ Sin(heightvalue) * 20
End With
Next I

For I = 1 To 5
D3DDevice.SetTransform D3DTS_WORLD, smallmatrix(I)
D3DDevice.SetTexture 0, texture2
For j = 0 To 5
D3DDevice.SetTexture 0, texture
If j >= 0 Then D3DDevice.SetTexture 0, texture3
If j >= 1 Then D3DDevice.SetTexture 0, texture
smallmesh.DrawSubset j
Next j
Next I
'ooooooooooo
'OOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOO

End Function
Public Function renderface()

'=============== super ship
If D3DPP.BackBufferWidth = 1024 Then smallship.Draw smallshippic, ByVal 0, MV(1, 1), MV(0, 0), 0, MV(900, 80), &HFFFFFFFF
'===============

End Function





'for 3

'For j = 0 To 4
'D3DDevice.SetTexture 0, texture2
'If j >= 1 Then D3DDevice.SetTexture 0, texture
'mesh.DrawSubset j
'Next j
'Next i


'for 4   'æÍÏ
'For j = 0 To 4
'D3DDevice.SetTexture 0, texture2
'If j = 1 Then D3DDevice.SetTexture 0, texture3
'If j >= 2 Then D3DDevice.SetTexture 0, texture
'mesh.DrawSubset j
'Next j




'===============
'===============

