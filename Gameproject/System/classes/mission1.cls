VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mission"
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


''''''''''''''''''''''''''''''''''''''''''''mission1 variables
Private rotangel            As Single
Private groundvertex(59000) As VERTEX
Private groundvb            As Direct3DVertexBuffer8
''''''''''''''''''''''''''''''''''''''''''''material
Private coulor               As D3DCOLORVALUE
'''''''''''''''''''''''''''''''''''''''''''''mesh
Private skymesh        As D3DXMesh
Private watermesh      As D3DXMesh
'''''''''''''''''''''''''''''''''''''''''''''matrix
Private skymat         As D3DMATRIX
Private watermat       As D3DMATRIX
'''''''''''''''''''''''''''''''''''''''''''''textures
Private sunskytex         As Direct3DTexture8
Private riseskytex         As Direct3DTexture8
Private nightskytex         As Direct3DTexture8
Private groundtex      As Direct3DTexture8
Private watertex       As Direct3DTexture8

Private redcol           As Single
Private greencol         As Single
Private bluecol          As Single

Private Sub Class_Initialize()
'''''''''''''''''''''''''''''''''''''''''' set texture
''''''''''''''''''''''''''''''''''''''''''
Set groundtex = D3DX.CreateTextureFromFile(D3DDevice, App.Path + "\Data\texture\ground\m1.jpg")
Set watertex = D3DX.CreateTextureFromFile(D3DDevice, App.Path + "\Data\texture\water\m1.jpg")

Set sunskytex = D3DX.CreateTextureFromFile(D3DDevice, App.Path + "\Data\texture\sky\m1day.jpg")
Set riseskytex = D3DX.CreateTextureFromFile(D3DDevice, App.Path + "\Data\texture\sky\m1rise.jpg")
Set nightskytex = D3DX.CreateTextureFromFile(D3DDevice, App.Path + "\Data\texture\sky\m1night.jpg")

'''''''''''''''''''''''''''''
''''''''''''''''''set mesh
Set skymesh = D3DX.LoadMeshFromX(App.Path & "\Data\xfiles\sky.x", D3DXMESH_MANAGED, D3DDevice, Nothing, Nothing, 0)
D3DXMATH_MATRIX.D3DXMatrixIdentity skymat

Set watermesh = D3DX.LoadMeshFromX(App.Path & "\Data\xfiles\water.x", D3DXMESH_MANAGED, D3DDevice, Nothing, Nothing, 0)
D3DXMATH_MATRIX.D3DXMatrixIdentity watermat

'''''''''''''''''''''''''''''''''''''''' draw ground
''''''''''''''''''''''''''''''''''''''''
Open App.Path + "\Data\map\map.txt" For Input As #2
Set groundvb = D3DDevice.CreateVertexBuffer(24 * 43500, 0, D3DFVF_VERTEX, D3DPOOL_SYSTEMMEM)
Dim h1, h2, h3, h4 As String

''''''''''
''''''''''
Dim X As Long
X = 1
For i = 1 To 49
DoEvents
For j = 1 To 49
Input #2, h1, h2, h3, h4

If j > 23 Then
groundvertex(X) = MakeVertex(j * 100, Val(h1) * 10, i * 60, 0, 0)
   X = X + 1
groundvertex(X) = MakeVertex((j * 100) - 100, Val(h2) * 10, i * 60, 1, 0)
   X = X + 1
groundvertex(X) = MakeVertex((j * 100) - 100, Val(h3) * 10, (i * 60) - 60, 1, 1)
   X = X + 1
groundvertex(X) = MakeVertex(j * 100, Val(h1) * 10, i * 60, 0, 0)
   X = X + 1
groundvertex(X) = MakeVertex((j * 100) - 100, Val(h3) * 10, (i * 60) - 60, 1, 1)
   X = X + 1
groundvertex(X) = MakeVertex(j * 100, Val(h4) * 10, (i * 60) - 60, 0, 1)
   X = X + 1
   
End If

Next j
Next i
Close #2
D3DVertexBuffer8SetData groundvb, 0, 24 * 33000, 0, groundvertex(1)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''

enablefog = 1
sun = 1
''''''''''''''''''''''''''''''''''''''''''''''''
'sky'
D3DXMATH_MATRIX.D3DXMatrixRotationX tempmat, -11
D3DXMATH_MATRIX.D3DXMatrixScaling skymat, 2, 2, 2
D3DXMATH_MATRIX.D3DXMatrixMultiply skymat, skymat, tempmat
With skymat
.m41 = 2600
.m42 = -1000
.m43 = 1400
End With

''''''''''''''''''''''''''''''''''''''''''
'water'
D3DXMATH_MATRIX.D3DXMatrixScaling watermat, 4, 0.5, 5
With watermat
.m41 = 2000
.m42 = -40
.m43 = 2450
End With
'''''''''''''''''''''''''''''''''''''''''''''
currentmaterial = scensematerial
'''''''''''''''''''''''''''''''''''''''''''''
End Sub

Public Function render()

'******
'*******
'********
If sun = 1 Then redcol = 10: greencol = 10: bluecol = 10: drawfog 0, 5000
'===================
If rise = 1 Then
redcol = redcol - 0.01: greencol = greencol - 0.01: bluecol = bluecol - 0.01 'ÇáÞíãå ÊÊäÇÞÕ
If redcol < 0.6 Then redcol = 0.6
If greencol < 0.5 Then greencol = 0.5
If bluecol < 0.4 Then bluecol = 0.4
End If
'===================
If night = 1 Then
redcol = redcol - 0.01: greencol = greencol - 0.01: bluecol = bluecol - 0.01 'ÇáÞíãå ÊÊäÇÞÕ
If redcol < 0.3 Then redcol = 0.3
If greencol < 0.3 Then greencol = 0.3
If bluecol < 0.3 Then bluecol = 0.3
End If
'===================
scensematerial = setupmaterials(redcol, greencol, bluecol)
'=================================

If nkey = 1 Then scensematerial = setupmaterials(0, 10, 0): enablefog = 0: D3DDevice.SetRenderState D3DRS_FOGENABLE, 0

'**************
'**************
'**************
D3DDevice.SetMaterial scensematerial
D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CW

'''''''''''''''''''''''''''''''''''''''''''''' sky render
D3DDevice.SetTexture 0, sunskytex
If redcol = 0.6 And greencol = 0.5 And bluecol = 0.4 Then D3DDevice.SetTexture 0, riseskytex
If redcol = 0.3 And greencol = 0.3 And bluecol = 0.3 Then D3DDevice.SetTexture 0, nightskytex

D3DDevice.SetTransform D3DTS_WORLD, skymat
skymesh.DrawSubset 0
'''''''''''''''''''''''''''''''''''''''''''''


D3DDevice.SetVertexShader D3DFVF_VERTEX
D3DDevice.SetTransform D3DTS_WORLD, matworld
 ''''''''''''''''''''''''''''''''''''''''''''' ground render
D3DDevice.SetTexture 0, groundtex
D3DDevice.SetStreamSource 0, groundvb, 24
D3DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 3000


''''''''''''''''''''''''''''''''''''''''''''''''''water render
''''''''''''''''''''''''''''''''''''''''''''''''''
rotangel = rotangel + 0.01
If rotangel > 360 Then rotangel = 0

With watermat
.m42 = -40 + Cos(rotangel) * 10
End With

D3DDevice.SetTexture 0, watertex
D3DDevice.SetTransform D3DTS_WORLD, watermat
watermesh.DrawSubset 0
'''''''''''''''''''''''''''''''''''''''''''''''''''''
If roundnum > 7 And roundnum <= 10 And enablefog = 1 Then drawfog 800, 1500
If roundnum > 5 And roundnum <= 7 And enablefog = 1 Then drawfog 0, 5000
If roundnum <= 5 And enablefog = 1 Then drawfog 0, 5000

'''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Function

