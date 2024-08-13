VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "explo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'####################
'## CONTROL privateANTS ##
'###################
Private nParticles As Long   '//How many particles we are going to use
Private ParticleSize As Single
Private Gravity As Single
Private XWind As Single
Private ZWind As Single
Private XVariation As Single  '//Small number = tight *funnel* of particles
Private YVariation As Single
Private ZVariation As Single  '//Big number = explosion like particle effect
Private PrtData(0 To 250 - 1) As PARTICLE
Private PrtVertList(0 To 250 - 1) As PARTICLEVERTEX
Dim LastUpdatedParticles As Long

Private explomat As D3DMATRIX
Dim ParticleTex As Direct3DTexture8

'############
'## GEOMETRY ##
'############
Private Type PARTICLEVERTEX
    v As D3DVECTOR
    color As Long
    tu As Single
    tv As Single
End Type
'##############
'##  PARTICLE FX  ##
'##############
Private Enum PARTICLE_STATUS
    Alive = 0
    Dead = 1
End Enum

Private Type PARTICLE
    X As Single     'World Space Coordinates
    Y As Single
    Z As Single
    vX As Single    'Speed and Direction
    vY As Single
    vZ As Single
    CurrentColor As D3DCOLORVALUE
    LifeTime As Long    'How long Mr. Particle Exists
    Created As Long 'When this particle was created...
    Status As PARTICLE_STATUS 'Does he even exist?
End Type
Private Sub Class_Initialize()
nParticles = 250
ParticleSize = 0.01
Gravity = -0.00005
XWind = 0
ZWind = 0
XVariation = 2 '//Small number = tight *funnel* of particles
YVariation = 0.85
ZVariation = 2  '//Big number = explosion like particle effect


D3DXMATH_MATRIX.D3DXMatrixIdentity explomat
D3DXMATH_MATRIX.D3DXMatrixScaling explomat, 0.1, 0.1, 0.1
With explomat
.m41 = 3300
.m42 = 255
.m43 = 1300
End With

Set ParticleTex = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\Data\texture\explo\explo.bmp", 20, 20, D3DX_DEFAULT, 0, D3DFMT_A1R5G5B5, D3DPOOL_DEFAULT, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, &HFFFF00FF, ByVal 0, ByVal 0)


Dim I As Integer
For I = 0 To nParticles - 1
    PrtData(I).Status = Alive
    PrtData(I).LifeTime = 10000 + ((Rnd * 5000) - 2500)
    PrtData(I).Created = GetTickCount
    PrtData(I).X = 0
    PrtData(I).Y = -0.5
    PrtData(I).Z = 0
    PrtData(I).vX = (Rnd * XVariation) - (XVariation / 2)
    PrtData(I).vY = (Rnd * YVariation) - (YVariation / 3)
    PrtData(I).vZ = (Rnd * ZVariation) - (ZVariation / 2)
        Randomize
Next I
Call GenerateVertexDataFromParticles
End Sub
Private Sub UpdateParticles()
'//0. Any variables required
    Dim I As Long
    

'//1. Loop through all particles
    For I = 0 To nParticles - 1
        If PrtData(I).Status = Alive Then
                '//Update the positions
                    PrtData(I).X = PrtData(I).X + ((PrtData(I).vX / 500) * 10000) ' (GetTickCount - LastUpdatedParticles))
                    PrtData(I).Y = PrtData(I).Y + ((PrtData(I).vY / 500) * 10000) '(GetTickCount - LastUpdatedParticles))
                    PrtData(I).Z = PrtData(I).Z + ((PrtData(I).vZ / 500) * 10000) ' (GetTickCount - LastUpdatedParticles))
                    
                '//Update the velocities
                    PrtData(I).vX = PrtData(I).vX + ((XWind / 500) * 10000) '(GetTickCount - LastUpdatedParticles))
                    PrtData(I).vY = PrtData(I).vY + ((Gravity / 500) * 100000) ' (GetTickCount - LastUpdatedParticles))
                    PrtData(I).vZ = PrtData(I).vZ + ((ZWind / 500) * 10000) '(GetTickCount - LastUpdatedParticles))
                    
                    'Randomize
                '//Check if the particle has gone below ground level...
                   ' If PrtData(I).Y < -1 Then PrtData(I).Status = Dead
                
        End If
    Next I

LastUpdatedParticles = GetTickCount

'//2. Update the raw vertex data
    Call GenerateVertexDataFromParticles
End Sub
Public Function render()
UpdateParticles
    
D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
'These next two calls define how we make colours transparent, for a specific colour
'leave them as they are
D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
'Enable the usage of transparencies
D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
'Enable point sprite rendering
D3DDevice.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1 '//Enable point sprite rendering
'Allow Direct3D to size the points
D3DDevice.SetRenderState D3DRS_POINTSCALE_ENABLE, 1 '//Allow Direct3D to set/alter the size of the Psprites
'Set the size
D3DDevice.SetRenderState D3DRS_POINTSIZE, FtoDW(ParticleSize)


D3DDevice.SetMaterial scensematerial

 '//Draw our particles
   D3DDevice.SetTransform D3DTS_WORLD, explomat
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    D3DDevice.SetRenderState D3DRS_ZWRITEENABLE, 1 '//Stops Particles screwing things up :)
    D3DDevice.SetTexture 0, ParticleTex ' ParticleTex
D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0
'D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, nParticles, PrtVertList(I), Len(PrtVertList(0))
D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, nParticles, PrtVertList(I), Len(PrtVertList(0))
    
End Function
Private Function GenerateVertexDataFromParticles()
Dim I As Long
For I = 0 To nParticles - 1
    PrtVertList(I).v = MakeVector(PrtData(I).X, PrtData(I).Y, PrtData(I).Z)
    PrtVertList(I).color = D3DColorARGB(PrtData(I).CurrentColor.A * 255, PrtData(I).CurrentColor.r * 255, PrtData(I).CurrentColor.g * 255, PrtData(I).CurrentColor.b * 255)
Next I
End Function
Private Function CreateColorVal(A As Integer, r As Integer, g As Integer, b As Integer) As D3DCOLORVALUE
    CreateColorVal.A = A
    CreateColorVal.r = r
    CreateColorVal.g = g
    CreateColorVal.b = b
End Function
'//Allows us to pass a floating point number as a long - for SetRenderState calls...
Function FtoDW(f As Single) As Long
    Dim buf As D3DXBuffer
    Dim l As Long
    Set buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData buf, 0, 4, 1, f
    D3DX.BufferGetData buf, 0, 4, 1, l
    FtoDW = l
End Function
