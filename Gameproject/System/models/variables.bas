Attribute VB_Name = "help"
''''''''''''''''''''''''''''''''ÈÓã Çááå ÇáÑÍãä ÇáÑÍíã
''''''''''''''''''''''''''''''''
'##############################
'#####################################################################
'##############################ãÊÛíÑÇÊ ÇááÚÈå
'#####################################################################
'=============================================ãÊÛíÑÇÊ ÚÇãå
Public roundnum         As Byte
Public startshow        As Byte
Public faceshow         As Byte
Public startgame        As Byte

Public face             As Byte
Public PLAY             As Byte
Public yesrot           As Byte
Public cdshow           As Byte
Public mouseinp         As Byte         'ÇáãÊÊÛíÑ ÇáÐí íÍãá ÑÞã ÇáÒÑ ÇáãÖÛæØ ãä ÇáãÇæÓ
Public tempval          As Single
Public EyePos           As D3DVECTOR    'ãßÇä ÇáßÇãÑå
Public EyeDir           As D3DVECTOR    'ÇÊÌÇå ÇáäÙÑ
Public temptext         As String
Public allshipbuff      As D3DXBuffer
Public rotangel         As Single
Public time             As Byte     'æÞÊ áÈÏÇíå ÇááÚÈå
Public ntime            As Byte

'====================================== for tick count
Public FrameCount As Long
Public FrameRate As Long
Public FrameRateLastCheck As Long
Public LastUpdated As Long
'=============================================************=========
Public ttext            As text
Public realangel           As Single
Public realangel2          As Single
Public collisionflag As Byte
Public bullnum    As Integer
Public lazerbullnum As Byte
Public scensematerial       As D3DMATERIAL8
Public enablefog As Byte
Public currentmaterial As D3DMATERIAL8
Public nkey As Byte
Public moving As Byte

Public smallshipattackflag   As Byte
Public helicomeflag          As Byte
Public lazershot               As Byte
Public enablegun As Byte


Public xpos(10)           As Integer

Public sun  As Byte
Public rise As Byte
Public night As Byte
'===============================================ÇæÈÌßÊÓ
Public startlogo       As logo
Public gameface        As face
Public inp             As Directinput
Public mission         As mission
Public mouseoversound       As MusicSOUND
Public moseclick       As MusicSOUND
Public curpos          As POINTAPI
Public comp            As compass
Public gun             As weapon
Public shotsound       As MusicSOUND
Public gunmovesound    As MusicSOUND
Public game            As gameengine
Public facemusic       As MusicMP3
Public ships           As ship
Public helicopter      As heli
Public gunmode        As Byte
'===============================================ËæÇÈÊ
Public Const FVF_TransformedAndLit As Long = (D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR)
Public Const D3DFVF_VERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)
Public Const FVF_PARTICLEVERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)
Public Const pi As Single = 3.14159275180032
'============================================== dx variables
Public DX         As New DirectX8
Public D3DX       As New D3DX8
Public D3D        As Direct3D8
Public D3DDevice  As Direct3DDevice8

Public DIDevice   As DirectInputDevice8
Public DispMode   As D3DDISPLAYMODE
Public D3DPP      As D3DPRESENT_PARAMETERS
'ÇáÇÚáÇä Úä ßÇÆä ÍÇáå ÇáãÝÇÊíÍ ááãÇæÓ
Public MouseState As DIMOUSESTATE
''''''''''''''''''' all matrices
Public matworld              As D3DMATRIX
Public matView               As D3DMATRIX
Public matProj               As D3DMATRIX
Public tempmat               As D3DMATRIX
''''''''''''''''''''
'EXPLOooooooooooooooooooo
Public explosion   As explo
'oooooooooooooooooooooooo




'#####################################################################
'#############################ÓÊÑßÌÑÒ
'#####################################################################

Public Type spaceship
mesh         As D3DXMesh
tex          As Direct3DTexture8
matrix       As D3DMATRIX
End Type
'=================================
Public Type VERTEX
    position As D3DVECTOR
    color As Byte
    tu As Single
    tv As Single
End Type
'=================================
Public Type TLVERTEX
    X As Single
    Y As Single
    Z As Single
    RHW As Single
    color As Long
    Specular As Long
    tu As Single
    tv As Single
End Type    '32 bytes
'=================================
Public Type POINTAPI
        X As Long
        Y As Long
End Type
'#####################################################################
'##############################ÃáÏæÇá
'#####################################################################

Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
'#####################################################################
'##############################ÝäßÔäÇÊ ãÓÇÚÏå
'#####################################################################
Public Function MakeVertex(X As Integer, Y As Integer, Z As Integer, U As Single, v As Single) As VERTEX
    With MakeVertex
        .position.X = X
        .position.Y = Y
        .position.Z = Z
        .tu = U
        .tv = v
    End With
End Function
Public Function MakeVector(X As Single, Y As Single, Z As Single) As D3DVECTOR
    With MakeVector
        .X = X
        .Y = Y
        .Z = Z
    End With
End Function
Public Function Normalise(v As D3DVECTOR, Optional CARRY As Single = 0, Optional Qual As Single = 1) As D3DVECTOR
    Dim Leng As Single
    With Normalise
        Leng = Sqr(v.X * v.X + v.Y * v.Y + v.Z * v.Z) / Qual
        If Leng > 0 Then
            .X = v.X / Leng
            .Y = v.Y / Leng
            .Z = v.Z / Leng
        End If
        CARRY = Leng
    End With
End Function
Public Function VectorSubtract(V1 As D3DVECTOR, V2 As D3DVECTOR) As D3DVECTOR
    With VectorSubtract
        .X = V1.X - V2.X
        .Y = V1.Y - V2.Y
        .Z = V1.Z - V2.Z
    End With
End Function
Public Function CrossProduct(V1 As D3DVECTOR, V2 As D3DVECTOR) As D3DVECTOR
    With CrossProduct
        .X = V1.Y * V2.Z - V1.Z * V2.Y
        .Y = V1.Z * V2.X - V1.X * V2.Z
        .Z = V1.X * V2.Y - V1.Y * V2.X
    End With
End Function
Public Function DotProduct(V1 As D3DVECTOR, V2 As D3DVECTOR) As Single
    DotProduct = V1.X * V2.X + V1.Y * V2.Y + V1.Z * V2.Z
End Function
Public Function Dist(V1 As D3DVECTOR, V2 As D3DVECTOR) As Single
    Dist = Sqr((V2.X - V1.X) * (V2.X - V1.X) + (V2.Y - V1.Y) * (V2.Y - V1.Y) + (V2.Z - V1.Z) * (V2.Z - V1.Z))
End Function
Public Function vectorspace(v As D3DVECTOR) As Single
vectorspace = Sqr(v.X * v.X + v.Y * v.Y + v.Z * v.Z)
End Function

Public Function FtoDW(flo As Single) As Long
    Dim buf As D3DXBuffer
    Set buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData buf, 0, 4, 1, flo
    D3DX.BufferGetData buf, 0, 4, 1, FtoDW
End Function
Public Function CreateD3DColorVal(A As Single, r As Single, g As Single, b As Single) As D3DCOLORVALUE
    CreateD3DColorVal.A = A
    CreateD3DColorVal.r = r
    CreateD3DColorVal.g = g
    CreateD3DColorVal.b = b
End Function
Public Function CreateTLVertex(X As Single, Y As Single, Z As Single, RHW As Single, Diffuse As Long, Specular As Long, tu As Single, tv As Single) As TLVERTEX
    CreateTLVertex.X = X
    CreateTLVertex.Y = Y
    CreateTLVertex.Z = Z
    CreateTLVertex.RHW = RHW
    CreateTLVertex.color = Diffuse
    CreateTLVertex.Specular = Specular
    CreateTLVertex.tu = tu
    CreateTLVertex.tv = tv
End Function
Public Function MV(X As Single, Y As Single) As D3DVECTOR2
With MV: .X = X: .Y = Y: End With
End Function
Public Function Mrect(left As Integer, top As Integer, Width As Integer, Height As Integer) As RECT
Mrect.left = left
Mrect.top = top
Mrect.Right = left + Width
Mrect.bottom = top + Height
End Function
Public Function RotateCamera(Alpha As Single, Beta As Single)

    Dim MatWorld2 As D3DMATRIX, matworld As D3DMATRIX
    Dim yy As Single
    D3DXMatrixRotationAxis MatWorld2, MakeVector(EyeDir.Z, 0, -EyeDir.X), Sin(Beta)
    D3DXMatrixRotationY matworld, Sin(Alpha)
    D3DXMatrixMultiply matworld, matworld, MatWorld2
    
    With EyeDir
        yy = .X * matworld.m12 + .Y * matworld.m22 + .Z * matworld.m32
        If yy < -0.3 Then
            yy = -0.3
        ElseIf yy > 0.8 Then
            yy = 0.8
        End If
        EyeDir = MakeVector(.X * matworld.m11 + .Y * matworld.m21 + .Z * matworld.m31, _
                                yy, _
                                .X * matworld.m13 + .Y * matworld.m23 + .Z * matworld.m33)
        EyeDir = Normalise(EyeDir)
    End With
End Function
Public Function setupmaterials(r As Single, g As Single, b As Single) As D3DMATERIAL8
Dim col As D3DCOLORVALUE
With col
.r = r
.g = g
.b = b
End With

With setupmaterials
.Ambient = col
.Diffuse = col
.emissive = col
.Ambient = col
.power = 10
End With


End Function
Public Function drawfog(startfog As Single, endfog As Single)
   
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCCOLOR
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_DESTCOLOR
    D3DDevice.SetRenderState D3DRS_FOGENABLE, 1
    D3DDevice.SetRenderState D3DRS_FOGTABLEMODE, D3DFOG_LINEAR
    D3DDevice.SetRenderState D3DRS_FOGVERTEXMODE, D3DFOG_LINEAR
    D3DDevice.SetRenderState D3DRS_RANGEFOGENABLE, 0
    
    D3DDevice.SetRenderState D3DRS_FOGSTART, FtoDW(startfog)
    D3DDevice.SetRenderState D3DRS_FOGEND, FtoDW(endfog)
    D3DDevice.SetRenderState D3DRS_FOGCOLOR, vbRed / 4 'vbgray
    
End Function
