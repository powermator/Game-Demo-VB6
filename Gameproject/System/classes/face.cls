VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "face"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''ÈÓã Çááå ÇáÑÍãä ÇáÑÍíã
''''''''''''''''''''''''''''''''
'===================================== Interface
'##############################
'#Written by : jaber al_ani   #
'#Date       : 22- 2 -2006    #
'#Time       : 11 :09 PM      #
'##############################

Private background        As Direct3DSurface8
Private text(5)           As text
Private fontsize          As Byte
Private color             As Long
Private face              As Byte
Public scan               As Byte
Private colorsaturation   As Single
Private inbord            As Byte
'''''''''''''''''''''''''''''''''''''''''''''''''''''quiit
Private quitmesh        As D3DXSprite
Private quittex         As Direct3DTexture8
Private quitmenue       As Byte
Private quittext(2)     As text
Private playsound       As Byte
Private Sub Class_Initialize()
scan = 1
inbord = 0
colorsaturation = 0.01

Set background = D3DDevice.CreateRenderTarget(D3DPP.BackBufferWidth, D3DPP.BackBufferHeight, D3DFMT_X8R8G8B8, D3DMULTISAMPLE_NONE, True)
D3DX.LoadSurfaceFromFile background, ByVal 0, ByVal 0, App.Path & "\Data\texture\face\face.jpg", ByVal 0, D3DX_FILTER_LINEAR, &HFF00FF00, ByVal 0

For i = 0 To 4
Set text(i) = New text
Next i

If D3DPP.BackBufferWidth = 1024 Then fontsize = 40
If D3DPP.BackBufferWidth = 800 Then fontsize = 34
text(0).create fontsize, &HFFFFFFFF, "ÅÈÏÇ áÚÈÜå ÌÏíÜÏå", 0
text(1).create fontsize, &HFFFFFFFF, "   ÊÍãÜíá áÚÈÜå   ", 0
text(2).create fontsize, &HFFFFFFFF, "    ÅÚÜÜÏÇÏÇÊ    ", 0
text(3).create fontsize, &HFFFFFFFF, "     ÃáÊÍÑíÑ     ", 0
text(4).create fontsize, &HFFFFFFFF, "       ÅäåÇÁ      ", 0

''''''''''''''''''''''''''''''''''''''quit
quitmenue = 0
Set quitmesh = D3DX.CreateSprite(D3DDevice)
Set quittex = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\Data\texture\face\quit.jpg", 324, 155, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_DEFAULT, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, vbRed, ByVal 0, ByVal 0)

For i = 0 To 1
Set quittext(i) = New text
Next i

quittext(0).create fontsize, &H5FFF0000, "äÚã", 0
quittext(1).create fontsize, &H5FFF0000, " áÇ ", 0
End Sub
Public Function render()
If playsound = 0 Then Set facemusic = Nothing: Set facemusic = New MusicMP3: facemusic.create "3": facemusic.switchsound 1: playsound = 1

D3DDevice.CopyRects background, ByVal 0, 0, D3DDevice.GetBackBuffer(0, D3DBACKBUFFER_TYPE_MONO), ByVal 0

drawbootom
If scan = 1 Then checkinput
If quitmenue = 1 Then quit

End Function
Public Function drawbootom()
''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''
colorsaturation = colorsaturation + 0.000000001
If D3DPP.BackBufferWidth = 1024 Then
inbord = 0

For i = 0 To 4
color = &HFF0000FF / colorsaturation
If i = 4 And scan = 1 And curpos.X > 460 And curpos.X < 540 And curpos.Y > 350 And curpos.Y < 390 Then color = &H5FFF0000 / 1.5: inbord = 1
If i = 3 And scan = 1 And curpos.X > 431 And curpos.X < 557 And curpos.Y > 288 And curpos.Y < 331 Then color = &H5FFF0000 / 1.5: inbord = 1
If i = 2 And scan = 1 And curpos.X > 416 And curpos.X < 570 And curpos.Y > 233 And curpos.Y < 258 Then color = &H5FFF0000 / 1.5: inbord = 1
If i = 1 And scan = 1 And curpos.X > 400 And curpos.X < 603 And curpos.Y > 168 And curpos.Y < 207 Then color = &H5FFF0000 / 1.5: inbord = 1
If i = 0 And scan = 1 And curpos.X > 355 And curpos.X < 614 And curpos.Y > 107 And curpos.Y < 137 Then color = &H5FFF0000 / 1.5: inbord = 1

text(i).render 355, 100 + (i * 60), 400, 100, color
Next i

If scan = 1 Then
If inbord = 0 Then ntime = 0
If inbord = 1 Then ntime = ntime + 1
If ntime > 100 Then ntime = 50
If ntime = 1 Then mouseoversound.PLAY
End If

'''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''
If D3DPP.BackBufferWidth = 800 Then
inbord = 0
For i = 0 To 4
color = &HFF0000FF / colorsaturation
If i = 4 And scan = 1 And curpos.X > 460 - 90 And curpos.X < 540 - 120 And curpos.Y > 350 - 50 And curpos.Y < 390 - 50 Then color = &H5FFF0000 / 1.5: inbord = 1
If i = 3 And scan = 1 And curpos.X > 431 - 90 And curpos.X < 557 - 120 And curpos.Y > 288 - 30 And curpos.Y < 331 - 40 Then color = &H5FFF0000 / 1.5: inbord = 1
If i = 2 And scan = 1 And curpos.X > 416 - 90 And curpos.X < 570 - 120 And curpos.Y > 233 - 30 And curpos.Y < 258 - 20 Then color = &H5FFF0000 / 1.5: inbord = 1
If i = 1 And scan = 1 And curpos.X > 400 - 90 And curpos.X < 603 - 120 And curpos.Y > 168 - 20 And curpos.Y < 207 - 20 Then color = &H5FFF0000 / 1.5: inbord = 1
If i = 0 And scan = 1 And curpos.X > 355 - 90 And curpos.X < 614 - 120 And curpos.Y > 107 And curpos.Y < 137 Then color = &H5FFF0000 / 1.5: inbord = 1

text(i).render 270, 100 + (i * 50), 400, 100, color
Next i

If scan = 1 Then             ' if
If inbord = 0 Then ntime = 0
If inbord = 1 Then ntime = ntime + 1
If ntime > 100 Then ntime = 50
If ntime = 1 Then mouseoversound.PLAY
End If                      ' end if

End If

''''''''''''''''''''''''''''''''áÈÏÇíå ÇááÚÈå''''
If scan = 1 And curpos.X > 355 And curpos.X < 614 And curpos.Y > 107 And curpos.Y < 137 And mouseinp = 0 Then startgame = 1: faceshow = 0: Set gameface = Nothing: moseclick.PLAYnow: SetCursorPos 100, 100
''''''''''''''''''''''''''''''''''''''''''''''''
End If

End Function
Private Sub checkinput()
If D3DPP.BackBufferWidth = 1024 Then
For i = 0 To 4
If i = 4 And curpos.X > 460 And curpos.X < 540 And curpos.Y > 350 And curpos.Y < 390 And mouseinp = 0 Then moseclick.PLAYnow: quitmenue = 1
If i = 3 And curpos.X > 431 And curpos.X < 557 And curpos.Y > 288 And curpos.Y < 331 Then color = &H5FFF0000 / 1.5
If i = 2 And curpos.X > 416 And curpos.X < 570 And curpos.Y > 233 And curpos.Y < 258 Then color = &H5FFF0000 / 1.5
If i = 1 And curpos.X > 400 And curpos.X < 603 And curpos.Y > 168 And curpos.Y < 207 Then color = &H5FFF0000 / 1.5
If i = 0 And curpos.X > 355 And curpos.X < 614 And curpos.Y > 107 And curpos.Y < 137 Then color = &H5FFF0000 / 1.5
Next i

End If
'''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''
If D3DPP.BackBufferWidth = 800 Then
For i = 0 To 4
If i = 4 And curpos.X > 460 - 90 And curpos.X < 540 - 120 And curpos.Y > 350 - 50 And curpos.Y < 390 - 50 And mouseinp = 0 Then quitmenue = 1: moseclick.PLAYnow
If i = 3 And curpos.X > 431 - 90 And curpos.X < 557 - 120 And curpos.Y > 288 - 30 And curpos.Y < 331 - 40 Then color = &H5FFF0000 / 1.5
If i = 2 And curpos.X > 416 - 90 And curpos.X < 570 - 120 And curpos.Y > 233 - 30 And curpos.Y < 258 - 20 Then color = &H5FFF0000 / 1.5
If i = 1 And curpos.X > 400 - 90 And curpos.X < 603 - 120 And curpos.Y > 168 - 20 And curpos.Y < 207 - 20 Then color = &H5FFF0000 / 1.5
If i = 0 And curpos.X > 355 - 90 And curpos.X < 614 - 120 And curpos.Y > 107 And curpos.Y < 137 Then color = &H5FFF0000 / 1.5
Next i
End If


End Sub
Private Function quit()
scan = 0
inbord = 0
'[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[1024
If D3DPP.BackBufferWidth = 1024 Then

quitmesh.Draw quittex, ByVal 0, MV(1, 1), MV(324, 155), 0, MV(240, 420), &H5FFF0000 / 1.5

color = &HFF0000FF / colorsaturation
If curpos.X > 600 And curpos.X < 650 And curpos.Y > 560 And curpos.Y < 600 Then color = &H5FFF0000 / 1.5: inbord = 1
quittext(0).render 600, 550, 300, 100, color

color = &HFF0000FF / colorsaturation
If curpos.X > 366 And curpos.X < 390 And curpos.Y > 560 And curpos.Y < 600 Then color = &H5FFF0000 / 1.5: inbord = 1
quittext(1).render 350, 550, 300, 100, color


'text(0).create fontsize, &H5FFF0000, Str(ntime) + Str(inbord), 0
If inbord = 0 Then ntime = 0
If inbord = 1 Then ntime = ntime + 1
If ntime > 100 Then ntime = 50
If ntime = 1 Then mouseoversound.PLAY
'''''###### quit or back
'äÚã
If curpos.X > 600 And curpos.X < 650 And curpos.Y > 560 And curpos.Y < 600 And mouseinp = 0 Then moseclick.PLAYnow: End
'áÇ
If curpos.X > 366 And curpos.X < 390 And curpos.Y > 560 And curpos.Y < 600 And mouseinp = 0 Then quitmenue = 0: moseclick.PLAYnow: scan = 1
'##########
End If
']]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]

'[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[ 800
If D3DPP.BackBufferWidth = 800 Then
quitmesh.Draw quittex, ByVal 0, MV(0.71, 0.71), MV(324, 155), 0, MV(214, 383), &H5FFF0000 / 1.5

color = &HFF0000FF / colorsaturation
If curpos.X > 480 And curpos.X < 523 And curpos.Y > 464 And curpos.Y < 495 Then color = &H5FFF0000 / 1.5: inbord = 1
quittext(0).render 480, 450, 300, 100, color

color = &HFF0000FF / colorsaturation
If curpos.X > 285 And curpos.X < 302 And curpos.Y > 455 And curpos.Y < 486 Then color = &H5FFF0000 / 1.5: inbord = 1
quittext(1).render 270, 450, 300, 100, color

'''''###### quit or back
'äÚã
If curpos.X > 480 And curpos.X < 523 And curpos.Y > 464 And curpos.Y < 495 And mouseinp = 0 Then moseclick.PLAYnow: End
'áÇ
If curpos.X > 285 And curpos.X < 302 And curpos.Y > 455 And curpos.Y < 486 And mouseinp = 0 Then quitmenue = 0: moseclick.PLAYnow: scan = 1
'##########
End If
']]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]
If inbord = 0 Then ntime = 0
If inbord = 1 Then ntime = ntime + 1
If ntime > 100 Then ntime = 50
If ntime = 1 Then mouseoversound.PLAY
End Function
