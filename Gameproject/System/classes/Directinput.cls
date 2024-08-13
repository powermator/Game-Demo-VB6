VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Directinput"
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


' ÇáÇÚáÇä Úä ßÇÆä ÇáÇÏÎÇá ÇáÑÆíÓí
Private DI         As DirectInput8

'ÇáÇÚáÇä Úä ÌåÇÒ ÇáãÇæÓ
Private Mouse As DirectInputDevice8


'ÇáÇÚáÇä Úä ÌåÇÒ ÇáßíÈæÑÏ
Private Keyboard As DirectInputDevice8
'ÇáÇÚáÇä Úä ßÇÆä ÍÇáå ÇáãÝÇÊíÍ ááßíÈæÑÏ
Private DIState As DIKEYBOARDSTATE

Private Sub Class_Initialize()
'ÇäÔÇÁ ßÇÆä ÇáÇÏÎÇá ÇáÑÆíÓí
Set DI = DX.DirectInputCreate()

'ÇäÔÇÁ ßÇÆä ÇÏÎÇá ÇáãÇæÓ
Set Mouse = DI.CreateDevice("GUID_SysMouse")
'ÖÈØ ÎÕÇÆÕ ÇáãÇæÓ
With Mouse
.SetCommonDataFormat DIFORMAT_MOUSE
.SetCooperativeLevel form3.hWnd, DISCL_NONEXCLUSIVE Or DISCL_BACKGROUND
.Acquire
End With


'ÇäÔÇÁ ßÇÆä ÇÏÎÇá ÇáßíÈæÑÏ
Set Keyboard = DI.CreateDevice("GUID_SysKeyboard")
'ÖÈØ ÎÕÇÆÕ ÇáßíÈæÑÏ
With Keyboard
.SetCommonDataFormat DIFORMAT_KEYBOARD
.SetCooperativeLevel form3.hWnd, DISCL_NONEXCLUSIVE Or DISCL_BACKGROUND
.Acquire
End With

End Sub

Public Function CheckKey()
'######################################## Mousw
mouseinp = 5
'ÇÎÐ ÍÇáå ÇáãÇæÓ Ýí ßÇÆä ÇáÍÇáå
Mouse.GetDeviceStateMouse MouseState

If MouseState.Buttons(0) <> 0 Then mouseinp = 0
If MouseState.Buttons(1) <> 0 Then mouseinp = 1
If MouseState.Buttons(2) <> 0 Then mouseinp = 2
If MouseState.Buttons(3) <> 0 Then mouseinp = 3

'ÅÐÇ ÍÑßäÇ ÇáãÇæÓ íãíäÇ Çæ íÓÇÑÇ
If MouseState.lX > 1 Or MouseState.lX < -1 Or MouseState.lY > 1 Or MouseState.lY < -1 Then moving = 1

'########################################## KeyBoard

Keyboard.GetDeviceStateKeyboard DIState
If DIState.Key(DIK_ESCAPE) Then End
End Function
