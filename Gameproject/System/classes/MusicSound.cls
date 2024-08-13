VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MusicSOUND"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================== DirectSound8
'##############################
'#Written by : jaber al_ani   #
'#Date       : 22-Nov-2005    #
'#Time       : 01.30 PM       #
'##############################


'ÇáÇÚáÇä Úä ßÇÆä ÇáÕæÊ
Private DS As DirectSound8

'ÇáÇÚáÇä Úä ÈÝÑ ÇáÕæÊ
Private DSBuf As DirectSoundSecondaryBuffer8

'ÇáÇÚáÇä Úä ßÇÆä ÍÇáå ÇáÕæÊ
Private DSBDesc As DSBUFFERDESC
'áÚÒÝ ÇáÕæÊ ãÑå æÇÍÏå ÝÞØ ÚäÏ ÇáæÞæÝ Úáì ÒÑ


Public Function create(name As String)
ntime = 0
'ÇäÔÇÁ ßÇÆä ÇáÕæÊ ÇáÑÆíÓí
Set DS = DX.DirectSoundCreate("")

'ÖÈØ ÎÕÇÆÕ ÇáßÇÆä ÇáÑÆíÓí ÈÇä áå ÇáÇÓÈÞíå Ýí ÇáÊäÝíÐ
DS.SetCooperativeLevel form3.hWnd, DSSCL_PRIORITY

'ÌÚá ßÇÆä ÍÇáå ÇáÕæÊ ãÓÄæá Úä ÇáÊÍßã ÈÇÑÊÝÇÚ ÇáÕæÊ ÝÞØ
DSBDesc.lFlags = DSBCAPS_CTRLVOLUME

'ÊÍãíá ÇáÕæÊ ÏÇÎá ÇáÈÝÑ
Set DSBuf = DS.CreateSoundBufferFromFile(App.Path & "\Data\Sounds\" + name + ".wav", DSBDesc)


End Function
Public Function PLAY()

DSBuf.PLAY DSBPLAY_DEFAULT

End Function

Public Function PLAYnow()

DSBuf.SetCurrentPosition 0

DSBuf.PLAY DSBPLAY_DEFAULT

End Function

Public Function stopplay()
DSBuf.Stop
DSBuf.SetCurrentPosition 0
End Function



