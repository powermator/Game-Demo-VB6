VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MusicMP3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================== DirectShow === to play MP3
'##############################
'#Written by : jaber al_ani   #
'#Date       : 23-Nov-2005    #
'#Time       : 07.11 PM       #
'##############################

Private DSAudio     As IBasicAudio         'Basic Audio Objectt
Private DSControl   As IMediaControl      'MediaControl Object
Private DSPosition  As IMediaPosition    'MediaPosition Object

Public Function create(name As String)

'// Setup a filter graph for the file
        Set DSControl = New FilgraphManager
        Call DSControl.RenderFile(App.Path + "\Data\Sounds\" + name + ".mp3")
    
'// Setup the basic audio object
        Set DSAudio = DSControl
        DSAudio.Volume = 0
        DSAudio.Balance = 0
    
'// Setup the media event and position objects
        Set DSEvent = DSControl
        Set DSPosition = DSControl
        If ObjPtr(DSPosition) Then DSPosition.Rate = 1#
        DSPosition.CurrentPosition = 0
 

End Function
Public Function switchsound(switch As Byte)

If switch = 1 Then DSControl.Run
If switch = 0 Then DSControl.Stop

End Function
