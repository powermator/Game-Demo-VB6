VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MusicMIDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================== DirectMusic ====to play .MID / .MIDI / .RMI
'##############################
'#Written by : jaber al_ani   #
'#Date       : 23-Mar-2006    #
'#Time       : 09.48 PM       #
'##############################


Private oDMPerf As DirectMusicPerformance8  'the master performance
Private oDMLoader As DirectMusicLoader8     'helps load music into buffers
Private oDMSeg As DirectMusicSegment8       'actually stores the music to be played
Private dmParams As DMUS_AUDIOPARAMS


Public Function create()
'##########################################FOR  play .MID / .MIDI / .RMI
    
Set oDMPerf = DX.DirectMusicPerformanceCreate
Set oDMLoader = DX.DirectMusicLoaderCreate
        
'You can play around with these settings to see what happens...
oDMPerf.InitAudio form3.hWnd, DMUS_AUDIOF_ALL, dmParams, Nothing, DMUS_APATH_DYNAMIC_STEREO, 128
oDMPerf.SetMasterAutoDownload True

Set oDMSeg = oDMLoader.LoadSegment(App.Path & "\Data\music\Midi.mid")
    
oDMSeg.SetRepeats -1
oDMPerf.PlaySegmentEx oDMSeg, DMUS_SEGF_DEFAULT, 0

End Function

