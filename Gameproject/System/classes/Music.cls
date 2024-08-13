VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Music"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private oDMPerf As DirectMusicPerformance8  'the master performance
Private oDMLoader As DirectMusicLoader8     'helps load music into buffers
Private oDMSeg As DirectMusicSegment8       'actually stores the music to be played
Private dmParams As DMUS_AUDIOPARAMS
Public Sub create()
   
        Set DX = New DirectX8
    'we must have a root DX8 object, it is used to control/create
    'many other important objects. As shown in the next two calls:
        Set oDMPerf = DX.DirectMusicPerformanceCreate
        Set oDMLoader = DX.DirectMusicLoaderCreate
        
    'You can play around with these settings to see what happens...
        oDMPerf.InitAudio form3.hWnd, DMUS_AUDIOF_ALL, dmParams, Nothing, DMUS_APATH_DYNAMIC_STEREO, 128
        oDMPerf.SetMasterAutoDownload True
        
    '//2. Setup the messaging system
    'this next part will allow DM8 to talk back to us; with most DX programs
    'we talk to it (function calls etc...), but we can also let it call our
    'functions (the DXCALLBACK function). For those that know about this sort
    'of thing, it's just a fairly simple Callback structure.
        oDMPerf.AddNotificationType DMUS_NOTIFY_ON_SEGMENT 'relay messages about the segment
    
    '//3. Configure the loader
        oDMLoader.SetSearchDirectory App.Path & "\"
            'the path above specifies where we are loading
            'the audio file from. If you load it from a different
            'folder, "\Music\Techno\" for example, then you need to
            'change the above parameter.
            
    '//4. Load in the audio file
        Set oDMSeg = oDMLoader.LoadSegment(App.Path & "\Data\music\Midi.mid")
    'the loader will correctly parse the following file formats:
    '.WAV; .MID; .SEG; .RMI
    'IT WONT LOAD MP3 files, if you require MP3 audio then look
    'at the DirectShow8 tutorials.
            
        oDMSeg.SetStandardMidiFile
    'for .MID / .MIDI / .RMI files we must make this additional call.
    'for other formats (SEG/WAV) dont make this call.
            
    '//5. setup tempo and volume
    oDMSeg.SetRepeats 0
    oDMPerf.PlaySegmentEx oDMSeg, DMUS_SEGF_DEFAULT, 0
'\\\\\\oDMSeg.SetRepeats -1 ' -1 = play infinitely = loop
'\\\\\\oDMPerf.SetMasterTempo scrlTempo.Value / 100
'\\\\\\oDMPerf.StopEx oDMSeg, 0, DMUS_SEGF_DEFAULT
'\\\\\\oDMPerf.SetMasterVolume scrlVol.Value
End Sub

    


    


