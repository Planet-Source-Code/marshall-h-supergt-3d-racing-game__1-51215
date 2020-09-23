Attribute VB_Name = "modSound"
'  _________________________________
' /                                 \
' |           modSound.bas          |
' \_________________________________/
'
Dim DSound As DirectSound
Public SndFreq As Long
Public EngineBuffer As DirectSoundBuffer
Public CheckBuffer As DirectSoundBuffer

Dim DsDesc As DSBUFFERDESC
Dim DsWave As WAVEFORMATEX

Sub SetupDSound()
    Set DSound = DirectX.DirectSoundCreate("")

    If Err.Number <> 0 Then
        MsgBox "Unable to continue, error creating Directsound object."
        Exit Sub
    End If

    DSound.SetCooperativeLevel frmGame.Hwnd, DSSCL_NORMAL

    DsDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
    DsWave.nFormatTag = WAVE_FORMAT_PCM 'Sound Must be PCM otherwise we get errors
    DsWave.nChannels = 2    '1= Mono, 2 = Stereo
    DsWave.lSamplesPerSec = 22050
    DsWave.nBitsPerSample = 8 '16 =16bit, 8=8bit
    DsWave.nBlockAlign = DsWave.nBitsPerSample / 8 * DsWave.nChannels
    DsWave.lAvgBytesPerSec = DsWave.lSamplesPerSec * DsWave.nBlockAlign

    Set EngineBuffer = DSound.CreateSoundBufferFromFile("engine.Wav", DsDesc, DsWave)
    SndFreq = EngineBuffer.GetFrequency

  Set CheckBuffer = CreateSound("checkpoint.wav")
End Sub

Function CreateSound(filename As String) As DirectSoundBuffer
  DsDesc.lFlags = DSBCAPS_STATIC
  
  Set CreateSound = DSound.CreateSoundBufferFromFile(filename, DsDesc, DsWave)
  If Err.Number <> 0 Then
    MsgBox "Unable to find sound file"
    MsgBox Err.Description
    End
  End If
End Function

Sub PlaySound(Sound As DirectSoundBuffer, CloseFirst As Boolean, LoopSound As Boolean)
  
  If CloseFirst Then
    Sound.Stop
    Sound.SetCurrentPosition 0
  End If
  If LoopSound Then
    Sound.Play 1
  Else
    Sound.Play 0
  End If
End Sub

