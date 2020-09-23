VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SSE_DXRecord"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim mvarSoundbuffer As DirectSoundSecondaryBuffer8
Dim mvarCaptureBuffer As DirectSoundCaptureBuffer8
Dim DsCDesc As DSCBUFFERDESC
Dim mvarDSCCAPS As Long

'most of this information is in the DirectX 8 SDK
'I also used some code from some of the submissions on
'Planet source code.
'With all of the reading I've done on this I've forgotten where I
'got most of it so if you see your stuff here thanks alot

Public Function SetRecordBuffer(Hz As Long, Channels As Integer, BITS As Integer) As Boolean
Dim DsCWaveDesc As WAVEFORMATEX
       
    DsCWaveDesc = WaveEx(Hz, Channels, BITS)
    DsCDesc.fxFormat = DsCWaveDesc
    DsCDesc.lBufferBytes = DsCWaveDesc.lAvgBytesPerSec * 20
    DsCDesc.lFlags = DSCBCAPS_WAVEMAPPED

    Set mvarCaptureBuffer = dc.CreateCaptureBuffer(DsCDesc)
    SetRecordBuffer = True
Exit Function

ErrorHandle:
    SetRecordBuffer = False

End Function

Public Sub ClearPlayBuffer()

    If Not mvarSoundbuffer Is Nothing Then
        mvarSoundbuffer.Stop
    End If
    Set mvarSoundbuffer = Nothing

End Sub

Public Sub ClearRecordBuffer()
    
    Set mvarCaptureBuffer = Nothing

End Sub

Public Sub WriteFile(FileName As String)
    
    If Not mvarSoundbuffer Is Nothing Then
        mvarSoundbuffer.SaveToFile (FileName)
    End If

End Sub

Public Sub StopPlay()
    
    If Not mvarSoundbuffer Is Nothing Then
        mvarSoundbuffer.Stop
    End If
End Sub

Public Sub Play()

    If Not mvarSoundbuffer Is Nothing Then
        mvarSoundbuffer.Play DSBPLAY_DEFAULT
    End If

End Sub

Public Sub StopRecord()

    mvarCaptureBuffer.Stop
    ConvertToSBuffer

End Sub


Public Sub Record()

    mvarCaptureBuffer.Start DSCBSTART_DEFAULT

End Sub

Public Property Get GetCaps() As Long

    GetCaps = mvarDSCCAPS

End Property

Private Sub Class_Initialize()
    'load the frm to get the hWnd value for the
    'ds.SetCooperativeLevel
    
    Load frm

    Set dx = New DirectX8
    Set dc = dx.DirectSoundCaptureCreate("")
    Set ds = dx.DirectSoundCreate("")
    ds.SetCooperativeLevel frm.hWnd, DSSCL_PRIORITY
    Dim dscaps As DSCCAPS
    dc.GetCaps dscaps
    mvarDSCCAPS = dscaps.lFormats


End Sub

Private Sub Class_Terminate()
    'need to put cleanup stuff here
    ClearPlayBuffer
    ClearRecordBuffer
    Set dc = Nothing
    Set ds = Nothing
    Set dx = Nothing
    Unload frm

End Sub

Private Function WaveEx(Hz As Long, Channels As Integer, BITS As Integer) As WAVEFORMATEX
'this is someone elses function and I can't remember where I got it from
'if you know who you are thanks alot it helped out greatly

'This is a useful little wrapper function
'for filling out wave descriptions...
    WaveEx.nFormatTag = WAVE_FORMAT_PCM
    WaveEx.nChannels = Channels
    WaveEx.lSamplesPerSec = Hz
    WaveEx.nBitsPerSample = BITS
    WaveEx.nBlockAlign = Channels * BITS / 8
    WaveEx.lAvgBytesPerSec = WaveEx.lSamplesPerSec * WaveEx.nBlockAlign
    WaveEx.nSize = 0

End Function

Private Sub ConvertToSBuffer()
'this function is also someoneelses I changed it a bit to make it
'work in this dll but not much

Dim capCURS As DSCURSORS
Dim dsb As DSBUFFERDESC
Dim ByteBuffer() As Integer
   
    mvarCaptureBuffer.GetCurrentPosition capCURS
    dsb.lBufferBytes = capCURS.lWrite + 1
    dsb.fxFormat = DsCDesc.fxFormat
    If capCURS.lWrite = 0 Then Exit Sub
    Set mvarSoundbuffer = ds.CreateSoundBuffer(dsb)
    ReDim ByteBuffer(capCURS.lWrite)
    mvarCaptureBuffer.ReadBuffer 0, capCURS.lWrite, ByteBuffer(0), DSCBLOCK_DEFAULT
    mvarSoundbuffer.WriteBuffer 0, capCURS.lWrite, ByteBuffer(0), DSBLOCK_DEFAULT



End Sub

