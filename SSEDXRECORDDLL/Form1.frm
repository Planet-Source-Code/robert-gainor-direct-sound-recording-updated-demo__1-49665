VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1245
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   1245
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboSoundSet 
      Height          =   315
      Left            =   1785
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   315
      Width           =   3690
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save File"
      Height          =   435
      Left            =   4305
      TabIndex        =   2
      Top             =   735
      Width           =   1170
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Play"
      Height          =   435
      Left            =   3045
      TabIndex        =   1
      Top             =   735
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Record"
      Height          =   435
      Left            =   1785
      TabIndex        =   0
      Top             =   735
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Recording Format:"
      Height          =   330
      Left            =   315
      TabIndex        =   4
      Top             =   315
      Width           =   1380
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Recorder As SSE_DXRecord

Enum CONST_WAVEFORMATFLAGS
  WAVE_FORMAT_1M08 = 1
  WAVE_FORMAT_1M16 = 4
  WAVE_FORMAT_1S08 = 2
  WAVE_FORMAT_1S16 = 8
  WAVE_FORMAT_2M08 = 16
  WAVE_FORMAT_2M16 = 64
  WAVE_FORMAT_2S08 = 32
  WAVE_FORMAT_2S16 = 128
  WAVE_FORMAT_4M08 = 256
  WAVE_FORMAT_4M16 = 1024
  WAVE_FORMAT_4S08 = 512
  WAVE_FORMAT_4S16 = 2048
End Enum



Private Sub Command1_Click()
If Command1.Caption = "Record" Then
    Command1.Caption = "Stop"
    Recorder.ClearRecordBuffer
    
    Select Case cboSoundSet.Text
    Case "44.1kHz 16-Bit Stereo"
        Recorder.SetRecordBuffer 44100, 2, 16
    Case "44.1kHz 8 Bit Stereo"
        Recorder.SetRecordBuffer 44100, 2, 8
    Case "44.1kHz 16-Bit Mono"
        Recorder.SetRecordBuffer 44100, 1, 16
    Case "44.1kHz 8-Bit Mono"
        Recorder.SetRecordBuffer 44100, 1, 8
    Case "22.05kHz 16-Bit Stereo"
        Recorder.SetRecordBuffer 22050, 2, 16
    Case "22.05kHz 8-Bit Stereo"
        Recorder.SetRecordBuffer 22050, 2, 8
    Case "22.05kHz 16-Bit Mono"
        Recorder.SetRecordBuffer 22050, 1, 16
    Case "22.05kHz 8-Bit Mono"
        Recorder.SetRecordBuffer 22050, 1, 8
    Case "11.025kHz 16-Bit Stereo"
        Recorder.SetRecordBuffer 11025, 2, 16
    Case "11.025kHz 8-Bit Stereo"
        Recorder.SetRecordBuffer 11025, 2, 8
    Case "11.025kHz 16-Bit Mono"
        Recorder.SetRecordBuffer 11025, 1, 16
    Case "11.025kHz 8-Bit Mono"
        Recorder.SetRecordBuffer 11025, 1, 8

End Select
    
    Recorder.Record
Else
    Command1.Caption = "Record"
    Recorder.StopRecord
End If

End Sub

Private Sub Command2_Click()
Recorder.Play

End Sub

Private Sub Command3_Click()
Recorder.WriteFile App.Path & "\Test.wav"

End Sub

Private Sub Form_Load()

    Set Recorder = New SSE_DXRecord
    
    
    
    If Recorder.GetCaps And WAVE_FORMAT_4S16 Then
            cboSoundSet.AddItem "44.1kHz 16-Bit Stereo"
            cboSoundSet.AddItem "44.1kHz 8 Bit Stereo"
            cboSoundSet.AddItem "44.1kHz 16-Bit Mono"
            cboSoundSet.AddItem "44.1kHz 8-Bit Mono"
            cboSoundSet.AddItem "22.05kHz 16-Bit Stereo"
            cboSoundSet.AddItem "22.05kHz 8-Bit Stereo"
            cboSoundSet.AddItem "22.05kHz 16-Bit Mono"
            cboSoundSet.AddItem "22.05kHz 8-Bit Mono"
            cboSoundSet.AddItem "11.025kHz 16-Bit Stereo"
            cboSoundSet.AddItem "11.025kHz 8-Bit Stereo"
            cboSoundSet.AddItem "11.025kHz 16-Bit Mono"
            cboSoundSet.AddItem "11.025kHz 8-Bit Mono"
            
    ElseIf Recorder.GetCaps And WAVE_FORMAT_4S08 Then
            cboSoundSet.AddItem "44.1kHz 8 Bit Stereo"
            cboSoundSet.AddItem "44.1kHz 8-Bit Mono"
            cboSoundSet.AddItem "22.05kHz 8-Bit Stereo"
            cboSoundSet.AddItem "22.05kHz 8-Bit Mono"
            cboSoundSet.AddItem "11.025kHz 8-Bit Stereo"
            cboSoundSet.AddItem "11.025kHz 8-Bit Mono"
            
    ElseIf Recorder.GetCaps And WAVE_FORMAT_4M16 Then
            cboSoundSet.AddItem "44.1kHz 16-Bit Mono"
            cboSoundSet.AddItem "44.1kHz 8-Bit Mono"
            cboSoundSet.AddItem "22.05kHz 16-Bit Mono"
            cboSoundSet.AddItem "22.05kHz 8-Bit Mono"
            cboSoundSet.AddItem "11.025kHz 16-Bit Mono"
            cboSoundSet.AddItem "11.025kHz 8-Bit Mono"
            
    ElseIf Recorder.GetCaps And WAVE_FORMAT_4M08 Then
            cboSoundSet.AddItem "44.1kHz 8-Bit Mono"
            cboSoundSet.AddItem "22.05kHz 8-Bit Mono"
            cboSoundSet.AddItem "11.025kHz 8-Bit Mono"
            
    ElseIf Recorder.GetCaps And WAVE_FORMAT_2S16 Then
            cboSoundSet.AddItem "22.05kHz 16-Bit Stereo"
            cboSoundSet.AddItem "22.05kHz 8-Bit Stereo"
            cboSoundSet.AddItem "22.05kHz 16-Bit Mono"
            cboSoundSet.AddItem "22.05kHz 8-Bit Mono"
            cboSoundSet.AddItem "11.025kHz 16-Bit Stereo"
            cboSoundSet.AddItem "11.025kHz 8-Bit Stereo"
            cboSoundSet.AddItem "11.025kHz 16-Bit Mono"
            cboSoundSet.AddItem "11.025kHz 8-Bit Mono"
            
    ElseIf Recorder.GetCaps And WAVE_FORMAT_2S08 Then
            cboSoundSet.AddItem "22.05kHz 8-Bit Stereo"
            cboSoundSet.AddItem "22.05kHz 8-Bit Mono"
            cboSoundSet.AddItem "11.025kHz 8-Bit Stereo"
            cboSoundSet.AddItem "11.025kHz 8-Bit Mono"
            
    ElseIf Recorder.GetCaps And WAVE_FORMAT_2M16 Then
            cboSoundSet.AddItem "22.05kHz 16-Bit Mono"
            cboSoundSet.AddItem "22.05kHz 8-Bit Mono"
            cboSoundSet.AddItem "11.025kHz 16-Bit Mono"
            cboSoundSet.AddItem "11.025kHz 8-Bit Mono"
            
    ElseIf Recorder.GetCaps And WAVE_FORMAT_2M08 Then
            cboSoundSet.AddItem "22.05kHz 8-Bit Mono"
            cboSoundSet.AddItem "11.025kHz 8-Bit Mono"
            
    ElseIf Recorder.GetCaps And WAVE_FORMAT_1S16 Then
            cboSoundSet.AddItem "11.025kHz 16-Bit Stereo"
            cboSoundSet.AddItem "11.025kHz 8-Bit Stereo"
            cboSoundSet.AddItem "11.025kHz 16-Bit Mono"
            cboSoundSet.AddItem "11.025kHz 8-Bit Mono"
            
    ElseIf Recorder.GetCaps And WAVE_FORMAT_1S08 Then
            cboSoundSet.AddItem "11.025kHz 8-Bit Stereo"
            cboSoundSet.AddItem "11.025kHz 8-Bit Mono"
            
    ElseIf Recorder.GetCaps And WAVE_FORMAT_1M16 Then
            cboSoundSet.AddItem "11.025kHz 16-Bit Mono"
            cboSoundSet.AddItem "11.025kHz 8-Bit Mono"
            
    ElseIf Recorder.GetCaps And WAVE_FORMAT_1M08 Then
            cboSoundSet.AddItem "11.025kHz 8-Bit Mono"
    Else
        MsgBox "Recording is not supported by your sound card; exiting.", vbApplicationModal
        Unload Me
    End If
    
    
    cboSoundSet.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Recorder = Nothing

End Sub
