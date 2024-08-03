Attribute VB_Name = "AudiostationMP3Player"
Option Explicit

' /////////////////////////////////////////////////////////////////////////////////
' Module:           AudiostationMP3Player
' Description:      Adds MP3 Player functionality
'
' Date Changed:     03-08-2024
' Date Created:     04-10-2021
' Author:           Sibra-Soft - Alex van den Berg
' /////////////////////////////////////////////////////////////////////////////////

Public MediaPlaylistMode As enumPlaylistMode
Public MediaPlayMode As enumPlayMode
Public MediaPlaystate As enumPlayStates

Public MediaPlaylist As New LocalStorage

Public ShowElapsedTime As Boolean
Public CurrentMediaFilename As String
Public CurrentTrackNumber As Integer

Private Const REWIND_FORWARD_SECONDS As Long = 5
Private Const BASSFALSE As Long = 0
Private chan As Long

Public Sub Init()
    MediaPlaystate = Stopped
    MediaPlaylistMode = RepeatPlaylist
    MediaPlayMode = PlaySingleTrack
    ShowElapsedTime = True
End Sub

Public Sub Rewind()
    AdjustPosition REWIND_FORWARD_SECONDS
End Sub

Public Sub Forward()
    AdjustPosition -REWIND_FORWARD_SECONDS
End Sub

Private Sub AdjustPosition(seconds As Long)
    Dim pos As Long
    pos = BASS_ChannelBytes2Seconds(chan, BASS_ChannelGetPosition(chan, BASS_POS_BYTE))
    Call BASS_ChannelSetPosition(chan, BASS_ChannelSeconds2Bytes(chan, pos + seconds), BASS_POS_BYTE)
End Sub

Public Sub Pause()
    Call BASS_ChannelPause(chan)
    MediaPlaystate = Paused
End Sub

Public Sub StartPlay()
    AudiostationCDPlayer.StopPlay
    AudiostationMIDIPlayer.StopMidiPlayback

    PlayStateMediaMode = MP3MediaMode

    If MediaPlaystate = Paused Then
        Call BASS_ChannelPlay(chan, False)
    Else
        If CurrentTrackNumber = 0 Then CurrentTrackNumber = 1
        
        Call BASS_StreamFree(chan)
        Call BASS_MusicFree(chan)
        
        CurrentMediaFilename = MediaPlaylist.GetItemByIndex(CurrentTrackNumber, 1)
        
        chan = BASS_StreamCreateFile(BASSFALSE, StrPtr(CurrentMediaFilename), 0, 0, BASS_STREAM_AUTOFREE)
        If chan = 0 Then 
            chan = BASS_MusicLoad(BASSFALSE, CurrentMediaFilename, 0, 0, BASS_STREAM_AUTOFREE, 1)
        End If
        
        If chan = 0 Then
            ' Handle error: unable to load the media file
            MsgBox "Error loading media file: " & CurrentMediaFilename, vbCritical
            Exit Sub
        End If
        
        Call BASS_ChannelPlay(chan, True)
    End If

    MediaPlaystate = Playing
End Sub

Public Sub StopPlay()
    Call BASS_ChannelStop(chan)
    MediaPlaystate = Stopped
End Sub

Public Sub NextTrack(Optional TrackNumber As Integer, Optional Force As Boolean = False)
    If MediaPlaylist.StorageContainer.count = 0 Or CurrentTrackNumber >= MediaPlaylist.StorageContainer.count Then Exit Sub
    
    If TrackNumber > 0 Then
        CurrentTrackNumber = TrackNumber
    Else
        Dim NextTrackNumber As Integer
        Randomize
        
        If Force Then
            NextTrackNumber = CurrentTrackNumber + 1
        Else
            Select Case MediaPlayMode
                Case enumPlayMode.Shuffle
                    NextTrackNumber = Extensions.RandomNumber(1, MediaPlaylist.StorageContainer.count)
                Case enumPlayMode.PlaySingleTrack
                    Exit Sub
                Case enumPlayMode.AutoNextTrack
                    If MediaPlaylistMode = RepeatSingleTrack Then
                        NextTrackNumber = CurrentTrackNumber
                    Else
                        NextTrackNumber = CurrentTrackNumber + 1
                    End If
            End Select
        End If
        
        CurrentTrackNumber = NextTrackNumber
    End If
    
    CurrentMediaFilename = MediaPlaylist.GetItemByIndex(CurrentTrackNumber, 1)
    Call StartPlay
End Sub

Public Sub PreviousTrack()
    If MediaPlaylist.StorageContainer.count = 0 Or CurrentTrackNumber <= 1 Then Exit Sub
    
    CurrentTrackNumber = CurrentTrackNumber - 1
    CurrentMediaFilename = MediaPlaylist.GetItemByIndex(CurrentTrackNumber, 1)
    Call StartPlay
End Sub
