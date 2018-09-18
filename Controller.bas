Attribute VB_Name = "Controller"
Option Explicit

Dim gui As frmMain

Dim m As MidiFile

Dim trackColors() As Long

Dim muteTracks() As Boolean

Dim curTrack As Integer

Dim isPauze As Boolean
Dim started As Boolean

Sub Main()
    Set gui = New frmMain
    gui.Show

    curDevice = 0
    rc = midiOutClose(hmidi)
    rc = midiOutOpen(hmidi, curDevice, 0, 0, 0)
       
    If (rc <> 0) Then
          MsgBox "Couldn't open midi out, rc = " & rc
    End If
    
    setVolume (0.5)
    
    Dim c() As Long
    ReDim c(16)
    c(0) = RGB(255, 0, 0)
    c(1) = RGB(0, 255, 0)
    c(2) = RGB(0, 0, 255)
    c(3) = RGB(255, 255, 0)
    c(4) = RGB(255, 0, 255)
    c(5) = RGB(0, 255, 255)
    c(6) = RGB(128, 0, 0)
    c(7) = RGB(0, 128, 0)
    c(8) = RGB(0, 0, 128)
    c(9) = RGB(128, 128, 0)
    c(10) = RGB(128, 0, 128)
    c(11) = RGB(0, 128, 128)
    c(12) = RGB(128, 0, 255)
    c(13) = RGB(128, 255, 0)
    c(14) = RGB(0, 255, 128)
    c(15) = RGB(255, 0, 128)
    c(16) = RGB(128, 255, 128)
    
    
    trackColors = c
    Call gui.gpiano.setTrackColors(c)
    
    Dim i As Integer
    For i = 2 To 20
        Load gui.notetrack(i)
    Next
End Sub

Sub OpenMidi(path As String)
        
    Set m = New MidiFile
    
    clearStatus
    gui.lststatus.Visible = False
    gui.lstPermStatus.Visible = False
    Dim i As Integer
    For i = 0 To gui.notetrack.UBound
        gui.notetrack(i).clear
    Next
    gui.sldSpeed.Value = 100
    
    
    curTrack = -1
    
    Call m.init(path)
    gui.lststatus.Visible = True
    gui.lstPermStatus.Visible = True
    
    Dim t() As String
    t = m.getTrackNames
    gui.lstTracks.ListItems.clear
    
    ReDim muteTracks(UBound(t))
    
    
    For i = 0 To UBound(t)
        Dim l As ListItem
        Set l = gui.lstTracks.ListItems.Add(, "track" & i, t(i))
        l.Checked = True
        
        l.ForeColor = trackColors(i)
    Next
    
    gui.Caption = "Midi Control - " & getFileTitle(path)
End Sub

Private Function isNoteTrackLoaded(nt As notetrack) As Boolean
    On Error GoTo errh
    
    nt.doNothing
    
    isNoteTrackLoaded = True
Exit Function
errh:
    isNoteTrackLoaded = False
End Function
Sub PlayMidi()
    If m Is Nothing Then
        Exit Sub
    End If
    
    If Not started Then
        started = True
        gui.cmdPlay.Caption = "Pause"
        
        gui.lblTotalTime.Caption = getTimeFromSec(m.getMidiLength \ 1000)
        m.play
        
        Exit Sub
    End If
    
    If isPauze Then
        isPauze = False
        gui.cmdPlay.Caption = "Pause"
        
        m.pause (False)
    Else
        isPauze = True
        gui.cmdPlay.Caption = "Play"
        
        m.pause (True)
    End If
    
End Sub

Sub StopMidi()
    On Error GoTo errh
    
    started = False
    gui.cmdPlay.Caption = "Play"
    
    If Not m Is Nothing Then
        'clear all notes first
        Dim i As Integer
        For i = 0 To 15
            MidiPlayer.AllNotesUp (i)
        Next
        
        m.StopPlay
    End If
    
Exit Sub
errh:
    MsgBox "Error: " & Err.Description
End Sub

Sub Step(nxt As Boolean)
    If Not m Is Nothing Then
        If nxt Then
            Call m.stepNext
        Else
            Call m.stepPrevious
        End If
    End If
End Sub

Sub setCurrentTrack(i As Integer)
    Call gui.pianoInstrument.clear
    Call gui.lstcurInstrumentKeys.clear
    curTrack = i
    
    Dim ntcount As Integer
    
    Dim j As Integer
    For j = 0 To gui.notetrack.UBound - 1
        If j <> i And j <> i + 1 Then
            gui.notetrack(j).Visible = False
        Else
            gui.notetrack(j).Top = 7800 + CLng(ntcount) * 1080
            gui.notetrack(j).Visible = True
            gui.notetrack(j).draw
            
            ntcount = ntcount + 1
        End If
    Next
End Sub

Sub setMuteTrack(idx As Integer, mute As Boolean)
    muteTracks(idx) = mute
    Call m.muteTrack(idx, mute)
End Sub

Sub NoteDown(idx As Integer, track As Integer, program As Integer, channel As Integer, velocity As Integer)
    On Error Resume Next
    If muteTracks(track) Then
        Exit Sub
    End If
    
    Call gui.gpiano.pressKeyIdx(idx, track)
    
    If curTrack = track Then
        Call gui.pianoInstrument.pressKeyIdx(idx)
        If gui.lstcurInstrumentKeys.ListCount > 32000 Then gui.lstcurInstrumentKeys.clear
        gui.lstcurInstrumentKeys.AddItem (gui.pianoInstrument.getKeyStrFromIdx(idx))
        gui.lstcurInstrumentKeys.ListIndex = gui.lstcurInstrumentKeys.ListCount - 1
    End If
    
    
    Call MidiPlayer.NoteDown(idx, program, channel, velocity)

End Sub

Sub NoteUp(idx As Integer, track As Integer, program As Integer, channel As Integer)
    Call gui.gpiano.releaseKeyIdx(idx)
    
    If curTrack = track Then
        Call gui.pianoInstrument.releaseKeyIdx(idx)
    End If
    
    Call MidiPlayer.NoteUp(idx, program, channel)
End Sub

Sub AllNotesUp(channel As Integer)
    Call gui.gpiano.clear
    Call gui.pianoInstrument.clear
    
    Call MidiPlayer.AllNotesUp(channel)
End Sub

Sub AllControllersOff(channel As Integer)
    Call MidiPlayer.AllControllersOff(channel)
End Sub

Sub setSheetCurPos(dt As Long)
    If curTrack <> -1 Then
        Call gui.notetrack(curTrack).setCurrentPos(dt)
        Call gui.notetrack(curTrack + 1).setCurrentPos(dt)
    End If
End Sub

Sub sheetRedraw()
    If curTrack <> -1 Then
        Call gui.notetrack(curTrack).draw
        Call gui.notetrack(curTrack + 1).draw
    End If
End Sub

Sub setSheetBarSize(barl As Integer)
    Dim i As Integer
    For i = 0 To gui.notetrack.UBound
        Call gui.notetrack(i).init(barl)
    Next
End Sub

Sub addNoteToSheet(trackIdx As Integer, idx As Integer, startDt As Long, dt As Long)
    Call gui.notetrack(trackIdx).addNoteIdx(idx, startDt, dt)
End Sub

Sub timeUpdate(ms As Long)
    gui.lblTime.Caption = getTimeFromSec(ms \ 1000)
    If m.getMidiLength() <> 0 Then
        gui.sldSeek.Value = ms / m.getMidiLength() * gui.sldSeek.Max
    End If
    
End Sub

Sub seekTime(perc As Double)
    If Not m Is Nothing Then
        Call m.seekTime(perc)
    End If
End Sub

Sub setSpeed(perc As Double)
    If Not m Is Nothing Then
        Call m.setTempoMultiplier(perc)
    End If
End Sub

Sub setVolume(vol As Double)
    MidiPlayer.volume = vol
    Dim i As Integer
    For i = 0 To 15
        Call MidiPlayer.changeVolume(127, i)
    Next
    
    'MidiPlayer.vol = Int(vol * 127)
    
End Sub

Sub changeVolume(vol As Integer, channel As Integer)
    Call MidiPlayer.changeVolume(vol, channel)
End Sub

Sub doRAWControlChange(param1 As Integer, param2 As Integer, channel As Integer)
    Call MidiPlayer.RAWControlChange(param1, param2, channel)
End Sub

Sub getBPM(bpm As Integer)
    gui.lblBPM.Caption = bpm
End Sub

Sub clearStatus()
    gui.lststatus.clear
    gui.lstcurInstrumentKeys.clear
End Sub

Sub AddStatus(str As String)
    'Debug.Print str
    If gui.lststatus.ListCount > 32000 Then
        gui.lststatus.clear
    End If
    
    gui.lststatus.AddItem str
    
    If gui.lstPermStatus.ListCount < 4 Then
        gui.lstPermStatus.AddItem str
    Else
        Dim i As Integer
        For i = 1 To 3
            gui.lstPermStatus.List(i - 1) = gui.lstPermStatus.List(i)
        Next
        gui.lstPermStatus.List(3) = str
    End If
End Sub
