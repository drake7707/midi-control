VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MidiFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Any, ByVal Source As Any, ByVal length As Long)

Private Type midiHeaderRaw
    hHeader(3) As Byte 'should be 4D 54 68 64 (MThd)
    hSize(3) As Byte 'should be 00 00 00 06 (6 bytes length)
    hFormat(1) As Byte
    hNrTracks(1) As Byte
    hNrDeltaTicks(1) As Byte
End Type

Private Type trackRaw
    tHeader(3) As Byte '/* This will be 'M','T','r','k' */
    tLength(3) As Byte ' /* This will be the actual size of Data[] */
    tData() As Byte '  /* Its actual size is Data[Length] */
End Type

Private Enum midiEventEnum
    menoteoff = &H8
    meNoteOn = &H9
    meKeyAfterTouch = &HA
    meControlChange = &HB
    meProgramChange = &HC
    meChannelAfterTouch = &HD
    mePitchWeelChange = &HE '2000H is normal
    meMeta = &HF
End Enum

Private Enum midiEventMetaEnum
    memTrackSequence = &H0
    memText = &H1
    memCopyright = &H2
    memTrackName = &H3
    memTrackInstrument = &H4
    memLyric = &H5
    memMarker = &H6
    memCuePoint = &H7
    
    memMidiChannel = &H20
    memMidiPort = &H21
    
    memEndOfTrack = &H2F
    memTempo = &H51
    memTimeSig = &H58
    memKeySig = &H59
    memSeqSpecificInfo = &H7F
End Enum

Private Enum midiEventControlChangeEnum
    ccBankSelectCoarse = 0
    ccModulationWheelCoarse = 1
    ccBreathControllerCoarse = 2
    ccFootPedalCoarse = 4
    ccPortamentoTimeCoarse = 5
    ccDataEntryCoarse = 6
    ccVolumeCoarse = 7
    ccBalanceCoarse = 8
    ccPanPositionCoarse = 10
    ccExpressionCoarse = 11
    ccEffectControl1Coarse = 12
    ccEffectControl2Coarse = 13
    ccGeneralPurposeSlider1 = 16
    ccGeneralPurposeSlider2 = 17
    ccGeneralPurposeSlider3 = 18
    ccGeneralPurposeSlider4 = 19
    ccBankSelectFine = 32
    ccModulationWheelFine = 33
    ccBreathControllerFine = 34
    ccFootPedalFine = 36
    ccPotamentoTimeFine = 37
    ccDataEntryFine = 38
    ccVolumeFine = 39
    ccBalanceFine = 40
    ccPanPositionFine = 42
    ccExpressionFine = 43
    ccEffectControl1Fine = 44
    ccEffectControl2Fine = 45
    ccHoldPedalOnOff = 64
    ccPortamentoOnOff = 65
    ccSustenutoPedalOnOff = 66
    ccSoftPedalOnOff = 67
    ccLegatoPedalOnOff = 68
    ccHold2PedalOnOff = 69
    ccSoundVariation = 70
    ccSoundTimbre = 71
    ccSoundReleaseTime = 72
    ccSoundAttackTime = 73
    ccSoundBrightness = 74
    ccSoundControl6 = 75
    ccSoundControl7 = 76
    ccSoundControl8 = 77
    ccSoundControl9 = 78
    ccSoundControl10 = 79
    ccGeneralPurposeButton1OnOff = 80
    ccGeneralPurposeButton2OnOff = 81
    ccGeneralPurposeButton3OnOff = 82
    ccGeneralPurposeButton4OnOff = 83
    ccEffectsLevel = 91
    ccTremuloLevel = 92
    ccChorusLevel = 93
    ccCelesteLevel = 94
    ccPhaserLevel = 95
    ccDataButtonIncrement = 96
    ccDataButtonDecrement = 97
    ccNonRegisteredParameterFine = 98
    ccNonRegisteredParameterCoarse = 99
    ccRegisteredParameterFine = 100
    ccRegisteredParameterCoarse = 101
    ccAllSoundOff = 120
    ccAllControllersOff = 121
    ccLocalKeyboard = 122
    ccAllNotesOff = 123
    ccOmniModeOff = 124
    ccOmniModeOn = 125
    ccMonoOperation = 126
    ccPolyOperation = 127
End Enum

Private Type midiEvent
    deltaTime As Long
    type As midiEventEnum
    channel As Byte
    param1 As Long
    param2 As Long
    
    command As midiEventMetaEnum
    text As String
    tempo As Long
    
    timestamp As Long 'ms since beginning of midi (calculated with tempo changes)
End Type

Private Type track
    tHeader As String
    tLength As Long
    tData() As midiEvent
    trackName As String
    instrumentName As String
    
    instrument As Integer
    length As Long 'total ms of a track
End Type

Private Type midiHeader
    hHeader As String 'should be 4D 54 68 64 (MThd)
    hSize As Long 'should be 00 00 00 06 (6 bytes length)
    hFormat As Integer
    hNrTracks As Integer
    hNrDeltaTicks As Integer
End Type

Private Type midi
    mh As midiHeader
    tracks() As track
    
    tempo As Long
    dt As Double
    
    text As String
    copyright As String
    
    length As Long
End Type

Dim mid As midi

Dim StopPlaying As Boolean


Dim trackMuted() As Boolean

Private Type playInfo
    curTimestamp As Double
    curMEidx() As Integer
    trackEnded() As Boolean
    pause As Boolean
End Type

Dim playInfo As playInfo

Dim baseTempo As Long
Dim tempoMult As Double

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)



Private Function getNameOfMidiEvent(mee As midiEventEnum) As String
    Select Case mee
        Case meNoteOn
            getNameOfMidiEvent = "meNoteOn"
        Case menoteoff
            getNameOfMidiEvent = "meNoteOff"
        Case meKeyAfterTouch
            getNameOfMidiEvent = "meKeyAfterTouch"
        Case meControlChange
            getNameOfMidiEvent = "meControlChange"
        Case meProgramChange
            getNameOfMidiEvent = "meProgramChange"
        Case meChannelAfterTouch
            getNameOfMidiEvent = "meChannelAfterTouch"
        Case mePitchWeelChange
            getNameOfMidiEvent = "mePitchWheelChange"
        Case meMeta
            getNameOfMidiEvent = "meMeta"
        Case Else
            getNameOfMidiEvent = "unknown 0x" & Hex(mee)
    End Select
End Function

Private Function getNameOfMidiMetaEvent(mee As midiEventMetaEnum) As String
    Select Case mee
        Case memTrackSequence
            getNameOfMidiMetaEvent = "memTrackSequence"
        Case memText
            getNameOfMidiMetaEvent = "memText"
        Case memCopyright
            getNameOfMidiMetaEvent = "memCopyright"
        Case memTrackName
            getNameOfMidiMetaEvent = "memTrackName"
        Case memTrackInstrument
            getNameOfMidiMetaEvent = "memTrackInstrument"
        Case memLyric
            getNameOfMidiMetaEvent = "memLyric"
        Case memMarker
            getNameOfMidiMetaEvent = "memMarker"
        Case memCuePoint
            getNameOfMidiMetaEvent = "memCuePoint"
        Case memMidiChannel
            getNameOfMidiMetaEvent = "memMidiChannel"
        Case memMidiPort
            getNameOfMidiMetaEvent = "memMidiPort"
        Case memEndOfTrack
            getNameOfMidiMetaEvent = "memEndOfTrack"
        Case memTempo
            getNameOfMidiMetaEvent = "memTempo"
        Case memTimeSig
            getNameOfMidiMetaEvent = "memTimeSig"
        Case memKeySig
            getNameOfMidiMetaEvent = "memKeySig"
        Case memSeqSpecificInfo
            getNameOfMidiMetaEvent = "memSeqSpecificInfo"
        Case Else
            getNameOfMidiMetaEvent = "unknown meta 0x" & Hex(mee)
    End Select

End Function

Private Function getControlChangeName(cc As midiEventControlChangeEnum) As String
    Select Case cc
    Case ccBankSelectCoarse
        getControlChangeName = "ccBankSelectCoarse"
    Case ccModulationWheelCoarse
        getControlChangeName = "ccModulationWheelCoarse"
    Case ccBreathControllerCoarse
        getControlChangeName = "ccBreathControllerCoarse"
    Case ccFootPedalCoarse
        getControlChangeName = "ccFootPedalCoarse"
    Case ccPortamentoTimeCoarse
        getControlChangeName = "ccPortamentoTimeCoarse"
    Case ccDataEntryCoarse
        getControlChangeName = "ccDataEntryCoarse"
    Case ccVolumeCoarse
        getControlChangeName = "ccVolumeCoarse"
    Case ccBalanceCoarse
        getControlChangeName = "ccBalanceCoarse"
    Case ccPanPositionCoarse
        getControlChangeName = "ccPanPositionCoarse"
    Case ccExpressionCoarse
        getControlChangeName = "ccExpressionCoarse"
    Case ccEffectControl1Coarse
        getControlChangeName = "ccEffectControl1Coarse"
    Case ccEffectControl2Coarse
        getControlChangeName = "ccEffectControl2Coarse"
    Case ccGeneralPurposeSlider1
        getControlChangeName = "ccGeneralPurposeSlider1"
    Case ccGeneralPurposeSlider2
        getControlChangeName = "ccGeneralPurposeSlider2"
    Case ccGeneralPurposeSlider3
        getControlChangeName = "ccGeneralPurposeSlider3"
    Case ccGeneralPurposeSlider4
        getControlChangeName = "ccGeneralPurposeSlider4"
    Case ccBankSelectFine
        getControlChangeName = "ccBankSelectFine"
    Case ccModulationWheelFine
        getControlChangeName = "ccModulationWheelFine"
    Case ccBreathControllerFine
        getControlChangeName = "ccBreathControllerFine"
    Case ccFootPedalFine
        getControlChangeName = "ccFootPedalFine"
    Case ccPotamentoTimeFine
        getControlChangeName = "ccPotamentoTimeFine"
    Case ccDataEntryFine
        getControlChangeName = "ccDataEntryFine"
    Case ccVolumeFine
        getControlChangeName = "ccVolumeFine"
    Case ccBalanceFine
        getControlChangeName = "ccBalanceFine"
    Case ccPanPositionFine
        getControlChangeName = "ccPanPositionFine"
    Case ccExpressionFine
        getControlChangeName = "ccExpressionFine"
    Case ccEffectControl1Fine
        getControlChangeName = "ccEffectControl1Fine"
    Case ccEffectControl2Fine
        getControlChangeName = "ccEffectControl2Fine"
    Case ccHoldPedalOnOff
        getControlChangeName = "ccHoldPedalOnOff"
    Case ccPortamentoOnOff
        getControlChangeName = "ccPortamentoOnOff"
    Case ccSustenutoPedalOnOff
        getControlChangeName = "ccSustenutoPedalOnOff"
    Case ccSoftPedalOnOff
        getControlChangeName = "ccSoftPedalOnOff"
    Case ccLegatoPedalOnOff
        getControlChangeName = "ccLegatoPedalOnOff"
    Case ccHold2PedalOnOff
        getControlChangeName = "ccHold2PedalOnOff"
    Case ccSoundVariation
        getControlChangeName = "ccSoundVariation"
    Case ccSoundTimbre
        getControlChangeName = "ccSoundTimbre"
    Case ccSoundReleaseTime
        getControlChangeName = "ccSoundReleaseTime"
    Case ccSoundAttackTime
        getControlChangeName = "ccSoundAttackTime"
    Case ccSoundBrightness
        getControlChangeName = "ccSoundBrightness"
    Case ccSoundControl6
        getControlChangeName = "ccSoundControl6"
    Case ccSoundControl7
        getControlChangeName = "ccSoundControl7"
    Case ccSoundControl8
        getControlChangeName = "ccSoundControl8"
    Case ccSoundControl9
        getControlChangeName = "ccSoundControl9"
    Case ccSoundControl10
        getControlChangeName = "ccSoundControl10"
    Case ccGeneralPurposeButton1OnOff
        getControlChangeName = "ccGeneralPurposeButton1OnOff"
    Case ccGeneralPurposeButton2OnOff
        getControlChangeName = "ccGeneralPurposeButton2OnOff"
    Case ccGeneralPurposeButton3OnOff
        getControlChangeName = "ccGeneralPurposeButton3OnOff"
    Case ccGeneralPurposeButton4OnOff
        getControlChangeName = "ccGeneralPurposeButton4OnOff"
    Case ccEffectsLevel
        getControlChangeName = "ccEffectsLevel"
    Case ccTremuloLevel
        getControlChangeName = "ccTremuloLevel"
    Case ccChorusLevel
        getControlChangeName = "ccChorusLevel"
    Case ccCelesteLevel
        getControlChangeName = "ccCelesteLevel"
    Case ccPhaserLevel
        getControlChangeName = "ccPhaserLevel"
    Case ccDataButtonIncrement
        getControlChangeName = "ccDataButtonIncrement"
    Case ccDataButtonDecrement
        getControlChangeName = "ccDataButtonDecrement"
    Case ccNonRegisteredParameterFine
        getControlChangeName = "ccNonRegisteredParameterFine"
    Case ccNonRegisteredParameterCoarse
        getControlChangeName = "ccNonRegisteredParameterCoarse"
    Case ccRegisteredParameterFine
        getControlChangeName = "ccRegisteredParameterFine"
    Case ccRegisteredParameterCoarse
        getControlChangeName = "ccRegisteredParameterCoarse"
    Case ccAllSoundOff
        getControlChangeName = "ccAllSoundOff"
    Case ccAllControllersOff
        getControlChangeName = "ccAllControllersOff"
    Case ccLocalKeyboard
        getControlChangeName = "ccLocalKeyboard"
    Case ccAllNotesOff
        getControlChangeName = "ccAllNotesOff"
    Case ccOmniModeOff
        getControlChangeName = "ccOmniModeOff"
    Case ccOmniModeOn
        getControlChangeName = "ccOmniModeOn"
    Case ccMonoOperation
        getControlChangeName = "ccMonoOperation"
    Case ccPolyOperation
        getControlChangeName = "ccPolyOperation"
    Case Else
        getControlChangeName = "unknown control change 0x" & Hex(cc)
    End Select
    
End Function

Sub init(path As String)
    tempoMult = 1
    
    readMidi (path)
End Sub

Private Sub readMidi(path As String)

    Dim f As Integer
    f = FreeFile
        
    Dim m As midi
        
    clearStatus
    
    Open path For Binary As #f
        Dim mhr As midiHeaderRaw
        Get #f, , mhr
        Dim mh As midiHeader
        mh = getProperMidiHeader(mhr)
        m.mh = mh
        
        AddStatus "hHeader: " & mh.hHeader
        AddStatus "hSize: " & mh.hSize
        AddStatus "hFormat: " & mh.hFormat
        AddStatus "hNrTracks: " & mh.hNrTracks
        AddStatus "hDeltaTicks: " & mh.hNrDeltaTicks
        
        
        If mh.hHeader <> "MThd" Then
            MsgBox "Not a valid midi file!"
            Close #f
            Exit Sub
        End If
    
        Dim thr() As trackRaw
        ReDim thr(mh.hNrTracks - 1)
        
        ReDim m.tracks(mh.hNrTracks - 1)
        
        Dim i As Integer
        For i = 0 To mh.hNrTracks - 1
            Get #f, , thr(i).tHeader
            
            AddStatus "--- Track " & i & " ---"
            AddStatus " tHeader: " & ByteArrayToString(thr(i).tHeader)
            
            If ByteArrayToString(thr(i).tHeader) <> "MTrk" Then
                MsgBox "Not a valid midi track!"
                Close #f
                Exit Sub
            End If
            
            Get #f, , thr(i).tLength
            
            AddStatus " tLength: " & ByteArrayToLong(thr(i).tLength)
            
            ReDim thr(i).tData(ByteArrayToLong(thr(i).tLength) - 1)
            
            Get #f, , thr(i).tData
            
            Dim tr As track
            tr.tHeader = ByteArrayToString(thr(i).tHeader)
            tr.tLength = ByteArrayToLong(thr(i).tLength)
            
            tr.tData = getMidiEvents(m, tr, thr(i).tData)
            
            m.tracks(i) = tr
            
            Dim j As Integer
            
            For j = 0 To UBound(tr.tData)
                Call printMidiEvent(tr.tData(j))
            Next
        Next
    Close #f
    
    mid = m
    
    calculateTimeStamps
    
    ReDim trackMuted(UBound(mid.tracks))

    resetPlayInfo
End Sub

Private Function getMidiEvents(ByRef m As midi, ByRef tr As track, b() As Byte) As midiEvent()
    On Error Resume Next
    
    Dim lastStatus As midiEventEnum
    Dim lastChannel As Integer
    lastStatus = -1
    lastChannel = -1
    
    Dim cur As Long
    cur = 0
    
    Dim mevs() As midiEvent
    ReDim mevs(50)
    Dim mevcount As Integer
    mevcount = 0
    
    Dim dontadd As Boolean
    Dim emptymev As midiEvent
    
    Dim eot As Boolean
    eot = False
    Do While Not cur > UBound(b) And Not eot
        Dim mev As midiEvent
        mev = emptymev
        dontadd = False
        mev.deltaTime = ReadVarLen(b, cur)
        
        Dim tb As Byte
        tb = b(cur)
        cur = cur + 1
        
        
        If tb < &HF0 Then
            mev.type = (tb And &HF0) \ &HF
            mev.channel = tb And &HF
            
processNormal:
            Select Case mev.type
                Case meNoteOn, menoteoff, meKeyAfterTouch, _
                     meControlChange
                     'param1 = note number
                     'param2 = velocity
                    tb = b(cur)
                    cur = cur + 1
                    mev.param1 = tb
                    
                    tb = b(cur)
                    cur = cur + 1
                    mev.param2 = tb
                    
                    lastStatus = mev.type
                    lastChannel = mev.channel
                Case meProgramChange
                    'param1 = new program
                    tb = b(cur)
                    cur = cur + 1
                    mev.param1 = tb
                    
                    tr.instrument = mev.param1
                    
                    lastStatus = mev.type
                    lastChannel = mev.channel
                Case meChannelAfterTouch
                    tb = b(cur)
                    cur = cur + 1
                    mev.param1 = tb
                
                    lastStatus = mev.type
                    lastChannel = mev.channel
                Case mePitchWeelChange
                    tb = b(cur)
                    cur = cur + 1
                        
                    tb = b(cur)
                    cur = cur + 1
                    
                    lastStatus = mev.type
                    lastChannel = mev.channel
                Case Else
                    If lastStatus <> -1 And Not isStatusByte(tb) Then
                        mev.type = lastStatus
                        mev.channel = lastChannel
                         
                         'the current byte was already the first of data
                         'so to be compatible with the previous crap, just
                         'roll back 1 byte
                        cur = cur - 1
                        GoTo processNormal
                    End If
            End Select
        Else
            mev.type = meMeta
            
            tb = b(cur)
            cur = cur + 1
            
            mev.command = tb
            
            Select Case mev.command
                Case memTrackSequence
                    tb = b(cur)
                    cur = cur + 1
                                        
                    If tb <> &H2 Then
                        MsgBox "Error track corrupt"
                        Exit Function
                    End If
                    
                    
                    Dim tbs() As Byte
                    ReDim tbs(1)
                    
                    tb = b(cur)
                    cur = cur + 1
                    tbs(0) = tb
                    
                    tb = b(cur)
                    cur = cur + 1
                    tbs(1) = tb
                    
                    mev.param1 = ByteArrayToInteger(tbs)
                Case memText, memCopyright, memTrackName, _
                     memTrackInstrument, memLyric, memMarker, memCuePoint
                     
                    tb = b(cur)
                    cur = cur + 1
                    Dim i As Integer
                    For i = 0 To tb - 1
                        tb = b(cur)
                        cur = cur + 1
                        mev.text = mev.text & Chr(tb)
                    Next
                    
                    If mev.command = memText Then m.text = mev.text
                    If mev.command = memCopyright Then m.copyright = mev.text
                    If mev.command = memTrackName Then tr.trackName = mev.text
                    If mev.command = memTrackInstrument Then tr.instrumentName = mev.text
                    
                    dontadd = True
                    
                Case memTempo
                    tb = b(cur)
                    cur = cur + 1
                                        
                    If tb <> &H3 Then
                        MsgBox "Error track corrupt"
                        Exit Function
                    End If
                    
                    ReDim tbs(3)
                    tb = b(cur)
                    cur = cur + 1
                    tbs(1) = tb
                    
                    tb = b(cur)
                    cur = cur + 1
                    tbs(2) = tb
                    
                    tb = b(cur)
                    cur = cur + 1
                    tbs(3) = tb
                    
                    tbs(0) = 0
                    
                    mev.tempo = ByteArrayToLong(tbs)
                    
                    
                Case memTimeSig
                    tb = b(cur)
                    cur = cur + 1
                                    
                    If tb <> &H4 Then
                        MsgBox "Error track corrupt"
                        Exit Function
                    End If
                    tb = b(cur)
                    cur = cur + 1
                    
                    tb = b(cur)
                    cur = cur + 1
                    
                    tb = b(cur)
                    cur = cur + 1
                    
                    tb = b(cur)
                    cur = cur + 1
                    
                Case memEndOfTrack
                    tb = b(cur)
                    cur = cur + 1
                    
                    If tb <> &H0 Then
                        MsgBox "Error track corrupt"
                        Exit Function
                    End If
                                            
                    eot = True
                Case Else
                    'retrieve length
                    tb = b(cur)
                    cur = cur + 1
                    
                    For i = 0 To tb - 1
                        tb = b(cur)
                        cur = cur + 1
                        mev.text = mev.text & Chr(tb)
                    Next
            End Select
        End If
        
        If Not dontadd Then
            If mevcount >= UBound(mevs) Then
                ReDim Preserve mevs(mevcount + 50)
            End If
            mevs(mevcount) = mev
            mevcount = mevcount + 1
        End If
    Loop
    
    'shrink to fit
    ReDim Preserve mevs(mevcount - 1)
    
    getMidiEvents = mevs
End Function

Private Function isStatusByte(b As Byte) As Boolean
    isStatusByte = b And &H80
End Function

Private Function getProperMidiHeader(m As midiHeaderRaw) As midiHeader
    getProperMidiHeader.hHeader = ByteArrayToString(m.hHeader)
    getProperMidiHeader.hSize = ByteArrayToLong(m.hSize)
    getProperMidiHeader.hFormat = ByteArrayToInteger(m.hFormat)
    getProperMidiHeader.hNrTracks = ByteArrayToInteger(m.hNrTracks)
    getProperMidiHeader.hNrDeltaTicks = ByteArrayToInteger(m.hNrDeltaTicks)
End Function

Private Sub resetPlayInfo()
    Dim uTracks As Integer
    
    
    uTracks = UBound(mid.tracks)

    'clear previous info
    Dim ePlayInfo As playInfo
    playInfo = ePlayInfo
    
    ReDim playInfo.curMEidx(uTracks)
    ReDim playInfo.trackEnded(uTracks)
    'clearStatus
    
    Call Controller.timeUpdate(CLng(playInfo.curTimestamp))
End Sub

Sub play()
    StopPlaying = False
    
    'how long is 1 delta tick in ms ?
    Const MICROSECONDS_PER_MINUTE = 60000000
    
    setTempo (MICROSECONDS_PER_MINUTE / 120) '120 BPM
    
    Dim uTracks As Integer
    uTracks = UBound(mid.tracks)
    
    Dim tick As Long
    tick = GetTickCount
    
    Dim i As Integer
    Dim lasttimeupdate As Long
    
    Do While Not allTracksDone(playInfo.trackEnded) And Not StopPlaying
        If GetTickCount - tick >= mid.dt Then
            DoEvents
            For i = 0 To uTracks
                If Not playInfo.trackEnded(i) Then
                If playInfo.curTimestamp >= mid.tracks(i).tData(playInfo.curMEidx(i)).timestamp Then
                    'process curNote
                    Call processEvent(mid.tracks(i).tData(playInfo.curMEidx(i)), i)
                    
                    If playInfo.curMEidx(i) + 1 > UBound(mid.tracks(i).tData) Then
                        playInfo.trackEnded(i) = True
                    Else
                        'also process the other notes at the same time
                        Do
                            If playInfo.curMEidx(i) + 1 > UBound(mid.tracks(i).tData) Then
                                playInfo.trackEnded(i) = True
                                Exit Do
                            Else
                                playInfo.curMEidx(i) = playInfo.curMEidx(i) + 1
                                
                                If mid.tracks(i).tData(playInfo.curMEidx(i)).deltaTime = 0 Then
                                    'process curNote
                                    Call processEvent(mid.tracks(i).tData(playInfo.curMEidx(i)), i)
                                End If
                            End If
                        Loop While mid.tracks(i).tData(playInfo.curMEidx(i)).deltaTime = 0
                    End If
                End If
                End If
            Next
                        
            playInfo.curTimestamp = playInfo.curTimestamp + (GetTickCount - tick)
            
            Call Controller.setSheetCurPos(playInfo.curTimestamp / mid.dt)
            
            If GetTickCount - lasttimeupdate >= 100 Then
            
                Call Controller.timeUpdate(CLng(playInfo.curTimestamp))
                lasttimeupdate = GetTickCount
            End If
            
            tick = GetTickCount
        Else
            Dim sleeptime As Long
            sleeptime = mid.dt - (GetTickCount - tick)
            If sleeptime > 25 Then
                sleeptime = 25
                DoEvents
            End If
            If sleeptime > 0 Then
                Sleep sleeptime
            End If
        End If
    Loop

    'Dim j As Integer
    'Dim k As Integer
    'For i = 0 To uTracks
    '    For j = 0 To 86
    '        For k = 0 To 15
    '            Call Controller.NoteUp(j, i, mid.tracks(i).instrument, k)
    '        Next
    '    Next
    'Next
    For i = 0 To 15
        Call Controller.AllNotesUp(i)
        Call Controller.AllControllersOff(i)
    Next
    
End Sub

Private Sub processEvent(n As midiEvent, trackNr As Integer)
    Call printMidiEvent(n)
    Select Case n.type
        Case meMeta
            Call processMetaEvent(n, trackNr)
        Case meNoteOn
            If n.param2 = 0 Then 'velocity = 0, do noteoff
                Call Controller.NoteUp(n.param1 - 9, trackNr, mid.tracks(trackNr).instrument, CInt(n.channel))
            Else
                Call Controller.NoteDown(n.param1 - 9, trackNr, mid.tracks(trackNr).instrument, CInt(n.channel), CInt(n.param2))
            End If
        Case menoteoff
            Call Controller.NoteUp(n.param1 - 9, trackNr, mid.tracks(trackNr).instrument, CInt(n.channel))
        Case meProgramChange
            mid.tracks(trackNr).instrument = n.param1
        Case meControlChange
            Select Case n.param1
                Case ccVolumeCoarse
                    Call Controller.changeVolume(CInt(n.param2), CInt(n.channel))
                
                Case Else
                    Call Controller.doRAWControlChange(CInt(n.param1), CInt(n.param2), CInt(n.channel))
            End Select
            
    End Select
End Sub

Private Sub processMetaEvent(n As midiEvent, trackNr As Integer)
    Select Case n.command
        Case memTempo
            setTempo (n.tempo)
        End Select
    
End Sub

Private Sub setTempo(t As Long, Optional ignoreMultiplier As Boolean = False)
    If ignoreMultiplier Then
        mid.tempo = t
    Else
        If tempoMult = 0 Then
            mid.tempo = 2 ^ 30
        Else
            mid.tempo = t / tempoMult
        End If
    End If
    
    baseTempo = t
    mid.dt = (mid.tempo / 1000) / mid.mh.hNrDeltaTicks
    
    Const MICROSECONDS_PER_MINUTE = 60000000
    Controller.getBPM (MICROSECONDS_PER_MINUTE / mid.tempo)
End Sub

Private Function allTracksDone(t() As Boolean) As Boolean
    Dim i As Integer
    For i = 0 To UBound(t)
        If Not t(i) Then
            allTracksDone = False
            Exit Function
        End If
    Next
    allTracksDone = True
End Function

Private Function allTracksAtStart(idx() As Integer) As Boolean
    Dim i As Integer
    For i = 0 To UBound(idx)
        If idx(i) <> 0 Then
            allTracksAtStart = False
            Exit Function
        End If
    Next
    allTracksAtStart = True
End Function

Sub StopPlay()
    StopPlaying = True
    resetPlayInfo
End Sub

Private Sub calculateTimeStamps()
    'how long is 1 delta tick in ms ?
    Const MICROSECONDS_PER_MINUTE = 60000000
    mid.tempo = MICROSECONDS_PER_MINUTE / 120
    
    mid.dt = (mid.tempo / mid.mh.hNrDeltaTicks) / 1000
    
    Dim uTracks As Integer
    uTracks = UBound(mid.tracks)
    Dim trackEnded() As Boolean
    
    ReDim trackEnded(uTracks)
    Dim eventIdx() As Long
    ReDim eventIdx(uTracks)
    
    Dim i As Integer
    
    Dim totalTime() As Double
    Dim totalDeltaTime() As Long
    
    ReDim totalTime(uTracks)
    ReDim totalDeltaTime(uTracks)
    
    Dim chk As Boolean
    
    Call Controller.setSheetBarSize(mid.mh.hNrDeltaTicks * 4)
    
    Do While Not allTracksDone(trackEnded)
        'DoEvents
        For i = 0 To uTracks
            If Not trackEnded(i) Then
                'totalTime(i) = totalTime(i) + roundDTToNote(mid.tracks(i).tData(eventIdx(i)).deltaTime) * mid.dt
                totalDeltaTime(i) = totalDeltaTime(i) + roundDTToNote(mid.tracks(i).tData(eventIdx(i)).deltaTime)
                
                totalTime(i) = totalDeltaTime(i) * mid.dt
                
                mid.tracks(i).tData(eventIdx(i)).timestamp = totalTime(i)
                
                
                If mid.tracks(i).tData(eventIdx(i)).type = meNoteOn And _
                   mid.tracks(i).tData(eventIdx(i)).param2 <> 0 Then
                    Call Controller.addNoteToSheet(i, CLng(mid.tracks(i).tData(eventIdx(i)).param1), totalDeltaTime(i), 0)
                End If
                
                If mid.tracks(i).tData(eventIdx(i)).type = meMeta And _
                   mid.tracks(i).tData(eventIdx(i)).command = memTempo And Not chk Then
                    'make sure we don't calculate with the multiplier!
                    Call setTempo(mid.tracks(i).tData(eventIdx(i)).tempo, False)
                    chk = True
                    'Call processMetaEvent(mid.tracks(i).tData(eventIdx(i)), i)
                End If
                
                If eventIdx(i) + 1 > UBound(mid.tracks(i).tData) Then
                    trackEnded(i) = True
                    mid.tracks(i).length = totalTime(i)
                Else
                    'also process the other notes at the same time
                    Do
                        If eventIdx(i) + 1 > UBound(mid.tracks(i).tData) Then
                            trackEnded(i) = True
                            mid.tracks(i).length = totalTime(i)
                            Exit Do
                        Else
                            eventIdx(i) = eventIdx(i) + 1

                            If mid.tracks(i).tData(eventIdx(i)).deltaTime = 0 Then
                                mid.tracks(i).tData(eventIdx(i)).timestamp = totalTime(i)
                                
                                If mid.tracks(i).tData(eventIdx(i)).type = meNoteOn And _
                                   mid.tracks(i).tData(eventIdx(i)).param2 <> 0 Then
                                    Call Controller.addNoteToSheet(i, CLng(mid.tracks(i).tData(eventIdx(i)).param1), totalDeltaTime(i), 0)
                                End If
                                
                                If mid.tracks(i).tData(eventIdx(i)).type = meMeta And _
                                   mid.tracks(i).tData(eventIdx(i)).command = memTempo And Not chk Then
                                    'make sure we don't calculate with the multiplier!
                                    Call setTempo(mid.tracks(i).tData(eventIdx(i)).tempo, False)
                                    'Call processMetaEvent(mid.tracks(i).tData(eventIdx(i)), i)
                                End If
                            End If
                        End If
                    Loop While mid.tracks(i).tData(eventIdx(i)).deltaTime = 0
                    
                End If
            End If
        Next
    Loop
    
    Dim maxlength As Long
    maxlength = mid.tracks(0).length
    
    For i = 1 To uTracks
        If mid.tracks(i).length > maxlength Then
            maxlength = mid.tracks(i).length
        End If
    Next
    mid.length = maxlength
    
    Call Controller.sheetRedraw
End Sub

Function roundDTToNote(dt As Long) As Long
    '1 bar = (mid.mh.hNrDeltaTicks * 4)
    '  4 * dt  => 1 note
    '  2 * dt  => 1/2 note
    '      dt  => 1/4 note
    '  1/2*dt  => 1/8 note
    '  1/4*dt  => 1/16 note
    '  1/8*dt  => 1/32 note
    '  1/16*dt => 1/64 note
    
    ' ((mid.mh.hNrDeltaTicks * 4) / dt) should be any of those ratio's
    
    'if between 1 and (1/64 of nrdeltaticks * 4) then
    Dim newdt As Double
    
    If dt > 0 And dt < 1 / 64 * (mid.mh.hNrDeltaTicks * 4) Then
        newdt = 1 / 64 * (mid.mh.hNrDeltaTicks * 4)
    ElseIf dt > 1 / 64 * (mid.mh.hNrDeltaTicks * 4) And dt < 1 / 32 * (mid.mh.hNrDeltaTicks * 4) Then
        newdt = 1 / 32 * (mid.mh.hNrDeltaTicks * 4)
    ElseIf dt > 1 / 32 * (mid.mh.hNrDeltaTicks * 4) And dt < 1 / 16 * (mid.mh.hNrDeltaTicks * 4) Then
        newdt = 1 / 16 * (mid.mh.hNrDeltaTicks * 4)
    ElseIf dt > 1 / 16 * (mid.mh.hNrDeltaTicks * 4) And dt < 1 / 8 * (mid.mh.hNrDeltaTicks * 4) Then
        newdt = 1 / 8 * (mid.mh.hNrDeltaTicks * 4)
    ElseIf dt > 1 / 8 * (mid.mh.hNrDeltaTicks * 4) And dt < 1 / 4 * (mid.mh.hNrDeltaTicks * 4) Then
        newdt = 1 / 4 * (mid.mh.hNrDeltaTicks * 4)
    ElseIf dt > 1 / 4 * (mid.mh.hNrDeltaTicks * 4) And dt < 1 / 2 * (mid.mh.hNrDeltaTicks * 4) Then
        newdt = 1 / 2 * (mid.mh.hNrDeltaTicks * 4)
    ElseIf dt > 1 / 2 * (mid.mh.hNrDeltaTicks * 4) And dt < 1 * (mid.mh.hNrDeltaTicks * 4) Then
        newdt = 1 * (mid.mh.hNrDeltaTicks * 4)
    Else
        newdt = dt 'don't change it's longer than 1 note
    End If

    
    roundDTToNote = dt
End Function
Function seekTime(perc As Double)
    playInfo.curTimestamp = perc * mid.length
    
    'adapt all event idx to the current time stamp
    Dim i As Integer
    Dim j As Integer
    For i = 0 To UBound(mid.tracks)
        For j = 0 To UBound(mid.tracks(i).tData)
            If mid.tracks(i).tData(j).timestamp >= playInfo.curTimestamp Then
                playInfo.curMEidx(i) = j
                Exit For
            End If
        Next
    Next
        
    'all notes up
    For i = 0 To 15
        Call Controller.AllNotesUp(i)
    Next
    
End Function

Sub pause(val As Boolean)
    playInfo.pause = val

    If val Then
        StopPlaying = True
    Else
        play
    End If
End Sub

Sub stepNext()
    Dim i As Integer
    For i = 0 To 15
        Call Controller.AllNotesUp(i)
        Call Controller.AllControllersOff(i)
    Next
    
    'play the first following note
    Dim notePlayed As Boolean
    Dim tick As Long
    tick = GetTickCount
    
    
    Dim uTracks As Integer
    uTracks = UBound(mid.tracks)
        
    Do While Not allTracksDone(playInfo.trackEnded) And Not notePlayed
        DoEvents
        For i = 0 To uTracks
            If Not playInfo.trackEnded(i) And Not trackMuted(i) Then
            If playInfo.curTimestamp > mid.tracks(i).tData(playInfo.curMEidx(i)).timestamp Then
                'process curNote
                If mid.tracks(i).tData(playInfo.curMEidx(i)).type = meNoteOn And _
                   mid.tracks(i).tData(playInfo.curMEidx(i)).param2 <> 0 Then
                    notePlayed = True
                    
                    Call Controller.setSheetCurPos(mid.tracks(i).tData(playInfo.curMEidx(i)).timestamp / mid.dt)
                End If
                Call processEvent(mid.tracks(i).tData(playInfo.curMEidx(i)), i)
                
                
                If playInfo.curMEidx(i) + 1 > UBound(mid.tracks(i).tData) Then
                    playInfo.trackEnded(i) = True
                Else
                    'also process the other notes at the same time
                    Do
                        If playInfo.curMEidx(i) + 1 > UBound(mid.tracks(i).tData) Then
                            playInfo.trackEnded(i) = True
                            Exit Do
                        Else
                            playInfo.curMEidx(i) = playInfo.curMEidx(i) + 1
                            
                            If mid.tracks(i).tData(playInfo.curMEidx(i)).deltaTime = 0 Then
                                'process curNote
                                If mid.tracks(i).tData(playInfo.curMEidx(i)).type = meNoteOn And _
                                   mid.tracks(i).tData(playInfo.curMEidx(i)).param2 <> 0 Then
                                    notePlayed = True
                                    
                                    Call Controller.setSheetCurPos(mid.tracks(i).tData(playInfo.curMEidx(i)).timestamp / mid.dt)
                                End If
                                Call processEvent(mid.tracks(i).tData(playInfo.curMEidx(i)), i)
                            End If
                        End If
                    Loop While mid.tracks(i).tData(playInfo.curMEidx(i)).deltaTime = 0
                End If
            End If
            End If
        Next
                    
        playInfo.curTimestamp = playInfo.curTimestamp + mid.dt
        
        Call Controller.timeUpdate(CLng(playInfo.curTimestamp))
    Loop
    
End Sub

Sub stepPrevious()
    Dim i As Integer
    For i = 0 To 15
        Call Controller.AllNotesUp(i)
        Call Controller.AllControllersOff(i)
    Next
    
    'play the first following note
    Dim notePlayed As Boolean
    Dim tick As Long
    tick = GetTickCount
    

    Dim uTracks As Integer
    uTracks = UBound(mid.tracks)
    
  '  Dim lasttimeupdate As Long
    
    Do While Not allTracksAtStart(playInfo.curMEidx) And Not notePlayed
        'If GetTickCount - tick >= mid.dt Then
            DoEvents
            For i = 0 To uTracks
                If playInfo.curMEidx(i) <> 0 And Not trackMuted(i) Then
                If playInfo.curTimestamp < mid.tracks(i).tData(playInfo.curMEidx(i)).timestamp Then

                    'process curNote
                    If mid.tracks(i).tData(playInfo.curMEidx(i)).type = meNoteOn And _
                       mid.tracks(i).tData(playInfo.curMEidx(i)).param2 <> 0 Then
                        playInfo.trackEnded(i) = False
                        
                        Call Controller.setSheetCurPos(mid.tracks(i).tData(playInfo.curMEidx(i)).timestamp / mid.dt)
                        
                        notePlayed = True
                    End If
                    Call processEvent(mid.tracks(i).tData(playInfo.curMEidx(i)), i)
                    
                    If playInfo.curMEidx(i) - 1 < 0 Then
                        'playInfo.trackEnded(i) = True
                    Else
                        'also process the other notes at the same time
                        Do
                            If playInfo.curMEidx(i) - 1 < 0 Then
                                'playInfo.trackEnded(i) = True
                                Exit Do
                            Else
                                playInfo.curMEidx(i) = playInfo.curMEidx(i) - 1
                                
                                If mid.tracks(i).tData(playInfo.curMEidx(i)).deltaTime = 0 Then
                                    'process curNote
                                    If mid.tracks(i).tData(playInfo.curMEidx(i)).type = meNoteOn And _
                                       mid.tracks(i).tData(playInfo.curMEidx(i)).param2 <> 0 Then
                                        playInfo.trackEnded(i) = False
                                        
                                        Call Controller.setSheetCurPos(mid.tracks(i).tData(playInfo.curMEidx(i)).timestamp / mid.dt)
                                        
                                        notePlayed = True
                                    End If
                                    Call processEvent(mid.tracks(i).tData(playInfo.curMEidx(i)), i)
                                End If
                            End If
                        Loop While mid.tracks(i).tData(playInfo.curMEidx(i)).deltaTime = 0
                    End If
                End If
                End If
            Next
                        
            playInfo.curTimestamp = playInfo.curTimestamp - mid.dt
    Loop
    
    Call Controller.timeUpdate(CLng(playInfo.curTimestamp))
End Sub

Function getMidiLength() As Long
    getMidiLength = mid.length
End Function

Function getBPM() As Integer
    Const MICROSECONDS_PER_MINUTE = 60000000
    getBPM = MICROSECONDS_PER_MINUTE / mid.tempo
End Function

Sub setTempoMultiplier(mult As Double)
    tempoMult = mult
    
    Dim curperc  As Double
    curperc = playInfo.curTimestamp / mid.length
    
    calculateTimeStamps
    
    playInfo.curTimestamp = curperc * mid.length
    'change curTimestamp on same percentage
    
    Call setTempo(baseTempo)
    
End Sub

Sub muteTrack(idx As Integer, m As Boolean)
    trackMuted(idx) = m
End Sub




Private Function ByteArrayToLong(b() As Byte) As Long
    If UBound(b) <> 3 Then
        Err.Raise 10000, , "Invalid byte array"
    Else
        'ByteArrayToLong = CLng(&HFFFFFF) * b(0) + CLng(&HFFFF) * b(1) + CLng(&HFF) * b(2) + b(3)
        Dim l As Long
        Dim b2() As Byte
        ReDim b2(UBound(b))
        b2(0) = b(3)
        b2(1) = b(2)
        b2(2) = b(1)
        b2(3) = b(0)
        Call CopyMemory(VarPtr(l), VarPtr(b2(0)), 4)
               
        ByteArrayToLong = l
    End If
End Function

Private Function ByteArrayToInteger(b() As Byte) As Long
    If UBound(b) <> 1 Then
        Err.Raise 10000, , "Invalid byte array"
    Else
        ByteArrayToInteger = &HFF * b(0) + b(1)
    End If
End Function

Private Function ByteArrayToString(b() As Byte) As String
    Dim i As Integer
    For i = 0 To UBound(b)
        ByteArrayToString = ByteArrayToString & Chr(b(i))
    Next
End Function









'Initialize the variable which will hold the value. Set it to 0. We'll call this variable 'result'.
'Read the next byte of the Variable Length quantity from the MIDI file.
'Shift all of the bits in 'result' 7 places to the left. (ie, Multiply 'result' by 128).
'Logically OR 'result' with the byte that was read in, but first mask off bit #7 of the byte. (ie, AND the byte with hexadecimal 7F before you OR with 'result'. But make sure you save the original value of the byte for the test in the next step).
'Test if bit #7 of the byte is set. (ie, Is the byte AND hexadecimal 80 equal to hexadecimal 80)? If so, loop back to step #2. Otherwise, you're done, and 'result' now has the appropriate value.
Function ReadVarLen(ByRef data() As Byte, ByRef cur As Long) As Long
    Dim b As Long
    
    Dim tb As Byte
    
    Dim c As Byte
    
    'Get #f, , tb
    tb = data(cur)
    cur = cur + 1
    
    b = tb
    
    If b And &H80 Then
        b = b And &H7F
        
        Do
            'Get #f, , c
             c = data(cur)
            cur = cur + 1
            
            Dim flagged7 As Boolean
            flagged7 = c And &H80
            c = c And &H7F
            b = b * 128 + c
        Loop While flagged7
    End If
    
    ReadVarLen = b
End Function

'unsigned long ReadVarLen()
'{
'    register unsigned long value;
'    register unsigned char c;
'
'    if ( (value = getc(infile)) & 0x80 )
'    {
'       value &= 0x7F;
'       Do
'       {
'         value = (value << 7) + ((c = getc(infile)) & 0x7F);
'       } while (c & 0x80);
'    }
'
'    return(value);
'}

Private Sub printArray(b() As Byte)
    Dim i As Integer
    For i = 0 To UBound(b)
        Dim str As String
        
        Dim c As String
        c = Hex(b(i))
        If Len(c) = 1 Then
            c = "0" & c
        End If
        str = str & c & " "
    Next
    
    Debug.Print str
End Sub

Private Sub printMidiEvent(mev As midiEvent)
    With mev
        If .type = meMeta Then
            Select Case .command
                Case memText, memCopyright, memTrackName
                    AddStatus "[" & .timestamp & "] " & getNameOfMidiMetaEvent(.command) & " " & .text
                Case memTempo
                    AddStatus "[" & .timestamp & "] " & getNameOfMidiMetaEvent(.command) & " " & .tempo
                Case Else
                    AddStatus "[" & .timestamp & "] " & getNameOfMidiMetaEvent(.command)
            End Select
        Else
            Select Case .type
                Case menoteoff, meNoteOn
                    AddStatus "[" & .timestamp & " - " & .deltaTime & "] " & getNameOfMidiEvent(.type) & " " & .channel & " " & .param1 & " " & .param2
                Case meControlChange
                    AddStatus "[" & .timestamp & "] " & getNameOfMidiEvent(.type) & " " & .channel & " " & getControlChangeName(.param1)
                Case meProgramChange
                    AddStatus "[" & .timestamp & "] " & getNameOfMidiEvent(.type) & " " & .channel & " " & "'" & getInstrumentName(CInt(.param1)) & "'"
                Case Else
                    AddStatus "[" & .timestamp & "] " & getNameOfMidiEvent(.type) & " " & .channel
            End Select
        End If
    End With
End Sub

Function getTrackNames() As String()
    On Error Resume Next
    
    Dim t() As String
    ReDim t(UBound(mid.tracks))
    
    Dim i As Integer
    For i = 0 To UBound(mid.tracks)
        t(i) = mid.tracks(i).trackName & " - " & mid.tracks(i).instrumentName & " - " & getInstrumentName(mid.tracks(i).instrument)
    Next

    getTrackNames = t
End Function


