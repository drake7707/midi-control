Attribute VB_Name = "General"
Option Explicit


Function getInstrumentName(i As Integer)
    Select Case i
        Case 0
        getInstrumentName = "1 Acoustic Grand Piano"
        Case 1
        getInstrumentName = "2 Bright Acoustic Piano"
        Case 2
        getInstrumentName = "3 Electric Grand Piano"
        Case 3
        getInstrumentName = "4 Honky-tonk Piano"
        Case 4
        getInstrumentName = "5 Electric Piano 1"
        Case 5
        getInstrumentName = "6 Electric Piano 2"
        Case 6
        getInstrumentName = "7 Harpsichord"
        Case 7
        getInstrumentName = "8 Clavinet"
        Case 8
        getInstrumentName = "9 Celesta"
        Case 9
        getInstrumentName = "10 Glockenspiel"
        Case 10
        getInstrumentName = "11 Music Box"
        Case 11
        getInstrumentName = "12 Vibraphone"
        Case 12
        getInstrumentName = "13 Marimba"
        Case 13
        getInstrumentName = "14 Xylophone"
        Case 14
        getInstrumentName = "15 Tubular Bells"
        Case 15
        getInstrumentName = "16 Dulcimer"
        Case 16
        getInstrumentName = "17 Drawbar Organ"
        Case 17
        getInstrumentName = "18 Percussive Organ"
        Case 18
        getInstrumentName = "19 Rock Organ"
        Case 19
        getInstrumentName = "20 Church Organ"
        Case 20
        getInstrumentName = "21 Reed Organ"
        Case 21
        getInstrumentName = "22 Accordion"
        Case 22
        getInstrumentName = "23 Harmonica"
        Case 23
        getInstrumentName = "24 Tango Accordion"
        Case 24
        getInstrumentName = "25 Guitar (nylon)"
        Case 25
        getInstrumentName = "26 Acoustic Guitar (steel)"
        Case 26
        getInstrumentName = "27 Electric Guitar (jazz)"
        Case 27
        getInstrumentName = "28 Electric Guitar (clean)"
        Case 28
        getInstrumentName = "29 Electric Guitar (muted)"
        Case 29
        getInstrumentName = "30 Overdriven Guitar"
        Case 30
        getInstrumentName = "31 Distortion Guitar"
        Case 31
        getInstrumentName = "32 Guitar Harmonics"
        Case 32
        getInstrumentName = "33 Acoustic Bass"
        Case 33
        getInstrumentName = "34 Electric Bass (finger)"
        Case 34
        getInstrumentName = "35 Electric Bass (pick)"
        Case 35
        getInstrumentName = "36 Fretless Bass"
        Case 36
        getInstrumentName = "37 Slap Bass 1"
        Case 37
        getInstrumentName = "38 Slap Bass 2"
        Case 38
        getInstrumentName = "39 Synth Bass 1"
        Case 39
        getInstrumentName = "40 Synth Bass 2"
        Case 40
        getInstrumentName = "41 Violin"
        Case 41
        getInstrumentName = "42 Viola"
        Case 42
        getInstrumentName = "43 Cello"
        Case 43
        getInstrumentName = " 44 Contrabass"
        Case 44
        getInstrumentName = "45 Tremolo Strings"
        Case 45
        getInstrumentName = "46 Pizzicato Strings"
        Case 46
        getInstrumentName = "47 Orchestral Harp"
        Case 47
        getInstrumentName = "48 Timpani"
        Case 48
        getInstrumentName = "49 String Ensemble 1"
        Case 49
        getInstrumentName = "50 String Ensemble 2"
        Case 50
        getInstrumentName = "51 SynthStrings 1"
        Case 51
        getInstrumentName = "52 SynthStrings 2"
        Case 52
        getInstrumentName = "53 Choir Aahs"
        Case 53
        getInstrumentName = "54 Voice Oohs"
        Case 54
        getInstrumentName = "55 Synth Voice"
        Case 55
        getInstrumentName = "56 Orchestra Hit"
        Case 56
        getInstrumentName = "57 Trumpet"
        Case 57
        getInstrumentName = "58 Trombone"
        Case 58
        getInstrumentName = "59 Tuba"
        Case 59
        getInstrumentName = "60 Muted Trumpet"
        Case 60
        getInstrumentName = "61 French Horn"
        Case 61
        getInstrumentName = "62 Brass Section"
        Case 62
        getInstrumentName = "63 SynthBrass 1"
        Case 63
        getInstrumentName = "64 SynthBrass 2"
        Case 64
        getInstrumentName = "65 Soprano Sax"
        Case 65
        getInstrumentName = "66 Alto Sax"
        Case 66
        getInstrumentName = "67 Tenor Sax"
        Case 67
        getInstrumentName = "68 Baritone Sax"
        Case 68
        getInstrumentName = "69 Oboe"
        Case 69
        getInstrumentName = "70 English Horn"
        Case 70
        getInstrumentName = "71 Bassoon"
        Case 71
        getInstrumentName = "72 Clarinet"
        Case 72
        getInstrumentName = "73 Piccolo"
        Case 73
        getInstrumentName = "74 Flute"
        Case 74
        getInstrumentName = "75 Recorder"
        Case 75
        getInstrumentName = "76 Pan Flute"
        Case 76
        getInstrumentName = "77 Blown Bottle"
        Case 77
        getInstrumentName = "78 Shakuhachi"
        Case 78
        getInstrumentName = "79 Whistle"
        Case 79
        getInstrumentName = "80 Ocarina"
        Case 80
        getInstrumentName = "81 Lead 1(square)"
        Case 81
        getInstrumentName = "82 Lead 2 (sawtooth)"
        Case 82
        getInstrumentName = "83 Lead 3 (calliope)"
        Case 83
        getInstrumentName = "84 Lead 4 (chiff)"
        Case 84
        getInstrumentName = "85 Lead 5 (charang)"
        Case 85
        getInstrumentName = "86 Lead 6 (voice)"
        Case 86
        getInstrumentName = " 87 Lead 7 (fifths)"
        Case 87
        getInstrumentName = "88 Lead 8 (bass+lead)"
        Case 88
        getInstrumentName = "89 Pad 1 (new age)"
        Case 89
        getInstrumentName = "90 Pad 2 (warm)"
        Case 90
        getInstrumentName = "91 Pad 3 (polysynth)"
        Case 91
        getInstrumentName = "92 Pad 4 (choir)"
        Case 92
        getInstrumentName = "93 Pad 5 (bowed)"
        Case 93
        getInstrumentName = "94 Pad 6 (metallic)"
        Case 94
        getInstrumentName = "95 Pad 7 (halo)"
        Case 95
        getInstrumentName = "96 Pad 8 (sweep)"
        Case 96
        getInstrumentName = "97 FX 1 (rain)"
        Case 97
        getInstrumentName = "98 FX 2 (soundtrack)"
        Case 98
        getInstrumentName = "99 FX 3 (crystal)"
        Case 99
        getInstrumentName = "100 FX 4 (atmosphere)"
        Case 100
        getInstrumentName = "101 FX 5 (brightness)"
        Case 101
        getInstrumentName = "102 FX 6 (goblins)"
        Case 102
        getInstrumentName = "103 FX 7 (echoes)"
        Case 103
        getInstrumentName = "104 FX 8 (sci-fi)"
        Case 104
        getInstrumentName = "105 Sitar"
        Case 105
        getInstrumentName = "106 Banjo"
        Case 106
        getInstrumentName = "107 Shamisen"
        Case 107
        getInstrumentName = "108 Koto"
        Case 108
        getInstrumentName = "109 Kalimba"
        Case 109
        getInstrumentName = "110 Bag Pipe"
        Case 110
        getInstrumentName = "111 Fiddle"
        Case 111
        getInstrumentName = "112 Shanai"
        Case 112
        getInstrumentName = "113 Tinkle Bell"
        Case 113
        getInstrumentName = "114 Agogo"
        Case 114
        getInstrumentName = "115 Steel Drums"
        Case 115
        getInstrumentName = "116 Woodblock"
        Case 116
        getInstrumentName = "117 Taiko Drum"
        Case 117
        getInstrumentName = "118 Melodic Tom"
        Case 118
        getInstrumentName = "119 Synth Drum"
        Case 119
        getInstrumentName = "120 Reverse Cymbal"
        Case 120
        getInstrumentName = "121 Guitar Fret Noise"
        Case 121
        getInstrumentName = "122 Breath Noise"
        Case 122
        getInstrumentName = "123 Seashore"
        Case 123
        getInstrumentName = "124 Bird Tweet"
        Case 124
        getInstrumentName = "125 Telephone Ring"
        Case 125
        getInstrumentName = "126 Helicopter"
        Case 126
        getInstrumentName = "127 Applause"
        Case 127
        getInstrumentName = "128 Gunshot"
    End Select
End Function

Function getTimeFromSec(S As Double) As String
    Dim secRemaining As Double
    Dim timeleft As String

    secRemaining = S

    If secRemaining > 1073741824 Then
        ' x mod y must be with longs :/
        secRemaining = 1073741824
    End If
    
  '  If secRemaining >= CLng(3600) * 24 Then
  '      'few days
  '      timeleft = Format$(Int(secRemaining / (CLng(3600) * 24)), "00")
  '      secRemaining = secRemaining Mod (CLng(3600) * 24)
  '  End If

    'If secRemaining >= 3600 Then
        'few hours
        timeleft = Format$(Int(secRemaining / (3600)), "00")
        secRemaining = secRemaining Mod 3600
    'End If

    'If secRemaining >= 60 Then
        'few minutes
        timeleft = timeleft & ":" & Format$(Int(secRemaining / 60), "00")
        secRemaining = secRemaining Mod 60
    'End If

    timeleft = timeleft & ":" & Format$(secRemaining, "00")

    getTimeFromSec = timeleft
End Function

Function getFileTitle(path As String) As String
    Dim p() As String
    p = Split(path, "\")
    
    If UBound(p) >= 0 Then
        getFileTitle = p(UBound(p))
    Else
        getFileTitle = ""
    End If
End Function
