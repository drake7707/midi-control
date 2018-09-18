VERSION 5.00
Begin VB.UserControl NoteTrack 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11715
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   43
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   781
   Begin VB.Shape shpCur 
      BorderColor     =   &H0000FF00&
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   135
   End
End
Attribute VB_Name = "NoteTrack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type note
    octave As Integer
    key As Integer
    start As Long
    length As Long
    
        
End Type

Const NR_BARS = 4

Dim notes() As note
Dim notecount As Integer

Dim barLength As Integer

Dim xOffset As Integer

Dim curpos As Long

Private Sub UserControl_Initialize()
    init (0)
    draw
End Sub

Sub init(bl As Integer)
    barLength = bl
    ReDim notes(50)
    notecount = 0
    
    
    'barLength = 48
    'Dim i As Integer
   ' For i = 0 To 11
   '     Call addNote(3, i, 6 * i, 5)
   ' Next
    
End Sub

Sub clear()
    ReDim notes(50)
    notecount = 0
End Sub

Sub addNoteIdx(idx As Integer, dtoffset As Long, dt As Long)
    Call addNote(getOctaveFromIdx(idx), idx Mod 12, dtoffset, dt)
    
End Sub

Sub addNote(octave As Integer, key As Integer, dtoffset As Long, dt As Long)
    If notecount >= UBound(notes) Then
        ReDim Preserve notes(notecount + 50)
    End If
    
    Dim n As note
    n.octave = octave
    n.key = key
    n.length = dt
    n.start = dtoffset
    
    notes(notecount) = n
    
    notecount = notecount + 1
End Sub

Sub draw()
    UserControl.Cls
    
    drawCurPos
    
    drawLines
    drawNotes
    
    
    UserControl.Refresh
End Sub

Private Sub drawLines()
    Dim i As Integer
    Dim lcount As Integer
    For i = 0 To 7
        UserControl.Line (0, i * (UserControl.ScaleHeight - 1) / 7)-(UserControl.ScaleWidth, i * (UserControl.ScaleHeight - 1) / 7)
        
        UserControl.CurrentX = 2
        UserControl.CurrentY = i * UserControl.ScaleHeight / 7 + (UserControl.ScaleHeight / 7) / 2 - UserControl.TextHeight("ABCDEFG") / 2
        
        If UserControl.ScaleHeight > 80 Then
            UserControl.FontSize = 8
        Else
            UserControl.FontSize = 1 + (UserControl.ScaleHeight) / 10 - 1
        End If
        
        Select Case i
            Case 0
                UserControl.Print "C"
            Case 1
                UserControl.Print "D"
            Case 2
                UserControl.Print "E"
            Case 3
                UserControl.Print "F"
            Case 4
                UserControl.Print "G"
            Case 5
                UserControl.Print "A"
            Case 6
                UserControl.Print "B"
        End Select
    Next
    
    Dim basex As Integer
    basex = UserControl.TextWidth("W") + 1
    
    If barLength <> 0 Then
        
        For i = basex To UserControl.ScaleWidth
            If i Mod (UserControl.ScaleWidth) \ NR_BARS = 0 Then
                            UserControl.Line (i, 0)-(i, UserControl.ScaleHeight)
                
            End If
            
        Next
    End If
    
    UserControl.Line (0, 0)-(0, UserControl.ScaleHeight)
    UserControl.Line (UserControl.TextWidth("W"), 0)-(UserControl.TextWidth("W"), UserControl.ScaleHeight)
End Sub

Private Sub drawNotes()
    Dim i As Integer
    Dim noteoffset As Long
    Dim basex As Integer
    basex = UserControl.TextWidth("W") + 1
    
    
    Dim noteY As Integer
    
    For i = 0 To notecount - 1
        Select Case notes(i).key
        Case 9 'A
            UserControl.FillStyle = 1
            noteY = 5 * (UserControl.ScaleHeight / 7) + UserControl.ScaleHeight / 7 / 2
        Case 10 'A#
            UserControl.FillStyle = vbSolid
            noteY = 5.5 * (UserControl.ScaleHeight / 7) + UserControl.ScaleHeight / 7 / 2
        Case 11 'B
            UserControl.FillStyle = 1
            noteY = 6 * (UserControl.ScaleHeight / 7) + UserControl.ScaleHeight / 7 / 2
        Case 0 'C
            UserControl.FillStyle = 1
            noteY = 0 * (UserControl.ScaleHeight / 7) + UserControl.ScaleHeight / 7 / 2
        Case 1 'c#
            UserControl.FillStyle = vbSolid
            noteY = 0.5 * (UserControl.ScaleHeight / 7) + UserControl.ScaleHeight / 7 / 2
        Case 2 'D
            UserControl.FillStyle = 1
            noteY = 1 * (UserControl.ScaleHeight / 7) + UserControl.ScaleHeight / 7 / 2
        Case 3 'D#
            UserControl.FillStyle = vbSolid
            noteY = 1.5 * (UserControl.ScaleHeight / 7) + UserControl.ScaleHeight / 7 / 2
        Case 4 'E
            UserControl.FillStyle = 1
            noteY = 2 * (UserControl.ScaleHeight / 7) + UserControl.ScaleHeight / 7 / 2
        Case 5 'F
            UserControl.FillStyle = 1
            noteY = 3 * (UserControl.ScaleHeight / 7) + UserControl.ScaleHeight / 7 / 2
        Case 6 'F#
            UserControl.FillStyle = vbSolid
            noteY = 3.5 * (UserControl.ScaleHeight / 7) + UserControl.ScaleHeight / 7 / 2
        Case 7 'G
            UserControl.FillStyle = 1
            noteY = 4 * (UserControl.ScaleHeight / 7) + UserControl.ScaleHeight / 7 / 2
        Case 8 'G#
            UserControl.FillStyle = vbSolid
            noteY = 4.5 * (UserControl.ScaleHeight / 7) + UserControl.ScaleHeight / 7 / 2
        End Select
        
        'UserControl.ScaleHeight / 7 / 4 is note width
        'usercontrol.ScaleWidth / 3 = 1 barLength in px
        '(notes(i).start / barLength) = offset of note in barlength
        '(notes(i).start / barLength) * (usercontrol.ScaleWidth / NR_BARS) = offset of note in pixels (1px widht note)
        noteoffset = (notes(i).start / barLength) * (UserControl.ScaleWidth / NR_BARS)
        
        Dim cx As Long
        Dim cy As Long
        cx = basex - xOffset + noteoffset + UserControl.ScaleHeight / 7 / 4
        cy = noteY
        
        Dim r1 As Integer
        Dim r2 As Integer
        Dim g1 As Integer
        Dim g2 As Integer
        Dim b1 As Integer
        Dim b2 As Integer
        
        Dim oldcolor As Long
        oldcolor = UserControl.ForeColor
        
        If notes(i).octave Mod 3 = 0 Then
            UserControl.ForeColor = RGB(128, 0, 0)
        ElseIf notes(i).octave Mod 3 = 1 Then
            UserControl.ForeColor = RGB(0, 128, 0)
        ElseIf notes(i).octave Mod 3 = 2 Then
            UserControl.ForeColor = RGB(0, 0, 128)
        End If
        'r1 = 0
        'r2 = 255
        'g1 = 0
        'g2 = 0
        'b1 = 255
        'b2 = 0
        'Dim a As Single
        'a = notes(i).octave / 8
        
        'UserControl.ForeColor = RGB(r1 * a + r2 * (1 - a), g1 * a + g2 * (1 - a), b1 * a - b2 * (1 - a))
        
        If cx - UserControl.ScaleHeight / 7 / 4 > 0 And cx + UserControl.ScaleHeight / 7 / 4 < UserControl.ScaleWidth Then
            UserControl.Circle (cx, cy), UserControl.ScaleHeight / 7 / 4
        End If
        UserControl.ForeColor = oldcolor
        
        
    Next
End Sub

Private Function drawCurPos() As Boolean
    On Error Resume Next
    Dim basex As Integer
    basex = UserControl.TextWidth("W") + 1

    shpCur.Height = UserControl.ScaleHeight
    shpCur.Width = UserControl.ScaleHeight / 7
    If barLength = 0 Then
        shpCur.Visible = False
    Else
        shpCur.Visible = True
        shpCur.Left = basex - xOffset + (curpos / barLength) * (UserControl.ScaleWidth / NR_BARS)
        
        
        If shpCur.Left < 0 Then
            xOffset = (UserControl.ScaleWidth / 2) / barLength + (curpos / barLength) * (UserControl.ScaleWidth / NR_BARS)
            drawCurPos = True
        ElseIf shpCur.Left > UserControl.ScaleWidth Then
            xOffset = (curpos / barLength) * (UserControl.ScaleWidth / NR_BARS)
            drawCurPos = True
        End If
    End If
End Function

Sub setCurrentPos(startDt As Long)
    curpos = startDt
    
    Dim x As Boolean
    x = drawCurPos
    
    If x Then draw
End Sub

Private Sub UserControl_Resize()
    draw
End Sub

Private Function getToneFromIdx(i As Integer) As String
    Select Case i Mod 12
        Case 0
            getToneFromIdx = "A"
        Case 1
            getToneFromIdx = "A#"
        Case 2
            getToneFromIdx = "B"
        Case 3
            getToneFromIdx = "C"
        Case 4
            getToneFromIdx = "C#"
        Case 5
            getToneFromIdx = "D"
        Case 6
            getToneFromIdx = "D#"
        Case 7
            getToneFromIdx = "E"
        Case 8
            getToneFromIdx = "F"
        Case 9
            getToneFromIdx = "F#"
        Case 10
            getToneFromIdx = "G"
        Case 11
            getToneFromIdx = "G#"
    End Select
    
End Function


Private Function getOctaveFromIdx(i As Integer) As Integer
    If i >= 0 And i <= 2 Then
        getOctaveFromIdx = 0
    Else
        getOctaveFromIdx = 1 + (i - 3) \ 12
    End If
End Function

Sub doNothing()

End Sub
