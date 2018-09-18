VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Midi Control"
   ClientHeight    =   10530
   ClientLeft      =   285
   ClientTop       =   585
   ClientWidth     =   12525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10530
   ScaleWidth      =   12525
   StartUpPosition =   3  'Windows Default
   Begin MidiControl.NoteTrack notetrack 
      Height          =   975
      Index           =   1
      Left            =   0
      TabIndex        =   24
      Top             =   8880
      Width           =   12495
      _ExtentX        =   18865
      _ExtentY        =   1931
   End
   Begin MidiControl.NoteTrack notetrack 
      Height          =   975
      Index           =   0
      Left            =   0
      TabIndex        =   23
      Top             =   7800
      Width           =   12495
      _ExtentX        =   18865
      _ExtentY        =   1931
   End
   Begin VB.CommandButton cmdStepPrevious 
      Caption         =   "Step <-"
      Height          =   375
      Left            =   4440
      TabIndex        =   22
      Top             =   6600
      Width           =   975
   End
   Begin VB.CommandButton cmdStepNext 
      Caption         =   "Step ->"
      Height          =   375
      Left            =   4440
      TabIndex        =   21
      Top             =   7080
      Width           =   975
   End
   Begin VB.ListBox lstPermStatus 
      Height          =   1035
      Left            =   7200
      TabIndex        =   20
      Top             =   5280
      Width           =   5295
   End
   Begin VB.Frame Frame2 
      Caption         =   "Speed"
      Height          =   615
      Left            =   6960
      TabIndex        =   17
      Top             =   7065
      Width           =   3615
      Begin MSComctlLib.Slider sldSpeed 
         Height          =   300
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   529
         _Version        =   393216
         LargeChange     =   50
         SmallChange     =   10
         Max             =   500
         SelStart        =   100
         TickFrequency   =   50
         Value           =   100
      End
      Begin VB.Label lblBPM 
         Height          =   255
         Left            =   2880
         TabIndex        =   19
         Top             =   240
         Width           =   615
      End
   End
   Begin MSComctlLib.Slider sldSeek 
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   10080
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   661
      _Version        =   393216
      LargeChange     =   10
      SmallChange     =   5
      Max             =   100
      TickFrequency   =   2
   End
   Begin VB.CommandButton cmdSolo 
      Caption         =   "Solo"
      Height          =   375
      Left            =   4440
      TabIndex        =   13
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdSelectNone 
      Caption         =   "Select None"
      Height          =   375
      Left            =   4440
      TabIndex        =   12
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "Select All"
      Height          =   375
      Left            =   4440
      TabIndex        =   11
      Top             =   2760
      Width           =   975
   End
   Begin VB.ListBox lststatus 
      Height          =   2595
      Left            =   7200
      TabIndex        =   10
      Top             =   2640
      Width           =   5295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Volume"
      Height          =   615
      Left            =   6960
      TabIndex        =   8
      Top             =   6450
      Width           =   3615
      Begin MSComctlLib.Slider sldVolume 
         Height          =   300
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   529
         _Version        =   393216
         LargeChange     =   10
         SmallChange     =   5
         Max             =   100
         SelStart        =   50
         TickFrequency   =   20
         Value           =   50
      End
   End
   Begin MSComctlLib.ListView lstTracks 
      Height          =   4935
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   8705
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   7056
      EndProperty
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      Top             =   7200
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   120
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Open midi"
      Filter          =   "*.mid|*.mid"
   End
   Begin MidiControl.piano gpiano 
      Height          =   1215
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   2143
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Height          =   375
      Left            =   5640
      TabIndex        =   3
      Top             =   6840
      Width           =   1215
   End
   Begin VB.ListBox lstcurInstrumentKeys 
      Height          =   3765
      Left            =   5520
      TabIndex        =   2
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   6480
      Width           =   1215
   End
   Begin MidiControl.piano pianoInstrument 
      Height          =   1215
      Left            =   0
      TabIndex        =   6
      Top             =   1320
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   2143
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   12480
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   12480
      Y1              =   9960
      Y2              =   9960
   End
   Begin VB.Label lblTotalTime 
      Caption         =   "00:00:00"
      Height          =   255
      Left            =   11280
      TabIndex        =   16
      Top             =   10080
      Width           =   1215
   End
   Begin VB.Label lblTime 
      Caption         =   "00:00:00"
      Height          =   255
      Left            =   9960
      TabIndex        =   15
      Top             =   10080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   11880
      TabIndex        =   0
      Top             =   7080
      Width           =   615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m As MidiFile

Private Sub cmdOpen_Click()
    Controller.StopMidi
    On Error GoTo errh
    cd.ShowOpen
    
   ' Set m = New MidiFile
   ' Call m.init("C:\Documents and Settings\Drake7707\My Documents\radicaldreamers-gpiano.mid", lstcurInstrumentKeys, gpiano)
    'Call m.init("C:\Documents and Settings\Drake7707\My Documents\SC4m02.mid", lstcurInstrumentKeys, gpiano)
    'Call m.init("E:\terranigma.mid", lstcurInstrumentKeys, gpiano)
    'Peril_gpiano.mid
    'Call m.init("C:\Documents and Settings\Drake7707\My Documents\Peril_gpiano.mid", lstcurInstrumentKeys, gpiano)
    'Call m.init(cd.FileName, lstcurInstrumentKeys)
    Call Controller.OpenMidi(cd.FileName)
    Label1.Caption = "Load complete"
Exit Sub
errh:
End Sub


Private Sub cmdPlay_Click()
'    Controller.StopMidi
    
    'If Not m Is Nothing Then
    '    Call m.play 'Midi(1)
    'End If
    Call Controller.PlayMidi
    
End Sub

Private Sub cmdStepNext_Click()
    Call Controller.Step(True)
    
End Sub

Private Sub cmdStepPrevious_Click()
    Call Controller.Step(False)
End Sub

Private Sub cmdStop_Click()
    'Call m.StopPlay
    Call Controller.StopMidi
    
End Sub

Private Sub cmdSelectAll_Click()
    Dim i As Integer
    For i = 1 To lstTracks.ListItems.Count
        lstTracks.ListItems.Item(i).Checked = True
        Call lstTracks_ItemCheck(lstTracks.ListItems.Item(i))
    Next
End Sub

Private Sub cmdSelectNone_Click()
    Dim i As Integer
    For i = 1 To lstTracks.ListItems.Count
        lstTracks.ListItems.Item(i).Checked = False
        Call lstTracks_ItemCheck(lstTracks.ListItems.Item(i))
    Next
End Sub

Private Sub cmdSolo_Click()
    If lstTracks.SelectedItem Is Nothing Then
        Exit Sub
    End If
    
    Dim i As Integer
    For i = 1 To lstTracks.ListItems.Count
        If i = lstTracks.SelectedItem.Index Then
            lstTracks.ListItems.Item(i).Checked = True
        Else
            lstTracks.ListItems.Item(i).Checked = False
        End If
        Call lstTracks_ItemCheck(lstTracks.ListItems.Item(i))
    Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call Controller.StopMidi
    
    DoEvents

    rc = midiOutClose(hmidi)
    
    DoEvents
    
    Unload Me
    End
End Sub



Private Sub gpiano_NoteDown(octave As Integer, key As String, idx As Integer)
    Label1.Caption = key & octave
    
    Call MidiPlayer.NoteDown(idx, 0, 15, 127)
End Sub

Private Sub gpiano_NoteUp(octave As Integer, key As String, idx As Integer)
    Label1.Caption = key & octave
    
    Call MidiPlayer.NoteUp(idx, 0, 15)
End Sub


Private Sub lstTracks_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call Controller.setMuteTrack(Item.Index - 1, Not Item.Checked)
End Sub

Private Sub lstTracks_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call Controller.setCurrentTrack(Item.Index - 1)
End Sub

Private Sub pianoInstrument_NoteDown(octave As Integer, key As String, idx As Integer)
    Call MidiPlayer.NoteDown(idx, 0, 15, 127)
End Sub

Private Sub pianoInstrument_NoteUp(octave As Integer, key As String, idx As Integer)
    Call MidiPlayer.NoteUp(idx, 0, 15)
End Sub

Private Sub sldSeek_Click()
    
    Controller.seekTime (sldSeek.Value / sldSeek.Max)
End Sub

Private Sub sldSeek_Scroll()
    
    Controller.seekTime (sldSeek.Value / sldSeek.Max)
    
End Sub

Private Sub sldSpeed_Click()
    sldSpeed.Value = (sldSpeed.Value \ 10) * 10
    
    Controller.setSpeed (sldSpeed.Value / 100)
End Sub

Private Sub sldSpeed_Scroll()
    sldSpeed.Value = (sldSpeed.Value \ 10) * 10
    sldSpeed.ToolTipText = sldSpeed.Value & "%"
    
    Controller.setSpeed (sldSpeed.Value / 100)
End Sub

Private Sub sldVolume_Click()
    Controller.setVolume (sldVolume.Value / sldVolume.Max)
End Sub

Private Sub sldVolume_Scroll()
    Controller.setVolume (sldVolume.Value / sldVolume.Max)
End Sub
