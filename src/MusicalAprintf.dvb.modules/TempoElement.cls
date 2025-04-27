VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TempoElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public typs As String
Public pitch
Public note As AbsoluteElement
Public x As Double
Public tempo As vtempoProperties
Public tuneNumber As Double
Public totalHeightInPitches As Double
Public tempoHeightAbove As Double
    
Public Sub init(tempo As vtempoProperties, tuneNumber As Integer)
        Me.Type = "TempoElement"
        Me.tempo = tempo
        Me.tempo.Type = "tempo"  ''/ TODO-PER: this should be set earlier, in the parser, probably.
        Me.tuneNumber = tuneNumber
        '' TODO: can these two properties be merged?
        Me.totalHeightInPitches = 6
        Me.tempoHeightAbove = Me.totalHeightInPitches
        Me.pitch = undefined  '' This will be set later
        If (Me.tempo.duration And Not Me.tempo.suppressBpm) Then
            Me.note = Me.createNote(tempo, tuneNumber)
        End If
End Sub

Public Sub setX(x As Double)
        Me.x = x
End Sub

 Public Function createNote(tempo As vtempoProperties, tuneNumber As Double) As AbsoluteElement
        Dim temposcale As Double
        Dim duration As Double
        Dim abselem As AbsoluteElement
        Dim dot As Integer
        Dim flag As String
        Dim note As String
        Dim ret As Dictionary
        Dim tempoNote As RelativeElement
        Dim stem As Double
        
        
        temposcale = 0.75
        duration = tempo.duration(0)  '' TODO when multiple durations
        Set abselem = New AbsoluteElement
        abselem.init tempo, duration, 1, "tempo", tuneNumber
        '' There aren't an infinite number of note values, but we are passed a float, so just in case something is off upstream,
        '' merge all of the in between points.
        dot = 0: flag = "": note = ""
        If (duration <= 1 / 32) Then note = "noteheads.quarter": flag = "flags.u32nd": dot = 0
        ElseIf (duration <= 1 / 16) Then note = "noteheads.quarter": flag = "flags.u16th": dot = 0
        ElseIf (duration <= 3 / 32) Then note = "noteheads.quarter": flag = "flags.u16nd": dot = 1
        ElseIf (duration <= 1 / 8) Then note = "noteheads.quarter": flag = "flags.u8th": dot = 0
        ElseIf (duration <= 3 / 16) Then note = "noteheads.quarter": flag = "flags.u8th": dot = 1
        ElseIf (duration <= 1 / 4) Then note = "noteheads.quarter": dot = 0
        ElseIf (duration <= 3 / 8) Then note = "noteheads.quarter": dot = 1
        ElseIf (duration <= 1 / 2) Then note = "noteheads.half": dot = 0
        ElseIf (duration <= 3 / 4) Then note = "noteheads.half": dot = 1
        ElseIf (duration <= 1) Then note = "noteheads.whole": dot = 0
        ElseIf (duration <= 1.5) Then note = "noteheads.whole": dot = 1
        ElseIf (duration <= 2) Then note = "noteheads.dbl": dot = 0
        Else
            note = "noteheads.dbl": dot = 1
        End If
        Dim tmpOpt1 As New Dictionary
        Dim noteOpt As New oNoteHeadOptions
        tmpD1("verticalPos") = 0
        noteOpt.dir = "up"
        noteOpt.flag = flag
        noteOpt.dot = dot
        noteOpt.scale_ = temposcale
        '' This is just temporary: we'll offset the vertical positioning when we get the actual vertical spot.
        Set ret = createNoteHead(abselem, note, tmpD1, noteOpt)
        Set tempoNote = ret("notehead")
        abselem.addHead tempoNote
        Dim p1, p2, dx, width
        If (note <> "noteheads.whole" And note <> "noteheads.dbl") Then
            p1 = 1 / 3 * temposcale
            p2 = 5 * temposcale
            dx = tempoNote.dx + tempoNote.w
            width = -0.6
            Set stem = New RelativeElement
            Set tmpD1 = New Dictionary
            tmpD1("typs") = "stem"
            tmpD1("pitch2") = p2
            tmpD1("linewidth") = width
            stem.init Null, dx, 0, p1, tmpD1
            abselem.addRight stem
        End If
        Set createNote = abselem
End Function
