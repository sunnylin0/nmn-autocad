VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BeamElem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public typs As String
Public duration As Double
Public isflat As Boolean
Public isGrace As Boolean
Public forceup As Boolean
Public forcedown As Boolean
Public elems As AbsoluteElementList
Public total As Double
Public average As Double
Public allrests As Boolean
Public stemHeight As Double
Public beams As vbeamRangeList    'BeamRange[]
Public stemsUp As Boolean
Public min As Double
Public max As Double

Public hint As Boolean
Public Count As Double
Public mainNote As AbsoluteElement
Public Sub init(stemHeight As Double, typs As String, flat, Optional firstElement As VoiceABC = Nothing)
        '' type is "grace", "up", "down", or undefined. flat is used to force flat beams, as it commonly found in the grace notes of bagpipe music.
        Me.typs = "BeamElem"
        Me.isflat = IIf(flat, True, False)
        Me.isGrace = IIf(typs = "grace", True, False)
        Me.forceup = IIf(Me.isGrace = True Or (typs = "up"), True, False)
        Me.forcedown = IIf(typs = "down", True, False)
        Set Me.elems = New AbsoluteElementList '' all the AbsoluteElements that this beam touches. It may include embedded rests.
        Me.total = 0
        Me.average = 6  '' use middle line as start for average.
        Me.allrests = True
        Me.stemHeight = stemHeight
        Set Me.beams = New vbeamRangeList   '' During the layout phase, this will become a list of the beams that need to be drawn.
        If Not (firstElement Is Nothing) Then
            If firstElement.duration > 0 Then
                Me.duration = firstElement.duration
                If (firstElement.startTriplet = True) Then
                    Me.duration = Me.duration * firstElement.tripletMultiplier
                End If
                Me.duration = CInt(Me.duration * 1000) / 1000
            End If
        Else
            Me.duration = 0
        End If
End Sub

Public Sub setHint()
        Me.hint = True
End Sub

Public Sub runningDirection(abcelem As VoiceABC)
        Dim pitch As Double
        pitch = abcelem.averagepitch
        If (pitch = 0) Then Exit Sub '' don't include elements like spacers in beams
        Me.total = CInt(Me.total + pitch)

        Me.Count = Me.Count + 1
End Sub

Public Sub Add(abselem)
        Dim pitch As Double
        pitch = abselem.abcelem.averagepitch
        If (pitch = 0) Then Exit Sub  '' don't include elements like spacers in beams
        If Not (abselem.abcelem Is Nothing) Then
            If Not (abselem.abcelem.Rest Is Nothing) Then
                Me.allrests = False
            End If
        End If
        Set abselem.Beam = Me
        Me.elems.Push abselem
        Me.total = CInt(Me.total + pitch)
        If (abselem.abcelem.minPitch < Me.min) Then Me.min = abselem.abcelem.minPitch
        
        If (abselem.abcelem.maxPitch > Me.max) Then Me.max = abselem.abcelem.maxPitch
        
End Sub

Public Sub addBeam(Beam As vbeamRange)
        Me.beams.Push Beam
End Sub

Public Sub setStemDirection()
        '' Have to figure this out before the notes are placed because placing the notes also places the decorations.
        Me.average = calcAverage(Me.total, Me.Count)
        If (Me.forceup = True) Then
            Me.stemsUp = True
        ElseIf (Me.forcedown = True) Then
            Me.stemsUp = False
        Else
            'var middleLine = 6  '' hardcoded 6 is B
            Me.stemsUp = Me.average < 6  '' true is up, false is down
        End If
        Me.Count = 0
        Me.total = 0
End Sub

Public Sub calcDir()

        Me.average = calcAverage(Me.total, Me.elems.Count)
        If (Me.forceup = True) Then
            Me.stemsUp = True
         ElseIf (Me.forcedown = True) Then
            Me.stemsUp = False
         Else
            'var middleLine = 6  '' hardcoded 6 is B
            Me.stemsUp = Me.average < 6  '' true is up, false is down
        End If
        Dim dir As String
        Dim i As Integer, j As Integer
        dir = IIf(Me.stemsUp = True, "up", "down")
        For i = 0 To Me.elems.Count - 1
            For j = 0 To Me.elems(i).heads.Count - 1
                Me.elems(i).heads(j).stemdir = dir
            Next
        Next
    
End Sub
Private Function calcAverage(total As Double, numElements As Double) As Double
    If (numElements = 0) Then calcAverage = 0
    calcAverage = total / numElements
End Function


Public Function Clone() As BeamElem
    Dim cope As New BeamElem
    cope.typs = Me.typs
    cope.duration = Me.duration
    cope.isflat = Me.isflat
    cope.isGrace = Me.isGrace
    cope.forceup = Me.forceup
    cope.forcedown = Me.forcedown
Set cope.elems = Me.elems
    cope.total = Me.total
    cope.average = Me.average
    cope.allrests = Me.allrests
    cope.stemHeight = Me.stemHeight
Set cope.beams = Me.beams
    cope.stemsUp = Me.stemsUp
    cope.min = Me.min
    cope.max = Me.max
    
    cope.hint = Me.hint
    cope.Count = Me.Count
Set cope.mainNote = Me.mainNote
    Set Clone = cope
End Function
