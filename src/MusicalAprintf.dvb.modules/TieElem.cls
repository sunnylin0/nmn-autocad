VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TieElem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public typs As String
Public isTie As Boolean
Public stemdir As String
Public dotted As Boolean
Public anchor1 As RelativeElement
Public anchor2 As RelativeElement
Public isGrace As Boolean
Public fixedY As Boolean
Public voicenumber As Integer
Public internalNotes As RelativeElementList
Public top As Double
Public bottom As Double
Public startLimitX As AbsoluteElement
Public endLimitX As AbsoluteElement
Public above As Boolean
Public startX As Double
Public endX As Double
Public startY As Double
Public endY As Double
Public elemset As iArray    'SVGPathElement[]
Public hint As Boolean
Public Sub init(options As Dictionary)
        Me.typs = "TieElem"
        ''  console.log("constructor", options.anchor1 ? options.anchor1.pitch : "N/A", options.anchor2 ? options.anchor2.pitch : "N/A", options.isTie, options.isGrace)
        Set Me.anchor1 = options("anchor1")  '' must have a .x and a .pitch, and a .parent property or be null (means starts at the "beginning" of the line - after keysig)
        Set Me.anchor2 = options("anchor2")  '' must have a .x and a .pitch property or be null (means ends at the end of the line)
        If (options("isGrace")) Then Me.isGrace = True
        If (options("fixedY")) Then Me.fixedY = True
        If (options("stemDir")) Then Me.stemdir = options("stemDir")
        If (options("voiceNumber")) Then Me.voicenumber = options("voiceNumber")
        If (options("style")) Then Me.dotted = True
        Set Me.internalNotes = New RelativeElementList
End Sub

Public Sub addInternalNote(note)
        Me.internalNotes.Push note
End Sub

Public Sub setEndAnchor(anchor2 As RelativeElement)
        ''  console.log("end", me.anchor1 ? me.anchor1.pitch : "N/A", anchor2 ? anchor2.pitch : "N/A", me.isTie, me.isGrace)
        Set Me.anchor2 = anchor2  '' must have a .x and a .pitch property or be null (means ends at the end of the line)

        '' we don't really have enough info to know what the vertical extent is yet and we won't until drawing. This will just give it enough
        '' room on either side (we don't even know if the slur will be above yet). We need to set this so that we can make sure the voice has
        '' at least enough room that the line doesn't get cut off if the tie or slur is the lowest thing.
        If Not (Me.anchor1 Is Nothing) Then
            Me.top = Math.max(Me.anchor1.pitch, Me.anchor2.pitch) + 4
            Me.bottom = Math.min(Me.anchor1.pitch, Me.anchor2.pitch) - 4
         Else
            Me.top = Me.anchor2.pitch + 4
            Me.bottom = Me.anchor2.pitch - 4
        End If
End Sub

    '' If we encounter a repeat sign, then we don't want to extend either a tie or a slur past it, so these are called to be a limit.
Public Sub setStartX(startLimitElem As AbsoluteElement)
        Set Me.startLimitX = startLimitElem
End Sub

Public Sub setEndX(endLimitElem As AbsoluteElement)
        Set Me.endLimitX = endLimitElem
End Sub

Public Sub setHint()
        Me.hint = True
End Sub

Public Sub calcTieDirection()
        '' The rules:
        '' 1) If it is in a grace note group, then the direction is always BELOW.
        '' 2) If it is in a single voice, then the direction is always OPPOSITE of the stem (or where the stem would have been in the case of whole notes.)
        '' 3) If the stem direction is forced (probably because there are two voices on the same line), then the direction is the SAME as the stem direction.

        If (Me.isGrace) Then
            Me.above = False
        ElseIf (Me.voicenumber = 0) Then
            Me.above = True
        ElseIf (Me.voicenumber > 0) Then
            Me.above = False
        Else
            Dim referencePitch As Integer
            If (Not Me.anchor1 Is Nothing) Then
                referencePitch = Me.anchor1.pitch
            ElseIf Not (Me.anchor2 Is Nothing) Then
                referencePitch = Me.anchor2.pitch
            Else
                referencePitch = 14  '' TODO-PER: this can't really happen normally. This would imply that a tie crossed over three lines, something like "C-\nz\nC"
            '' Put the arc in the opposite direction of the stem. That isn't always the pitch if one or both of the notes are beamed with something that affects its stem.
            If (Not Me.anchor1 Is Nothing) And (Not Me.anchor2 Is Nothing) Then
                    If (Me.anchor1.stemdir = "down") And (Me.anchor2.stemdir = "down") Then
                        Me.above = True
                    ElseIf (Me.anchor1.stemdir = "up") And (Me.anchor2.stemdir = "up") Then
                        Me.above = False
                    Else
                        Me.above = referencePitch >= 6
                    End If
            ElseIf (Not Me.anchor1 Is Nothing) Then
                Me.above = (Me.anchor1.stemdir = "down")
            ElseIf (Not Me.anchor2 Is Nothing) Then
                Me.above = (Me.anchor2.stemdir = "down")
            Else
                Me.above = referencePitch >= 6
        End If
End Sub

    '' From "standard music notation practice" by Music Publishers¡¦ Association:
    '' 1) Slurs are placed under the note heads if all stems go up.
    '' 2) Slurs are placed over the note heads if all stems go down.
    '' 3) If there are both up stems and down stems, prefer placing the slur over.
    '' 4) When the staff has opposite stemmed voices, all slurs should be on the stemmed side.

Public Sub calcSlurDirection()
        If (Me.isGrace) Then
            Me.above = False
        ElseIf (Me.voicenumber = 0) Then
            Me.above = True
        ElseIf (Me.voicenumber > 0) Then
            Me.above = False
        Else
            Dim hasDownStem   As Boolean
            Dim i As Integer
            Dim rel As RelativeElement
            hasDownStem = False
            
            If (Not Me.anchor1 Is Nothing) Then
                hasDownStem = IIf(Me.anchor1.stemdir = "down", True, False)
            Else
                hasDownStem = False
            End If
            If (Not Me.anchor2 Is Nothing) Then
                hasDownStem = IIf(Me.anchor2.stemdir = "down", True, False)
            Else
                hasDownStem = False
            End If
            
            For i = 0 To Me.internalNotes.Count - 1
                Set rel = Me.internalNotes(i)
                If (rel.stemdir = "down") Then
                    hasDownStem = True
                End If
            Next
            Me.above = hasDownStem
        
End Sub

Public Sub calcX(lineStartX As Double, lineEndX As Double)
        If (Not Me.anchor1 Is Nothing) Then
            Me.startX = Me.anchor1.x  '' The normal case where there is a starting element to attach to.
            If (Me.anchor1.scalex < 1) Then '' this is a grace note - don't offset the tie as much.
                Me.startX = Me.startX - 3
        ElseIf Not (Me.startLimitX Is Nothing) Then
            Me.startX = Me.startLimitX.x + Me.startLimitX.w  '' if there is no start element, but there is a repeat mark before the start of the line.
        Else
            If (Not Me.anchor2 Is Nothing) Then
                Me.startX = Me.anchor2.x - 20  '' There is no element and no repeat mark: make a small arc
            Else
                Me.startX = lineStartX  '' Don't have any guidance, so extend to beginning of line
            End If
        End If
        If (Me.anchor1 Is Nothing) And Me.dotted Then
            Me.startX = Me.startX - 3 '' The arc needs to be long enough to tell that it is dotted.
        End If
        

        If Not (Me.anchor2 Is Nothing) Then
            Me.endX = Me.anchor2.x  '' The normal case where there is a starting element to attach to.
        ElseIf Not (Me.endLimitX Is Nothing) Then
            Me.endX = Me.endLimitX.x  '' if there is no start element, but there is a repeat mark before the start of the line.
        Else
            Me.endX = lineEndX  '' There is no element and no repeat mark: extend to the beginning of the line.

End Sub

Public Sub calcTieY()
        '' If the tie comes from another line, then one or both anchors will be missing.
        If Not (Me.anchor1 Is Nothing) Then
            Me.startY = Me.anchor1.pitch
        ElseIf Not (Me.anchor2 Is Nothing) Then
            Me.startY = Me.anchor2.pitch
        Else
            Me.startY = IIf(Me.above, 14, 0)

        If Not (Me.anchor2 Is Nothing) Then
            Me.endY = Me.anchor2.pitch
        ElseIf Not (Me.anchor1 Is Nothing) Then
            Me.endY = Me.anchor1.pitch
        Else
            Me.endY = IIf(Me.above, 14, 0)
End Sub

    '' From "standard music notation practice" by Music Publishers¡¦ Association:
    '' 1) If the anchor note is down stem, the slur points to the note head.
    '' 2) If the anchor note is up stem, and the slur is over, then point to middle of stem.

Public Sub calcSlurY()
        If Not (Me.anchor1 Is Nothing) And Not (Me.anchor2 Is Nothing) Then
            If (Me.above And Me.anchor1.stemdir = "up" And Not Me.fixedY) Then
                Me.startY = (Me.anchor1.highestVert + Me.anchor1.pitch) / 2
                Me.startX = Me.startX + Me.anchor1.w / 2 '' When going to the middle of the stem, bump the line to the right a little bit to make it look right.
             Else
                Me.startY = Me.anchor1.pitch
            End If
            Dim beamInterferes As Boolean
            '' If the closing note has an up stem, and it is beamed, and it isn't the first note in the beam, then the beam will get in the way.
            beamInterferes = Me.anchor2.parent.Beam And Me.anchor2.parent.Beam.stemsUp And (Me.anchor2.parent.Beam.elems(0) = Me.anchor2.parent)
            Dim Midpoint As Double
            Midpoint = (Me.anchor2.highestVert + Me.anchor2.pitch) / 2
            If (Me.above And Me.anchor2.stemdir = "up" And Not Me.fixedY And _
                Not beamInterferes And (Midpoint < Me.startY)) Then
                Me.endY = Midpoint
                Me.endX = Me.endX + Math.Round(Me.anchor2.w / 2) '' When going to the middle of the stem, bump the line to the right a little bit to make it look right.
             Else
                Me.endY = IIf(Me.above And beamInterferes, Me.anchor2.highestVert, Me.anchor2.pitch)
            End If

         ElseIf Not (Me.anchor1 Is Nothing) Then
            Me.startY = Me.endY = Me.anchor1.pitch
         ElseIf Not (Me.anchor2 Is Nothing) Then
            Me.startY = Me.endY = Me.anchor2.pitch
         Else
            '' This is the case where the slur covers the entire line.
            '' TODO-PER: figure out where the real top and bottom of the line are.
            Me.startY = IIf(Me.above, 14, 0)
            Me.endY = IIf(Me.above, 14, 0)
        End If
End Sub

Public Sub avoidCollisionAbove()
        Dim i As Integer, maxInnerHeight As Double
        '' Double check that an interior note in the slur isn't so high that it interferes.
        If (Me.above) Then
            maxInnerHeight = -50
            For i = 0 To Me.internalNotes.Count - 1
                If (Me.internalNotes(i).highestVert > maxInnerHeight) Then
                    maxInnerHeight = Me.internalNotes(i).highestVert
                End If
            Next
            If (maxInnerHeight > Me.startY And maxInnerHeight > Me.endY) Then
                Me.startY = Me.endY = maxInnerHeight - 1
            End If
        End If
End Sub
Public Function clone() As TieElem
    Dim cope As New TieElem
    
    cope.typs = Me.typs
    cope.isTie = Me.isTie
    cope.stemdir = Me.stemdir
    cope.dotted = Me.dotted
Set cope.anchor1 = Me.anchor1
Set cope.anchor2 = Me.anchor2
    cope.isGrace = Me.isGrace
    cope.fixedY = Me.fixedY
    cope.voicenumber = Me.voicenumber
Set cope.internalNotes = Me.internalNotes
    cope.top = Me.top
    cope.bottom = Me.bottom
Set cope.startLimitX = Me.startLimitX
Set cope.endLimitX = Me.endLimitX
    cope.above = Me.above
    cope.startX = Me.startX
    cope.endX = Me.endX
    cope.startY = Me.startY
    cope.endY = Me.endY
Set cope.elemset = Me.elemset
    cope.hint = Me.hint
    
    Set clone = cope
End Function
