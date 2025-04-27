Attribute VB_Name = "LayoutModule"
Option Explicit


Function layout(renderer As RendererModule, abctune As TuneData, width As Double, space As Double) As Double
    Dim i As Integer
    Dim j As Integer
    
    Dim abcLine As TuneLine
    '' Adjust the x-coordinates to their absolute positions
    '' �N x ���нվ��䵴���m
    Dim maxWidth  As Double
    maxWidth = width
    For i = 0 To abctune.lines.Count - 1
        abcLine = abctune.lines(i)
        If Not (abcLine.Staff Is Nothing) Then
            setXSpacing renderer, width, space, abcLine.staffGroup, abctune.formatting, i = abctune.lines.Count - 1, False
            If (abcLine.staffGroup.w > maxWidth) Then maxWidth = abcLine.staffGroup.w
        End If
    Next

    '' Layout the beams and add the stems to the beamed notes.
    '' �G�����ñN�D�F�K�[���筵�Ť��C
    For i = 0 To abctune.lines.clount - 1
        Set abcLine = abctune.lines(i)
        If Not (abcLine.staffGroup Is Nothing) Then
            If Not (abcLine.staffGroup.voices Is Nothing) Then
                For j = 0 To abcLine.staffGroup.voices.Count
                    layoutVoice (abcLine.staffGroup.voices(j))
                Next
                setUpperAndLowerElements renderer, abcLine.staffGroup
            End If
        End If
    Next

    '' Set the staff spacing
    '' TODO-PER: we should have been able to do this by the time we called setUpperAndLowerElements, but for some reason the "bottom" element seems to be set as a side effect of setting the X spacing.
    '' �]�m���u�ж��Z
    '' TODO-PER�G��ڭ̽ե� setUpperAndLowerElements �ɡA�ڭ����ӯ������o�@�I�A
    ''���ѩ�Y�ح�]�A���������������G�Q�]�m���]�m X ���Z���Ƨ@�ΡC
    For i = 0 To abctune.lines.Count - 1
        Set abcLine = abctune.lines(i)
        If Not (abcLine.staffGroup Is Nothing) Then
            abcLine.staffGroup.setHeight
        End If
    Next
    layout = maxWidth
End Function
'/**
' * Do the x-axis positioning for a single line (a group of related staffs)
' * ���u�]�@�լ����H���^�i��x�b�w��
' */
Public Sub setXSpacing(renderer As RendererModule, width As Double, space As Double, staffGroup As StaffGroupElement, formatting As vFormatting, isLastLine As Boolean, isDebug As Boolean)
    Dim leftEdge As Double
    Dim newspace As Double
    leftEdge = getLeftEdgeOfStaff(renderer, staffGroup.getTextSize, staffGroup.voices, staffGroup.brace, staffGroup.bracket)
    newspace = space
    '' TODO-PER: shouldn't need multiple passes, but each pass gets it closer to the right spacing. (Only affects long lines: normal lines break out of this loop quickly.)
    '' TODO-PER�G���ݭn�h���q�L�A���C���q�L���|�Ϩ�󱵪񥿽T�����Z�C �]�ȼv�T����G���q��ܧִN�|�������`���C�^
    Dim it As Integer
    Dim ret
    For it = 0 To 7
        setret = layoutStaffGroup(newspace, renderer, isDebug, staffGroup, leftEdge)
        newspace = calcHorizontalSpacing(isLastLine, formatting.stretchLast, width + renderer.padding.left, staffGroup.w, newspace, ret.spacingUnits, ret.minSpace, renderer.padding.left + renderer.padding.right)
'        If (isDebug) Then
'            console.log("setXSpace", it, staffGroup.w, newspace, staffGroup.minspace)
'        End If
        If (newspace = 0) Then break
    Next
    centerWholeRests staffGroup.voices
End Sub

Function calcHorizontalSpacing(isLastLine As Boolean, stretchLast As Boolean, targetWidth, lineWidth, spacing, spacingUnits, minSpace, padding) As Double
    If (isLastLine) Then
        If (stretchLast) Then
            If (lineWidth / targetWidth < 0.66) Then
                calcHorizontalSpacing = 0 '' keep this for backward compatibility. The break isn't quite the same for some reason.
            End If
         Else
            '' "Stretch the last music line of a tune when it lacks less than the float fraction of the page width."
            Dim lack As Double
            Dim stretch As Boolean
            lack = 1 - (lineWidth + padding) / targetWidth
            stretch = lack < stretchLast
            If Not (stretch) Then calcHorizontalSpacing = 0 '' don't stretch last line too much
        End If
    End If
    If (Math.Abs(targetWidth - lineWidth) < 2) Then calcHorizontalSpacing = 0 '' if we are already near the target width, we're done.
    Dim relSpace As Double
    Dim constSpace As Double
    relSpace = spacingUnits * spacing
    constSpace = lineWidth - relSpace
    If (spacingUnits > 0) Then
        spacing = (targetWidth - constSpace) / spacingUnits
        If (spacing * minSpace > 50) Then
            spacing = 50 / minSpace
        End If
        calcHorizontalSpacing = spacing
    End If
    calcHorizontalSpacing = 0
End Function
'/**
' * ���Ÿm��
' * ��ӥ��ŬO�@�دS���p�G�p�G���̦b�@�Ӥp�`����W�s�b�A���򥦭����ө~���C
' * (�p�G���̤��O��W�s�b�A�h�i��O�Τ���~�A���ڭ̱N��m��ⰼ����Ӷ��ؤ��������ߡC�^
' * whole rests are a special case: if they are by themselves in a measure, then they should be centered.
' * (If they are not by themselves, that is probably a user error, but we'll just center it between the two items to either side of it.)
' */
Public Sub centerWholeRests(voices As VoiceABCList)
    '' whole rests are a special case: if they are by themselves in a measure, then they should be centered.
    '' (If they are not by themselves, that is probably a user error, but we'll just center it between the two items to either side of it.)
    Dim i As Integer, j As Integer
    Dim voice As VoiceElement
    Dim abselem As VoiceElement
    Dim befor As VoiceElement
    Dim after As VoiceElement
    For i = 0 To voices.Count - 1
        Set voice = voices(i)
        '' Look through all of the elements except for the first and last. If the whole note appears there then there isn't anything to center it between anyway.
        For j = 1 To voice.children.Count - 1
            Set abselem = voice.children(j)
            If Not (abselem.abcelem.Rest Is Nothing) Then
                If (abselem.abcelem.Rest.typs = "whole" Or abselem.abcelem.Rest.typs = "multimeasure") Then
                Set before = voice.children(j - 1)
                Set after = voice.children(j + 1)
                abselem.center before, after
                End If
            End If
        Next
    Next
End Sub

