Attribute VB_Name = "LayoutModule"
Option Explicit


Function layout(abcTuneLines As TuneLineList, width As Double, space As Double) As Double
    Dim i As Integer
    Dim j As Integer
    
    Dim abcLine As TuneLine
    '' Adjust the x-coordinates to their absolute positions
    '' �N x ���нվ��䵴���m
    Dim maxWidth  As Double
    maxWidth = width
    For i = 0 To abcTuneLines.Count - 1
        Set abcLine = abcTuneLines(i)
        If Not (abcLine.Staffs Is Nothing) Then
        
            Call setXSpacing(New RendererModule, width, space, abcLine.StaffGroup, New vFormatting, i = abcTuneLines.Count - 1, False)
            
            If (abcLine.StaffGroup.w > maxWidth) Then
                maxWidth = abcLine.StaffGroup.w
            End If
        End If
    Next
'
'    '' Layout the beams and add the stems to the beamed notes.
'    '' �G�����ñN�D�F�K�[���筵�Ť��C
'    For i = 0 To abctune.lines.clount - 1
'        Set abcLine = abctune.lines(i)
'        If Not (abcLine.StaffGroup Is Nothing) Then
'            If Not (abcLine.StaffGroup.voices Is Nothing) Then
'                For j = 0 To abcLine.StaffGroup.voices.Count
'                    layoutVoice (abcLine.StaffGroup.voices(j))
'                Next
'                setUpperAndLowerElements renderer, abcLine.StaffGroup
'            End If
'        End If
'    Next
'
'    '' Set the staff spacing
'    '' TODO-PER: we should have been able to do this by the time we called setUpperAndLowerElements, but for some reason the "bottom" element seems to be set as a side effect of setting the X spacing.
'    '' �]�m���u�ж��Z
'    '' TODO-PER�G��ڭ̽ե� setUpperAndLowerElements �ɡA�ڭ����ӯ������o�@�I�A
'    ''���ѩ�Y�ح�]�A���������������G�Q�]�m���]�m X ���Z���Ƨ@�ΡC
'    For i = 0 To abctune.lines.Count - 1
'        Set abcLine = abctune.lines(i)
'        If Not (abcLine.StaffGroup Is Nothing) Then
'            abcLine.StaffGroup.setHeight
'        End If
'    Next
    layout = maxWidth
End Function
'/**
' * Do the x-axis positioning for a single line (a group of related staffs)
' * ���u�]�@�լ����H���^�i��x�b�w��
' */
Public Sub setXSpacing(renderer As RendererModule, width As Double, space As Double, StaffGroup As StaffGroupElement, formatting As vFormatting, isLastLine As Boolean, isDebug As Boolean)
    Dim leftEdge As Double
    Dim newspace As Double
    Dim lastSpace As Double
    'leftEdge = getLeftEdgeOfStaff(renderer, StaffGroup.getTextSize, StaffGroup.voices, StaffGroup.brace, StaffGroup.bracket)
    leftEdge = 0
    newspace = space
    '' TODO-PER: shouldn't need multiple passes, but each pass gets it closer to the right spacing. (Only affects long lines: normal lines break out of this loop quickly.)
    '' TODO-PER�G���ݭn�h���q�L�A���C���q�L���|�Ϩ�󱵪񥿽T�����Z�C �]�ȼv�T����G���q��ܧִN�|�������`���C�^
    Dim it As Integer
    
    For it = 0 To 7
        Call layoutStaffGroup2(newspace, New RendererModule, isDebug, StaffGroup, leftEdge)
        newspace = calcHorizontalSpacing(isLastLine, formatting.stretchLast, width - G.LeftSpace - G.RightSpace, StaffGroup.w, newspace, StaffGroup.spacingUnits, StaffGroup.minSpace, G.LeftSpace + G.RightSpace) ', renderer.padding.left + renderer.padding.right)
        If newspace <> lastSpace Then
            lastSpace = newspace
        Else
            Exit For
        End If
'        If (isDebug) Then
'            console.log("setXSpace", it, staffGroup.w, newspace, staffGroup.minspace)
'        End If
        If (newspace = 0) Then
            Exit For
        End If
    Next
    '����Ť�m��
    ''centerWholeRests StaffGroup.voices
End Sub

Function calcHorizontalSpacing(isLastLine As Boolean, stretchLast As Double, targetWidth As Double, lineWidth As Double, spacing As Double, spacingUnits As Double, minSpace As Double, padding As Double) As Double
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
    If (Abs(targetWidth - lineWidth) < 2) Then calcHorizontalSpacing = 0 '' if we are already near the target width, we're done.
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
        Exit Function
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
Public Sub centerWholeRests(voices As VoiceElementList)
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


'/** ��X�C�q�־��W�̤j���e��  */
Function getLeftEdgeOfStaff(renderer As RendererModule, getTextSize As Double, voices As VoiceElementList, brace, bracket) As Double

    Dim x   As Double
    Dim voiceheaderw   As Double
    Dim size   As Double
    Dim sizeW As Double
    Dim ofs As Double
    Dim i   As Integer
    
    Dim x   As Double
    x = renderer.padding.left

    '' find out how much space will be taken up by voice headers
     voiceheaderw = 0
    Dim gTextSize As New getTextSize
    For i = 0 To voices.Length - 1
        If (voices(i).header Is Nothing) Then
            size = gTextSize.calc(voices(i).header, "voicefont", "")
            voiceheaderw = Math.max(voiceheaderw, size.width)
        End If
    Next
    voiceheaderw = addBraceSize(voiceheaderw, brace, getTextSize)
    voiceheaderw = addBraceSize(voiceheaderw, bracket, getTextSize)
    
    
    If (voiceheaderw > 0) Then
        '' Give enough spacing to the right - we use the width of an A for the amount of spacing.
        '' ���k�����������Z - �ڭ̨ϥ� A ���e�ק@�����Z�q�C
        '' �b�[�W 1 �Ӧr���e��
        sizeW = gTextSize.calc("A", "voicefont", "")
        voiceheaderw = voiceheaderw + sizeW.width
    End If
    x = x + voiceheaderw

    ofs = 0
    ofs = setBraceLocation(brace, x, ofs)
    ofs = setBraceLocation(bracket, x, ofs)
    getLeftEdgeOfStaff = x + ofs

End Function

Function addBraceSize(voiceheaderw, brace, getTextSize) As Double
    Dim i As Integer
    Dim size As Double
    If Not (brace Is Nothing) Then
        For i = 0 To i < brace.Length - 1
            If Not (brace(i).header Is Nothing) Then
                 size = gTextSize.calc(brace(i).header, "voicefont", "")
                voiceheaderw = Math.max(voiceheaderw, size.width)
            End If
        Next
    End If
    addBraceSize = voiceheaderw

End Function

Function setBraceLocation(brace, x, ofs) As Double
    Dim i As Integer
    If Not (brace Is Nothing) Then
        For i = 0 To brace.Length - 1
            Call setLocation(x, brace(i))
            ofs = Math.max(ofs, brace(i).getWidth())
        Next
    End If
    setBraceLocation = ofs


End Function

Function setLocation(x, Element)
    Element.x = x
End Function
