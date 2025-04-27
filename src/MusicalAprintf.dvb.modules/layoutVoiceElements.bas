Attribute VB_Name = "layoutVoiceElements"
Option Explicit

Sub beginLayout(startX As Double, voice As VoiceElement)

    voice.i = 0
    voice.durationIndex = 0
    ''this.ii=this.children.length
    voice.startX = startX
    voice.minX = startX '' furthest left To where negatively positioned elements are allowed To go
    voice.nextX = startX '' x position where the Next element Of this voice should be placed assuming no other voices And no fixed width constraints
    voice.spacingduration = 0 '' duration left To be laid out In current iteration (omitting additional spacing due To other aspects, such As bars, dots, sharps And flats)
End Sub


Function layoutEnded(voice As VoiceElement) As Boolean
    layoutEnded = (voice.i >= voice.children.Count)
End Function


Function getNextX(voice As VoiceElement) As Double
    getNextX = Math.max(Array(voice.minX, voice.nextX))
        'Return Math.max(voice.minx, voice.nextx)
End Function

'' number of spacing units expected for next positioning
Function getSpacingUnits(voice As VoiceElement)
   getSpacingUnits = Sqrt(voice.spacingduration# / 1000 * 8)
End Function

'**
'* ���էG������ this.i �B������
'* �C�����N����h���I�s�����
'* @param x position to try to layout the element at
'*          ���էG����������m
'* @param spacing base spacing
'*                ��¦���Z
'*'
Function layoutOneItem(x As Double, spacing As Double, voice As VoiceElement, minPadding As Double, firstVoice As VoiceElement)
    Dim child As MusicItem 'voiceItem
    Dim er As Double
    Dim pad  As Double
    Dim firstChild As MusicItem 'voiceItem
    Dim overlaps As Boolean
    Dim j As Integer
    Set child = voice.children(voice.i)
    If (child Is Nothing) Then
        layoutOneItem = 0
        Exit Function
    End If
    '' er : available extrawidth To the left  �����i�Ϊ��B�~�e��
    er = x - voice.minX
    '' pad : only add padding To the items that aren't fixed to the left edge.
    '' pad : �Ȧb���T�w�쥪��t�����طs�W�񺡡C
    If child.duration > 0 Then
        pad = voice.durationIndex / 1000 + minPadding
    Else
        pad = voice.durationIndex / 1000
    End If
        '' See if this item overlaps the item in the first voice. If firstVoice Is undefined then there's nothing to compare.
        '' �d�ݸӶ��جO�_�P�Ĥ@�ӻy���������ح��|�C �p�G firstVoice ���w�q�A�h�S������i������C
    '    If (child.abcelem.el_type === "note" && !child.abcelem.rest && voice.voicenumber! == 0 && firstVoice) Then
    '    firstChild = firstVoice.children(firstVoice.i)
    '    '' It overlaps if the either the child's top or bottom is inside the firstChild's or at least within 1
    '    '' A special case Is if the element Is on the same line then it can share a note head, if the notehead Is the same
    '    '' �p�G�Ĥl�������Ω����b�Ĥ@�ӫĤl�������Φܤ֦b 1 �����A�h�����|
    '    '' �@�دS���p�O�p�G�����b�P�@�樺�򥦥i�H�@�Τ@�ӭ����Y�A�p�G�����Y�ۦP
    '    overlaps = firstChild &&
    '            ((child.abcelem.maxpitch <= firstChild.abcelem.maxpitch + 1 && child.abcelem.maxpitch >= firstChild.abcelem.minpitch - 1) ||
    '                (child.abcelem.minpitch <= firstChild.abcelem.maxpitch + 1 && child.abcelem.minpitch >= firstChild.abcelem.minpitch - 1))
    '        '' See if they can share a note head
    '        '' �ݬݥL�̬O�_�i�H�@�ɤ@�ӭ����Y
    '        If (overlaps && child.abcelem.minpitch === firstChild.abcelem.minpitch && child.abcelem.maxpitch === firstChild.abcelem.maxpitch &&
    '            firstChild.heads && firstChild.heads.length > 0 && child.heads && child.heads.length > 0 &&
    '            firstChild.heads[0].c === child.heads[0].c) Then overlaps = False
    '    '' If this note overlaps the note in the first voice And we haven't moved the note yet (this can be called multiple times)
    '    '' �p�G�o�ӭ��ŻP�Ĥ@���n���������ŭ��|�A�åB�ڭ��٨S�����ʸӭ��š]�i�H�h���I�s�^
    '    If (overlaps) Then
    '        '' I think that firstChild should always have at least one note head, but defensively make sure.
    '        '' There was a problem with this being called more than once so if a value Is adjusted then it Is saved so it Is only adjusted once.
    '        var firstChildNoteWidth = firstChild.heads && firstChild.heads.length > 0 ? firstChild.heads[0].realWidth : firstChild.fixed.w
    '        If (!child.adjustedWidth) Then
    '            child.adjustedWidth = firstChildNoteWidth + child.w
    '            child.w = child.adjustedWidth
    '            For (j = 0 j < child.children.length j++) then
    '                var relativeChild = child.children[j]
    '                If (relativeChild.name.indexOf("accidental") < 0) Then {
    '                    If (!relativeChild.adjustedWidth) Then
    '                    relativeChild.adjustedWidth = relativeChild.dx + firstChildNoteWidth
    '                End If
    '                relativeChild.dx = relativeChild.adjustedWidth
    '            End If

    '        Next
    '    End If
    'End If

    Dim extraWidth As Double
    extraWidth = getExtraWidth(child, pad)
    If (er < extraWidth) Then  '' shift right by needed amount �k���һݼƶq
        '' There's an exception if a bar element is after a Part element, there is no shift.
        '' �p�G bar ������� Part ��������A�h���ҥ~���p�A�h���|�o�Ͳ���C
        If (voice.i = 0 Or child.typs <> Cg.bar) Then
            x = x + extraWidth - er
        ElseIf voice.children(voice.i - 1).typs <> Cg.part And voice.children(voice.i - 1).typs <> Cg.tempo Then
            x = x + extraWidth - er
        End If
    End If
    child.setX x

    voice.spacingduration = child.duration
    ''update minx
    voice.minX = x + getMinWidth(child) '' add necessary layout space �s�W���n���G���Ŷ�
    If (voice.i <> voice.children.Count - 1) Then
        voice.minX = voice.minX + child.minspacing '' add minimumspacing except On last elem �s�W���̫�@�Ӥ������~���̤p���Z
    End If

    '' �p��U�@�Ӥ��󪺦�m
    Call updateNextX(x, spacing, voice)

    '' contribute to staff y position
    ''this.staff.top = Math.max(child.top,this.staff.top)
    ''this.staff.bottom = Math.min(child.bottom,this.staff.bottom)

    layoutOneItem = x '' where we End up having placed the child
End Function

Sub shiftRight(dx As Double, voice As VoiceElement)
    Dim child As voiceItem
    Set child = voice.children(voice.i)
    If (child = Empty) Then Exit Sub
    child.setX (child.x + dx)
    voice.minX = voice.minX + dx
    voice.nextX = voice.nextX + dx
End Sub

'' call when spacingduration has been updated
Sub updateNextX(x As Double, spacing As Double, voice As VoiceElement)
    voice.nextX = x + (spacing * Math.Sqrt(voice.spacingduration / 1000 * 8))
End Sub

Sub updateIndices(voice As VoiceElement)
    If (layoutEnded(voice) = False) Then
        voice.durationIndex = voice.durationIndex + voice.children(voice.i).duration
        If (voice.children(voice.i).typs = Cg.bar) Then
            '' everytime we meet a barline, do rounding to nearest 64th
            '' �C���J��p�`�u�ɡA���|�|�ˤ��J��̱��񪺲� 64 ��
            'voice.durationIndex = Round(voice.durationIndex / 6400) * 6400 '' 64) / 64 '' everytime we meet a barline, Do rounding To nearest 64th
        End If
        voice.i = voice.i + 1
    End If

End Sub
Function getExtraWidth(child As MusicItem, minPadding As Double) As Double   '' space needed To the left Of the note
    Dim padding As Double
    If (child.typs = Cg.note Or child.typs = Cg.bar) Then
        padding = minPadding
    End If
    getExtraWidth = -child.extraw + padding
End Function

Function getMinWidth(child As MusicItem) As Double  '' absolute space taken To the right Of the note
    getMinWidth = child.w
End Function



