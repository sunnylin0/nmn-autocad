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
'* 嘗試佈局索引 this.i 處的元素
'* 每次迭代不能多次呼叫此函數
'* @param x position to try to layout the element at
'*          嘗試佈局元素的位置
'* @param spacing base spacing
'*                基礎間距
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
    '' er : available extrawidth To the left  左側可用的額外寬度
    er = x - voice.minX
    '' pad : only add padding To the items that aren't fixed to the left edge.
    '' pad : 僅在未固定到左邊緣的項目新增填滿。
    If child.duration > 0 Then
        pad = voice.durationIndex / 1000 + minPadding
    Else
        pad = voice.durationIndex / 1000
    End If
        '' See if this item overlaps the item in the first voice. If firstVoice Is undefined then there's nothing to compare.
        '' 查看該項目是否與第一個語音中的項目重疊。 如果 firstVoice 未定義，則沒有什麼可比較的。
    '    If (child.abcelem.el_type === "note" && !child.abcelem.rest && voice.voicenumber! == 0 && firstVoice) Then
    '    firstChild = firstVoice.children(firstVoice.i)
    '    '' It overlaps if the either the child's top or bottom is inside the firstChild's or at least within 1
    '    '' A special case Is if the element Is on the same line then it can share a note head, if the notehead Is the same
    '    '' 如果孩子的頂部或底部在第一個孩子的內部或至少在 1 之內，則它重疊
    '    '' 一種特殊情況是如果元素在同一行那麼它可以共用一個音符頭，如果音符頭相同
    '    overlaps = firstChild &&
    '            ((child.abcelem.maxpitch <= firstChild.abcelem.maxpitch + 1 && child.abcelem.maxpitch >= firstChild.abcelem.minpitch - 1) ||
    '                (child.abcelem.minpitch <= firstChild.abcelem.maxpitch + 1 && child.abcelem.minpitch >= firstChild.abcelem.minpitch - 1))
    '        '' See if they can share a note head
    '        '' 看看他們是否可以共享一個音符頭
    '        If (overlaps && child.abcelem.minpitch === firstChild.abcelem.minpitch && child.abcelem.maxpitch === firstChild.abcelem.maxpitch &&
    '            firstChild.heads && firstChild.heads.length > 0 && child.heads && child.heads.length > 0 &&
    '            firstChild.heads[0].c === child.heads[0].c) Then overlaps = False
    '    '' If this note overlaps the note in the first voice And we haven't moved the note yet (this can be called multiple times)
    '    '' 如果這個音符與第一個聲音中的音符重疊，並且我們還沒有移動該音符（可以多次呼叫）
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
    If (er < extraWidth) Then  '' shift right by needed amount 右移所需數量
        '' There's an exception if a bar element is after a Part element, there is no shift.
        '' 如果 bar 元素位於 Part 元素之後，則有例外情況，則不會發生移位。
        If (voice.i = 0 Or child.typs <> Cg.bar) Then
            x = x + extraWidth - er
        ElseIf voice.children(voice.i - 1).typs <> Cg.part And voice.children(voice.i - 1).typs <> Cg.tempo Then
            x = x + extraWidth - er
        End If
    End If
    child.setX x

    voice.spacingduration = child.duration
    ''update minx
    voice.minX = x + getMinWidth(child) '' add necessary layout space 新增必要的佈局空間
    If (voice.i <> voice.children.Count - 1) Then
        voice.minX = voice.minX + child.minspacing '' add minimumspacing except On last elem 新增除最後一個元素之外的最小間距
    End If

    '' 計算下一個元件的位置
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
            '' 每次遇到小節線時，都會四捨五入到最接近的第 64 位
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



