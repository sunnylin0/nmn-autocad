Attribute VB_Name = "layoutStaffGroup"
Option Explicit
'
'import StaffGroupElement from '../creation/elements/staff-group-element'
'import VoiceElement from '../creation/elements/voice-element'
'import Renderer from '../renderer'
'import layoutVoiceElements from './voice-elements'

Function checkLastBarX(voices As VoiceElementList)
''這是把最後的小節線對齊
    Dim maxX As Double
    Dim i As Integer
    For i = 0 To voices.Count - 1
        Dim curVoice
        Dim LastChildNum As Integer
        Dim maxChild As MusicItem 'VoiceElement
        Dim barX As Double
        Set curVoice = voices(i)
        If (curVoice.children.Count > 0) Then
            LastChildNum = curVoice.children.Count - 1
            Set maxChild = curVoice.children(LastChildNum)
'            If (maxChild.abcelem.el_type = "bar") Then
'                barX = maxChild.children(0).x
'                If (barX > maxX) Then
'                    maxX = barX
'                 Else
'                    maxChild.children [0].x = maxX
'                End If
'            End If
            If (maxChild.typs = Cg.bar) Then
                barX = maxChild.x
                If (barX > maxX) Then
                    maxX = barX
                 Else
                    maxChild.x = maxX
                End If
            End If
        End If
    Next
End Function

Public Function layoutStaffGroup2(spaci As Double, renderer As RendererModule, debug_ As Boolean, staffGroup As StaffGroupElement, leftEdge As Double) As Dictionary

    Dim currentduration
    Dim durationIndex   As Double
    Const Epsilon As Double = 0.0000001
    Dim spacingunit As Double
    Dim spacingUnits As Double
    Dim minSpace As Double
    Dim x As Double
    Dim i As Integer, j As Integer, k As Integer
    
    
    
    For i = 0 To staffGroup.voices.Count - 1
        layoutVoiceElements.beginLayout x, staffGroup.voices(i)
    Next
    
    
    Dim errCount As Long
    minSpace = 1000
    '這迴圈是設定 X 軸向
    Do While (finished(staffGroup.voices) = False And errCount < 500)  ' Inner loop.
'       Dim currVoice As VoiceElement
'       Set currVoice = StaffGroup.voices(1)
'       Debug.Print "layoutStaffGroup loop :" & currVoice.i
       errCount = errCount + 1
        
        '' 找到要在跨聲音的候選者之間佈置的第一個持續時間級別
        currentduration = Empty '' candidate smallest duration level
        For i = 0 To staffGroup.voices.Count - 1
            If Not layoutVoiceElements.layoutEnded(staffGroup.voices(i)) Then
                If IsEmpty(currentduration) Then
                    currentduration = getDurationIndex(staffGroup.voices(i))
                ElseIf getDurationIndex(staffGroup.voices(i)) < currentduration Then
                    currentduration = getDurationIndex(staffGroup.voices(i))
                    
                End If
            End If
        Next
        
        
              
        
        ''isolate voices at current duration level
        ''隔離目前持續時間等級的語音
        Dim currentvoices As New iArray ' VoiceElement[] = []
        Dim othervoices As New iArray ' VoiceElement[] = []
        currentvoices.Clear
        othervoices.Clear
        For i = 0 To staffGroup.voices.Count - 1
            durationIndex = getDurationIndex(staffGroup.voices(i))
            '' PER: Because of the inexactness of JS floating point math, we just get close.
            '' PER：由於 JS 浮點數學的不精確性，我們只是接近而已。
            If (durationIndex - currentduration > Epsilon) Then
                othervoices.Push staffGroup.voices(i)
            Else
                currentvoices.Push staffGroup.voices(i)
            End If
            
        Next
        
        '' among the current duration level find the one which needs starting furthest right
        '' 在目前持續時間級別中找到需要從最右邊開始的持續時間級別
        spacingunit = 0 '' number of spacingunits coming from the previously laid out element to this one
        Dim spacingduration As Double
        For i = 0 To currentvoices.Count - 1
            
            If (layoutVoiceElements.getNextX(currentvoices(i)) > x) Then
                x = layoutVoiceElements.getNextX(currentvoices(i))
                spacingunit = layoutVoiceElements.getSpacingUnits(currentvoices(i))
                spacingduration = currentvoices(i).spacingduration
            End If
        Next
        spacingUnits = spacingUnits + spacingunit
        minSpace = Math.min(Array(minSpace, spacingunit))
        

        Dim lastTopVoice
        For i = 0 To currentvoices.Count - 1
            Dim v As VoiceElement
            Dim topVoice As VoiceElement
            Dim voicechildx As Double
            Dim dx As Double
            Set v = currentvoices(i)
            If (v.voicenumber = 0) Then lastTopVoice = i
            If lastTopVoice <> Empty And currentvoices(lastTopVoice).voicenumber <> v.voicenumber Then
                Set topVoice = currentvoices(lastTopVoice)
            Else
                Set topVoice = Nothing
            End If
            ''line 不知到 if (~isSameStaff(v, topVoice)) then   Set topVoice = Empty
            'Debug.Print v.children(v.I).notes(0).mnote
            voicechildx = layoutVoiceElements.layoutOneItem(x, spaci, v, 0, topVoice)
            dx = voicechildx - x
            ''這是看是否有前倚音
            ''如果有，全部的音符就在加前倚音的距離
            If (dx > 0) Then
                x = voicechildx ''update x
                For j = 0 To i '' shift over all previously laid out elements
                    Call layoutVoiceElements.shiftRight(dx, currentvoices(j))
                Next
            End If
        Next

        '' remove the value of already counted spacing units in other voices
        '' (e.g.if a voice had planned to use up 5 spacing units but is not in line to be laid out at this duration level -
        '' where we've used 2 spacing units - then we must use up 3 spacing units, not 5)
        '' 刪除其他語音中已計算的間隔單位的值
        ''（例如，如果一個語音計劃使用 5 個間隔單位，但未在此持續時間級別上排列 -
        '' 我們使用了 2 個間隔單位 - 那麼我們必須用完3個間距單位，而非5 個）
        '' 測試後不需要
        For i = 0 To othervoices.Count - 1
            othervoices(i).spacingduration = othervoices(i).spacingduration - spacingduration
            Call layoutVoiceElements.updateOtherNextX(x, spaci, othervoices(i))   '' adjust other voices expectations
        Next
        
                    
              
        '' 更新目前佈局元素的索引
        For i = 0 To currentvoices.Count - 1
            Dim voice As VoiceElement
            Set voice = currentvoices(i)
            '' 把每一個 voice.i 加 1 為下一個子元素
            '' 還有修改 voice.durationindex 加上現在已經讀取 的音符長度
            '' 4分音附=0.25 2分音附=0.5 全分音附=1
            '' 每一小節總長(分母)為 1
            Call layoutVoiceElements.updateIndices(voice)
        Next
    Loop
    i = i + 1



    '' find the greatest remaining x as a base for the width
    '' 找出最大的剩餘 x 作為寬度的基數
    For i = 0 To i < staffGroup.voices.Count - 1
        If (layoutVoiceElements.getNextX(staffGroup.voices(i)) > x) Then
            x = layoutVoiceElements.getNextX(staffGroup.voices(i))
            spacingunit = layoutVoiceElements.getSpacingUnits(staffGroup.voices(i))
        End If
    Next

    '' adjust lastBar when needed (multi staves)
    Call checkLastBarX(staffGroup.voices)
    ''console.log("greatest remaining",spacingunit,x)
    spacingUnits = spacingUnits + spacingunit
    ''把一組的 V 寬度設定 最寬
    staffGroup.setWidth (x)
    
    
    staffGroup.spacingUnits = spacingUnits
    staffGroup.minSpace = minSpace
    
End Function


Public Function finished(voices As VoiceElementList) As Boolean
    Dim i As Integer
    Dim v As VoiceElement
    For i = 0 To voices.Count - 1
        Set v = voices(i)
        If (layoutVoiceElements.layoutEnded(v) = False) Then
            finished = False
            Exit Function
        End If
    Next
            
    finished = True
End Function

Function getDurationIndex(element As VoiceElement) As Double
    '' if the ith element doesn't have a duration (is not a note), its duration index is fractionally before.
    '' This enables CLEF KEYSIG TIMESIG PART, etc.to be laid out before we get to the first note of other voices
    '' 如果第 i 個元素沒有持續時間（不是音符），則其持續時間索引在前面。
    '' 這使得 CLEF KEYSIG TIMESIG PART 等能夠在我們到達其他聲音的第一個音符之前進行佈局
    Dim getItemDuration As Double
    
    If Not (element.children(element.i) Is Nothing) Then
        If TypeOf element.children(element.i) Is MusicItem Then
            If element.children(element.i).duration > 0 Then
                getItemDuration = 0
            Else
                getItemDuration = 0.0000005
            End If
        Else
            getItemDuration = 0.0000005
        End If
        
    Else
        getItemDuration = 0.0000005
    End If
    
    getDurationIndex = element.durationIndex - getItemDuration
End Function

Public Function isSameStaff(voice1 As VoiceElement, voice2 As VoiceElement) As Boolean
    If (IsEmpty(voice1) Or IsEmpty(voice1.Staff) Or IsEmpty(voice1.Staff.voices) Or voice1.Staff.voices.Count = 0) Then
        isSameStaff = False
        Exit Function
    End If
    If (IsEmpty(voice2) Or IsEmpty(voice2.Staff) Or IsEmpty(voice2.Staff.voices) Or voice2.Staff.voices.Count = 0) Then
        isSameStaff = False
        Exit Function
    End If
    isSameStaff = Not (voice1.Staff.voices(0) <> voice2.Staff.voices(0))
End Function



