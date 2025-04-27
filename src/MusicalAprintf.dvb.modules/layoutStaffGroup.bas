Attribute VB_Name = "layoutStaffGroup"
Option Explicit
'
'import StaffGroupElement from '../creation/elements/staff-group-element'
'import VoiceElement from '../creation/elements/voice-element'
'import Renderer from '../renderer'
'import layoutVoiceElements from './voice-elements'

Function checkLastBarX(voices As VoiceElementList)
''�o�O��̫᪺�p�`�u���
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

Public Function layoutStaffGroup2(spaci As Double, renderer As RendererModule, debug_ As Boolean, StaffGroup As StaffGroupElement, leftEdge As Double) As Dictionary

    Dim currentduration
    Dim durationIndex   As Double
    Const Epsilon As Double = 0.0000001
    Dim spacingunit As Double
    Dim spacingUnits As Double
    Dim minSpace As Double
    Dim x As Double
    Dim i As Integer, j As Integer, k As Integer
    
    
    
    For i = 0 To StaffGroup.voices.Count - 1
        layoutVoiceElements.beginLayout x, StaffGroup.voices(i)
    Next
    
    
    Dim errCount As Long
    minSpace = 1000
    '�o�j��O�]�w X �b�V
    Do While (finished(StaffGroup.voices) = False And errCount < 500)  ' Inner loop.
'       Dim currVoice As VoiceElement
'       Set currVoice = StaffGroup.voices(1)
'       Debug.Print "layoutStaffGroup loop :" & currVoice.i
       errCount = errCount + 1
        
        '' ���n�b���n�����Կ�̤����G�m���Ĥ@�ӫ���ɶ��ŧO
        currentduration = Empty '' candidate smallest duration level
        For i = 0 To StaffGroup.voices.Count - 1
            If currentduration = Empty Then
                currentduration = getDurationIndex(StaffGroup.voices(i))
            ElseIf getDurationIndex(StaffGroup.voices(i)) < currentduration Then
                currentduration = getDurationIndex(StaffGroup.voices(i))
                
            End If
        Next
        
        
              
        
        ''isolate voices at current duration level
        ''�j���ثe����ɶ����Ū��y��
        Dim currentvoices As New iArray ' VoiceElement[] = []
        Dim othervoices As New iArray ' VoiceElement[] = []
        currentvoices.Clear
        othervoices.Clear
        For i = 0 To StaffGroup.voices.Count - 1
            durationIndex = getDurationIndex(StaffGroup.voices(i))
            '' PER: Because of the inexactness of JS floating point math, we just get close.
            '' PER�G�ѩ� JS �B�I�ƾǪ�����T�ʡA�ڭ̥u�O����Ӥw�C
            If (durationIndex - currentduration > Epsilon) Then
                othervoices.Push StaffGroup.voices(i)
             Else
                currentvoices.Push StaffGroup.voices(i)
            End If
'            If i = 0 Then
'                Debug.Print "out tracks :" & i & " " & StaffGroup.voices(i).i & "/" & StaffGroup.voices(i).children.Count & _
'                " allDur: " & StaffGroup.voices(i).durationIndex & "<+ " & StaffGroup.voices(i).children(StaffGroup.voices(i).i).duration;
'            Else
'                Debug.Print "   " & i & " " & StaffGroup.voices(i).i & "/" & StaffGroup.voices(i).children.Count & _
'                " allDur: " & StaffGroup.voices(i).durationIndex & " <+ " & StaffGroup.voices(i).children(StaffGroup.voices(i).i).duration
'            End If
            
        Next
        
         
        

        
        '' among the current duration level find the one which needs starting furthest right
        '' �b�ثe����ɶ��ŧO�����ݭn�q�̥k��}�l������ɶ��ŧO
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
            Dim V As VoiceElement
            Dim topVoice As VoiceElement
            Dim voicechildx As Double
            Dim dx As Double
            Set V = currentvoices(i)
            If (V.voicenumber = 0) Then lastTopVoice = i
            If lastTopVoice <> Empty And currentvoices(lastTopVoice).voicenumber <> V.voicenumber Then
                Set topVoice = currentvoices(lastTopVoice)
            Else
                Set topVoice = Nothing
            End If
            ''line ������ if (~isSameStaff(v, topVoice)) then   Set topVoice = Empty
            voicechildx = layoutVoiceElements.layoutOneItem(x, spaci, V, 0, topVoice)
            dx = voicechildx - x
            ''�o�O�ݬO�_���e�ʭ�
            ''�p�G���A���������ŴN�b�[�e�ʭ����Z��
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
        '' �R����L�y�����w�p�⪺���j��쪺��
        ''�]�Ҧp�A�p�G�@�ӻy���p���ϥ� 5 �Ӷ��j���A�����b������ɶ��ŧO�W�ƦC -
        '' �ڭ̨ϥΤF 2 �Ӷ��j��� - ����ڭ̥����Χ�3�Ӷ��Z���A�ӫD5 �ӡ^
        '' ���իᤣ�ݭn
        For i = 0 To othervoices.Count - 1
            othervoices(i).spacingduration = othervoices(i).spacingduration - spacingduration
            Call layoutVoiceElements.updateNextX(x, spaci, othervoices(i))   '' adjust other voices expectations
        Next
        
                    
              
        '' ��s�ثe�G������������
        For i = 0 To currentvoices.Count - 1
            Dim voice As VoiceElement
            Set voice = currentvoices(i)
            '' ��C�@�� voice.i �[ 1 ���U�@�Ӥl����
            '' �٦��ק� voice.durationindex �[�W�{�b�w�gŪ�� �����Ū���
            '' 4������=0.25 2������=0.5 ��������=1
            '' �C�@�p�`�`��(����)�� 1
            Call layoutVoiceElements.updateIndices(voice)
        Next
    Loop
    i = i + 1



    '' find the greatest remaining x as a base for the width
    '' ��X�̤j���Ѿl x �@���e�ת����
    For i = 0 To i < StaffGroup.voices.Count - 1
        If (layoutVoiceElements.getNextX(StaffGroup.voices(i)) > x) Then
            x = layoutVoiceElements.getNextX(StaffGroup.voices(i))
            spacingunit = layoutVoiceElements.getSpacingUnits(StaffGroup.voices(i))
        End If
    Next

    '' adjust lastBar when needed (multi staves)
    Call checkLastBarX(StaffGroup.voices)
    ''console.log("greatest remaining",spacingunit,x)
    spacingUnits = spacingUnits + spacingunit
    ''��@�ժ� V �e�׳]�w �̼e
    StaffGroup.setWidth (x)
    
    
    StaffGroup.spacingUnits = spacingUnits
    StaffGroup.minSpace = minSpace
    
End Function


Public Function finished(voices As VoiceElementList) As Boolean
    Dim i As Integer
    Dim V As VoiceElement
    For i = 0 To voices.Count - 1
        Set V = voices(i)
        If (layoutVoiceElements.layoutEnded(V) = False) Then
            finished = False
            Exit Function
        End If
    Next
            
    finished = True
End Function

Function getDurationIndex(Element As VoiceElement) As Double
    '' if the ith element doesn't have a duration (is not a note), its duration index is fractionally before.
    '' This enables CLEF KEYSIG TIMESIG PART, etc.to be laid out before we get to the first note of other voices
    '' �p�G�� i �Ӥ����S������ɶ��]���O���š^�A�h�����ɶ����ަb�e���C
    '' �o�ϱo CLEF KEYSIG TIMESIG PART ������b�ڭ̨�F��L�n�����Ĥ@�ӭ��Ť��e�i��G��
    Dim getItemDuration As Double
    
    If Not (Element.children(Element.i) Is Nothing) Then
        If TypeOf Element.children(Element.i) Is MusicItem Then
            If Element.children(Element.i).duration > 0 Then
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
    
    getDurationIndex = Element.durationIndex - getItemDuration
End Function

Public Function isSameStaff(voice1 As VoiceElement, voice2 As VoiceElement) As Boolean
    If (voice1 = Empty Or voice1.Staff = Empty Or voice1.Staff.voices = Empty Or voice1.Staff.voices.Count = 0) Then
        isSameStaff = False
        Exit Function
    End If
    If (voice2 = Empty Or voice2.Staff = Empty Or voice2.Staff.voices = Empty Or voice2.Staff.voices.Count = 0) Then
        isSameStaff = False
        Exit Function
    End If
    isSameStaff = Not (voice1.Staff.voices(0) <> voice2.Staff.voices(0))
End Function



