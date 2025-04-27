Attribute VB_Name = "layoutStaffGroup"
Option Explicit
'
'import StaffGroupElement from '../creation/elements/staff-group-element'
'import VoiceElement from '../creation/elements/voice-element'
'import Renderer from '../renderer'
'import layoutVoiceElements from './voice-elements'

Function checkLastBarX(voices)
    Dim maxX As Double
    Dim i As Integer
    For i = 0 To voices.Count - 1
        Dim curVoice
        Dim LastChild As Integer
        Dim maxChild
        Dim barX As Double
        Set curVoice = voices(i)
        If (curVoice.children.Count > 0) Then
            Set LastChild = curVoice.children.Count - 1
            Set maxChild = curVoice.children(LastChild)
            If (maxChild.abcelem.el_type = "bar") Then
                barX = maxChild.children(0).x
                If (barX > maxX) Then
                    maxX = barX
                 Else
                    maxChild.children [0].x = maxX
                End If
            End If
        End If
    Next
End Function

Function layoutStaffGroup(spaci As Double, renderer As RendererModule, debug_ As Boolean, staffGroup As StaffGroupElement, leftEdge As Double) As Double

    Dim currentduration
    Dim durationIndex   As Double
    Const Epsilon As Double = 0.0000001
    Dim spacingunit As Double
    Dim spacingUnits As Double
    Dim minSpace As Double
    Dim x As Double
    Dim i As Integer, j As Integer, k As Integer
    
    
    minSpace = 1000
    '�o�j��O�]�w X �b�V
    Do While (finished(staffGroup.voices) = False)   ' Inner loop.
       Dim currVoice As VoiceElement
       Set currVoice = staffGroup.voices(1)
       Debug.Print currVoice.i
       
        
        '' ���n�b���n�����Կ�̤����G�m���Ĥ@�ӫ���ɶ��ŧO
        currentduration = Empty '' candidate smallest duration level
        For i = 0 To staffGroup.voices.Count - 1
            If currentduration = Empty Then
                currentduration = getDurationIndex(staffGroup.voices(i))
            Else
                If getDurationIndex(staffGroup.voices(i)) < currentduration Then
                    currentduration = getDurationIndex(staffGroup.voices(i))
                End If
            End If
        Next
        
        
              
        
        ''isolate voices at current duration level
        ''�j���ثe����ɶ����Ū��y��
        Dim currentvoices As New iArray ' VoiceElement[] = []
        Dim othervoices As New iArray ' VoiceElement[] = []
        currentvoices.Clear
        othervoices.Clear
        For i = 0 To staffGroup.voices.Count - 1
            durationIndex = getDurationIndex(staffGroup.voices(i))
            '' PER: Because of the inexactness of JS floating point math, we just get close.
            '' PER�G�ѩ� JS �B�I�ƾǪ�����T�ʡA�ڭ̥u�O����Ӥw�C
            If (durationIndex - currentduration > Epsilon) Then
                othervoices.Push staffGroup.voices(i)
                ''console.log("out: voice ",i)
             Else
                currentvoices.Push staffGroup.voices(i)
                ''if (debug) console.log("in: voice ",i)
            
            End If
        Next
        
         
        

        
        '' among the current duration level find the one which needs starting furthest right
        '' �b�ثe����ɶ��ŧO�����ݭn�q�̥k��}�l������ɶ��ŧO
        spacingunit = 0 '' number of spacingunits coming from the previously laid out element to this one
        Dim spacingduration As Double
        For i = 0 To currentvoices.Count - 1
            
            If (layoutVoiceElement.getNextX(currentvoices(i)) > x) Then
                x = layoutVoiceElement.getNextX(currentvoices(i))
                spacingunit = layoutVoiceElement.getSpacingUnits(currentvoices(i))
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
            ''line ������ if (~isSameStaff(v, topVoice)) then   Set topVoice = Empty
            voicechildx = layoutVoiceElement.layoutOneItem(x, spaci, v, 0, topVoice)
            dx = voicechildx - x
            ''�o�O�ݬO�_���e�ʭ�
            ''�p�G���A���������ŴN�b�[�e�ʭ����Z��
            If (dx > 0) Then
                x = voicechildx ''update x
                For j = 0 To i '' shift over all previously laid out elements
                    Call layoutVoiceElement.shiftRight(dx, currentvoices(j))
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
            Call layoutVoiceElement.updateNextX(x, spaci, othervoices(i))   '' adjust other voices expectations
        Next
        
                    
              
        '' ��s�ثe�G������������
        For i = 0 To currentvoices.Count - 1
            Dim voice As VoiceElement
            Set voice = currentvoices(i)
            '' ��C�@�� voice.i �[ 1 ���U�@�Ӥl����
            '' �٦��ק� voice.durationindex �[�W�{�b�w�gŪ�� �����Ū���
            '' 4������=0.25 2������=0.5 ��������=1
            '' �C�@�p�`�`��(����)�� 1
            Call layoutVoiceElement.updateIndices(voice)
        Next
    Loop
    i = i + 1



    '' find the greatest remaining x as a base for the width
    '' ��X�̤j���Ѿl x �@���e�ת����
    For i = 0 To i < staffGroup.voices.Count - 1
        If (layoutVoiceElements.getNextX(staffGroup.voices(i)) > x) Then
            x = layoutVoiceElements.getNextX(staffGroup.voices(i))
            spacingunit = layoutVoiceElements.getSpacingUnits(staffGroup.voices(i))
        End If
    Next

    '' adjust lastBar when needed (multi staves)
    checkLastBarX (staffGroup.voices)
    ''console.log("greatest remaining",spacingunit,x)
    spacingUnits = spacingUnits + spacingunit
    ''��@�ժ� V �e�׳]�w �̼e
    staffGroup.setWidth (x)
    Dim returnDict As New Dictionary
    returnDict("spacingUnits") = spacingUnits
    returnDict("minSpace") = minSpace
    Set layoutStaffGroup = returnDict
End Function


Public Function finished(voices As iArray) As Boolean
    Dim i As Integer
    Dim v As VoiceElement
    For i = 0 To i < voices.Count
        Set v = voices(i)
        If (layoutVoiceElements.layoutEnded(v) = False) Then
            finished = False
            Exit Function
        End If
    End If
            
    finished = True
End Function

Function getDurationIndex(Element As VoiceElement) As Double
    '' if the ith element doesn't have a duration (is not a note), its duration index is fractionally before.
    '' This enables CLEF KEYSIG TIMESIG PART, etc.to be laid out before we get to the first note of other voices
    '' �p�G�� i �Ӥ����S������ɶ��]���O���š^�A�h�����ɶ����ަb�e���C
    '' �o�ϱo CLEF KEYSIG TIMESIG PART ������b�ڭ̨�F��L�n�����Ĥ@�ӭ��Ť��e�i��G��
    Dim getItemDuration As Double
    If TypeOf Element.children(Element.i) Is voiceItem Then
        If Element.children(Element.i).dur > 0 Then
            getItemDuration = Element.durationIndex - 0
        Else
            getItemDuration = Element.durationIndex - 0.0000005
        End If
    End If
    
    getDurationIndex = getItemDuration
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



