VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StaffGroupElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public line As Integer
Public startX As Double
Public w As Double
Public height As Double
Public gTextSize As getTextSize
Public voices As VoiceElementList   ' VoiceElement
Public staffInfos As vStaffInfoList
Public brace As BraceElemList
Public bracket As BraceElemList
Private sa_(100) As Size
Private saIs As New Dictionary

Public Property Get sa(Index As Integer) As Size
    If saIs(Index) = True Then
        sa = sa_(Index)
        Exit Property
    ElseIf UBound(sa_) >= Index Then
        sa = Empty
    End If
End Property

Public Property Let sa(Index As Integer, si As Size)
    If UBound(sa_) >= Index Then
        sa_(Index) = si
        saIs(Index) = True
    Else
        'ReDim Me.sa_(index * 2)
        sa_(Index) = si
        saIs(Index) = True
    End If
    
End Property

Public Sub init(gTextSize As getTextSize)

        Set Me.gTextSize = gTextSize
        Set Me.voices = New VoiceElementList
        Set Me.staffInfos = New vStaffInfoList
        Set Me.brace = New BraceElemList  ''tony
        Set Me.bracket = New BraceElemList

End Sub

Public Sub setLimit(member As String, voice As VoiceElement)
        If (voice.specialY(member) Is Nothing) Then Exit Sub
        If (voice.Staff.specialY(member) Is Nothing) Then
            voice.Staff.specialY(member) = voice.specialY(member)
        Else
            voice.Staff.specialY(member) = Math.max(voice.Staff.specialY(member), voice.specialY(member))
        End If
End Sub

Public Sub addVoice(voice As VoiceElement, staffnumber As Integer, stafflines As Integer)
        Dim voiceNum As Integer
        Dim staffEle As vStaffInfo
        Dim zeor0
        voiceNum = Me.voices.Count
        Me.voices.Push voice
        If Not (Me.staffInfos(staffnumber) Is Nothing) Then
            Me.staffInfos(staffnumber).voices.Push (voiceNum)
        Else
            '' TODO-PER: how does the min/max change when stafflines is not 5?
            Set staffEle = New vStaffInfo
            staffEle.top = 10
            staffEle.bottom = 2
            staffEle.lines = stafflines
            staffEle.voices.Push voiceNum
            zeor0 = 0
            Dim spY As New Dictionary
            spY("tempoHeightAbove") = 0
            spY("partHeightAbove") = 0
            spY("volumeHeightAbove") = 0
            spY("dynamicHeightAbove") = 0
            spY("endingHeightAbove") = 0
            spY("chordHeightAbove") = 0
            spY("lyricHeightAbove") = 0
            spY("lyricHeightBelow") = 0
            spY("chordHeightBelow") = 0
            spY("volumeHeightBelow") = 0
            spY("dynamicHeightBelow") = 0
            Set staffEle.specialY = spY
            
            Me.staffInfos.Push staffEle
        End If
        Set voice.staffInfo = Me.staffInfos(staffnumber)
End Sub

Public Sub setHeight()
        Me.height = calcHeight(Me)
End Sub

Public Sub setWidth(width As Double)
        Dim i As Integer
        Me.w = width
        For i = 0 To Me.voices.Count - 1
            Call Me.voices(i).setWidth(width)
        Next
End Sub

Public Sub setStaffLimits(voice As VoiceElement)
        voice.Staff.top = Math.max(voice.Staff.top, voice.top)
        voice.Staff.bottom = Math.min(voice.Staff.bottom, voice.bottom)
        Me.setLimit "tempoHeightAbove", voice
        Me.setLimit "partHeightAbove", voice
        Me.setLimit "volumeHeightAbove", voice
        Me.setLimit "dynamicHeightAbove", voice
        Me.setLimit "endingHeightAbove", voice
        Me.setLimit "chordHeightAbove", voice
        Me.setLimit "lyricHeightAbove", voice
        Me.setLimit "lyricHeightBelow", voice
        Me.setLimit "chordHeightBelow", voice
        Me.setLimit "volumeHeightBelow", voice
        Me.setLimit "dynamicHeightBelow", voice
    
End Sub
