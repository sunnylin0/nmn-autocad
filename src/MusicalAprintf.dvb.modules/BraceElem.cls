VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BraceElem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public typs As String
Public header As String
Public startVoice As VoiceElement
Public endVoice As VoiceElement
Public lastContinuedVoice As VoiceElement
Public Sub init(voice As VoiceElement, typs As String)
        Me.startVoice = voice
        Me.typs = typs
End Sub

Public Sub setBottomStaff(voice As VoiceElement)
        Me.endVoice = voice
        '' If only the start brace has a name then the name belongs to the brace instead of the staff.
        If (Me.startVoice.header And Me.endVoice.header) Then
            Me.header = Me.startVoice.header
            delete Me.startVoice.header
        End If
End Sub

Public Sub continuing(voice As VoiceElement)
        '' If the final staff isn"t present, then use the last one we saw.
        Me.lastContinuedVoice = voice
End Sub

Public Function getWidth() As Double

        getWidth = 10 '' TODO-PER as  right now the drawing function doesn"t vary the width at all. If it does in the future then this will change.
End Function

Public Function isStartVoice(voice As VoiceElement) As Boolean
        If (Not Me.startVoice Is Nothing And Not Me.startVoice.Staff Is Nothing) Then
            If (Me.startVoice.Staff.voices.Length > 0 And Me.startVoice.Staff.voices(0) Is voice) Then
                isStartVoice = True
            End If
        End If
        isStartVoice = False
    
End Function
