VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "newABCModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function VoiceElem(voicenumber As Integer, voicetotal As Integer) As VoiceElement
    Set VoiceElem = New VoiceElement
    VoiceElem.init voicenumber, voicetotal
End Function

Public Function AbstractEng(gTextSize As getTextSize, tuneNumber As Integer, options As Dictionary) As AbstractEngraver
    Set AbstractEng = New AbstractEngraver
    AbstractEng.init gTextSize, tuneNumber, options

End Function

Function RelativeElem(c As String, dx As Double, w As Double, pitch As Double, Optional opt As oRelativeOptions = Nothing) As RelativeElement
    Dim rElem As New RelativeElement
    rElem.init c, dx, w, pitch, opt
    
    Set RelativeElem = rElem
End Function

Function AbsoluteElem(slefelem, duration As Double, minspacing As Double, typs As String, tuneNumber As Integer, Optional options1 As Dictionary = Nothing)
    Dim abselem As New AbsoluteElement
    Dim vABC As VoiceABC
    If (TypeOf slefelem Is vClefProperties) Then
        Set vABC = slefelem.ToVoiceABC
    ElseIf (TypeOf slefelem Is vKeySignature) Then
        Set vABC = slefelem.ToVoiceABC
    ElseIf (TypeOf slefelem Is VoiceABC) Then
        Set vABC = slefelem
    Else
        Set vABC = slefelem.ToVoiceABC
    End If

    abselem.init vABC, duration, minspacing, typs, tuneNumber, options1
    Set AbsoluteElem = abselem
End Function

Public Function BeamElem(stemHeight As Double, typs As String, flat, Optional firstElement As VoiceABC = Nothing) As BeamElem
    Set BeamElem = New BeamElem
    BeamElem.init stemHeight, typs, flat, firstElement
End Function

Public Function BraceElem(voice As VoiceElement, typs As String) As BraceElem
    Dim brelem As New BraceElem
    brelem.init voice, typs
    Set BraceElem = brelem
End Function


Function TempoElem(tempo As vtempoProperties, tuneNumber As Integer) As TempoElement
    Set telem = New TempoElement
    telem.init tempo.ToVoiceABC, tuneNumber
    Set TempoElem = telem
End Function

Function TieElem(options As Dictionary) As TieElem
    Set tieEle = New TieElem
    tieEle.init options
    Set TieElem = tieEle
End Function



Public Function StaffGroupElem(gTextSize As getTextSize) As StaffGroupElement
    Set StaffGroupElem = New StaffGroupElement
    StaffGroupElem.init gTextSize
End Function
