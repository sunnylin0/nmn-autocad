VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtempoProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public typs As String
Public duration As iArray 'Array<double>
Public bpm As Double
Public preString As String
Public postString As String
Public startChar As Integer
Public endChar As Integer
Public suppress As Boolean
Public suppressBpm As Boolean

Private Sub Class_Initialize()
    Set Me.duration = New iArray
End Sub


Public Function ToVoiceABC() As VoiceABC
        Set ToVoiceABC = New VoiceABC
        
   
        ToVoiceABC.typs = Me.typs
        Set ToVoiceABC.duration = Me.duration
        ToVoiceABC.bpm = Me.bpm
        ToVoiceABC.preString = Me.preString
        ToVoiceABC.postString = Me.postString
        ToVoiceABC.startChar = Me.startChar
        ToVoiceABC.endChar = Me.endChar
        ToVoiceABC.suppress = Me.suppress
        ToVoiceABC.suppressBpm = Me.suppressBpm
        
End Function

