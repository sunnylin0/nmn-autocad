VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vKeySignature"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public el_typs As String
Public accidentals As iArray     'Array<Accidental>;
Public root As String
Public acc As String
Public mode As String

Public Function ToVoiceABC() As VoiceABC
        Set ToVoiceABC = New VoiceABC
        ToVoiceABC.el_typs = Me.el_typs
        Set ToVoiceABC.accidentals = Me.accidentals
        ToVoiceABC.root = Me.root
        ToVoiceABC.acc = Me.acc
        ToVoiceABC.mode = Me.mode
End Function

