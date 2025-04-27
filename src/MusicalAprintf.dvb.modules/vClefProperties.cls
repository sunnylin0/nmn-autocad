VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vClefProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public typs As String
Public el_typs As String
Public verticalPos As Double
Public clefPos As Double

Public stafflines As Double
Public staffscale As Double
Public transpose As Double

Public Function ToVoiceABC() As VoiceABC
        Set ToVoiceABC = New VoiceABC
        ToVoiceABC.typs = Me.typs
        ToVoiceABC.el_typs = Me.el_typs
        ToVoiceABC.verticalPos = Me.verticalPos
        ToVoiceABC.clefPos = Me.clefPos
        ToVoiceABC.stafflines = Me.stafflines
        ToVoiceABC.staffscale = Me.staffscale
        ToVoiceABC.transpose = Me.transpose


End Function
