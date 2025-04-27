VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vMeter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public el_typs As String
Public typs As String
Public value As vMeterFractionList       ' Array<MeterFraction>;
Public beat_division   As vMeterFractionList   'Array<MeterFraction>;

Public Function ToVoiceABC() As VoiceABC
        Set ToVoiceABC = New VoiceABC
        ToVoiceABC.el_typs = Me.el_typs
        ToVoiceABC.typs = Me.typs
        Set ToVoiceABC.value = Me.value
        Set ToVoiceABC.beat_division = Me.beat_division
        
End Function


