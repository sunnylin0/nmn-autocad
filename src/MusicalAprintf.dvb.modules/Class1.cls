VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public barNumber      As Double
Public brace As BracePosition
Public bracket As BracePosition
Public clef As ClefProperties
Public connectBarLines As BracePosition
Public annotationfont    As vFont
Public gchordfont     As vFont
Public tripletfont    As vFont
Public vocalfont    As vFont
Public key As KeySignature
Public meter As meter
Public spacingAbove    As Double
Public spacingBelow   As Double
Public stafflines   As Double
Public isJianpu As Boolean
Public staffscale    As Double
Public title As iArray
Public voices As VoiceABC    'Array<Array<VoiceItem>>
