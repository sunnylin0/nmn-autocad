VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Staff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public barNumber As Integer
Public brace As String     'BracePosition
Public bracket As String     'BracePosition
Public clef As vClefProperties
Public connectBarLines As String    'BracePosition
Public annotationfont As vFont
Public gchordfont As vFont
Public tripletfont As vFont
Public vocalfont As vFont
Public key As vKeySignature
Public meter As vMeter
Public spacingAbove As Double
Public spacingBelow As Double
Public stafflines As Integer
Public isJianpu As Boolean
Public staffscale As Double
Public title As iArray    ' Array<string>
Public voices As MusicItemList  'iArray 'VoiceABCList        'Array<Array<VoiceItem>>
