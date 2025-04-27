VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "voicNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public typs As String
Public el_type As String
Public abselem As AbsoluteElement
Public beambr As Double
Public chord As iArray  ' Array<ChordProperties>
Public decoration As iArray ' Array<Decorations> ''TODO enumerate these
Public duration As Double
Public endBeam As Boolean
Public endSlur As Double
Public endTriplet As Boolean
'    fonts as {
'        annotationfont as Font,
'        gchordfont as Font,
'        measurefont as Font,
'        repeatfont as Font,
'        tripletfont as Font,
'        vocalfont As font
'    }
Public gracenotes As vgraceNotes
'    {   acciaccatura As Boolean
'        accidental As AccidentalName
'        duration As Double
'        midipitch As Double
'        name As String
'        pitch As Double
'        verticalPos As Double
'        startBeam As Boolean
'        endBeam As Boolean
'        startSlur as {
'            label As Double
'            direction as 'up', 'down'
'            style as 'dotted'
'        }[]
'        endSlur as double[]
'        startTie as {
'            direction as 'up', 'down'
'            style as 'dotted'
'        }
'        endTie As Boolean
'    }[]
    Public lyric As iArray ' Array<{ syllable: string  divider: ' ' | '-' | '_'  }>
    Public nostem As Boolean
    'Public midiPitches As iArray 'midiPitches
    'Public midiGraceNotePitches As iArray 'MidiGracePitches
    Public pitches As iArray
'    positioning as {
'        chordPosition as 'above' | 'below' | 'hidden'
'        dynamicPosition as 'above' | 'below' | 'hidden'
'        ornamentPosition as 'above' | 'below' | 'hidden'
'        vocalPosition as 'above' | 'below' | 'hidden'
'        volumePosition as 'above' | 'below' | 'hidden'
'    }
    Public Rest As vrest
    
    Public startBeam As Boolean
    Public startSlur As iArray
'    {
'        label As Integer
'        direction as 'up', 'down'
'        style as 'dotted'
'    }[]
    Public startTriplet
    Public tripletMultiplier As Double
    Public tripletR As Double
    Public stemConnectsToAbove As Boolean
    Public style As Dictionary 'NoteHeadType
    Public startChar As Double
    Public endChar As Double


