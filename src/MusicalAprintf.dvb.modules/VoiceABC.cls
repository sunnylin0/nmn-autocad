VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VoiceABC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public el_typs As String
Public typs As String
    ''VoiceItemClef {
Public stafflines As Double
Public staffscale As Double
Public transpose As Double
Public verticalPos As Double
Public clefPos As Double
Public startChar As Integer
Public endChar As Integer
    '' VoiceItemBar {
Public barNumber As Integer
Public chord As iArray     'Array<ChordProperties>
Public decoration As iArray     'Array<Decorations>
Public endEnding As Boolean
Public startEnding As String
    '' VoiceItemGap {
Public gap As Double
Public accidentals As iArray     'Array<Accidental>;
Public root As String
Public acc As String
Public mode As String
    '' VoiceItemKey extends KeySignature {
    '' VoiceItemMeter extends Meter {
    '' Meter {
Public value As vMeterFractionList     ' Array<MeterFraction>
Public beat_division As vMeterFractionList     '  Array<MeterFraction>
    '' VoiceItemMidi {

Public cmd As Dictionary ' MidiCommands
Public params As iArray    ' Array<string | number>
    '' VoiceItemOverlay {
    ''el_type as  "overlay"
    ''startChar as  number
    ''endChar as  number
    ''overlay as  Array<NoteProperties>
    '' VoiceItemPart {

Public title As String
    '' VoiceItemScale {

Public Size As Double
    '' VoiceItemStem {

Public Direction As String ' StemDirection
    '' VoiceItemStyle {

Public head As String ' NoteHeadType
    '' VoiceItemTempo extends TempoProperties {

    '' TempoProperties {
    ''duration as  Array<number>
Public bpm As Double
Public preString As String
Public postString As String
Public suppress As Boolean
Public suppressBpm As Boolean
    '' VoiceItemTranspose {

Public steps As Double
    '' VoiceItemNote extends NoteProperties {

    '' NoteProperties {'' TODO
    ''type as  string
    ''el_type as  string  ''TODO enumerate these
    ''abselem as  AbsoluteElement

Public beambr As Double
Public duration As Double
Public endBeam As Boolean
Public endSlur As iArray    ' Array<number>
Public endTriplet As Boolean     'true
Public fonts As Dictionary
'{
'        annotationfont as  Font,
'        gchordfont as  Font,
'        measurefont as  Font,
'        repeatfont as  Font,
'        tripletfont as  Font,
'        vocalfont As font
'    }
Public gracenotes As vpitchesNoteList
Public lyric As iArray    'Array<{ syllable as  string  divider as  ' ' | '-' | '_'  }>
Public nostem As Boolean
Public midiPitches As Dictionary ' midiPitches
Public midiGraceNotePitches As Dictionary  ' MidiGracePitches
Public pitches As vpitchesNoteList
'Array<{
'        accidental As AccidentalName
'        pitch As Double
'        name As String
'        startSlur as  {
'            label As Double
'            direction as  'up', 'down'
'            style as  'dotted'
'        }[]
'        endSlur as  Array<number>
'        startTie as  {
'            direction as  'up', 'down'
'            style as  'dotted'
'        }
'        endTie As Boolean
'        verticalPos As Double
'        highestVert As Double
'        style As NoteHeadType
'    }>
Public positioning As Dictionary
'{
'        chordPosition as  'above' | 'below' | 'hidden'
'        dynamicPosition as  'above' | 'below' | 'hidden'
'        ornamentPosition as  'above' | 'below' | 'hidden'
'        vocalPosition as  'above' | 'below' | 'hidden'
'        volumePosition as  'above' | 'below' | 'hidden'
'    }
Public Rest As vrestNote
'{
'        type as  'invisible' | 'spacer' | 'rest' | 'multimeasure' | 'invisible-multimeasure' | 'whole'
' text As Double
'        startTie as  {
'            direction as  'up', 'down'
'            style as  'dotted'
'        }
' endTie As Boolean
'    }
Public startBeam As Boolean
Public startSlur As vSlurList
Public startTriplet As Double
Public tripletMultiplier As Double
Public tripletR As Double
Public stemConnectsToAbove As Boolean    'true
Public style As String  'NoteHeadType
    ''startChar as  number
    ''endChar as  number
    '' VoiceItemBeam {
Public averagepitch As Double
Public minPitch As Double
Public maxPitch As Double
    '' VoiceItemChord extends ChordProperties {
    '' ChordProperties {
Public Name As String
Public position As String 'ChordPlacement
Public rel_position As point
Public dots As Integer


Public Function Clone() As VoiceABC
    Dim cope As New VoiceABC
    cope.averagepitch = Me.averagepitch
    cope.barNumber = Me.barNumber
    cope.beambr = Me.beambr
 Set cope.beat_division = Me.beat_division
    cope.bpm = Me.bpm
   Set cope.chord = Me.chord
    cope.clefPos = Me.clefPos
   Set cope.cmd = Me.cmd
   Set cope.decoration = Me.decoration
    cope.Direction = Me.Direction
    cope.duration = Me.duration
    cope.el_typs = Me.el_typs
    cope.endBeam = Me.endBeam
    cope.endChar = Me.endChar
    cope.endEnding = Me.endEnding
 Set cope.endSlur = Me.endSlur
    cope.endTriplet = Me.endTriplet
 Set cope.fonts = Me.fonts
    cope.gap = Me.gap
 Set cope.gracenotes = Me.gracenotes
    cope.head = Me.head
  Set cope.lyric = Me.lyric
    cope.maxPitch = Me.maxPitch
  Set cope.midiPitches = Me.midiPitches
  Set cope.midiGraceNotePitches = Me.midiGraceNotePitches
    cope.minPitch = Me.minPitch
    cope.Name = Me.Name
    cope.nostem = Me.nostem
   Set cope.params = Me.params
   Set cope.pitches = Me.pitches
    cope.position = Me.position
   Set cope.positioning = Me.positioning
    cope.postString = Me.postString
    cope.preString = Me.preString
  Set cope.rel_position = Me.rel_position
  Set cope.Rest = Me.Rest
    cope.Size = Me.Size
    cope.stafflines = Me.stafflines
    cope.staffscale = Me.staffscale
    cope.startBeam = Me.startBeam
    cope.startChar = Me.startChar
    cope.startEnding = Me.startEnding
  Set cope.startSlur = Me.startSlur
    cope.startTriplet = Me.startTriplet
    cope.stemConnectsToAbove = Me.stemConnectsToAbove
    cope.steps = Me.steps
    cope.style = Me.style
    cope.suppress = Me.suppress
    cope.suppressBpm = Me.suppressBpm
    cope.title = Me.title
    cope.transpose = Me.transpose
    cope.tripletMultiplier = Me.tripletMultiplier
    cope.tripletR = Me.tripletR
    cope.typs = Me.typs
 Set cope.value = Me.value
    cope.verticalPos = Me.verticalPos
        
    Set Clone = cope
    
End Function
