VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vFormatting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public alignbars As Double
Public aligncomposer As String 'Alignment
Public auquality As String
Public bagpipes As Boolean
Public botmargin As Double
Public botspace As Double
Public bstemdown As Boolean
Public composerspace As Double
Public continueall As Boolean
Public continuous As String
Public dynalign As Boolean
Public exprabove As Boolean
Public exprbelow As Boolean
Public flatBeams As Boolean
Public footer As String
Public freegchord As Boolean
Public gchordbox As Boolean
Public graceSlurs As Boolean
Public gracespacebefore As Double
Public gracespaceinside As Double
Public gracespaceafter As Double
Public header As String
Public INDENT As Double
Public infoline As Boolean
Public infospace As Double
Public leftmargin As Double
Public linesep As Double
Public lineskipfac As Double
Public map As String
Public maxshrink As Double
Public maxstaffsep As Double
Public maxsysstaffsep As Double
Public measurebox As Boolean
Public midi As New Dictionary
'{
'Public barlines As MidiParam
'Public bassprog As MidiParam
'Public bassvol As MidiParam
'Public beat As MidiParam
'Public beataccents As MidiParam
'Public beatmod As MidiParam
'Public beatstring As MidiParam
'Public bendvelocity As MidiParam
'Public c As MidiParam
'Public channel As MidiParam
'Public chordattack As MidiParam
'Public chordname As MidiParam
'Public chordprog As MidiParam
'Public chordvol As MidiParam
'Public control As MidiParam
'Public controlcombo As MidiParam
'Public deltaloudness As MidiParam
'Public drone As MidiParam
'Public droneoff As MidiParam
'Public droneon As MidiParam
'Public drum As MidiParam
'Public drumbars As MidiParam
'Public drummap: MidiParam
'Public drumoff As MidiParam
'Public drumon As MidiParam
'Public expand As MidiParam
'Public fermatafixed As MidiParam
'Public fermataproportional As MidiParam
'Public gchord As MidiParam
'Public gchordon As MidiParam
'Public gchordoff As MidiParam
'Public grace As MidiParam
'Public gracedivider As MidiParam
'Public makechordchannels As MidiParam
'Public nobarlines As MidiParam
'Public nobeataccents As MidiParam
'Public noportamento As MidiParam
'Public pitchbend As MidiParam
'Public program As MidiParam
'Public portamento As MidiParam
'Public ptstress As MidiParam
'Public randomchordattack As MidiParam
'Public ratio As MidiParam
'Public rtranspose As MidiParam
'Public snt As MidiParam
'Public stressmodel As MidiParam
'Public temperamentlinear As MidiParam
'Public temperamentnormal As MidiParam
'Public transpose As MidiParam
'Public trim As MidiParam
'Public volinc: MidiParam
'public   }
Public musicspace As Double
Public nobarcheck As String
Public noteSpacingFactor As Double
Public parskipfac As Double
Public partsbox As Boolean
Public partsspace As Double
Public percmap As New iArray   'Array<PercMapElement>
Public playtempo As String
Public rightmargin As Double
Public scale_ As Double
Public score As String
Public slurHeight As Double
Public splittune As Boolean
Public squarebreve As Boolean
Public staffsep As Double
Public staffwidth As Double
Public stemHeight As Double
Public straightflags As Boolean
Public stretchLast As Double
Public stretchstaff As Boolean
Public subtitlespace As Double
Public sysstaffsep As Double
Public systemsep As Double
Public textspace As Double
Public titleformat As String
Public titleleft As Boolean
Public titlespace As Double
Public topmargin As Double
Public topspace As Double
Public vocalabove As Boolean
Public vocalspace As Double
Public wordsspace As Double
Public annotationfont As vFont
Public composerfont As vFont
Public footerfont As vFont
Public gchordfont As vFont
Public headerfont As vFont
Public historyfont As vFont
Public infofont As vFont
Public measurefont As vFont
Public pageheight As Double
Public pagewidth As Double
Public partsfont As vFont
Public repeatfont As vFont
Public subtitlefont As vFont
Public tabgracefont As vFont
Public tablabelfont As vFont
Public tabnumberfont As vFont
Public tempofont As vFont
Public textfont As vFont
Public titlefont As vFont
Public tripletfont As vFont
Public vocalfont As vFont
Public voicefont As vFont
Public wordsfont As vFont


Public jazzchords As Boolean
