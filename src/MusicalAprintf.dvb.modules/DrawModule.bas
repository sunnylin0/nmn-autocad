Attribute VB_Name = "DrawModule"
Option Explicit

Type Rect
    left As Double
    top As Double
    width As Double
    height As Double
End Type
Type positionxxx
    x As Double
    y As Double
End Type
Type SVGRect
    x As Double
    y As Double
    width As Double
    height As Double
End Type
Type box
    left As Double
    top As Double
    rigth  As Double
    bottom  As Double
End Type

Type size
    width As Double
    height As Double
End Type

'Public Enum positionEnum
'below = "below"
'above = "above"
'Hidden = "hidden"
'left = "left"
'right = "right"
'relative = "relative"
'End Enum

Enum spacing
    FONTEM = 360
    fontsize = 30
    Step = 30 * 93 / 720 '/*spacing.FONTSIZE*/
    space = 10
    TOPNOTE = 15
    STAVEHEIGHT = 100
    INDENT = 50
End Enum



'Enum VoiceElemTYPS
'    note = "note"
'    bar = "bar"
'    meter = "meter"
'    clef = "clef"
'    Key = "key"
'    stem = "stem"
'    part = "part"
'    tempo = "tempo"
'    style = "style"
'    hint = "hint"
'    midi = "midi"
'    'SCALE = "scale"
'    color = "color"
'
'    Gap = "gap"
'    overlay = "overlay"
'    Transpose = "transpose"
'    beam = "beam"
'
'
'End Enum


'Enum AbsoluteTYPS
'    symbol = "symbol"
'    tempo = "tempo"
'    part = "part"
'    rest = "rest"
'    note = "note"
'    bar = "bar"
'    staff_extra = "staff-extra"
'    staff_extra_clef = "staff_extra_clef"
'    staff_extra_key_signature = "staff_extra_key_signature"
'    staff_extra_time_signature = "staff_extra_time_signature"
'End Enum

'Type FixedCoords
'     x As Double
'     y As Double
'     w As Double
'     h As Double
'     t As Double
'     b As Double
'End Type
'Type RelativeOptions
'    typs As String
'    klass As String * 20
'    name As String * 20
'    anchor As String * 20 '"start" | "middle" | "end"
'    position As String
'    dim As Variant
'    pitch2 As Double
'    realWidth As Double
'    lineWidth As Double
'    thickness As Double
'    stemHeight As Double
'    scalex As Double
'    scaley As Double
'    top As Double
'    bottom As Double
'    height As Double
'    width As Double
'End Type


'Type TempoProperties
'    type As String
'    duration As iArray
'    bpm As Double
'    preString As String
'    postString As String
'    startChar As Integer
'    endChar As Integer
'    suppress As Boolean
'    suppressBpm As Boolean
'End If




Public gTuneLine As TuneLine
