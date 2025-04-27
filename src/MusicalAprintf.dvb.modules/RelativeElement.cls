VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RelativeElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public typs As String
Public c As String
Public pitch As Double
Public x As Double
Public w As Double
Public dx As Double
Public scalex As Double
Public scaley As Double
Public pitch2 As Double
Public lineWidth As Double
Public klass As String
Public anchor As String '"start" | "middle" | "end"
Public top As Double
Public bottom As Double
Public height As Double
Public Name As String
Public realWidth As Double
Public centerVertically As Boolean
Public parent As AbsoluteElement
Public graphelem As Dictionary  'SVGPathElement
Public stemdir As String
Public highestVert As Double
Public dim_ As Dictionary
' {
'    fontas {
'        faceas string
'        sizeas double
'        decorationas string
'        styleas string
'        weightas string
'        box?
'        paddingas double
'    }
'    attras {
'        'font-size'as double
'        'font-style'as string
'        'font-family'as string
'        'font-weight'as string
'        'text-decoration'as string
'        classas string
'        'text-anchor'as string
'        x As Double
'        y As Double
'        'data-name'as string
'    }
'}

 Public position As String ' 'below' | 'above' | 'hidden' | 'left' | 'right' | 'relative'
 Public partHeightAbove
 Public chordHeightAbove
 Public chordHeightBelow
 Public lyricHeightAbove
 Public lyricHeightBelow
Public lane As Integer

 Public Sub init(c As String, dx As Double, w As Double, pitch As Double, Optional opt As oRelativeOptions = Nothing)
    If (opt Is Nothing) Then Set opt = New oRelativeOptions
    Me.x = 0
    Me.c = c       '' character or path or string
    Me.dx = dx     '' relative x position
    Me.w = w       '' minimum width taken up by this element (can include gratuitous space)
    Me.pitch = pitch  '' relative y position by pitch
    If opt.scalex = 0 Then Me.scalex = 1
    If opt.scaley = 0 Then Me.scaley = 1
    If opt.typs = "" Then Me.typs = "symbol"
    Me.pitch2 = opt.pitch2
    Me.lineWidth = opt.lineWidth
    Me.klass = opt.klass
    If opt.anchor = "" Then Me.anchor = "middle"
    Me.top = pitch
    Me.bottom = pitch
'    if (me.pitch2 !== undefined && me.pitch2 > me.top) me.top = me.pitch2
'    if (me.pitch2 !== undefined && me.pitch2 < me.bottom) me.bottom = me.pitch2
'    if (opt.thickness) {
'        me.top += opt.thickness / 2
'        me.bottom -= opt.thickness / 2
'    }
'    if (opt.stemHeight) {
'        if (opt.stemHeight > 0)
'            me.top += opt.stemHeight
'        Else
'            me.bottom += opt.stemHeight
'    }
'    if (opt.dim)
'        Me.dim = opt.dim
'    if (opt.position)
'        Me.position = opt.position
'    me.height = opt.height ? opt.height : 4  '' The +1 is to give a little bit of padding.
'    if (opt.top)
'        Me.top = opt.top
'    if (opt.bottom)
'        Me.bottom = opt.bottom
'    if (opt.name)
'        Me.name = opt.name
'    else if (me.c)
'        Me.name = Me.c
'    Else
'        Me.name = Me.typs
'    if (opt.realWidth)
'        Me.realWidth = opt.realWidth
'    Else
'        Me.realWidth = Me.w
    Me.centerVertically = False
'    Select Case Me.typs
'        Case "debug":
'            Me.chordHeightAbove = Me.height
'
'        Case "lyric":
'            if (opt.position && opt.position === 'below')
'                Me.lyricHeightBelow = Me.height
'            Else
'                Me.lyricHeightAbove = Me.height
'
'        Case "chord":
'            if (opt.position && opt.position === 'below')
'                Me.chordHeightBelow = Me.height
'            Else
'                Me.chordHeightAbove = Me.height
'
'        Case "text":
'            if (me.pitch === undefined) {
'                if (opt.position && opt.position === 'below')
'                    Me.chordHeightBelow = Me.height
'                Else
'                    Me.chordHeightAbove = Me.height
'            } else
'                Me.centerVertically = True
'
'        Case "part": Me.partHeightAbove = Me.height
'        Case Else
'            Debug.Print "Not between 1 and 10"
'    End Select
End Sub

Public Function getChordDim() As Rect
    If Me.typs = "debug" Then
        Exit Function
    End If
    If IsEmpty(Me.chordHeightAbove) And IsEmpty(Me.chordHeightBelow) Then
        Exit Function
    End If
    '' Chords are centered, annotations are left justified.
    '' NOTE: the font reports extra space to the left and right anyway, so there is a built in margin.
    '' We add a little margin so that items can't touch - we use half the font size as the margin, so that is 1/4 on each side.
    '' if there is only one character that we're printing, use half of that margin.
    '' var margin = me.dim.font.size/4
    '' if (me.c.length === 1)
    ''  margin = margin / 2
    Dim margin As Double
    margin = 0

    Dim Offset As Double
    If Me.typs = "chord" Then
        Offset = Me.realWidth / 2
    Else
        Offset = 0
    End If
    Dim r As Rect
    
    r.left = Me.x - Offset - margin
    r.right = left + Me.realWidth + margin
    getChordDim = r
End Function

Public Sub invertLane(total As Integer)
    Me.lane = total - Me.lane - 1
End Sub

Public Sub putChordInLane(i As Integer)
    Me.lane = i
    '' Add some extra space to account for the character's descenders.
    If IsEmpty(Me.chordHeightAbove) Then
        Me.chordHeightAbove = (Me.height * 1.25) * Me.lane
    Else
        Me.chordHeightBelow = (Me.height * 1.25) * Me.lane
End Sub

Public Function getLane() As Double
    getLane = Me.lane
End Function

Public Sub setX(x As Double)
    Me.x = x + Me.dx
End Sub


Public Function clone() As RelativeElement
    Dim cope As New RelativeElement
    
                               
    cope.typs = Me.typs
    cope.c = Me.c
    cope.pitch = Me.pitch
    cope.x = Me.x
    cope.w = Me.w
    cope.dx = Me.dx
    cope.scalex = Me.scalex
    cope.scaley = Me.scaley
    cope.pitch2 = Me.pitch2
    cope.lineWidth = Me.lineWidth
    cope.klass = Me.klass
    cope.anchor = Me.anchor
    cope.top = Me.top
    cope.bottom = Me.bottom
    cope.height = Me.height
    cope.Name = Me.Name
    cope.realWidth = Me.realWidth
    cope.centerVertically = Me.centerVertically
Set cope.parent = Me.parent
Set cope.graphelem = Me.graphelem
    cope.stemdir = Me.stemdir
    cope.highestVert = Me.highestVert
Set cope.dim_ = Me.dim_
    cope.position = Me.position
    cope.partHeightAbove = Me.partHeightAbove
    cope.chordHeightAbove = Me.chordHeightAbove
    cope.chordHeightBelow = Me.chordHeightBelow
    cope.lyricHeightAbove = Me.lyricHeightAbove
    cope.lyricHeightBelow = Me.lyricHeightBelow
    cope.lane = Me.lane
    Set clone = cope
End Function

