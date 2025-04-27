VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AbsoluteElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit




Public typs As String
''''''''''''''''增加
Public notes As iArray      'MusicNoteItem Array
    'mfingering As String  '指法
    'mnote As String           '高低音
    'mscale As String          '音階
 Public mtempo As String             '拍子
 Public duration As Double      '拍子的延遲長度
    'mtow_fingering As String  '指法 第二個
 Public slurStart As Boolean                '連音符
 Public slurEnd As Boolean                '連音符
 Public tripStart As Boolean           '3 連音開始
 Public tripEnd As Boolean             '3 連音結束
 Public mand As Boolean                  '記錄合音 true為跟後面一音合音 ,false為沒有
 Public what As Integer                 '記錄一些不一樣的指法
 
 
 Public mete As Integer     '拍號設定改變
 Public mete2 As Integer     '拍號設定改變
 Public barsperstaff  As Integer    '設定每行幾小節
  
 Public ln As Integer       '要記錄讀進來的資料是在文件裡的第幾 "行"
 Public WD As Integer      '要記錄貢進來的資料是在文件要的第幾 "字"
 
''''''''''''''''增加

Public x As Double
Public y As Double
Public w As Double     '正位的寬度

Public extce As Double '這是前置的位置
Public children As RelativeElementList
Public heads As RelativeElementList
Public abcelem As VoiceABC
Public tuneNumber As Integer
Public durationClass As Double

Public minspacing As Double
Public extra As RelativeElementList
Public extraw As Double    '前綴元素的寬度 是負值
Public right As RelativeElementList
Public invisible As Boolean
Public bottom As Double
Public top As Double
Public fixed  As vFixedCoords  ' FixedCoords
Public specialY As Dictionary  'SpecialYProperties
Public elemset As iArray  'SVGElement[]
'    counters as  {
'        line as  number
'        measure as  number
'        measureTotal as  number
'        voice as  number
'        note as  number
'    }
Public notePositions As New PointList
Public isClef As Boolean
Public isKeySig As Boolean
Public Beam As BeamElem
Public startChar As Double
Public endChar As Double

Public hint As Boolean
Public startTie As Boolean
Public klass As String


 
Public Sub init(abcelem As VoiceABC, duration As Double, minspacing As Double, typs As String, tuneNumber As Integer, Optional options1 As Dictionary = Nothing)

    ''  console.log("Absolute:",abcelem, duration, minspacing, type, tuneNumber, options)
    If (options1 Is Nothing) Then
        Set options1 = New Dictionary
    End If
    Me.tuneNumber = tuneNumber
    Set Me.abcelem = abcelem
    Me.duration = duration
    If options1("durationClassOveride") <> 0 Then
        Me.durationClass = options1("durationClassOveride")
    Else
        Me.durationClass = Me.duration
    End If
    Me.minspacing = minspacing
    Me.x = 0
    Set Me.children = New RelativeElementList
    Set Me.heads = New RelativeElementList
    Set Me.extra = New RelativeElementList
    Me.extraw = 0
    Me.w = 0
    Set Me.right = New RelativeElementList
    Me.invisible = False
    'Me.bottom = undefined
    'Me.top = undefined
    Me.typs = typs

    '' The following are the dimensions of the fixed part of the element.
    '' That is, the chord text will be a different height depending on lot of factors, but the 8th flag will always be in the same place.
    Set Me.fixed = New vFixedCoords ' { w: 0, t: undefined, b: undefined }  '' there is no x-coord here, because that is set later.

    '' these are the heights of all of the vertical elements that can't be placed until the end of the line.
    '' the vertical order of elements that are above is: tempo, part, volume/dynamic, ending/chord, lyric
    '' the vertical order of elements that are below is: lyric, chord, volume/dynamic
        Dim spY As Dictionary
        Set spY = New Dictionary
        
        spY("tempoHeightAbove") = 0
        spY("partHeightAbove") = 0
        spY("volumeHeightAbove") = 0
        spY("dynamicHeightAbove") = 0
        spY("endingHeightAbove") = 0
        spY("chordHeightAbove") = 0
        spY("lyricHeightAbove") = 0
        spY("lyricHeightBelow") = 0
        spY("chordHeightBelow") = 0
        spY("volumeHeightBelow") = 0
        spY("dynamicHeightBelow") = 0
        Set Me.specialY = spY
    

End Sub

Public Function getFixedCoords() As vFixedCoords
    Set getFixedCoords = vFixedCoords
    getFixedCoords.x = Me.x
    getFixedCoords.w = Me.fixed.w
    getFixedCoords.t = Me.fixed.t
    getFixedCoords.B = Me.fixed.B
    
End Function

Public Sub addExtra(extra As RelativeElement)
    '' used for accidentals, multi-measure rest text,
    '' left-side decorations, gracenote heads,
    '' left annotations, gracenote stems.


    Me.fixed.w = Math.max(Me.fixed.w, extra.dx + extra.w)

    Me.fixed.t = Math.max(Me.fixed.t, extra.top)
    
    Me.fixed.B = Math.min(Me.fixed.B, extra.bottom)
    If (extra.dx < Me.extraw) Then Me.extraw = extra.dx
    Me.extra(Me.extra.Count) = extra
    Me.addChild_ extra
End Sub

Public Sub addHead(head As RelativeElement)

    If (head.dx < Me.extraw) Then Me.extraw = head.dx
    Me.heads.Push head
    Me.addRight head
End Sub

Public Sub addRight(right As RelativeElement)


    '' These are the elements that are the fixed part.
    
    Me.fixed.w = Math.max(Me.fixed.w, right.dx + right.w)
    Me.fixed.t = Math.max(Me.fixed.t, right.top)
    Me.fixed.B = Math.min(Me.fixed.B, right.bottom)
    
    '' if (isNaN(me.fixed.t) || isNaN(me.fixed.b))
    ''  debugger
    If (right.dx + right.w > Me.w) Then Me.w = right.dx + right.w
    Me.right.Push right
    Me.addChild_ right

End Sub

Public Sub addFixed(elem As RelativeElement)
    '' used for elements that can't move relative to other elements after they have been placed.
    '' used for ledger lines, bar numbers, debug msgs, clef, key sigs, time sigs
    Me.addChild_ elem

End Sub

Public Sub addFixedX(elem As RelativeElement)
    '' used for elements that can't move horizontally relative to other elements after they have been placed.
    '' used for parts, tempo, decorations
    Me.addChild_ elem

End Sub

Public Sub addCentered(elem As RelativeElement)
    Dim half As Double
    half = elem.w / 2
    If (-half < Me.extraw) Then Me.extraw = -half
    Set Me.extra(Me.extra.Count) = elem
    If (elem.dx + half > Me.w) Then Me.w = elem.dx + half
    Set Me.right(Me.right.Count) = elem
    Me.addChild_ (elem)
End Sub

'Public Sub setLimit(member, child)
'
'
'    if (!child[member]) return
'    if (!me.specialY[member])
'        me.specialY[member] = child[member]
'    Else
'        me.specialY[member] = Math.max(me.specialY[member], child[member])
'End Sub

Public Sub addChild_(child As RelativeElement)
    ''  console.log("Relative:",child)
    Set child.parent = Me
    Me.children.Push child
    Me.pushTop child.top
    Me.pushBottom child.bottom
'    Me.setLimit "tempoHeightAbove", child
'    Me.setLimit "partHeightAbove", child
'    Me.setLimit "volumeHeightAbove", child
'    Me.setLimit "dynamicHeightAbove", child
'    Me.setLimit "endingHeightAbove", child
'    Me.setLimit "chordHeightAbove", child
'    Me.setLimit "lyricHeightAbove", child
'    Me.setLimit "lyricHeightBelow", child
'    Me.setLimit "chordHeightBelow", child
'    Me.setLimit "volumeHeightBelow", child
'    Me.setLimit "dynamicHeightBelow", child
End Sub

Public Sub pushTop(top As Double)

            Me.top = Math.max(top, Me.top)

End Sub
Public Sub pushBottom(bottom As Double)
            Me.bottom = Math.min(Array(bottom, Me.bottom))
End Sub


Public Sub setX(x As Double)
    Dim i As Integer
    Me.x = x
    For i = 0 To Me.children.Length - 1
        Me.children(i).setX (x)
    Next i
End Sub

Public Sub center(before As AbsoluteElement, after As AbsoluteElement)
    '' Used to center whole rests
    Dim Midpoint As Double
    Midpoint = (after.x - before.x) / 2 + before.x
    Me.x = Midpoint - Me.w / 2
    Dim k As Integer
    For k = 0 To Me.children.Length - 1
        Me.children(k).setX (Me.x)
    Next
End Sub

Public Sub setHint()
    Me.hint = True
End Sub

Public Sub highlight(klass, color)
    'highlight.bind(this)(klass, color)
End Sub

Public Sub unhighlight(klass, color)
    'unhighlight.bind(this)(klass, color)
End Sub

Public Function Clone() As AbsoluteElement
    Dim cope As New AbsoluteElement
    cope.typs = Me.typs
    cope.duration = Me.duration
    cope.x = Me.x
    cope.y = Me.y
    cope.w = Me.w
Set cope.children = Me.children
Set cope.heads = Me.heads
Set cope.abcelem = Me.abcelem
    cope.tuneNumber = Me.tuneNumber
    cope.durationClass = Me.durationClass
    
    cope.minspacing = Me.minspacing
Set cope.extra = Me.extra
    cope.extraw = Me.extraw
Set cope.right = Me.right
    cope.invisible = Me.invisible
    cope.bottom = Me.bottom
    cope.top = Me.top
Set cope.fixed = Me.fixed
Set cope.specialY = Me.specialY
Set cope.elemset = Me.elemset

Set cope.notePositions = Me.notePositions
    cope.isClef = Me.isClef
    cope.isKeySig = Me.isKeySig
Set cope.Beam = Me.Beam
    cope.startChar = Me.startChar
    cope.endChar = Me.endChar
    
    cope.hint = Me.hint
    cope.startTie = Me.startTie
    cope.klass = Me.klass
    
    Set Clone = cope
    
End Function
