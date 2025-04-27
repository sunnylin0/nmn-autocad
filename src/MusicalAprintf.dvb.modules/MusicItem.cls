VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MusicItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

 Public typs As Cg
 Public notes As MusicNoteItemList
    'mfingering As String  '指法
    'mnote As String           '高低音
    'mscale As String          '音階
 Public mtempo As String             '拍子
 Public duration As Double      '拍子的延遲長度
    'mtow_fingering As String  '指法 第二個
 Public accidental As String    '升降記號
 Public a_deco As iArray        '動態指法
 Public slurStart As Boolean           '連音符
 Public slurEnd As Boolean             '連音符
 Public tripStart As Boolean           '3 連音開始
 Public tripEnd As Boolean             '3 連音結束
 Public tripCount As Integer           '是幾連音
 
 Public mand As Boolean                  '記錄合音 true為跟後面一音合音 ,false為沒有
 Public what As Integer                 '記錄一些不一樣的指法
 Public nflags As Integer    '下拍線個數
 Public dots   As Integer    '負點音符個數
 
 
 Public mete As Integer     '拍號設定改變
 Public mete2 As Integer     '拍號設定改變
 Public barsperstaff  As Integer    '設定每行幾小節
 Public barNumber As Integer    '這小節號碼
  
 Public ln As Integer       '要記錄讀進來的資料是在文件裡的第幾 "行"
 Public WD As Integer      '要記錄貢進來的資料是在文件要的第幾 "字"
 
 Public x As Double
 Public y As Double
 Public w As Double     '正位的寬度
 Public extce As Double '這是前置的位置
 Public oX As Double   '最終的位置
 Public oY As Double   '最終的位置
 


Public children As RelativeElementList
Public heads As RelativeElementList
Public abcelem As VoiceABC
Public tuneNumber As Integer
Public durationClass As Double

Public minspacing As Double
Public extra As RelativeElementList
Public extraw As Double    '前綴元素的寬度 是負值
Public extraObjs As MusicNoteItemList  '前綴元素
Public rightObjs As MusicNoteItemList  '後綴元素
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
Public remark As String


'/**
' * get head type, dots, flags of note/rest for a duration
' * 取得一段時間內的頭類型、點、音符 / 休息標誌
' */
Public Function identify_Duration(dur_o As Double)
    
    If dur_o <= 0 Then Exit Function
    Dim flags As Integer
    Dim dots As Long
    Dim dur As Long
        dots = 0
        dur = dur_o

    If (dur Mod 12 <> 0) Then
       Debug.Print Me.typs, "Invalid note duration $1", dur
    End If
    
    dur = dur / 12         ''/* see C.BLEN for values */
    If (dur <= 1) Then Debug.Print 1, Me.typs, "Note too short"
    
'    For flags = 5 To 0 Step -1
'        dur = Math.BitMoveRight(dur, 1)
'        If dur And 1 Then Exit For ''找出奇數
'    Next
    flags = 5
    Do While Not dur And 1
        ''找出奇數
        dur = Math.BitMoveRight(dur, 1)
        flags = flags - 1
        ''If dur And 1 Then Exit Do
    Loop
    
    dur = Math.BitMoveRight(dur, 1)
'    If ((dur + 1) And dur) Then
'        if (s.type != C.REST || dur_o != s.p_v.wmeasure)
'            error(0, s, "Non standard note duration $1", dur_o)
'    End If
    While dur > 0
        dots = dots + 1
        dur = Math.BitMoveRight(dur, dots)
    Wend

    flags = flags - dots
    Me.duration = dur_o
    Me.nflags = CInt(flags)
    Me.dots = dots
End Function

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
'    For i = 0 To Me.children.Count - 1
'        Me.children(i).setX (x)
'    Next i
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

Public Sub calMeWidth()

        Select Case Me.typs
           Case Cg.bar:
                Me.w = amt.wBar * G.FONTSIZE
           Case Cg.meter:
                Me.w = amt.wMete * G.FONTSIZE
           Case Cg.Rest, Cg.note:
                If Me.dots > 0 Then
                    Me.w = 3 * amt.wNote * G.FONTSIZE
                Else
                    Me.w = amt.wNote * G.FONTSIZE
                End If
           Case Else
        End Select
        '計算後綴寬度
        If Not Me.rightObjs Is Nothing Then
            Me.w = Me.w + Me.rightObjs.Count * G.FONTSIZE * amt.extraScale
        End If

End Sub

