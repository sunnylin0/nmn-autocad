VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataVoiceXX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'
'Public i As Double             ''站存設定現在要作用的子元素
'Public durationIndex As Double ''累記 現在已經讀取每小節的音符長度總合
'Public startX As Double        ''樂器名後 的 x 為開始位置
'Public minx As Double
'Public nextx As Double         ''下一個子元素的位置，是用子元素的時長(duration)*8 開跟號 * spacing
'     ''如子元素 (bar) 沒時長則用 前一次取到的 (note) 時長 spacingduration
'Public spacingduration As Double
'Public header As String     ''每段頭的樂器名
'
'Public children As MusicItemList 'AbsoluteElementList
'Public beams As BeamElemList
'Public otherchildren As iArray 'voiceItem[] '' ties, slurs, triplets
'     '' "GlissandoElem" as    ''滑音
'     '' "CrescendoElem" as    ''漸強
'     '' "DynamicDecoration" as ''動態裝飾
'     '' "TripletElem" as  ''三連音
'     '' "EndingElem" as   ''結尾
'     '' "TieElem" as      ''圓滑線
'Public w As Double
'Public duplicate As Boolean
'Public voicenumber As Double
'Public voicetotal As Double
'Public bottom As Double
'Public top As Double
'Public specialY As Dictionary  ' SpecialYProperties
'Public barfrom As Boolean
'Public barto As Boolean
'Public headerPosition As Double
'Public staffInfo As vStaffInfo
''{
'' top As Integer
'' bottom As Integer
'' lines As Integer
'' voices As Integer[]
'' specialY As SpecialYProperties
'' absoluteY As Integer
'' bottomLine As Integer
'' topLine As Integer
''     }
'
'     ''engraver 加的
'Public isPercussion As Boolean  '' 是否是打擊樂
'Public isJianpu As Boolean      '' 是否是簡譜
'
'
'
'Public Sub init(voicenumber As Integer, voicetotal As Integer)
'
'        Set Me.children = New MusicItemList 'AbsoluteElementList
'        Set Me.beams = New BeamElemList
'        Set Me.otherchildren = New iArray '[]  '' ties, slurs, triplets
'        Me.i = 0
'        Me.w = 0
'        Me.duplicate = False
'        Me.voicenumber = voicenumber  ''number of the voice on a given stave (not staffgroup)
'        Me.voicetotal = voicetotal
'        Me.bottom = 7
'        Me.top = 7
'        Dim spY As Dictionary
'        Set spY = New Dictionary
'
'        spY("tempoHeightAbove") = 0
'        spY("partHeightAbove") = 0
'        spY("volumeHeightAbove") = 0
'        spY("dynamicHeightAbove") = 0
'        spY("endingHeightAbove") = 0
'        spY("chordHeightAbove") = 0
'        spY("lyricHeightAbove") = 0
'        spY("lyricHeightBelow") = 0
'        spY("chordHeightBelow") = 0
'        spY("volumeHeightBelow") = 0
'        spY("dynamicHeightBelow") = 0
'        Set Me.specialY = spY
'
'End Sub
'
'Public Sub addChild(abselem As AbsoluteElement)
'        Dim firstItem As Boolean
'        Dim i As Integer
'        '' This is always passed an AbsoluteElement
'
'        If (abselem.typs = "bar") Then
'            firstItem = True
'            For i = 0 To Me.children.Count
'                If (InStr(1, Me.children(i).typs, "staff-extra") < 0 And Me.children(i).typs <> "tempo") Then
'                    firstItem = False
'                End If
'            Next
'
'            If (firstItem = False) Then
'                Me.beams.Push abselem '("bar")
'                Me.otherchildren.Push abselem '("bar")
'            End If
'        End If
'        Me.children.Push abselem
'        Me.setRange abselem
'End Sub
'
'Public Sub setLimit(member As String, child)
'        '' Sometimes we get an absolute element in here and sometimes we get some type of relative element.
'        '' If there is a "specialY" element, then assume it is an absolute element. If that doesn't exist, look for the
'        '' same members at the top level, because that's where they are in relative elements.
'        Dim specialY As Dictionary
'        Set specialY = child.specialY
'        If Not (specialY Is Nothing) Then Set specialY = child.specialY
'        If (specialY(member) <> 0) Then Exit Sub
'        If (Me.specialY(member) = 0) Then
'            'line Me.specialY(member) = specialY(member)
'        Else
'            Dim aa1, aa2
'            aa1 = Math.max(Me.specialY(member), specialY(member))
'            Set Me.specialY(member) = aa1 ' Math.max(aa1, aa2)
'        End If
'End Sub
'
'Public Sub adjustRange(child)
'        If (child.bottom <> 0) Then
'            Me.bottom = Math.min(Me.bottom, child.bottom)
'        End If
'        If (child.top <> 0) Then Me.top = Math.max(Me.top, child.top)
'End Sub
'
'Public Sub setRange(child)
'        Call Me.adjustRange(child)
'        Call Me.setLimit("tempoHeightAbove", child)
'        Call Me.setLimit("partHeightAbove", child)
'        Call Me.setLimit("volumeHeightAbove", child)
'        Call Me.setLimit("dynamicHeightAbove", child)
'        Call Me.setLimit("endingHeightAbove", child)
'        Call Me.setLimit("chordHeightAbove", child)
'        Call Me.setLimit("lyricHeightAbove", child)
'        Call Me.setLimit("lyricHeightBelow", child)
'        Call Me.setLimit("chordHeightBelow", child)
'        Call Me.setLimit("volumeHeightBelow", child)
'        Call Me.setLimit("dynamicHeightBelow", child)
'End Sub
'
'Public Sub addOther(child)
'        Me.otherchildren.Push child
'        Me.setRange child
'End Sub
'
'Public Sub addBeam(child As BeamElem)
'        Me.beams.Push child
'End Sub
'
'Public Sub setWidth(width As Double)
'        Me.w = width
'End Sub
'
'
