VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AbstractEngraver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public decoration As Variant  ' decoration        ''裝飾模組
Public gTextSize As getTextSize  'getTextSize        ''這是計算文字的大小
Public tuneNumber As Double
Public isBagpipes
Public flatBeams
Public graceSlurs As Boolean         ''倚音／Grace notes 圓滑線
Public percmap
Public initialClef As Variant
Public jazzchords As Boolean
Public germanAlphabet As Boolean

Public slurs As TieElemList
Public ties As TieElemList
Public voiceScale As Double
Public voiceColor
Public slursbyvoice     'as  {
'        s0v0?
'        s1v0?
'        s2v0?
'    }
Public tiesbyvoice     'as  {
'        s0v0 As []
'        s1v0 As []
'        s2v0 As []
'    }
Public endingsbyvoice     'as  {
'        s0v0?
'        s1v0?
'        s2v0?
'    }
Public scaleByVoice     'as  {
'        s0v0 As Double
'        s1v0 As Double
'        s2v0 As Double
'    }
'    colorByVoice?
Public tripletMultiplier As Double

Public abcLine As VoiceABCList  'voiceItem []
Public accidentalSlot            ''將保存該和弦上所有臨時記號的清單。
Public accidentalshiftx
Public dotshiftx
Public hasVocals As Boolean
Public minY As Double
Public partstartelem
Public startLimitElem As AbsoluteElement
Public stemdir
Public stemHeight As Double
Public measureLength As Double
Public tempoSet As Boolean
Public tiesSave As TieElemList
Public slursSave As TieElemList
Public slursbyvoiceSave
Public tiesbyvoiceSave
Public triplet As Object 'TripletElem           ''三連音模組

Public style As String

Private isJianpu As Double
Private hint  As Double
Private chartable As Dictionary
Private glyphs As New GlyphsModule




Private Function getDuration(elem As VoiceABC) As Double
    getDuration = elem.duration
End Function

Public Function cloneObject(dest, src)

    Dim prop
    For Each prop In src
'        if (src.hasOwnProperty(prop)) {
'            if (!(Array.isArray(src[prop]) || isObject(src[prop]))) {
'                dest[prop] = src[prop]
'            }
'        }
    Next
End Function


Public Function cloneAbsolute(absSrc As AbsoluteElement)
    Dim returned As New AbsoluteElement
'    Call returned.init("", 0, 0, "", 0)
'    cloneObject(returned, absSrc)
'    returned.top = 0
'    returned.bottom = -1
'    If (absSrc.abcelem) Then
'        returned.abcelem = {}
'        Call cloneObject(returned.abcelem, absSrc.abcelem)
'    End If
    Set cloneAbsolute = returned
End Function

'{
'    jianpu: { 0: "jianpu.0", 1: "jianpu.1", 2: "jianpu.2", 3: "jianpu.3", 4: "jianpu.4", 5: "jianpu.5", 6: "jianpu.6", 7: "jianpu.7", 8: "jianpu.8", 9: "jianpu.9" },
'    rest: { 0: "rests.whole", 1: "rests.half", 2: "rests.quarter", 3: "rests.8th", 4: "rests.16th", 5: "rests.32nd", 6: "rests.64th", 7: "rests.128th", "multi": "rests.multimeasure" },
'    note: { "-1": "noteheads.dbl", 0: "noteheads.whole", 1: "noteheads.half", 2: "noteheads.quarter", 3: "noteheads.quarter", 4: "noteheads.quarter", 5: "noteheads.quarter", 6: "noteheads.quarter", 7: "noteheads.quarter", 'nostem': "noteheads.quarter" },
'    rhythm: { "-1": "noteheads.slash.whole", 0: "noteheads.slash.whole", 1: "noteheads.slash.whole", 2: "noteheads.slash.quarter", 3: "noteheads.slash.quarter", 4: "noteheads.slash.quarter", 5: "noteheads.slash.quarter", 6: "noteheads.slash.quarter", 7: "noteheads.slash.quarter", nostem: "noteheads.slash.nostem" },
'    x: { "-1": "noteheads.indeterminate", 0: "noteheads.indeterminate", 1: "noteheads.indeterminate", 2: "noteheads.indeterminate", 3: "noteheads.indeterminate", 4: "noteheads.indeterminate", 5: "noteheads.indeterminate", 6: "noteheads.indeterminate", 7: "noteheads.indeterminate", nostem: "noteheads.indeterminate" },
'    harmonic: { "-1": "noteheads.harmonic.quarter", 0: "noteheads.harmonic.quarter", 1: "noteheads.harmonic.quarter", 2: "noteheads.harmonic.quarter", 3: "noteheads.harmonic.quarter", 4: "noteheads.harmonic.quarter", 5: "noteheads.harmonic.quarter", 6: "noteheads.harmonic.quarter", 7: "noteheads.harmonic.quarter", nostem: "noteheads.harmonic.quarter" },
'    triangle: { "-1": "noteheads.triangle.quarter", 0: "noteheads.triangle.quarter", 1: "noteheads.triangle.quarter", 2: "noteheads.triangle.quarter", 3: "noteheads.triangle.quarter", 4: "noteheads.triangle.quarter", 5: "noteheads.triangle.quarter", 6: "noteheads.triangle.quarter", 7: "noteheads.triangle.quarter", nostem: "noteheads.triangle.quarter" },
'    uflags: { 3: "flags.u8th", 4: "flags.u16th", 5: "flags.u32nd", 6: "flags.u64th" },
'    dflags: { 3: "flags.d8th", 4: "flags.d16th", 5: "flags.d32nd", 6: "flags.d64th" }
'}

Private Sub Class_Initialize()

    isJianpu = True
    hint = False

    Set chartable = New Dictionary
    Dim dd As New Dictionary
    dd.Add "jianpu", Array("0:jianpu.0", "1:jianpu.1", "2:jianpu.2", "3:jianpu.3", "4:jianpu.4", "5:jianpu.5", "6:jianpu.6", "7:jianpu.7", "8:jianpu.8", "9:jianpu.9")
    dd.Add "rest", Array("0:rests.whole", "1:rests.half", "2:rests.quarter", "3:rests.8th", "4:rests.16th", "5:rests.32nd", "6:rests.64th", "7:rests.128th", "multi:rests.multimeasure")
    dd.Add "note", Array("-1:noteheads.dbl", "0:noteheads.whole", "1:noteheads.half", "2:noteheads.quarter", "3:noteheads.quarter", "4:noteheads.quarter", "5:noteheads.quarter", "6:noteheads.quarter", "7:noteheads.quarter", "nostem:noteheads.quarter")
    dd.Add "rhythm", Array("-1:noteheads.slash.whole", "0:noteheads.slash.whole", "1:noteheads.slash.whole", "2:noteheads.slash.quarter", "3:noteheads.slash.quarter", "4:noteheads.slash.quarter", "5:noteheads.slash.quarter", "6:noteheads.slash.quarter", "7:noteheads.slash.quarter", "nostem:noteheads.slash.nostem")
    dd.Add "x", Array("-1:noteheads.indeterminate", "0:noteheads.indeterminate", "1:noteheads.indeterminate", "2:noteheads.indeterminate", "3:noteheads.indeterminate", "4:noteheads.indeterminate", "5:noteheads.indeterminate", "6:noteheads.indeterminate", "7:noteheads.indeterminate", "nostem:noteheads.indeterminate")
    dd.Add "harmonic", Array("-1:noteheads.harmonic.quarter", "0:noteheads.harmonic.quarter", "1:noteheads.harmonic.quarter", "2:noteheads.harmonic.quarter", "3:noteheads.harmonic.quarter", "4:noteheads.harmonic.quarter", "5:noteheads.harmonic.quarter", "6:noteheads.harmonic.quarter", "7:noteheads.harmonic.quarter", "nostem:noteheads.harmonic.quarter")
    dd.Add "triangle", Array("-1:noteheads.triangle.quarter", "0:noteheads.triangle.quarter", "1:noteheads.triangle.quarter", "2:noteheads.triangle.quarter", "3:noteheads.triangle.quarter", "4:noteheads.triangle.quarter", "5:noteheads.triangle.quarter", "6:noteheads.triangle.quarter", "7:noteheads.triangle.quarter", "nostem:noteheads.triangle.quarter")
    dd.Add "uflags", Array("3:flags.u8th", "4:flags.u16th", "5:flags.u32nd", "6:flags.u64th")
    dd.Add "dflags", Array("3:flags.d8th", "4:flags.d16th", "5:flags.d32nd", "6:flags.d64th")
    
    Dim ii As Integer
    Dim jj As Integer
    Dim distChild As New Dictionary
    Dim arr, arrK
    
    Dim tmpll
    For ii = 0 To dd.Count - 1
        distChild.RemoveAll
        arr = dd.Items(ii)
        For jj = 0 To UBound(arr)
            tmpll = Split(arr(jj), ":")
            distChild.Add tmpll(0), tmpll(1)
        Next
        chartable.Add dd.keys(ii), distChild
    Next
End Sub


''/** 這是取得 note 有連音線的集合，如果不用連音線，則傳回一個元素 */
Public Function getBeamGroup(abcLine As VoiceABCList, pos As Integer) As VoiceABCList
    '' If there are notes beamed together, they are handled as a group, so find all of them here.
    Dim elem As VoiceABC
    Dim group As New VoiceABCList
    Set elem = abcLine(pos)
    If (elem.el_typs <> "note" And Not elem.startBeam And elem.endBeam) Then
        group.Push elem
        Set getBeamGroup = group
        Exit Function
    End If
        
    
    While (pos < abcLine.Count And abcLine(pos).el_typs = "note")
        group.Push abcLine(pos)
        If (abcLine(pos).endBeam) Then
            Exit Function
        End If
        pos = pos + 1
    Wend
    'return { count: group.length, elem: group }
        Set getBeamGroup = group
End Function


Public Sub init(gTextSize As getTextSize, tuneNumber As Integer, options As Dictionary)
        Set Me.decoration = Nothing ' New decoration
        Set Me.gTextSize = gTextSize
        Me.tuneNumber = tuneNumber
'        Me.isBagpipes = options.bagpipes
'        Me.flatBeams = options.flatBeams
'        Me.graceSlurs = options.graceSlurs
'        Me.percmap = options.percmap
'        Me.initialClef = options.initialClef
'        me.jazzchords = !!options.jazzchords
'        me.germanAlphabet = !!options.germanAlphabet
        Me.reset
 End Sub

Public Sub reset()

        Set Me.slurs = New TieElemList
        Set Me.ties = New TieElemList
        Me.voiceScale = 1
'        Me.voiceColor =
'        me.slursbyvoice = {}
'        me.tiesbyvoice = {}
'        me.endingsbyvoice = {}
'        me.scaleByVoice = {}
'        me.colorByVoice = {}
        Me.tripletMultiplier = 1
'
'        Me.abcline = undefined
'        Me.accidentalSlot = undefined
'        Me.accidentalshiftx = undefined
'        Me.dotshiftx = undefined
'        Me.hasVocals = False
'        Me.minY = undefined
'        Me.partstartelem = undefined
'        Me.startlimitelem = undefined
'        Me.stemdir = undefined
End Sub
  
Public Sub setStemHeight(heightInPixels As Double)
        Me.stemHeight = CInt(heightInPixels * 10 / spacing.Step) / 10
End Sub

Public Function getCurrentVoiceId(s As Double, v As Double)
        getCurrentVoiceId = "s" + s + "v" + v
End Function

Public Sub pushCrossLineElems(s As Double, v As Double)
        Set Me.slursbyvoice(Me.getCurrentVoiceId(s, v)) = Me.slurs
        Set Me.tiesbyvoice(Me.getCurrentVoiceId(s, v)) = Me.ties
        Set Me.endingsbyvoice(Me.getCurrentVoiceId(s, v)) = Me.partstartelem
        Set Me.scaleByVoice(Me.getCurrentVoiceId(s, v)) = Me.voiceScale
        If (Me.voiceColor) Then
            Set Me.colorByVoice(Me.getCurrentVoiceId(s, v)) = Me.voiceColor
        End If
End Sub

Public Sub popCrossLineElems(s, v)
        Set Me.slurs = Me.slursbyvoice(Me.getCurrentVoiceId(s, v)) ' || {}
        Set Me.ties = Me.tiesbyvoice(Me.getCurrentVoiceId(s, v)) ' || []
        Set Me.partstartelem = Me.endingsbyvoice(Me.getCurrentVoiceId(s, v))
        Me.voiceScale = Me.scaleByVoice(Me.getCurrentVoiceId(s, v))
        If (Me.voiceScale = 0) Then Me.voiceScale = 1
        Me.voiceColor = Me.colorByVoice(Me.getCurrentVoiceId(s, v))
End Sub

Public Function containsLyrics(staves)   ':Staff[])
        Dim i As Integer, j As Integer, k As Integer
        Dim el As VoiceABC
        For i = 0 To staves.Count - 1
            For j = 0 To staves(i).voices.Count - 1
                For k = 0 To staves(i).voices(j).Count - 1
                    Set el = staves(i).voices(j)(k)
                    If Not (el.lyric Is Nothing) Then
                        '' We just want to see if there are vocals below the music to know where to put the dynamics.
                        If Not (el.positioning Is Nothing) Then
                            If (el.positioning("vocalPosition") = "below") Then Me.hasVocals = True
                        End If
                        Return
                    End If
                Next
            Next
        Next
End Function

Public Function createABCLine(Staffs As StaffList, tempo As vtempoProperties, l As Integer) As StaffGroupElement
        '' PER: This will be the lowest that any note reaches. It will be used to set the dynamics row.
        '' PER：這將是任何音符達到的最低值。 它將用於設定動態行。
        Me.minY = 2
        '' See if there are any lyrics on this line.
        '' 看看這一行是否有歌詞。
        Me.containsLyrics Staffs
        
        Me.tempoSet = False
        Dim staffGroup As StaffGroupElement
        Dim s As Integer
        Set staffGroup = N.StaffGroupElem(Me.gTextSize)
        
        For s = 0 To Staffs.Count - 1
            If (hint) Then Me.restoreState
            hint = False
            ''if (staffs[s].isJianpu == true)
            Me.createABCJianpu staffGroup, Staffs(s), tempo, s, l
            ''else
            ''    me.createABCStaff(staffgroup, staffs[s], tempo, s, l)
        Next
        Set createABCLine = staffGroup
End Function
Public Sub createABCJianpu(staffGroup As StaffGroupElement, abcStaff As Staff, tempo As vtempoProperties, s As Integer, l As Integer)
        '' If the tempo is passed in, then the first element should get the tempo attached to it.
        '' 如果傳入速度，則第一個元素應取得附加的速度。
        staffGroup.gTextSize.updateFonts abcStaff
        Dim v As Integer
        Dim voice As VoiceElement
        Dim clef As AbsoluteElement, keySig As AbsoluteElement, ts As AbsoluteElement
        
        For v = 0 To abcStaff.voices.Count
            Set voice = New VoiceElement
            voice.init v, abcStaff.voices.Count
            voice.isJianpu = True
            If (v = 0) Then
                voice.barfrom = (abcStaff.connectBarLines = "start" And abcStaff.connectBarLines = "continue")
                voice.barto = (abcStaff.connectBarLines = "continue" And abcStaff.connectBarLines = "end")
            Else
                '' duplicate = true  bar lines and other duplicate info need not be created
                '' duplicate = true  空白無資料 無需創建條線和其他重複訊息
                voice.duplicate = True
            End If
            If Not (abcStaff.title Is Nothing) Then
                If (abcStaff.title(v)) Then
                    'line voice.header = abcstaff.title(v).replace(/\\n/g, "\n")
                    voice.headerPosition = 6 + staffGroup.gTextSize.baselineToCenter(voice.header, "voicefont", "staff-extra voice-name", v, abcStaff.voices.Count) / spacing.Step
                End If
            End If
            If Not (abcStaff.clef Is Nothing) Then
                If (abcStaff.clef.typs = "perc") Then voice.isPercussion = True
                If (abcStaff.clef.typs = "jianpu") Then voice.isJianpu = True
            End If
            
            'set clef = (!me.initialClef || l === 0) && createClef(abcstaff.clef, me.tuneNumber)
            If (Me.initialClef = False Or l = 0) Then
             Set clef = createClef(abcStaff.clef, Me.tuneNumber)
            End If
            If Not (clef Is Nothing) Then
                If (v = 0 And abcStaff.barNumber) Then
                    Me.addMeasureNumber abcStaff.barNumber, clef
                End If
                voice.addChild clef
                Set Me.startLimitElem = clef  '' limit ties here
            End If
            Set keySig = createKeySignature(abcStaff.key, Me.tuneNumber)
            If Not (keySig Is Nothing) Then
                voice.addChild keySig
                Set Me.startLimitElem = keySig  '' limit ties here
            End If
            If Not (abcStaff.meter Is Nothing) Then
                If (abcStaff.meter.typs = "specified") Then
                    Me.measureLength = abcStaff.meter.value(0).num / abcStaff.meter.value(0).den
                 Else
                    Me.measureLength = 1
                End If
                Set ts = createTimeSignature(abcStaff.meter, Me.tuneNumber)
                voice.addChild ts
                Set Me.startLimitElem = ts  '' limit ties here
            End If
            If (voice.duplicate) Then
                voice.children = New AbsoluteElementList  '' we shouldn't reprint the above if we're reusing the same staff. We just created them to get the right spacing.
            End If
            Dim stafflines As Integer
            Dim isSingleLineStaff As Boolean
            stafflines = IIf((abcStaff.clef.stafflines And abcStaff.clef.stafflines = 0), abcStaff.clef.stafflines, 5)
            staffGroup.addVoice voice, s, stafflines
            isSingleLineStaff = (stafflines = 1)
            Me.createABCVoice abcStaff.voices(v), tempo, s, v, isSingleLineStaff, voice
            staffGroup.setStaffLimits (voice)
            If (v = 0) Then
                '' only do brace and bracket processing on the first voice, otherwise it would be done twice.
                '' 只對第一個語音進行大括號和方括號處理，否則會進行兩次。
                If (abcStaff.brace = "start" And ((staffGroup.brace Is Nothing) And abcStaff.brace)) Then
                    If (staffGroup.brace Is Nothing) Then
                        Set staffGroup.brace = New BraceElemList
                    End If
                    staffGroup.brace.Push N.BraceElem(voice, "brace")
                 ElseIf (abcStaff.brace = "end" And Not (staffGroup.brace Is Nothing)) Then
                    staffGroup.brace(staffGroup.brace.Count - 1).setBottomStaff voice
                 ElseIf (abcStaff.brace = "continue" And Not (staffGroup.brace Is Nothing)) Then
                    staffGroup.brace(staffGroup.brace.Count - 1).continuing voice
                End If
                If (abcStaff.bracket = "start" Or (staffGroup.bracket Is Nothing And abcStaff.bracket)) Then
                    If (staffGroup.bracket Is Nothing) Then
                        Set staffGroup.bracket = New BraceElemList
                    End If
                    staffGroup.bracket.Push N.BraceElem(voice, "bracket")
                 ElseIf (abcStaff.bracket = "end" And (Not staffGroup.bracket Is Nothing)) Then
                    staffGroup.bracket(staffGroup.bracket.Count - 1).setBottomStaff voice
                 ElseIf (abcStaff.bracket = "continue" And (Not staffGroup.bracket Is Nothing)) Then
                    staffGroup.bracket(staffGroup.bracket.Count - 1).continuing voice
                End If
            End If
        Next
End Sub

Public Sub createABCStaff(staffGroup As StaffGroupElement, abcStaff As vStaffInfo, tempo As vtempoProperties, s As Double, l As Double)
        '' If the tempo is passed in, then the first element should get the tempo attached to it.
        staffGroup.getTextSize.updateFonts (abcStaff)
        Dim v, isSingleLineStaff As Boolean
        Dim voice As VoiceElement, keySig, ts, stafflines
        
        For v = 0 To abcStaff.voices.Count
            Set voice = New VoiceElement
            voice.init v, abcStaff.voices.Count
            If (v = 0) Then
                voice.barfrom = (abcStaff.connectBarLines = "start" And abcStaff.connectBarLines = "continue")
                voice.barto = (abcStaff.connectBarLines = "continue" And abcStaff.connectBarLines = "end")
            Else
                voice.duplicate = True  '' bar lines and other duplicate info need not be created
            End If
            
            If (abcStaff.title) Then
                If (abcStaff.title(v)) Then
                    'line voice.header = abcstaff.title[v].replace(/\\n/g, "\n")
                    'line voice.headerPosition = 6 + staffgroup.getTextSize.baselineToCenter(voice.header, "voicefont", 'staff-extra voice-name', v, abcstaff.voices.length) / spacing.STEP
                End If
            End If
            If Not (abcStaff.clef Is Nothing) Then
                If (abcStaff.clef.typs = "perc") Then
                    voice.isPercussion = True
                End If
                If (abcStaff.clef.typs = "jianpu") Then
                    voice.isJianpu = True
                End If
            End If
            If (Me.initialClef = False Or l = 0) Then
             Set clef = createClef(abcStaff.clef, Me.tuneNumber)
            End If
            If Not (clef Is Nothing) Then
                If (v = 0 And abcStaff.barNumber) Then
                    Me.addMeasureNumber abcStaff.barNumber, clef
                End If
                voice.addChild (clef)
                Me.startLimitElem = clef  '' limit ties here
            End If
            Set keySig = createKeySignature(abcStaff.key, Me.tuneNumber)
            If Not (keySig Is Nothing) Then
                voice.addChild keySig
                Me.startLimitElem = keySig  '' limit ties here
            End If
            If Not (abcStaff.meter Is Nothing) Then
                If (abcStaff.meter.typs = "specified") Then
                    Me.measureLength = abcStaff.meter.value(0).num / abcStaff.meter.value(0).den
                 Else
                    Me.measureLength = 1
                End If
                Set ts = createTimeSignature(abcStaff.meter, Me.tuneNumber)
                voice.addChild (ts)
                Me.startLimitElem = ts  '' limit ties here
            End If
            If (voice.duplicate) Then
               Set voice.children = New AbsoluteElementList  '' we shouldn't reprint the above if we're reusing the same staff. We just created them to get the right spacing.
            End If
            Dim stafflines As Double
            Dim isSingleLineStaff As Boolean
            
            stafflines = IIf((abcStaff.clef.stafflines And abcStaff.clef.stafflines = 0), abcStaff.clef.stafflines, 5)
            staffGroup.addVoice voice, s, stafflines
            isSingleLineStaff = (stafflines = 1)
            Me.createABCVoice abcStaff.voices(v), tempo, s, v, isSingleLineStaff, voice
            staffGroup.setStaffLimits voice
            
  
            If (v = 0) Then
                '' only do brace and bracket processing on the first voice, otherwise it would be done twice.
                '' 只對第一個語音進行大括號和方括號處理，否則會進行兩次。
                If (abcStaff.brace = "start" And Not (staffGroup.brace And abcStaff.brace)) Then
                    If (staffGroup.brace = False) Then
                        Set staffGroup.brace = New iArray
                    End If
                    Dim bEle As New BeamElem
                    bEle.init voice, "brace"
                    staffGroup.brace.Push bEle
                 ElseIf (abcStaff.brace = "end" And staffGroup.brace) Then
                    staffGroup.brace(staffGroup.brace.Length - 1).setBottomStaff voice
                 ElseIf (abcStaff.brace = "continue" And staffGroup.brace) Then
                    staffGroup.brace(staffGroup.brace.Length - 1).continuing voice
                End If
                If (abcStaff.bracket = "start" Or (staffGroup.bracket = False And abcStaff.bracket)) Then
                    If (staffGroup.bracket = False) Then
                        Set staffGroup.bracket = New iArray
                    End If
                    Dim bEle As New BeamElem
                    bEle.init voice, "bracket"
                    staffGroup.bracket.Push bEle
                 ElseIf (abcStaff.bracket = "end" And staffGroup.bracket) Then
                    staffGroup.bracket(staffGroup.bracket.Length - 1).setBottomStaff voice
                 ElseIf (abcStaff.bracket = "continue" And staffGroup.bracket) Then
                    staffGroup.bracket(staffGroup.bracket.Count - 1).continuing voice
                End If
            End If
        Next
End Sub


  Public Sub createABCVoice(abcLine As VoiceABCList, tempo As vtempoProperties, s As Integer, v As Integer, isSingleLineStaff As Boolean, voice As VoiceElement)
'        me.popCrossLineElems(s, v)
'        me.stemdir = (me.isBagpipes) ? "down" : null
        Set Me.abcLine = abcLine
'        if (me.partstartelem) {
'            me.partstartelem = new EndingElem("", null, null)
'            voice.addOther(me.partstartelem)
'        }
'        var voiceNumber = voice.voicetotal < 2 ? -1 : voice.voicenumber
'        for (var slur in me.slurs) {
'            if (me.slurs.hasOwnProperty(slur)) {
'                '' this is already a slur element, but it was created for the last line, so recreate it.
'                '' 這已經是一個 slur 元素，但它是為最後一行建立的，因此重新建立它。
'                me.slurs[slur] = new TieElem({ force: me.slurs[slur].force, voiceNumber: voiceNumber, stemDir: me.slurs[slur].stemDir, style: me.slurs[slur].dotted })
'                if (hint) me.slurs[slur].setHint()
'                voice.addOther(me.slurs[slur])
'            }
'        }
'        for (var i = 0  i < me.ties.length  i++) {
'            '' this is already a tie element, but it was created for the last line, so recreate it.
'            '' 這已經是一個 tie 元素，但它是為最後一行創建的，因此重新創建它。
'            me.ties[i] = new TieElem({ force: me.ties[i].force, stemDir: me.ties[i].stemDir, voiceNumber: voiceNumber, style: me.ties[i].dotted })
'            if (hint) me.ties[i].setHint()
'            voice.addOther(me.ties[i])
'        }
        Dim j As Integer
        For j = 0 To Me.abcLine.Count - 1
            setAveragePitch Me.abcLine(j)
            Me.minY = Math.min(Me.abcLine(j).minPitch, Me.minY)
        Next

        Dim isFirstStaff As Boolean
        If s = 0 Then
            isFirstStaff = True
        Else
            isFirstStaff = False
        End If
        Dim pos As Integer
        Dim elemList As VoiceABCList
        Dim i As Integer
        While (pos < Me.abcLine.Count - 1)
            Debug.Print pos
            If pos = 84 Then
                pos = 84
            End If
            Set elemList = getBeamGroup(Me.abcLine, pos)
            Dim abselems As AbsoluteElementList
            Dim TempoElem As AbsoluteElement
            Set abselems = Me.createABCElement(isFirstStaff, isSingleLineStaff, voice, elemList)
            If Not abselems Is Nothing Then
                For i = 0 To i < abselems.Count
                    If Me.tempoSet <> 0 Or tempo Or tempo.suppress <> 0 Then
                        Me.tempoSet = True
                        
                        Set TempoElem = N.AbsoluteElem(tempo, 0, 0, "tempo", Me.tuneNumber, Nothing)
                        TempoElem.addFixedX N.TempoElem(tempo, Me.tuneNumber)
                        

                        '''' ************
                        'line tempoElement.addFixedX(new TempoElement(tempo, me.tuneNumber, createNoteHead))
                        voice.addChild TempoElem
                    End If
                    voice.addChild abselems(i)
                Next
            End If
            pos = pos + elemList.Count
        Wend
        'me.pushCrossLineElems(s, v)
    End Sub

Public Sub saveState()
'        set me.tiesSave = parseCommon.cloneArray(set me.ties);
'        set me.slursSave = parseCommon.cloneHashOfHash(set me.slurs);
'        set me.slursbyvoiceSave = parseCommon.cloneHashOfHash(set me.slursbyvoice);
'        set me.tiesbyvoiceSave = parseCommon.cloneHashOfArrayOfHash(set me.tiesbyvoice);
End Sub

Public Sub restoreState()
'        set me.ties = parseCommon.cloneArray(set me.tiesSave);
'        set me.slurs = parseCommon.cloneHashOfHash(set me.slursSave);
'        set me.slursbyvoice = parseCommon.cloneHashOfHash(set me.slursbyvoiceSave);
'        set me.tiesbyvoice = parseCommon.cloneHashOfArrayOfHash(set me.tiesbyvoiceSave);
End Sub



    '' return an array of AbsoluteElement
Public Function createABCElement(isFirstStaff As Boolean, isSingleLineStaff As Boolean, voice As VoiceElement, elems As VoiceABCList) As AbsoluteElementList
        Dim elemset As New AbsoluteElementList
        
        If elems.Count > 1 Then
                '' it is undefined if we were passed an array in - an array means a set of notes that should be beamed together.
            Set elemset = Me.createBeam(isSingleLineStaff, voice, elems)
            Set createABCElement = elemset
            Exit Function
        End If
        Dim elem As VoiceABC
        
        Set elem = elems(0)
            
        Select Case elem.el_typs

            Case "note":
                If voice.isJianpu = True Then
                    ''簡譜是否是長拍子
                    
                    ''console.log(`test d.ts Constants.COLORS.SILVER-> ${Constants.COLORS.SILVER}`)
                    ''Object.keys(Constants.COLORS).forEach((ths) => {
                    ''    console.log(`test d.ts Constants.COLORS.SILVER for-> ${ths}`)
                    ''})
                    
                    Set elemset = Me.createLongNote(isSingleLineStaff, voice, elem)
                 Else
                    ''elemset = me.createBeam(isSingleLineStaff, voice, [elem])
                    elemset(0) = Me.createNote(elem, False, isSingleLineStaff, voice)
                    If Not Me.triplet Is Nothing And Me.triplet.isClosed() Then
                        voice.addOther (Me.triplet)
                        Me.triplet = Null
                        Me.tripletMultiplier = 1
                    End If
                End If
            Case "bar":
                elemset(0) = Me.createBarLine(voice, elem, isFirstStaff)
                If (voice.duplicate = True And elemset.Count > 0) Then
                    elemset(0).invisible = True
                
                End If
            Case "meter":
                Set elemset(0) = createTimeSignature(elem, Me.tuneNumber)
                Set Me.startLimitElem = elemset(0)  '' limit ties here
                If (voice.duplicate = True And elemset.Count > 0) Then
                    elemset(0).invisible = True
                End If
            Case "clef":
                 Set elemset(0) = createClef(elem, Me.tuneNumber)
                If (elemset(0) Is Nothing) Then
                    Set createABCElement = Nothing
                    Exit Function
                End If
                If (voice.duplicate = True And elemset.Count > 0) Then
                    elemset(0).invisible = True
                End If
            Case "key":
                Dim absKey As AbsoluteElement
                Set absKey = createKeySignature(elem, Me.tuneNumber)
                If Not (absKey Is Nothing) Then
                    Set elemset(0) = absKey
                    Set Me.startLimitElem = elemset(0)  '' limit ties here
                End If
                If (voice.duplicate = True And elemset.Count > 0) Then elemset(0).invisible = True
                
            Case "stem":
                Me.stemdir = IIf(elem.Direction = "auto", "", elem.Direction)
                
            Case "part":
                Dim abselem As New AbsoluteElement
                Call abselem.init(elem, 0, 0, "part", Me.tuneNumber)
                Dim calsize As size
                calsize = Me.gTextSize.calc(elem.title, "partsfont", "part")
                
                Dim option2 As oRelativeOptions
                option2.typs = "part"
                option2.height = calsize.height / spacing.Step
                
                abselem.addFixedX N.RelativeElem(elem.title, 0, 0, 0, option2)
                Set elemset(0) = abselem
                
'            Case tempo:
'                Dim abselem3 As New AbsoluteElement
'                Dim tEle As New TempoElement
'                abselem3.init elem, 0, 0, "tempo", Me.tuneNumber
'                tEle.init elem, Me.tuneNumber, createNoteHead
'                abselem3.addFixedX tEle
'                Set elemset(0) = abselem3
'            Case VoiceElemTYPS.style:
'                if (elem.head === "normal")
'                    delete Me.style
'                Else
'                    Me.style = elem.head
'                break
'            Case VoiceElemTYPS.hint:
'                hint = True
'                me.saveState()
'                break
'            Case VoiceElemTYPS.midi:
'                '' This has no effect on the visible music, so just skip it.
'                break
'            Case VoiceElemTYPS.scale:
'                Me.voiceScale = elem.Size
'                break
'            Case VoiceElemTYPS.color:
'                Me.voiceColor = elem.color
'                voice.color = Me.voiceColor
'                break

            Case Else
                Dim abselem2 As AbsoluteElement
                Dim tmpDic As New oRelativeOptions
                
                Set abselem2 = N.AbsoluteElem(elem, 0, 0, "unsupported", Me.tuneNumber)
                tmpDic.typs = "debug"
                
                
                abselem2.addFixed N.RelativeElem("element typs " + elem.el_typs, 0, 0, 0, tmpDic)
                Set elemset(0) = abselem2
        End Select
    Set createABCElement = elemset
End Function
Public Function createBeam(isSingleLineStaff As Boolean, voice As VoiceElement, elems As VoiceABCList) As AbsoluteElementList
        Dim abselemset As AbsoluteElementList
        Dim i As Integer
        Dim beamEle As BeamElem
        Dim elem As VoiceABC
        Dim abselem As AbsoluteElement
        Dim tempStemDir As String
        
        Set abselemset = New AbsoluteElementList
        Set beamEle = N.BeamElem(Me.stemHeight * Me.voiceScale, Me.stemdir, Me.flatBeams, elems(0))
        
        For i = 0 To elems.Count - 1
            '' Do a first pass to figure out the stem direction before creating the notes, so that staccatos and other decorations can be placed correctly.
            beamEle.runningDirection elems(i)
        Next
        beamEle.setStemDirection
         tempStemDir = Me.stemdir
        Me.stemdir = IIf(beamEle.stemsUp, "up", "down")
        For i = 0 To elems.Count - 1
            Set elem = elems(i)
            Set abselem = Me.createNote(elem, True, isSingleLineStaff, voice)
            abselemset.Push abselem
            beamEle.Add abselem
            If (Not Me.triplet Is Nothing) Then
                If (Me.triplet.isClosed()) Then
                    voice.addOther (Me.triplet)
                    Me.triplet = Null
                    Me.tripletMultiplier = 1
                End If
            End If
        Next
        beamEle.calcDir
        voice.addBeam beamEle
        Me.stemdir = tempStemDir
        Set createBeam = abselemset
   End Function

'/** 簡譜  超過1拍要加入長拍子 */
Public Function createLongNote(isSingleLineStaff As Boolean, voice As VoiceElement, elem As VoiceABC) As AbsoluteElementList
        Dim abselemset As AbsoluteElementList
        Dim abselem As AbsoluteElement
        Dim longelem As AbsoluteElement
        Dim totalduration As Double
        Dim durlog As Double
        Dim dot As Integer
        Dim cloneelem As VoiceABC
        Dim rOpt As oRelativeOptions
        
        Dim tmpR As RelativeElement
        

        totalduration = getDuration(elem)
        durlog = Math.Floor(Log(totalduration) / Log(2))  ''TODO use getDurlog
        dot = 0
        Set abselemset = New AbsoluteElementList
        If (totalduration < 0.5) Then
           Set abselem = Me.createNote(elem, True, isSingleLineStaff, voice)
            abselemset.Push abselem
         Else
            '' 超過二拍加入長拍子
            Set cloneelem = elem.clone
            cloneelem.duration = 0.25
            Set abselem = Me.createNote(cloneelem, True, isSingleLineStaff, voice)
            abselemset.Push abselem
            totalduration = totalduration - 0.25
            For totalduration = totalduration / 0.25 To 0 Step -1
            
                longelem = cloneAbsolute(abselem)
                Set rOpt = New oRelativeOptions
                rOpt.typs = "beatline"
                rOpt.pitch2 = 2
                rOpt.lineWidth = 2
                rOpt.bottom = 1
                
                longelem.addFixed N.RelativeElem("", 0, 14, 5, rOpt)
                abselemset.Push longelem
      
            Next

        End If

        '' 計算符點音符數量 dot
        ''for (var tot = Math.pow(2, durlog), inc = tot / 2
        ''    tot < duration
        ''    dot++, tot += inc, inc /= 2)

        ''for (durlog  durlog > -2  durlog--) {
        ''    var index = -durlog - 3
        ''    longelem = cloneAbsolute(abselem)
        ''    longelem.addFixed(new RelativeElement(null, 0, 14, 1.5 - index *3, { typs: "beatline", pitch2: 1, linewidth: 1.2, bottom: 1 }))
        ''    abselemset.push(longelem)
        ''}
        If Not (Me.triplet Is Nothing) Then
            If (Me.triplet.isClosed()) Then
                voice.addOther (Me.triplet)
                Me.triplet = Null
                Me.tripletMultiplier = 1
            End If
        End If
        Set createLongNote = abselemset
End Function

Public Function addGraceNotes(elem As VoiceABC, voice As VoiceElement, abselem As AbsoluteElement, notehead As RelativeElement, stemHeight As Double, isBagpipes As Boolean, roomTaken As Double) As Double
        Dim gracescale  As Double
        Dim graceScaleStem  As Double
        Dim gracebeam  As BeamElem
        Dim flag
        Dim i  As Integer
        Dim graceoffsets  As iArray  'number
        Dim tmpRel As RelativeElement
        Dim tmpDic As Dictionary
        
         gracescale = 3 / 5
        graceScaleStem = 3.5 / 5  '' TODO-PER: empirically found constant.
        stemHeight = Math.Round(stemHeight * graceScaleStem)
         
    

        If (elem.gracenotes.Count > 1) Then
            Set gracebeam = New BeamElem
            gracebeam.init stemHeight, "grace", isBagpipes
            If (hint) Then gracebeam.setHint
            Set gracebeam.mainNote = abselem    '' this gives us a reference back to the note this is attached to so that the stems can be attached somewhere.
        End If

        
        Set graceoffsets = New iArray
        For i = elem.gracenotes.Count - 1 To 0 Step -1   '' figure out where to place each gracenote
            roomTaken = roomTaken + 10
            graceoffsets(i) = roomTaken
            If (elem.gracenotes(i).accidental) Then
                roomTaken = roomTaken + 7
            End If
        Next
        Dim gracepitch As Double
        Dim accidentalSlot As iArray
        Dim ret As Double
        Dim pos As Double
        Dim dAcciaccatura As Double
        Dim graceDuration As Double
        Dim pseudoabselem As Double
        Dim isInvisibleRest As Double
        Dim graceDuration As Double
        Dim grace

        For i = 0 To elem.gracenotes.Count - 1
            gracepitch = elem.gracenotes(i).verticalPos

            flag = IIf((Not gracebeam Is Nothing), Empty, chartable.uflags(IIf(isBagpipes, 5, 3)))
            Set accidentalSlot = New iArray
            Set tmpDic = New Dictionary
            tmpDic("dir") = "up"
            tmpDic("headx") = -graceoffsets(i)
            tmpDic("extrax") = -graceoffsets(i)
            tmpDic("flag") = flag
            tmpDic("scale") = gracescale * Me.voiceScale
            tmpDic("accidentalSlot") = accidentalSlot
            
            Set ret = createNoteHead(abselem, "noteheads.quarter", elem.gracenotes(i), tmpDic)
            ret.notehead.highestVert = ret.notehead.pitch + stemHeight
            Set grace = ret.notehead
            Me.addSlursAndTies abselem, elem.gracenotes(i), grace, voice, "up", True

            abselem.addExtra grace
            '' PER: added acciaccatura slash
            If (elem.gracenotes(i).acciaccatura) Then
                pos = elem.gracenotes(i).verticalPos + 7 * gracescale         '' the same formula that determines the flag position.
                dAcciaccatura = IIf(Not gracebeam Is Nothing, 5, 6)         '' just an offset to make it line up correctly.
                
                Set tmpR = New RelativeElement
                Set tmpDic = New Dictionary
                tmpDic("scalex") = gracescale
                tmpDic("scaley") = gracescale
                
                tmpR.init "flags.ugrace", -graceoffsets(i) + dAcciaccatura, 0, pos, tmpDic
                abselem.addRight tmpR
            End If
            If Not (gracebeam Is Nothing) Then '' give the beam the necessary info
                graceDuration = elem.gracenotes(i).duration / 2
                If (isBagpipes) Then graceDuration = graceDuration / 2
                    Set pseudoabselem = New Dictionary
                    Set tmpDic = New Dictionary
                    tmpDic("averagepitch") = gracepitch
                    tmpDic("minpitch") = gracepitch
                    tmpDic("maxpitch") = gracepitch
                    tmpDic("duration") = graceDuration
                    pseudoabselem("heads") = Array(grace)
                    pseudoabselem("abcelem") = tmpDic
                    
                End If
                gracebeam.Add (pseudoabselem)
             Else  '' draw the stem
                Dim p1, p2, dx, width
                p1 = gracepitch + 1 / 3 * gracescale
                p2 = gracepitch + 7 * gracescale
                dx = grace.dx + grace.w
                width = -0.6
                Set tmpR = New RelativeElement
                Set tmpDic = New Dictionary
                tmpDic("typs") = "stem"
                tmpDic("pitch2") = p2
                tmpDic("linewidth") = width
                tmpR.init Null, dx, 0, p1, tmpDic
                abselem.addExtra tmpR
            End If
            ledgerLines abselem, gracepitch, gracepitch, False, glyphs.getSymbolWidth("noteheads.quarter"), Array(), True, grace.dx - 1, 0.6

            '' if this is the first grace note, we might want to start a slur.
            '' there is a slur if graceSlurs is specifically set.
            '' there is no slur if it is bagpipes.
            '' there is not a slur if the element is a spacer or invisible rest.
            Dim isInvisibleRest As Boolean
            If Not elem.rest Is Nothing Then
                isInvisibleRest = (elem.rest.typs = "spacer" Or elem.rest.typs = "invisible")
            Else
                isInvisibleRest = False
            End If
            
            Dim tieE As TieElem
            If (i = 0 And Not isBagpipes And Me.graceSlurs And Not isInvisibleRest) Then
                '' This is the overall slur that is under the grace notes.
                Set tieE = New TieElem
                Set tmpDic = New Dictionary
                tmpDic("anchor1") = grace
                tmpDic("anchor2") = notehead
                tmpDic("isGrace") = True
                tieE.init tmpDic
                
                voice.addOther tieE
            End If
        Next

        If Not (gracebeam Is Nothing) Then
            gracebeam.calcDir
            voice.addBeam gracebeam
        End If
        addGraceNotes = roomTaken
End Function

Public Function addNoteToAbcElement(abselem As AbsoluteElement, elem As VoiceABC, dot As Integer, stemdir As String, style As String, zeroDuration As Boolean, durlog As Double, nostem, voice As VoiceElement) As Dictionary
        Dim dotshiftx As Double '' room taken by chords with displaced noteheads which cause dots to shift
        Dim notehead As RelativeElement
        Dim roomTaken As Double
        Dim roomTakenRight As Double
        Dim min As Double
        Dim i As Double
        Dim additionalLedgers As New iArray
        Dim accidentalSlot As New iArray
        Dim symbolWidth As Double
        Dim dir As String
        Dim noteSymbol As String
        Dim tmpR As RelativeElement
        Dim tmpDic As Dictionary
        
        '' The accidentalSlot will hold a list of all the accidentals on this chord. Each element is a vertical place,
        '' and contains a pitch, which is the last pitch that contains an accidental in that slot. The slots are numbered
        '' from closest to the note to farther left. We only need to know the last accidental we placed because
        '' we know that the pitches are sorted by now.
        ''accidentalSlot 將保存該和弦上所有臨時記號的清單。 每個元素都是一個垂直的地方，
        '' 並包含一個音高，這是該插槽中包含臨時記號的最後一個音高。 插槽已編號
        '' 從最靠近音符到最左邊。 我們只需要知道我們放置的最後一個意外項，因為
        '' 我們知道現在音高已經排序了。

         dir = IIf((elem.averagepitch >= 6), "down", "up")
        If (stemdir <> "") Then dir = stemdir

        style = IIf(elem.style <> "", elem.style, style)  '' get the style of note head.
        If (style = "" Or style = "normal") Then style = "note"

        If (zeroDuration) Then
            noteSymbol = chartable(style).nostem
        Else
            noteSymbol = chartable(style)(CStr(-durlog))
        End If
        If IsEmpty(noteSymbol) Then
            Debug.Print "noteSymbol:"  ''console.log("noteSymbol:", style, durlog, zeroDuration)
        End If
        '' determine elements of chords which should be shifted
        '' 決定應該要移動的和弦元素
        Dim p As Integer
        Dim prev
        Dim curr
        Dim delta As Double
        Dim pp As Integer
        Dim st, stmax, stepa
        st = IIf(dir = "down", elem.pitches.Count - 2, 1)
        stmax = IIf((dir = "down"), p >= 0, p < elem.pitches.Count)
        stepa = IIf((dir = "down"), -1, 1)
        For p = st To stmax Step p = p + stepa
        
            Set prev = elem.pitches(IIf((dir = "down"), p + 1, p - 1))
            Set curr = elem.pitches(p)
            delta = IIf((dir = "down"), prev.pitch - curr.pitch, curr.pitch - prev.pitch)
            If (delta <= 1 And Not prev.printer_shift) Then
                curr.printer_shift = IIf(delta, "different", "same")
                If (curr.verticalPos > 11 Or curr.verticalPos < 1) Then        '' PER: add extra ledger line
                    additionalLedgers.Push curr.verticalPos - (curr.verticalPos Mod 2)
                End If
                If (dir = "down") Then
                    roomTaken = glyphs.getSymbolWidth(noteSymbol) + 2
                 Else
                    dotshiftx = glyphs.getSymbolWidth(noteSymbol) + 2
                End If
            End If
        Next

        '' 疊音製作
        
        Dim flag
        Dim c As String
        Dim percHead
        Dim isTopWhenStemIsDown As Boolean
        Dim isBottomWhenStemIsUp As Boolean
        pp = elem.pitches.Count
        For p = 0 To elem.pitches.Count - 1

            If (Not nostem) Then
                
                If ((dir = "down" And p <> 0) Or (dir = "up" And p <> pp - 1)) Then '' not the stemmed elem of the chord
                    flag = Empty
                 Else
                    flag = chartable(IIf((dir = "down"), "dflags", "uflags"))(-durlog)
                End If
            End If
            c = "0"
            '' There is a style for the whole group of pitches, but there could also be an override for a particular pitch.
            '' 整個音高組都有一個樣式，但也可能有一個特定音高的覆蓋。
            If (elem.pitches(p).style <> "") Then
                c = chartable(elem.pitches(p).style)(-durlog)
             ElseIf (voice.isPercussion And Me.percmap) Then
                c = noteSymbol
                percHead = Me.percmap(pitchesToPerc(elem.pitches(p)))
                If (percHead And percHead.notehead) Then
                    If (chartable(percHead.notehead)) Then
                        c = chartable(percHead.notehead)(-durlog)
                End If
             Else
                c = noteSymbol
            End If
            ''jianpu 改數字譜
            c = (elem.pitches(p).pitch Mod 7) + 1
            elem.pitches(p).verticalPos = 2 + p * 4.5     ''y位置

            '' The highest position for the sake of placing slurs is itself if the slur is internal. It is the highest position possible if the slur is for the whole chord.
            '' If the note is the only one in the chord, then any slur it has counts as if it were on the whole chord.
            '' 如果連線是內部的，則放置連線的最高位置就是連線本身。 如果連線針對整個和弦，則這是可能的最高位置。
            '' 如果該音符是和弦中唯一的音符，則它所具有的任何連線都將被視為在整個和弦上。
            elem.pitches(p).highestVert = elem.pitches(p).verticalPos
            isTopWhenStemIsDown = (stemdir = "up" Or dir = "up") And p = 0
            isBottomWhenStemIsUp = (stemdir = "down" Or dir = "down") And p = pp - 1
            
            If (isTopWhenStemIsDown Or isBottomWhenStemIsUp) Then '' place to put slurs if not already on pitches

                If (Not elem.startSlur Is Nothing And pp = 1) Then
                    elem.pitches(p).highestVert = elem.pitches(pp - 1).verticalPos
                    If (getDuration(elem) < 1 And (stemdir = "up" Or dir = "up")) Then
                        elem.pitches(p).highestVert = elem.pitches(p).highestVert + 6       '' If the stem is up, then compensate for the length of the stem
                    End If
                End If
                If (Not elem.startSlur Is Nothing) Then
                    If (elem.pitches(p).startSlur Is Nothing) Then elem.pitches(p).startSlur = New vSlur  ''TODO possibly redundant, provided array is not optional
                    For i = 0 To elem.startSlur.Count
                        addIfNotExist elem.pitches(p).startSlur, elem.startSlur(i)
                    Next
                End If

                If Not (elem.endSlur Is Nothing) Then
                    elem.pitches(p).highestVert = elem.pitches(pp - 1).verticalPos
                    If (getDuration(elem) < 1 And (stemdir = "up" Or dir = "up")) Then
                        elem.pitches(p).highestVert = elem.pitches(p).highestVert + 6       '' If the stem is up, then compensate for the length of the stem
                    End If
                    If (elem.pitches(p).endSlur Is Nothing) Then elem.pitches(p).endSlur = New iArray    ''TODO possibly redundant, provided array is not optional
                        For i = 0 To elem.endSlur.Count - 1
                            addIfNotExist elem.pitches(p).endSlur, elem.endSlur(i)
                        Next
                    End If
                End If
            End If
            
            Dim hasStem As Boolean
            Dim ret As Dictionary
            Dim noteOpt As oNoteHeadOptions
            hasStem = Not nostem And durlog <= -1
            If (isJianpu) Then
                Set noteOpt = New oNoteHeadOptions
                noteOpt.dir = "down"
                noteOpt.extrax = --roomTaken
                noteOpt.flag = flag
                noteOpt.dot = dot
                noteOpt.dotshiftx = dotshiftx
                noteOpt.scale_ = Me.voiceScale
                Set noteOpt.accidentalSlot = accidentalSlot
                noteOpt.shouldExtendStem = (stemdir <> "")
                noteOpt.printAccidentals = Not voice.isPercussion
                
                Set ret = createNoteHead(abselem, c, elem.pitches(p), noteOpt)
            Else
                Set noteOpt = New oNoteHeadOptions
                noteOpt.dir = "down"
                noteOpt.extrax = --roomTaken
                noteOpt.flag = flag
                noteOpt.dot = dot
                noteOpt.dotshiftx = dotshiftx
                noteOpt.scale_ = Me.voiceScale
                Set noteOpt.accidentalSlot = accidentalSlot
                noteOpt.shouldExtendStem = (stemdir <> "")
                noteOpt.printAccidentals = Not voice.isPercussion
                
                Set ret = createNoteHead(abselem, c, elem.pitches(p), noteOpt)
            End If
            symbolWidth = Math.max(glyphs.getSymbolWidth(c), symbolWidth)
            abselem.extraw = abselem.extraw - ret("extraLeft")
            Set notehead = ret("notehead")
            If Not (notehead Is Nothing) Then
                Me.addSlursAndTies abselem, elem.pitches(p), notehead, voice, IIf(hasStem, dir, ""), False

                If Not (elem.gracenotes Is Nothing) Then
                    If elem.gracenotes.Count > 0 Then
                        '' If there is a tie to the grace notes, leave a little more room for the note to avoid collisions.
                        '' 如果裝飾音有聯繫，請為音符留出更多空間以避免衝突。
                        notehead.bottom = notehead.bottom - 1
                    End If
                End If
                abselem.addHead notehead
            End If
            roomTaken = roomTaken + ret("accidentalshiftx")
            roomTakenRight = Math.max(roomTakenRight, ret("dotshiftx"))
        Next

        '' draw stem from the furthest note to a pitch above/below the stemmed note
        '' 將詞幹從最遠的音符繪製到高於/低於詞幹音符的音高
        '' 繪製符桿
        Dim stemHeight As Double
        Dim p1 As Double, p2 As Double
        Dim dx As Double, width As Double
        If (Not isJianpu And hasStem) Then
            stemHeight = CInt(70 * Me.voiceScale) / 10
            p1 = IIf((dir = "down"), elem.minPitch - stemHeight, elem.minPitch + 1 / 3)
            '' PER added stemdir test to make the line meet the note.
            '' PER新增了stemdir測試以使該行符合註解。
            If (p1 > 6 And Not stemdir) Then p1 = 6
            p2 = IIf((dir = "down"), elem.maxPitch - 1 / 3, elem.maxPitch + stemHeight)
            '' PER added stemdir test to make the line meet the note.
            '' PER新增了stemdir測試以使該行符合註解。
            If (p2 < 6 And Not stemdir) Then p2 = 6
            dx = IIf((dir = "down" And abselem.heads.Count = 0), 0, abselem.heads(0).w)
            width = IIf((dir = "down"), 1, -1)
            '' TODO-PER-HACK: One typs of note head has a different placement of the stem. This should be more generically calculated:
            '' TODO-PER-HACK：一種類型的符頭具有不同的符幹位置。 這應該更通用地計算：

            If (notehead And notehead.c = "noteheads.slash.quarter") Then
                If (dir = "down") Then
                    p2 = p2 - 1
                Else
                    p1 = p1 + 1
                End If
            End If
            Dim rOpt As New oRelativeOptions
            rOpt.typs = "stem"
            rOpt.pitch2 = p2
            rOpt.lineWidth = width
            rOpt.bottom = p1 - 1
             
            abselem.addRight N.RelativeElem("", dx, 0, p1, rOpt)

            min = Math.min(p1, p2)
        End If

        Set tmpDic = New Dictionary
        Set tmpDic("noteHead") = notehead
        tmpDic("roomTaken") = roomTaken
        tmpDic("roomTakenRight") = roomTakenRight
        tmpDic("min") = min
        Set tmpDic("additionalLedgers") = additionalLedgers
        tmpDic("dir") = dir
        tmpDic("symbolWidth") = symbolWidth
        Set addNoteToAbcElement = tmpDic
End Function

'    addNoteToAbcElementOld(abselem: AbsoluteElement, elem: VoiceItem, dot: number, stemdir: StemDirection, style, zeroDuration, durlog, nostem, voice: VoiceElement) {
'        var dotshiftx = 0  '' room taken by chords with displaced noteheads which cause dots to shift
'        var notehead: RelativeElement
'        var roomtaken = 0
'        var roomtakenright = 0
'        var Min
'        var i
'        var additionalLedgers = []
'        '' The accidentalSlot will hold a list of all the accidentals on this chord. Each element is a vertical place,
'        '' and contains a pitch, which is the last pitch that contains an accidental in that slot. The slots are numbered
'        '' from closest to the note to farther left. We only need to know the last accidental we placed because
'        '' we know that the pitches are sorted by now.
'        ''accidentalSlot 將保存該和弦上所有臨時記號的清單。 每個元素都是一個垂直的地方，
'        '' 並包含一個音高，這是該插槽中包含臨時記號的最後一個音高。 插槽已編號
'        '' 從最靠近音符到最左邊。 我們只需要知道我們放置的最後一個意外項，因為
'        '' 我們知道現在音高已經排序了。
'        var accidentalSlot = []
'        var symbolWidth = 0
'
'        var dir = (elem.averagepitch >= 6) ? "down" : "up"
'        if (stemdir) dir = stemdir
'
'        style = elem.style ? elem.style : style  '' get the style of note head.
'        if (!style || style === "normal") style = "note"
'        var noteSymbol
'        if (zeroDuration)
'            noteSymbol = chartable[style].nostem
'        Else
'            noteSymbol = chartable[style][-durlog]
'        if (!noteSymbol)
'            console.log("noteSymbol:", style, durlog, zeroDuration)
'
'        '' determine elements of chords which should be shifted
'        '' 決定應該要移動的和弦元素
'        var p
'        for (p = (dir === "down") ? elem.pitches.length - 2 : 1  (dir === "down") ? p >= 0 : p < elem.pitches.length  p = (dir === "down") ? p - 1 : p + 1) {
'            var prev = elem.pitches[(dir === "down") ? p + 1 : p - 1]
'            var curr = elem.pitches[p]
'            var delta = (dir === "down") ? prev.pitch - curr.pitch : curr.pitch - prev.pitch
'            if (delta <= 1 && !prev.printer_shift) {
'                curr.printer_shift = (delta) ? "different" : "same"
'                if (curr.verticalPos > 11 || curr.verticalPos < 1) {        '' PER: add extra ledger line
'                    additionalLedgers.push(curr.verticalPos - (curr.verticalPos % 2))
'                }
'                if (dir === "down") {
'                    roomtaken = glyphs.getSymbolWidth(noteSymbol) + 2
'                } else {
'                    dotshiftx = glyphs.getSymbolWidth(noteSymbol) + 2
'                }
'            }
'        }
'
'        var pp = elem.pitches.length
'        for (p = 0  p < elem.pitches.length  p++) {
'
'            if (!nostem) {
'                var flag
'                if ((dir === "down" && p !== 0) || (dir === "up" && p !== pp - 1)) { '' not the stemmed elem of the chord
'                    flag = Null
'                } else {
'                    flag = chartable[(dir === "down") ? "dflags" : "uflags"][-durlog]
'                }
'            }
'            var c
'            '' There is a style for the whole group of pitches, but there could also be an override for a particular pitch.
'            '' 整個音高組都有一個樣式，但也可能有一個特定音高的覆蓋。
'            if (elem.pitches[p].style) {
'                c = chartable[elem.pitches[p].style][-durlog]
'            } else if (voice.isPercussion && me.percmap) {
'                c = noteSymbol
'                var percHead = me.percmap[pitchesToPerc(elem.pitches[p])]
'                if (percHead && percHead.noteHead) {
'                    if (chartable[percHead.noteHead])
'                        c = chartable[percHead.noteHead][-durlog]
'                }
'            } else
'                c = noteSymbol
'            ''jianpu 改數字譜
'            c = `jianpu.${(elem.verticalPos % 7) + 1}`
'            elem.pitches [p].verticalPos = 2
'
'            '' The highest position for the sake of placing slurs is itself if the slur is internal. It is the highest position possible if the slur is for the whole chord.
'            '' If the note is the only one in the chord, then any slur it has counts as if it were on the whole chord.
'            '' 如果連線是內部的，則放置連線的最高位置就是連線本身。 如果連線針對整個和弦，則這是可能的最高位置。
'            '' 如果該音符是和弦中唯一的音符，則它所具有的任何連線都將被視為在整個和弦上。
'            elem.pitches[p].highestVert = elem.pitches[p].verticalPos
'            var isTopWhenStemIsDown = (stemdir === "up" || dir === "up") && p === 0
'            var isBottomWhenStemIsUp = (stemdir === "down" || dir === "down") && p === pp - 1
'            if (isTopWhenStemIsDown || isBottomWhenStemIsUp) { '' place to put slurs if not already on pitches
'
'                if (elem.startSlur || pp === 1) {
'                    elem.pitches[p].highestVert = elem.pitches[pp - 1].verticalPos
'                    if (getDuration(elem) < 1 && (stemdir === "up" || dir === "up"))
'                        elem.pitches[p].highestVert += 6         '' If the stem is up, then compensate for the length of the stem
'                }
'                if (elem.startSlur) {
'                    if (!elem.pitches[p].startSlur) elem.pitches[p].startSlur = []  ''TODO possibly redundant, provided array is not optional
'                    for (i = 0  i < elem.startSlur.length  i++) {
'                        addIfNotExist(elem.pitches[p].startSlur, elem.startSlur[i])
'                    }
'                }
'
'                if (elem.endSlur) {
'                    elem.pitches[p].highestVert = elem.pitches[pp - 1].verticalPos
'                    if (getDuration(elem) < 1 && (stemdir === "up" || dir === "up"))
'                        elem.pitches[p].highestVert += 6         '' If the stem is up, then compensate for the length of the stem
'                    if (!elem.pitches[p].endSlur) elem.pitches[p].endSlur = []  ''TODO possibly redundant, provided array is not optional
'                    for (i = 0  i < elem.endSlur.length  i++) {
'                        addIfNotExist(elem.pitches[p].endSlur, elem.endSlur[i])
'                    }
'                }
'            }
'
'            var hasStem = !nostem && durlog <= -1
'            var ret = createNoteHead(abselem, c, elem.pitches[p],
'                { dir: dir, extrax: -roomTaken, flag: flag, dot: dot, dotshiftx: dotshiftx, scale: me.voiceScale, accidentalSlot: accidentalSlot, shouldExtendStem: !stemdir, printAccidentals: !voice.isPercussion })
'            symbolWidth = Math.Max(glyphs.getSymbolWidth(c), symbolWidth)
'            abselem.extraw -= ret.extraLeft
'            notehead = ret.notehead
'            if (noteHead) {
'                me.addSlursAndTies(abselem, elem.pitches[p], noteHead, voice, hasStem ? dir : null, false)
'
'                if (elem.gracenotes && elem.gracenotes.length > 0) {
'                    '' If there is a tie to the grace notes, leave a little more room for the note to avoid collisions.
'                    '' 如果裝飾音有聯繫，請為音符留出更多空間以避免衝突。
'                    notehead.bottom = notehead.bottom - 1
'                }
'                abselem.addHead (notehead)
'            }
'            roomTaken += ret.accidentalshiftx
'            roomtakenright = Math.Max(roomtakenright, ret.dotshiftx)
'        }
'
'        '' draw stem from the furthest note to a pitch above/below the stemmed note
'        '' 將詞幹從最遠的音符繪製到高於/低於詞幹音符的音高
'        '' 繪製符桿
'        if (hasStem) {
'            var stemheight = Math.Round(70 * Me.voiceScale) / 10
'            var p1 = (dir === "down") ? elem.minpitch - stemHeight : elem.minpitch + 1 / 3
'            '' PER added stemdir test to make the line meet the note.
'            '' PER新增了stemdir測試以使該行符合註解。
'            if (p1 > 6 && !stemdir) p1 = 6
'            var p2 = (dir === "down") ? elem.maxpitch - 1 / 3 : elem.maxpitch + stemHeight
'            '' PER added stemdir test to make the line meet the note.
'            '' PER新增了stemdir測試以使該行符合註解。
'            if (p2 < 6 && !stemdir) p2 = 6
'            var dx = (dir === "down" || abselem.heads.length === 0) ? 0 : abselem.heads[0].w
'            var width = (dir === "down") ? 1 : -1
'            '' TODO-PER-HACK: One typs of note head has a different placement of the stem. This should be more generically calculated:
'            '' TODO-PER-HACK：一種類型的符頭具有不同的符幹位置。 這應該更通用地計算：
'
'            if (noteHead && noteHead.c === 'noteheads.slash.quarter') {
'                if (dir === 'down')
'                    p2 -= 1
'                Else
'                    p1 += 1
'            }
'            abselem.addRight(new RelativeElement(null, dx, 0, p1, { typs: "stem", pitch2: p2, linewidth: width, bottom: p1 - 1 }))
'            ''var RelativeElement = function RelativeElement(c, dx, w, pitch, opt) {
'            Min = Math.Min(p1, p2)
'        }
'        return { noteHead: noteHead, roomTaken: roomTaken, roomTakenRight: roomTakenRight, min: min, additionalLedgers: additionalLedgers, dir: dir, symbolWidth: symbolWidth }
'    }

Public Sub addLyric(abselem As AbsoluteElement, elem As VoiceABC)
        Dim lyricStr As String
        Dim div As String
        Dim lyricDim As String
        Dim position As String
        Dim ly
        Dim tmpRel As RelativeElement
        Dim tmpDic As Dictionary
        lyricStr = ""
        
        For Each ly In elem.lyric
            div = IIf(ly.divider = " ", "", ly.divider)
            lyricStr = lyricStr & ly.syllable & div & vbCrLf
        Next
         lyricDim = Me.getTxtSize.calc(lyricStr, "vocalfont", "lyric")
         position = IIf(elem.positioning, elem.positioning.vocalPosition, "below")
        Set tmpRel = New RelativeElement
        Set tmpDic = New Dictionary
        tmpDic("typs") = "lyric"
        tmpDic("position") = position
        tmpDic("height") = lyricDim.height / spacing.Step
        tmpDic("dim") = Me.getTxtSize.attr("vocalfont", "lyric")
        tmpR.init lyricStr, 0, lyricDim.width, Nothing, tmpDic
        
        abselem.addCentered tmpRel
End Sub
'    /**
'         * stem presence: true for drawing stemless notehead
'         * 符幹存在：適用於繪製無符符頭
'         */
Public Function createNote(elem As VoiceABC, nostem As Boolean, isSingleLineStaff As Boolean, voice As VoiceElement) As AbsoluteElement
        Dim notehead As RelativeElement
        Dim roomTaken As Double
        Dim roomTakenRight As Double
        Dim symbolWidth As Double
        Dim additionalLedgers As iArray
        Dim dir As String
        Dim duration As Double
        Dim zeroDuration As Boolean
        Dim durlog As Double
        Dim dot As Integer
        Dim ret3
        
        Set notehead = Nothing
        roomTaken = 0  '' room needed to the left of the note
        roomTakenRight = 0  '' room needed to the right of the note
        symbolWidth = 0
        Set additionalLedgers = New iArray  '' PER: handle the case of [bc'], where the b doesn't have a ledger line

        duration = getDuration(elem)
        zeroDuration = False
        If (duration = 0) Then ''PER: zero duration will draw a quarter note head.
            zeroDuration = True
            duration = 0.25
            nostem = True
        End If
         durlog = Math.Floor(Log(duration) / Log(2))  ''TODO use getDurlog
         dot = 0

        '' 計算符點音符數量 dot
        Dim tot As Double
        Dim Inc As Double
        For tot = Math.Pow(2, durlog) To duration - 1
            Inc = tot / 2
            dot = dot + 1
            tot = tot + Inc
            Inc = Inc / 2
        Next


        If (elem.startTriplet <> 0) Then
            Me.tripletMultiplier = elem.tripletMultiplier
        End If

        Dim durationForSpacing As Double
        Dim abstyps As String
        Dim abselem As AbsoluteElement
        Dim ret1
        Dim ret2
        Dim rOpt As oRelativeOptions
        Dim tmpDic As Dictionary
        Dim tmpRelem As RelativeElement
        
        durationForSpacing = duration * Me.tripletMultiplier
        If Not (elem.rest Is Nothing) Then
            If (elem.rest.typs = "multimeasure") Then
                durationForSpacing = 1
            ElseIf (elem.rest.typs = "invisible-multimeasure") Then
                durationForSpacing = Me.measureLength * elem.rest.text
            End If
        End If
        
        abstyps = IIf(Not elem.rest Is Nothing, "rest", "note")
        
        Set tmpDic = New Dictionary
        tmpDic("durationClassOveride") = elem.duration * Me.tripletMultiplier
        Set abselem = N.AbsoluteElem(elem, durationForSpacing, 1, abstyps, Me.tuneNumber, tmpDic)
        If (hint = True) Then abselem.setHint

        If Not (elem.rest Is Nothing) Then
            If (Me.measureLength = duration And elem.rest.typs <> "invisible" And elem.rest.typs <> "spacer" And InStr(1, elem.rest.typs, "multimeasure") = 0) Then
                '' If the rest is exactly a measure, always use a whole rest
                '' 如果休止符恰好是一個度量，則始終使用整個休止符
                elem.rest.typs = "whole"  '' If the rest is exactly a measure, always use a whole rest
            End If
            Set ret1 = addRestToAbsElement(abselem, elem, duration, dot, voice.voicetotal > 1, Me.stemdir, isSingleLineStaff, durlog, Me.voiceScale)
            Set notehead = ret1.notehead
            roomTaken = ret1.roomTaken
            roomTakenRight = ret1.roomTakenRight
         Else
            '' 重要加入 音符
            Set ret2 = Me.addNoteToAbcElement(abselem, elem, dot, Me.stemdir, Me.style, zeroDuration, durlog, nostem, voice)
            If (ret2("min") <> 0) Then
                Me.minY = Math.min(ret2.min, Me.minY)
            End If
            Set notehead = ret2("noteHead")
            roomTaken = ret2("roomtaken")
            roomTakenRight = ret2("roomtakenright")
            Set additionalLedgers = ret2("additionalLedgers")
            dir = ret2("dir")
            symbolWidth = ret2("symbolWidth")
        End If

        If Not (elem.lyric Is Nothing) Then Me.addLyric abselem, elem
        

        If Not (elem.gracenotes Is Nothing) Then
            roomTaken = roomTaken + Me.addGraceNotes(elem, voice, abselem, notehead, Me.stemHeight * Me.voiceScale, Me.isBagpipes, roomTaken)
        End If
        '' 繪製裝飾
        If Not (elem.decoration Is Nothing) Then
            'line me.decoration.createDecoration(voice, elem.decoration, abselem.top, (notehead) ? notehead.w : 0, abselem, roomtaken, dir, abselem.bottom, elem.positioning, me.hasVocals)
        End If

        If (elem.barNumber <> 0) Then
            Set rOpt = New oRelativeOptions
            rOpt.typs = "barNumber"
            abselem.addFixed N.RelativeElem(elem.barNumber, -10, 0, 0, rOpt)
        End If

        '' ledger lines
        ledgerLines abselem, elem.minPitch, elem.maxPitch, Not (elem.rest Is Nothing), symbolWidth, additionalLedgers, dir, -2, 1

        If Not (elem.chord Is Nothing) Then
            Set ret3 = addChord(Me.gTextSize, abselem, elem, roomTaken, roomTakenRight, symbolWidth, Me.jazzchords, Me.germanAlphabet)
            roomTaken = ret3.roomTaken
            roomTakenRight = ret3.roomTakenRight
        End If

        '' 開始繪製圓滑線
        If (elem.startTriplet = True) Then
            'line me.triplet = new TripletElem(elem.startTriplet, notehead, { flatBeams: me.flatBeams })  '' above is opposite from case of slurs
        End If

        '' 結束圓滑線
        If (elem.endTriplet = True And Not Me.triplet Is Nothing) Then
            'line Me.triplet.setCloseAnchor (notehead)
        End If
        '' 結束圓滑線後加入
        If Not (Me.triplet Is Nothing) And Not elem.startTriplet And Not elem.endTriplet And Not (elem.rest Is Nothing) Then
            If (elem.rest.typs = "spacer") Then
                Me.triplet.middleNote notehead
            End If
        End If

        Set createNote = abselem
End Function

Public Sub addSlursAndTies(abselem As AbsoluteElement, pitchelem, notehead As RelativeElement, voice As VoiceElement, dir As String, isGrace As Boolean)
        Dim found As Boolean
        Dim i, j
        Dim voicenumber As Integer
        Dim tie As TieElem
        
        If (pitchelem.endTie) Then
            If (Me.ties.Count > 0) Then
                '' If there are multiple open ties, find the one that applies by matching the pitch, if possible.
                found = False
                For j = 0 To Me.ties.Count - 1
                    If Not (Me.ties(j).anchor1 Is Nothing) Then
                        If (Me.ties(j).anchor1.pitch = notehead.pitch) Then
                            Me.ties(j).setEndAnchor notehead
                            voice.setRange Me.ties(j)
                            Me.ties.Splice j, 1
                            found = True
                            Exit For
                        End If
                    End If
                Next
                If (Not found) Then
                    Me.ties(0).setEndAnchor notehead
                    voice.setRange Me.ties(0)
                    Me.ties.Splice 0, 1
                End If
            End If
        End If

         voicenumber = IIf(voice.voicetotal < 2, -1, voice.voicenumber)
        If Not (pitchelem.startTie Is Nothing) Then
            Dim tmpDic As Dictionary
            Set tie = New TieElem
            Set tmpDic = New Dictionary
            tmpDic("anchor1") = notehead
            tmpDic("force") = (Me.stemdir = "down" Or Me.stemdir = "up")
            tmpDic("stemDir") = Me.stemdir
            tmpDic("isGrace") = isGrace
            tmpDic("voiceNumber") = voicenumber
            tmpDic("style") = pitchelem.startTie.style
            
            tie.init tmpDic
            If (hint) Then tie.setHint

            Me.ties.Push tie
            voice.addOther tie
            '' HACK-PER: For the animation, we need to know if a note is tied to the next one, so here's a flag.
            '' Unfortunately, only some of the notes in the current event might be tied, but this will consider it
            '' tied if any one of them is. That will work for most cases.
            abselem.startTie = True
        End If

        Dim slur As TieElem
        Dim slurid As Integer
        
        If Not (pitchelem.endSlur Is Nothing) Then
            For i = 0 To pitchelem.endSlur.Count - 1
                slurid = pitchelem.endSlur(i)
                If Not (Me.slurs(slurid) Is Nothing) Then
                    Set slur = Me.slurs(slurid)
                    slur.setEndAnchor notehead
                    voice.setRange slur
                    Set Me.slurs(slurid) = Nothing
                 Else
                    Set slur = New TieElem
                    Set tmpDic = New Dictionary
                    tmpDic("anchor2") = notehead
                    tmpDic("stemDir") = Me.stemdir
                    tmpDic("voiceNumber") = voicenumber
                    
                    slur.init tmpDic
                    If (hint) Then slur.setHint
                    voice.addOther (slur)
                End If
                If Not (Me.startLimitElem Is Nothing) Then
                    slur.setStartX Me.startLimitElem
                End If
            Next
         ElseIf (Not isGrace) Then
'            Dim s
'            For Each s In Me.slurs
'
'                If Not s Is Nothing Then
'                'line 不知如何寫 'if (me.slurs.hasOwnProperty(s)) {
'                    Me.slurs(s).addInternalNote (notehead)
'                End If
'            Next
        End If

        If Not (pitchelem.startSlur Is Nothing) Then
            For i = 0 To pitchelem.startSlur.Count - 1
                slurid = pitchelem.startSlur(i).label
                Set slur = New TieElem
                
                Set slur = New TieElem
                Set tmpDic = New Dictionary
                tmpDic("anchor2") = notehead
                tmpDic("stemDir") = Me.stemdir
                tmpDic("voiceNumber") = voicenumber
                tmpDic("style") = pitchelem.startSlur(i).style
                
                slur.init tmpDic
                
              
                If (hint) Then slur.setHint
                Set Me.slurs(slurid) = slur
                voice.addOther (slur)
            Next
        End If
End Sub

Public Sub addMeasureNumber(Number As Integer, abselem As AbsoluteElement)
        Dim measureNumDim As size
        Dim dx  As Double
        Dim vert As Double
        Dim tmpR As New RelativeElement
        Dim tmpDic As New Dictionary
        measureNumDim = Me.getTxtSize.calc(Number, "measurefont", "bar-number")
        dx = 0
        If (abselem.isClef = True) Then '' If this is a clef rather than bar line, then the number shouldn't be centered because it could overlap the left side. This is an easy way to let it be centered but move it over, too.
            dx = dx + measureNumDim.width / 2
        End If
        vert = IIf(measureNumDim.width > 10 And abselem.abcelem.typs = "treble", 13, 11)
        tmpDic("typs") = "barNumber"
        tmpDic("dim") = Me.getTxtSize.attr("measurefont", "bar-number")
        Call tmpR.init(Number, dx, measureNumDim.width, vert + measureNumDim.height / spacing.Step, tmpDic)
        
        abselem.addFixed tmpR
End Sub

Public Function createBarLine(voice As VoiceElement, elem As VoiceABC, isFirstStaff As Boolean) As AbsoluteElement
        '' bar_thin, bar_thin_thick, bar_thin_thin, bar_thick_thin, bar_right_repeat, bar_left_repeat, bar_double_repeat
        Dim abselem As New AbsoluteElement
        Dim anchor As RelativeElement   '' place to attach part lines
        Dim anchor2 As RelativeElement
        Dim dx As Double
        Dim tmpDic As Dictionary
         dx = 0
         Call abselem.init(elem, 0, 10, "bar", Me.tuneNumber)
        If (elem.barNumber <> Empty) Then
            Call Me.addMeasureNumber(elem.barNumber, abselem)
        End If
        
        Dim firstdots As Boolean
        Dim firstthin As Boolean
        Dim thick As Boolean
        Dim secondthin As Boolean
        Dim seconddots As Boolean
        
         firstdots = (elem.typs = "bar_right_repeat" Or elem.typs = "bar_dbl_repeat")
         firstthin = (elem.typs <> "bar_left_repeat" And elem.typs <> "bar_thick_thin" And elem.typs <> "bar_invisible")
         thick = (elem.typs = "bar_right_repeat" Or elem.typs = "bar_dbl_repeat" Or elem.typs = "bar_left_repeat" Or _
            elem.typs = "bar_thin_thick" Or elem.typs = "bar_thick_thin")
         secondthin = (elem.typs = "bar_left_repeat" Or elem.typs = "bar_thick_thin" Or elem.typs = "bar_thin_thin" Or elem.typs = "bar_dbl_repeat")
         seconddots = (elem.typs = "bar_left_repeat" Or elem.typs = "bar_dbl_repeat")
        
        Dim slur
        '' limit positioning of slurs
        If (firstdots <> Empty Or seconddots <> Empty) Then
            For Each slur In Me.slurs
                If (Me.slurs.hasOwnProperty(slur)) Then
                    Me.slurs [slur].setEndX(abselem)
                End If
            Next
          Set Me.startLimitElem = abselem
        End If

        If (firstdots = True) Then
            anchor = New RelativeElement
            anchor2 = New RelativeElement
            Call anchor.init("dots.dot", dx, 1, 7)
            Call anchor2.init("dots.dot", dx, 1, 5)
            abselem.addRight anchor
            abselem.addRight anchor2
            dx = dx + 6 ''2 hardcoded, twice
        End If

        If (firstthin = True) Then
            anchor = New RelativeElement
            tmpDic = New Dictionary
            tmpDic("typs") = "bar"
            tmpDic("pitch2") = 10
            tmpDic("linewidth") = 0.6
            anchor.init Null, dx, 1, 2, tmpDic
            abselem.addRight anchor
        End If

        If (elem.typs = "bar_invisible") Then
            anchor = New RelativeElement
            tmpDic = New Dictionary
            tmpDic("typs") = "none"
            tmpDic("pitch2") = 10
            tmpDic("linewidth") = 0.6
            anchor.init Null, dx, 1, 2, tmpDic
            abselem.addRight (anchor)
        End If

        If (elem.decoration <> Empty) Then
            'line me.decoration.createDecoration(voice, elem.decoration, 12, (thick) ? 3 : 1, abselem, 0, "down", 2, elem.positioning, me.hasVocals)
        End If

        If (thick = True) Then
            dx = dx + 4 ''3 hardcoded
            anchor = New RelativeElement
            tmpDic = New Dictionary
            tmpDic("typs") = "bar"
            tmpDic("pitch2") = 10
            tmpDic("linewidth") = 4
            anchor.init Null, dx, 4, 2, tmpDic
            abselem.addRight anchor
            dx = dx + 5
        End If

        '' if (me.partstartelem && (thick || (firstthin && secondthin))) { '' means end of nth part
        '' me.partstartelem.anchor2=anchor
        '' me.partstartelem = null
        '' }

        If (Me.partstartelem <> Empty And elem.endEnding <> Empty) Then
            Me.partstartelem.anchor2 = anchor
            Me.partstartelem = Nothing
        End If

        If (secondthin = True) Then
            dx = dx + 3 ''3 hardcoded
            anchor = New RelativeElement
            tmpDic = New Dictionary
            tmpDic("typs") = "bar"
            tmpDic("pitch2") = 10
            tmpDic("linewidth") = 0.6
            anchor.init Nothing, dx, 1, 2, tmpDic
            abselem.addRight anchor '' 3 is hardcoded
        End If

        If (seconddots = True) Then
            dx = dx + 3 ''3 hardcoded
            anchor = New RelativeElement
            anchor2 = New RelativeElement
            Call anchor.init("dots.dot", dx, 1, 7)
            Call anchor2.init("dots.dot", dx, 1, 5)
            abselem.addRight anchor
            abselem.addRight anchor2
        End If '' 2 is hardcoded
        
'line
'        if (elem.startEnding && isFirstStaff) { '' only put the first & second ending marks on the first staff
'            var textWidth = me.GetTxtSize.calc(elem.startEnding, "repeatfont", '').width
'            abselem.minspacing += textWidth + 10  '' Give plenty of room for the ending number.
'            me.partstartelem = new EndingElem(elem.startEnding, anchor, null)
'            voice.addOther (Me.partstartelem)
'        }

        '' Add a little space to the left of the bar line so that nothing can crowd it.
        abselem.extraw = abselem.extraw - 5

        If Not (elem.chord Is Nothing) Then
          'line   var ret3 = addChord(Me.GetTxtSize, abselem, elem, 0, 0, 0, False, Me.germanAlphabet)
        End If

        Set createBarLine = abselem

End Function


'''''''''**************************************************** down function is private

Private Sub setAveragePitch(elem)
    Dim sum As Integer
    Dim p As Integer
    If Not (elem.pitches Is Nothing) Then
        For p = 0 To elem.pitches.Count - 1
            sortPitch elem
        Next
        sum = 0
        For p = 0 To elem.pitches.Count - 1
            sum = sum + elem.pitches(p).verticalPos
        Next
        elem.averagepitch = sum / elem.pitches.Count
        elem.minPitch = elem.pitches(0).verticalPos
        elem.maxPitch = elem.pitches(elem.pitches.Count - 1).verticalPos
    End If
End Sub
Public Sub sortPitch(elem)
    Dim sorted As Boolean
    Dim p As Integer
    Dim tmp As vpitchesNote
    Do
        sorted = True
        For p = 0 To elem.pitches.Count - 1
            If p + 1 >= elem.pitches.Count Then
                sorted = True
            ElseIf (elem.pitches(p).pitch > elem.pitches(p + 1).pitch) Then
                sorted = False
                Set tmp = elem.pitches(p)
                elem.pitches(p) = elem.pitches(p + 1)
                elem.pitches(p + 1) = tmp
            End If
        Next
    Loop While (Not (sorted))
End Sub

Public Sub ledgerLines(abselem As AbsoluteElement, minPitch As Double, maxPitch As Double, isRest As Boolean, _
    symbolWidth As Double, additionalLedgers, dir As String, dx As Double, scale_ As Double)
    Dim i As Double
    Dim ofs As Double
    Dim rOpt As oRelativeOptions
    For i = maxPitch To 11 Step -1
        If (i Mod 2 = 0 And Not isRest) Then
            Set rOpt = New oRelativeOptions
            rOpt.typs = "ledger"
            abselem.addFixed N.RelativeElem("", dx, (symbolWidth + 4) * scale_, i, rOpt)
        End If
    Next

    For i = minPitch To 1
        If (i Mod 2 = 0 And Not isRest) Then
            Set rOpt = New oRelativeOptions
            rOpt.typs = "ledger"
            abselem.addFixed N.RelativeElem("", dx, (symbolWidth + 4) * scale_, i, rOpt)
        End If
    Next

    For i = 0 To additionalLedgers.Count - 1 '' PER: draw additional ledgers
        ofs = symbolWidth
        If (dir = "down") Then ofs = -ofs
            Set rOpt = New oRelativeOptions
            rOpt.typs = "ledger"
        Dim ledger As Double
        ledger = additionalLedgers(i)
        abselem.addFixed N.RelativeElem("", ofs + dx, (symbolWidth + 4) * scale_, ledger, rOpt)
    Next
End Sub

'
'function addRestToAbsJianpuElement(abselem: AbsoluteElement, elem, duration: number, dot: number, isMultiVoice, stemdir:string, isSingleLineStaff, durlog, voiceScale) {
'    var c;
'    var restpitch = 7;
'    var noteHead: RelativeElement;
'    var roomTaken: number;
'    var roomTakenRight: number;
'
'    if (isMultiVoice) {
'        if (stemdir === "down") restpitch = 3;
'        if (stemdir === "up") restpitch = 11;
'    }
'    // There is special placement for the percussion staff. If there is one staff line, then move the rest position.
'    if (isSingleLineStaff) {
'        // The half and whole rests are attached to different lines normally, so we need to tweak their position to get them to both be attached to the same one.
'        if (duration < 0.5)
'            restpitch = 7;
'        else if (duration < 1)
'            restpitch = 7;  // half rest
'        Else
'            restpitch = 5; // whole rest
'    }
'    switch (elem.rest.type) {
'        Case "whole":
'            c = chartable.rest[0];
'            elem.averagepitch = restpitch;
'            elem.minpitch = restpitch;
'            elem.maxpitch = restpitch;
'            dot = 0;
'            break;
'        Case "rest":
'            if (elem.style === "rhythm") // special case for rhythm: rests are a handy way to express the rhythm.
'                c = chartable.rhythm[-durlog];
'            Else
'                c = chartable.rest[-durlog];
'            elem.averagepitch = restpitch;
'            elem.minpitch = restpitch;
'            elem.maxpitch = restpitch;
'            break;
'        Case "invisible":
'        Case "invisible-multimeasure":
'        Case "spacer":
'            c = "";
'            elem.averagepitch = restpitch;
'            elem.minpitch = restpitch;
'            elem.maxpitch = restpitch;
'            break;
'        Case "multimeasure":
'            c = chartable.rest['multi'];
'            elem.averagepitch = restpitch;
'            elem.minpitch = restpitch;
'            elem.maxpitch = restpitch;
'            dot = 0;
'            var mmWidth = glyphs.getSymbolWidth(c);
'            abselem.addHead(new RelativeElement(c, mmWidth, mmWidth * 2, 7));
'            var numMeasures = new RelativeElement("" + elem.rest.text, mmWidth, mmWidth, 16, { type: "multimeasure-text" });
'            abselem.addExtra(numMeasures);
'    }
'    if (elem.rest.type.indexOf("multimeasure") < 0 && elem.rest.type !== "invisible") {
'        var ret = createNoteHead(abselem, c, { verticalPos: restpitch },
'            { dot: dot, scale: voiceScale });
'        noteHead = ret.notehead;
'        if (noteHead) {
'            abselem.addHead(noteHead);
'            roomTaken = ret.accidentalshiftx;
'            roomTakenRight = ret.dotshiftx;
'        }
'    }
'    return { noteHead: noteHead, roomTaken: roomTaken, roomTakenRight: roomTakenRight };
'}

Public Function addRestToAbsElement(abselem As AbsoluteElement, elem, duration As Double, dot As Integer, isMultiVoice As Boolean, stemdir As String, isSingleLineStaff As Boolean, durlog, voiceScale) As Dictionary
    Dim c As String
    Dim restpitch As Integer
    Dim notehead As RelativeElement
    Dim roomTaken As Double
    Dim roomTakenRight As Double
    Dim mmWidth As Double
    Dim numMeasures As RelativeElement
    Dim rOpt As oRelativeOptions
    restpitch = 7

    If (isMultiVoice) Then
        If (stemdir = "down") Then restpitch = 3
        If (stemdir = "up") Then restpitch = 11
    End If
    '' There is special placement for the percussion staff. If there is one staff line, then move the rest position.
    If (isSingleLineStaff) Then
        '' The half and whole rests are attached to different lines normally, so we need to tweak their position to get them to both be attached to the same one.
        If (duration < 0.5) Then
            restpitch = 7
        ElseIf (duration < 1) Then
            restpitch = 7   '' half rest
        Else
            restpitch = 5  '' whole rest
        End If
    End If
    Select Case (elem.rest.typs)
        Case "whole":
            c = chartable.rest(0)
            elem.averagepitch = restpitch
            elem.minPitch = restpitch
            elem.maxPitch = restpitch
            dot = 0
        Case "rest":
            If (elem.style = "rhythm") Then '' special case for rhythm: rests are a handy way to express the rhythm.
                c = chartable.rhythm(-durlog)
            Else
                c = chartable.rest(-durlog)
            End If
            elem.averagepitch = restpitch
            elem.minPitch = restpitch
            elem.maxPitch = restpitch
        Case "invisible":
        Case "invisible-multimeasure":
        Case "spacer":
            c = ""
            elem.averagepitch = restpitch
            elem.minPitch = restpitch
            elem.maxPitch = restpitch
        Case "multimeasure":
            c = chartable.rest("multi")
            elem.averagepitch = restpitch
            elem.minPitch = restpitch
            elem.maxPitch = restpitch
            dot = 0
            mmWidth = glyphs.getSymbolWidth(c)
            abselem.addHead N.RelativeElem(c, mmWidth, mmWidth * 2, 7)
            Set rOpt = New oRelativeOptions
            rOpt.typs = "multimeasure-text"
            Set numMeasures = N.RelativeElem("" + elem.rest.text, mmWidth, mmWidth, 16, rOpt)
            abselem.addExtra numMeasures
    End Select
    If (InStr(1, elem.rest.typs, "multimeasure") < 0 And elem.rest.typs <> "invisible") Then
        Dim ret As Dictionary
        Dim noteOpt As New oNoteHeadOptions
        Dim pitchOpt As New vpitchesNote
        pitchOpt.verticalPos = restpitch
        noteOpt.dot = dot
        noteOpt.scale_ = voiceScale
        Set ret = createNoteHead(abselem, c, pitchOpt, noteOpt)
        notehead = ret.notehead
        If Not (notehead Is Nothing) Then
            abselem.addHead (notehead)
            roomTaken = ret.accidentalshiftx
            roomTakenRight = ret.dotshiftx
        End If
    End If
    Dim outret As New Dictionary
    outret("noteHead") = notehead
    outret("roomTaken") = roomTaken
    outret("roomTakenRight") = roomTakenRight
    Set addRestToAbsElement = outret
    
End Function

Public Sub addIfNotExist(arr, item)
    Dim i As Integer
    For i = 0 To UBound(arr) - 1
        'line 不知要如何改 if (JSON.stringify(arr(i)) === JSON.stringify(item))             Return
    Next
    arr.Push (item)
End Sub
