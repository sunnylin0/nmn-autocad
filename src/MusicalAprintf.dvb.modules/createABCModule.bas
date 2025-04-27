Attribute VB_Name = "createABCModule"
'Option Explicit
'
'''    abc_create_clef.js
'
'import AbsoluteElement from './elements/absolute-element'
'import glyphs from './glyphs'
'import RelativeElement from './elements/relative-element'

Public N As New newABCModule
Private glyphs As New GlyphsModule
Function createClef(elem As vClefProperties, tuneNumber As Integer) As AbsoluteElement
    Dim clef As String
    Dim octave As Integer
    Dim abselem As AbsoluteElement
    octave = 0
    elem.el_typs = "clef"
    Set abselem = N.AbsoluteElem(elem, 0, 10, "staff-extra clef", tuneNumber)
    abselem.isClef = True
    Select Case elem.typs
        Case "treble": clef = "clefs.G"
        Case "tenor": clef = "clefs.C"
        Case "alto": clef = "clefs.C"
        Case "bass": clef = "clefs.F"
        Case "treble+8": clef = "clefs.G": octave = 1
        Case "tenor+8": clef = "clefs.C": octave = 1
        Case "bass+8": clef = "clefs.F": octave = 1
        Case "alto+8": clef = "clefs.C": octave = 1
        Case "treble-8": clef = "clefs.G": octave = -1
        Case "tenor-8": clef = "clefs.C": octave = -1
        Case "bass-8": clef = "clefs.F": octave = -1
        Case "alto-8": clef = "clefs.C": octave = -1
        Case "none": Set createClef = Nothing: Exit Function
        Case "perc": clef = "clefs.perc"
        Case "jianpu": clef = "clefs.none"
        Case Else:
        Dim rElem As New RelativeElement
        Dim opt As oRelativeOptions
        opt.typs = "debug"
        abselem.addFixed N.RelativeElem("clef=" + elem.typs, 0, 0, 0, opt)
        
        
    End Select

    '' if (elem.verticalPos) {
    '' pitch = elem.verticalPos
    '' }
    Dim dx As Double
    dx = 5
    If (clef <> "") Then
        Dim height As Double
        Dim ofs As Double
        Dim scale_ As Double, adjustspacing As Double, pitch As Double
        Dim top As Double, bottom As Double
        
        height = glyphs.symbolHeightInPitches(clef)
        ofs = clefOffsets(clef)
        Dim opt2  As New oRelativeOptions
        opt2.top = height + elem.clefPos + ofs
        opt2.bottom = elem.clefPos + ofs
        
        abselem.addRight N.RelativeElem(clef, dx, glyphs.getSymbolWidth(clef), elem.clefPos, opt2)
        
        If (octave <> 0) Then
            scale_ = 2 / 3
            adjustspacing = (glyphs.getSymbolWidth(clef) - glyphs.getSymbolWidth("8") * scale_) / 2
            pitch = IIf(octave > 0, abselem.top + 3, abselem.bottom - 1)
            top = IIf(octave > 0, abselem.top + 3, abselem.bottom - 3)
            bottom = top - 2
            If (elem.typs = "bass-8") Then
                '' The placement for bass octave is a little different. It should hug the clef.
                pitch = 3
                adjustspacing = 0
            End If
            Dim opt3 As oRelativeOptions
            opt3.scalex = scale_
            opt3.scaley = scale_
            opt3.top = top
            opt3.bottom = bottom
            abselem.addRight N.RelativeElem("8", dx + adjustspacing, glyphs.getSymbolWidth("8") * scale_, pitch, opt3)
            ''abselem.top += 2
        End If
    End If
    Set createClef = abselem
End Function

Public Function clefOffsets(clef As String) As Integer

    Select Case clef
        Case "clefs.G": clefOffsets = -5: Exit Function
        Case "clefs.C": clefOffsets = -4: Exit Function
        Case "clefs.F": clefOffsets = -4: Exit Function
        Case "clefs.perc": clefOffsets = -2: Exit Function
        Case Else: clefOffsets = 0: Exit Function
    End Select
End Function



Function createKeySignature(elem As vKeySignature, tuneNumber As Integer) As AbsoluteElement
    Dim abselem As AbsoluteElement
    Dim dx As Double, i As Long
    Dim symbol As String
    Dim fudge As Double
    Dim acc As vAccidental
        
    elem.el_typs = "keySignature"
    If Not (elem.accidentals Is Nothing) Then
        If (elem.accidentals.Count = 0) Then
            Exit Function
        End If
    End If
    Set abselem = N.AbsoluteElem(elem, 0, 10, "staff-extra key-signature", tuneNumber)
    abselem.isKeySig = True
    dx = 0
    For i = 0 To elem.accidentals.Count - 1
        Set acc = elem.accidentals(i)
        symbol = ""
        fudge = 0
        Select Case acc.acc
            Case "sharp": symbol = "accidentals.sharp": fudge = -3
            Case "natural": symbol = "accidentals.nat"
            Case "flat": symbol = "accidentals.flat": fudge = -1.2
            Case "quartersharp": symbol = "accidentals.halfsharp": fudge = -2.5
            Case "quarterflat": symbol = "accidentals.halfflat": fudge = -1.2
            Case Else: symbol = "accidentals.flat"
        End Select
        Dim opt As New oRelativeOptions
        opt.thickness = glyphs.symbolHeightInPitches(symbol)
        opt.top = acc.verticalPos + glyphs.symbolHeightInPitches(symbol) + fudge
        opt.bottom = acc.verticalPos + fudge
        
        abselem.addRight N.RelativeElem(symbol, dx, glyphs.getSymbolWidth(symbol), acc.verticalPos, opt)
        dx = dx + glyphs.getSymbolWidth(symbol) + 2
    Next
    Set createKeySignature = abselem
End Function




Function createTimeSignature(elem As vMeter, tuneNumber As Integer) As AbsoluteElement
    Dim abselem As AbsoluteElement
    Dim rElem As RelativeElement
    Dim opt As oRelativeOptions
    Dim x As Double
    Dim i As Integer, i2 As Integer, i3 As Integer
    Dim numWidth As Double, denWidth As Double
    Dim maxWidth As Double, thisWidth As Double
    Dim optO As oRelativeOptions
    
    elem.el_typs = "timeSignature"
    Set abselem = N.AbsoluteElem(elem, 0, 10, "staff-extra time-signature", tuneNumber)
    If (elem.typs = "specified") Then
        x = 0
        For i = 0 To elem.value.Count - 1
            If (i <> 0) Then
                opt.thickness = glyphs.symbolHeightInPitches("+")
                abselem.addRight N.RelativeElem("+", x + 1, glyphs.getSymbolWidth("+"), 6, opt)
                x = x + glyphs.getSymbolWidth("+") + 2
            End If
            If (elem.value(i).den) Then
                numWidth = 0
                Dim ss As String
                ss = CStr(elem.value(i).num)
                For i2 = 0 To Len(ss) - 1
                    numWidth = numWidth + glyphs.getSymbolWidth(Mid(ss, i2 + 1, 1))
                Next
                denWidth = 0
                ss = CStr(elem.value(i).den)
                For i2 = 0 To Len(ss) - 1
                    denWidth = denWidth + glyphs.getSymbolWidth(Mid(ss, i2 + 1, 1))
                Next
                maxWidth = Math.max(numWidth, denWidth)
                Set opt = New oRelativeOptions
                opt.thickness = glyphs.symbolHeightInPitches(elem.value(i).num)
                abselem.addRight N.RelativeElem(elem.value(i).num, x + (maxWidth - numWidth) / 2, numWidth, 8, opt)
                opt.thickness = glyphs.symbolHeightInPitches(elem.value(i).den)
                abselem.addRight N.RelativeElem(elem.value(i).den, x + (maxWidth - denWidth) / 2, denWidth, 4, opt)
                x = x + maxWidth
            Else
                thisWidth = 0
                ss = elem.value(i).num
                For i3 = 0 To Len(ss) - 1
                    thisWidth = thisWidth + glyphs.getSymbolWidth(Mid(ss, i3 + 1, 1))
                Next
                opt.thickness = glyphs.symbolHeightInPitches(elem.value(i).num)
                abselem.addRight N.RelativeElem(elem.value(i).num, x, thisWidth, 6, opt)
                x = x + thisWidth
            End If
        Next
     ElseIf (elem.typs = "common_time") Then
        opt.thickness = glyphs.symbolHeightInPitches("timesig.common")
        abselem.addRight N.RelativeElem("timesig.common", 0, glyphs.getSymbolWidth("timesig.common"), 6, opt)

     ElseIf (elem.typs = "cut_time") Then
        opt.thickness = glyphs.symbolHeightInPitches("timesig.cut")
        abselem.addRight N.RelativeElem("timesig.cut", 0, glyphs.getSymbolWidth("timesig.cut"), 6, opt)
     ElseIf (elem.typs = "tempus_imperfectum") Then
        opt.thickness = glyphs.symbolHeightInPitches("timesig.imperfectum")
        abselem.addRight N.RelativeElem("timesig.imperfectum", 0, glyphs.getSymbolWidth("timesig.imperfectum"), 6, opt)
     ElseIf (elem.typs = "tempus_imperfectum_prolatio") Then
        opt.thickness = glyphs.symbolHeightInPitches("timesig.imperfectum2")
        abselem.addRight N.RelativeElem("timesig.imperfectum2", 0, glyphs.getSymbolWidth("timesig.imperfectum2"), 6, opt)
     ElseIf (elem.typs = "tempus_perfectum") Then
        opt.thickness = glyphs.symbolHeightInPitches("timesig.perfectum")
        abselem.addRight N.RelativeElem("timesig.perfectum", 0, glyphs.getSymbolWidth("timesig.perfectum"), 6, opt)
     ElseIf (elem.typs = "tempus_perfectum_prolatio") Then
        opt.thickness = glyphs.symbolHeightInPitches("timesig.perfectum2")
        abselem.addRight N.RelativeElem("timesig.perfectum2", 0, glyphs.getSymbolWidth("timesig.perfectum2"), 6, opt)
     Else
       Debug.Print ("time signature:" + elem)
    End If
    Set createTimeSignature = abselem
End Function
Function germanNote(note As String) As String
    Dim gn As String
    Select Case (note)
        Case "B#": gn = "H#"
        Case "B＃": gn = "H＃"
        Case "B": gn = "H"
        Case "Bb": gn = "B"
        Case "Bｂ": gn = "B"
        Case eles: gn = note
    End Select
    germanNote = gn

End Function




Function translateChord(chordString, jazzchords, germanAlphabet) As String
    Dim lines
    Dim i As Integer
    Dim chord As String
    Dim baseChord As String
    Dim modifier As String
    Dim bassNote As String
    Dim marker As String
    Dim bass As String
    
    Dim reg
    Dim regEx As New RegExp, matches
    Dim str, MatchContent As String

    strDatePattern = "^((ABCDEFG)(??)?)?((^\/)+)?(\/((ABCDEFG)(#b??)?))?"

    With regEx
        .Global = True      ' 搜索字符串中的全部字符，如果?假，?找到匹配的字符就停止搜索！
        .MultiLine = False  ' 是否指定多行搜索
        .IgnoreCase = True  ' 指定大小寫敏感（True）
        .Pattern = strDatePattern   ' 所匹配的正則
    End With
    lines = Split(chordString, "\n")
    For i = 0 To UBound(lines) - 1
        chord = lines(i)
        '' If the chord isn"t in a recognizable format then just skip it.
        Set reg = regEx.Execute(chord)
        ''reg = chord.match(/^((ABCDEFG)(??)?)?((^\/)+)?(\/((ABCDEFG)(#b??)?))?/)
        If (UBound(reg) >= 4) Then
            baseChord = reg(1).value
            modifier = reg(2).value
            bassNote = reg(4).value
            If (germanAlphabet) Then
                baseChord = germanNote(baseChord)
                bassNote = germanNote(bassNote)
            End If
            '' This puts markers in the pieces of the chord that are read by the svg creator.
            '' After the main part of the chord (the letter, a sharp or flat, and "m") a marker is added. Before a slash a marker is added.
            marker = IIf(jazzchords, "0x03", "")
            bass = IIf(bassNote, "/" + bassNote, "")
            lines(i) = Join(Array(baseChord, modifier, bass), marker)
        End If
    Next
   translateChord = Join(lines, vbCrLf)

End Function

Public Function addChord(gTextSize As getTextSize, abselem As AbsoluteElement, elem As VoiceABC, roomTaken, roomTakenRight, noteheadWidth, jazzchords, germanAlphabet) As Dictionary
    Dim i As Integer, j As Integer
    Dim pos As Double
    Dim rel_position As Double
    Dim chords 'As Array
    Dim chord As String
    Dim x As Double, y As Double
    Dim font As vFont
    Dim klass As String
    Dim attr As Double
    Dim dime As Double
    Dim chordWidth As Double
    Dim chordHeight As Double
    Dim rOpt As oRelativeOptions
    
    
    For i = 0 To elem.chord.Count - 1
        pos = elem.chord(i).position
        rel_position = elem.chord(i).rel_position
        chords = elem.chord(i).Name.Split("\n")
        For j = UBound(chords) - 1 To 0 Step -1    '' parse these in opposite order because we place them from bottom to top.
            chord = chords(j)
            x = 0
            If (pos = "left" Or pos = "right" Or pos = "below" Or pos = "above" Or Not Not (rel_position)) Then
                font = "annotationfont"
                klass = "annotation"
            Else
                font = "gchordfont"
                klass = "chord"
                chord = translateChord(chord, jazzchords, germanAlphabet)
            End If
            attr = gTextSize.attr(font, klass)
            dime = gTextSize.calc(chord, font, klass)
            chordWidth = dime.width
            chordHeight = dime.height / spacing.Step
            Select Case (pos)
                Case "left":
                    roomTaken = roomTaken + chordWidth + 7
                    x = -roomTaken         '' TODO-PER: This is just a guess from trial and error
                    y = elem.averagepitch
                    Set rOpt = New oRelativeOptions
                    rOpt.typs = "text"
                    rOpt.height = chordHeight
                    rOpt.dime = attr
                    rOpt.position = "left"
                    
                    abselem.addExtra N.RelativeElem(chord, x, chordWidth + 4, y, rOpt)
                    break
                Case "right":
                    roomTakenRight = roomTakenRight + 4
                    x = roomTakenRight '' TODO-PER: This is just a guess from trial and error
                    y = elem.averagepitch
                    Set rOpt = New oRelativeOptions
                    rOpt.typs = "text"
                    rOpt.height = chordHeight
                    rOpt.dime = attr
                    rOpt.position = "right"
                    abselem.addRight N.RelativeElem(chord, x, chordWidth + 4, y, rOpt)
                Case "below":
                    '' setting the y-coordinate to undefined for now: it will be overwritten later on, after we figure out what the highest element on the line is.
                    Set rOpt = New oRelativeOptions
                    rOpt.typs = "text"
                    rOpt.height = chordHeight
                    rOpt.dime = attr
                    rOpt.position = "below"
                    rOpt.realWidth = chordWidth
                    
                    abselem.addRight N.RelativeElem(chord, 0, 0, 0, rOpt)

                Case "above":
                    '' setting the y-coordinate to undefined for now: it will be overwritten later on, after we figure out what the highest element on the line is.
                    Set rOpt = New oRelativeOptions
                    rOpt.typs = "text"
                    rOpt.height = chordHeight
                    rOpt.dime = attr
                    rOpt.position = "above"
                    rOpt.realWidth = chordWidth
                    abselem.addRight N.RelativeElem(chord, 0, 0, 0, rOpt)
                Case Else:
                    If (rel_position) Then
                        Dim relPositionY As Double
                        relPositionY = rel_position.y + 3 * spacing.Step  '' TODO-PER: this is a fudge factor to make it line up with abcm2ps
                        Set rOpt = New oRelativeOptions
                        rOpt.typs = "text"
                        rOpt.height = chordHeight
                        rOpt.dime = attr
                        rOpt.position = "relative"
                        abselem.addRight N.RelativeElem(chord, x + rel_position.x, 0, elem.minPitch + relPositionY / spacing.Step, rOpt)
                     Else
                        '' setting the y-coordinate to undefined for now: it will be overwritten later on, after we figure out what the highest element on the line is.
                        Dim pos2 As String
                        pos2 = "above"
                        If Not (elem.positioning Is Nothing) Then
                            If (elem.positioning.chordPosition) Then pos2 = elem.positioning.chordPosition
                        End If
                        If (pos2 <> "hidden") Then
                            Set rOpt = New oRelativeOptions
                            rOpt.typs = "chord"
                            rOpt.height = chordHeight
                            rOpt.dime = attr
                            rOpt.position = pos2
                            rOpt.realWidth = chordWidth
                            abselem.addCentered N.RelativeElem(chord, noteheadWidth / 2, chordWidth, 0, rOpt)
                        End If
                    End If
            End Select
        Next
    Next
    Dim ret As New Dictionary
    ret("roomTaken") = roomTaken
    ret("roomTakenRight") = roomTakenRight
    
    Set addChord = ret
End Function


Function createNoteHead(abselem As AbsoluteElement, c As String, pitchelem, options As oNoteHeadOptions) As Dictionary
    If (options Is Nothing) Then Set options = New oNoteHeadOptions
    Dim dir, headx, extrax, dot, dotshiftx, scale_, accidentalSlot, shouldExtendStem, printAccidentals
    Dim flag As String
    
    dir = IIf(options.dir <> "", options.dir, "")
    headx = IIf(options.headx, options.headx, 0)
    extrax = IIf(options.extrax, options.extrax, 0)
    flag = IIf(options.flag <> "", options.flag, "")
    dot = IIf(options.dot, options.dot, 0)
    dotshiftx = IIf(options.dotshiftx, options.dotshiftx, 0)
    scale_ = IIf(options.scale_, options.scale_, 1)
    Set accidentalSlot = IIf(Not (options.accidentalSlot Is Nothing), options.accidentalSlot, New iArray)
    shouldExtendStem = IIf(options.shouldExtendStem, options.shouldExtendStem, False)
    printAccidentals = IIf(options.printAccidentals, options.printAccidentals, True)

    '' TODO scale the dot as well
    Dim pitch As Double
    Dim notehead As RelativeElement
    Dim accidentalshiftx As Double
    Dim newDotShiftX As Double
    Dim extraLeft As Double
    Dim adjust As Double
    Dim shiftheadx As Double
    Dim pos As Double
    Dim xdelta As Double
    Dim dotadjusty As Double
    Dim rOpt As oRelativeOptions
    Dim opts As oRelativeOptions
    
    pitch = pitchelem.verticalPos
    accidentalshiftx = 0
    newDotShiftX = 0
    extraLeft = 0
    If c = Empty Then
        Set rOpt = New oRelativeOptions
        rOpt.typs = "debug"
        abselem.addFixed N.RelativeElem("pitch is undefined", 0, 0, 0, rOpt)
    
    ElseIf (c = "") Then
        Set notehead = N.RelativeElem(Empty, 0, 0, pitch)
    Else
        shiftheadx = headx
'line        If (pitchelem.printer_shift) Then
'            adjust = IIf(pitchelem.printer_shift = "same", 1, 0)
'            shiftheadx = IIf(dir = "down", -glyphs.getSymbolWidth(c) * scale_ + adjust, glyphs.getSymbolWidth(c) * scale_ - adjust)
'        End If
        Set opts = New oRelativeOptions
        opts.scalex = scale_
        opts.scaley = scale_
        opts.thickness = glyphs.symbolHeightInPitches(c) * scale_
        opts.Name = pitchelem.Name
        Set notehead = N.RelativeElem(c, shiftheadx, glyphs.getSymbolWidth(c) * scale_, pitch, opts)
        notehead.stemdir = dir
        ''這邊是看是否要加入單個 1/4 1/8 1/16 的符桿符號
        If (flag <> "") Then
            pos = pitch + IIf(dir = "down", -7, 7) * scale_
            '' if this is a regular note, (not grace or tempo indicator) then the stem will have been stretched to the middle line if it is far from the center.
            If (shouldExtendStem) Then
                If (dir = "down" And pos > 6) Then pos = 6
                If (dir = "up" And pos < 6) Then pos = 6
            End If
            ''if (scale=1 && (dir="down")?(pos>6):(pos<6)) pos=6
            xdelta = IIf(dir = "down", headx, headx + notehead.w - 0.6)
            Set opts = New oRelativeOptions
            opts.scalex = scale_
            opts.scaley = scale_
            abselem.addRight N.RelativeElem(flag, xdelta, glyphs.getSymbolWidth(flag) * scale_, pos, opts)
        End If
        newDotShiftX = notehead.w + dotshiftx - 2 + 5 * dot
        For dot = dot To 0 Step -1
            dotadjusty = (1 - Abs(pitch) Mod 2)  ''PER: take abs value of the pitch. And the shift still happens on ledger lines.
            abselem.addRight N.RelativeElem("dots.dot", notehead.w + dotshiftx - 2 + 5 * dot, glyphs.getSymbolWidth("dots.dot"), pitch + dotadjusty)
        Next
    End If
    If Not (notehead Is Nothing) Then
        notehead.highestVert = pitchelem.highestVert
    End If
    ''加入升降記號
    If (printAccidentals And pitchelem.accidental <> "") Then
        Dim symb As String
        Select Case (pitchelem.accidental)
            Case "quartersharp":    symb = "accidentals.halfsharp"
            Case "dblsharp":        symb = "accidentals.dblsharp"
            Case "sharp":           symb = "accidentals.sharp"
            Case "quarterflat":     symb = "accidentals.halfflat"
            Case "flat":            symb = "accidentals.flat"
            Case "dblflat":         symb = "accidentals.dblflat"
            Case "natural":         symb = "accidentals.nat"
        End Select
        '' if a note is at least a sixth away, it can share a slot with another accidental
        Dim accSlotFound As Boolean
        Dim accPlace  As Double
        Dim j As Integer
        Dim h As Double
        Dim opt3 As oRelativeOptions
        
        accSlotFound = False
        accPlace = extrax
        For j = 0 To UBound(accidentalSlot) - 1
            If (pitch - accidentalSlot(j, 0) >= 6) Then
                accidentalSlot(j, 0) = pitch
                accPlace = accidentalSlot(j, 1)
                accSlotFound = True
                Exit For
            End If
        Next
        If (accSlotFound = False) Then
            accPlace = accPlace - (glyphs.getSymbolWidth(symb) * scale_ + 2)
            accidentalSlot.Push ([pitch, accPlace])
            accidentalshiftx = (glyphs.getSymbolWidth(symb) * scale_ + 2)
        End If
        Set opt3 = New oRelativeOptions
        h = glyphs.symbolHeightInPitches(symb)
         opt3.scalex = scale_
         opt3.scaley = scale_
         opt3.top = pitch + h / 2
         opt3.bottom = pitch - h / 2
        
        abselem.addExtra N.RelativeElem(symb, accPlace, glyphs.getSymbolWidth(symb), pitch + 3, opt3)
        extraLeft = glyphs.getSymbolWidth(symb) / 2  '' TODO-PER: We need a little extra width if there is an accidental, but I'm not sure why it isn't the full width of the accidental.
    End If
    Dim retDict As New Dictionary
    Set retDict("notehead") = notehead
    retDict("accidentalshiftx") = accidentalshiftx
    retDict("dotshiftx") = dotshiftx
    retDict("extraLeft") = extraLeft
    
    Set createNoteHead = retDict
End Function

'export var createNoteHeadJianpu = function (abselem: AbsoluteElement, c, pitchelem, options) {
'    if (!options) options = {}
'    var dir = (options.dir !== undefined) ? options.dir : null
'    '' dx 軸的位移
'    var headx = (options.headx !== undefined) ? options.headx : 0
'    var extrax = (options.extrax !== undefined) ? options.extrax : 0
'    var flag = (options.flag !== undefined) ? options.flag : null
'    var dot = (options.dot !== undefined) ? options.dot : 0
'    var dotshiftx = (options.dotshiftx !== undefined) ? options.dotshiftx : 0
'    var scale = (options.scale !== undefined) ? options.scale : 1
'    var accidentalSlot = (options.accidentalSlot !== undefined) ? options.accidentalSlot : []
'    var shouldExtendStem = (options.shouldExtendStem !== undefined) ? options.shouldExtendStem : false
'    var printAccidentals = (options.printAccidentals !== undefined) ? options.printAccidentals : true
'
'    '' TODO scale the dot as well
'    var pitch = pitchelem.verticalPos
'    var notehead: RelativeElement
'    var accidentalshiftx = 0
'    var newDotShiftX = 0
'    var extraLeft = 0
'    if (c === undefined) {
'        abselem.addFixed(new RelativeElement("pitch is undefined", 0, 0, 0, { type: "debug" }))
'    }
'    elseif (c === "") {
'        notehead = new RelativeElement(null, 0, 0, pitch)
'    } else {
'        var shiftheadx = headx
'        '' jianpu 不用這個
'        ''if (pitchelem.printer_shift) {
'        ''  var adjust = (pitchelem.printer_shift === "same") ? 1 : 0
'        ''  shiftheadx = (dir === "down") ? -glyphs.getSymbolWidth(c) * scale + adjust : glyphs.getSymbolWidth(c) * scale - adjust
'        ''}
'        shiftheadx = -glyphs.getSymbolWidth(c) * scale
'        var opts = { scalex: scale, scaley: scale, thickness: glyphs.symbolHeightInPitches(c) * scale, name: pitchelem.name }
'
'        notehead = new RelativeElement(c, 0, glyphs.getSymbolWidth(c) * scale, pitch, opts)
'        notehead.stemDir = dir
'        ''這邊是看是否要加入單個 1/4 1/8 1/16 的符桿符號
'        '' jianpu 不用
'        ''if (flag) {
'        ''  var pos = pitch + ((dir === "down") ? -7 : 7) * scale
'        ''  '' if this is a regular note, (not grace or tempo indicator) then the stem will have been stretched to the middle line if it is far from the center.
'        ''  if (shouldExtendStem) {
'        ''      if (dir === "down" && pos > 6)
'        ''          pos = 6
'        ''      if (dir === "up" && pos < 6)
'        ''          pos = 6
'        ''  }
'        ''  ''if (scale===1 && (dir==="down")?(pos>6):(pos<6)) pos=6
'        ''  var xdelta = (dir === "down") ? headx : headx + notehead.w - 0.6
'        ''  abselem.addRight(new RelativeElement(flag, xdelta, glyphs.getSymbolWidth(flag) * scale, pos, { scalex: scale, scaley: scale }))
'        ''}
'        var getDurlog = function (duration) {
'            '' TODO-PER: This is a hack to prevent a Chrome lockup. Duration should have been defined already,
'            '' but there's definitely a case where it isn't. [Probably something to do with triplets.]
'            if (duration === undefined) {
'                return 0
'            }
'            ''        console.log("getDurlog: " + duration)
'            return Math.floor(Math.log(duration) / Math.log(2))
'        }
'        '' 加入 jianpu 的節拍線
'        if (flag) {
'            var duration = abselem.duration  '' get the duration via abcelem because of triplets
'            if (duration === 0) duration = 0.25  '' if this is stemless, then we use quarter note as the duration.
'            for (var durlog = getDurlog(duration)  durlog < -2  durlog++) {
'                var index = -durlog - 3
'                abselem.addFixed(new RelativeElement(null, 0, glyphs.getSymbolWidth(c) * scale, 1.5 - index * 0.8, { type: "beatline", pitch2: 1, linewidth: 1.2, bottom: 1 }))
'            }
'        }
'        ''abselem.addRight(new RelativeElement(null, dx, 0, p1, { type: "stem", pitch2: p2, linewidth: width, bottom: p1 - 1 }))
'        newDotShiftX = notehead.w + dotshiftx - 2 + 5 * dot
'        for (  dot > 0  dot--) {
'            ''PER: take abs value of the pitch. And the shift still happens on ledger lines.
'            ''PER：取音高的絕對值。 這種轉變仍然發生在帳本上。
'            var dotadjusty = (1 - Math.abs(pitch) % 2)
'            abselem.addRight(new RelativeElement("dots.dot", notehead.w + dotshiftx - 2 + 5 * dot, glyphs.getSymbolWidth("dots.dot"), pitch + dotadjusty))
'        }
'    }
'    if (notehead)
'        notehead.highestVert = pitchelem.highestVert
'    ''加入升降記號
'    if (printAccidentals && pitchelem.accidental) {
'        var symb
'        switch (pitchelem.accidental) {
'            Case "quartersharp":
'                symb = "accidentals.halfsharp"
'                break
'            Case "dblsharp":
'                symb = "accidentals.dblsharp"
'                break
'            Case "sharp":
'                symb = "accidentals.sharp"
'                break
'            Case "quarterflat":
'                symb = "accidentals.halfflat"
'                break
'            Case "flat":
'                symb = "accidentals.flat"
'                break
'            Case "dblflat":
'                symb = "accidentals.dblflat"
'                break
'            Case "natural":
'                symb = "accidentals.nat"
'        }
'        '' if a note is at least a sixth away, it can share a slot with another accidental
'        var accSlotFound = false
'        '' jianpu 不用 var accPlace = extrax
'        var accPlace = 0
'        for (var j = 0  j < accidentalSlot.length  j++) {
'            if (pitch - accidentalSlot[j][0] >= 6) {
'                accidentalSlot[j][0] = pitch
'                accPlace = accidentalSlot[j][1]
'                accSlotFound = true
'                break
'            }
'        }
'        if (accSlotFound === false) {
'            accPlace -= (glyphs.getSymbolWidth(symb) * scale + 2)
'            accidentalSlot.push([pitch, accPlace])
'            accidentalshiftx = (glyphs.getSymbolWidth(symb) * scale + 2)
'        }
'        var h = glyphs.symbolHeightInPitches(symb)
'        abselem.addExtra(new RelativeElement(symb, accPlace, glyphs.getSymbolWidth(symb), pitch + 3, { scalex: scale, scaley: scale, top: pitch + h / 2, bottom: pitch - h / 2 }))
'        extraLeft = glyphs.getSymbolWidth(symb) / 2  '' TODO-PER: We need a little extra width if there is an accidental, but I'm not sure why it isn't the full width of the accidental.
'    }
'
'    return { notehead: notehead, accidentalshiftx: accidentalshiftx, dotshiftx: newDotShiftX, extraLeft: extraLeft }
'
'}
'export default createNoteHead


Public Function pitchesToPerc(pitchObj) As String
    Dim pitchDict As New Dictionary
    pitchDict("f0") = "_C":  pitchDict("s8") = "^d":
    pitchDict("n0") = "=C":  pitchDict("x8") = "d":
    pitchDict("s0") = "^C":  pitchDict("f9") = "_e":
    pitchDict("x0") = "C":   pitchDict("n9") = "=e":
    pitchDict("f1") = "_D":  pitchDict("s9") = "^e":
    pitchDict("n1") = "=D":  pitchDict("x9") = "e":
    pitchDict("s1") = "^D":  pitchDict("f10") = "_f":
    pitchDict("x1") = "D":   pitchDict("n10") = "=f":
    pitchDict("f2") = "_E":  pitchDict("s10") = "^f":
    pitchDict("n2") = "=E":  pitchDict("x10") = "f":
    pitchDict("s2") = "^E":  pitchDict("f11") = "_g":
    pitchDict("x2") = "E":   pitchDict("n11") = "=g":
    pitchDict("f3") = "_F":  pitchDict("s11") = "^g":
    pitchDict("n3") = "=F":  pitchDict("x11") = "g":
    pitchDict("s3") = "^F":  pitchDict("f12") = "_a":
    pitchDict("x3") = "F":   pitchDict("n12") = "=a":
    pitchDict("f4") = "_G":  pitchDict("s12") = "^a":
    pitchDict("n4") = "=G":  pitchDict("x12") = "a":
    pitchDict("s4") = "^G":  pitchDict("f13") = "_b":
    pitchDict("x4") = "G":   pitchDict("n13") = "=b":
    pitchDict("f5") = "_A":  pitchDict("s13") = "^b":
    pitchDict("n5") = "=A":  pitchDict("x13") = "b":
    pitchDict("s5") = "^A":  pitchDict("f14") = "_c'":
    pitchDict("x5") = "A":   pitchDict("n14") = "=c'":
    pitchDict("f6") = "_B":  pitchDict("s14") = "^c'":
    pitchDict("n6") = "=B":  pitchDict("x14") = "c'":
    pitchDict("s6") = "^B":  pitchDict("f15") = "_d'":
    pitchDict("x6") = "B":   pitchDict("n15") = "=d'":
    pitchDict("f7") = "_c":  pitchDict("s15") = "^d'":
    pitchDict("n7") = "=c":  pitchDict("x15") = "d'":
    pitchDict("s7") = "^c":  pitchDict("f16") = "_e'":
    pitchDict("x7") = "c":   pitchDict("n16") = "=e'":
    pitchDict("f8") = "_d":  pitchDict("s16") = "^e'":
    pitchDict("n8") = "=d":  pitchDict("x16") = "e'":
    
    Dim pitch As String

    pitch = IIf(Not pitchObj.accidental Is Nothing, pitchObj.accidental(0), "x") & pitchObj.verticalPos
    pitchesToPerc = pitchDict(pitch)
End Function

