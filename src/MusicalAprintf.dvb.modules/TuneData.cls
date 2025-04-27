VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TuneData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public formatting As vFormatting
Public engraver As EngraverController
Public lines As TuneLineList
Public media As String 'Dictionary media
Public metaText As vMetaText
Public metaTextInfo As vMetaTextInfo
Public version As String
'    lineBreaksas  Array<number>
'    visualTransposeas Number
    
Public staffNum As Double
Public voiceNum As Double
Public lineNum As Double
Public runningFonts As vFormatting   'vFont
Public lineBreaks As iArray     ' Array<double>
Public visualTranspose As Double
Public TopText As Dictionary 'TopText
Public BottomText As Dictionary 'BottomText
Public totalTime
Public totalBeats
Public meter As vMeterFraction
Public tempoLocations
Public noteTimings As iArray
    
    
Private Sub Class_Initialize()
    Me.reset
End Sub

    
Public Sub reset()


        Me.version = "1.1.0"
        Me.media = "screen"
        Set Me.metaText = New vMetaText
        Set Me.metaTextInfo = New vMetaTextInfo
        Set Me.formatting = New vFormatting
        Set Me.lines = New TuneLineList
        Me.staffNum = 0
        Me.voiceNum = 0
        Me.lineNum = 0
        Set Me.runningFonts = New vFormatting
        'delete Me.visualTranspose
End Sub

'    copy(dest, src, prop, attrs) {
'        for (var i = 0  i < attrs.length  i++)
'            dest[prop][attrs[i]] = src[prop][attrs[i]]
'    }
'
'    copyTopInfo(src) {
'        var attrs = ['tempo', 'title', 'header', 'rhythm', 'origin', 'composer', 'author', 'partOrder']
'        me.copy(this, src, "metaText", attrs)
'        me.copy(this, src, "metaTextInfo", attrs)
'    }
'
'    copyBottomInfo(src) {
'        var attrs = ['unalignedWords',
'            'book',
'            'source',
'            'discography',
'            'notes',
'            'transcription',
'            'history',
'            'abc-copyright',
'            'abc-creator',
'            'abc-edited-by',
'            'footer']
'        me.copy(this, src, "metaText", attrs)
'        me.copy(this, src, "metaTextInfo", attrs)
'    }

    '' The structure consists of a hash with the following two items:
    '' metaText: a hash of {key, value}, where key is one of: title, author, rhythm, source, transcription, unalignedWords, etc...
    '' tempo: { noteLength: number (e.g. .125), bpm: number }
    '' lines: an array of elements, or one of the following:
    ''
    '' STAFF: array of elements
    '' SUBTITLE: string
    ''
    '' TODO: actually, the start and end char should modify each part of the note type
    '' The elements all have a type field and a start and end char
    '' field. The rest of the fields depend on the type and are listed below:
    '' REST: duration=1,2,4,8  chord: string
    '' NOTE: accidental=none,dbl_flat,flat,natural,sharp,dbl_sharp
    ''      pitch: "C" is 0. The numbers refer to the pitch letter.
    ''      duration: .5 (sixteenth), .75 (dotted sixteenth), 1 (eighth), 1.5 (dotted eighth)
    ''          2 (quarter), 3 (dotted quarter), 4 (half), 6 (dotted half) 8 (whole)
    ''      chord: { name:chord, position: one of 'default', 'above', 'below' }
    ''      end_beam = true or undefined if this is the last note in a beam.
    ''      lyric: array of { syllable: xxx, divider: one of " -_" }
    ''      startTie = true|undefined
    ''      endTie = true|undefined
    ''      startTriplet = num <- that is the number to print
    ''      endTriplet = true|undefined (the last note of the triplet)
    '' TODO: actually, decoration should be an array.
    ''      decoration: upbow, downbow, accent
    '' BAR: type=bar_thin, bar_thin_thick, bar_thin_thin, bar_thick_thin, bar_right_repeat, bar_left_repeat, bar_double_repeat
    ''  number: 1 or 2: if it is the start of a first or second ending
    '' CLEF: type=treble,bass
    '' KEY-SIG:
    ''      accidentals[]: { acc:sharp|dblsharp|natural|flat|dblflat,  note:a|b|c|d|e|f|g }
    '' METER: type: common_time,cut_time,specified
    ''      if specified, { num: 99, den: 99 }

Public Function getBeatLength() As Double
        '' This returns a fraction: for instance 1/4 for a quarter
        '' There are two types of meters: compound and regular. Compound meter has 3 beats counted as one.
        Dim meter As vMeterFraction
        Dim multiplier As Integer
        Set meter = Me.getMeterFraction()
        multiplier = 1
        If (meter.num = 6 Or meter.num = 9 Or meter.num = 12) Then
            multiplier = 3
        ElseIf (meter.num = 3 And meter.den = 8) Then
            multiplier = 3
        End If

        getBeatLength = multiplier / meter.den
End Function

'    computePickupLength(lines, barLength): number {
'        var pickupLength = 0
'        for (var i = 0  i < lines.length  i++) {
'            if (lines[i].staff) {
'                for (var j = 0  j < lines[i].staff.length  j++) {
'                    for (var v = 0  v < lines[i].staff[j].voices.length  v++) {
'                        var voice = lines[i].staff[j].voices[v]
'                        var tripletMultiplier = 1
'                        for (var el = 0  el < voice.length  el++) {
'                            var isSpacer = voice[el].rest && voice[el].rest.type === "spacer"
'                            if (voice[el].startTriplet)
'                                tripletMultiplier = voice[el].tripletMultiplier
'                            if (voice[el].duration && !isSpacer && voice[el].el_type !== "tempo")
'                                pickupLength += voice[el].duration * tripletMultiplier
'                            if (voice[el].endTriplet)
'                                tripletMultiplier = 1
'                            if (pickupLength >= barLength)
'                                pickupLength -= barLength
'                            if (voice[el].el_type === 'bar')
'                                return pickupLength
'                        }
'                    }
'                }
'            }
'        }
'
'        return pickupLength
'    }
'    getPickupLength(): number {
'        var barLength = Me.getBarLength()
'        var pickupLength = Me.computePickupLength(Me.lines, barLength)
'
'        '' If computed pickup length is very close to 0 or the bar length, we assume
'        '' that we actually have a full bar and hence no pickup.
'        return (pickupLength < 1e-8 || barLength - pickupLength < 1e-8) ? 0 : pickupLength
'    }

Public Function getBarLength() As Double
        Dim meter As vMeterFraction
        Set meter = Me.getMeterFraction()
        getBarLength = meter.num / meter.den
End Function

Public Function getTotalTime() As Double
       getTotalTime = Me.totalTime
End Function

Public Function getTotalBeats() As Double
        getTotalBeats = Me.totalBeats
End Function

Public Function millisecondsPerMeasure(bpmOverride As Integer) As Double
        Dim bpm As Integer
        Dim tempo As vtempoProperties
        If (bpmOverride <> 0) Then
            bpm = bpmOverride
        Else
            Set tempo = IIf(Me.metaText Is Nothing, Nothing, Me.metaText.tempo)
            bpm = Me.getBpm(tempo)
        End If
        If (bpm <= 0) Then
            bpm = 1  '' I don't think this can happen, but we don't want a possibility of dividing by zero.
        End If

        Dim beatsPerMeasure
        Dim minutesPerMeasure
         beatsPerMeasure = Me.getBeatsPerMeasure()

         minutesPerMeasure = beatsPerMeasure / bpm
        millisecondsPerMeasure = minutesPerMeasure * 60000
End Function

Public Function getBeatsPerMeasure() As Double
        Dim beatLen, barLen
        beatLen = Me.getBeatLength()
        barLen = Me.getBarLength()
        getBeatsPerMeasure = barLen / beatLen
End Function

Public Function getMeter() As vMeter
        Dim i As Integer, j As Integer
        Dim line As TuneLine
        Dim meter As vMeter
        For i = 0 To Me.lines.Count - 1
            Set line = Me.lines(i)
            If Not (line.Staffs Is Nothing) Then
                For j = 0 To line.Staffs.Count - 1
                    Set meter = line.Staffs(j).meter
                    If Not (meter Is Nothing) Then
                        Set getMeter = meter
                        Exit Function
                    End If
                Next
            End If
        Next
        Set meter = New vMeter
        meter.typs = "common_time"
        Set getMeter = meter
End Function

Public Function getMeterFraction() As vMeterFraction
        Dim meter As vMeter
        Dim num As Integer, den As Integer
        Set meter = Me.getMeter()
        num = 4
        den = 4
        If Not (meter Is Nothing) Then
            If (meter.typs = "specified") Then
                num = CInt(meter.value(0).num)
                den = CInt(meter.value(0).den)
            ElseIf (meter.typs = "cut_time") Then
                num = 2
                den = 2
            ElseIf (meter.typs = "common_time") Then
                num = 4
                den = 4
            End If
        End If
        Set Me.meter = New vMeterFraction
        Me.meter.num = num
        Me.meter.den = den
        Set getMeterFraction = Me.meter  '' TODO-PER: is this saved value used anywhere? A get function shouldn't change state.
End Function

Public Function getKeySignature() As vKeySignature


        Dim i As Integer, j As Integer
        Dim line As TuneLine
        For i = 0 To Me.lines.Count - 1
            Set line = Me.lines(i)
            If Not (line.Staff Is Nothing) Then
                For j = 0 To line.Staff.Count - 1
                    If Not (line.Staff(j).key Is Nothing) Then
                        Set getKeySignature = line.Staff(j).key
                    End If
                Next
            End If
        Next
        Set getKeySignature = Nothing
End Function

'    getElementFromChar(charPos: number): VoiceItem {
'        for (var i = 0  i < me.lines.length  i++) {
'            var line = me.lines[i]
'            if (line.staff) {
'                for (var j = 0  j < line.staff.length  j++) {
'                    var staff = line.staff[j]
'                    for (var k = 0  k < staff.voices.length  k++) {
'                        var voice = staff.voices[k]
'                        for (var ii = 0  ii < voice.length  ii++) {
'                            var elem = voice[ii]
'                            if (elem.startChar && elem.endChar &&
'                                elem.startChar <= charPos && elem.endChar > charPos)
'                                return elem
'                        }
'                    }
'                }
'            }
'        }
'        return null
'    }
'
'    addVerticalInfo(timingEvents) {
'        '' Add vertical info to the bar events: put the next event's top, and the event after the next measure's top.
'        var lastBarTop
'        var lastBarBottom
'        var lastEventTop
'        var lastEventBottom
'        for (var e = timingEvents.length - 1  e >= 0  e--) {
'            var ev = timingEvents[e]
'            if (ev.type === 'bar') {
'                ev.top = lastEventTop
'                ev.nextTop = lastBarTop
'                lastBarTop = lastEventTop
'
'                ev.bottom = lastEventBottom
'                ev.nextBottom = lastBarBottom
'                lastBarBottom = lastEventBottom
'            } else if (ev.type === 'event') {
'                lastEventTop = ev.top
'                lastEventBottom = ev.top + ev.height
'            }
'        }
'    }
'
'    makeSortedArray(hash) {
'        var arr = []
'        for (var k in hash) {
'            if (hash.hasOwnProperty(k))
'                arr.push(hash[k])
'        }
'        arr = arr.sort(function (a, b) {
'            var Diff = a.milliseconds - b.milliseconds
'            '' if the events have the same time, make sure a bar comes before a note
'            if (diff !== 0) {
'                return diff
'            }
'            else {
'                return a.type === "bar" ? -1 : 1
'            }
'        })
'        return arr
'    }
'
'    addElementToEvents(eventHash, element, voiceTimeMilliseconds, top, height, line, measureNumber, timeDivider, isTiedState, nextIsBar) {
'        if (element.hint)
'            return { isTiedState: undefined, duration: 0 }
'        var realDuration = element.durationClass ? element.durationClass : element.duration
'        if (element.abcelem.rest && element.abcelem.rest.type === "spacer")
'            realDuration = 0
'        if (realDuration > 0) {
'            var es = []
'            '' If there is an invisible rest, then there are not elements, so don't push a null one.
'            for (var i = 0  i < element.elemset.length  i++) {
'                if (element.elemset[i] !== null)
'                    es.push(element.elemset[i])
'            }
'            var isTiedToNext = Element.startTie
'            if (isTiedState !== undefined) {
'                eventHash ["event" + isTiedState].elements.Push(es) '' Add the tied note to the first note that it is tied to
'                if (nextIsBar) {
'                    if (!eventHash["event" + voiceTimeMilliseconds]) {
'                        eventHash["event" + voiceTimeMilliseconds] = {
'                            type: "event",
'                            milliseconds: voiceTimeMilliseconds,
'                            line: line,
'                            measureNumber: measureNumber,
'                            top: top,
'                            height: height,
'                            left: null,
'                            width: 0,
'                            elements: [],
'                            startChar: null,
'                            endChar: null,
'                            startCharArray: [],
'                            endCharArray:[],
'                        }
'                    }
'                    eventHash ["event" + voiceTimeMilliseconds].measureStart = True
'                    nextIsBar = False
'                }
'                if (!isTiedToNext)
'                    isTiedState = undefined
'            } else {
'                '' the last note wasn't tied.
'                if (!eventHash["event" + voiceTimeMilliseconds]) {
'                    eventHash["event" + voiceTimeMilliseconds] = {
'                        type: "event",
'                        milliseconds: voiceTimeMilliseconds,
'                        line: line,
'                        measureNumber: measureNumber,
'                        top: top,
'                        height: height,
'                        left: element.x,
'                        width: element.w,
'                        elements: [es],
'                        startChar: element.abcelem.startChar,
'                        endChar: element.abcelem.endChar,
'                        startCharArray: [element.abcelem.startChar],
'                        endCharArray: [element.abcelem.endChar],
'                        midiPitches: element.abcelem.midiPitches ? parseCommon.cloneArray(element.abcelem.midiPitches) : []
'                    }
'                    if (element.abcelem.midiGraceNotePitches)
'                        eventHash ["event" + voiceTimeMilliseconds].midiGraceNotePitches = parseCommon.cloneArray(Element.abcelem.midiGraceNotePitches)
'                } else {
'                    '' If there is more than one voice then two notes can fall at the same time. Usually they would be lined up in the same place, but if it is a whole rest, then it is placed funny. In any case, the left most element wins.
'                    if (eventHash["event" + voiceTimeMilliseconds].left)
'                        eventHash["event" + voiceTimeMilliseconds].left = Math.min(eventHash["event" + voiceTimeMilliseconds].left, element.x)
'                    Else
'                        eventHash ["event" + voiceTimeMilliseconds].left = Element.x
'                    eventHash ["event" + voiceTimeMilliseconds].elements.Push(es)
'                    eventHash ["event" + voiceTimeMilliseconds].startCharArray.Push(Element.abcelem.startChar)
'                    eventHash ["event" + voiceTimeMilliseconds].endCharArray.Push(Element.abcelem.endChar)
'                    if (eventHash["event" + voiceTimeMilliseconds].startChar === null)
'                        eventHash ["event" + voiceTimeMilliseconds].startChar = Element.abcelem.startChar
'                    if (eventHash["event" + voiceTimeMilliseconds].endChar === null)
'                        eventHash ["event" + voiceTimeMilliseconds].endChar = Element.abcelem.endChar
'                    if (element.abcelem.midiPitches && element.abcelem.midiPitches.length) {
'                        if (!eventHash["event" + voiceTimeMilliseconds].midiPitches)
'                            eventHash ["event" + voiceTimeMilliseconds].midiPitches = []
'                        for (var i = 0  i < element.abcelem.midiPitches.length  i++)
'                            eventHash["event" + voiceTimeMilliseconds].midiPitches.push(element.abcelem.midiPitches[i])
'                    }
'                    if (element.abcelem.midiGraceNotePitches && element.abcelem.midiGraceNotePitches.length) {
'                        if (!eventHash["event" + voiceTimeMilliseconds].midiGraceNotePitches)
'                            eventHash ["event" + voiceTimeMilliseconds].midiGraceNotePitches = []
'                        for (var j = 0  j < element.abcelem.midiGraceNotePitches.length  j++)
'                            eventHash["event" + voiceTimeMilliseconds].midiGraceNotePitches.push(element.abcelem.midiGraceNotePitches[j])
'                    }
'                }
'                if (nextIsBar) {
'                    eventHash ["event" + voiceTimeMilliseconds].measureStart = True
'                    nextIsBar = False
'                }
'                if (isTiedToNext)
'                    isTiedState = voiceTimeMilliseconds
'            }
'        }
'        return { isTiedState: isTiedState, duration: realDuration / timeDivider, nextIsBar: nextIsBar || element.type === 'bar' }
'    }
'
'    makeVoicesArray(): Array<Selectable[]> {
'        '' First make a new array that is arranged by voice so that the repeats that span different lines are handled correctly.
'        var voicesArr = []
'        var measureNumber = []
'        var tempos = {}
'        for (var line = 0  line < me.engraver.staffgroups.length  line++) {
'            var group = me.engraver.staffgroups[line]
'            if (group && group.staffs && group.staffs.length > 0) {
'                var firstStaff = group.staffs[0]
'                var middleC = firstStaff.absoluteY
'                var top = middleC - firstStaff.top * spacing.STEP
'                var lastStaff = group.staffs[group.staffs.length - 1]
'                middleC = lastStaff.absoluteY
'                var bottom = middleC - lastStaff.bottom * spacing.STEP
'                var height = bottom - top
'
'                var voices = group.voices
'                for (var v = 0  v < voices.length  v++) {
'                    if (voices[v].staff && voices[v].staff.isTabStaff)
'                        continue
'                    var noteFound = False
'                    if (!voicesArr[v])
'                        voicesArr [v] = []
'                    if (measureNumber[v] === undefined)
'                        measureNumber [v] = 0
'                    var elements = voices[v].children
'                    for (var elem = 0  elem < elements.length  elem++) {
'                        if (elements[elem].type === "tempo")
'                            tempos[measureNumber[v]] = me.getBpm(elements[elem].abcelem)
'                        voicesArr[v].push({ top: top, height: height, line: group.line, measureNumber: measureNumber[v], elem: elements[elem] })
'                        if (elements[elem].type === 'bar' && noteFound) '' Count the measures by counting the bar lines, but skip a bar line that appears at the left of the music, before any notes.
'                            measureNumber[v]++
'                        if (elements[elem].type === 'note' || elements[elem].type === 'rest')
'                            noteFound = True
'                    }
'                }
'            }
'        }
'        Me.tempoLocations = tempos  '' This should be passed back, but the function is accessible publicly so that would break the interface.
'        return voicesArr
'    }
'
'    setupEvents(startingDelay, timeDivider, startingBpm, warp) {
'        if (!warp) warp = 1
'        var timingEvents = []
'
'        var eventHash = {}
'        '' The time is the number of seconds from the beginning of the piece.
'        '' The units we are scanning are in notation units (i.e. 0.25 is a quarter note)
'        var time = startingDelay
'        var isTiedState
'        var nextIsBar = True
'        var voices = Me.makeVoicesArray()
'        var maxVoiceTimeMilliseconds = 0
'        for (var v = 0  v < voices.length  v++) {
'            var voiceTime = time
'            var voiceTimeMilliseconds = Math.Round(voiceTime * 1000)
'            var startingRepeatElem = 0
'            var endingRepeatElem = -1
'            var elements = voices[v]
'            var bpm = startingBpm
'            timeDivider = Me.getBeatLength() * bpm / 60
'            var tempoDone = -1
'            for (var elem = 0  elem < elements.length  elem++) {
'                var thisMeasure = elements[elem].measureNumber
'                if (tempoDone !== thisMeasure && me.tempoLocations[thisMeasure]) {
'                    bpm = me.tempoLocations[thisMeasure]
'                    timeDivider = warp * Me.getBeatLength() * bpm / 60
'                    tempoDone = thisMeasure
'                }
'                var element = elements[elem].elem
'                var ret = me.addElementToEvents(eventHash, element, voiceTimeMilliseconds, elements[elem].top, elements[elem].height, elements[elem].line, elements[elem].measureNumber, timeDivider, isTiedState, nextIsBar)
'                isTiedState = ret.isTiedState
'                nextIsBar = ret.nextIsBar
'                voiceTime += ret.duration
'                var lastHash
'                if (element.duration > 0 && eventHash["event" + voiceTimeMilliseconds]) '' This won't exist if this is the end of a tie.
'                    lastHash = "event" + voiceTimeMilliseconds
'                voiceTimeMilliseconds = Math.Round(voiceTime * 1000)
'                if (element.type === 'bar') {
'                    var barType = Element.abcelem.Type
'                    var endRepeat = (barType === "bar_right_repeat" || barType === "bar_dbl_repeat")
'                    var startEnding = (element.abcelem.startEnding === '1')
'                    var startRepeat = (barType === "bar_left_repeat" || barType === "bar_dbl_repeat" || barType === "bar_right_repeat")
'                    if (endRepeat) {
'                        '' Force the end of the previous note to the position of the measure - the cursor won't go past the end repeat
'                        if (elem > 0) {
'                            eventHash [lastHash].endX = Element.x
'                        }
'
'                        if (endingRepeatElem === -1)
'                            endingRepeatElem = elem
'                        var lastVoiceTimeMilliseconds = 0
'                        tempoDone = -1
'                        for (var el2 = startingRepeatElem  el2 < endingRepeatElem  el2++) {
'                            thisMeasure = elements[el2].measureNumber
'                            if (tempoDone !== thisMeasure && me.tempoLocations[thisMeasure]) {
'                                bpm = me.tempoLocations[thisMeasure]
'                                timeDivider = warp * Me.getBeatLength() * bpm / 60
'                                tempoDone = thisMeasure
'                            }
'                            var element2 = elements[el2].elem
'                            ret = me.addElementToEvents(eventHash, element2, voiceTimeMilliseconds, elements[el2].top, elements[el2].height, elements[el2].line, elements[el2].measureNumber, timeDivider, isTiedState, nextIsBar)
'                            isTiedState = ret.isTiedState
'                            nextIsBar = ret.nextIsBar
'                            voiceTime += ret.duration
'                            lastVoiceTimeMilliseconds = voiceTimeMilliseconds
'                            voiceTimeMilliseconds = Math.Round(voiceTime * 1000)
'                        }
'                        if (eventHash["event" + lastVoiceTimeMilliseconds]) '' This won't exist if it is the beginning of the next line. That's ok because we will just count the end of the last line as the end.
'                            eventHash["event" + lastVoiceTimeMilliseconds].endX = elements[endingRepeatElem].elem.x
'                        nextIsBar = True
'                        endingRepeatElem = -1
'                    }
'                    if (startEnding)
'                        endingRepeatElem = elem
'                    if (startRepeat)
'                        startingRepeatElem = elem
'                }
'            }
'            maxVoiceTimeMilliseconds = Math.max(maxVoiceTimeMilliseconds, voiceTimeMilliseconds)
'        }
'        '' now we have all the events, but if there are multiple voices then there may be events out of order or duplicated, so normalize it.
'        timingEvents = Me.makeSortedArray(eventHash)
'        Me.addVerticalInfo (timingEvents)
'        me.addEndPoints(me.lines, timingEvents)
'        timingEvents.push({ type: "end", milliseconds: maxVoiceTimeMilliseconds })
'        me.addUsefulCallbackInfo(timingEvents, bpm * warp)
'        return timingEvents
'    }
'
'    addUsefulCallbackInfo(timingEvents, bpm) {
'        var millisecondsPerMeasure = Me.millisecondsPerMeasure(bpm)
'        for (var i = 0  i < timingEvents.length  i++) {
'            var ev = timingEvents[i]
'            ev.millisecondsPerMeasure = millisecondsPerMeasure
'        }
'    }
'
'    skipTies(elements, index) {
'        while (index < elements.length && elements[index].left === null)
'            index++
'        return elements[index]
'    }
'    addEndPoints(lines, elements) {
'        if (elements.length < 1)
'            Return
'        for (var i = 0  i < elements.length - 1  i++) {
'            var el = elements[i]
'            var next = me.skipTies(elements, i + 1)
'            if (el.left !== null) {
'                '' If there is no left element that is because this is a tie so it should be skipped.
'                var endX = (next && el.top === next.top) ? next.left : lines[el.line].staffGroup.w
'                '' If this is already set, it is because the notes aren't sequential here, like the next thing is a repeat bar line.
'                '' In that case, the right-most position is passed in. There could still be an intervening note in another voice, so always look for the closest position.
'                '' If there is a repeat that stays on the same line, the endX set above won't be right because the next note will be before. In that case, use the endX that was calculated.
'                if (el.endX !== undefined) {
'                    if (endX > el.left)
'                        el.endX = Math.min(el.endX, endX)
'                } else
'                    el.endX = endX
'            }
'        }
'        var lastEl = elements[elements.length - 1]
'        lastEl.endX = lines[lastEl.line].staffGroup.w
'    }
'
'    getBpm(tempo): number {
'        var bpm
'        if (tempo) {
'            bpm = tempo.bpm
'            var beatLength = Me.getBeatLength()
'            var statedBeatLength = tempo.duration && tempo.duration.length > 0 ? tempo.duration[0] : beatLength
'            bpm = bpm * statedBeatLength / beatLength
'        }
'        if (!bpm) {
'            bpm = 180
'            '' Compensate for compound meter, where the beat isn't a beat.
'            var meter = Me.getMeterFraction()
'            if (meter && meter.num !== 3 && (meter.num % 3 === 0)) {
'                bpm = 120
'            }
'        }
'        return bpm
'    }
'
'    setTiming(bpm?: number, measuresOfDelay?: number) {
'        measuresOfDelay = measuresOfDelay || 0
'        if (!me.engraver || !me.engraver.staffgroups) {
'            console.Log ("setTiming cannot be called before the tune is drawn.")
'            Me.noteTimings = []
'            return me.noteTimings
'        }
'
'        var tempo = me.metaText ? me.metaText.tempo : null
'        var naturalBpm = Me.getBpm(tempo)
'        var warp = 1
'        if (bpm) {
'            if (tempo)
'                warp = bpm / naturalBpm
'        } else
'            bpm = naturalBpm
'
'        '' Calculate the basic midi data. We only care about the qpm variable here.
'        ''me.setUpAudio({qpm: bpm})
'
'        var beatLength = Me.getBeatLength()
'        var beatsPerSecond = bpm / 60
'
'        var measureLength = Me.getBarLength()
'
'        var startingDelay = measureLength / beatLength * measuresOfDelay / beatsPerSecond
'        if (startingDelay)
'            startingDelay -= me.getPickupLength() / beatLength / beatsPerSecond
'        var timeDivider = beatLength * beatsPerSecond
'
'        Me.noteTimings = Me.setupEvents(startingDelay, timeDivider, bpm, warp)
'        if (me.noteTimings.length > 0) {
'            me.totalTime = me.noteTimings[me.noteTimings.length - 1].milliseconds / 1000
'            Me.totalBeats = Me.totalTime * beatsPerSecond
'        } else {
'            Me.totalTime = undefined
'            Me.totalBeats = undefined
'        }
'        return me.noteTimings
'    }
'
'    setUpAudio(options: SynthOptions): AudioTracks {
'        if (!options) options = {}
'        var seq = sequence(This, options)
'        return flatten(seq, options, me.formatting.percmap, me.formatting.midi)
'    }
'    deline(options) {
'        return delineTune(me.lines, options)
'    }
'
