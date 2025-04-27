VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EngraverController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public classes As Dictionary
Public dragColor As String
Public dragIndex As Double
Public dragMouseStart As point 'as  { x as  number, y as  number  }
Public dragTarget 'as  null | any
Public dragYStep As Double
Public dragging As Boolean
Public engraver As AbstractEngraver
Public GetFontAttr As GetFontAndAttr
Public getTxtSize As getTextSize
Public listeners As iArray     'ClickListener []
Public rangeHighlight 'as  any
Public renderer As RendererModule
Public responsive As String
Public scale_ As Double
Public initialClef 'as  any
Public selectTypes As iArray  'boolean | Array<DragTypes>
Public selectables As iArray '  Array<Selectable>
Public selected As iArray  'Array<any>
Public selectionColor As String
Public space As Double
Public staffgroups As iArray 'StaffGroupElement
Public staffwidthPrint As Double
Public staffwidthScreen As Double
Public width As Double
Public oneSvgPerLine
Public jazzchords
Public germanAlphabet
Public lineThickness
Public svgs As iArray 'SVGmodel

    
Public Sub init(paper, params As Dictionary)

        'params = params || {}
        Set params = New Dictionary
        Me.oneSvgPerLine = params("oneSvgPerLine")
        Me.selectionColor = params("selectionColor")
        Me.dragColor = IIf(params("dragColor") <> Empty, params("dragColor"), params("selectionColor"))
        Me.dragging = IIf(params("dragging") <> Empty, True, False)
       'line Set Me.selectTypes = params("selectTypes")
        Me.responsive = params("responsive")
        Me.space = 3 * spacing.space
        Me.initialClef = params("initialClef")
        Me.scale_ = IIf(params("scale") <> 0, CDbl(params("scale")), 1)
        Set Me.classes = New Dictionary 'Classes({ shouldAddClasses: params("add_classes") })
        If (Not (Me.scale_ > 0.1)) Then
            Me.scale_ = 0
        End If

        If (params("staffwidth") <> Empty) Then
            '' Note: Normally all measurements to the engraver are in POINTS. However, if a person is formatting for the
            '' screen and directly inputting the width, then it is more logical to have the measurement in pixels.
            Me.staffwidthScreen = params("staffwidth")
            Me.staffwidthPrint = params("staffwidth")
         Else
            Me.staffwidthScreen = 740  '' TODO-PER: Not sure where this number comes from, but this is how it's always been.
            Me.staffwidthPrint = 680  '' The number of pixels in 8.5", after 1cm of margin has been removed.
        End If
       Set Me.listeners = New iArray
        If (params("clickListener") <> Empty) Then
            Me.addSelectListener (params("clickListener"))
        End If

       Set Me.renderer = New RendererModule   ' Renderer(paper)
       Me.renderer.init
       Me.renderer.setPaddingOverride params
        If (params("showDebug") <> Empty) Then
            Me.renderer.showDebug = params("showDebug")
        End If
        If (params("jazzchords") <> Empty) Then
            Me.jazzchords = params("jazzchords")
        End If
        If (params("germanAlphabet") <> Empty) Then
            Me.germanAlphabet = params("germanAlphabet")
        End If
        If (params("lineThickness") <> Empty) Then
            Me.lineThickness = params("lineThickness")
        End If
        Set Me.renderer.controller = Me  '' TODO-GD needed for highlighting
        If params("foregroundColor") <> Empty Then
            Me.renderer.foregroundColor = params("foregroundColor")
        Else
            Me.renderer.foregroundColor = "currentColor"
        End If
        If (params("ariaLabel") <> Empty) Then
            Me.renderer.ariaLabel = params("ariaLabel")
        End If
        
        If (params("minPadding") <> Empty) Then
           Me.renderer.minPadding = params("minPadding")
        Else
           Me.renderer.minPadding = 0
        End If
        Me.reset

End Sub

Public Sub reset()

        Set Me.selected = New iArray
        Set Me.staffgroups = New iArray 'StaffGroupElement
        If Not (Me.engraver Is Nothing) Then
            Me.engraver.reset
        End If
        Set Me.engraver = Nothing
        Me.renderer.reset
        Set Me.dragTarget = Nothing
        Me.dragIndex = -1
        Set Me.dragMouseStart = New point
        Me.dragMouseStart.x = -1
        Me.dragMouseStart.y = -1
        Me.dragYStep = 0
        If (Me.lineThickness) Then
            Me.renderer.setLineThickness Me.lineThickness
        End If

End Sub

'    /**
'     * run the engraving process
'     */
Public Sub engraveABC(abctunes As TuneData, Optional tuneNumber As Integer = 0, Optional lineOffset As Integer = 0)
'        If (abctunes(0) = Empty) Then
'            abctunes = Array(abctunes)
'        End If
        Me.reset
        Dim i As Integer
        'For i = 0 To i < abctunes.Count
            If (tuneNumber = 0) Then
                tuneNumber = i
            End If
            Set Me.GetFontAttr = New GetFontAndAttr
            Set Me.getTxtSize = New getTextSize
            Me.GetFontAttr.init abctunes.formatting, Me.classes
            Me.getTxtSize.init Me.GetFontAttr, Me.renderer.paper
            Me.engraveTune abctunes, tuneNumber, lineOffset
        'End If
End Sub

'    /**
'     * Some of the items on the page are not scaled, so adjust them in the opposite direction of scaling to cancel out the scaling.
'     */
Public Sub adjustNonScaledItems(scale_ As Double)
        Me.width = Me.width / scale_
        Me.renderer.adjustNonScaledItems scale_
End Sub
Public Function getMeasureWidths(abctune As Dictionary) As iArray

        Me.reset
        Set Me.GetFontAttr = New GetFontAndAttr
        Set Me.getTxtSize = New getTextSize
        Me.GetFontAttr.init abctune.formatting, Me.classes
        Me.getTxtSize.init Me.GetFontAttr, Me.renderer.paper

        Call Me.setupTune(abctune, 0)
        Call Me.constructTuneElements(abctune)
        '' layout() sets the x-coordinate of the abcTune element here:
        '' abcTune.lines[0].staffGroup.voices[0].children[0].x
        Call layout(Me.renderer, abctune, 0, Me.space)

        Dim ret As New iArray
        Dim section As New Dictionary
        Dim abcLine
        Dim i As Integer
        Dim needNewSection   As Boolean
        
        needNewSection = True
        For i = 0 To i < abctune.lines.Count
            abcLine = abctune.lines(i)
            If Not (abcLine.Staff Is Nothing) Then
                If (needNewSection = True) Then
                    section("left") = 0
                    Set section("measureWidths") = New iArray
                    section("height") = Me.renderer.padding.top + Me.renderer.spacing.music + Me.renderer.padding.bottom + 24  '' the 24 is the empirical value added to the bottom of all tunes.
                    section("total") = 0
                    
                    ret.Push section
                    needNewSection = False
                End If
                '' At this point, the voices are laid out so that the bar lines are even with each other. So we just need to get the placement of the first voice.
                If (abcLine.staffGroup.voices.Length > 0) Then
                    Dim voice As VoiceElement
                    Dim foundNotStaffExtra  As Boolean
                    Dim lastXPosition As Double
                    Dim k As Integer
                    Dim child As AbsoluteElement
                    Set voice = abcLine.staffGroup.voices(0)
                    foundNotStaffExtra = False
                    lastXPosition = 0
                    For k = 0 To voice.children.Count - 1
                        Set child = voice.children(k)
                        If (foundNotStaffExtra = False And child.isClef = False And child.isKeySig = False) Then
                            foundNotStaffExtra = True
                            section.left = child.x
                            lastXPosition = child.x
                        End If
                        If (child.typs = "bar") Then
                            section("measureWidths").Push (child.x - lastXPosition)
                            section("total") = section("total") + (child.x - lastXPosition)
                            lastXPosition = child.x
                        End If
                    Next
                End If
                ''section.height += calcHeight(abcLine.staffGroup) * spacing.STEP
             Else
                needNewSection = True
        Next
        Set getMeasureWidths = ret

End Function
Public Function setupTune(abctune As TuneData, tuneNumber As Integer) As Double
        'line Me.classes.reset

        If (abctune.formatting.jazzchords = Empty) Then
            Me.jazzchords = abctune.formatting.jazzchords
        End If
        
        Me.renderer.newTune abctune
        Dim tDic As New Dictionary
        tDic("bagpipes") = abctune.formatting.bagpipes
        tDic("flatbeams") = abctune.formatting.flatBeams
        tDic("graceSlurs") = abctune.formatting.graceSlurs <> False '' undefined is the default, which is true
        Set tDic("percmap") = abctune.formatting.percmap
        tDic("initialClef") = Me.initialClef
        tDic("jazzchords") = Me.jazzchords
        tDic("germanAlphabet") = Me.germanAlphabet
        Set Me.engraver = n.AbstractEng(Me.getTxtSize, tuneNumber, tDic)
        Me.engraver.setStemHeight Me.renderer.spacingRenderer.stemHeight
        Me.engraver.measureLength = abctune.getMeterFraction().num / abctune.getMeterFraction().den
        If (abctune.formatting.staffwidth <> 0) Then
            Me.width = abctune.formatting.staffwidth * 1.33  '' The width is expressed in pt  convert to px.
         Else
            If (Me.renderer.isPrint = True) Then
                Me.width = Me.staffwidthPrint
            Else
                Me.width = Me.staffwidthScreen
            End If
        End If

        Dim scale_ As Double
        If abctune.formatting.scale_ <> 0 Then
            scale_ = abctune.formatting.scale_
        Else
            scale_ = Me.scale_
        End If
        If (Me.responsive = "resize") Then '' The resizing will mess with the scaling, so just don't do it explicitly.
            scale_ = 0
        End If
        If (scale_ = 0) Then
            If Me.renderer.isPrint = True Then
                scale_ = 0.75
            Else
                scale_ = 1
            End If
        End If
        Me.adjustNonScaledItems scale_
        setupTune = scale_
End Function
    
Public Sub constructTuneElements(abctune As TuneData)
'/** 生成原始五線譜數據  */

        'line abcTune.topText = new TopText(abcTune.metaText, abcTune.metaTextInfo, abcTune.formatting, abcTune.lines, me.width, me.renderer.isPrint, me.renderer.padding.left, me.renderer.spacing, me.GetTxtSize)

        '' Generate the raw staff line data
        Dim i As Integer
        Dim abcLine As TuneLine
        Dim hasPrintedTempo  As Boolean
        Dim hasSeenNonSubtitle As Boolean
        Dim tmpPrintTempo As vtempoProperties
        hasPrintedTempo = False
        hasSeenNonSubtitle = False
        If hasPrintedTempo = False Then
            Set tmpPrintTempo = abctune.metaText.tempo
        End If
        
        '自訂
        Set abcLine = New TuneLine
        Set abcLine.staffGroup = Me.engraver.createABCLine(abctune.lines(0).Staffs, tmpPrintTempo, i)
        
'        for (i = 0  i < abcTune.lines.length  i++) {
'            abcLine = abcTune.lines[i]
'            if (abcLine.staff) {
'                hasSeenNonSubtitle = True
'                abcLine.staffGroup = me.engraver.createABCLine(abcLine.staff, !hasPrintedTempo ? abcTune.metaText.tempo : null, i)
'                hasPrintedTempo = True
'            } else if (abcLine.subtitle) {
'                '' If the subtitle is at the top, then it was already accounted for. So skip all subtitles until the first non-subtitle line.
'                if (hasSeenNonSubtitle) {
'                    var center = Me.width / 2 + Me.renderer.padding.left
'                    abcLine.nonMusic = new Subtitle(me.renderer.spacing.subtitle, abcTune.formatting, abcLine.subtitle, center, me.renderer.padding.left, me.GetTxtSize)
'                }
'            } else if (abcLine.text !== undefined) {
'                hasSeenNonSubtitle = True
'                abcLine.nonMusic = new FreeText(abcLine.text, abcLine.vskip, me.GetFontAttr, me.renderer.padding.left, me.width, me.GetTxtSize)
'            } else if (abcLine.separator !== undefined && abcLine.separator.lineLength) {
'                hasSeenNonSubtitle = True
'                abcLine.nonMusic = new Separator(abcLine.separator.spaceAbove, abcLine.separator.lineLength, abcLine.separator.spaceBelow)
'            }
'        }
'        abcTune.bottomText = new BottomText(abcTune.metaText, me.width, me.renderer.isPrint, me.renderer.padding.left, me.renderer.spacing, me.GetTxtSize)
End Sub

 Public Sub engraveTune(abctune As TuneData, tuneNumber As Integer, lineOffset As Integer)

 
        Dim scale_ As Double
        scale_ = Me.setupTune(abctune, tuneNumber)

        '' Create all of the element objects that will appear on the page.
        '' 建立將出現在頁面上的所有元素物件。
        '' 將建立 StaffGroup 物件
        Me.constructTuneElements abctune

        '' Do all the positioning, both horizontally and vertically
        '' 進行所有的水平和垂直定位 x 的位置
        Dim maxWidth  As Double
        maxWidth = layout(Me.renderer, abctune, Me.width, Me.space)

        '' Deal with tablature for staff
        '' 處理五線譜的樂譜
'        If Not (abctune.tablatures Is Nothing) Then
'            tablatures.layoutTablatures Me.renderer, abctune
'        End If

'        '' Do all the writing to the SVG
'        '' 用 StaffGroup 將所有內容寫入 SVG
'        Dim ret
'        Set ret = draw(Me.renderer, Me.classes, abctune, Me.width, maxWidth, Me.responsive, scale_, Me.selectTypes, tuneNumber, lineOffset)
'        Set Me.staffgroups = ret.staffgroups
'        Set Me.selectables = ret.selectables
'
'        If (Me.oneSvgPerLine) Then
'            var div = Me.renderer.paper.svg.parentNode
'            Me.svgs = splitSvgIntoLines(Me.renderer, div, abctune.metaText.title, Me.responsive)
'         Else
'            Me.svgs = Array(Me.renderer.paper.svg)
'        End If
'        setupSelection Me, Me.svgs
End Sub


Public Function getDim(historyEl As Dictionary) As Dictionary


        '' Get the dimensions on demand because the getBBox call is expensive.
        If (historyEl("dim") = Empty) Then
            Dim sRect As Dictionary
            Dim tdict As New Dictionary
            sRect = historyEl.svgEl.getBBox()
            
            tdict("left") = Int(sRect("x"))
            tdict("top") = Int(sRect("y"))
            tdict("rigth") = Int(sRect("x") + sRect("width"))
            tdict("bottom") = Int(sRect("y") + sRect("height"))
            Set historyEl("dim") = tdict
        End If
        getDim = historyEl("dim")
End Function

Public Sub addSelectListener(clickListener)
        Me.listeners(Me.listeners.Count) = clickListener
    End Sub

