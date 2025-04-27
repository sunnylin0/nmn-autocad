VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RendererModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

 ''    abc_renderer.js: API to render to SVG/Raphael/whatever rendering engine

'/**
' * Implements the API for rendering ABCJS Abstract Rendering Structure to a canvas/paper (e.g. SVG, Raphael, etc)
' * @param {Object} paper
' */
Public paper As SVGmodel      'svg
Public controller As EngraverController
Public space As Double
Public padding As vBox
Public paddingOverride As vBox
Public y As Double
Public abcTune As TuneData

Public path As String
Public isPrint As Boolean
Public lineThickness As Double
Public spacingRenderer As vSpacingRenderer
Public firefox112 As Boolean
Public foregroundColor As String
Public minPadding As Double
Public staffbottom As Double
Public showDebug As Boolean
Public ariaLabel As String


Public Sub init() 'paper As Dictionary)

        Set Me.paper = New SVGmodel  ' new Svg(paper)
        Me.paper.init
        Set Me.controller = Nothing

        Me.space = 3 * spacing.space
        'me.padding = {}   '' renderer's padding is managed by the controller
        Me.reset
        'me.firefox112 = navigator.userAgent.indexOf('Firefox/112.0') >= 0
End Sub


Public Sub reset()

        Me.paper.Clear  'clear()
        Me.y = 0
        Set Me.abcTune = Nothing
        Me.path = ""
        Me.isPrint = False
        Me.lineThickness = 0
        initVerticalSpace
End Sub

Public Sub newTune(abcTune As TuneData)
        Set Me.abcTune = abcTune   '' TODO-PER: this is just to get the font info.
        Me.setVerticalSpace abcTune.formatting
         ''me.measureNumber = null
         ''me.noteNumber = null
        Me.isPrint = (abcTune.media = "print")
        Me.setPadding abcTune
End Sub

 Public Sub setLineThickness(lineThickness As Integer)
        Me.lineThickness = lineThickness
End Sub

Public Sub setPaddingOverride(params As Dictionary)
        Set Me.paddingOverride = New vBox
        With Me.paddingOverride
            .top = params("paddingtop")
            .bottom = params("paddingbottom")
            .right = params("paddingright")
            .left = params("paddingleft")
        End With

End Sub
Public Function getPaddingVariable(self, paddingOverKey, formattingKey, printDefault, screenDefault) As Double
            Dim paddingKey As Double
            If (formattingKey <> 0) Then
                paddingKey = formattingKey
            ElseIf (paddingOverKey <> 0) Then
                paddingKey = paddingOverKey
            ElseIf (self.isPrint) Then
                paddingKey = printDefault
            Else
                paddingKey = screenDefault
            End If
            getPaddingVariable = paddingKey
End Function
Public Sub setPadding(abcTune As TuneData)
         '' If the padding is set in the tune, then use that.
         '' Otherwise, if the padding is set in the override, use that.
         '' Otherwise, use the defaults (there are a different set of defaults for screen and print.)

         '' 1cm x 0.393701in/cm x 72pt/in x 1.33px/pt = 38px
         '' 1.8cm x 0.393701in/cm x 72pt/in x 1.33px/pt = 68px
        Set padding = New vBox
        padding.top = getPaddingVariable(Me, paddingOverride.top, abcTune.formatting.topmargin, 38, 15)
        padding.bottom = getPaddingVariable(Me, paddingOverride.bottom, abcTune.formatting.botmargin, 38, 15)
        padding.left = getPaddingVariable(Me, paddingOverride.left, abcTune.formatting.leftmargin, 68, 15)
        padding.right = getPaddingVariable(Me, paddingOverride.right, abcTune.formatting.rightmargin, 68, 15)
End Sub

'    /**
'     * Some of the items on the page are not scaled, so adjust them in the opposite direction of scaling to cancel out the scaling.
'     * @param {float} scale
'     */
Public Sub adjustNonScaledItems(scale_ As Double)
        Me.padding.top = Me.padding.top / scale_
        Me.padding.bottom = Me.padding.bottom / scale_
        Me.padding.left = Me.padding.left / scale_
        Me.padding.right = Me.padding.right / scale_
        Set Me.abcTune.formatting.headerfont = New vFont        'line 沒有建立 要初始化
        Set Me.abcTune.formatting.footerfont = New vFont        'line 沒有建立 要初始化
        Me.abcTune.formatting.headerfont.size = Me.abcTune.formatting.headerfont.size / scale_
        Me.abcTune.formatting.footerfont.size = Me.abcTune.formatting.footerfont.size / scale_
End Sub

'    /**
'     * Set the the values for all the configurable vertical space options.
'     */
Public Sub initVerticalSpace()
         '' conversion: 37.7953 = conversion factor for cm to px.
         '' All of the following values are in px.
         Set Me.spacingRenderer = New vSpacingRenderer
        With Me.spacingRenderer
            .composer = 7.56   '' Set the vertical space above the composer.
            .graceBefore = 8.67   '' Define the space before, inside and after the grace notes.
            .graceInside = 10.67
            .graceAfter = 16
            .info = 0   '' Set the vertical space above the infoline.
            .lineSkipFactor = 1.1   '' Set the factor for spacing between lines of text. (multiply this by the font size)
            .music = 7.56   '' Set the vertical space above the first staff.
            .paragraphSkipFactor = 0.4   '' Set the factor for spacing between text paragraphs. (multiply this by the font size)
            .parts = 11.33   '' Set the vertical space above a new part.
            .slurHeight = 1#    '' Set the slur height factor.
            .staffSeparation = 61.33   '' Do not put a staff system closer than <unit> from the previous system.
            .stemHeight = 26.67 + 10   '' Set the stem height.
            .subtitle = 3.78   '' Set the vertical space above the subtitle.
            .systemStaffSeparation = 48   '' Do not place the staves closer than <unit> inside a system. * This values applies to all staves when in the tune header. Otherwise, it applies to the next staff
            .text = 18.9   '' Set the vertical space above the history.
            .title = 7.56   '' Set the vertical space above the title.
            .top = 30.24   ''Set the vertical space above the tunes and on the top of the continuation pages.
            .vocal = 0   '' Set the vertical space above the lyrics under the staves.
            .words = 0 '' Set the vertical space above the lyrics at the end of the tune.
        End With
'    /*
'        TODO-PER: Handle the x-coordinate spacing items, too.
'    maxshrink <float>Default: 0.65
'    Set how much to compress horizontally when music line breaks
'    are automatic.
'    <float> must be between 0 (natural spacing)
'    and 1 (max shrinking).
'
'     '' This next value is used to compute the natural spacing of
'     '' the notes. The base spacing of the crotchet is always
'     '' 40 pts. When the duration of a note type is twice the
'     '' duration of an other note type, its spacing is multiplied
'     '' by this factor.
'     '' The default value causes the note spacing to be multiplied
'     '' by 2 when its duration is multiplied by 4, i.e. the
'     '' space of the semibreve is 80 pts and the space of the
'     '' semiquaver is 20 pts.
'     '' Setting this value to 1 sets all note spacing to 40 pts.
'    noteSpacingFactor: 1.414,  '' Set the note spacing factor to <float> (range 1..2).
'
'    scale <float> Default: 0.75 Set the page scale factor. Note that the header and footer are not scaled.
'
'    stretchlast <float>Default: 0.8
'    Stretch the last music line of a tune when it exceeds
'    the <float> fraction of the page width.
'    <float> range is 0.0 to 1.0.
'         */
End Sub

Public Sub setVerticalSpace(formatting As vFormatting)
         '' conversion from pts to px 4/3
        If (formatting.staffsep <> 0) Then Me.spacingRenderer.staffSeparation = formatting.staffsep * 4 / 3
        If (formatting.composerspace <> 0) Then Me.spacingRenderer.composer = formatting.composerspace * 4 / 3
        If (formatting.partsspace <> 0) Then Me.spacingRenderer.parts = formatting.partsspace * 4 / 3
        If (formatting.textspace <> 0) Then Me.spacingRenderer.text = formatting.textspace * 4 / 3
        If (formatting.musicspace <> 0) Then Me.spacingRenderer.music = formatting.musicspace * 4 / 3
        If (formatting.titlespace <> 0) Then Me.spacingRenderer.title = formatting.titlespace * 4 / 3
        If (formatting.sysstaffsep <> 0) Then Me.spacingRenderer.systemStaffSeparation = formatting.sysstaffsep * 4 / 3
        If (formatting.subtitlespace <> 0) Then Me.spacingRenderer.subtitle = formatting.subtitlespace * 4 / 3
        If (formatting.topspace <> 0) Then Me.spacingRenderer.top = formatting.topspace * 4 / 3
        If (formatting.vocalspace <> 0) Then Me.spacingRenderer.vocal = formatting.vocalspace * 4 / 3
        If (formatting.wordsspace <> 0) Then Me.spacingRenderer.words = formatting.wordsspace * 4 / 3
End Sub


'    /**
'     * Calculates the y for a given pitch value (relative to the stave the renderer is currently printing)
'     * @param {number} ofs pitch value (bottom C on a G clef = 0, D=1, etc.)
'     */
Public Function calcY(ofs As Double) As Double
        calcY = Me.y - ofs * spacing.Step
End Function

Public Sub moveY(em As Double, numLines As Integer)
        If (numLines = 0) Then numLines = 1
        Me.y = Me.y + em * numLines
End Sub

Public Sub absolutemoveY(y As Double)
        Me.y = y
End Sub


