VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GlyphsModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
''
''import spacing from '../helpers/spacing'
''import Svg from '../svg'

Public glyphs As New Dictionary


Public Function pathClone(pathArray)
    Dim res(UBound(pathArray))
    'Dim charr()
    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To UBound(pathArray) - 1
        res [i] = []
        Dim charr(UBound(pathArray(i)))
        For j = 0 To UBound(pathArray(i)) - 1
        
            charr(j) = pathArray(i, j)
        Next
        res(i) = charr
    Next
    pathClone = res
End Function

Public Sub pathScale(ByRef pathArray, kx, ky)
    Dim i As Integer, ii As Integer
    Dim j As Integer, jj As Integer
    Dim p
    ii = UBound(pathArray) - 1
    For i = 0 To ii
        Set p = pathArray(i)
        jj = UBound(p) - 1
        For j = 1 To jj
            p(j) = p(j) * IIf((j Mod 2), kx, ky)
        Next
    Next
End Sub

Public Function printSymbol(x As Double, y As Double, symb As String, paper As SVGmodel, attrs As Dictionary)
        Dim pathArray, path, i
        If (glyphs(symb) Is Nothing) Then
            Set printSymbol = Nothing
        End If

        Set pathArray = pathClone(glyphs(symb).d)
        pathArray(0)(1) = pathArray(0)(1) + x
        pathArray(0)(2) = pathArray(0)(2) + y
        path = ""
        For i = 0 To UBound(pathArray) - 1
            path = path + Join(pathArray(i), " ")
        Next
        attrs.path = path
        Debug.Print "svg 要繪圖的地方: code:123ew81as2"
        'return paper.path(attrs)
End Function

Public Function getPathForSymbol(x As Double, y As Double, symb As String, scalex As Double, scaley As Double)
        scalex = IIf(scalex, scalex, 1)
        scaley = IIf(scaley, scaley, 1)
        If (glyphs(symb) Is Nothing) Then
            Set getPathForSymbol = Nothing
            Exit Function
        End If
        Dim pathArray
        pathArray = pathClone(glyphs(symb).d)
        If (scalex <> 1 Or scaley <> 1) Then pathScale pathArray, scalex, scaley
        pathArray(0, 1) = pathArray(0, 1) + x
        pathArray(0, 2) = pathArray(0, 2) + y

        getPathForSymbol = pathArray
End Function

Public Function getSymbolWidth(symbol As String) As Double
        If Not (glyphs(symbol) Is Nothing) Then
            getSymbolWidth = glyphs(symbol).w
        Else
            getSymbolWidth = 0
        End If
End Function

Public Function symbolHeightInPitches(symbol As String) As Double
        Dim height As Double
        height = IIf(Not glyphs(symbol) Is Nothing, glyphs(symbol).h, 0)
        symbolHeightInPitches = height / spacing.Step
End Function

Public Function getSymbolAlign(symbol As String) As String
        If (InStr(1, symbol, "scripts") > 0 And _
            symbol <> "scripts.roll") Then
            getSymbolAlign = "center"
        End If
        getSymbolAlign = "left"
End Function

Public Function getYCorr(symbol As String) As Integer
        Select Case (symbol)
            Case "0", "1", "2":
            Case "3", "4", "5", "6", "7", "8", "9", "+":
                getYCorr = -1
            Case "timesig.common", "timesig.cut": getYCorr = 0
            Case "flags.d32nd": getYCorr = -1
            Case "flags.d64th": getYCorr = -2
            Case "flags.u32nd": getYCorr = 1
            Case "flags.u64th": getYCorr = 3
            Case "rests.whole": getYCorr = 1
            Case "rests.half": getYCorr = -1
            Case "rests.8th": getYCorr = -1
            Case "rests.quarter": getYCorr = -1
            Case "rests.16th": getYCorr = -1
            Case "rests.32nd": getYCorr = -1
            Case "rests.64th": getYCorr = -1
            Case "f", "m", "p", "s", "z":
                getYCorr = -4
            Case "scripts.trill", "scripts.upbow", "scripts.downbow":
                getYCorr = -2
            Case "scripts.ufermata", "scripts.wedge", "scripts.roll", "scripts.shortphrase", "scripts.longphrase":
                getYCorr = -1
            Case "scripts.dfermata":
                getYCorr = 1
            Case Else
                getYCorr = 0
        End Select
End Function
Public Sub setSymbol(Name, path)
        glyphs("name") = path
    
End Sub

Function createSym(d As String, w As Double, h As Double) As vGlyphTyps
    Set createSym = New vGlyphTyps
    createSym.d = d
    createSym.w = w
    createSym.h = h

End Function

Private Sub Class_Initialize()


    
    Set glyphs("note.0") = createSym("path as string", 14, 16)
    Set glyphs("note.1") = createSym("path as string", 14, 16)
    Set glyphs("note.2") = createSym("path as string", 14, 16)
    Set glyphs("note.3") = createSym("path as string", 14, 16)
    Set glyphs("note.4") = createSym("path as string", 14, 16)
    Set glyphs("note.5") = createSym("path as string", 14, 16)
    Set glyphs("note.6") = createSym("path as string", 14, 16)
    Set glyphs("note.7") = createSym("path as string", 14, 16)
    Set glyphs("note.8") = createSym("path as string", 14, 16)
    Set glyphs("note.9") = createSym("path as string", 14, 16)
    Set glyphs("clefs.none") = createSym("path as string", 9, 9)
    Set glyphs("0") = createSym("path as string", 10.78, 14.959)
    Set glyphs("1") = createSym("path as string", 8.94, 15.058)
    Set glyphs("2") = createSym("path as string", 10.764, 14.97)
    Set glyphs("3") = createSym("path as string", 9.735, 14.967)
    Set glyphs("4") = createSym("path as string", 11.795, 14.994)
    Set glyphs("5") = createSym("path as string", 10.212, 14.997)
    Set glyphs("6") = createSym("path as string", 9.956, 14.982)
    Set glyphs("7") = createSym("path as string", 10.561, 15.093)
    Set glyphs("8") = createSym("path as string", 10.926, 14.989)
    Set glyphs("9") = createSym("path as string", 9.959, 14.986)
    Set glyphs("rests.multimeasure") = createSym("path as string", 42, 18)
    Set glyphs("rests.whole") = createSym("path as string", 11.25, 4.68)
    Set glyphs("rests.half") = createSym("path as string", 11.25, 4.68)
    Set glyphs("rests.quarter") = createSym("path as string", 7.888, 21.435)
    Set glyphs("rests.8th") = createSym("path as string", 7.534, 13.883)
    Set glyphs("rests.16th") = createSym("path as string", 9.724, 21.383)
    Set glyphs("rests.32nd") = createSym("path as string", 11.373, 28.883)
    Set glyphs("rests.64th") = createSym("path as string", 12.453, 36.383)
    Set glyphs("rests.128th") = createSym("path as string", 12.992, 43.883)
    Set glyphs("accidentals.sharp") = createSym("path as string", 8.25, 22.462)
    Set glyphs("accidentals.halfsharp") = createSym("path as string", 5.25, 20.174)
    Set glyphs("accidentals.nat") = createSym("path as string", 5.4, 22.8)
    Set glyphs("accidentals.flat") = createSym("path as string", 6.75, 18.801)
    Set glyphs("accidentals.halfflat") = createSym("path as string", 6.728, 18.801)
    Set glyphs("accidentals.dblflat") = createSym("path as string", 12.1, 18.804)
    Set glyphs("accidentals.dblsharp") = createSym("path as string", 7.95, 7.977)
    Set glyphs("dots.dot") = createSym("path as string", 3.45, 3.45)
    Set glyphs("noteheads.dbl") = createSym("path as string", 16.83, 8.145)
    Set glyphs("noteheads.whole") = createSym("path as string", 14.985, 8.097)
    Set glyphs("noteheads.half") = createSym("path as string", 10.37, 8.132)
    Set glyphs("noteheads.quarter") = createSym("path as string", 9.81, 8.094)
    Set glyphs("noteheads.slash.nostem") = createSym("path as string", 12.81, 15.63)
    Set glyphs("noteheads.indeterminate") = createSym("path as string", 9.843, 8.139)
    Set glyphs("scripts.ufermata") = createSym("path as string", 19.748, 11.289)
    Set glyphs("scripts.dfermata") = createSym("path as string", 19.744, 11.274)
    Set glyphs("scripts.sforzato") = createSym("path as string", 13.5, 7.5)
    Set glyphs("scripts.staccato") = createSym("path as string", 2.989, 3.004)
    Set glyphs("scripts.tenuto") = createSym("path as string", 8.985, 1.08)
    Set glyphs("scripts.umarcato") = createSym("path as string", 7.5, 8.245)
    Set glyphs("scripts.dmarcato") = createSym("path as string", 7.5, 8.25)
    Set glyphs("scripts.stopped") = createSym("path as string", 8.295, 8.295)
    Set glyphs("scripts.upbow") = createSym("path as string", 9.73, 15.608)
    Set glyphs("scripts.downbow") = createSym("path as string", 11.22, 9.992)
    Set glyphs("scripts.turn") = createSym("path as string", 16.366, 7.893)
    Set glyphs("scripts.trill") = createSym("path as string", 17.963, 16.49)
    Set glyphs("scripts.segno") = createSym("path as string", 15, 22.504)
    Set glyphs("scripts.coda") = createSym("path as string", 16.035, 21.062)
    Set glyphs("scripts.comma") = createSym("path as string", 3.042, 9.237)
    Set glyphs("scripts.roll") = createSym("path as string", 10.817, 6.125)
    Set glyphs("scripts.prall") = createSym("path as string", 15.011, 7.5)
    Set glyphs("scripts.arpeggio") = createSym("path as string", 5, 10)
    Set glyphs("scripts.mordent") = createSym("path as string", 15.011, 10.012)
    Set glyphs("flags.u8th") = createSym("path as string", 6.692, 22.59)
    Set glyphs("flags.u16th") = createSym("path as string", 6.693, 26.337)
    Set glyphs("flags.u32nd") = createSym("path as string", 6.697, 32.145)
    Set glyphs("flags.u64th") = createSym("path as string", 6.682, 39.694)
    Set glyphs("flags.d8th") = createSym("path as string", 8.492, 21.691)
    Set glyphs("flags.ugrace") = createSym("path as string", 12.019, 9.954)
    Set glyphs("flags.dgrace") = createSym("path as string", 15.12, 9.212)
    Set glyphs("flags.d16th") = createSym("path as string", 8.475, 22.591)
    Set glyphs("flags.d32nd") = createSym("path as string", 8.385, 29.191)
    Set glyphs("flags.d64th") = createSym("path as string", 8.485, 32.932)
    Set glyphs("clefs.C") = createSym("path as string", 20.31, 29.97)
    Set glyphs("clefs.F") = createSym("path as string", 20.153, 23.142)
    Set glyphs("clefs.G") = createSym("path as string", 19.051, 57.057)
    Set glyphs("clefs.perc") = createSym("path as string", 21, 14.97)
    Set glyphs("tab.big") = createSym("path as string", 19.643, 43.325)
    Set glyphs("tab.tiny") = createSym("path as string", 15.709, 34.656)
    Set glyphs("timesig.common") = createSym("path as string", 13.038, 15.689)
    Set glyphs("timesig.cut") = createSym("path as string", 13.038, 20.97)
    Set glyphs("timesig.imperfectum") = createSym("path as string", 13.038, 20.97)
    Set glyphs("timesig.imperfectum2") = createSym("path as string", 13.038, 20.97)
    Set glyphs("timesig.perfectum") = createSym("path as string", 13.038, 20.97)
    Set glyphs("timesig.perfectum2") = createSym("path as string", 13.038, 20.97)
    Set glyphs("f") = createSym("path as string", 16.155, 19.445)
    Set glyphs("m") = createSym("path as string", 14.687, 9.126)
    Set glyphs("p") = createSym("path as string", 14.689, 13.127)
    Set glyphs("r") = createSym("path as string", 9.41, 9.132)
    Set glyphs("s") = createSym("path as string", 6.632, 8.758)
    Set glyphs("z") = createSym("path as string", 8.573, 8.743)
    Set glyphs("+") = createSym("path as string", 7.507, 7.515)
    Set glyphs(",") = createSym("path as string", 3.452, 8.143)
    Set glyphs("-") = createSym("path as string", 5.001, 0.81)
    Set glyphs(".") = createSym("path as string", 3.413, 3.402)
    Set glyphs("scripts.wedge") = createSym("path as string", 7.49, 7.752)
    Set glyphs("scripts.thumb") = createSym("path as string", 5.955, 9.75)
    Set glyphs("scripts.open") = createSym("path as string", 5.955, 7.5)
    Set glyphs("scripts.longphrase") = createSym("path as string", 2.16, 23.04)
    Set glyphs("scripts.mediumphrase") = createSym("path as string", 2.16, 15.54)
    Set glyphs("scripts.shortphrase") = createSym("path as string", 2.16, 8.04)
    Set glyphs("scripts.snap") = createSym("path as string", 10.38, 6.84)
    Set glyphs("noteheads.slash.whole") = createSym("path as string", 10.81, 15.63)
    Set glyphs("noteheads.slash.quarter") = createSym("path as string", 9, 9)
    Set glyphs("noteheads.harmonic.quarter") = createSym("path as string", 7.5, 8.165)
    Set glyphs("noteheads.triangle.quarter") = createSym("path as string", 9, 9)

End Sub
