VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GetTextSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public gFontAndAttr As GetFontAndAttr
Public svg As SVGmodel        'svg
Public Sub init(g_FontAndAttr As GetFontAndAttr, Optional svg As SVGmodel)
        Set Me.gFontAndAttr = g_FontAndAttr
        Set Me.svg = svg
End Sub

Public Sub updateFonts(fontOverrides As Staff)
        Me.gFontAndAttr.updateFonts fontOverrides
End Sub

Public Function attr(typs, klass)
        attr = Me.gFontAndAttr.calc(typs, klass)
End Function

Public Function calc(text, typs, klass, Optional el As Dictionary) As Size
        Dim hash As Dictionary
        '' This can be passed in either a string or a font. If it is a string it names one of the standard fonts.
        If (VarType(typs) = vbString) Then
            hash = Me.attr(typs, klass)
        Else
            Dim f As New vFont
            Dim ar As New Dictionary
            f.face = typs.face
            f.Size = typs.Size
            f.decoration = typs.decoration
            f.style = typs.style
            f.weight = typs.weight
            
            aa("font-size") = typs.Size
            aa("font-style") = typs.style
            aa("font-family") = typs.face
            aa("font-weight") = typs.weight
            aa("text-decoration") = typs.decoration
            aa("class") = Me.GetFontAttr.classes.generate(klass)
            hash("font") = f
            hash("attr") = ar
        End If
        
        Dim si As Size
        si = Me.svg.getTxtSize(text, hash("attr"), el)
        If (hash.font.box) Then
            '' Add padding and an equal margin to each side.
            si.height = Size.height + hash.font.padding * 4
            si.width = Size.width + hash.font.padding * 4
        End If
        calc = si
End Function

Public Function baselineToCenter(text As String, typs As String, klass As String, Index As Integer, total As Double) As Double
        '' This is for the case where SVG wants to use the baseline of the first line as the Y coordinate.
        '' If there are multiple lines of text or there is an array of text then that will not be centered so this adjusts it.
        Dim height  As Double
        Dim fontHeight  As Double
        height = Me.calc(text, typs, klass).height
        fontHeight = Me.attr(typs, klass).font.Size

        baselineToCenter = height * 0.5 + (total - Index - 2) * fontHeight
End Function
