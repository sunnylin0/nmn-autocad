VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GetFontAndAttr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public formatting As vFormatting
Public classes As Dictionary '     classes
Public Sub init(formatting As vFormatting, classes As Dictionary)
        Set Me.formatting = formatting
        Set Me.classes = classes
    End Sub

Public Sub updateFonts(fontOverrides As Staff)

   If Not fontOverrides.gchordfont Is Nothing Then Set Me.formatting.gchordfont = fontOverrides.gchordfont
   If Not fontOverrides.tripletfont Is Nothing Then Set Me.formatting.tripletfont = fontOverrides.tripletfont
   If Not fontOverrides.annotationfont Is Nothing Then Set Me.formatting.annotationfont = fontOverrides.annotationfont
   If Not fontOverrides.vocalfont Is Nothing Then Set Me.formatting.vocalfont = fontOverrides.vocalfont


End Sub

Public Function calc(typs, klass) As Dictionary


        Dim font As vFont
        If (VarType(typs) = vbString) Then
            Set font = Me.formatting(typs)
            '' Raphael deliberately changes the font units to pixels for some reason, so we need to change points to pixels here.
            If Not font Is Nothing Then
                font.face = font.face
                font.Size = Int(font.Size * 4 / 3)
                font.decoration = font.decoration
                font.style = font.style
                font.weight = font.weight
                font.box = font.box
                
            Else
                font.face = "Arial"
                font.Size = Int(12 * 4 / 3)
                font.decoration = "underline"
                font.style = "normal"
                font.weight = "normal"
         Else
             If TypeOf typs Is vFont Then
                font.face = typs.face
                font.Size = Int(typs.Size * 4 / 3)
                font.decoration = typs.decoration
                font.style = typs.style
                font.weight = typs.weight
                font.box = typs.box
            End If
        End If
        Dim paddingPercent
        paddingPercent = IIf(Me.formatting.fontboxpadding = True, Me.formatting.fontboxpadding, 0.1)
        
        font.padding = font.Size * paddingPercent

        Dim attr As New Dictionary
        attr("font-size") = font.Size
        attr("font-style") = font.style
        attr("font-family") = font.face
        attr("font-weight") = font.weight
        attr("text-decoration") = font.decoration
        attr("class") = Me.classes.generate(klass)
        Set calc = New Dictionary
        Set calc("font") = font
        Set calc("attr") = attr
End Function
