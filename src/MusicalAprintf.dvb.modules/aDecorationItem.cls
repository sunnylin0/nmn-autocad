VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aDecorationItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String
Public func
Public glyph As String
Public h
Public hd
Public wl
Public wr
Public str As String
Public dd_st As aDecorationItem
Public dd_en As aDecorationItem

Public Sub ConvertVarName(key, value)
        Select Case key    ' Evaluate Number.
            ''一般
            Case "name":    Me.Name = value
            Case "func":    Me.func = value
            Case "glyph":   Me.glyph = value
            Case "h":       Me.h = value
            Case "hd":      Me.hd = value
            Case "wl":      Me.wl = value
            Case "wr":      Me.wr = value
            Case "str":     Me.str = value
            Case "dd_st": Set Me.dd_st = value
            Case "dd_st": Set Me.dd_st = value
            Case Else
                Debug.Print "沒有變數 " & key & "名"
        End Select
End Sub

