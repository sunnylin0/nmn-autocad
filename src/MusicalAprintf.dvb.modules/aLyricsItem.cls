VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aLyricsItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public t As String
Public font As aFontItem
Public iStart
Public iend
Public ln
Public shift


Public Sub ConvertVarName(key, value)
        Select Case key    ' Evaluate Number.
            ''一般
            Case "t":        Me.t = value
            Case "font":  Set Me.font = value
            Case "istart":   Me.iStart = value
            Case "iend":     Me.iend = value
            Case "ln":       Me.ln = value
            Case "shift":    Me.shift = value
            Case Else
                Debug.Print "沒有變數 " & key & "名"
        End Select
End Sub

