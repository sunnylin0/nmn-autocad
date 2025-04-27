VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aGchordItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public typs As String
Public text As String
Public x
Public pos
Public otext As String
Public iStart
Public iend
Public font As aFontItem

Public Sub ConvertVarName(key, value)
        Select Case key    ' Evaluate Number.
            ''一般
            Case "type":    Me.typs = value
            Case "text":    Me.text = value
            Case "name":    Me.Name = value
            Case "x":       Me.x = value
            Case "pos":     Me.pos = value
            Case "used":    Me.used = value
            Case "otext":   Me.otext = value
            Case "istart":  Me.iStart = value
            Case "iend":  Me.iend = value
            Case "font":  Set Me.font = value
            Case Else
                Debug.Print "沒有變數 " & key & "名"
        End Select
End Sub

