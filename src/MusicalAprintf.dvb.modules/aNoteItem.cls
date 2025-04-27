VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aNoteItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public pit
Public shhd
Public shac
Public dur
Public midi
Public jn
Public jo
Public opit
Public acc
Public a_dd As iArray

Public Sub ConvertVarName(key, value)
        Select Case key    ' Evaluate Number.
            ''一般
            Case "pit":   Me.pit = value
            Case "shhd":  Me.shhd = value
            Case "shac":  Me.shac = value
            Case "dur":   Me.dur = value
            Case "midi":  Me.midi = value
            Case "jn":    Me.jn = value
            Case "jo":    Me.jo = value
            Case "acc":   Me.acc = value
            Case "a_dd":  Me.a_dd = value
            Case Else
                Debug.Print "沒有變數 " & key & "名"
        End Select
End Sub

