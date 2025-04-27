VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aFontItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public box As Boolean
Public pad
Public Name As String
Public weight As String
Public size
Public used As Boolean
Public fid
Public swfac
Public fname As String
Public style As String
Public src As String
Public Sub ConvertVarName(key, value)
        Select Case key    ' Evaluate Number.
            ''一般
            Case "box":    Me.box = value
            Case "pad":    Me.pad = value
            Case "name":   Me.Name = value
            Case "weight": Me.weight = value
            Case "size":   Me.size = value
            Case "used":   Me.used = value
            Case "fid":    Me.fid = value
            Case "swfac":  Me.swfac = value
            Case "fname":  Me.fname = value
            Case "style":  Me.style = value
            Case "src":    Me.src = value
            Case Else
                Debug.Print "沒有變數 " & key & "名"
        End Select
End Sub

