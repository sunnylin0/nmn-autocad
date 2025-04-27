VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "point"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public x As Double
Public Y As Double
Public z As Double

Sub c(xPt As Variant, yPt As Variant, zPt As Variant)
        x = xPt
        Y = yPt
        z = zPt
End Sub
Sub a(varPt As Variant)
    If UBound(varPt) = 2 Then
        x = varPt(0)
        Y = varPt(1)
        z = varPt(2)
    End If
    
End Sub

Function at() As Double()
    Dim dd(0 To 2) As Double
    dd(0) = x
    dd(1) = Y
    dd(2) = z
     at = dd
End Function
Function copy() As point
    Dim pp As point
    Set pp = New point
    pp.c x, Y, z
    Set copy = pp
End Function
Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()

End Sub
