VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "point_xxx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public x As Double
Public y As Double
Public Z As Double

Function c(xPt As Variant, yPt As Variant, Optional zPt As Variant = 0)
        x = xPt
        y = yPt
        Z = zPt
        Set c = Me
End Function
Function a(varPt As Variant)
    If UBound(varPt) = 2 Then
        x = varPt(0)
        y = varPt(1)
        Z = varPt(2)
    ElseIf UBound(varPt) = 1 Then
        x = varPt(0)
        y = varPt(1)
        Z = 0
    End If
    Set a = Me
End Function

Function at() As Double()
    Dim dd(0 To 2) As Double
    dd(0) = x
    dd(1) = y
    dd(2) = Z
     at = dd
End Function
Function Clone() As point
    Dim pp As point
    Set pp = New point
    pp.c x, y, Z
    Set Clone = pp
End Function

Function ToArray()
    Dim tarr(2)
    tarr(0) = x
    tarr(1) = y
    tarr(2) = Z
    ToArrayXYZ = tarr
End Function

'Private Sub Class_Initialize(Optional ix As Variant = 0, Optional iy As Variant = 0, Optional iz As Variant = 0)
'        X = ix
'        Y = iy
'        Z = iz
'End Sub
