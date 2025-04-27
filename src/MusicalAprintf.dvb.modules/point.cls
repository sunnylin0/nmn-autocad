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
Public y As Double
Public Z As Double

' #####################
' #    CONSTRUCTOR    #
' #####################
'Private Sub Class_Initialize()
'  Set iArray = New Collection
'End Sub

' #####################
' #  Getter / Setter  #
' #####################
'Attribute Element.VB_UserMemId = 0 saved in both procedures

'@DefaultMember
'Public Property Get element() As point
'    ' add Attribute element.VB_UserMemId = 0
'    Dim mm
'    Set mm = Me
'   Set element = mm
'End Property

'@DefaultMember
Public Property Let element(ByVal tmp As Variant)
Attribute element.VB_UserMemId = 0
    'add Attribute element.VB_UserMemId = 0
    If TypeOf tmp Is point Then
        x = tmp.x
        y = tmp.y
        Z = tmp.Z
    ElseIf (VarType(tmp) And vbArray) = vbArray Then
        If UBound(tmp) = 2 Then
            x = tmp(0)
            y = tmp(1)
            Z = tmp(2)
        ElseIf UBound(tmp) = 1 Then
            x = tmp(0)
            y = tmp(1)
            Z = 0
        End If
    ElseIf VarType(tmp) And vbVariant Then
              
    End If

End Property

' #############################
' # COMMON COLLECTION METHODS #
' #############################


Function c(xPt As Variant, yPt As Variant, Optional zPt As Variant = 0)
        x = xPt
        y = yPt
        Z = zPt
        Set c = Me
End Function
Function a(varPt As Variant)
    If TypeOf varPt Is point Then
        x = varPt.x
        y = varPt.y
        Z = varPt.Z
    ElseIf UBound(varPt) >= 2 Then
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
Function clone() As point
    Dim pp As point
    Set pp = New point
    pp.c x, y, Z
    Set clone = pp
End Function

Function ToArray()
    Dim tarr(2)
    tarr(0) = x
    tarr(1) = y
    tarr(2) = Z
    ToArray = tarr
End Function
Function ToDouble() As Double()
    Dim arrD(2) As Double
    arrD(0) = x
    arrD(1) = y
    arrD(2) = Z
     ToDouble = arrD
End Function

