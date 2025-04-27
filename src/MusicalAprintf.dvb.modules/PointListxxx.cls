VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PointListxxx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim pntList() As point
Dim mSize As Long
Dim mLogSize As Long
Sub clean()
    mSize = 0
End Sub
Function Size() As Long
    Size = mSize
End Function
Function list() As Double()
    If mSize <= 0 Then
        Exit Function
    End If

    Dim tmplist() As Double
    ReDim tmplist(0 To (mSize * 3) - 1)
    Dim i As Long
    For i = 0 To mSize - 1
        tmplist(i * 3) = pntList(i).x
        tmplist(i * 3 + 1) = pntList(i).y
        tmplist(i * 3 + 2) = pntList(i).Z
    Next
    list = tmplist
End Function
Function at(ln As Long) As point
    Dim item  As New point
    If ln < mSize Then
        Set item = pntList(ln)
    Else
        Set item = New point
    End If
    Set at = item
End Function

Sub Add(xi As Variant, yi As Variant, zi As Variant)
    If UBound(pntList) < mSize + 1 Then
        '°}¦C¤£°÷¡A¦b©µ¦ù
        ReDim Preserve pntList(UBound(pntList) + 100)
    End If
    
    Set pntList(mSize) = New point
    pntList(mSize).x = xi
    pntList(mSize).y = yi
    pntList(mSize).Z = zi
    mSize = mSize + 1
    
End Sub

Sub addpt(Pnt As point)

    If UBound(pntList) < mSize + 1 Then
        '°}¦C¤£°÷¡A¦b©µ¦ù
        ReDim Preserve pntList(UBound(pntList) + 100)
    End If
    
    Set pntList(mSize) = New point
    Set pntList(mSize) = Pnt
    mSize = mSize + 1
    
End Sub

Private Sub Class_Initialize()
    ReDim pntList(8)
End Sub
