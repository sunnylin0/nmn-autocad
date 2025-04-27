VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "runTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'MIT License
'
'Copyright (c) 2019-2021 Petr Katerinak
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.
'
' Version of iArray: 1.0
''修正 元素從 0 開始

Private iArray As Collection
Private dict As Dictionary
Private i As Long ' Counter

' #####################
' #    CONSTRUCTOR    #
' #####################
Private Sub Class_Initialize()
  Set iArray = New Collection
  Set dict = New Dictionary
End Sub

' #####################
' #  Getter / Setter  #
' #####################
'Attribute Element.VB_UserMemId = 0 saved in both procedures

'@DefaultMember
Public Sub element(ByVal timeName As String)
Attribute element.VB_UserMemId = 0
    If timeName = "" Or timeName = " " Then
        
        Exit Sub
    End If
    Dim tt
    Dim iaa As iArray
    If IsObject(dict(timeName)) Then
        dict(timeName).Push Timer
        
    ElseIf IsEmpty(dict(timeName)) Then
    
        Set dict(timeName) = New iArray
        dict(timeName).Push Timer
        
    End If
    tt = dict.keys
    
    
 
End Sub
'
''@DefaultMember
'Public Property Let element(ByVal index As Long, ByVal val As Variant)
'  If index + 1 > iArray.Count Then
'    Dim eMsg As String
'    eMsg = IIf(iArray.Count > 0, _
'               "Index (" & index + 1 & ") out of iArray bounds (1-" & iArray.Count & ")", _
'               "iArray doesn't have any indexes")
'    err.Raise Number:=9, Description:=eMsg
'    Exit Property
'  End If
'  iArray.Remove index + 1
'  iArray.Add val, before:=index + 1
'End Property

' #############################
' # COMMON COLLECTION METHODS #
' #############################
Public Function Count() As Long
  Count = iArray.Count
End Function
Public Sub Add(ByVal val As Variant, Optional ByVal bef As Variant = Empty, Optional ByVal aft As Variant = Empty)
  If IsEmpty(bef) And IsEmpty(aft) Then
    iArray.Add val
  ElseIf IsEmpty(bef) Then
    iArray.Add val, , , aft
  ElseIf IsEmpty(aft) Then
    iArray.Add val, , bef
  Else
    iArray.Add val, , bef, aft
  End If
End Sub
Public Sub Remove(index As Long)
  iArray.Remove index + 1
End Sub

'Clear
' Empties iArray
Public Function Clear() As Boolean
  For i = iArray.Count - 1 To 0 Step -1
    iArray.Remove i + 1
  Next i
  Clear = IIf(iArray.Count = 0, True, False)
End Function

'Clone
' Makes a hard copy of the iArray
' @return iArray | Return exact copy of itself
Public Function clone() As iArray
  Set clone = New iArray
  For i = 0 To Me.Count - 1
    clone.Push Me(i)
  Next i
End Function


'Sum
' Calculates the sum of the numeric iArray.
' @return Variant | Sum value / "NaN" if iArray contains non-numeric value(s)
Public Function sum() As Variant
  If Not Me.ContainsOnlyNumeric Then sum = "NaN": Exit Function
  sum = 0
  Dim el As Variant
  For Each el In iArray
    sum = sum + el
  Next el
End Function



Public Function ToList()
  If dict.Count = 0 Then
    Exit Function
  End If
  
  Dim sn As String
  Dim j As Integer
  Dim k As Integer
  Dim iaa As iArray
  Dim sum As Double
  Dim stime As Double
  Dim etime As Double
  Debug.Print "******** run Time sum   *********"
  For j = 0 To UBound(dict.keys)
    sn = dict.keys(j)
    Set iaa = dict(sn)
    If iaa.Count Mod 2 <> 0 Then
        Debug.Print sn & " 資料沒結束."
    End If
    sum = 0
    For k = 0 To iaa.Count - 1 Step 2
        If k + 1 < iaa.Count Then
           sum = sum + iaa(k + 1) - iaa(k)
        End If
    Next
    Debug.Print Format(sn, "@@@@@@@@@@@@@@@@") & " : " & Format(CStr(sum), "000.000000") & "秒"
  Next
  
  
End Function

'ToString
' Creates string representation of the iArray.
' @param String delimiter | Optional character to separate the iArray's elements (default = ";")
' @return String | Formated representation of tha iArray
Public Function ToString(Optional ByVal delimiter As String = ";") As String
  If Me.Count = 0 Then ToString = "{}": Exit Function
  ToString = vbNullString
  For i = 0 To Me.Count - 1
    If i = 0 Then ToString = ToString + "{"
    If i > 0 Then ToString = ToString + delimiter
    If VarType(Me(i)) = vbString Then
        ToString = ToString + Chr$(34)
        ToString = ToString + CStr(Me(i))
        ToString = ToString + Chr$(34)
    ElseIf TypeOf Me(i) Is iArray Then
        ToString = ToString + Me(i).ToString(delimiter)
    ElseIf (VarType(Me(i)) And vbArray) = vbArray Then
        ToString = ToString + "[Array]"
    ElseIf VarType(Me(i)) = vbObject Then
        ToString = ToString + "[Object]"
    Else
        ToString = ToString + CStr(Me(i))
    End If
    
    If i = iArray.Count - 1 Then ToString = ToString + "}"
  Next i
End Function

