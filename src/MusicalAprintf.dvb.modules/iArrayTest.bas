Attribute VB_Name = "iArrayTest"
Option Explicit
Private iArr As iArray
Private iArr1 As iArray
Private iArr2 As iArray
Private cool As New Collection
Public Type miss2
    Name As String * 20
    idx As Integer
End Type
Public Type qq
     aaas As String
End Type

Public Function getMiss() As miss2()
    Dim gg As miss2
    Dim gm(3) As miss2
    gm(0).Name = "33"
    gm(0).idx = 74
    getMiss = gm
End Function
Sub objAdd()
    Dim mm As New Dictionary
    mm("asdfwe") = 4
 
    
End Sub
Public Function Splice( _
        ByVal StartIndex As Long, _
        ParamArray Args() As Variant _
    ) As Variant()
    Dim max As Integer
    Dim min As Integer
    max = UBound(Args)
    min = LBound(Args)
    End Function

Public Function getSlur(ss As String) As vSlur
    Set getSlur = New vSlur
    getSlur.label = ss
End Function
Public Sub sortPitch1()
    Dim sorted As Boolean
    Dim p As Integer
    Dim tmp
    Dim es As New iArray
    es.Push getSlur("6")
    es.Push getSlur("5")
    es.Push getSlur("4")
        For p = 0 To es.Count - 1
            If p + 1 >= es.Count Then
                sorted = True
            ElseIf (es(p).label > es(p + 1).label) Then
                Set tmp = es(p)
                es(p) = es(p + 1)
                es(p + 1) = tmp
            End If
        Next
    
End Sub



Public Sub iArrayVoiceTest()
    
    
    
    Dim vv As New iArray
    Dim v2
    
    vv.Add 0
    vv.Add 1
    vv.Add 2
    vv.Add 3
    vv.Add 4
    vv.Add 5
    vv.Add 6
    vv.Add 7
    vv.Add 8
    vv.Add 9
    vv.Add 0
    v2 = vv.Splice(3, 3, "a", "b")
    
    
    vv.Add "aa", , 3


End Sub
Sub oobjectIS()
    Dim aa As New iArray
    Dim bb(4) As Double
    Dim cc
    Dim dd As New point
    Dim ee As Date
    Dim ff As New point
    Dim gg As Boolean
    
    Dim ans As Boolean
    

    
    cc = Array(3, 4, 5, 23)
    
    Debug.Print "iArray" & vbTab & IsObject(aa)
    Debug.Print "Double()" & vbTab & IsObject(bb)
    Debug.Print "Variant" & vbTab & IsObject(cc)
    Debug.Print "Point " & vbTab & IsObject(dd)
    Debug.Print "Date" & vbTab & IsObject(ee)
    Debug.Print "miss2" & vbTab & IsObject(ff)
    Debug.Print "boolean" & vbTab & IsObject(gg)
    
    Debug.Print "iArray" & vbTab & VarType(aa)
    Debug.Print "Double()" & vbTab & VarType(bb)
    Debug.Print "Variant" & vbTab & VarType(cc)
    Debug.Print "Point" & vbTab & VarType(dd)
    Debug.Print "Date" & vbTab & VarType(ee)
    Debug.Print "miss2" & vbTab & VarType(ff)
    Debug.Print "boolean" & vbTab & VarType(gg)
    
    
End Sub
Public Sub iArrayPointTest()

    Set iArr = New iArray
    Dim pt As New point
    Call pt.c(3, 2)
    iArr.Push pt.Clone
    Call pt.c(13, 12)
    iArr.Push pt.Clone
    Call pt.c(23, 22)
    iArr.Push pt.Clone
    Dim tmp
    Set tmp = iArr.head


End Sub


Public Sub iArrTest2()
    Dim llpt As New PointList
    Dim pt As New point
    Call pt.c(3, 2)
    llpt.Push pt.Clone
    Call pt.c(13, 12)
    llpt.Push pt.Clone
    Call pt.c(23, 22)
    llpt.Push pt.Clone
    
    Dim tttmp
    tttmp = llpt.AddArrayAfter(1, llpt)
End Sub
Sub mae()
    Dim aa()
    Dim ppt As New point
    Dim p1 As New point
    Dim ss
    aa = Array(3, 4, 5)
    ss = TypeName(aa)
    p1.c 3, 4, 5
    
    ppt = p1
    'if TypeOf aa is Array then
    'End If
    
    
End Sub


Public Sub test_iTemplate_test()
    Dim it As iTemplate
    Dim tarr As New iArrayTemplateList
    Dim a2 As New iArrayTemplateList
    Dim aa(3) As Double
    Dim bb(3) As Double
    Dim cc
    
    cc = Array(33, 44, 11)
    If VarType(aa) = VarType(bb) Then
        cc = 33
    End If
    
    
    Set it = New iTemplate
    it.Name = "abc"
    it.typs = 12
    tarr.Push it
    
    Set it = New iTemplate
    it.Name = "Tita"
    it.typs = 33
    tarr.Push it
    
    Set it = New iTemplate
    it.Name = "May"
    it.typs = 23
    tarr.Push it
    
    Set a2 = tarr.head
    
    Set cc = tarr.Pop
    
End Sub


Public Sub test_type_test()
    Dim it
    Dim tarr As New iArray
    Dim a2
    Dim aa
    Dim bb
    
    it = Array(3, 4, 5, 6)
    tarr.Push it
    
    it = Array("_a_", "_b_", "_c_")
    tarr.Push it

    Set a2 = tarr.head
    
End Sub

Public Sub iArrayTest()
'Dim aa As Variant
Dim bb As Variant
'aa = Array(10, 20, 30)
bb = Array("_A_", "_B_", "_C_")
Dim aa As New point
'Dim bb As New point
aa.c 3, 4, 5
'bb.c 3, 4, 5


'ii(0) = 4

  ' ##### PUSH / POP TEST
  Debug.Print vbCrLf & " #### Push/Pop test"
  Set iArr = New iArray
  iArr.PushArray Array(aa, "a", True, 1, aa, "ba", bb)
  Call validate("Push", 4, iArr.Push("Hello world"))
  Debug.Print iArr(1)
  iArr.Pop
  iArr.Pop
  Call validate("Pop", True, iArr.Pop)
    
  ' ##### SHIFT / UNSHIFT TEST
  Debug.Print vbCrLf & " #### Shift/Unshift test"
  Set iArr = New iArray
  iArr.Unshift "..."
  Call validate("Unshift", "{""...""}", iArr.ToString)
  iArr.Unshift 123456
  iArr.UnshiftArray Array(aa, 3.1415, Empty, vbNullString, bb, "a")
  Call validate("UnshiftArray", Array("{3.1415  """" ""a"" 123456 ""...""}", "{3,1415  """" ""a"" 123456 ""...""}"), iArr.ToString)
  iArr.Shift
  iArr.Shift
  iArr.Shift
  Call validate("Shift", "{""a"" 123456 ""...""}", iArr.ToString)

  ' ##### ENQUEUE / DEQUEUE TEST
  Debug.Print vbCrLf & " #### Enqueue/Dequeue test"
  Set iArr = New iArray
  iArr.Enqueue ("Queued element")
  Call validate("Enqueue", "{""Queued element""}", iArr.ToString)
  iArr.EnqueueArray Array(aa, 1, "2", 3.14, False, "Last", bb)
  Call validate("EnqueueArray", Array("{""Queued element"" 1 ""2"" 3.14 False ""Last""}", "{""Queued element"" 1 ""2"" 3,14 False ""Last""}"), iArr.ToString)
  iArr.Dequeue
  iArr.Dequeue
  Call validate("Dequeue", 1, iArr.Dequeue)

  ' ##### DEFAULT MEMBERS TEST
  Debug.Print vbCrLf & " #### Default members test"
  Set iArr = New iArray
  iArr.PushArray Array(aa, "1", 2, "3", 4, bb)
  Call validate("Default Members set", 2, iArr(2))
  iArr(2) = "Two"
  Call validate("Default Members edit", "Two", iArr(2))

  ' ##### CLEAR ARRAY TEST
  Debug.Print vbCrLf & " #### Clear array test"
  Set iArr = New iArray
  iArr.PushArray Array(1, 2, 3, 4, 5, "a", "b", "c", True, Empty)
  iArr.Clear
  Call validate("Clear", "{}", iArr.ToString)

  ' ##### COUNT OCCURRENCES TEST
  Debug.Print vbCrLf & " #### Count occurrences test"
  Set iArr = New iArray
  iArr.PushArray Array(1, 2, 2, 1, 3, 1, 2)
  Call validate("Count occurrences (yes)", 3, iArr.CountOccurrences(2))
  Call validate("Count occurrences (not)", 0, iArr.CountOccurrences(4))

  ' ##### CONTAINS TEST
  Debug.Print vbCrLf & " #### Contains test"
  Set iArr = New iArray
  iArr.PushArray Array(1, 2, 2, 1, 3, 1, 2, aa, bb)
  Call validate("Contains (yes)", True, iArr.Contains(1))
  Call validate("Contains (not)", False, iArr.Contains(5))

  ' ##### CONTAINS ALL TEST
  Debug.Print vbCrLf & " #### Contains All test"
  Set iArr = New iArray
  iArr.PushArray Array(aa, bb, 1, 2, 2, 1, 3, 1, 2, aa, bb)
  Call validate("ContainsAll (yes)", True, iArr.ContainsAll(Array(1, 3)))
  Call validate("ContainsAll (not)", False, iArr.ContainsAll(Array(1, 4, 5)))

  ' ##### FIND DIFFERENCES TEST
  Debug.Print vbCrLf & " #### Difference test"
  Set iArr1 = New iArray
  iArr1.PushArray Array(1, 2, 3, aa, bb)
  Set iArr2 = New iArray
  iArr2.PushArray Array(2, 3, 4, aa, bb)
  Set iArr = iArr2.Difference(iArr1)
  Call validate("Difference", "{1 4}", iArr.ToString)
  Set iArr = iArr2.Difference(iArr1, "d")
  Call validate("Difference (with ""d"" param)", "{1}", iArr.ToString)
  Set iArr = iArr2.Difference(iArr1, "a")
  Call validate("Difference (with ""a"" param)", "{4}", iArr.ToString)

  ' ##### JOINING ARRAYS TEST
  Debug.Print vbCrLf & " #### Joining arrays test"
  Set iArr1 = New iArray
  iArr1.PushArray Array(1, 2, 3, "a", "b", "c", aa, bb)
  Set iArr2 = New iArray
  iArr2.PushArray Array(4, 5, 6, "d", "e", "f", aa, bb)
  Set iArr = iArr1.Join(iArr2)
  Call validate("Join", "{1 2 3 ""a"" ""b"" ""c"" 4 5 6 ""d"" ""e"" ""f""}", iArr.ToString)

  ' ##### DROP LEFT/RIGHT TEST
  Debug.Print vbCrLf & " #### Drop left/right test"
  Set iArr = New iArray
  iArr.PushArray Array(aa, "1", "Two", "3", 4, 1, 2, 3, 4, 5, "a", "b", "c", True, bb)
  Call validate("DropLeft return", "{""1"" ""Two""}", iArr.DropLeft(2).ToString)
  Call validate("DropLeft", "{""3"" 4 1 2 3 4 5 ""a"" ""b"" ""c"" True}", iArr.ToString)
  Call validate("DropRight return", "{""b"" ""c"" True}", iArr.DropRight(3).ToString)
  Call validate("DropRight", "{""3"" 4 1 2 3 4 5 ""a""}", iArr.ToString)

  ' ##### UNIQUE / REMOVE DUPLICATES TEST
  Debug.Print vbCrLf & " #### Unique / Remove duplicates test"
  Set iArr = New iArray
  iArr.PushArray Array("3", 4, 1, 2, 3, 4, 5, "a", "b", "3", "c", "a", True, aa, bb)
  Dim uniqueArr As New iArray
  Set uniqueArr = iArr.Unique
  Call validate("Unique", "{""3"" 4 1 2 3 5 ""a"" ""b"" ""c"" True}", uniqueArr.ToString)
  Call validate("RemoveDuplicates (removed count)", 3, iArr.RemoveDuplicates)
  Call validate("RemoveDuplicates", "{""3"" 4 1 2 3 5 ""a"" ""b"" ""c"" True}", iArr.ToString)

  ' ##### CLONE ARRAY TEST
  Debug.Print vbCrLf & " #### Clone array test"
  Set iArr = New iArray
  iArr.PushArray Array("3", 4, 1, 2, 3, 4, 5, "a", "b", "c", True, aa, bb)
  Dim arrCloned As New iArray
  Set arrCloned = iArr.Clone
  iArr.Clear
  Call validate("Clone", "{""3"" 4 1 2 3 4 5 ""a"" ""b"" ""c"" True [Object] [Array]}", arrCloned.ToString)

  ' ##### SHUFFLE ARRAY TEST
  Debug.Print vbCrLf & " #### Shuffle array test"
  Set iArr = New iArray
  iArr.PushArray Array("3", 4, 1, 2, 3, 4, 5, "a", "b", "c", True, aa, bb)
  Dim arrShufl As New iArray
  Set arrShufl = iArr.Shuffle
  Call validate("Shuffle", True, arrShufl.ToString <> iArr.ToString And arrShufl.ContainsAll(iArr))

  ' ##### REVERSE ARRAY TEST
  Debug.Print vbCrLf & " #### Reverse array test"
  Set iArr = New iArray
  iArr.PushArray Array("3", 4, 1, 2, 3, 4, 5, "a", "b", "c", True)
  Dim arrRev As New iArray
  Set arrRev = iArr.Reverse
  Call validate("Reverse", "{True ""c"" ""b"" ""a"" 5 4 3 2 1 4 ""3""}", arrRev.ToString)

  ' ##### FIRST/LAST TEST
  Debug.Print vbCrLf & " #### First/Last test"
  Set iArr = New iArray
  iArr.PushArray Array(1, 2, 3, 4, 5)
  Call validate("First", 1, iArr.First)
  Call validate("Last", 5, iArr.Last)

  ' ##### ADD AFTER/BEFORE TEST
  Debug.Print vbCrLf & " #### Add after/before test"
  Set iArr = New iArray
  iArr.AddBefore 2, "Something"
  iArr.PushArray Array(1, 2, 3, 4, 5)
  iArr.AddBefore 1, "New First"
  Call validate("AddBefore", "{""New First"" ""Something"" 1 2 3 4 5}", iArr.ToString)
  iArr.AddAfter 4, "Hello"
  Call validate("AddAfter", "{""New First"" ""Something"" 1 2 ""Hello"" 3 4 5}", iArr.ToString)

  ' ##### ADD AFTER/BEFORE ARRAY TEST
  Debug.Print vbCrLf & " #### Add after/before array test"
  Set iArr = New iArray
  iArr.PushArray Array(1, 2, 3, 4, 5)
  iArr.AddArrayAfter 2, Array("a", "b", "c")
  Call validate("AddArrayBefore", "{1 2 ""a"" ""b"" ""c"" 3 4 5}", iArr.ToString)
  iArr.AddArrayBefore 7, Array(True, False)
  Call validate("AddArrayBefore", "{1 2 ""a"" ""b"" ""c"" 3 True False 4 5}", iArr.ToString)
  
  ' ##### TAIL/HEAD TEST
  Debug.Print vbCrLf & " #### Tail / Head test"
  Set iArr = New iArray
  iArr.PushArray Array("3", 4, 1, 2, 3, 4, 5, "a", "b", "c", True)
  Dim tailArr As New iArray
  Set tailArr = iArr.Tail
  Call validate("Tail", "{4 1 2 3 4 5 ""a"" ""b"" ""c"" True}", tailArr.ToString)
  Dim headArr As New iArray
  Set headArr = tailArr.head
  Call validate("Head", "{4 1 2 3 4 5 ""a"" ""b"" ""c""}", headArr.ToString)
  
  ' ##### CONTAINS ONLY NUMERIC TEST
  Debug.Print vbCrLf & " #### Contains Only Numeric test"
  Set iArr = New iArray
  iArr.PushArray Array("3", 4, "3.1415", 2, "1E-2")
  Call validate("ContainsOnlyNumeric", True, iArr.ContainsOnlyNumeric)
  Set iArr = New iArray
  iArr.PushArray Array("3a", 4, 3.1415, 2, "1E-2")
  Call validate("ContainsOnlyNumeric", False, iArr.ContainsOnlyNumeric)
  
  ' ##### AVERAGE / SUM TEST
  Debug.Print vbCrLf & " #### Average / Sum test"
  Set iArr = New iArray
  iArr.PushArray Array("3a", 4, "3.1415", 2, "1E-2")
  Call validate("Sum", "NaN", iArr.sum)
  Set iArr = New iArray
  iArr.PushArray Array("3", 4, 3.1415, 2, "1E-2")
  Call validate("Sum", 12.1515, iArr.sum)
  Call validate("Average", 2.4303, iArr.Avg)
  
  ' ##### OCCURRENCE INDEXES TEST
  Debug.Print vbCrLf & " #### Occurence Indexes test"
  Set iArr = New iArray
  iArr.PushArray Array(1, 2, True, "Abc", 2, "1", 3, 1, 2)
  Call validate("Occurrence indexes", "{1 8}", iArr.OccurenceIndexes(1).ToString)
  
  ' ##### INTERSECT ARRAYS TEST
  Debug.Print vbCrLf & " #### Intersect arrays test"
  Set iArr1 = New iArray
  iArr1.PushArray Array(1, 2, 3, "a", "b", "a")
  Set iArr2 = New iArray
  iArr2.PushArray Array(3, 2, 6, "a", "a", "f")
  Set iArr = iArr1.Intersect(iArr2)
  Call validate("Intersect", "{2 3 ""a""}", iArr.ToString)
  
  ' ##### UNION ARRAYS TEST
  Debug.Print vbCrLf & " #### Union arrays test"
  Set iArr1 = New iArray
  iArr1.PushArray Array(1, 2, 3, "a", "b", "a")
  Set iArr2 = New iArray
  iArr2.PushArray Array(3, 2, 6, "a", "a", "f")
  Set iArr = iArr1.Union(iArr2)
  Call validate("Union", "{1 2 3 ""a"" ""b"" 6 ""f""}", iArr.ToString)
End Sub

Private Sub validate(Name As String, Expected, Actual As String)
  If Not IsArray(Expected) Then Expected = Array(Expected)
  
  Dim found As Boolean: found = False
  Dim possibleResult As String
  
  Dim i As Integer
  For i = LBound(Expected) To UBound(Expected)
    If Expected(i) = Actual Then found = True: Exit For
  Next i
  
  If found Then
    Debug.Print Name + " - OK"
  Else
    Debug.Print Name + " - NOK"
    Debug.Print " - Actual value: " + Actual
    
    Dim expectedString As Variant
    For i = LBound(Expected) To UBound(Expected)
      If i > LBound(Expected) Then expectedString = expectedString + " or "
      expectedString = expectedString + Expected(i)
    Next i
    Debug.Print " - Expected value: " + CStr(expectedString)
  End If
End Sub

Public Function ToString(self, Optional ByVal Delimiter As String = " ") As String
    Dim j As Long
     If (VarType(self) And vbArray) = vbArray Then
     ElseIf TypeOf self Is iArray Then
     
     End If
  If Me.Count = 0 Then ToString = "{}": Exit Function
  ToString = vbNullString
  For i = 0 To Me.Count - 1
    If i = 0 Then ToString = ToString + "{"
    If i > 0 Then ToString = ToString + Delimiter
    If VarType(Me(i)) = vbString Then
        ToString = ToString + chr$(34)
        ToString = ToString + CStr(Me(i))
        ToString = ToString + chr$(34)
    ElseIf TypeOf Me(i) Is iArray Then
        ToString = ToString + Me(i).ToString(Delimiter)
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
