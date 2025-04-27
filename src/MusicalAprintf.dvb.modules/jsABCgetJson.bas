Attribute VB_Name = "jsABCgetJson"
Option Explicit

Public ts As New iArray
Public Function getAbcJson() As iArray
    ''把 abc2svg 的 TS 資料 轉換成  VBA 物件 =>ts
    Dim prod As Variant
    Baidu_appition
    Dim cdtxt As String
    Dim tsCount As Integer
    Dim childCount As Integer
    Dim i As Integer
    Dim j As Integer
    Dim childItem As voiceItem
    Dim childList As iArray
    tsCount = WD.ExecuteScript("return ts.length")
    ts.Clear
    For i = 0 To tsCount - 1
        If i >= ts.Count - 1 Then
            ts.Push New iArray
        End If
        Set childList = ts(i)
        
        childCount = WD.ExecuteScript("return ts[" & i & "].length")
        For j = 0 To childCount - 1
            cdtxt = "ts[" & i & "][" & j & "]"
            prod = WD.ExecuteScript("return JSON.stringify(" & cdtxt & ", JsonReplacerABC2SVG, 2)")
            'Debug.Print prod
            Set childItem = childGetTsItem(prod)
            childItem.fname = "ts[" & i & "][" & j & "]"
            childList.Push childItem
        Next
    Next
    
 Set getAbcJson = ts
End Function

Function childGetTsItem(prod) As voiceItem
    On Error GoTo ErrorHandler
    Dim JsonObject
    Dim Json As Dictionary
    Dim key
    Dim abcItem As New iArray
    Dim getArrList As iArray
    Dim vItem As voiceItem
    
    Set vItem = New voiceItem
    Set Json = JsonConverter.ParseJson(prod)
    For Each key In Json.keys
        Select Case key    ' Evaluate Number.
            Case "notes":
                Set getArrList = objArrayItem(Json.item(key), AbcObject.noteItem)
                'Call vItem.ConvertVarName(key, Json.Item(key))
                CallByName vItem, key, VbSet, getArrList
            Case "slurStart":
                Set getArrList = objArrayItem(Json.item(key), AbcObject.Number)
                'Call vItem.ConvertVarName(key, Json.Item(key))
                CallByName vItem, key, VbSet, getArrList
            Case "slurEnd":
                Set getArrList = objArrayItem(Json.item(key), AbcObject.Number)
                'Call vItem.ConvertVarName(key, Json.Item(key))
                CallByName vItem, key, VbSet, getArrList
            Case "k_map", "a_gch":
                Set getArrList = objArrayItem(Json.item(key), AbcObject.KMapItem)
                'Call vItem.ConvertVarName(key, Json.Item(key))
                CallByName vItem, key, VbSet, getArrList
            Case "a_meter":
            Case "x_meter":
            Case "tempo_notes":
            Case "tempo_wh":
            Case "slur":
            Case "pos":
            Case "sy":
            Case "type":
                CallByName vItem, "typs", VbLet, Json.item(key)
            Case "next":
                CallByName vItem, "nexs", VbLet, Json.item(key)
            Case "__TSpos":
                CallByName vItem, "TS_pos", VbLet, Json.item(key)
            Case Else
                CallByName vItem, key, VbLet, Json.item(key)
                'Call vItem.ConvertVarName(key, Json.Item(key))
callBack1:
        End Select
    Next
    Set childGetTsItem = vItem
    Exit Function
ErrorHandler:
    If VarType(Json.item(key)) = vbObject Then
        Debug.Print "key-> " & key & " value-> [object]"
    Else
        Debug.Print "key-> " & key & " value-> " & Json.item(key)
    End If
GoTo callBack1

End Function
Function objArrayItem(ListObject, ByVal objectEnum As AbcObject) As iArray
    On Error GoTo ErrorHandler
        Dim ky1, lsNote, o, o2, ks
        Dim i, j
        Dim ListNote
        Dim aObj
        Dim aObjArr As New iArray
        Dim retArrList
        If objectEnum = Number Or objectEnum = Booling Then
            For Each o In ListObject
                aObjArr.Push o
            Next
        Else
            For Each o In ListObject
                Set aObj = newObjectItem(objectEnum)
                For Each ks In o.keys
                    CallByName aObj, ks, VbLet, o.item(ks)
                    Select Case key    ' Evaluate Number.
                        Case "font":
                            Set retArrList = objArrayItem(o.item(key), AbcObject.FontItem)
                            'Call vItem.ConvertVarName(key, Json.Item(key))
                            CallByName aObj, ks, VbSet, retArrList
                      
                        Case Else
                            CallByName aObj, ks, VbLet, o.item(ks)

                     End Select
                    
                    'Call aNote.ConvertVarName(o2, lsNote.Item(o2))
callback2:
                Next
                Call aObjArr.Push(aObj)
            Next
        End If
        Set objArrayItem = aObjArr
        Exit Function
ErrorHandler:
    Debug.Print "objArrayItemkey=> "& "[" &   &"]" & ks & "  value-> " & o.item(ks)
GoTo callback2
        ' Resume execution at same line
        ' that caused the error.

End Function
Function newObjectItem(objectName As AbcObject)
    '創建任合物件
        Select Case objectName
            ''一般
            Case AbcObject.noteItem:        Set newObjectItem = New aNoteItem
            Case AbcObject.DecorationItme:  Set newObjectItem = New aDecorationItem
            Case AbcObject.FontItem:        Set newObjectItem = New aFontItem
            Case AbcObject.GchordItem:      Set newObjectItem = New aGchordItem
            Case AbcObject.LyrcsItem:       Set newObjectItem = New aLyricsItem
            Case AbcObject.VoiecItem:       Set newObjectItem = New voiceItem
            Case AbcObject.KMapItem:        Set newObjectItem = New aKMap
            Case Else
                Debug.Print "沒有變數 " & key & "名"
        End Select
End Function
Function testJSON()
    Dim JsonString As String
    Dim JsonObject
    Dim Json As Dictionary
    Set Json = JsonConverter.ParseJson("{""a"":123,""b"":[1,2,3,4],""c"":{""d"":456}}")
    
    ' Json("a") -> 123
    ' Json("b")(2) -> 2
    ' Json("c")("d") -> 456
    Json("c")("e") = 789
    
    Debug.Print JsonConverter.ConvertToJson(Json)
    ' -> "{"a":123,"b":[1,2,3,4],"c":{"d":456,"e":789}}"
    
    Debug.Print JsonConverter.ConvertToJson(Json, Whitespace:=2)
    Debug.Print Json.Count

End Function

Function voieMusic()
    Dim vvbase As New voiceBase
    Debug.Print vvbase.typs
    vvbase.typs = 3
    Debug.Print vvbase.typs
    
    Dim vvbar As New voiceBar
    Debug.Print vvbar.typs
End Function
