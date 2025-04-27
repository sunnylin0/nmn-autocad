VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dictionary_xx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' Dictionary v1.2.0
' (c) Tim Hall - https:''github.com/timhall/VBA-Dictionary
'
' Drop-in replacement for Scripting.Dictionary on Mac
'
' @author: tim.hall.engr@gmail.com
' @license: MIT (http:''www.opensource.org/licenses/mit-license.php
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

#Const UseScriptingDictionaryIfAvailable = True

#If Mac Or Not UseScriptingDictionaryIfAvailable Then

' KeyValue 0: FormattedKey, 1: OriginalKey, 2: Value
Private pKeyValues As Collection
Private pKeys() As Variant
Private pItems() As Variant
Private pCompareMode As CompareMethod

#Else

Private pDictionary As Object

#End If


Public LabelPos As Integer

' --------------------------------------------- '
' Types
' --------------------------------------------- '

Public Enum CompareMethod
    BinaryCompare = vbBinaryCompare
    TextCompare = vbTextCompare
    DatabaseCompare = vbDatabaseCompare
End Enum

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public Property Get CompareMode() As CompareMethod
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    CompareMode = pCompareMode
#Else
    CompareMode = pDictionary.CompareMode
#End If
End Property
Public Property Let CompareMode(value As CompareMethod)
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    If Me.Count > 0 Then
        ' Can't change CompareMode for Dictionary that contains data
        ' http:''msdn.microsoft.com/en-us/library/office/gg278481(v=office.15).aspx
        err.Raise 5 ' Invalid procedure call or argument
    End If
    
    pCompareMode = value
#Else
    pDictionary.CompareMode = value
#End If
End Property

Public Property Get Count() As Long
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    Count = pKeyValues.Count
#Else
    Count = pDictionary.Count
#End If
End Property

Public Property Get item(key As Variant) As Variant
Attribute item.VB_UserMemId = 0
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    Dim KeyValue As Variant
    KeyValue = GetKeyValue(key)
    
    If Not IsEmpty(KeyValue) Then
        If IsObject(KeyValue(2)) Then
            Set item = KeyValue(2)
        Else
            item = KeyValue(2)
        End If
    Else
        ' Not found -> Returns Empty
    End If
#Else
    If IsObject(pDictionary.item(key)) Then
        Set item = pDictionary.item(key)
    Else
        item = pDictionary.item(key)
    End If
#End If
End Property
Public Property Let item(key As Variant, value As Variant)
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    If Me.Exists(key) Then
        ReplaceKeyValue GetKeyValue(key), key, value
    Else
        AddKeyValue key, value
    End If
#Else
    pDictionary.item(key) = value
#End If
End Property
Public Property Set item(key As Variant, value As Variant)
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    If Me.Exists(key) Then
        ReplaceKeyValue GetKeyValue(key), key, value
    Else
        AddKeyValue key, value
    End If
#Else
    Set pDictionary.item(key) = value
#End If
End Property

Public Property Let key(Previous As Variant, Updated As Variant)
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    Dim KeyValue As Variant
    KeyValue = GetKeyValue(Previous)
    
    If Not IsEmpty(KeyValue) Then
        ReplaceKeyValue KeyValue, Updated, KeyValue(2)
    End If
#Else
    pDictionary.key(Previous) = Updated
#End If
End Property

' ============================================= '
' Public Methods
' ============================================= '

''
' Add an item with the given key
'
' @param {Variant} Key
' @param {Variant} Item
' --------------------------------------------- '
Public Sub Add(key As Variant, item As Variant)
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    If Not Me.Exists(key) Then
        AddKeyValue key, item
    Else
        ' This key is already associated with an element of this collection
        err.Raise 457
    End If
#Else
    pDictionary.Add key, item
#End If
End Sub

''
' Check if an item exists for the given key
'
' @param {Variant} Key
' @return {Boolean}
' --------------------------------------------- '
Public Function Exists(key As Variant) As Boolean
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    Exists = Not IsEmpty(GetKeyValue(key))
#Else
    Exists = pDictionary.Exists(key)
#End If
End Function

''
' Get an array of all items
'
' @return {Variant}
' --------------------------------------------- '
Public Function Items() As Variant
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    If Me.Count > 0 Then
        Items = pItems
    Else
        ' Split("") creates initialized empty array that matches Dictionary Keys and Items
        Items = Split("")
    End If
#Else
    Items = pDictionary.Items
#End If
End Function

''
' Get an array of all keys
'
' @return {Variant}
' --------------------------------------------- '
Public Function keys() As Variant
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    If Me.Count > 0 Then
        keys = pKeys
    Else
        ' Split("") creates initialized empty array that matches Dictionary Keys and Items
        keys = Split("")
    End If
#Else
    keys = pDictionary.keys
#End If
End Function

''
' Remove an item for the given key
'
' @param {Variant} Key
' --------------------------------------------- '
Public Sub Remove(key As Variant)
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    Dim KeyValue As Variant
    KeyValue = GetKeyValue(key)
    
    If Not IsEmpty(KeyValue) Then
        RemoveKeyValue KeyValue
    Else
        ' Application-defined or object-defined error
        err.Raise 32811
    End If
#Else
    pDictionary.Remove key
#End If
End Sub

''
' Remove all items
' --------------------------------------------- '
Public Sub RemoveAll()
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    Set pKeyValues = New Collection
    
    Erase pKeys
    Erase pItems
#Else
    pDictionary.RemoveAll
#End If
End Sub

' ============================================= '
' Private Functions
' ============================================= '

#If Mac Or Not UseScriptingDictionaryIfAvailable Then

Private Function GetKeyValue(key As Variant) As Variant
    On Error Resume Next
    GetKeyValue = pKeyValues(GetFormattedKey(key))
    err.Clear
End Function

Private Sub AddKeyValue(key As Variant, value As Variant, Optional index As Long = -1)
    If Me.Count = 0 Then
        ReDim pKeys(0 To 0)
        ReDim pItems(0 To 0)
    Else
        ReDim Preserve pKeys(0 To UBound(pKeys) + 1)
        ReDim Preserve pItems(0 To UBound(pItems) + 1)
    End If
    
    Dim FormattedKey As String
    FormattedKey = GetFormattedKey(key)
    
    If index > 0 And index <= pKeyValues.Count Then
        Dim i As Long
        For i = UBound(pKeys) To index Step -1
            pKeys(i) = pKeys(i - 1)
            If IsObject(pItems(i - 1)) Then
                Set pItems(i) = pItems(i - 1)
            Else
                pItems(i) = pItems(i - 1)
            End If
        Next i
        
        pKeys(index - 1) = key
        If IsObject(value) Then
            Set pItems(index - 1) = value
        Else
            pItems(index - 1) = value
        End If
        
        pKeyValues.Add Array(FormattedKey, key, value), FormattedKey, before:=index
    Else
        pKeys(UBound(pKeys)) = key
        If IsObject(value) Then
            Set pItems(UBound(pItems)) = value
        Else
            pItems(UBound(pItems)) = value
        End If
        
        pKeyValues.Add Array(FormattedKey, key, value), FormattedKey
    End If
End Sub

Private Sub ReplaceKeyValue(KeyValue As Variant, key As Variant, value As Variant)
    Dim index As Long
    Dim i As Integer
    
    For i = 0 To UBound(pKeys)
        If pKeys(i) = KeyValue(1) Then
            index = i + 1
            Exit For
        End If
    Next i
    
    ' Remove existing value
    RemoveKeyValue KeyValue, index
    
    ' Add new key value back
    AddKeyValue key, value, index
End Sub

Private Sub RemoveKeyValue(KeyValue As Variant, Optional ByVal index As Long = -1)
    Dim i As Long
    If index = -1 Then
        For i = 0 To UBound(pKeys)
            If pKeys(i) = KeyValue(1) Then
                index = i
            End If
        Next i
    Else
        index = index - 1
    End If
    
    If index >= 0 And index <= UBound(pKeys) Then
        For i = index To UBound(pKeys) - 1
            pKeys(i) = pKeys(i + 1)
            
            If IsObject(pItems(i + 1)) Then
                Set pItems(i) = pItems(i + 1)
            Else
                pItems(i) = pItems(i + 1)
            End If
        Next i
        
        If UBound(pKeys) = 0 Then
            Erase pKeys
            Erase pItems
        Else
            ReDim Preserve pKeys(0 To UBound(pKeys) - 1)
            ReDim Preserve pItems(0 To UBound(pItems) - 1)
        End If
    End If
    
    pKeyValues.Remove KeyValue(0)
End Sub

Private Function GetFormattedKey(key As Variant) As String
    GetFormattedKey = CStr(key)
    If Me.CompareMode = CompareMethod.BinaryCompare Then
        ' Collection does not have method of setting key comparison
        ' So case-sensitive keys aren't supported by default
        ' -> Approach: Append lowercase characters to original key
        '    AbC -> AbC__b, abc -> abc__abc, ABC -> ABC
        '    Won't work in very strange cases, but should work for now
        '    AbBb -> AbBb__bb matches AbbB -> AbbB__bb
        Dim Lowercase As String
        Lowercase = ""
        
        Dim i As Integer
        Dim Ascii As Integer
        Dim Char As String
        For i = 1 To Len(GetFormattedKey)
            Char = VBA.Mid$(GetFormattedKey, i, 1)
            Ascii = Asc(Char)
            If Ascii >= 97 And Ascii <= 122 Then
                Lowercase = Lowercase & Char
            End If
        Next i
        
        If Lowercase <> "" Then
            GetFormattedKey = GetFormattedKey & "__" & Lowercase
        End If
    End If
End Function

#End If


Private Sub Class_Initialize()

    
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    Set pKeyValues = New Collection
    
    Erase pKeys
    Erase pItems
#Else
    Set pDictionary = CreateObject("Scripting.Dictionary")
#End If
End Sub

Private Sub Class_Terminate()

#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    Set pKeyValues = Nothing
#Else
    Set pDictionary = Nothing
#End If
End Sub
