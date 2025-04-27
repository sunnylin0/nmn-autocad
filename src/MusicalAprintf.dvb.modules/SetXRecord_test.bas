Attribute VB_Name = "SetXRecord_test"
Option Explicit
Sub testt()
'ด๚ธี setAddXrecord
    Dim arr1(2) As Variant
    Dim a As Variant
    SetXRecord "tlscad", "A", Array(1, 2, "A")
    SetXRecord "tlscad", "B", Array(3, 4, "B")
    a = GetXRecord("tlscad", "A")
    If Not IsNull(a) Then MsgBox a(2)
End Sub


Public Function SetXRecord(ByVal DictName As String, ByVal Keyword As String, ByVal XRecordData)
    Dim pDict As AcadDictionary
    Dim pXRecord As AcadXRecord
    Dim XRecordType() As Integer
    Dim pLen As Integer
    Dim i As Integer
    
    Set pDict = ThisDrawing.Dictionaries.Add(DictName)
    Set pXRecord = pDict.AddXRecord(Keyword)
    
    pLen = UBound(XRecordData)
    ReDim XRecordType(pLen) As Integer
    For i = 0 To pLen
        Select Case VarType(XRecordData(i))
            Case vbInteger, vbLong
                XRecordType(i) = 70
            Case vbSingle, vbDouble
                XRecordType(i) = 40
            Case vbString
                XRecordType(i) = 1
        End Select
    Next i
    
    pXRecord.SetXRecordData XRecordType, XRecordData
End Function
Public Function GetXRecord(ByVal DictName As String, ByVal Keyword As String)
On Error GoTo ErrHandle
    Dim pDict As AcadDictionary
    Dim pXRecord As AcadXRecord
    Dim xt
    Set pDict = ThisDrawing.Dictionaries(DictName)
    Set pXRecord = pDict.GetObject(Keyword)
    pXRecord.GetXRecordData xt, GetXRecord
    Exit Function
ErrHandle:
    GetXRecord = Null
End Function
 
Public Function CreateArray(ByVal TypeName As VbVarType, ParamArray ValArray())
    Dim nCount As Integer
    Dim i
    Dim mArray
    
    nCount = UBound(ValArray)
    
    Select Case TypeName
    Case vbDouble
        Dim dArray() As Double
        ReDim dArray(nCount)
        For i = 0 To nCount
            dArray(i) = ValArray(i)
        Next i
        CreateArray = dArray
    Case vbInteger
        Dim nArray() As Integer
        ReDim nArray(nCount)
        For i = 0 To nCount
            nArray(i) = ValArray(i)
        Next i
        CreateArray = nArray
    Case vbString
        Dim sArray() As String
        ReDim sArray(nCount)
        For i = 0 To nCount
            sArray(i) = ValArray(i)
        Next i
        CreateArray = sArray
    Case vbVariant
        Dim vArray()
        ReDim vArray(nCount)
        For i = 0 To nCount
            vArray(i) = ValArray(i)
        Next i
        CreateArray = vArray
    Case vbObject
        Dim oArray() As Object
        ReDim oArray(nCount)
        For i = 0 To nCount
            Set oArray(i) = ValArray(i)
        Next i
        CreateArray = oArray
    End Select
End Function
