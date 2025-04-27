Attribute VB_Name = "AttribCode"
Option Explicit
Function ACAD_Ver() As Integer
'傳回AutoCAD 的版本號碼
    ' This example returns AutoCAD version as a string
    
    Dim VERSION As String
    VERSION = ThisDrawing.Application.VERSION
    ACAD_Ver = Val(VERSION)
    
End Function
Function GetAttribTextString(entAttrib As AcadEntity, title As String) As String
'找 屬性(titile)的文字
    Dim Array1 As Variant
    Dim Pnt As Variant
    'Dim entObj As AcadEntity
    Dim count As Integer
        With entAttrib
            '當一個塊引用表行被找到後，檢查它是否有屬性
            If StrComp(.EntityName, "AcDbBlockReference", 1) = 0 Then
                '如果有屬性
                If .HasAttributes = True Then
                    '提取塊引用中的屬性
                    Array1 = .GetAttributes
                    '這一輪迴圈用來查找標題，如果有填在第1行
                    For count = LBound(Array1) To UBound(Array1)
                        '如果還沒有標題
                        If Array1(count).TagString = title Then
                            GetAttribTextString = Array1(count).textString
                            Exit Function
                        End If
                    Next count
                    'MsgBox """" & title & """屬性未找到"
                End If
            End If
        End With
End Function

Function SetAttribTextString(entAttrib As AcadEntity, title As String, sVal As String) As String
'設定 屬性(titile)的文字
    Dim Array1 As Variant
    Dim Pnt As Variant
    'Dim entObj As AcadEntity
    Dim count As Integer
        With entAttrib
            '當一個塊引用表行被找到後，檢查它是否有屬性
            If StrComp(.EntityName, "AcDbBlockReference", 1) = 0 Then
                '如果有屬性
                If .HasAttributes = True Then
                    '提取塊引用中的屬性
                    Array1 = .GetAttributes
                    '這一輪迴圈用來查找標題，如果有填在第1行
                    For count = LBound(Array1) To UBound(Array1)
                        '如果還沒有標題
                        If Array1(count).TagString = title Then
                            Array1(count).textString = sVal
                            Exit Function
                        End If
                    Next count
                    'MsgBox """" & title & """屬性未找到"
                    SetAttribTextString = ""
                End If
            Else
                'MsgBox "這不是屬性圖元"
                SetAttribTextString = ""
            End If
        End With
End Function

Function SetTowAttrib(entAttrib As AcadEntity, towAttrib As AcadEntity) As String
'設定 屬性(titile)的文字
    Dim Array1 As Variant
    Dim Pnt As Variant
    Dim title As String
    Dim sVal As String
    'Dim entObj As AcadEntity
    Dim count As Integer
        With entAttrib
            '當一個塊引用表行被找到後，檢查它是否有屬性
            If StrComp(.EntityName, "AcDbBlockReference", 1) = 0 Then
                '如果有屬性
                If .HasAttributes = True Then
                    '提取塊引用中的屬性
                    Array1 = .GetAttributes
                    '這一輪迴圈用來查找標題，如果有填在第1行
                    For count = LBound(Array1) To UBound(Array1)
                       
                        title = Array1(count).TagString
                        sVal = Array1(count).textString
                        
                        Call SetAttribTextString(towAttrib, title, sVal)
                        
                    Next count
                    
                    
                End If
            Else
                'MsgBox "這不是屬性圖元"
            End If
        End With
End Function

Function GetAttribList(entAttrib As AcadEntity) As String
'取得 屬性列表
    Dim Array1 As Variant
    Dim Pnt As Variant
    Dim strList As String
    Dim count As Integer
    strList = ""
        With entAttrib
            '當一個塊引用表行被找到後，檢查它是否有屬性
            If StrComp(.EntityName, "AcDbBlockReference", 1) = 0 Then
                '如果有屬性
                If .HasAttributes = True Then
                    '提取塊引用中的屬性
                    Array1 = .GetAttributes
                    '這一輪迴圈用來查找標題，如果有填在第1行
                    For count = LBound(Array1) To UBound(Array1)
                        strList = strList & Array1(count).TagString & vbCrLf
                    Next count
                    GetAttribList = strList
                    Exit Function
                End If
            End If
        End With
End Function












