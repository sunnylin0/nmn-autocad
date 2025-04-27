Attribute VB_Name = "AttribCode"
Option Explicit

Function ACAD_Ver() As Integer
'傳回AutoCAD 的版本號碼
    ' This example returns AutoCAD version as a string
    
    Dim version As String
    version = ThisDrawing.Application.version
    ACAD_Ver = val(version)
    
End Function
Function GetAttribTextString(entAttrib As AcadEntity, title As String) As String
'找 屬性(titile)的文字
    Dim Array1 As Variant
    Dim Pnt As Variant
    'Dim entObj As AcadEntity
    Dim Count As Integer
        With entAttrib
            '當一個塊引用表行被找到後，檢查它是否有屬性
            If StrComp(.EntityName, "AcDbBlockReference", 1) = 0 Then
                '如果有屬性
                If .HasAttributes = True Then
                    '提取塊引用中的屬性
                    Array1 = .GetAttributes
                    '這一輪迴圈用來查找標題，如果有填在第1行
                    For Count = LBound(Array1) To UBound(Array1)
                        '如果還沒有標題
                        If Array1(Count).TagString = title Then
                            GetAttribTextString = Array1(Count).textString
                            Exit Function
                        End If
                    Next Count
                    'MsgBox """" & title & """屬性未找到"
                End If
            End If
        End With
End Function

Function SetAttribTextString(entAttrib As AcadEntity, title As String, sval As String) As String
'設定 屬性(titile)的文字
    Dim Array1 As Variant
    Dim Pnt As Variant
    'Dim entObj As AcadEntity
    Dim Count As Integer
        With entAttrib
            '當一個塊引用表行被找到後，檢查它是否有屬性
            If StrComp(.EntityName, "AcDbBlockReference", 1) = 0 Then
                '如果有屬性
                If .HasAttributes = True Then
                    '提取塊引用中的屬性
                    Array1 = .GetAttributes
                    '這一輪迴圈用來查找標題，如果有填在第1行
                    For Count = LBound(Array1) To UBound(Array1)
                        '如果還沒有標題
                        If Array1(Count).TagString = title Then
                            Array1(Count).textString = sval
                            Exit Function
                        End If
                    Next Count
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
    Dim sval As String
    'Dim entObj As AcadEntity
    Dim Count As Integer
        With entAttrib
            '當一個塊引用表行被找到後，檢查它是否有屬性
            If StrComp(.EntityName, "AcDbBlockReference", 1) = 0 Then
                '如果有屬性
                If .HasAttributes = True Then
                    '提取塊引用中的屬性
                    Array1 = .GetAttributes
                    '這一輪迴圈用來查找標題，如果有填在第1行
                    For Count = LBound(Array1) To UBound(Array1)
                       
                        title = Array1(Count).TagString
                        sval = Array1(Count).textString
                        
                        Call SetAttribTextString(towAttrib, title, sval)
                        
                    Next Count
                    
                    
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
    Dim Count As Integer
    strList = ""
        With entAttrib
            '當一個塊引用表行被找到後，檢查它是否有屬性
            If StrComp(.EntityName, "AcDbBlockReference", 1) = 0 Then
                '如果有屬性
                If .HasAttributes = True Then
                    '提取塊引用中的屬性
                    Array1 = .GetAttributes
                    '這一輪迴圈用來查找標題，如果有填在第1行
                    For Count = LBound(Array1) To UBound(Array1)
                        strList = strList & Array1(Count).TagString & vbCrLf
                    Next Count
                    GetAttribList = strList
                    Exit Function
                End If
            End If
        End With
End Function












