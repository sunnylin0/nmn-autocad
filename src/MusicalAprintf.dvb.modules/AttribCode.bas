Attribute VB_Name = "AttribCode"
Option Explicit
Function ACAD_Ver() As Integer
'�Ǧ^AutoCAD ���������X
    ' This example returns AutoCAD version as a string
    
    Dim VERSION As String
    VERSION = ThisDrawing.Application.VERSION
    ACAD_Ver = Val(VERSION)
    
End Function
Function GetAttribTextString(entAttrib As AcadEntity, title As String) As String
'�� �ݩ�(titile)����r
    Dim Array1 As Variant
    Dim Pnt As Variant
    'Dim entObj As AcadEntity
    Dim count As Integer
        With entAttrib
            '��@�Ӷ��ޥΪ��Q����A�ˬd���O�_���ݩ�
            If StrComp(.EntityName, "AcDbBlockReference", 1) = 0 Then
                '�p�G���ݩ�
                If .HasAttributes = True Then
                    '�������ޥΤ����ݩ�
                    Array1 = .GetAttributes
                    '�o�@���j��ΨӬd����D�A�p�G����b��1��
                    For count = LBound(Array1) To UBound(Array1)
                        '�p�G�٨S�����D
                        If Array1(count).TagString = title Then
                            GetAttribTextString = Array1(count).textString
                            Exit Function
                        End If
                    Next count
                    'MsgBox """" & title & """�ݩʥ����"
                End If
            End If
        End With
End Function

Function SetAttribTextString(entAttrib As AcadEntity, title As String, sVal As String) As String
'�]�w �ݩ�(titile)����r
    Dim Array1 As Variant
    Dim Pnt As Variant
    'Dim entObj As AcadEntity
    Dim count As Integer
        With entAttrib
            '��@�Ӷ��ޥΪ��Q����A�ˬd���O�_���ݩ�
            If StrComp(.EntityName, "AcDbBlockReference", 1) = 0 Then
                '�p�G���ݩ�
                If .HasAttributes = True Then
                    '�������ޥΤ����ݩ�
                    Array1 = .GetAttributes
                    '�o�@���j��ΨӬd����D�A�p�G����b��1��
                    For count = LBound(Array1) To UBound(Array1)
                        '�p�G�٨S�����D
                        If Array1(count).TagString = title Then
                            Array1(count).textString = sVal
                            Exit Function
                        End If
                    Next count
                    'MsgBox """" & title & """�ݩʥ����"
                    SetAttribTextString = ""
                End If
            Else
                'MsgBox "�o���O�ݩʹϤ�"
                SetAttribTextString = ""
            End If
        End With
End Function

Function SetTowAttrib(entAttrib As AcadEntity, towAttrib As AcadEntity) As String
'�]�w �ݩ�(titile)����r
    Dim Array1 As Variant
    Dim Pnt As Variant
    Dim title As String
    Dim sVal As String
    'Dim entObj As AcadEntity
    Dim count As Integer
        With entAttrib
            '��@�Ӷ��ޥΪ��Q����A�ˬd���O�_���ݩ�
            If StrComp(.EntityName, "AcDbBlockReference", 1) = 0 Then
                '�p�G���ݩ�
                If .HasAttributes = True Then
                    '�������ޥΤ����ݩ�
                    Array1 = .GetAttributes
                    '�o�@���j��ΨӬd����D�A�p�G����b��1��
                    For count = LBound(Array1) To UBound(Array1)
                       
                        title = Array1(count).TagString
                        sVal = Array1(count).textString
                        
                        Call SetAttribTextString(towAttrib, title, sVal)
                        
                    Next count
                    
                    
                End If
            Else
                'MsgBox "�o���O�ݩʹϤ�"
            End If
        End With
End Function

Function GetAttribList(entAttrib As AcadEntity) As String
'���o �ݩʦC��
    Dim Array1 As Variant
    Dim Pnt As Variant
    Dim strList As String
    Dim count As Integer
    strList = ""
        With entAttrib
            '��@�Ӷ��ޥΪ��Q����A�ˬd���O�_���ݩ�
            If StrComp(.EntityName, "AcDbBlockReference", 1) = 0 Then
                '�p�G���ݩ�
                If .HasAttributes = True Then
                    '�������ޥΤ����ݩ�
                    Array1 = .GetAttributes
                    '�o�@���j��ΨӬd����D�A�p�G����b��1��
                    For count = LBound(Array1) To UBound(Array1)
                        strList = strList & Array1(count).TagString & vbCrLf
                    Next count
                    GetAttribList = strList
                    Exit Function
                End If
            End If
        End With
End Function












