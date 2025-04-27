Attribute VB_Name = "AttribCode"
Option Explicit

Function ACAD_Ver() As Integer
'�Ǧ^AutoCAD ���������X
    ' This example returns AutoCAD version as a string
    
    Dim version As String
    version = ThisDrawing.Application.version
    ACAD_Ver = val(version)
    
End Function
Function GetAttribTextString(entAttrib As AcadEntity, title As String) As String
'�� �ݩ�(titile)����r
    Dim Array1 As Variant
    Dim Pnt As Variant
    'Dim entObj As AcadEntity
    Dim Count As Integer
        With entAttrib
            '��@�Ӷ��ޥΪ��Q����A�ˬd���O�_���ݩ�
            If StrComp(.EntityName, "AcDbBlockReference", 1) = 0 Then
                '�p�G���ݩ�
                If .HasAttributes = True Then
                    '�������ޥΤ����ݩ�
                    Array1 = .GetAttributes
                    '�o�@���j��ΨӬd����D�A�p�G����b��1��
                    For Count = LBound(Array1) To UBound(Array1)
                        '�p�G�٨S�����D
                        If Array1(Count).TagString = title Then
                            GetAttribTextString = Array1(Count).textString
                            Exit Function
                        End If
                    Next Count
                    'MsgBox """" & title & """�ݩʥ����"
                End If
            End If
        End With
End Function

Function SetAttribTextString(entAttrib As AcadEntity, title As String, sval As String) As String
'�]�w �ݩ�(titile)����r
    Dim Array1 As Variant
    Dim Pnt As Variant
    'Dim entObj As AcadEntity
    Dim Count As Integer
        With entAttrib
            '��@�Ӷ��ޥΪ��Q����A�ˬd���O�_���ݩ�
            If StrComp(.EntityName, "AcDbBlockReference", 1) = 0 Then
                '�p�G���ݩ�
                If .HasAttributes = True Then
                    '�������ޥΤ����ݩ�
                    Array1 = .GetAttributes
                    '�o�@���j��ΨӬd����D�A�p�G����b��1��
                    For Count = LBound(Array1) To UBound(Array1)
                        '�p�G�٨S�����D
                        If Array1(Count).TagString = title Then
                            Array1(Count).textString = sval
                            Exit Function
                        End If
                    Next Count
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
    Dim sval As String
    'Dim entObj As AcadEntity
    Dim Count As Integer
        With entAttrib
            '��@�Ӷ��ޥΪ��Q����A�ˬd���O�_���ݩ�
            If StrComp(.EntityName, "AcDbBlockReference", 1) = 0 Then
                '�p�G���ݩ�
                If .HasAttributes = True Then
                    '�������ޥΤ����ݩ�
                    Array1 = .GetAttributes
                    '�o�@���j��ΨӬd����D�A�p�G����b��1��
                    For Count = LBound(Array1) To UBound(Array1)
                       
                        title = Array1(Count).TagString
                        sval = Array1(Count).textString
                        
                        Call SetAttribTextString(towAttrib, title, sval)
                        
                    Next Count
                    
                    
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
    Dim Count As Integer
    strList = ""
        With entAttrib
            '��@�Ӷ��ޥΪ��Q����A�ˬd���O�_���ݩ�
            If StrComp(.EntityName, "AcDbBlockReference", 1) = 0 Then
                '�p�G���ݩ�
                If .HasAttributes = True Then
                    '�������ޥΤ����ݩ�
                    Array1 = .GetAttributes
                    '�o�@���j��ΨӬd����D�A�p�G����b��1��
                    For Count = LBound(Array1) To UBound(Array1)
                        strList = strList & Array1(Count).TagString & vbCrLf
                    Next Count
                    GetAttribList = strList
                    Exit Function
                End If
            End If
        End With
End Function












