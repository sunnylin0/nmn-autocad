Attribute VB_Name = "��r�Х�"
Option Explicit


Public Sub select_qin()
On Error Resume Next
Dim k As Integer
For k = 0 To 500
    Dim ent As Object

    'Dim tmp_ent() As Variant
    
   
    Dim SSetColl As AcadSelectionSets
    Set SSetColl = ThisDrawing.SelectionSets
    
    ' Create a SelectionSet named "TEST" in the current drawing
    Dim ssetObj As AcadSelectionSet
    
  
    If SSetColl.count > 0 Then
        SSetColl.item(0).Delete
    End If
    
    Set ssetObj = SSetColl.add("TEST")
    

    ssetObj.SelectOnScreen
    If ssetObj.count = 0 Then
    
       Exit Sub
    End If
    '��X�ϭ������C�@�ӿﶰ
    ReDim tmp_ent(ssetObj.count - 1)
    Dim lll As AcadLine
    
    
    Dim minExt As Variant
    Dim maxExt As Variant
    Dim minPT(2) As Double
    Dim maxPT(2) As Double
    Dim midPT(2) As Double
    Dim toPT(2) As Double
    Dim aa1 As AcadText
    Dim i As Integer
    Dim pos_y As Double
    i = 0
    For Each ent In ssetObj
        '���o�̤j�̤p
        ent.GetBoundingBox minExt, maxExt
        If i = 0 Then
            minPT(0) = minExt(0): minPT(1) = minExt(1)
            maxPT(0) = maxExt(0): maxPT(1) = maxExt(1)
        Else
            If minPT(0) > minExt(0) Then
                minPT(0) = minExt(0)
            End If
            If minPT(1) > minExt(1) Then
                minPT(1) = minExt(1)
            End If
            
        End If
        i = i + 1
    Next
    
    '����
    toPT(0) = 60000 * k: toPT(1) = 20000:
    For Each ent In ssetObj
        '���o�̤j�̤p
        
        ent.Move minPT, toPT
    Next
    
Next
End Sub
