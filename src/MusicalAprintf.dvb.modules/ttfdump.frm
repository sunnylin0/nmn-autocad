VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ttfdump 
   Caption         =   "���o ttfdump �����"
   ClientHeight    =   3675
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4752
   OleObjectBlob   =   "ttfdump.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ttfdump"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim xylist As New PointList

Private Sub comDrawing_Click()
    Dim pt As Variant
    Dim xyString As String
    Dim sX As String
    Dim sY As String
    Dim i As Long
    
    Me.Hide
    ' Return a point using a prompt
    pt = ThisDrawing.Utility.GetPoint(, "\n��ܭn���J���I �GEnter insertion point: ")
    '�e�X�w��u
    
    For i = 0 To xylist.size - 1
        sX = xylist.at(i).x + pt(0)
        sY = xylist.at(i).y + pt(1)
        xyString = xyString & sX & "," & sY & " "
    Next
    
    ThisDrawing.SendCommand "line " & xyString
    
End Sub

Private Sub CommandButton1_Click()
    Dim arr As Variant
    Dim SearchString, SearchChar, MyPos
    
    SearchChar = "abs"    ' �n�M��r�� "abs"�C
    Dim i As Long
    Dim y As Long
    Dim pt As New point
    Dim SStr As String
    arr = Split(Me.TextBox1.text, vbCrLf)

    xylist.clean
    Me.TextBox2.text = ""
    For i = 0 To UBound(arr) - 1
        If InStr(1, arr(i), "ABS", 1) <> 0 Then
            SearchString = arr(i)   ' �Q�j�M���r��C
            ' �p�g p �M�j�g P �b [��r���] �U�O�@�˪��C
            MyPos = InStr(1, SearchString, SearchChar, 1)
            arr(i) = Mid(arr(i), MyPos, Len(arr(i)))    '���o ABS()�����
            
                
            SStr = Replace(arr(i), " ", "")       '�ťղ���
            
            Set pt = getABS_XY(SStr)
            Call xylist.addpt(pt)
        End If
        Me.TextBox2 = Me.TextBox2 & pt.x & "," & pt.y & vbCrLf
        'Me.TextBox2 = Me.TextBox2.text & vbCrLf & ARR(i) '�ݸ��
    Next
'
'    For i = 0 To UBound(ARR) - 1
'        Me.TextBox2 = Me.TextBox2.text & vbCrLf & ARR(i)
'    Next
End Sub

Private Function getABS_XY(dataXY As String) As point
'���o�A������ xy ���
    Dim fpos As Integer
    Dim mpos As Integer
    Dim lpos As Integer
    Dim i As Integer
    Dim sX As String
    Dim sY As String
    Dim tPt As New point
    
    fpos = InStr(1, dataXY, "(")
    mpos = InStr(1, dataXY, ",")
    lpos = InStr(fpos, dataXY, ")")
    
    sX = Mid(dataXY, fpos + 1, mpos - fpos - 1)
    sY = Mid(dataXY, mpos + 1, lpos - mpos - 1)
    tPt.x = CDbl(sX)
    tPt.y = CDbl(sY)
    Set getABS_XY = tPt
End Function
