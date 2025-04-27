VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMusical 
   Caption         =   "Musical"
   ClientHeight    =   3990
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4860
   OleObjectBlob   =   "frmMusical.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMusical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    '插入單一Text
    Dim MusicText As NnmText
    Dim basePnt(0 To 2) As Double
    Dim cenPnt(0 To 2) As Double
    Dim returnPnt As Variant
    Dim pt As Variant
    Dim rad As Integer
    Me.Hide
    returnPnt = ThisDrawing.Utility.GetPoint(basePnt, "\n選擇要插入的點")
    

    'Set SEOBJ = New SmileyEntity
    pt = returnPnt
    Set MusicText = ThisDrawing.ModelSpace.AddCustomObject("MUSICTEXT")
    'Call MusicText.setNewObject(returnPnt, 1)
    Call MusicText.setData(pt, "ba!1=z", "MMP2005", 44)
    returnPnt(0) = returnPnt(0) + 31.83
    
End Sub

Private Sub CommandButton2_Click()
    '插入多個Text
    Dim MusicText As NnmText
    Dim basePnt(0 To 2) As Double
    Dim cenPnt(0 To 2) As Double
    Dim returnPnt As Variant
    Dim pt As Variant
    Dim rad As Integer
    Me.Hide
    returnPnt = ThisDrawing.Utility.GetPoint(basePnt, "\n選擇要插入的點")
    

    'Set SEOBJ = New SmileyEntity
    pt = returnPnt
    pt(0) = 0: pt(1) = 0: pt(2) = 0:
    Set MusicText = ThisDrawing.ModelSpace.AddCustomObject("MUSICTEXT")
    'Call MusicText.setNewObject(returnPnt, 1)
    Call MusicText.setData(returnPnt, "bf!1=z", "MMP2005", 44)
    returnPnt(0) = returnPnt(0) + 31.83
    Set MusicText = ThisDrawing.ModelSpace.AddCustomObject("MUSICTEXT")
    Call MusicText.setNewObject(returnPnt, 1)
    Call MusicText.setData(returnPnt, " A?2=c", "MMP2005", 44)
    returnPnt(0) = returnPnt(0) + 31.83
    Set MusicText = ThisDrawing.ModelSpace.AddCustomObject("MUSICTEXT")
    Call MusicText.setNewObject(returnPnt, 1)
    Call MusicText.setData(returnPnt, "#t.3=a", "MMP2005", 44)
    
End Sub

Private Sub CommandButton33_Click()
'取得物件 MusicText
'用x 軸來鏈結
On Error GoTo EndJoin

    Dim ent As Object
    Me.Hide

StartJoin:
    Dim SSetColl As AcadSelectionSets
    Set SSetColl = ThisDrawing.SelectionSets
    
    ' Create a SelectionSet named "TEST" in the current drawing
    Dim ssetObj As AcadSelectionSet
  

    If SSetColl.Count > 0 Then
        SSetColl.item(0).delete
    End If
    Set ssetObj = SSetColl.Add("TEST")

    ssetObj.SelectOnScreen
    
    '排序陣列
    Dim TheArray As Variant
    Dim TheArrayEnt As Variant
    Dim coI As Integer
    Dim i As Long
    
    ' Add the dimstyle to the dimstyles collection
    coI = 0
  ' 找出圖面中的每一個選集
    For Each ent In ssetObj
        If (ent.objectName = "MusicText") Then
            coI = coI + 1
        End If
    Next
    
    '只有一個，不用排序了
    If coI < 1 Then
        GoTo StartJoin
    End If
    
    ReDim TheArray(coI - 1)
    ReDim TheArrayEnt(coI - 1)
    
    coI = 0
    Dim MTobj As NnmText
    Dim pt As Variant
    For Each ent In ssetObj
        If (ent.objectName = "MusicText") Then
            pt = ent.insertionPoint()
            TheArray(coI) = pt(0)
            Set TheArrayEnt(coI) = ent
            coI = coI + 1
        End If
    Next
    
    
    
    '排序開始
    SelectionSortMaxTow TheArray, TheArrayEnt
    
    For i = 0 To UBound(TheArray)
        If i = 0 Then
            TheArrayEnt(i).LeftLink = 0
            TheArrayEnt(i).RightLink = TheArrayEnt(i + 1).ObjectID
        ElseIf i = UBound(TheArray) Then
            TheArrayEnt(i).LeftLink = TheArrayEnt(i - 1).ObjectID
            TheArrayEnt(i).RightLink = 0
        Else
            TheArrayEnt(i).RightLink = TheArrayEnt(i + 1).ObjectID
        End If
        TheArrayEnt(i).Update
    Next
GoTo StartJoin

EndJoin:
End Sub

Private Sub CommandButton3_Click()

'取得物件 MusicText
'用x 軸來鏈結
On Error GoTo EndJoin

    Dim ent As Object
    Me.Hide

StartJoin:
    Dim SSetColl As AcadSelectionSets
    Set SSetColl = ThisDrawing.SelectionSets
    
    ' Create a SelectionSet named "TEST" in the current drawing
    Dim ssetObj As AcadSelectionSet
  

    If SSetColl.Count > 0 Then
        SSetColl.item(0).delete
    End If
    Set ssetObj = SSetColl.Add("TEST")

    ssetObj.SelectOnScreen
    
    '排序陣列
    Dim TheArray As Variant
    Dim TheArrayEnt As Variant
    Dim coI As Integer
    Dim i As Long
    
    ' Add the dimstyle to the dimstyles collection
    coI = 0
  ' 找出圖面中的每一個選集
    For Each ent In ssetObj
        If (ent.objectName = "MusicText") Then
            coI = coI + 1
        End If
    Next
    
    ReDim TheArray(coI - 1)
    ReDim TheArrayEnt(coI - 1)
    
    coI = 0
    Dim MTobj As NnmText
    Dim pt As Variant
    For Each ent In ssetObj
        If (ent.objectName = "MusicText") Then
            pt = ent.insertionPoint()
            TheArray(coI) = pt(0)
            Set TheArrayEnt(coI) = ent
            coI = coI + 1
        End If
    Next
    
    Me.addMusicJoin TheArrayEnt
EndJoin:
End Sub

Private Sub CommandButton4_Click()
'清除 MusicText 的鏈結
On Error Resume Next
    Dim ent As Object

   
    Dim SSetColl As AcadSelectionSets
    Set SSetColl = ThisDrawing.SelectionSets
    
    ' Create a SelectionSet named "TEST" in the current drawing
    Dim ssetObj As AcadSelectionSet
  
    If SSetColl.Count > 0 Then
        SSetColl.item(0).delete
    End If
    Set ssetObj = SSetColl.Add("TEST")
    Me.Hide
    ssetObj.SelectOnScreen
    
    '排序陣列
    Dim TheArray As Variant
    Dim TheArrayEnt As Variant
    Dim coI As Integer
    Dim i As Long
    
    ' Add the dimstyle to the dimstyles collection
    coI = 0
  ' 找出圖面中的每一個選集
    For Each ent In ssetObj
        If (ent.objectName = "MusicText") Then
            ent.RightLink = 0
            ent.LeftLink = 0
        End If
    Next
End Sub


Public Function addMusicJoin(ByVal the_Ids As Variant)

'取得物件 MusicText
'用x 軸來鏈結
On Error Resume Next

    Dim ent As Object
    
    '排序陣列
    Dim TheArray As Variant
    Dim TheArrayEnt As Variant
    Dim coI As Integer
    Dim i As Long
    
    ' Add the dimstyle to the dimstyles collection
    coI = 0
  ' 找出圖面中的每一個選集
    For i = 0 To UBound(the_Ids, 1)
        If the_Ids(i).objectName = "MusicText" Then
            coI = coI + 1
        End If
    Next

    
    '只有一個，不用排序了
    If coI < 1 Then
        Exit Function
    End If
    
    ReDim TheArray(coI - 1)
    ReDim TheArrayEnt(coI - 1)
    
    coI = 0
    Dim MTobj As NnmText
    Dim pt As Variant
    
    For i = 0 To UBound(the_Ids, 1)
        Set ent = the_Ids(i)
        If (ent.objectName = "MusicText") Then
            pt = ent.insertionPoint()
            TheArray(coI) = pt(0)
            Set TheArrayEnt(coI) = ent
            coI = coI + 1
            
            
            Dim minExt As Variant
            Dim maxExt As Variant
    
            ent.GetBoundingBox minExt, maxExt
            Set MTobj = ent
            minExt = MTobj.GripLeft()
            
            
        End If
    Next
    
    
    
    '排序開始
    SelectionSortMaxTow TheArray, TheArrayEnt
    
    For i = 0 To UBound(TheArray)
        If i = 0 Then
            TheArrayEnt(i).LeftLink = 0
            TheArrayEnt(i).RightLink = TheArrayEnt(i + 1).ObjectID
        ElseIf i = UBound(TheArray) Then
            TheArrayEnt(i).LeftLink = TheArrayEnt(i - 1).ObjectID
            TheArrayEnt(i).RightLink = 0
        Else
            TheArrayEnt(i).RightLink = TheArrayEnt(i + 1).ObjectID
        End If
        TheArrayEnt(i).Update
    Next


EndJoin:
End Function

Sub getGrip()

End Sub
        

Public Function eraseMTlink(ByVal the_id As Variant)
End Function

Private Sub CommandButton5_Click()

Dim ttt As NnmText
Dim minExt As Variant
    Dim maxExt As Variant
    
    
    lineObj.GetBoundingBox

    ttt.GetBoundingBox minExt, maxExt
End Sub


Private Sub CommandButton6_Click()

'取得 grip 數值

On Error Resume Next
    Dim Object As Object
    Dim PickedPoint As Variant, TransMatrix As Variant, ContextData As Variant
    
        
TRYAGAIN:
        Me.Hide

    ' Get information about selected object
    ThisDrawing.Utility.GetSubEntity Object, PickedPoint, TransMatrix, ContextData


    Dim MTobj As NnmText
    Dim pt As Variant
    Dim SStr As String
    
        If (Object.objectName = "MusicText") Then
            Set MTobj = Object
            
            pt = MTobj.GripLeft()
            SStr = "GripLeft = " & Format(pt(0), "0.00") & ", " & Format(pt(1), "0.00") & vbCrLf
            
            pt = MTobj.GripMid()
            SStr = SStr & "nGripMid = " & Format(pt(0), "0.00") & ", " & Format(pt(1), "0.00") & vbCrLf
            
            pt = MTobj.GripRight()
            SStr = SStr & "nGripRight = " & Format(pt(0), "0.00") & ", " & Format(pt(1), "0.00") & vbCrLf
             
            pt = MTobj.GripLeftDown()
            SStr = SStr & "nGripLeftDown = " & Format(pt(0), "0.00") & ", " & Format(pt(1), "0.00") & vbCrLf
            
            pt = MTobj.GripMidUp()
            SStr = SStr & "nGripMidUp = " & Format(pt(0), "0.00") & ", " & Format(pt(1), "0.00") & vbCrLf
            
            
            ThisDrawing.Utility.Prompt SStr
        End If
End Sub







Private Sub CommandButton7_Click()
    ' This example returns the current layer
    ' and then adds a new layer.
    ' Finally, it returns the layer to the previous setting.
    Dim currLayer As AcadLayer
    Dim newLayer As AcadLayer
    
    ' Return the current layer of the active document
    Set currLayer = ThisDrawing.ActiveLayer
    MsgBox "The current layer is " & currLayer.Name, vbInformation, "ActiveLayer Example"
    
    ' Create a Layer and make it the active layer
    Dim i As Integer
    For i = 0 To ThisDrawing.Layers.Count - 1
        Set newLayer = ThisDrawing.Layers(i)
        Me.TextBox1 = Me.TextBox1 & "datalayer(i,1) = """ & newLayer.Name & """ : datalayer(i,2) =  " & newLayer.TrueColor.ColorIndex & vbCrLf
        
        
    Next
    
End Sub

Private Sub CommandButton8_Click()
    Dim TextColl As AcadTextStyles
    Set TextColl = ThisDrawing.TextStyles
    
    ' Create a Text style named "TEST" in the current drawing
    Dim textStyle As AcadTextStyle
    
    Dim i As Integer
    For i = 0 To ThisDrawing.TextStyles.Count - 1
        Set textStyle = ThisDrawing.TextStyles(i)
        Me.TextBox1 = Me.TextBox1 & "dataStyles(" & i & ",1)=""" & textStyle.Name & """" & vbCrLf _
            & "dataStyles(" & i & ",2)=""" & textStyle.fontFile & """" & vbCrLf _
            & "dataStyles(" & i & ",3)=""" & textStyle.BigFontFile & """" & vbCrLf
    Next
    
    
End Sub



Private Sub CommandButton9_Click()
    Dim str As Variant
    Dim i As Integer
    str = GetWindowsFonts
    
    For i = 0 To UBound(str)
        Me.TextBox1 = Me.TextBox1 & str(i) & vbCrLf
    Next
End Sub
