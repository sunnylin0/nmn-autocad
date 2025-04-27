VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPoline2Txt 
   Caption         =   "UserForm1"
   ClientHeight    =   3225
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4704
   OleObjectBlob   =   "frmPoline2Txt.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPoline2Txt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


    Dim g_object As AcadLWPolyline
Function get_mulitpoint() As Integer
    Dim str As String
    Dim i As Integer
    Dim getPT As Double
    Dim ina As Integer
    str = ""
    If (g_object.objectName <> "AcDbPolyline") Then
        get_mulitpoint = 1
        Exit Function
    End If
    Dim pointWCS(2) As Double
    Dim pointUCS As Variant
    pointWCS(0) = 0#
    pointWCS(1) = 0#
    pointWCS(2) = 0#
    


    For i = 0 To UBound(g_object.Coordinates)
        getPT = g_object.Coordinates(i)
        If (i Mod 2) = 0 Then
            pointWCS(0) = getPT
            pointUCS = ThisDrawing.Utility.TranslateCoordinates(pointWCS, acWorld, acUCS, False)
            ina = pointUCS(0)
        Else
            pointWCS(1) = getPT
            pointUCS = ThisDrawing.Utility.TranslateCoordinates(pointWCS, acWorld, acUCS, False)
            ina = pointUCS(1)
        End If
        
        If i = 0 Then
            str = ina
        Else
            str = str & "," & ina
        End If
        If Len(str) > 200 Then
            ThisDrawing.Utility.Prompt (str)
            str = ""
        End If
    Next
    ThisDrawing.Utility.Prompt (str)
    Me.TextBox1 = Me.TextBox1 & str
    get_mulitpoint = 0
End Function

Sub Example_SelectOnScreen() 'AutoCAD 取得 LWPolyLine 的各點
 ' Have the user enter a point
    Dim pointG As Variant
    pointG = ThisDrawing.Utility.GetPoint(, "Enter a point to translate:")
    

    Dim ucsObj As AcadUCS
    Dim origin(0 To 2) As Double
    Dim xAxisPnt(0 To 2) As Double
    Dim yAxisPnt(0 To 2) As Double
    
    ' Define the UCS
    xAxisPnt(0) = pointG(0) + 5: xAxisPnt(1) = pointG(1): xAxisPnt(2) = pointG(2)
    yAxisPnt(0) = pointG(0): yAxisPnt(1) = pointG(1) + 5: yAxisPnt(2) = pointG(2)
    
    ' Add the UCS to the UserCoordinatesSystems collection
    Set ucsObj = ThisDrawing.UserCoordinateSystems.Add(pointG, xAxisPnt, yAxisPnt, "New_UCS")
    ThisDrawing.ActiveUCS = ucsObj


    
    ' Create the selection set
    Dim ssetObj As AcadSelectionSet
    'Set ssetObj = ThisDrawing.SelectionSets.Add("TEST_SSET")
    If ThisDrawing.SelectionSets.Count = 0 Then
        Set ssetObj = ThisDrawing.SelectionSets.Add("TEST_SSET")
    Else
        Set ssetObj = ThisDrawing.SelectionSets.item(0)
        ssetObj.Clear
    
    End If
    
    ' Add objects to a selection set by prompting user to select on the screen
    ssetObj.SelectOnScreen
    ssetObj.item (0)
    Dim returnPnt As Variant
    Dim str As String
    Dim strAll As String
    Dim i As Integer
    Dim ina As Integer
    Dim counta As Integer
    counta = ssetObj.Count()
    
    str = ""
    strAll = ""
    ThisDrawing.Utility.Prompt ("drawXY=")
    
    ' Return a point using a prompt
    ' Translate the point into UCS coordinates
    For i = 0 To counta - 1
        Set g_object = ssetObj.item(i)
        ina = get_mulitpoint
        If (i <> counta - 1) And (ina = 0) Then
            ThisDrawing.Utility.Prompt (",o,")
            
            Me.TextBox1 = Me.TextBox1 & ",o,"
        End If
    Next


End Sub



Sub scale_SelectOnScreen() 'AutoCAD 取得 LWPolyLine 的各點
 ' Have the user enter a point
    Dim pointG As Variant
    pointG = ThisDrawing.Utility.GetPoint(, "Enter a point to translate:")
    

    Dim ucsObj As AcadUCS
    Dim origin(0 To 2) As Double
    Dim xAxisPnt(0 To 2) As Double
    Dim yAxisPnt(0 To 2) As Double
    
    ' Define the UCS
    origin(0) = pointG(0): origin(1) = pointG(1): origin(2) = pointG(2)
    xAxisPnt(0) = pointG(0) + 5: xAxisPnt(1) = pointG(1): xAxisPnt(2) = pointG(2)
    yAxisPnt(0) = pointG(0): yAxisPnt(1) = pointG(1) + 5: yAxisPnt(2) = pointG(2)
    


    
    ' Create the selection set
    Dim ssetObj As AcadSelectionSet
    'Set ssetObj = ThisDrawing.SelectionSets.Add("TEST_SSET")
    If ThisDrawing.SelectionSets.Count = 0 Then
        Set ssetObj = ThisDrawing.SelectionSets.Add("TEST_SSET")
    Else
        Set ssetObj = ThisDrawing.SelectionSets.item(0)
        ssetObj.Clear
    
    End If
    
    ' Add objects to a selection set by prompting user to select on the screen
    ssetObj.SelectOnScreen
    ssetObj.item (0)
    Dim returnPnt As Variant
    Dim str As String
    Dim dd As AcadLWPolyline
    Dim i As Integer
    Dim ina As Integer
    Set dd = ssetObj.item(0)
    str = ""
    ' Return a point using a prompt
        ' Translate the point into UCS coordinates
    Dim pointWCS(2) As Double
    Dim pointUCS As Variant
    pointWCS(0) = 0#
    pointWCS(1) = 0#
    pointWCS(2) = 0#
    

    dd.ScaleEntity pointG, 1.03179

    'ThisDrawing.Utility.Prompt (str)
    Me.TextBox1 = str

End Sub



Private Sub CommandButton1_Click()
    Me.TextBox1 = ""
    Me.Hide
    Example_SelectOnScreen
    Me.Show
End Sub

Private Sub CommandButton2_Click()

'使用 VB 建立線： (工程引用'AutoCad 2004 Type Library')
     On Error Resume Next
'    Dim acadApp As AcadApplication
'    Set acadApp = GetObject(, "AutoCAD.Application.16")
'    If Err Then
'        Err.clear
'        Set acadApp = CreateObject("AutoCAD.Application.16")
'        If Err Then
'            MsgBox Err.Description
'            Exit Sub
'        End If
'    End If
'    acadApp.Visible = True
'    Dim acadDoc As AcadDocument
'    Set acadDoc = acadApp.ActiveDocument
    

    Dim plineObj As AcadPolyline
    Dim stXY() As String
    Dim ReadData() As String
    Dim pt(1000, 2) As Double
    Dim ptlist() As Double
  
    
    Dim ssst As String
    ssst = Replace(Me.TextBox1, vbCrLf, ",")
    ssst = Replace(ssst, vbTab, ",")
    ssst = Replace(ssst, ",", ",")
    ssst = Replace(ssst, "  ", ",")
    ssst = Replace(ssst, "  ", ",")
    
    ReadData = Split(ssst, ",")
    
    Dim i As Integer
    Dim posI As Integer
    Dim chkXY As Integer
    
    posI = 0
    chkXY = 0
    For i = 0 To UBound(ReadData)
        If ReadData(i) <> "" Then
            If chkXY = 0 Then
                pt(posI, 0) = ReadData(i)
                chkXY = 1
            Else
                pt(posI, 1) = ReadData(i)
                posI = posI + 1
                chkXY = 0
            End If
        End If
    Next
    If posI <= 1 Then
        Exit Sub
    End If
    ReDim ptlist((posI * 3) - 1)
    Dim iScale As Double
    iScale = tbSCALE
    If iScale = 0 Then
        iScale = 1
    End If
    
    For i = 0 To posI
        ptlist(i * 3) = pt(i, 0) * iScale
        ptlist(i * 3 + 1) = pt(i, 1) * iScale
    Next
    
    Dim pointG As Variant
    pointG = ThisDrawing.Utility.GetPoint(, "要畫的原點:")
    
    Dim ucsObj As AcadUCS
    Dim origin(0 To 2) As Double
    Dim xAxisPnt(0 To 2) As Double
    Dim yAxisPnt(0 To 2) As Double
    
    ' Define the UCS
    xAxisPnt(0) = pointG(0) + 5: xAxisPnt(1) = pointG(1): xAxisPnt(2) = pointG(2)
    yAxisPnt(0) = pointG(0): yAxisPnt(1) = pointG(1) + 5: yAxisPnt(2) = pointG(2)
    
    ' Add the UCS to the UserCoordinatesSystems collection
    Set ucsObj = ThisDrawing.UserCoordinateSystems.Add(pointG, xAxisPnt, yAxisPnt, "New_UCS")
    ThisDrawing.ActiveUCS = ucsObj
    
    
    Set plineObj = ThisDrawing.ModelSpace.AddPolyline(ptlist)
    
    
    ZoomAll
    
'    Set acadDoc = Nothing
'    Set acadApp = Nothing
   
End Sub
