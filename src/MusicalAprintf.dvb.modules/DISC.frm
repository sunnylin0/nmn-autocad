VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DISC 
   Caption         =   "UserForm1"
   ClientHeight    =   3225
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4704
   OleObjectBlob   =   "DISC.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DISC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbLoad_Click()
  
    Dim TrackingDictionary As AcadDictionary, TrackingXRecord As AcadXRecord
    Dim XRecordDataType As Variant, XRecordData As Variant
    Dim ArraySize As Long, iCount As Long
    Dim DataType As Integer, data As String, msg As String
    
        ' Unique identifiers to distinguish our XRecordData from other XRecordData
    Const TYPE_STRING = 1
    Const TAG_DICTIONARY_NAME = "ObjectTrackerDictionary"
    Const TAG_XRECORD_NAME = "ObjectTrackerXRecord"

    ' Connect to the dictionary we store the XRecord in
    On Error GoTo create
    Set TrackingDictionary = ThisDrawing.Dictionaries(TAG_DICTIONARY_NAME)
    Set TrackingXRecord = TrackingDictionary.GetObject(TAG_XRECORD_NAME)
    On Error GoTo 0
    
 
    ' Read back all XRecordData entries
    TrackingXRecord.GetXRecordData XRecordDataType, XRecordData
    ArraySize = UBound(XRecordDataType)
    
    ' Retrieve and display stored XRecordData
    For iCount = 0 To ArraySize
        ' Get information for this element
        DataType = XRecordDataType(iCount)
        data = XRecordData(iCount)
        
        If DataType = TYPE_STRING Then
            msg = msg & data & vbCrLf
        End If
    Next
    Me.TextBox1 = msg
    ''MsgBox "The data in the XRecord is: " & vbCrLf & vbCrLf & msg, vbInformation
    
    Exit Sub

create:
    ' Create the entities that hold our XRecordData
    If TrackingDictionary Is Nothing Then  ' Make sure we have our tracking object
        Set TrackingDictionary = ThisDrawing.Dictionaries.Add(TAG_DICTIONARY_NAME)
        Set TrackingXRecord = TrackingDictionary.AddXRecord(TAG_XRECORD_NAME)
    End If
    
    Resume
End Sub

Private Sub CommandButton1_Click()

    ' This example creates a new XRecord if one doesn't exist,
    ' appends data to the XRecord and reads it back.  To see data being added
    ' run the example more than once.
    
    Dim TrackingDictionary As AcadDictionary, TrackingXRecord As AcadXRecord
    Dim XRecordDataType As Variant, XRecordData As Variant
    Dim ArraySize As Long, iCount As Long
    Dim DataType As Integer, data As String, msg As String
    
    ' Unique identifiers to distinguish our XRecordData from other XRecordData
    Const TYPE_STRING = 1
    Const TAG_DICTIONARY_NAME = "ObjectTrackerDictionary"
    Const TAG_XRECORD_NAME = "ObjectTrackerXRecord"

    ' Connect to the dictionary we store the XRecord in
    On Error GoTo create
    Set TrackingDictionary = ThisDrawing.Dictionaries(TAG_DICTIONARY_NAME)
    Set TrackingXRecord = TrackingDictionary.GetObject(TAG_XRECORD_NAME)
    On Error GoTo 0
    
    ' Get current XRecordData
    TrackingXRecord.GetXRecordData XRecordDataType, XRecordData
    
    ' If we don't have an array already then create one
    If VarType(XRecordDataType) And vbArray = vbArray Then
        ArraySize = UBound(XRecordDataType) + 1       ' Get the size of the data elements returned
        ArraySize = ArraySize + 1                        ' Increase to hold new data
    
        ReDim Preserve XRecordDataType(0 To ArraySize)
        ReDim Preserve XRecordData(0 To ArraySize)
    Else
        ArraySize = 0
        ReDim XRecordDataType(0 To ArraySize) As Integer
        ReDim XRecordData(0 To ArraySize) As Variant
    End If
    
    ' Append new XRecord Data
    '
    ' For this sample we only append the current time to the XRecord
    XRecordDataType(ArraySize) = TYPE_STRING: XRecordData(ArraySize) = CStr(Now)
    TrackingXRecord.SetXRecordData XRecordDataType, XRecordData
    
    ' Read back all XRecordData entries
    TrackingXRecord.GetXRecordData XRecordDataType, XRecordData
    ArraySize = UBound(XRecordDataType)
    
    ' Retrieve and display stored XRecordData
    For iCount = 0 To ArraySize
        ' Get information for this element
        DataType = XRecordDataType(iCount)
        data = XRecordData(iCount)
        
        If DataType = TYPE_STRING Then
            msg = msg & data & vbCrLf
        End If
    Next
    
    MsgBox "The data in the XRecord is: " & vbCrLf & vbCrLf & msg, vbInformation
    
    Exit Sub

create:
    ' Create the entities that hold our XRecordData
    If TrackingDictionary Is Nothing Then  ' Make sure we have our tracking object
        Set TrackingDictionary = ThisDrawing.Dictionaries.Add(TAG_DICTIONARY_NAME)
        Set TrackingXRecord = TrackingDictionary.AddXRecord(TAG_XRECORD_NAME)
    End If
    
    Resume
End Sub

Private Sub CommandButton2_Click()
    
    Dim TrackingDictionary As AcadDictionary, TrackingXRecord As AcadXRecord
    Dim XRecordDataType As Variant, XRecordData As Variant
    Dim ArraySize As Long, iCount As Long
    Dim DataType As Integer, data As String, msg As String
    
    ' Unique identifiers to distinguish our XRecordData from other XRecordData
    Const TYPE_STRING = 1
    Const TAG_DICTIONARY_NAME = "ObjectTrackerDictionary"
    Const TAG_XRECORD_NAME = "ObjectTrackerXRecord"

    ' Connect to the dictionary we store the XRecord in
    On Error GoTo create
    Set TrackingDictionary = ThisDrawing.Dictionaries(TAG_DICTIONARY_NAME)
    Set TrackingXRecord = TrackingDictionary.GetObject(TAG_XRECORD_NAME)
End Sub
Option Explicit

'使用擴展字典
'AutoCAD有兩種字典
'第一種：擴展字典——一種與具體的對象關聯的字典，每個對象僅能擁有一個擴展字典，
'                    其中可以保存與該對象有關的信息。
'第二種：命名對象字典(Named Object Dictionary)——用來保存與實體無關的數據，
'                    AutoCAD本身就使用對象命名字典來保存一些信息，例如組字典（保存編組的信息）
'                    多線樣式字典等。

'A 使用擴展字典的一般步驟：
'  保存信息的基本步驟：
'  1、使用GetExtensionDictionary方法創建一個對象的擴展字典
'  2、使用AddXData方法向擴展字典添加一條擴展記錄
'  3、使用SetXRecordData將數據保存在擴展記錄中。
'  讀取擴展字典的基本步驟：
'  1、使用GetExtensionDictionary獲取對象的擴展字典
'  2、使用GetObject方法獲得指定的擴展記錄
'  3、使用GetXRecordData讀取保存在擴展記錄中的數據。

'B 使用命名對象字典的一般步驟：
'  保存命名字典信息的基本步驟
'  1、使用Dictionaries.add添加一個命名對象字典
'  2、同A
'  3、同A
'  讀取命名對象字典的基本步驟：
'  1、使用對象的基本操作獲得Dictionaries集合中制定的字典
'  2、同A
'  3、同A

Public Function HasXRecord(ByVal ent As AcadEntity, ByVal key As String) As Boolean
  '判斷對象是否已經擁有擴展字典
  Dim objDict As AcadDictionary
  Dim objXRecord As AcadXRecord
 
  If ent.HasExtensionDictionary Then
    '獲得擴展字典
    Set objDict = ent.GetExtensionDictionary
   
    On Error Resume Next
    Set objXRecord = objDict.GetObject(key)
    Set objDict = Nothing
   
    If err Then
      err.Clear
      HasXRecord = False
    Else
      HasXRecord = True
    End If
  Else
    HasXRecord = False
  End If
   
End Function

Public Sub CreateXRecord(ByRef xDataType As Variant, ByRef xData As Variant, ParamArray Filter())
  '創建擴展記錄的數據集
  Debug.Assert (UBound(Filter) Mod 2 = 1)
 
  Dim Count As Integer
  Count = (UBound(Filter) + 1) / 2
  Dim DataType() As Integer, data() As Variant
  ReDim DataType(Count - 1)
  ReDim data(Count - 1)
 
  Dim i As Integer
  For i = 0 To Count - 1
    DataType(i) = Filter(2 * i)
    data(i) = Filter(2 * i + 1)
  Next i
    xDataType = DataType
    xData = data
End Sub

Public Sub AddXRecord(ByVal ent As AcadEntity, ByVal key As String, ByVal xDataType As Variant, ByVal xData As Variant)
  '向指定的實體添加擴展字典
  Dim objDict As AcadDictionary
  Dim objXRecord As AcadXRecord
 
  Set objDict = ent.GetExtensionDictionary()
  Set objXRecord = objDict.AddXRecord(key)
  objXRecord.SetXRecordData xDataType, xData
End Sub
  
Public Sub GetXRecord(ByVal ent As AcadEntity, ByVal key As String, ByRef xDataType As Variant, ByRef xData As Variant)
  '獲取擴展字典
  Dim objDict As AcadDictionary
  Dim objXRecord As AcadXRecord
 
  Set objDict = ent.GetExtensionDictionary
  Set objXRecord = objDict.GetObject(key)
  objXRecord.GetXRecordData xDataType, xData
End Sub

Public Function HasNamedDictionary(ByVal DictName As String, ByVal key As String) As Boolean
  '判斷是否已經存在鍵位dictName的字典
  On Error Resume Next
 
  Dim objDict As AcadDictionary
  Dim objXRecord As AcadXRecord
 
  Set objDict = ThisDrawing.Dictionaries(DictName)
  If err Then
    err.Clear
    HasNamedDictionary = False
  Else
    Set objXRecord = objDict.GetObject(key)
    If err Then
      err.Clear
      HasNamedDictionary = False
    Else
      HasNamedDictionary = True
    End If
  End If
End Function

Public Sub AddNamedDictionary(ByVal DictName As String, ByVal key As String, ByVal xDataType As Variant, ByVal xData As Variant)
  '用於向當前圖形添加指定的命名對象字典
  Dim objDict As AcadDictionary
  Dim objXRecord As AcadXRecord
 
  Set objDict = ThisDrawing.Dictionaries.Add(DictName)
  Set objXRecord = objDict.AddXRecord(key)
  objXRecord.SetXRecordData xDataType, xData
End Sub

Public Sub GetNamedDictionary(ByVal DictName As String, ByVal key As String, ByRef xDataType As Variant, ByRef xData As Variant)
  '從當前圖形中獲得指定的命名對象字典
  Dim objDict As AcadDictionary
  Dim objXRecord As AcadXRecord
 
  Set objDict = ThisDrawing.Dictionaries(DictName)
  Set objXRecord = objDict.GetObject(key)
  objXRecord.GetXRecordData xDataType, xData
End Sub

 

調用類:

Option Explicit

Public Sub AddEntXRecord()
  Dim ent As AcadEntity
  Dim PickPoint As Variant
 
  ThisDrawing.Utility.GetEntity ent, PickPoint, vbCr & "請選擇對象："
 
  Dim point(0 To 2) As Double
  SetPoint3d point, 100, 100, 0
 
  Dim xRecord As New clsXRecord
  Dim DataType As Variant, data As Variant
 
  xRecord.GetXRecord ent, "EX02", DataType, data
 
  Dim s As String
  s = data(1)
 
  xRecord.CreateXRecord DataType, data, _
      1, "道路", _
      8, ent.Layer, _
      40, PickPoint(0), _
      10, point
 
  If xRecord.HasXRecord(ent, "EX02") Then
    ThisDrawing.Utility.Prompt vbNewLine & "實體已經包含指定名稱的擴展記錄 "
  Else
    xRecord.AddXRecord ent, "EX02", DataType, data
    ThisDrawing.Utility.Prompt vbNewLine & "成功為實體添加擴展記錄 "
  End If
 
 
 
End Sub


Public Sub SetPoint3d(ByVal pt As Variant, ByRef x As Double, ByRef y As Double, ByRef Z As Double)
  ReDim pt(0 To 2) As Double
  pt(0) = x
  pt(1) = y
  pt(2) = Z
 
End Sub

Private Sub CommandButton3_Click()
    '把 textbox1 內的多文字，轉換成多個 AcadText
    Dim returnPnt As Variant
    
    Me.Hide
    returnPnt = ThisDrawing.Utility.GetPoint(, "選擇插入點： ")
    
    
    Dim textObj As AcadMText
    Dim textString As String
    Dim insertionPoint(0 To 2) As Double
    Dim height As Double
    Dim i As Integer
    
    For i = 1 To Len(Me.TextBox1.text)
        
        textString = MidB(Me.TextBox1.text, i, 1)
        insertionPoint(0) = returnPnt(0) + i * 10
        insertionPoint(1) = returnPnt(1)
        insertionPoint(2) = 0
        height = 3
    
        ' Create the text object in model space
        Set textObj = ThisDrawing.ModelSpace.AddMText(insertionPoint, 3, textString)
    Next
 
End Sub

Private Sub CommandButton4_Click()
    Dim returnObj As AcadObject
    Dim basePnt As Variant
    Dim arrStr As Variant
    On Error Resume Next
    
    ' The following example waits for a selection from the user
    Me.Hide
    ThisDrawing.Utility.GetEntity returnObj, basePnt, "Select an object"
    arrStr = Split(returnObj.textString, "\P")
    
    Dim textObj As AcadMText
    Dim textString As String
    Dim insertionPoint(0 To 2) As Double
    Dim height As Double
    Dim i As Integer
    
    For i = 0 To UBound(arrStr)
        
        textString = arrStr(i)
        insertionPoint(0) = basePnt(0) + 30
        insertionPoint(1) = basePnt(1) - (i * 5)
        insertionPoint(2) = 0
        height = 3
    
        ' Create the text object in model space
        Set textObj = ThisDrawing.ModelSpace.AddMText(insertionPoint, 5, textString)
        textObj.height = 5
    Next
    
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Deactivate()

End Sub

Private Sub UserForm_ QueryClose(Cancel As Integer, CloseMode As Integer)

End Sub
