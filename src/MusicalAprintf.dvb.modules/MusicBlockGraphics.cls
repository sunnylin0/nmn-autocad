VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MusicBlockGraphics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    

Const MAINLAYER As String = "main"     '主要的圖層
'    '音附的駐點位置
'Private Type GripPoints
'    osnPtMid(0 To 2) As Double       '在正中間
'    osnPtNode1(0 To 2) As Double     '在左下
'    osnPtNode2(0 To 2) As Double     '在中上
'    osnPtEnd1(0 To 2) As Double      '在左中
'    osnPtEnd2(0 To 2) As Double      '在右中
'    atPt(0 To 2) As Double           '備注點
'End Type
Dim gTextSize As Double
Dim gAtPt As New point
Dim MData As String

Dim MFS As MusicFing

Dim Grippt As GripPoints  '音附指法 狀態


    Dim kFing(30 To 255) As String
    Dim posFing(0 To 255) As point
    Dim kTowFing(30 To 255) As String
    Dim posTowFing(0 To 255) As point
    
Const MAX200 As Integer = 200

Function Grip() As GripPoints
      Grip = Grippt
End Function

Property Let TextSize(TSize As Double)
    gTextSize = TSize
End Property
Property Get TextSize() As Double
    TextSize = gTextSize
End Property

Property Let DataTest(dtest As String)
    MData = dtest
End Property
Property Get DataTest() As String
    DataTest = MData
End Property
Property Get copy() As MusicBlockGraphics
    Dim copy_mbg As New MusicBlockGraphics
    copy_mbg.setDataText gAtPt, MData, gTextSize
    Set copy = copy_mbg
End Property

Sub setDataText(position As point, cst As String, TextSize As Double)

    'position 插入的位置
    'cst  文字資料=指法1 八度音 主音 拍子 指法2
    '          共5筆 每筆一個字元 所以共5個字元
    'textSize  字形大小
    
    MData = cst
    
    gAtPt.x = position.x
    gAtPt.Y = position.Y
    gAtPt.z = position.z
    
    gTextSize = TextSize
    setGripPoints gAtPt
    
    
    
'Public Type MusicFing
'    sTONE As String        ' 升降記號 行
'    sFinge As String   '這是指法行     _+)(*&
'    sScale As String   '這是高低音行   .:! ,;?
'    sNote As String   '這為主行       1234567.|l
'    sTempo As String   '這是拍子       -=368acefz
'    sTowFinge As String   '這是指法行第二行  _+)(*&
'    sSlur As String
'End Type
    

    Dim cc As String
    cc = ""

    '取得關鍵字  主音 AMT.iNote
    cc = Mid(MData, amt.iNote, 1)
    Select Case cc
       Case "0": MFS.sNote = "M-NOTE0"
       Case "1": MFS.sNote = "M-NOTE1"
       Case "2": MFS.sNote = "M-NOTE2"
       Case "3": MFS.sNote = "M-NOTE3"
       Case "4": MFS.sNote = "M-NOTE4"
       Case "5": MFS.sNote = "M-NOTE5"
       Case "6": MFS.sNote = "M-NOTE6"
       Case "7": MFS.sNote = "M-NOTE7"
       Case ".": MFS.sNote = "M-NOTE."
       Case "-": MFS.sNote = "M-NOTE-"
       Case "o": MFS.sNote = "M-NOTEo"
       Case "x": MFS.sNote = "M-NOTEx"
       Case "X": MFS.sNote = "M-NOTEX2"
       Case "e": MFS.sNote = "M-NOTEe"
       Case "E": MFS.sNote = "M-NOTEE2"
       Case " ": MFS.sNote = ""
       Case Else
            MFS.sNote = ""
    End Select
    
    '取得關鍵字 拍子  -=368acefz amt.iTempo
    Dim iLe As Integer
    Dim cn As String
    
    iLe = Tempo_c2line(Mid(cst, amt.iTempo, 1))
    Select Case iLe
        Case 0: MFS.sTempo = ""
                MFS.iTempo = 0
        Case 1: MFS.sTempo = "M-TEMPO1"
                MFS.iTempo = 1
        Case 2: MFS.sTempo = "M-TEMPO2"
                MFS.iTempo = 2
        Case 3: MFS.sTempo = "M-TEMPO3"
                MFS.iTempo = 3
        Case 4: MFS.sTempo = "M-TEMPO4"
                MFS.iTempo = 4
        Case Else
             MFS.sTempo = ""
                MFS.iTempo = 0
    End Select
    
    '取得關鍵字高低音行   .:! ,;?
    '    sScale As String
    cc = Mid(MData, amt.iScale, 1)
    Select Case cc
       Case ".": MFS.sScale = "M-SCALE1"
                 MFS.iScale = 1
       Case ":": MFS.sScale = "M-SCALE2"
                 MFS.iScale = 2
       Case "!": MFS.sScale = "M-SCALE3"
                 MFS.iScale = 3
       Case ",": MFS.sScale = "M-SCALEa"
                 MFS.iScale = -1
       Case ";": MFS.sScale = "M-SCALEb"
                 MFS.iScale = -2
       Case "?": MFS.sScale = "M-SCALEc"
                 MFS.iScale = -3
       Case " ": MFS.sScale = ""
                 MFS.iScale = 0
       Case Else
            MFS.sScale = ""
                 MFS.iScale = 0
    End Select
    

'    '取得關鍵字 -指法1
    cc = Mid(cst, amt.iFinge, 1)
    MFS.sFinge = ""
    Select Case cc
        Case "a" To "z", "A" To "Z", "-", "=", "@"
            MFS.sFinge = kFing(Asc(cc))
            Set MFS.mfPosition = posFing(Asc(cc)).copy
        Case " "
        Case "("
        Case "|"
        Case Else
        MFS.sFinge = ""
            
    End Select



'    '取得關鍵字 -指法2
    cc = Mid(cst, amt.iTowFinge, 1)
    MFS.sTowFinge = ""
    Select Case cc
        Case "a" To "z", "A" To "Z", "-", "=", "@"
            MFS.sTowFinge = kTowFing(Asc(cc))
            Set MFS.mfTowPosition = posTowFing(Asc(cc)).copy
        Case " "
        Case "("
        Case "|"
        Case Else
        MFS.sTowFinge = ""
            
    End Select

'    '取得升降字 -特疏字法
    cc = Mid(cst, amt.iTONE, 1)
    MFS.sTONE = cc


'    '取得關鍵字 -指法1
'     Select Case Mid(cst, amt.iFinge, 1)
'        Case "0": tmp_erhu_fing.fing1 = tmp_erhu_fing.fing1 + amt.f空弦 '空弦
'        Case "1": tmp_erhu_fing.fing1 = tmp_erhu_fing.fing1 + amt.f一指
'        Case "2": tmp_erhu_fing.fing1 = tmp_erhu_fing.fing1 + amt.f二指
'        Case "3": tmp_erhu_fing.fing1 = tmp_erhu_fing.fing1 + amt.f三指
'        Case "4": tmp_erhu_fing.fing1 = tmp_erhu_fing.fing1 + amt.f四指
'        Case "W", "w": tmp_erhu_fing.fing1 = tmp_erhu_fing.fing1 + amt.f揉弦
'
'        Case "E", "e": tmp_erhu_fing.Push = "拉"
'        Case "V", "v": tmp_erhu_fing.Push = "推"
'        Case "Q", "q": tmp_erhu_fing.InOut = "內"
'        Case "A", "a": tmp_erhu_fing.InOut = "外"
'        Case Else
'     End Select
'
'     '取得關鍵字 -指法2
'     Select Case Mid(cst, amt.iTowFinge, 1)
'        Case "0": tmp_erhu_fing.fing1 = tmp_erhu_fing.fing1 + amt.f空弦 '空弦
'        Case "1": tmp_erhu_fing.fing1 = tmp_erhu_fing.fing1 + amt.f一指
'        Case "2": tmp_erhu_fing.fing1 = tmp_erhu_fing.fing1 + amt.f二指
'        Case "3": tmp_erhu_fing.fing1 = tmp_erhu_fing.fing1 + amt.f三指
'        Case "4": tmp_erhu_fing.fing1 = tmp_erhu_fing.fing1 + amt.f四指
'        Case "W", "w": tmp_erhu_fing.fing1 = tmp_erhu_fing.fing1 + amt.f揉弦
'
'        Case "E", "e": tmp_erhu_fing.Push = "拉"
'        Case "V", "v": tmp_erhu_fing.Push = "推"
'        Case "Q", "q": tmp_erhu_fing.InOut = "內"
'        Case "A", "a": tmp_erhu_fing.InOut = "外"
'
'        Case Else
'     End Select



End Sub

Private Sub setGripPoints(ths_aPt As point)
'駐點轉換
    'Dim Grippt As New GripPoints
    Set Grippt.atPt = ths_aPt
    Set Grippt.gptLeft = New point
    Set Grippt.gptLeftDown = New point
    Set Grippt.gptMid = New point
    Set Grippt.gptMidUp = New point
    Set Grippt.gptRight = New point
    
    
    
    '在正中間
    Grippt.gptMid.x = ths_aPt.x + (amt.A_TEXT_WIDTH * gTextSize) / 2
    Grippt.gptMid.Y = ths_aPt.Y + (amt.bkMidHight * gTextSize)
    
    '在中上
    Grippt.gptMidUp.x = ths_aPt.x + (amt.A_TEXT_WIDTH * gTextSize) / 2
    Grippt.gptMidUp.Y = ths_aPt.Y + (amt.bkMidUpHight * gTextSize)
    
    '在左下
    Set Grippt.gptLeftDown = ths_aPt
    
    '在左中
    Grippt.gptLeft.x = ths_aPt.x
    Grippt.gptLeft.Y = ths_aPt.Y + (amt.bkMidHight * gTextSize)
    '在右中
    Grippt.gptRight.x = ths_aPt.x + (amt.A_TEXT_WIDTH * gTextSize)
    Grippt.gptRight.Y = ths_aPt.Y + (amt.bkMidHight * gTextSize)
    
End Sub
Public Function InsterEnt() As AcadBlockReference '插入音附
    Dim textObj As AcadText
    Dim blockRefObj  As AcadBlockReference
    Dim textString As String
    Dim aligPoint(0 To 2) As Double
    Dim height As Double
    Dim insertionPoint(0 To 2) As Double
    
    
    Dim varAttributes As Variant
    Dim entObj As AcadEntity
    Dim pickPnt As Variant
    Dim blkRefObj As AcadBlockReference
    Dim attVars As Variant
    Dim i As Integer
    Dim TheObj As AcadEntity
    Dim towPos As Double
    Dim cc As String
    
    
    Dim point1(0 To 2) As Double
    Dim point2(0 To 2) As Double
    point1(0) = 0: point1(1) = 0: point1(2) = 0
    point2(0) = 2: point2(1) = 0: point2(2) = 0

    insertionPoint(0) = Grippt.gptMid.x
    insertionPoint(1) = Grippt.gptMid.Y
    insertionPoint(2) = Grippt.gptMid.z
    If MFS.sNote <> "" Then
        Set InsterEnt = ThisDrawing.ModelSpace.InsertBlock(insertionPoint, MFS.sNote, gTextSize, gTextSize, gTextSize, 0)
        
        InsterEnt.Layer = MAINLAYER
    
        ' Create the text object in model space
        'cc = Mid(MData, amt.iNote, 1)
        'Set textObj = ThisDrawing.ModelSpace.AddText(cc, insertionPoint, gTextSize)
        'textObj.styleName = "musictxt"
        'textObj.Alignment = acAlignmentCenter
        'Call textObj.Move(point1, insertionPoint)
    
        
'        If MFS.sTempo <> "" Then
'            Call ThisDrawing.ModelSpace.InsertBlock(insertionPoint, MFS.sTempo, gTextSize, gTextSize, gTextSize, 0)
'        End If
        
        aligPoint(0) = insertionPoint(0)
        aligPoint(1) = insertionPoint(1)
        aligPoint(2) = insertionPoint(2)
        If MFS.sScale <> "" Then
            If MFS.iScale <= 0 Then
                If MFS.iTempo >= 1 Then    '下點音附 要移下去一些
                '有線時 要下去的距離
                    aligPoint(0) = insertionPoint(0)
                    aligPoint(1) = insertionPoint(1) - (amt.DROP_ADD_LINE_INTERVAL * gTextSize * MFS.iTempo)
                    aligPoint(2) = insertionPoint(2)
                End If
            End If
            Set TheObj = ThisDrawing.ModelSpace.InsertBlock(aligPoint, MFS.sScale, gTextSize, gTextSize, gTextSize, 0)
            TheObj.Layer = MAINLAYER
        End If
        If frmEDIT.chkOption1.Value = True Then
            '古箏指法
            If MFS.sFinge <> "" Then
                
                
                aligPoint(0) = insertionPoint(0) + (gTextSize * MFS.mfPosition.x)
                aligPoint(1) = insertionPoint(1) + (gTextSize * MFS.mfPosition.Y)
                aligPoint(2) = insertionPoint(2)
                
                
                Set TheObj = ThisDrawing.ModelSpace.InsertBlock(aligPoint, MFS.sFinge, gTextSize, gTextSize, gTextSize, 0)
                TheObj.Layer = MAINLAYER
            End If
            
            towPos = 0
            If MFS.sTowFinge <> "" Then
                
                aligPoint(0) = insertionPoint(0) + (gTextSize * MFS.mfTowPosition.x)
                aligPoint(1) = insertionPoint(1) + (gTextSize * MFS.mfTowPosition.Y)
                aligPoint(2) = insertionPoint(2)
                
                '減 拍線
                towPos = (amt.DROP_ADD_LINE_INTERVAL * MFS.iTempo)
                If MFS.iScale < 0 Then
                    '減 下點音附
                    towPos = towPos + (amt.DROP_INTERVAL * (-MFS.iScale))
                End If
                If towPos > 0.3 Then
                    aligPoint(1) = aligPoint(1) - (towPos * gTextSize)
                Else
                    aligPoint(1) = aligPoint(1) - (0.3 * gTextSize)
                End If
                
                Set TheObj = ThisDrawing.ModelSpace.InsertBlock(aligPoint, MFS.sTowFinge, gTextSize, gTextSize, gTextSize, 0)
                TheObj.Layer = MAINLAYER
            End If
            
        '*****畫特疏指法*******************************************
            If MFS.sTONE = "*" Then
                InsertMusicStar insertionPoint, gTextSize
            End If
            
        ElseIf frmEDIT.chkOption2.Value = True Then
            '二胡指法用
            
            
        End If
        
        
    
        '寫入屬性
        '本段代碼從選中的圖塊中獲取屬性值，並對其修改

    If InsterEnt Is Nothing Then
        Exit Function
    End If

    '如果選擇的是塊引用，將其賦給塊引用對象
    Set blkRefObj = InsterEnt
    '判斷該塊引用是否含有屬性值
    If Not blkRefObj.HasAttributes Then
        'MsgBox "你選擇的圖塊沒有塊屬性，程序將退出！"
        '如果不含由屬性值退出
        Exit Function
    End If
    '獲取塊引用中的塊屬性對象
    attVars = blkRefObj.GetAttributes
    '對塊屬性對象進行遍歷
    For i = LBound(attVars) To UBound(attVars)
'       MsgBox "第" & I + 1 & "屬性對象的屬性值分別如下：" & Chr(13) & Chr(13) & _
'               "屬性標籤為：" & attVars(I).TagString & Chr(13) & _
'               "屬性值為：" & attVars(I).textString
        
        '將塊屬性的標籤和值進行修改
    Select Case attVars(i).TagString
        Case "ITONE"
            attVars(i).textString = Mid(MData, amt.iTONE, 1)
        Case "IFINGE"
            attVars(i).textString = Mid(MData, amt.iFinge, 1)
        Case "ISCALE"
            attVars(i).textString = Mid(MData, amt.iScale, 1)
        Case "INOTE"
            attVars(i).textString = Mid(MData, amt.iNote, 1)
        Case "ITEMPO"
            '取得關鍵字 拍子  -=368acefz amt.iTempo
            attVars(i).textString = Mid(MData, amt.iTempo, 1)
        Case "ITOWFINGE"
            attVars(i).textString = Mid(MData, amt.iTowFinge, 1)
        Case "ISLUR"
            attVars(i).textString = Mid(MData, amt.iSlur, 1)
        Case "IRE1"
        Case Else
            
    End Select

    Next
    End If
End Function

Private Function Tempo_c2line(theC As String) As Integer
'拍子的 字元 轉 成幾條線

    
    Dim tempo_hj As String
    Dim tempo_val As Variant
    Dim tempo_ll As Variant
    
    tempo_hj = " -2=45368aAcCfFgGzZ"
    tempo_ll = Array(1, 2, 2, 4, 4, 5, 3, 6, 8, 10, 10, 12, 12, 15, 15, 16, 32, 32)
    tempo_val = Array(0, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3, 4, 4, 4, 4, 4, 4, 4, 4)
    

    Dim ii As Integer
    Dim cn As String
    If theC <> "" Then
    
    For ii = 0 To Len(tempo_hj) - 1
        cn = Mid(tempo_hj, ii + 1, 1)
        If cn = Mid(theC, 1, 1) Then
            Tempo_c2line = tempo_val(ii)
            Exit Function

        End If
    Next ii
    
    End If
    Tempo_c2line = 0
End Function

Private Function Tempo_c2number(cc As String) As Double
'拍子的 字元 轉成 時值

    
    Dim tempo_hj As String
    Dim tempo_val As Variant
    Dim tempo_ll As Variant
    
    tempo_hj = " -2=45368aAcCfFgGzZ"
    tempo_ll = Array(1, 2, 2, 4, 4, 5, 3, 6, 8, 10, 10, 12, 12, 15, 15, 16, 16, 32, 32)
    tempo_val = Array(0, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3, 4, 4, 4, 4, 4, 4, 4, 4)
    

    Dim ii As Integer
    Dim cn As String
    For ii = 0 To Len(tempo_hj) - 1
        cn = Mid(tempo_hj, ii + 1, 1)
        If cn = Mid(theC, 1, 1) Then
            Tempo_c2line = tempo_ll(ii)
            Exit Function
        End If
    Next ii
    Tempo_c2line = 0
End Function




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
        If the_Ids(i).ObjectName = "AcDbBlockReference" Then
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
    'Dim MTobj As NnmText
    Dim pt As Variant
    
    For i = 0 To UBound(the_Ids, 1)
        Set ent = the_Ids(i)
        If (Mid(ent.Name, 1, 6) = "M-NOTE") Then
            pt = ent.insertionPoint()
            TheArray(coI) = pt(0)
            Set TheArrayEnt(coI) = ent
            coI = coI + 1
            
            
            Dim minExt As Variant
            Dim maxExt As Variant
    
            ent.GetBoundingBox minExt, maxExt
            'Set MTobj = ent
            'minExt = MTobj.GripLeft()
            
        End If
    Next
    
    
    
    '排序開始
    SelectionSortMaxTow TheArray, TheArrayEnt
    
    
    Dim stPt As point
    Dim endPt As point
    Dim tempo_num(0 To 20) As Integer
    Dim ptTheArr(0 To 20) As point
    Dim cIT As String
    Dim insertionPoint As Variant
    Dim entArr As AcadEntity
    
    Set stPt = New point
    Set endPt = New point
    
    '取得拍子的線條數量及 物件位置
    For i = 0 To UBound(TheArray)
        Set entArr = TheArrayEnt(i)
        cIT = GetAttribTextString(entArr, "ITEMPO")
        
        tempo_num(i) = Tempo_c2line(cIT)
        insertionPoint = TheArrayEnt(i).insertionPoint
        Set ptTheArr(i) = New point
        ptTheArr(i).c insertionPoint(0), insertionPoint(1), insertionPoint(2)
    Next
    
    

    Dim x As Integer
'    for (int x =1; x<=the_lineOne; x++)
    For x = 1 To 4
        For i = 0 To UBound(TheArray)
            If tempo_num(i) >= x Then
                If stPt.x = 0 Then
                    Set stPt = ptTheArr(i).copy '取得第一個位置
                Else
                    Set endPt = ptTheArr(i).copy '取得第 後面的位置
                End If
                If i = UBound(TheArray) Then  '如是最後一個就畫出來
                    If endPt.x = 0 Then
                        Set endPt = stPt.copy
                    End If
                    line_link x - 1, stPt, endPt
                    stPt.c 0, 0, 0
                    endPt.c 0, 0, 0
                End If
            Else
                If stPt.x <> 0 Then
                    If endPt.x = 0 Then
                        Set endPt = stPt.copy
                    End If
                    line_link x - 1, stPt, endPt
                    stPt.c 0, 0, 0
                    endPt.c 0, 0, 0
                    
                End If
            
            End If
        Next
    Next
            



EndJoin:
End Function

'
'//畫拍子的線
'// num_col   要畫幾個線
'// l_pt, r_pt  左右鏈結的中點

Sub line_link(num_col As Integer, lPt As point, rPt As point)

    If num_col < 0 Then
        Exit Sub
    End If
    
    Dim xt As Double
    Dim yt As Double
    
    Dim plineObj As AcadPolyline
    Dim points(0 To 5) As Double
    
    ' Define the 2D polyline points
    points(0) = lPt.x - amt.A_TEXT_WIDTH / 2 * gTextSize
    points(1) = lPt.Y - (amt.bkInster_Line_H * gTextSize) - (num_col * amt.LINE_JUMP * gTextSize)
    points(2) = lPt.z
    
    
    points(3) = rPt.x + amt.A_TEXT_WIDTH / 2 * gTextSize
    points(4) = points(1)
    points(5) = rPt.z
        
    ' Create a lightweight Polyline object in model space
    Set plineObj = ThisDrawing.ModelSpace.AddPolyline(points)
    plineObj.SetWidth 0, gTextSize * amt.LINE_THICKNESS, gTextSize * amt.LINE_THICKNESS
    plineObj.Layer = MAINLAYER
    'Call ThisDrawing.ModelSpace.AddPolyline .InsertBlock(insertionPoint, MFS.sFinge, gTextSize, gTextSize, gTextSize, 0)

End Sub

Private Function InsertMusicStar(insertionPoint() As Double, size As Double)
'插入古箏 前音 琵音 * 號
'insertionPoint  是本音符的位置

    
    Dim textObj As AcadText
    Dim blockRefObj  As AcadBlockReference
    Dim textString As String
    Dim alignmentPoint(0 To 2) As Double
    Dim midDownPt(0 To 2) As Double
    Dim height As Double
        
    
    
    Dim xStar As Double
    Dim yStar As Double
    Dim ssize As Double
    xStar = -1.06 '星指法的 X 上增量
    yStar = 1.3 '星指法的 Y 上增量
    ssize = 0.65 '小字的比例
    

    alignmentPoint(0) = insertionPoint(0) + (size * xStar)
    alignmentPoint(1) = insertionPoint(1) + (size * yStar)
    alignmentPoint(2) = insertionPoint(2)

    
    midDownPt(0) = alignmentPoint(0) + (amt.A_TEXT_WIDTH / 2 * size * ssize)
    midDownPt(1) = alignmentPoint(1)
    midDownPt(2) = alignmentPoint(2)


'*******畫星字*****************
    
    Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(midDownPt, "M-NOTESTAR", size, size, size, 0)
    blockRefObj.Layer = "裝飾符號"
    
'*******畫拍線*****************
    Dim sMGB As New MusicBlockGraphics
    Dim ppt As New point
    Dim lPt As New point
    Dim rPt As New point
    
    sMGB.setDataText ppt, "", size * ssize
    
        
    lPt.x = midDownPt(0)
    lPt.Y = midDownPt(1) + (0.15 * size * ssize)
    lPt.z = midDownPt(2)
    
    rPt.x = lPt.x
    rPt.Y = lPt.Y
    rPt.z = lPt.z

    sMGB.line_link 1, lPt, rPt
    sMGB.line_link 2, lPt, rPt
    
'*******畫弧*****************
      
    Dim arcObj As AcadArc
    Dim centerPoint(0 To 2) As Double
    Dim radius As Double
    Dim startAngleInDegree As Double
    Dim endAngleInDegree As Double
    
    ' Define the circle
    
    radius = 0.38 * size * ssize
    centerPoint(0) = midDownPt(0) + radius
    centerPoint(1) = midDownPt(1) - 0.3 * size * ssize
    centerPoint(2) = midDownPt(2)
    startAngleInDegree = 180#
    endAngleInDegree = 275#
    
    ' Convert the angles in degrees to angles in radians
    Dim startAngleInRadian As Double
    Dim endAngleInRadian As Double
    startAngleInRadian = startAngleInDegree * 3.141592 / 180#
    endAngleInRadian = endAngleInDegree * 3.141592 / 180#
    
    ' Create the arc object in model space
    Set arcObj = ThisDrawing.ModelSpace.AddArc(centerPoint, radius, startAngleInRadian, endAngleInRadian)
    arcObj.Layer = "裝飾符號"
    
End Function


Private Sub linkGrapTempoMaray(the_lineOne As Integer)
    '畫拍子聯結
    'pW
    'lineOne 是要看，現在是第幾條線
    
'    MusicData tmp_md = this->m_MusicData ;
'
'    string cst;
'    cst += "MT::linkGrapTempoMaray";
'    cst += tmp_md.strNnmTextData ;
'
'    LOG1W(cst.c_str(),(long )__LINE__);
'
'
'    AcGePoint3d stPt, endPt;
'    stPt = AcGePoint3d(0,0,0);
'    endPt = AcGePoint3d(0,0,0);
'
'    MusicText *tmp_objMTXT;
'
'    deque<structMT> sDeque;
'    deque<structMT>::iterator  tmp_iterSDeque;
'
'    //mapData::iterator pos;
'
'    structMT tmp_sMT;
'    //Acad::ErrorStatus es;
'
'    tmp_sMT.pt = tmp_md.point_;
'    tmp_sMT.m_tempo = tmp_md.tempo;
'
'    sDeque.push_back(tmp_sMT);
'
'
'    if (tmp_md.LlinkobjId != NULL)
'    {
'        return;
'    }
'    else if (tmp_md.RlinkobjId != NULL)
'    {
'        while(tmp_md.RlinkobjId != NULL )
'        {
'
'            if (acdbOpenObject((AcDbObject*&)tmp_objMTXT, tmp_md.RlinkobjId,
'                AcDb::kForWrite) == Acad::eOk)
'            {
'
'                tmp_md = tmp_objMTXT->m_MusicData ;
'
'                tmp_sMT.pt = tmp_md.point_;
'                tmp_sMT.m_tempo = tmp_md.tempo ;
'
'                sDeque.push_back(tmp_sMT);
'                tmp_objMTXT->close();
'            }
'            Else
'            {
'
'                cst = "";
'                cst += "未取得 id- pu438";
'                LOG1W(cst.c_str(),(long )__LINE__);
'
'                acutPrintf( "\n 未取得 id- pu438" );
'                tmp_md.RlinkobjId = NULL;
'
'            }
'        }
'    }
'
'    LOG1W(cst.c_str(),(long )__LINE__);
'
'    for (int x =1; x<=the_lineOne; x++)
'    {
'        for( tmp_iterSDeque = sDeque.begin(); tmp_iterSDeque != sDeque.end(); tmp_iterSDeque++)
'        {
'            if(((*tmp_iterSDeque).m_tempo) >= x)
'            {
'                if(stPt.x == 0)
'                {
'                    stPt = (*tmp_iterSDeque).pt;    //取得第一個位置
'                }
'                Else
'                {
'                    endPt = (*tmp_iterSDeque).pt;   //取得第後面的位置
'                }
'                if(tmp_iterSDeque == (sDeque.end()-1))//如是最後一個就畫出來
'                {
'                    if (endPt.x == 0)
'                    {
'                        endPt = stPt;
'                    }
'                    m_pGObject->line_link(pW, x-1, stPt, endPt);
'                    stPt = AcGePoint3d(0,0,0);
'                    endPt = AcGePoint3d(0,0,0);
'                }
'
'            }
'            Else
'            {
'                if(stPt.x != 0)
'                {
'                    if (endPt.x == 0)
'                    {
'                        endPt = stPt;
'                    }
'                    m_pGObject->line_link(pW, x-1, stPt, endPt);
'                    stPt = AcGePoint3d(0,0,0);
'                    endPt = AcGePoint3d(0,0,0);
'                }
'            }
'        }
'    }
'
'
'}

End Sub


Private Sub Class_Initialize()

    
    Dim kNUM1 As String, kNUM2 As String, kNUM3 As String, kNUM4 As String
    Dim kNUM5 As String, kNUM6 As String, kNUM7 As String, kNUM0 As String
    Dim k延長一拍 As String
    Dim kLIN1 As String: Dim kLIN2 As String: Dim kLIN3 As String
    Dim kUP1 As String
    Dim kUP2 As String
    Dim kDOW1 As String
    Dim kDOW2 As String
    Dim kAdot As String
    Dim kBarLine As String
    'Dim kFing(30 To 255) As String
    'Dim posFing(0 To 255) As point
    Dim posLinDown As Variant
    
    posLinDown = Array(0, 12, 26, 40)   '這是看低音下點要移多少下去，是當200字型時計算
    
    kNUM1 = ChrW(&H4E52)
    kNUM2 = ChrW(&H4E53)
    kNUM3 = ChrW(&H4E56)
    kNUM4 = ChrW(&H4E58)
    kNUM5 = ChrW(&H4E59)
    kNUM6 = ChrW(&H4E5C)
    kNUM7 = ChrW(&H4E5D)
    kNUM0 = ChrW(&H5D4C)
    k延長一拍 = ChrW(&H5D1F)
    kLIN1 = ChrW(&H5D39): kLIN2 = ChrW(&H5D3A): kLIN3 = ChrW(&H5D3C)
    kUP1 = ChrW(&H66E0): kUP2 = ChrW(&H66E3): kDOW1 = ChrW(&H66E6): kDOW2 = ChrW(&H66E8)
    kAdot = ChrW(&H5D3D)            '符點音符
    kBarLine = ChrW(&H601D)
    kFing(Asc("z")) = "MT" & Format(Asc("z"), "0000") 'ASC_z 122 ChrW(&H6416)  ' z 大指
    kFing(Asc("x")) = "MT" & Format(Asc("x"), "0000") 'ASC_x 120 ChrW(&H641A)  ' x 食指
    kFing(Asc("c")) = "MT" & Format(Asc("c"), "0000") 'ASC_c 99 ChrW(&H641C)  ' c 中指
    kFing(Asc("v")) = "MT" & Format(Asc("v"), "0000") 'ASC_v 118 ChrW(&H641F)  ' v 無名指
    kFing(Asc("a")) = "MT" & Format(Asc("a"), "0000") 'ASC_a 97 ChrW(&H6421)  ' a 八度
    kFing(Asc("s")) = "MT" & Format(Asc("s"), "0000") 'ASC_s 115 ChrW(&H66FC)  ' s 王字指法
    kFing(Asc("d")) = "MT" & Format(Asc("d"), "0000") 'ASC_d 100 ChrW(&H6700)  ' d 搖指
    'kFing(Asc("p")) = "MT" & Format(Asc("p"), "0000") 'ASC_p 112 ChrW(&H6381)  ' p 卜指法
    kFing(Asc("Z")) = "MT" & Format(Asc("Z"), "0000") 'ASC_Z 90 ChrW(&H6418)  ' 大寫Z 反大指法
    kFing(Asc("A")) = "MT" & Format(Asc("A"), "0000") 'ASC_A 65 ChrW(&H6413)  ' 大寫A 反八度
    kFing(Asc("-")) = "MT" & Format(Asc("-"), "0000") 'ASC_- 45 ChrW(&H66F7)  ' - 間一旋
    kFing(Asc("=")) = "MT" & Format(Asc("="), "0000") 'ASC_= 61 ChrW(&H66F8)  ' = 間二旋
    'kFing(Asc("t")) = "MT" & Format(Asc("t"), "0000") 'ASC_t 116 ChrW(&H670B)  ' t 打圓 托勾
    kFing(Asc("w")) = "MT" & Format(Asc("w"), "0000") 'ASC_w 119 ChrW(&H670B)  ' w 柔音
    kFing(Asc("@")) = "MT" & Format(Asc("@"), "0000") 'ASC_w 119 ChrW(&H670B)  '
    
    
    kTowFing(Asc("z")) = "MT" & Format(Asc("z"), "0000") 'ASC_z 122 ChrW(&H6416)  ' z 大指
    kTowFing(Asc("x")) = "MT" & Format(Asc("x"), "0000") 'ASC_x 120 ChrW(&H641A)  ' x 食指
    kTowFing(Asc("c")) = "MT" & Format(Asc("c"), "0000") 'ASC_c 99 ChrW(&H641C)  ' c 中指
    kTowFing(Asc("v")) = "MT" & Format(Asc("v"), "0000") 'ASC_v 118 ChrW(&H641F)  ' v 無名指
    kTowFing(Asc("a")) = "MT" & Format(Asc("a"), "0000") 'ASC_a 97 ChrW(&H6421)  ' a 八度
    kTowFing(Asc("s")) = "MT" & Format(Asc("s"), "0000") 'ASC_s 115 ChrW(&H66FC)  ' s 王字指法
    kTowFing(Asc("d")) = "MT" & Format(Asc("d"), "0000") 'ASC_d 100 ChrW(&H6700)  ' d 搖指
    'kTowFing(Asc("p")) = "MT" & Format(Asc("p"), "0000") 'ASC_p 112 ChrW(&H6381)  ' p 卜指法
    kTowFing(Asc("Z")) = "MT" & Format(Asc("Z"), "0000") 'ASC_Z 90 ChrW(&H6418)  ' 大寫Z 反大指法
    kTowFing(Asc("A")) = "MT" & Format(Asc("A"), "0000") 'ASC_A 65 ChrW(&H6413)  ' 大寫A 反八度
    kTowFing(Asc("-")) = "MT" & Format(Asc("-"), "0000") 'ASC_- 45 ChrW(&H66F7)  ' - 間一旋
    kTowFing(Asc("=")) = "MT" & Format(Asc("="), "0000") 'ASC_= 61 ChrW(&H66F8)  ' = 間二旋
    kTowFing(Asc("t")) = "MT" & Format(Asc("t"), "0000") 'ASC_t 116 ChrW(&H670B)  ' t 打圓 托勾
    kTowFing(Asc("w")) = "MT" & Format(Asc("w"), "0000") 'ASC_w 119 ChrW(&H670B)  ' w 柔音
    kTowFing(Asc("m")) = "MT" & Format(Asc("m"), "0000") 'ASC_w 119 ChrW(&H670B)  '
    kTowFing(Asc("e")) = "MT" & Format(Asc("e"), "0000") 'ASC_w 119 ChrW(&H670B)  '
    kTowFing(Asc("r")) = "MT" & Format(Asc("r"), "0000") 'ASC_w 119 ChrW(&H670B)  '
    kTowFing(Asc("k")) = "MT" & Format(Asc("k"), "0000") 'ASC_w 119 ChrW(&H670B)  '
    kTowFing(Asc("n")) = "MT" & Format(Asc("n"), "0000") 'ASC_t 116 ChrW(&H670B)  ' n 按音
    kTowFing(Asc("f")) = "MT" & Format(Asc("f"), "0000") 'ASC_t 116 ChrW(&H670B)  ' f 吟音
    
    '要上下移多少
    Dim ip As Integer
    For ip = 0 To 255
        Set posFing(ip) = New point
        Set posTowFing(ip) = New point
        
    Next
    
    posFing(Asc("z")).c -0.363, 1.57, 0 ' z 大指
    posFing(Asc("x")).c -0.363, 1.65, 0 ' x 食指
    posFing(Asc("c")).c -0.363, 1.57, 0 ' c 中指
    posFing(Asc("v")).c -0.363, 1.57, 0 ' v 無名指
    posFing(Asc("a")).c -0.363, 1.57, 0 ' a 八度
    posFing(Asc("s")).c -0.363, 1.57, 0 ' s 王字指法
    posFing(Asc("d")).c -0.363, -0.85, 0 ' d 搖指
    posFing(Asc("p")).c -0.363, -0.01, 0 ' p 卜指法
    posFing(Asc("Z")).c -0.363, 1.57, 0 ' 大寫Z 反大指法
    posFing(Asc("A")).c -0.363, 1.57, 0 ' 大寫A 反八度
    posFing(Asc("-")).c -0.363, 1.57, 0 ' - 間一旋
    posFing(Asc("=")).c -0.363, 1.57, 0 ' = 間二旋
    posFing(Asc("t")).c -0.363, 1.6, 0 ' t 打圓 托勾
    posFing(Asc("@")).c -0.363, 1.57, 0 ' w
    
    
    posTowFing(Asc("z")).c -0.363, -0.7, 0 ' z 大指
    posTowFing(Asc("x")).c -0.363, -0.7, 0 ' x 食指
    posTowFing(Asc("c")).c -0.363, -0.7, 0 ' c 中指
    posTowFing(Asc("v")).c -0.363, -0.7, 0 ' v 無名指
    posTowFing(Asc("a")).c -0.363, -0.7, 0 ' a 八度
    posTowFing(Asc("s")).c -0.363, -0.7, 0 ' s 王字指法
    'posTowFing(Asc("d")).c -0.363, -0.85, 0 ' d 搖指
    'posTowFing(Asc("p")).c -0.363, -0.01, 0 ' p 卜指法
    posTowFing(Asc("Z")).c -0.363, -0.7, 0 ' 大寫Z 反大指法
    posTowFing(Asc("A")).c -0.363, -0.7, 0 ' 大寫A 反八度
    posTowFing(Asc("-")).c -0.363, -0.7, 0 ' - 間一旋
    posTowFing(Asc("=")).c -0.363, -0.7, 0 ' = 間二旋
    posTowFing(Asc("t")).c -0.363, -0.7, 0 ' t 打圓 托勾
    posTowFing(Asc("w")).c -0.363, -0.7, 0 ' m
    posTowFing(Asc("m")).c -0.363, -0.7, 0 ' w
    posTowFing(Asc("e")).c -0.363, -0.7, 0 ' w
    posTowFing(Asc("r")).c -0.363, -0.7, 0 ' w
    posTowFing(Asc("k")).c -0.363, -0.7, 0 ' w
    posTowFing(Asc("n")).c -0.363, -0.7, 0 ' w
    posTowFing(Asc("f")).c -0.363, -0.7, 0 ' w
    
    
End Sub
