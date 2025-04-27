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
Dim MData As MusicItem


Dim MFS() As MusicFing

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

Property Let DataTest(dtest As MusicItem)
    MData = dtest
End Property
Property Get DataTest() As MusicItem
    DataTest = MData
End Property
Property Get copy() As MusicBlockGraphics
    Dim copy_mbg As New MusicBlockGraphics
    copy_mbg.setDataText gAtPt, MData, gTextSize
    Set copy = copy_mbg
End Property
Function checkBlockName(Name As String) As Boolean
    Dim ck As Boolean
    Dim i As Integer
    ck = False
    For i = 0 To ThisDrawing.Blocks.Count - 1
        If ThisDrawing.Blocks.item(i).Name = Name Then
            ck = True
            Exit For
        End If
    Next
    checkBlockName = ck
End Function

Static Function getGlyphs()

End Function


Sub setDataText(position As point, musItem As MusicItem, TextSize As Double)

    'position 插入的位置
    'cst  文字資料=指法1 八度音 主音 拍子 指法2
    '          共5筆 每筆一個字元 所以共5個字元
    'textSize  字形大小
    
    Set MData = musItem
    
    gAtPt.x = position.x
    gAtPt.y = position.y
    gAtPt.Z = position.Z
    
    gTextSize = TextSize
    setGripPoints gAtPt
    
    
    
'Public Type MusicFing
'    sTONE As String        ' 升降記號 行
'    sFinge As String   '這是指法行     _+)(*&
'    sScale As String   '這是高低音行   .:! , ?
'    sNote As String   '這為主行       1234567.|l
'    sTempo As String   '這是拍子       -=368acefz
'    sTowFinge As String   '這是指法行第二行  _+)(*&
'    sSlur As String
'End Type
    Dim currenCount As Integer
    Dim sarr() As String
    Dim reg As New RegExp
    Dim cc As String
    Dim sNote As String
    Dim iTempo As Integer
    Dim sTempo As Integer
    Dim iScale As Integer
    Dim sScale As Integer
    Dim i As Integer
    
    cc = ""
    
    '取得關鍵字  主音 AMT.iNote
    Dim noteListCode As New iArray
    Dim noteListType As iArray
  
            
        noteListCode.PushArray Array("0 M-NOTE0", _
            "1 M-NOTE1", _
            "2 M-NOTE2", _
            "3 M-NOTE3", _
            "4 M-NOTE4", _
            "5 M-NOTE5", _
            "6 M-NOTE6", _
            "7 M-NOTE7", _
            ". M-NOTE.", _
            "- M-NOTE-", _
            "o M-NOTEo", _
            "x M-NOTEx", _
            "X M-NOTEX2", _
            "e M-NOTEe", _
            "E M-NOTEE2")
            
    Dim iLe As Integer
    Dim cn As String
    Dim tempoListCode As New iArray

    tempoListCode.PushArray Array( _
            "1 M-TEMPO1", _
            "2 M-TEMPO2", _
            "3 M-TEMPO3", _
            "4 M-TEMPO4")
            
    
    Dim scaleListCode As New iArray
    scaleListCode.PushArray Array( _
            ". M-OCTAVE1 1", _
            ": M-OCTAVE2 2", _
            "! M-OCTAVE3 3", _
            ", M-OCTAVEa -1", _
            "; M-OCTAVEb -2", _
            "? M-OCTAVEc -3")
            
    '取得升降字 -特疏字法
    Dim toneListCode As New iArray
    toneListCode.PushArray Array( _
            "# MT-SHARP", _
            "b MT-FLAT", _
            "o MT-NATURAL")
    
    ReDim MFS(MData.notes.Count - 1)
    For currenCount = 0 To MData.notes.Count - 1
        
        cc = MData.notes(currenCount).mnote
        For i = 0 To noteListCode.Count - 1
            sarr = Split(noteListCode(i), " ")
            If cc = sarr(0) Then
                MFS(currenCount).sNote = sarr(1)
                Exit For
            Else
                MFS(currenCount).sNote = ""
            End If
        Next
    
        
        '取得關鍵字 拍子  -=368acefz amt.iTempo
        iLe = Tempo_c2line(MData.notes(currenCount).mtempo)
        cc = MData.notes(currenCount).mtempo
        For i = 0 To tempoListCode.Count - 1
            sarr = Split(tempoListCode(i), " ")
            If iLe = sarr(0) Then
                MFS(currenCount).sTempo = sarr(1)
                MFS(currenCount).iTempo = iLe
                Exit For
            Else
                MFS(currenCount).sTempo = cc
                MFS(currenCount).iTempo = iLe
            End If
        Next
        
        
        '取得關鍵字高低音行   .:! , ?
        cc = MData.notes(currenCount).mscale
        For i = 0 To scaleListCode.Count - 1
            sarr = Split(scaleListCode(i), " ")
            If cc = sarr(0) Then
                MFS(currenCount).sScale = sarr(1)
                MFS(currenCount).iScale = sarr(2)
                Exit For
            Else
                MFS(currenCount).sScale = ""
                MFS(currenCount).iScale = 0
            End If
        Next
    
    
        '取得關鍵字 -指法1
        MFS(currenCount).sFinge = ""
        cc = MData.notes(currenCount).mfingering
        Select Case cc
            Case "a" To "z", "A" To "Z", "-", "=", "@"
                MFS(currenCount).sFinge = kFing(Asc(cc))
                Set MFS(currenCount).mfPosition = posFing(Asc(cc)).clone
            Case Else
            MFS(currenCount).sFinge = ""
            Set MFS(currenCount).mfPosition = New point
        End Select
    
    
    
    '    '取得關鍵字 -指法2
        cc = MData.notes(currenCount).mtow_fingering
        MFS(currenCount).sTowFinge = ""
        Select Case cc
            Case "a" To "z", "A" To "Z", "-", "=", "@"
                MFS(currenCount).sTowFinge = kTowFing(Asc(cc))
                Set MFS(currenCount).mfTowPosition = posTowFing(Asc(cc)).clone
            Case Else
            MFS(currenCount).sTowFinge = ""
            Set MFS(currenCount).mfTowPosition = New point
        End Select
    
        '取得升降字 -特疏字法
        cc = MData.notes(currenCount).mtone
        For i = 0 To toneListCode.Count - 1
            sarr = Split(toneListCode(i), " ")
            If cc = sarr(0) Then
                MFS(currenCount).sTONE = sarr(1)
                Exit For
            Else
                MFS(currenCount).sTONE = cc
            End If
        Next
    Next

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
    Grippt.gptMid.y = ths_aPt.y + (amt.bkMidHight * gTextSize)
    
    '在中上
    Grippt.gptMidUp.x = ths_aPt.x + (amt.A_TEXT_WIDTH * gTextSize) / 2
    Grippt.gptMidUp.y = ths_aPt.y + (amt.bkMidUpHight * gTextSize)
    
    '在左下
    Set Grippt.gptLeftDown = ths_aPt
    
    '在左中
    Grippt.gptLeft.x = ths_aPt.x
    Grippt.gptLeft.y = ths_aPt.y + (amt.bkMidHight * gTextSize)
    '在右中
    Grippt.gptRight.x = ths_aPt.x + (amt.A_TEXT_WIDTH * gTextSize)
    Grippt.gptRight.y = ths_aPt.y + (amt.bkMidHight * gTextSize)
    
End Sub
Public Function drawSlur(mt_slur_left As MusicBlockGraphics, mt_slur_right As MusicBlockGraphics)
    '繪圓滑線
        Dim plineSlur As AcadLWPolyline
        Dim points(0 To 7) As Double
        Dim points_s(0 To 5) As Double
        Dim lenght As Double
        ' Find the bulge of the third segment
        Dim color As New AcadAcCmColor
        
        
        Dim islurAddX As Double
        Dim islurAddY As Double
        Dim islurBx As Double
        Dim islurBy As Double
        
        islurAddX = 0.6245 * gTextSize
        islurAddY = 0.4333 * gTextSize
        
        islurBy = 0.45 * gTextSize
        
        '看是否弧線的距離太近
        
        lenght = Abs(mt_slur_left.Grip.gptMid.x - mt_slur_right.Grip.gptMid.x)
        If lenght >= islurAddX * 2 Then
        
            points(0) = mt_slur_left.Grip.gptMid.x
            points(1) = mt_slur_left.Grip.gptMidUp.y + islurBy
            points(2) = points(0) + islurAddX
            points(3) = points(1) + islurAddY
            points(4) = points(2) + lenght - (islurAddX * 2)
            points(5) = points(3)
            points(6) = points(4) + islurAddX
            points(7) = points(1)
            ' Create a lightweight Polyline object in model space
           Set plineSlur = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
    
           
           ' Change the bulge of the third segment
           plineSlur.SetBulge 0, (-0.5858 / 1.4142)
           plineSlur.SetBulge 2, (-0.5858 / 1.4142)
           plineSlur.setWidth 0, 0, 0.15
           plineSlur.setWidth 1, 0.15, 0.15
           plineSlur.setWidth 2, 0.15, 0
                           
        Else
        '距離太近的畫線
            points_s(0) = mt_slur_left.Grip.gptMid.x
            points_s(1) = mt_slur_left.Grip.gptMidUp.y + islurBy
            points_s(2) = points_s(0) + (lenght / 2)
            points_s(3) = points_s(1) + islurAddY
            points_s(4) = points_s(2) + (lenght / 2)
            points_s(5) = points_s(1)
            ' Create a lightweight Polyline object in model space
           Set plineSlur = ThisDrawing.ModelSpace.AddLightWeightPolyline(points_s)
           ' Change the bulge of the third segment
           plineSlur.SetBulge 0, (-0.5858 / 1.4142)
           plineSlur.SetBulge 1, (-0.5858 / 1.4142)
           
           plineSlur.setWidth 0, 0, 0.15
           plineSlur.setWidth 1, 0.15, 0
        End If
    plineSlur.Layer = "fige"
    
    color.ColorIndex = 3
    plineSlur.TrueColor = color
    plineSlur.Update
                        
    Set drawSlur = plineSlur
End Function
Public Function drawSlurToEnd(mt_slur_left As MusicBlockGraphics, endX As Double)
    '繪圓滑線 到尾部
        Dim plineSlur As AcadLWPolyline
        Dim points(0 To 5) As Double
        Dim sLenght As Double
        ' Find the bulge of the third segment
        Dim color As New AcadAcCmColor
        
        
        Dim islurAddX As Double
        Dim islurAddY As Double
        Dim islurBx As Double
        Dim islurBy As Double
        
        islurAddX = 0.6245 * gTextSize
        islurAddY = 0.4333 * gTextSize
        
        islurBy = 0.45 * gTextSize
        
        '看是否弧線的距離太近
        
        sLenght = Abs(mt_slur_left.Grip.gptMid.x - endX) - islurAddX
        If sLenght < 1 Then
            sLenght = 1
        End If
        
        points(0) = mt_slur_left.Grip.gptMid.x
        points(1) = mt_slur_left.Grip.gptMidUp.y + islurBy
        points(2) = points(0) + islurAddX
        points(3) = points(1) + islurAddY
        points(4) = points(0) + sLenght + islurAddX
        points(5) = points(3)
        ' Create a lightweight Polyline object in model space
        Set plineSlur = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
        
        
        ' Change the bulge of the third segment
        plineSlur.SetBulge 0, (-0.5858 / 1.4142)
        plineSlur.setWidth 0, 0, 0.15
        plineSlur.setWidth 1, 0.15, 0.15
                           
        plineSlur.Layer = "fige"
        
        color.ColorIndex = 3
        plineSlur.TrueColor = color
        plineSlur.Update
                        
        Set drawSlurToEnd = plineSlur
End Function

Public Function drawSlurStarTo(starX As Double, mt_slur_right As MusicBlockGraphics)
    '繪圓滑線 從頭 到 元素
        Dim plineSlur As AcadLWPolyline
        Dim points(0 To 5) As Double
        Dim sLenght As Double
        ' Find the bulge of the third segment
        Dim color As New AcadAcCmColor
        
        
        Dim islurAddX As Double
        Dim islurAddY As Double
        Dim islurBx As Double
        Dim islurBy As Double
        Dim sX As Double
        Dim sY As Double
        
        islurAddX = 0.6245 * gTextSize
        islurAddY = 0.4333 * gTextSize
        
        islurBy = 0.45 * gTextSize
        
        '看是否弧線的距離太近
        
        sLenght = Abs(starX - mt_slur_right.Grip.gptMid.x) - islurAddX
        If sLenght < 1 Then
            sLenght = 1
        End If
            sX = mt_slur_right.Grip.gptMid.x - islurAddX - sLenght
            sY = mt_slur_right.Grip.gptMidUp.y + islurBy + islurAddY
        
            points(0) = sX
            points(1) = sY
            points(2) = points(0) + sLenght
            points(3) = points(1)
            points(4) = points(2) + islurAddX
            points(5) = points(3) - islurAddY
        ' Create a lightweight Polyline object in model space
        Set plineSlur = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
        
        
        ' Change the bulge of the third segment
        plineSlur.SetBulge 1, (-0.5858 / 1.4142)
        plineSlur.setWidth 0, 0.15, 0.15
        plineSlur.setWidth 1, 0.15, 0
                           
        plineSlur.Layer = "fige"
        
        color.ColorIndex = 3
        plineSlur.TrueColor = color
        plineSlur.Update
                        
        Set drawSlurStarTo = plineSlur
End Function

Public Function drawBlockItem(blockName As String, x As Double, y As Double, Optional size As Double = 1, _
    Optional dx As Double = 0, Optional dy As Double = 0, _
    Optional scX As Double = 1, Optional scY As Double = 1, Optional scZ As Double = 1)
    '繪 圖塊
    Dim textObj As AcadText
    Dim blockRefObj  As AcadBlockReference
    Dim textString As String
    Dim aligPoint(0 To 2) As Double
    
    aligPoint(0) = x + (size * dx)
    aligPoint(1) = y + (size * dy)
    aligPoint(2) = 0
    
    If checkBlockName(blockName) Then
        Set drawBlockItem = ThisDrawing.ModelSpace.InsertBlock(aligPoint, blockName, size * scX, size * scY, size * scZ, 0)
        drawBlockItem.Layer = MAINLAYER
    Else
        If blockName = "" Then blockName = " "
        Set drawBlockItem = ThisDrawing.ModelSpace.AddText(blockName, aligPoint, size)
    End If
End Function
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
    Dim i As Integer
    Dim TheObj As AcadEntity
    Dim towPos As Double
    Dim cc As String
    Dim mX As Double
    Dim mY As Double
    Dim mItem As MusicFing
    Dim ItemCount As Integer
    Dim currentY As Double
    Dim x As Double, y As Double, Z As Double
    Dim lastOctave As Integer
    Dim scY As Double
    
    x = Grippt.gptMid.x
    y = Grippt.gptMid.y
    Z = Grippt.gptMid.Z
    currentY = 0
    For ItemCount = 0 To UBound(MFS)
        mItem = MFS(ItemCount)
        If mItem.sNote <> "" Then
        If UBound(MFS) > 1 Then
         cc = cc
         End If
        
        '設訂疊音的 y 軸比例
        If UBound(MFS) = 3 Then
            scY = 0.7
        ElseIf UBound(MFS) >= 4 Then
            scY = 0.6
        Else
            scY = 1
        End If
        
            mX = 0: mY = 0
                If mItem.iScale <= 0 Then
                    If ItemCount >= 1 Then
                        lastOctave = lastOctave + Abs(mItem.iScale)
                    ElseIf mItem.iTempo >= 1 Then    '下點音附 要移下去一些
                        '有線時 要下去的距離
                        mX = 0
                        mY = -(amt.DROP_ADD_LINE_INTERVAL * mItem.iTempo)
                    End If
                    mY = mY - 0.2
                Else
                    mY = mY + 1 + 0.2
                End If
                
                If ItemCount > 0 Then
                    If lastOctave > 0 Then
                        y = y + (lastOctave * 0.3) * gTextSize - (lastOctave * 0.3) * gTextSize * (1 - scY)
                    End If
                    y = y + 1.1 * gTextSize - gTextSize * (1 - scY)
                End If
            If mItem.sScale <> "" Then
                Call drawBlockItem(mItem.sScale, x, y, gTextSize, mX, mY)
            End If
            
            '計算多音符的 y 軸移動
            If ItemCount = 0 Then lastOctave = 0
        
        
            '繪主音
            'Y = Y + (gTextSize * itemCount) + totalOctave
            mX = 0: mY = 0
'            If (ItemCount = 0) Then
'                '只取得合音的第一個
'                Set TheObj = drawBlockItem(mItem.sNote, x, y, gTextSize, mX, mY, 1, 1, 1)
'            Else
                 Set TheObj = drawBlockItem(mItem.sNote, x, y, gTextSize, mX, mY, 1, scY, 1)
'            End If
            
            '繪負點音符
            If (MData.dots > 0) Then
                Call drawBlockItem("M-NOTE.", x + (amt.wNote * gTextSize), y, gTextSize, mX, mY)
            End If
                        
            
            
            If mItem.iScale > 0 Then   '低八度的點
                lastOctave = mItem.iScale
            Else
                lastOctave = 0
            End If
            
            
            '升降記號
            Call drawBlockItem(mItem.sTONE, x, y, gTextSize, 0, 0)
        
            
            If frmEDIT.chkOption1.value = True Then
                '古箏指法 1
                mX = mItem.mfPosition.x
                mY = mItem.mfPosition.y
                Call drawBlockItem(mItem.sFinge, x, y, gTextSize, mX, mY)
                
                '古箏指法 2
                towPos = 0
                If mItem.sTowFinge <> "" Then
                    
                    mX = mItem.mfTowPosition.x
                    mY = mItem.mfTowPosition.y
                    
                    '減 拍線
                    towPos = (amt.DROP_ADD_LINE_INTERVAL * mItem.iTempo)
                    If mItem.iScale < 0 Then
                        '減 下點音附
                        towPos = towPos + (amt.DROP_INTERVAL * (-mItem.iScale))
                    End If
                    If towPos > 0.3 Then
                        mY = mY - towPos
                    Else
                        mY = mY - 0.3
                    End If
                    
                    Call drawBlockItem(mItem.sTowFinge, x, y, gTextSize, mX, mY)
                    
                End If
                
                
            '*****畫特疏指法*******************************************
                If mItem.sTONE = "*" Then
                    InsertMusicStar insertionPoint, gTextSize
                End If
    
                
            ElseIf frmEDIT.chkOption2.value = True Then
                '二胡指法用
                
                
            End If
            
            '寫入屬性
            '本段代碼從選中的圖塊中獲取屬性值，並對其修改
    
            
            BlockReferenceSetMusicData TheObj, MData.notes(ItemCount)
            Set InsterEnt = TheObj
            
                
            
        End If
    Next
End Function

Private Sub BlockReferenceSetMusicData(TheObj, note1 As MusicNoteItem)
    Dim blkRefObj As AcadBlockReference
    Dim attVars As Variant
    Dim i As Integer
    If TheObj.objectName = "AcDbBlockReference" Then
        '如果選擇的是塊引用，將其賦給塊引用對象
        Set blkRefObj = TheObj
        '判斷該塊引用是否含有屬性值
        If blkRefObj.HasAttributes Then
            'MsgBox "你選擇的圖塊沒有塊屬性，程序將退出！"
            '如果不含由屬性值退出
    
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
                        attVars(i).textString = IIf(IsEmpty(note1.mtone), "", note1.mtone)
                    Case "IFINGE"
                        attVars(i).textString = IIf(IsEmpty(note1.mfingering), "", note1.mfingering)
                    Case "ISCALE"
                        attVars(i).textString = IIf(IsEmpty(note1.mscale), "", note1.mscale)
                    Case "INOTE"
                        attVars(i).textString = IIf(IsEmpty(note1.mnote), "", note1.mnote)
                    Case "ITEMPO"
                        '取得關鍵字 拍子  -=368acefz amt.iTempo
                        attVars(i).textString = IIf(IsEmpty(note1.mtempo), "", note1.mtempo)
                    Case "ITOWFINGE"
                        attVars(i).textString = IIf(IsEmpty(note1.mtow_fingering), "", note1.mtow_fingering)
                    Case "ISLUR"
                        attVars(i).textString = IIf(IsEmpty(MData.slurStart), "", MData.slurStart)
                    Case "IRE1"
                    Case Else
                End Select
            Next
            
        End If
    End If
End Sub

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




Public Function addMusicJoin(ByVal the_Ids As iArray)

'取得物件 MusicText
'用x 軸來鏈結
On Error Resume Next

    Dim ent As Object
    
    '排序陣列
    Dim TheArray As New iArray
    Dim TheArrayEnt As New iArray
    'Dim coI As Integer
    Dim i As Long
    TheArray.Clear
    TheArrayEnt.Clear
    ' Add the dimstyle to the dimstyles collection

  ' 找出圖面中的每一個選集
'    For i = 0 To the_Ids.Count - 1
'        If the_Ids(i).objectName = "AcDbBlockReference" Then
'            coI = coI + 1
'        End If
'    Next

    

    
    'ReDim TheArray(coI - 1)
    'ReDim TheArrayEnt(coI - 1)
    
    
    'Dim MTobj As NnmText
    Dim pt As Variant
    
    For i = 0 To the_Ids.Count - 1
        Set ent = the_Ids(i)
        If (Mid(ent.Name, 1, 6) = "M-NOTE") Then
            pt = ent.insertionPoint()
            TheArray.Push pt(0)
            TheArrayEnt.Push ent
                       
            
            Dim minExt As Variant
            Dim maxExt As Variant
    
            ent.GetBoundingBox minExt, maxExt
            'Set MTobj = ent
            'minExt = MTobj.GripLeft()
            
        End If
    Next
    
    
    '只有一個，不用排序了
    If the_Ids.Count > 1 Then
        SelectionSortMaxTow TheArray.ToList, TheArrayEnt.ToList

    End If
    '排序開始
    
    
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
    For i = 0 To TheArray.Count - 1
        Set entArr = TheArrayEnt(i)
        cIT = GetAttribTextString(entArr, "ITEMPO")
        
        tempo_num(i) = Tempo_c2line(cIT)
        insertionPoint = TheArrayEnt(i).insertionPoint
        Set ptTheArr(i) = New point
        ptTheArr(i).c insertionPoint(0), insertionPoint(1), insertionPoint(2)
    Next
    
    

    Dim x As Integer
'    for (int x =1  x<=the_lineOne  x++)
    For x = 1 To 4
        For i = 0 To TheArray.Count - 1
            If tempo_num(i) >= x Then
                If stPt.x = 0 Then
                    Set stPt = ptTheArr(i).clone '取得第一個位置
                Else
                    Set endPt = ptTheArr(i).clone '取得第 後面的位置
                End If
                If i = TheArray.Count - 1 Then '如是最後一個就畫出來
                    If endPt.x = 0 Then
                        Set endPt = stPt.clone
                    End If
                    line_link x - 1, stPt, endPt
                    stPt.c 0, 0, 0
                    endPt.c 0, 0, 0
                End If
            Else
                If stPt.x <> 0 Then
                    If endPt.x = 0 Then
                        Set endPt = stPt.clone
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
'''畫拍子的線
''' num_col   要畫幾個線
''' l_pt, r_pt  左右鏈結的中點

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
    points(1) = lPt.y - (amt.bkInster_Line_H * gTextSize) - (num_col * amt.LINE_JUMP * gTextSize)
    points(2) = lPt.Z
    
    
    points(3) = rPt.x + amt.A_TEXT_WIDTH / 2 * gTextSize
    points(4) = points(1)
    points(5) = rPt.Z
        
    ' Create a lightweight Polyline object in model space
    Set plineObj = ThisDrawing.ModelSpace.AddPolyline(points)
    plineObj.setWidth 0, gTextSize * amt.LINE_THICKNESS, gTextSize * amt.LINE_THICKNESS
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
    lPt.y = midDownPt(1) + (0.15 * size * ssize)
    lPt.Z = midDownPt(2)
    
    rPt.x = lPt.x
    rPt.y = lPt.y
    rPt.Z = lPt.Z

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
    
'    MusicData tmp_md = this->m_MusicData
'
'    string cst
'    cst += "MT::linkGrapTempoMaray"
'    cst += tmp_md.strNnmTextData
'
'    LOG1W(cst.c_str(),(long )__LINE__)
'
'
'    AcGePoint3d stPt, endPt
'    stPt = AcGePoint3d(0,0,0)
'    endPt = AcGePoint3d(0,0,0)
'
'    MusicText *tmp_objMTXT
'
'    deque<structMT> sDeque
'    deque<structMT>::iterator  tmp_iterSDeque
'
'    ''mapData::iterator pos
'
'    structMT tmp_sMT
'    ''Acad::ErrorStatus es
'
'    tmp_sMT.pt = tmp_md.point_
'    tmp_sMT.m_tempo = tmp_md.tempo
'
'    sDeque.push_back(tmp_sMT)
'
'
'    if (tmp_md.LlinkobjId != NULL)
'    {
'        return
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
'                tmp_md = tmp_objMTXT->m_MusicData
'
'                tmp_sMT.pt = tmp_md.point_
'                tmp_sMT.m_tempo = tmp_md.tempo
'
'                sDeque.push_back(tmp_sMT)
'                tmp_objMTXT->close()
'            }
'            Else
'            {
'
'                cst = ""
'                cst += "未取得 id- pu438"
'                LOG1W(cst.c_str(),(long )__LINE__)
'
'                acutPrintf( "\n 未取得 id- pu438" )
'                tmp_md.RlinkobjId = NULL
'
'            }
'        }
'    }
'
'    LOG1W(cst.c_str(),(long )__LINE__)
'
'    for (int x =1  x<=the_lineOne  x++)
'    {
'        for( tmp_iterSDeque = sDeque.begin()  tmp_iterSDeque != sDeque.end()  tmp_iterSDeque++)
'        {
'            if(((*tmp_iterSDeque).m_tempo) >= x)
'            {
'                if(stPt.x == 0)
'                {
'                    stPt = (*tmp_iterSDeque).pt     ''取得第一個位置
'                }
'                Else
'                {
'                    endPt = (*tmp_iterSDeque).pt    ''取得第後面的位置
'                }
'                if(tmp_iterSDeque == (sDeque.end()-1))''如是最後一個就畫出來
'                {
'                    if (endPt.x == 0)
'                    {
'                        endPt = stPt
'                    }
'                    m_pGObject->line_link(pW, x-1, stPt, endPt)
'                    stPt = AcGePoint3d(0,0,0)
'                    endPt = AcGePoint3d(0,0,0)
'                }
'
'            }
'            Else
'            {
'                if(stPt.x != 0)
'                {
'                    if (endPt.x == 0)
'                    {
'                        endPt = stPt
'                    }
'                    m_pGObject->line_link(pW, x-1, stPt, endPt)
'                    stPt = AcGePoint3d(0,0,0)
'                    endPt = AcGePoint3d(0,0,0)
'                }
'            }
'        }
'    }
'
'
'}

End Sub
Public Function insterPositionBox(insterPt As point, G As Glode, Optional layerName As String = "Defpoints")
'插入定位框
    Dim plineObj As AcadPolyline
    Dim Pnt As New PointList
    Dim ip As New point
    ip = insterPt
    Pnt.Push ip.c(insterPt.x, insterPt.y - 200, 0).clone
    Pnt.Push ip.c(insterPt.x, insterPt.y + G.FONTSIZE * 9, 0).clone
    Pnt.Push ip.c(insterPt.x + G.pagewidth, insterPt.y + G.FONTSIZE * 9, 0).clone
    Pnt.Push ip.c(insterPt.x + G.pagewidth, insterPt.y - 200, 0).clone
    
'    Call Pnt.Push(.Add(insterPt.X, insterPt.Y - 200, 0)
'    Call Pnt.Add(insterPt.X, insterPt.Y + G.FontSize * 9, 0)
'    Call Pnt.Add(insterPt.X + G.PageWidth, insterPt.Y + G.FontSize * 9, 0)
'    Call Pnt.Add(insterPt.X + G.PageWidth, insterPt.Y - 200, 0)

    Set plineObj = ThisDrawing.ModelSpace.AddPolyline(Pnt.ToXYZList)
    plineObj.Layer = "Defpoints"
    Set insterPositionBox = plineObj
End Function

   
   
   

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
    kTowFing(Asc("b")) = "MT" & Format(Asc("b"), "0000") 'ASC_t 116 ChrW(&H670B)  ' b 按音
    
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
    posTowFing(Asc("b")).c -0.363, -0.7, 0 ' w
    
      'xnote.0': { d: [['M', 7.16, 0], ['C', 8.98, 0, 10.39, 0.91, 11.41, 2.74], ['C', 12, 3.9, 12.3, 5.37, 12.3, 7.15], ['L', 12.3, 9.42], ['C', 12.3, 10.72, 11.93, 12.13, 11.18, 13.64], ['C', 10.47, 15.19, 9.15, 15.98, 7.21, 16], ['C', 5.61, 15.99, 4.5, 15.61, 3.88, 14.87], ['C', 3.26, 14.13, 2.84, 13.42, 2.62, 12.75], ['C', 2.4, 12.07, 2.24, 11.33, 2.14, 10.51], ['C', 2.05, 9.69, 2, 8.61, 2, 7.26], ['C', 2.05, 5.31, 2.34, 3.85, 2.87, 2.88], ['C', 3.7, 0.99, 5.13, 0.03, 7.16, 0], ['Z']], w: 14, h: 16 },
    'xnote.1': { d: [['M', 6.46, 0], ['L', 8.66, 0], ['L', 8.84, 0.19], ['L', 8.84, 15.8], ['L', 8.66, 16], ['L', 5.37, 16], ['L', 5.17, 15.8], ['L', 5.17, 3.82], ['L', 3.22, 3.82], ['L', 3, 3.65], ['L', 3, 2], ['L', 3.03, 1.97], ['L', 3.25, 1.78], ['C', 4.17, 1.81, 4.85, 1.73, 5.29, 1.56], ['C', 5.72, 1.38, 6.04, 0.91, 6.24, 0.16], ['L', 6.46, 0], ['Z']], w: 14, h: 16 },
  
    
    
    
End Sub
