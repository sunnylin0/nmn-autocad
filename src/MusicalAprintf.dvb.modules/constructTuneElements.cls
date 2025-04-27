VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "constructTuneElements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public line As Integer
Public startX As Double
Public w As Double
Public height As Double
Public gTextSize As getTextSize
Public groupVoices123 As New VoiceElementList    '這是每行的 voices
Public staffInfos As vStaffInfoList
Public brace As BraceElemList
Public bracket As BraceElemList
Private sa_(100) As size
Private saIs As New Dictionary
Private igroup As Integer
Private ich As Integer
Private mBuf As DataBuffer
Private barConfig() As aBarConfig
    'ReDim barConfig(200)



Function translate2Staffs(mBuf As DataBuffer) As TuneLineList
 
    Dim abcLines As TuneLineList
    Dim rr As runTime
    Set rr = rTime
    
    rr ("tidyDataToRow")
    Set groupVoices123 = tidyDataToRow(mBuf)
    rr ("tidyDataToRow")
    Dim i As Integer

    rr ("create2TuneLines")
    Set abcLines = create2TuneLines(groupVoices123)
    rr ("create2TuneLines")
    'layoutxy abcLines, 30
    
    
    Set translate2Staffs = abcLines
End Function
'' call when spacingduration has been updated
Sub updateNextX(x As Double, spacing As Double, voice As VoiceElement)
    voice.nextX = x + (spacing * Math.Sqrt(voice.spacingduration / 1000 * 8))
End Sub
Function layoutxy(abcLines As TuneLineList, spaceing As Double)
    Dim aLine As TuneLine
    Dim vo As VoiceElement
    Dim mus As MusicItem
    Dim i As Integer
    Dim j As Integer
    Dim num As Integer
    Dim currSpace As Double
    Dim spaceUnit As Double
    Dim parentWidth As Double
    Dim calW As Double
    For i = 0 To abcLines.Count - 1
        Set aLine = abcLines(i)
        For j = 0 To aLine.staffGroup.voices.Count - 1
            Set vo = aLine.staffGroup.voices(j)
            currSpace = spaceing
            For num = 0 To 7
                spaceUnit = layoutOneItem(vo, currSpace)
                '計算寬度
                parentWidth = 210 - G.RightSpace + G.LeftSpace
                
                If (vo.w >= parentWidth) Then
                    calW = vo.w - parentWidth
                    currSpace = calW / spaceUnit
                Else
                    Exit For
                End If
            Next
            
           
        Next
        
    Next

End Function
Function layoutOneItem(vele As VoiceElement, spacing As Double) As Double
    Dim k As Integer
    Dim mus As MusicItem
    Dim spaceCurr As Double
    Dim spaceUnit As Double
    spaceUnit = 0
    vele.nextX = 0
    For k = 0 To vele.children.Count - 1
        Set mus = vele.children(k)
                    
        If (mus Is Nothing) Then
            Exit For
        End If
        
        Select Case mus.typs
           Case Cg.bar:
                mus.x = vele.nextX
                mus.w = amt.wBar * G.fontName
           Case Cg.meter:
                mus.x = vele.nextX
                mus.w = amt.wMete * G.fontName
           Case Cg.Rest, Cg.note:
                mus.x = vele.nextX
                mus.w = amt.wNote * G.fontName
                calExtraw mus   '計算前綴元素
           Case Else
                mus.x = vele.nextX
                mus.w = amt.wOther * G.fontName
        End Select
        vele.w = vele.nextX
        If (mus.duration > 0) Then
            spaceCurr = (spacing * Math.Sqrt(mus.duration / 1000 * 8))
        Else
            spaceCurr = (spacing * Math.Sqrt(1 / 1000 * 8))
        End If
            vele.nextX = mus.x + mus.w + spaceCurr
            vele.minX = mus.x + mus.w + 0.1
            vele.nextX = Math.max(vele.nextX, vele.minX)
            spaceUnit = spaceUnit + spaceCurr
        
    Next
    layoutOneItem = spaceUnit

End Function
Function create2TuneLines(groupVoices As VoiceElementList) As TuneLineList
    On Error Resume Next
    Dim abcLines As TuneLineList
    Dim abcStaffs As StaffList
    Dim tmpTL As TuneLine
    Dim vo As VoiceElement
    Dim sff As Staff
    Dim i As Integer
    Dim N As Integer
    Dim posLine As Integer
    Dim currVoiceNumber As Integer
       
    Set abcLines = New TuneLineList
    
    ''寫入 abcLines.staffs
    For i = 0 To groupVoices.Count - 1
        Set vo = groupVoices(i)
        
        If (vo.voicenumber > currVoiceNumber) Then
            posLine = 0
            currVoiceNumber = vo.voicenumber
        End If
        If (abcLines(posLine) Is Nothing) Then
             abcLines.Push New TuneLine
        End If
        If (abcLines(posLine).Staffs Is Nothing) Then
            Set abcLines(posLine).Staffs = New StaffList
        End If
        
        ''寫入 abcLines.staffs
        Set sff = New Staff
        Set sff.voices = vo.children
        abcLines(posLine).Staffs.Push sff

        
        ''寫入 abcLines.StaffGroup
        If (abcLines(posLine).staffGroup Is Nothing) Then
            Set abcLines(posLine).staffGroup = New StaffGroupElement
            Set abcLines(posLine).staffGroup.voices = New VoiceElementList
        End If
        abcLines(posLine).staffGroup.voices.Push vo
        posLine = posLine + 1
        

    Next
    
    posLine = 0

    
    
    Set create2TuneLines = abcLines
End Function
Function tidyDataToRow(mBuf As DataBuffer) As VoiceElementList
'資料整理成小節數 一行一行的array


    Dim tmp_track As Integer
    Dim tmp_track_item As Long

    Dim s1 As MusicItem
    Dim s2 As MusicItem
    
    
     Call mBuf.init
     
    Dim i As Long
    Dim voiceList As New VoiceElementList
    Dim vo As VoiceElement
    Dim trackIndex As Integer
    Dim barId As Integer
    Dim barEle As MusicItem
    Dim isMeasureOne As Boolean  '是否小節第一個
    ReDim barConfig(2000)

        For tmp_track = 0 To mBuf.GetTrackSize() - 1

            Call mBuf.init
            
            G.durationIndex = 0
            G.currLine = 0
            G.currMeasure = 0
            
            barId = 0
            Set vo = New VoiceElement
            Call vo.init(tmp_track, mBuf.GetTrackSize() - 1)
            isMeasureOne = True
                        

            
            For tmp_track_item = 0 To mBuf.GetTrackBufferSize(tmp_track) - 1
                
                Set s1 = mBuf.GetData(tmp_track, tmp_track_item)
                If (s1 Is Nothing) Then
                    Exit For
                End If
                

                If (G.durationIndex >= (Cg.BLEN / G.mete2) * G.mete) Then
                    '是否以到小節長度
                
                    G.currMeasure = G.currMeasure + 1
                    G.durationIndex = 0
                        
                    If s1.typs <> Cg.bar Then vo.insterBar barId
                    
                    barId = barId + 1
                    If barConfig(barId) Is Nothing Then
                        Set barConfig(barId) = New aBarConfig
                        barConfig(barId).barId = barId
                    ElseIf barConfig(barId).barLineQuantity > 1 Then
                        G.barsperstaff = barConfig(barId).barLineQuantity
                    End If


                    
                    If G.currMeasure >= G.barsperstaff Then  '看是否以到一行小節個數了
                        isMeasureOne = True
                        G.currMeasure = 0
                        G.currLine = G.currLine + 1
                        
                        voiceList.Push vo
                        
                        Set vo = New VoiceElement
                        Call vo.init(tmp_track, mBuf.GetTrackSize() - 1)
                    End If
                End If

                If isMeasureOne = True Then
                    '這是看第一小節 ，是否要加入小節元素
                    If (G.durationIndex = 0 And (s1.typs = Cg.note Or s1.typs = Cg.meter)) Then
                        Set s2 = vo.insterBar(barId)
                        s2.remark = "t" & tmp_track & "_i" & tmp_track_item
                        isMeasureOne = False
                    End If
                End If
                
                If (isMeasureOne = False And tmp_track > 0 And G.durationIndex = 0) Then
                    If Not barConfig(barId) Is Nothing Then
                        '取得第一軌的 第 barId 小節設定
                        If barConfig(barId).mete > 0 Then
                            Set s2 = vo.insterMete(barConfig(barId).mete, barConfig(barId).mete2)
                            s2.remark = "t" & tmp_track & "_i" & tmp_track_item
                            G.mete = barConfig(barId).mete
                            G.mete2 = barConfig(barId).mete2
                            G.durationIndex = G.durationIndex + 0.000005
                        End If
                        If barConfig(barId).barLineQuantity > 0 Then
                            G.barsperstaff = barConfig(barId).barLineQuantity
                        End If
                    End If
                End If

                Select Case s1.typs
                    Case Cg.Config:
                        If (s1.barsperstaff >= 1 And tmp_track = 0) Then
                            G.barsperstaff = s1.barsperstaff
                            If barConfig(barId) Is Nothing Then
                                Set barConfig(barId) = New aBarConfig
                            End If
                            barConfig(barId).barId = barId
                            barConfig(barId).barLineQuantity = s1.barsperstaff
                        End If
                        If (s1.setbarstaffid > 1) Then
                            If barConfig(s1.setbarstaffid) Is Nothing Then
                                Set barConfig(s1.setbarstaffid) = New aBarConfig
                            End If
                            barConfig(s1.setbarstaffid).barId = s1.setbarstaffid
                            barConfig(s1.setbarstaffid).barLineQuantity = s1.setbarstaff
                        End If
                        
                        If (s1.barlinebreak > 1) Then
                            If barConfig(setbarstaffid) Is Nothing Then
                                Set barConfig(setbarstaffid) = New aBarConfig
                            End If
                            barConfig(setbarstaffid).barId = setbarstaffid
                            barConfig(setbarstaffid).barLineQuantity = s1.setbarstaff
                        End If
                    Case Cg.bar:
                        s1.w = amt.wBar * G.fontsize
                        s1.barNumber = barId
                        s1.remark = "t" & tmp_track & "_i" & tmp_track_item
                        vo.children.Push s1
                        isMeasureOne = False
                        
                    Case Cg.meter:
                        
                        
                        If (tmp_track = 0) Then
                            G.mete = s1.mete
                            G.mete2 = s1.mete2
                            Set s2 = vo.insterMete(s1.mete, s1.mete2)
                            s2.remark = "t" & tmp_track & "_i" & tmp_track_item
                            Set barConfig(barId) = New aBarConfig
                            barConfig(barId).barId = barId
                            barConfig(barId).mete = s1.mete
                            barConfig(barId).mete2 = s1.mete2
                        End If
                        
                   Case Cg.Rest, Cg.note:
                        s1.w = amt.wNote * G.fontsize
                        If s1.dots > 0 Then
                            s1.w = 2 * amt.wNote * G.fontsize
                        End If
                        calExtraw s1 '計算前綴元素
                        s1.remark = "t" & tmp_track & "_i" & tmp_track_item
                        vo.children.Push s1
                        G.durationIndex = G.durationIndex + s1.duration
          
                   Case Else
                End Select
                

                
            Next
            If vo.children.Count >= 1 Then '看有無資料
                G.currMeasure = 0
                G.currLine = G.currLine + 1
                
                If (vo.children.Last.typs <> bar) Then
                    '看最後一個是否是 bar
                    '如不是增加
                    vo.insterBar barId
                End If
                voiceList.Push vo
                
                Set vo = New VoiceElement
                Call vo.init(tmp_track, mBuf.GetTrackSize() - 1)
            End If
            
        Next
    Set tidyDataToRow = voiceList
End Function
'Public Sub read2BarGold(measure As Integer, voice As VoiceElement)
'    '取得第 0 軌的 第 measure 小節設定
'    Dim staffn As Integer
'    Dim m1 As Integer
'    Dim m2 As Integer
'    staffn = barConfig(measure).barLineQuantity
'    m1 = barConfig(measure).mete
'    m2 = barConfig(measure).mete2
'    G.barsperstaff = IIf(staffn > 0, staffn, G.barsperstaff)
'    If (m1 >= 1 And m2 >= 1) Then
'        G.mete = m1
'        G.mete2 = m2
'
'        '還要插入一個 mete
'        voice.insterMete m1, m2
'    End If
'
'End Sub
'
'
