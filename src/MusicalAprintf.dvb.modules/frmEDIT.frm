VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEDIT 
   Caption         =   "�϶��� EDIT"
   ClientHeight    =   5844
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9660
   OleObjectBlob   =   "frmEDIT.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEDIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit
'2024.11.28  isVirtualChar As Boolean '�]�w�Ťߦr��  v2.13
'2024.11.28  �[�J I:setbar 5/3 �b�x�Y�]�w "�ĴX�p�`�}�l/�C��X�p�`"
'2024.11.27  v2.12 tuplet "{" "}" �[�J 3 5 6 7 �s�����\��
'2024.08.27  �ק� �|�����������I��m
'2024.08.23  "v2.1" �[�J �e�󤸯� '>'  ��󤸯� '<'
'2024.03.30  "v2.0" �n�[�J �縹 M:3/8 M:4/4 �A�j�ﭵ�űƧǤ覡
'2024.03.29  �ק� 4/4 3/8 ��ƹ����D
'2013.11.21  �ק� DataBuffer ������
'            �h�[ iAdd �X����
'2013.03.17  V3 ���n�ק� �G�J�������A�]�{�����e�O�Υj�媺���k�ϡA�{�b�令�G�J�����k��

Const version  As String = "v2.13" '�n�鸹�X
Const c1 As Integer = 60   'C��1����W��
'Const FOURPAINUM   As Integer = 64 '1/4���ŭp��
'Const MIDICLOCK As Integer = 24   '�C1/64���Ū�MIDICLOCK��
'Const TEMPO_DEF As Integer = 90   '�w�]�C����90��
Const PARTITION_DEF As Integer = 384   '�w�]�C����ά�384�p�ɳ椸
Const VOLUME_DEF As Integer = 64
Const MAINLAYER As String = "MAIN"    '�D�n���ϼh
Dim partLayer As String     '�O���{�b�ĴX��



'�G�J���k ���A
Private Type ErhuFing
    fing1 As Integer    '�� �� �� �� "b�Ų�"
    fing2 As String  '�� �� �� �� "b�Ų�"
    Push As String      '�ԡ� ��V
    InOut As String     '��   �~
End Type

Dim m_buf As New DataBuffer
Dim TuneLines As New TuneLineList
Private constTE As New constructTuneElements

'1 �bvb�u�{���ޥ�autocad����
'2 �w��autocad���H
Private acadapp As AcadApplication
Private acadDoc As AcadDocument
'3 �����{�ׄ�autocad����ۡA�H�U�O�ڇ���
'--------------------------------------------------------------
'����Cad
'-------------------------------------------------------------
Private Function AcadConnect() As Boolean
Dim flag As Boolean
On Error Resume Next
    Set acadapp = GetObject(, "AutoCAD.Application")
    flag = True
    If err Then
       err.Clear
       Set acadapp = CreateObject("AutoCAD.Application")
       flag = True
       If err Then
          flag = False
          MsgBox "�����b��AutoCAD,�[��d�O�_�w�E�I", vbOKCancel, "ĵ�i�I"
          Exit Function
       End If
    End If
    AcadConnect = flag
    Set acadDoc = acadapp.ActiveDocument
    'acadDoc.Close False
End Function


Private Sub cbCANCLE_Click()
    Me.Hide
End Sub


Private Sub init()
    G.fontName = Me.cobFontName.text
    G.fontsize = Me.cobFontSize.text
    
    Dim mete_mete As Variant
    mete_mete = Split(Me.cobMete, "/")
    G.mete = mete_mete(0)
    G.mete2 = mete_mete(1)
    
    G.Many = Me.cobMany
    G.bar = Me.cobBar.text
    G.barsperstaff = Me.cobBar.text
    
    G.pagewidth = Me.tbPageWidth '�e��
    G.LeftSpace = Me.tbLeftSpace  '���ť�
    G.RightSpace = Me.tbRightSpace  '�k�ť�
    G.BarToNoteSpace = Me.tbBarToNote    '�p�`�쭵��
    G.TrackToTrack = Me.tbTrackToTrack  '�n�����Z
    G.LineToLine = Me.tbLineToLine      '�C�涡�Z
    G.check1 = True
    G.MIN_X = Me.tbMIN_X            '�L��
    G.Beat_MIN_X = Me.tbBeat_min_x  '��L��
    G.IsBarAlign = Me.cbIsBarAlign
    G.isVirtualChar = Me.cbVirtualChar  '�Ťߦr
End Sub

Private Sub cmOK_Click()

    Call init

    database
    Set m_buf = New DataBuffer '.Clear
    
    Dim rr As runTime
    Set rr = rTime
    rr ("LoadDataToBuf")
    Call m_buf.LoadDataToBuf(Me.TextBox8.text)
    rr ("LoadDataToBuf")
    'MsgBox Me.TextBox8.text
    Me.Hide

    '��z �縹�Τp�`�� �����@��@�檺 TuneLines
    Set TuneLines = constTE.translate2Staffs(m_buf)
    
    'layout �N x ���нվ��䵴���m
    rr ("layout")
    Call layout(TuneLines, 210#, 30)
    rr ("layout")
    
    rr ("drawLayoutStaff")
    Call drawLayoutStaff(TuneLines)
    rr ("drawLayoutStaff")
    'Call draw_many_text1
    'Call setLayoutMusicItem
    
    rr.ToList
End Sub


Private Sub CommandButton2_Click()
    'abc ����

    Call init

    database
    Set m_buf = New DataBuffer '.Clear
    Call m_buf.LoadDataToBuf(Me.TextBox8.text)
    'MsgBox Me.TextBox8.text
    Me.Hide
    
    Set tune = New TuneData
    Set gTuneLine = New TuneLine
    Set gTuneLine.staffGroup = New StaffGroupElement
    Set gTuneLine.Staffs = getToStaffList(m_buf)
    ''GTuneLine..Staffs getInitABCElement
    
    Dim v As VoiceABCList
    Set v = gTuneLine.Staffs(0).voices(0)
    
    ''v.AddArrayAfter , m_Buf.getToVoiceABCs
    Dim eg As New EngraverController
    eg.init Nothing, Nothing
    'getToStaff m_Buf
    
    Set tune.engraver = eg
    If (tune.lines Is Nothing) Then
        Set tune.lines = New TuneLineList
        tune.lines.Push gTuneLine
    Else
        tune.lines.Push gTuneLine
    End If
    tune.version = "1.0.0"
    
    eg.engraveABC tune
    
    Call draw_many_text1

    
End Sub



Private Function getInitABCElement() As Staff
'�]�w�ո� ���� �縹
    Dim staffEle As New Staff
    Dim cl As vClefProperties
    Dim ky As vKeySignature
    Dim mf As vMeter
    
    '�s �ո�
    Set cl = New vClefProperties
    cl.el_typs = "clef"
    cl.typs = "treble"
    cl.verticalPos = 0
    
    
    '�s ����
    Dim vAcc As New vAccidental
    vAcc.acc = "sharp"
    vAcc.note = "f"
    vAcc.verticalPos = 10
    Set ky = New vKeySignature
    Set ky.accidentals = New iArray
    ky.el_typs = "keySignature"
    ky.root = "E"
    ky.accidentals.Push vAcc
    
    
    '�縹
    Dim mFrac As New vMeterFraction
    mFrac.num = 4
    mFrac.den = 4
    Set mf = New vMeter
    Set mf.value = New vMeterFractionList
    mf.el_typs = "timeSignature"
    mf.typs = "specified"
    mf.value.Push mFrac
    
    
    Set staffEle.clef = cl
    Set staffEle.key = ky
    Set staffEle.meter = mf
    
    Set staffEle.voices = New iArray
    Set getInitABCElement = staffEle
End Function
Private Function getToStaffList(buf As DataBuffer) As StaffList
    Dim trkCount As Long
    Dim noteCount As Long
    Dim i As Integer, j As Integer
    Dim staffElem As Staff
    Dim currVoice As VoiceABCList
    
    Set getToStaffList = New StaffList
    trkCount = buf.GetTrackSize
    For i = 0 To trkCount - 1
        Set staffElem = getInitABCElement()
        staffElem.voices.Push buf.getVoiceListData(i)
        getToStaffList.Push staffElem
'        buf.ConverToVoiceABCList i
'        If (staffEle.voices Is Nothing) Then
'            ReDim staffEle.voices(trkCount)
'        End If
'        noteCount = buf.GetTrackBufferSize(i)
'        Set currVoice = staffEle.voices(i)
'        For j = 0 To noteCount
'            buf.ConverToVoiceABCList i
'        Next
    
    Next
    


End Function
Private Sub database()
    ' Create new layer
    Dim i As Integer
    Dim color As AcadAcCmColor
    Dim layerObj As AcadLayer
    Dim sarr() As String
    Dim datalayer As New iArray
    '�]�w --> "�ϼh�W�r �C��"
    datalayer.PushArray Array( _
    "FIGE 6", _
    "TEXT 2", _
    "bar 181", _
    "�˹��Ÿ� 4", _
    "��ø�u 1", _
    "main 7", _
    "SimpErhu�Ÿ� 151")

    
    '�s�عϼh
    For i = 0 To datalayer.Count - 1
        sarr = Split(datalayer(i), " ")
        Set layerObj = ThisDrawing.Layers.Add(sarr(0))
        Set color = AcadApplication.GetInterfaceObject("AutoCAD.AcCmColor." & ACAD_Ver)
        color.ColorIndex = sarr(1)
        layerObj.TrueColor = color
    Next


   
    ' Create a Text style named "TEST" in the current drawing
    Dim textStyle As AcadTextStyle
    Dim dataStyles(20, 3) As String
    dataStyles(0, 1) = "Standard"
    dataStyles(0, 2) = "txt.shx"
    dataStyles(0, 3) = "chineset.shx"
    dataStyles(1, 1) = "EUDC"
    dataStyles(1, 2) = "EUDC.TTE"
    dataStyles(1, 3) = ""
    dataStyles(2, 1) = "����"
    dataStyles(2, 2) = "KAIU.TTF"
    dataStyles(2, 3) = ""
    dataStyles(3, 1) = "MMP2005"
    dataStyles(3, 2) = ""
    dataStyles(3, 3) = ""
    dataStyles(4, 1) = "����_�Ʀr"
    dataStyles(4, 2) = "SimSun.ttc"
    dataStyles(4, 3) = ""
    dataStyles(5, 1) = "����_�ө���"
    dataStyles(5, 2) = "MingLiU.ttc"
    'dataStyles(5, 2) = "PMingLiU.ttf"
    dataStyles(5, 3) = ""
    dataStyles(6, 1) = "�r��"
    dataStyles(6, 2) = "MAESTRO.TTF"
    dataStyles(6, 3) = ""
    dataStyles(7, 1) = "�رd�ʶ�"
    dataStyles(7, 2) = "DFFT_C7.ttc"
    'dataStyles(7, 2) = "DFLiHeiBold.ttf"
    dataStyles(7, 3) = ""
    dataStyles(8, 1) = "��r"
    dataStyles(8, 2) = "KAIU.TTF"
    dataStyles(8, 3) = ""
    dataStyles(9, 1) = "SimpErhu"
    dataStyles(9, 2) = "SimpErhuFont.ttf"
    dataStyles(9, 3) = ""

    'Dim i As Integer
'    For i = 0 To 9
'        Set textStyle = ThisDrawing.TextStyles.add(dataStyles(i, 1))
'        If dataStyles(i, 2) <> "" Then
'
'            Set textStyle = CreateTextStyle(dataStyles(i, 2), dataStyles(i, 1), FontType.BigFont)
'        End If
'        If dataStyles(i, 3) <> "" Then
'            textStyle.BigFontFile = dataStyles(i, 3)
'        End If
'    Next
End Sub
Private Sub inst_G(the_G As Glode, aPt As point)
    '���J�]�w���
    Dim mtxt As AcadMText
    Dim textObj As AcadText
    Dim textString As String
    Dim insertionPt As New point
    Dim height As Double
    
    ' Define the text object
    textString = version & vbCrLf
    textString = textString & "size " & the_G.fontsize & vbCrLf
    
    textString = textString & "���ť� " & the_G.LeftSpace & "mm" & vbCrLf
    textString = textString & "�k�ť� " & the_G.RightSpace & "mm" & vbCrLf
    textString = textString & "�p�`�쭵�� " & the_G.BarToNoteSpace & "mm" & vbCrLf
    textString = textString & "�n��  " & the_G.TrackToTrack & "mm" & vbCrLf
    textString = textString & "�C��  " & the_G.LineToLine & "mm" & vbCrLf
    textString = textString & "�L��  " & the_G.MIN_X & "mm" & vbCrLf
    textString = textString & "��L�� " & the_G.Beat_MIN_X
    
    
    insertionPt.x = aPt.x - 30
    insertionPt.y = aPt.y
    insertionPt.Z = aPt.Z
    height = 3
    
    ' Create the text object in model space
    Set mtxt = ThisDrawing.ModelSpace.AddMText(insertionPt.ToDouble, height, textString)
    mtxt.width = 40
    mtxt.styleName = "Standard"
    
    '���J���
    insertionPt.x = insertionPt.x - 3
    Set mtxt = ThisDrawing.ModelSpace.AddMText(insertionPt.ToDouble, 0.01, Me.TextBox8.text)
    mtxt.height = 0.01
    
End Sub
Private Sub drawLayoutStaff(abcTuneLines As TuneLineList)
    Dim MBG As New MusicBlockGraphics
    
    '��m��r��
    Dim retPt As Variant
    Dim insPt As New point
    
    ' Return a point using a prompt
    retPt = ThisDrawing.Utility.GetPoint(, "\n��ܭn���J���I �GEnter insertion point: ")
    insPt.a retPt
    '���J�]�w��Ƥ�r����
    Call inst_G(G, insPt)
'***********************************************************************************
    '�e�X�w��u-�e��
    Dim plineObj As AcadPolyline
    Set plineObj = MBG.insterPositionBox(insPt, G)

'*********************************************************************************
    '���J��@���D
    Dim objText As AcadText
    Dim titlePT As New point
    Dim ooPt As New point
    
    titlePT.x = insPt.x + (G.pagewidth / 2)
    titlePT.y = insPt.y + G.fontsize * 5.5
    Set objText = ThisDrawing.ModelSpace.AddText(m_buf.getTITLE, titlePT.ToDouble, 6)
    ooPt.a objText.insertionPoint
    objText.Layer = "TEXT"
    objText.Alignment = acAlignmentCenter
    objText.styleName = "��r"
    ooPt.x = ooPt.x + objText.insertionPoint(0)
    ooPt.y = ooPt.y + objText.insertionPoint(1)
    
    objText.Move objText.insertionPoint, ooPt.ToDouble

'*********************************************************************************
'�إߥD�n ����
    Dim layerObj As AcadLayer
    Set layerObj = ThisDrawing.Layers.Add(MAINLAYER)


    Dim tmp_joinApp As New iArray
    Dim tmp_joinIds As New iArray
    'double lastTemp
    Dim tmp_track As Integer
    Dim tmp_track_item As Long
    Dim A_TEMPO_add As Integer

    Dim tmp_delaytime As Double  '�p���l �O�n�p�W�@�Ӧr������
    
    Dim tmp_xy As point
    
    Dim BNewObj As AcadBlockReference
    Dim ptlist As New PointList
    Dim tmp_pLWPoly As AcadPolyline
    Dim tmp_name As String
    
    'Dim cst As String
    'Dim cst_no_fing As String
    Dim ptGripMid As Variant
    Dim s1 As MusicItem
    
    tmp_delaytime = 0
    
    tmp_name = G.fontName
    Dim tmp_erhu_fing As ErhuFing
    Dim midDownPt(2) As Double
    Dim mt_slur_left As MusicBlockGraphics
    Dim mt_slur_right As MusicBlockGraphics
    Dim plineSlur As AcadLWPolyline
    Dim mt_tuplet_left As MusicBlockGraphics
    Dim mt_tuplet_right As MusicBlockGraphics
    Dim plineTuplet As AcadLWPolyline
    
    Dim barConfig(1000) As aBarConfig
    Dim barId As Integer
    Dim currStaffGroup As StaffGroupElement
    Dim vo As VoiceElement
    Dim trackY As Double
    Dim iLineToLine As Integer
    Dim iTrackToTrack As Integer
    Dim currX As Double
    Dim currY As Double
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim startPt As New point
    Dim endPt As New point
    'Dim sumLine As Integer  '�p��ĴX��
    'Dim sumMeasure As Integer  '�p��C�檺�ĴX��
    
''******ø�s�U��u �ܼ�
    Dim durationIndex   As Double
    Dim nlast As Integer
    Dim s2 As MusicItem
    Dim Temo1 As Double
    Dim Temo2 As Double
    Dim beamMuseList As New MusicItemList
''***********************************
    trackY = 0
    iLineToLine = -1
    iTrackToTrack = 0
    MBG.setVirtualChar True
    For i = 0 To abcTuneLines.Count - 1
        Set currStaffGroup = abcTuneLines(i).staffGroup
        
        
        For j = 0 To abcTuneLines(i).staffGroup.voices.Count - 1
'            If j = 0 Or j = 2 Then
'                G.FONTSIZE = 3.6
'            Else
'                G.FONTSIZE = 3.4
'            End If
            partLayer = j
            
            
            Set vo = abcTuneLines(i).staffGroup.voices(j)
            startPt.c 0, 0
            endPt.c 0, 0
            If j = 0 Then
                iLineToLine = iLineToLine + 1
            Else
                iTrackToTrack = iTrackToTrack + 1
            End If
            'currY = (G.TrackToTrack * j) + (G.LineToLine * i) + (G.TrackToTrack * (abcTuneLines(i).StaffGroup.voices.Count - 1)) * i
            currY = G.TrackToTrack * iTrackToTrack + G.LineToLine * iLineToLine
            currY = insPt.y - currY
            currX = insPt.x + G.LeftSpace
            
            durationIndex = 0
            Temo1 = 0
            nlast = 0
            For k = 0 To vo.children.Count - 1
                Set s1 = vo.children(k)
                Select Case s1.typs
                   Case Cg.bar:
                        'ø�s�p�`�u  ********************************
                        If j > 0 And k < 3 Then
                            '�ĤG�����Ĥ@�Ӥp�`�u�n�Ԫ�
                            startPt.x = currX + s1.x + s1.w / 2
                            startPt.y = G.TrackToTrack + currY
                            endPt.x = startPt.x
                            endPt.y = currY
                        Else
                            startPt.x = currX + s1.x + s1.w / 2
                            startPt.y = ((amt.LINE_PASE + amt.DROP_UP) * G.fontsize) + currY
                            endPt.x = startPt.x
                            endPt.y = currY
                        
                        End If
                        
                        ptlist.Clear
                        ptlist.Add startPt
                        ptlist.Add endPt
                        
                        s1.oX = startPt.x
                        s1.oY = startPt.y
                        Set tmp_pLWPoly = ThisDrawing.ModelSpace.AddPolyline(ptlist.ToXYZList)
                        tmp_pLWPoly.ConstantWidth = amt.BAR_WITCH * G.fontsize / 4.6
                        tmp_pLWPoly.Layer = "bar"
                        
                        'ø�p�`�� ***********************************
                        If j = 0 And durationIndex = 0 Then
                            Dim ipt(3) As point
                            Dim objT As AcadText
                            Dim txtmidPT As New point
                            Set ipt(0) = New point: Set ipt(1) = New point: Set ipt(2) = New point: Set ipt(3) = New point
                            ipt(0).x = startPt.x
                            ipt(0).y = startPt.y + 0.2 * G.fontsize
                            ipt(1).x = startPt.x: ipt(1).y = startPt.y + (0.2 + amt.barInsterNumberSize * 1.3) * G.fontsize
                            ipt(2).x = ipt(0).x - amt.barInsterNumberSize * (Len(CStr(s1.barNumber + 1)) * 1.1) * G.fontsize
                            ipt(2).y = ipt(1).y
                            ipt(3).x = ipt(2).x
                            ipt(3).y = ipt(0).y
                            ptlist.Clear
                            ptlist.Add ipt(0)
                            ptlist.Add ipt(1)
                            ptlist.Add ipt(2)
                            ptlist.Add ipt(3)
                            ptlist.Add ipt(0)
                            Set tmp_pLWPoly = ThisDrawing.ModelSpace.AddPolyline(ptlist.ToXYZList)
                            tmp_pLWPoly.Layer = "TEXT"
                            
                            txtmidPT.x = ipt(1).x - Abs(ipt(1).x - ipt(2).x) / 2
                            txtmidPT.y = ipt(0).y + Abs(ipt(1).y - ipt(0).y) / 2
                            Set objT = ThisDrawing.ModelSpace.AddText(CStr(s1.barNumber + 1), txtmidPT.ToDouble, amt.barInsterNumberSize * G.fontsize)
                            objT.Layer = "Text"
                            objT.styleName = "txt"
                            ooPt.a objT.insertionPoint
                            objT.Alignment = acAlignmentMiddleCenter
                            
                            ooPt.x = ooPt.x + objT.insertionPoint(0)
                            ooPt.y = ooPt.y + objT.insertionPoint(1)
                            
                            objT.Move objT.insertionPoint, ooPt.ToDouble
                            If s1.barNumber + 1 = 37 Then
                                Debug.Print 37
                            End If
                            
                        End If
                   Case Cg.meter:
                        'ø�s�縹  ********************************
                        Dim metePt1 As New point
                        Dim metePt2 As New point
                        metePt2.x = currX + s1.x
                        metePt2.y = currY + 0.3 * G.fontsize
                        
                        metePt1.x = currX + s1.x
                        metePt1.y = currY + 1.3 * G.fontsize
                        
                        Set objText = ThisDrawing.ModelSpace.AddText(s1.mete2, metePt2.ToDouble, G.fontsize * 0.7)
                        objText.Layer = "�˹��Ÿ�":         objText.styleName = "����_�Ʀr"
                        
                        Set objText = ThisDrawing.ModelSpace.AddText(s1.mete, metePt1.ToDouble, G.fontsize * 0.7)
                        objText.Layer = "�˹��Ÿ�":         objText.styleName = "����_�Ʀr"
                        s1.oX = metePt2.x
                        s1.oY = metePt2.y
                        G.mete = s1.mete
                        G.mete2 = s1.mete2
                   Case Cg.Rest, Cg.note:
                        'ø�s����********************************
                        Dim N As Integer
                        startPt.x = currX + s1.x
                        startPt.y = currY
                        s1.oX = startPt.x
                        s1.oY = startPt.y
                        MBG.setDataText startPt, s1, G.fontsize
                        MBG.nowPartLayer = partLayer
                        Set BNewObj = MBG.InsterEnt '���J���ŤΫ��k
                        
                        ''** ��u�p�� ***********************
                        nlast = s1.nflags
                        Temo1 = Fix(durationIndex / (Cg.BLEN / G.mete2)) ''���o���`�[���e
                        durationIndex = durationIndex + s1.duration
                        
                        If (nlast > 0) Then
                            beamMuseList.Push s1
                             '�o�O�n�s���u���A�H m_Mete2 ���ɭ�
                            Temo2 = durationIndex - (Temo1 * (Cg.BLEN / G.mete2))
                            
                            If Temo2 Mod (Cg.BLEN / G.mete2) = 0 Or Temo2 > (Cg.BLEN / G.mete2) Then
                                MBG.draw_dur beamMuseList
                                beamMuseList.Clear
                            End If
                         ElseIf beamMuseList.Count >= 1 Then
                            MBG.draw_dur beamMuseList
                            beamMuseList.Clear
                        End If
                        
                        '*******ø�s��ƽu'**************************************************************************************'
                        'AMT.iSlur = 7        ' �s���Ŧ�    (3456)
                        If s1.slurStart = True Then
                            Set mt_slur_left = New MusicBlockGraphics
                            Set mt_slur_left = MBG.copy
        
                        ElseIf s1.slurEnd = True Then
                        
                            If mt_slur_left Is Nothing Then
                                Set plineSlur = MBG.drawSlurStarTo(vo.children(0).oX, MBG)
                            Else
                                Set plineSlur = MBG.drawSlur(mt_slur_left, MBG)
                                Set mt_slur_left = Nothing
                            End If
                        End If
                        '*****************************************************************
                               
                                        
                
                        If s1.tupletStart = True Then
                            'Set mt_Tuplet_left = New MusicBlockGraphics
                            Set mt_tuplet_left = MBG.copy
        
                        ElseIf s1.tupletEnd = True Then
        
                            Set plineTuplet = MBG.drawTuplet(mt_tuplet_left, MBG, s1.tupletCount)
        
                        End If
                
                               
                   Case Else
                End Select
                
                
            Next
            
            If Not mt_slur_left Is Nothing Then
                '��ƽu�O�_���e��
                Set plineSlur = MBG.drawSlurToEnd(mt_slur_left, s1.oX + s1.w)
                Set mt_slur_left = Nothing
            End If
        

        Next
    Next
  


End Sub

Private Sub draw_many_text1()
  
  
    Dim MBG As New MusicBlockGraphics
    
    '��m��r��
    Dim insPt As Variant
    Dim ipt As New point
    
    ' Return a point using a prompt
    insPt = ThisDrawing.Utility.GetPoint(, "\n��ܭn���J���I �GEnter insertion point: ")
    '���J�]�w��Ƥ�r����
    Call inst_G(G, insPt)
'***********************************************************************************
    '�e�X�w��u-�e��
    Dim plineObj As AcadPolyline
    ipt = insPt
    Set plineObj = MBG.insterPositionBox(ipt, G)

'*********************************************************************************
    '���J��@���D
    Dim objText As AcadText
    Dim titlePT As Variant
    Dim ooPt As Variant
    titlePT = insPt
    titlePT(0) = titlePT(0) + (G.pagewidth / 2)
    titlePT(1) = titlePT(1) + G.fontsize * 5.5
    Set objText = ThisDrawing.ModelSpace.AddText(m_buf.getTITLE, titlePT, 6)
    ooPt = objText.insertionPoint
    objText.Layer = "TEXT"
    objText.Alignment = acAlignmentCenter
    objText.styleName = "��r"
    ooPt(0) = ooPt(0) + objText.insertionPoint(0)
    ooPt(1) = ooPt(1) + objText.insertionPoint(1)
    
    objText.Move objText.insertionPoint, ooPt
    
'*********************************************************************************
'�إߥD�n���ϼh
    Dim layerObj As AcadLayer
    Set layerObj = ThisDrawing.Layers.Add(MAINLAYER)


    Dim tmp_joinApp As New iArray
    Dim tmp_joinIds As New iArray
    'double lastTemp
    Dim tmp_track As Integer
    Dim tmp_track_item As Long
    Dim A_TEMPO_add As Integer

    Dim tmp_delaytime As Double  '�p���l �O�n�p�W�@�Ӧr������
    
    Dim tmp_xy As point
    
    Dim BNewObj As AcadBlockReference
    Dim tmp_name As String
    
    'Dim cst As String
    'Dim cst_no_fing As String
    Dim ptGripMid As Variant
    Dim s1 As MusicItem
    
    tmp_delaytime = 0
    
    tmp_name = G.fontName
    Dim tmp_erhu_fing As ErhuFing
    Dim midDownPt(2) As Double
    Dim mt_slur_left As MusicBlockGraphics
    Dim mt_slur_right As MusicBlockGraphics
    Dim mt_trip_left As MusicBlockGraphics
    Dim mt_trip_right As MusicBlockGraphics
    Dim plineSlur As AcadLWPolyline
    Dim plineTuplet As AcadLWPolyline
    Dim barConfig(1000) As aBarConfig
    Dim barId As Integer
    
    'Dim sumLine As Integer  '�p��ĴX��
    'Dim sumMeasure As Integer  '�p��C�檺�ĴX��
    
  
        For tmp_track = 0 To m_buf.GetTrackSize() - 1
            'NewObj = Nothing
'lin            tmp_joinApp.clear()
            tmp_delaytime = 0
            A_TEMPO_add = 1
            tmp_joinApp.Clear
            tmp_joinIds.Clear
            
            G.durationIndex = 0
            G.currLine = 0
            G.currMeasure = 0
            For tmp_track_item = 0 To m_buf.GetTrackBufferSize(tmp_track) - 1

                '�o�O�n�s���u���A�H m_Mete2 ���ɭ�
                If G.durationIndex >= (Cg.BLEN / G.mete2 * A_TEMPO_add) Then
                    If A_TEMPO_add = G.mete Then
                        A_TEMPO_add = 1
                    Else
                        A_TEMPO_add = A_TEMPO_add + 1
                    End If

                    If tmp_joinIds.Count >= 1 Then
                        Dim pp As Long
                        'ReDim tmp_joinApp(tmp_joinIds.Count - 1)
'                        For pp = 0 To tmp_joinIds.Count - 1
'                            Set tmp_joinApp(pp) = tmp_joinIds(pp)
'                        Next
                        tmp_joinApp.PushArray tmp_joinIds
                        MBG.addMusicJoin tmp_joinApp
                        tmp_joinApp.Clear
                    End If
                    tmp_joinIds.Clear

                End If

                If (G.durationIndex >= (Cg.BLEN / G.mete2) * G.mete) Then
                    G.currMeasure = G.currMeasure + 1
                    G.durationIndex = 0

                    If G.currMeasure >= G.barsperstaff Then  '�ݨC��p�`�Ʀ��S���W�L
                        G.currMeasure = 0
                        G.currLine = G.currLine + 1
                    End If
                End If


                Set s1 = m_buf.GetData(tmp_track, tmp_track_item)
                If (s1 Is Nothing) Then GoTo CallBackFor
                Select Case s1.typs
                   Case Cg.Config:
                        If (s1.barsperstaff >= 1 And tmp_track = 0) Then
                            G.barsperstaff = s1.barsperstaff
                            Set barConfig(barId) = New aBarConfig
                            barConfig(barId).barId = barId
                            barConfig(barId).barLineQuantity = s1.barsperstaff
                        If (s1.setbarstaffid > 1) Then
                            Set barConfig(setbarstaffid) = New aBarConfig
                            barConfig(setbarstaffid).barId = setbarstaffid
                            barConfig(setbarstaffid).barLineQuantity = s1.setbarstaff
                        End If
                        GoTo CallBackFor
                   Case Cg.meter:
                        If (tmp_track = 0) Then
                            G.mete = s1.mete
                            G.mete2 = s1.mete2
                            barConfig(barId).barId = barId
                            barConfig(barId).mete = s1.mete
                            barConfig(barId).mete2 = s1.mete2
                            
                        End If
                        
                        GoTo CallBackFor
                   Case Cg.Rest, Cg.note:
                   Case Else
                End Select


                If s1.notes(0).mnote = " " Or s1.notes(0).mnote = "" Then
                    GoTo CallBackFor
                End If

'*******���J�C�檺�p�`�u'**************************************************************************************'

                Dim ppnt As New point
                Dim bbo As Boolean
                ppnt.a insPt
                Call atDraw_BarLine(ppnt, tmp_track, G.currLine, G.currMeasure, G.durationIndex)

                
                If s1.notes(0).mnote = "." Then
                    Set tmp_xy = atBarXYpos(ppnt, tmp_track, G.currLine, G.currMeasure, G.durationIndex, True)
                Else
                    Set tmp_xy = atBarXYpos(ppnt, tmp_track, G.currLine, G.currMeasure, G.durationIndex, False)
                    'Set tmp_xy = atTableDraw(ppnt, tmp_track, tmp_AllTempo, False)
                End If



                Dim atPt As Variant
                atPt = tmp_xy.at
'**************************************************************************************'
'  ���JMusicText ����
'**************************************************************************************'
                If Me.chkOption1 = True Then
                    '(�j���)
                    '�϶��Ϊ���)
                    ppnt.a atPt
                    MBG.setDataText ppnt, s1, G.fontsize
                    Set BNewObj = MBG.InsterEnt '���J���ŤΫ��k



                ElseIf Me.chkOption2 = True Then

                '(�G�J��)
                '�o�O�S�����k��
'                    AMT.iTONE = 1        ' * ��
'                    AMT.iFinge = 2    '�o�O���k��     _+)(*&
'                    AMT.iScale = 3    '�o�O���C����   .,:
'                    AMT.iNote = 4     '�o���D��       1234567.|l
'                    AMT.iTempo = 5    '�o�O��l       -=368acefz
'                    AMT.iTowFinge = 6    '�o�O���k��ĤG��  _+)(*&
'                    AMT.iSlur = 7        ' �s���Ŧ�    (3456)


                    'Call NewObj.setData(atPt, cst_no_fing, G.fontName, G.FontSize)
                    'NewObj.Layer = "main"
                    tmp_erhu_fing.fing1 = 0

                    tmp_erhu_fing.Push = ""
                    tmp_erhu_fing.InOut = ""

                    '���o����r -���k1
                    Select Case s1.notes(0).mfingering
                       Case "0": tmp_erhu_fing.fing1 = tmp_erhu_fing.fing1 + amt.f�ũ� '�ũ�
                       Case "1": tmp_erhu_fing.fing1 = tmp_erhu_fing.fing1 + amt.f�@��
                       Case "2": tmp_erhu_fing.fing1 = tmp_erhu_fing.fing1 + amt.f�G��
                       Case "3": tmp_erhu_fing.fing1 = tmp_erhu_fing.fing1 + amt.f�T��
                       Case "4": tmp_erhu_fing.fing1 = tmp_erhu_fing.fing1 + amt.f�|��
                       Case "W", "w": tmp_erhu_fing.fing1 = tmp_erhu_fing.fing1 + amt.f�|��

                       Case "E", "e": tmp_erhu_fing.Push = "��"
                       Case "V", "v": tmp_erhu_fing.Push = "��"
                       Case "Q", "q": tmp_erhu_fing.InOut = "��"
                       Case "A", "a": tmp_erhu_fing.InOut = "�~"
                       Case Else
                    End Select

                    '���o����r -���k2
                    Select Case s1.notes(0).mtow_fingering
                       Case "0": tmp_erhu_fing.fing1 = tmp_erhu_fing.fing1 + amt.f�ũ� '�ũ�
                       Case "1": tmp_erhu_fing.fing1 = tmp_erhu_fing.fing1 + amt.f�@��
                       Case "2": tmp_erhu_fing.fing1 = tmp_erhu_fing.fing1 + amt.f�G��
                       Case "3": tmp_erhu_fing.fing1 = tmp_erhu_fing.fing1 + amt.f�T��
                       Case "4": tmp_erhu_fing.fing1 = tmp_erhu_fing.fing1 + amt.f�|��
                       Case "W", "w": tmp_erhu_fing.fing1 = tmp_erhu_fing.fing1 + amt.f�|��

                       Case "E", "e": tmp_erhu_fing.Push = "��"
                       Case "V", "v": tmp_erhu_fing.Push = "��"
                       Case "Q", "q": tmp_erhu_fing.InOut = "��"
                       Case "A", "a": tmp_erhu_fing.InOut = "�~"

                       Case Else
                    End Select

                    ppnt.x = atPt(0)
                    ppnt.y = atPt(1)
                    ppnt.Z = atPt(2)
                    MBG.setDataText ppnt, s1, G.fontsize
                    Set BNewObj = MBG.InsterEnt '���J���ŤΫ��k


                    '���J���k ����(�G�J��)
                    InsertErhuFinge ppnt, tmp_erhu_fing, G.fontsize

                End If

'*******���J��ƽu'**************************************************************************************'


                'AMT.iSlur = 7        ' �s���Ŧ�    (3456)
                If s1.slurStart = True Then
                    'Set mt_slur_left = New MusicBlockGraphics
                    Set mt_slur_left = MBG.copy

                ElseIf s1.slurEnd = True Then

                    Set plineSlur = MBG.drawSlur(mt_slur_left, MBG)

                End If
                
                
'*****************************************************************
                '�s���u��

                tmp_joinIds.Push BNewObj
                Set BNewObj = Nothing

                Select Case (s1.notes(0).mnote)
                Case "|"
                    tmp_delaytime = 0
                Case "-"
                    tmp_delaytime = PARTITION_DEF
                Case "."
                    tmp_delaytime = tmp_delaytime / 2
                Case " "

                Case Else

                    Dim tempo_hj As String
                    Dim tempo_ll As Variant

                    tempo_hj = " -2=45368aAcCfFgGzZ"
                    tempo_ll = Array(1, 2, 2, 4, 4, 5, 3, 6, 8, 10, 10, 12, 12, 15, 15, 16, 16, 32, 32)

                    Dim ii As Integer
                    Dim cn As String
                    For ii = 0 To Len(tempo_hj) - 1
                        cn = Mid(tempo_hj, ii + 1, 1)
                        If cn = s1.notes(0).mtempo Then
                            tmp_delaytime = PARTITION_DEF / tempo_ll(ii)
                            Exit For
                        ElseIf s1.notes(0).mtempo = "" Then
                            tmp_delaytime = PARTITION_DEF
                            Exit For
                        Else
                            tmp_delaytime = 0
                        End If
                    Next ii
                End Select


                G.durationIndex = G.durationIndex + CInt(Fix(tmp_delaytime))
                
CallBackFor:
            Next
        Next

End Sub

Private Function InsertMusicText(insertionPoint() As Double, cst As String, size As Double)
'���J����
    Dim textObj As AcadText
    Dim blockRefObj  As AcadBlockReference
    Dim textString As String
    Dim alignmentPoint(0 To 2) As Double
    Dim height As Double
    Dim midDownPt(0 To 2) As Double
    
    
    insertionPoint (0)
    insertionPoint (1)
    insertionPoint (2)
    
    Dim ipos As Integer
    Dim yAdd As Double
        '���o����r  �D�� AMT.iNote
    Select Case Mid(cst, amt.iNote, 1)
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
       Case " ": MFS.sNote = ""
       Case Else
            MFS.sNote = ""
    End Select
    Dim midGirPnt(2) As Double
    
    If MFS.sNote <> "" Then
        Call ThisDrawing.ModelSpace.InsertBlock(insertionPoint, MFS.sNote, size, size, size, 0)
        

    End If

End Function

Private Sub CommandButton1_Click()
    Me.Hide
    TESTSTAR
End Sub
Public Sub TESTSTAR()
    '��m��r��
    Dim pt As Variant
    Dim insertionPoint(0 To 2) As Double
    
    ' Return a point using a prompt
    pt = ThisDrawing.Utility.GetPoint(, "\n��ܭn���J���I �GEnter insertion point: ")
    
    insertionPoint(0) = pt(0)
    insertionPoint(1) = pt(1)
    insertionPoint(2) = pt(2)
    InsertMusicStar insertionPoint, 3.5
End Sub




Private Function InsertErhuFinge(midDownPt As point, this_ef As ErhuFing, size As Double)
'���J�G�J���k
    Dim textObj As AcadText
    Dim blockRefObj  As AcadBlockReference
    Dim textString As String
    Dim insertionPoint(0 To 2) As Double, alignmentPoint(0 To 2) As Double
    Dim height As Double
    

    insertionPoint(0) = midDownPt.x
    insertionPoint(1) = midDownPt.y
    insertionPoint(2) = midDownPt.Z
    
    Dim ipos As Integer
    Dim yAdd As Double
    ipos = 0
    yAdd = 0.7 '���k���V�W�W�q
    If this_ef.fing1 <> 0 Then
        
        '�ݬO�_�O�϶�
        If this_ef.fing1 And amt.f�ũ� Then
            'Call ThisDrawing.ModelSpace.InsertBlock(insertionPoint, "�G�J_��", 0.75, 0.75, 0.75, 0)
            textString = "\U+5B80"
            height = size * 0.47
            alignmentPoint(0) = insertionPoint(0) + (amt.A_TEXT_WIDTH * size / 2)
            alignmentPoint(1) = insertionPoint(1) + size * 2.13 + (ipos * yAdd * size)
            alignmentPoint(2) = insertionPoint(2)
            
            Set textObj = ThisDrawing.ModelSpace.AddText(textString, insertionPoint, height)
            textObj.Alignment = acAlignmentCenter
            textObj.TextAlignmentPoint = alignmentPoint
            textObj.styleName = "����_�Ʀr"
            textObj.Layer = "�˹��Ÿ�"
        ipos = ipos + 1
        End If
        If this_ef.fing1 And amt.f�|�� Then
            'textString = "\U+5B80"
            height = size * 0.47
            alignmentPoint(0) = insertionPoint(0) + (amt.A_TEXT_WIDTH * size / 2)
            alignmentPoint(1) = insertionPoint(1) + size * 2.13 + (ipos * yAdd * size)
            alignmentPoint(2) = insertionPoint(2)
            
            Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(alignmentPoint, "�G�J_ݬ�|", 1#, 1#, 1#, 0)
            blockRefObj.Layer = "�˹��Ÿ�"
        ipos = ipos + 1
        End If
        '���O �N���J��r
        textString = ""
        If this_ef.fing1 And amt.f�@�� Then
            textString = "��"
        ElseIf this_ef.fing1 And amt.f�G�� Then
            textString = "��"
        ElseIf this_ef.fing1 And amt.f�T�� Then
            textString = "��"
        ElseIf this_ef.fing1 And amt.f�|�� Then
            textString = "��"
        End If
        If textString <> "" Then
            height = size * 0.47
            alignmentPoint(0) = insertionPoint(0) + (amt.A_TEXT_WIDTH * size / 2)
            alignmentPoint(1) = insertionPoint(1) + size * 2.13 + (ipos * yAdd * size)
            alignmentPoint(2) = insertionPoint(2)
            
            ' Create the text object in model space
            Set textObj = ThisDrawing.ModelSpace.AddText(textString, insertionPoint, height)
            textObj.Alignment = acAlignmentCenter
            textObj.TextAlignmentPoint = alignmentPoint
            textObj.styleName = "����_�Ʀr"
            textObj.Layer = "SimpErhu�Ÿ�"
            ipos = ipos + 1
        End If
        
    End If
    
    If this_ef.InOut <> "" Then
    '���~
        textString = this_ef.InOut
        height = size * 0.47
        
        alignmentPoint(0) = insertionPoint(0) + (amt.A_TEXT_WIDTH * size / 2)
        alignmentPoint(1) = insertionPoint(1) + (size * 2.13) + (ipos * yAdd * size)
        alignmentPoint(2) = insertionPoint(2)
        
        
        ' Create the text object in model space
        Set textObj = ThisDrawing.ModelSpace.AddText(textString, insertionPoint, height)
        textObj.Alignment = acAlignmentCenter
        textObj.TextAlignmentPoint = alignmentPoint
        textObj.styleName = "����_�Ʀr"
        textObj.Layer = "�˹��Ÿ�"
        ipos = ipos + 1
    End If
    
        


    If this_ef.Push <> "" Then

        alignmentPoint(0) = insertionPoint(0) + (amt.A_TEXT_WIDTH * size / 2)
        alignmentPoint(1) = insertionPoint(1) + (size * 2.13) + (ipos * yAdd * size)
        alignmentPoint(2) = insertionPoint(2)

        If this_ef.Push = "��" Then
            textString = "b��"
        ElseIf this_ef.Push = "��" Then
            textString = "b��"
        End If
        
        Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(alignmentPoint, textString, size, size, size, 0)
       
        'blockRefObj.styleName = "SimpErhu"
        blockRefObj.Layer = "SimpErhu�Ÿ�"
'
        ipos = ipos + 1
    End If
    
    
'    If this_ef.InOut <> "" Then
'    '�� �~
'       '���O �N���J��r
'            textString = this_ef.InOut
'            height = size * 0.47
'            alignmentPoint(0) = insertionPoint(0)
'            alignmentPoint(1) = insertionPoint(1) + (size * 2.13) + (ipos * yAdd * size)
'            alignmentPoint(2) = insertionPoint(2)
'            ' Create the text object in model space
'            Set textObj = ThisDrawing.ModelSpace.AddText(textString, insertionPoint, height)
'            textObj.Alignment = acAlignmentCenter
'            textObj.TextAlignmentPoint = alignmentPoint
'            textObj.styleName = "����_�Ʀr"
'            textObj.Layer = "�˹��Ÿ�"
''
'        ipos = ipos + 1
'    End If
'



End Function


Private Function atBarXYpos(ByVal the_pt As point, ByVal the_track As Integer, _
 ByVal the_line As Integer, ByVal the_measure As Integer, ByVal the_alltempo As Double, ByVal the_isDorp As Boolean) As point
'���o�{�b�������m
'atBarXYpos()
'*the_pt ���I
'*the_the_track �{�b�O�ĴX�y
'*the_line �{�b�O�ĴX��
'*the_measure �{�b�O�ĴX�p�`
'*the_allTempo �{�b������O�h��
'*the_isDorp �{�b�O�_�O���I����
'


    
    
    'Dim ROW_ALL_DEF As Integer
    Dim MeasureDEF As Integer
    'Dim tmp_modTempo As Integer
    'Dim row As Integer


    Dim tmp_modCol As Integer
    Dim Col As Double
    Dim col_b As Integer
    Dim tmp_barSpaceWidth As Double '�@�`���Z��
    Dim tmp_rowspacing As Double '�C���C�檺�Z��
    Dim tmp_NoteDist As Double   '�C�筵�Ū��۹��m
    

    tmp_barSpaceWidth = (G.pagewidth - G.LeftSpace - G.RightSpace) / G.barsperstaff   'x �b��
    tmp_NoteDist = (tmp_barSpaceWidth - G.BarToNoteSpace * 2) / G.mete
    tmp_rowspacing = (G.TrackToTrack * (G.Many - 1) + G.LineToLine)             'y �b��

    MeasureDEF = (G.mete * (Cg.BLEN / G.mete2))   '�p��C�p�`���h�ֳ�� (4������Ŧ��h�ֳ��)
    'tmp_modTempo = the_alltempo Mod MeasureDEF '���������`���� ��A���٦��h�֪� ����
    '

    tmp_modCol = tmp_barSpaceWidth Mod G.mete2
    Col = the_alltempo / (Cg.BLEN / G.mete2)  '���o�{�b�p�`���ĴX��
    col_b = (Col Mod G.mete)  '���o�C�p�`���ĴX��

    
    
    

    Static lastPoint As New point
    Dim pp As New point
    'col �O�C�@�檺�ĴX��A�O�H�@�笰���Ӽ�
    'tmp_modCol �O�C�@�窺�ĴX�Ӧr����m
    pp.x = G.LeftSpace + the_measure * tmp_barSpaceWidth + G.BarToNoteSpace  '��X�p�`����m + �p�`�u�쭵�Ū��ť�
    'pp.x = pp.x + (CDbl(tmp_modCol) / CDbl(PARTITION_DEF / (4))) * (amt.LINE_LEN * G.FontSize) '���� 4�����@�� �����D
    pp.x = pp.x + CDbl(Col) * tmp_NoteDist  '���� 4�����@�� �����D

    pp.y = (G.TrackToTrack) * the_track + ((G.TrackToTrack) * (G.Many - 1) + G.LineToLine) * the_line
    pp.y = -pp.y


    '�C�`����A���L��
    '��  1 5    2 6             1 5  2 6
    '    ----   ----  ->�e�i��  ---- ----
    '    123456789AB            123456789AB
    pp.x = pp.x + CDbl(Col) * G.Beat_MIN_X '�L��
    '�C�@����A���L��
    '��  1  5 3   2  6 4             1 53     2 64
    '    ---====  ---====  ->�e�i��  --==     --==
    '    123456789ABCDEF             123456789ABCDEF
    Static ismodcol As Integer
'    If tmp_modCol > 0 Then
'        ismodcol = ismodcol + 1
'        pp.x = pp.x + ismodcol * G.MIN_X  '�L��
'    Else
'        ismodcol = 0
'    End If

    If the_isDorp Then '�p�G�O���I���šA�N�e��@�b
        pp.x = (lastPoint.x + pp.x) / 2
    End If
    
    '�O�_2�Ӧr�Ӫ�����F
    '����1�Ӧr���e��
    If Abs(lastPoint.x - pp.x) <= amt.A_TEXT_WIDTH * G.fontsize Then
        pp.x = lastPoint.x + amt.A_TEXT_WIDTH * G.fontsize * 1.05
    End If
    
    Set lastPoint = pp  '�s�J�̫᪺�I
    
    
    Set atBarXYpos = New point
    atBarXYpos.x = pp.x + the_pt.x
    atBarXYpos.y = pp.y + the_pt.y
End Function

Private Function atDraw_BarLine(ByVal the_pt As point, ByVal the_track As Integer, _
 ByVal the_line As Integer, ByVal the_measure As Integer, ByVal the_alltempo As Double)
   
    Dim tmp_pLWPoly As AcadPolyline
    'Dim initPoint As New AcGePoint2d()
    Dim startPt As New point
    Dim endPt As New point
    Dim ptlist As New PointList
    
    Dim tmp_trackitem As Integer
    Dim tmp_bardist As Double '�@�`���Z��
    Dim tmp_rowspacing As Double '�C���C�檺�Z��
    
    If the_measure = 0 And the_alltempo = 0 Then
    '�C�檺�Ĥ@�p�`�A�e�X�C�檺�p�`�u
        tmp_bardist = (G.pagewidth - G.LeftSpace - G.RightSpace) / G.barsperstaff
        tmp_rowspacing = (G.TrackToTrack * (G.Many - 1) + G.LineToLine)
        '�o�O�n�e�p�`�u
        'ROW_ALL_DEF �O�C�檺�`�������Ӧh�֡A�p �C�p�`2/4��A�C�榳 7 �Ӥp�`�A�h�C�檺 �`�������ӬO (240*2)*7=3360
        'row �ĴX��
        'colunm �ĴX��(�b�ĴX�檺�ĴX��)
        Dim ROW_ALL_DEF As Integer
        Dim tmp_modTempo As Integer
        'Dim row As Integer
    
        Dim tmp_modCol As Integer
        Dim Col As Integer
        Dim col_b As Integer
    
    
        ROW_ALL_DEF = (G.barsperstaff * G.mete * PARTITION_DEF / (G.mete2 / 4))
        tmp_modTempo = the_alltempo Mod ROW_ALL_DEF '���������`���� ��A���٦��h�֪� ����
        'row = (the_alltempo - tmp_modTempo) \ ROW_ALL_DEF
    
        tmp_modCol = tmp_modTempo Mod (PARTITION_DEF / (G.mete2 / 4))
        Col = (tmp_modTempo - tmp_modCol) / (PARTITION_DEF / (G.mete2 / 4)) '���o�C�檺�ĴX��
        col_b = (Col Mod G.mete)  '���o�C�p�`���ĴX��
        
        If the_track = 0 Then '�u���Ĥ@�y�~�n�e
            
            '�o�O�b�e�Ĥ@�p�`
    
            startPt.x = the_pt.x + G.LeftSpace
            startPt.y = -((amt.LINE_PASE + amt.DROP_UP) * G.fontsize) + tmp_rowspacing * the_line
            startPt.y = -startPt.y + the_pt.y

            endPt.x = the_pt.x + G.LeftSpace
            endPt.y = G.TrackToTrack * (G.Many - 1) + tmp_rowspacing * the_line
            endPt.y = -endPt.y + the_pt.y

            ptlist.Clear
            ptlist.Add startPt
            ptlist.Add endPt
            Set tmp_pLWPoly = ThisDrawing.ModelSpace.AddPolyline(ptlist.ToXYZList)
            tmp_pLWPoly.ConstantWidth = amt.BAR_WITCH * G.fontsize / 4.6
            tmp_pLWPoly.Layer = "bar"
        
            

            Dim j As Integer
            Dim barIndex As Integer
            For barIndex = 0 To G.barsperstaff - 1
                For j = 0 To G.Many - 1
    
    
                    startPt.x = the_pt.x + G.LeftSpace + tmp_bardist + barIndex * tmp_bardist
    
                    startPt.y = -((amt.LINE_PASE + amt.DROP_UP) * G.fontsize) + (G.TrackToTrack * j) + (tmp_rowspacing * the_line)
    
                    startPt.y = -startPt.y + the_pt.y
    
    
                    endPt.x = the_pt.x + G.LeftSpace + tmp_bardist + barIndex * tmp_bardist
    
                    endPt.y = (G.TrackToTrack * j) + (tmp_rowspacing * the_line)
    
                    endPt.y = -endPt.y + the_pt.y
    
                    ptlist.Clear
                    ptlist.Add startPt
                    ptlist.Add endPt
                    Set tmp_pLWPoly = ThisDrawing.ModelSpace.AddPolyline(ptlist.ToXYZList)
                    tmp_pLWPoly.ConstantWidth = amt.BAR_WITCH * G.fontsize / 4.6
                    tmp_pLWPoly.Layer = "bar"
                Next j
                'm_pLWPoly->setThickness(plineInfo.m_thick)
                'm_pLWPoly->setConstantWidth(plineInfo.m_width)

            Next barIndex
    
    
        End If
    End If
End Function





Private Function atTableDraw(ByVal the_pt As point, ByVal the_track As Integer, ByVal the_alltempo As Long, ByVal the_isDorp As Boolean) As point

'atTableDraw()
'*the_pt ���I
'*the_pt �{�b�O�ĴX�y
'*the_allTempo �{�b������O�h��
'*the_isDorp �{�b�O�_�O���I����
'


    atTableDraw_bar the_pt, the_track, the_alltempo
    
    Dim ROW_ALL_DEF As Integer
    Dim tmp_modTempo As Integer
    Dim row As Integer


    Dim tmp_modCol As Integer
    Dim Col As Integer
    Dim col_b As Integer


    ROW_ALL_DEF = (G.bar * G.mete * PARTITION_DEF / (G.mete2 / 4))
    tmp_modTempo = the_alltempo Mod ROW_ALL_DEF '���������`���� ��A���٦��h�֪� ����
    row = (the_alltempo - tmp_modTempo) \ ROW_ALL_DEF

    tmp_modCol = tmp_modTempo Mod (PARTITION_DEF / (G.mete2 / 4))
    Col = (tmp_modTempo - tmp_modCol) / (PARTITION_DEF / (G.mete2 / 4)) '���o�C�檺�ĴX��
    col_b = (Col Mod G.mete)  '���o�C�p�`���ĴX��

    '�ǥX�C�`���۹��m
    Dim tmp_xbarInterval As Double
    tmp_xbarInterval = (G.pagewidth - G.LeftSpace - G.RightSpace) / ((G.bar * G.mete))

    Static lastPoint As New point
    Dim pp As New point
    'col �O�C�@�檺�ĴX��A�O�H�@�笰���Ӽ�
    'tmp_modCol �O�C�@�窺�ĴX�Ӧr����m
    pp.x = G.LeftSpace + G.BarToNoteSpace + Col * tmp_xbarInterval
    pp.x = pp.x + (CDbl(tmp_modCol) / CDbl(PARTITION_DEF / (4))) * (amt.LINE_LEN * G.fontsize) '���� 4�����@�� �����D

    pp.y = (G.TrackToTrack) * the_track + ((G.TrackToTrack) * (G.Many - 1) + G.LineToLine) * row
    pp.y = -pp.y
    pp.x = pp.x + the_pt.x

    '�C�`����A���L��
    '��  1 5    2 6             1 5  2 6
    '    ----   ----  ->�e�i��  ---- ----
    '    123456789AB            123456789AB
    pp.x = pp.x + CDbl(col_b) * G.Beat_MIN_X '�L��
    '�C�@����A���L��
    '��  1  5 3   2  6 4             1 53     2 64
    '    ---====  ---====  ->�e�i��  --==     --==
    '    123456789ABCDEF             123456789ABCDEF
    Static ismodcol As Integer
    If tmp_modCol > 0 Then
        ismodcol = ismodcol + 1
        pp.x = pp.x + ismodcol * G.MIN_X  '�L��
    Else
        ismodcol = 0
    End If
    pp.y = pp.y + the_pt.y
    If the_isDorp Then '�p�G�O���I���šA�N�e��@�b
        pp.x = (lastPoint.x + pp.x) / 2
    End If
    
    '�O�_2�Ӧr�Ӫ�����F
    '����1�Ӧr���e��
    If Abs(lastPoint.x - pp.x) <= amt.A_TEXT_WIDTH * G.fontsize Then
        pp.x = lastPoint.x + amt.A_TEXT_WIDTH * G.fontsize * 1.05
    End If
    
    Set lastPoint = pp
    Set atTableDraw = pp
End Function

Private Function atTableDraw_bar(ByVal the_pt As point, ByVal the_track As Integer, ByVal the_alltempo As Long)
   
    Dim tmp_pLWPoly As AcadPolyline
    'Dim initPoint As New AcGePoint2d()
    Dim startPt As New point
    Dim endPt As New point
    Dim ptlist As New PointList
    
    Dim tmp_trackitem As Integer
    Dim tmp_bardist As Double '�@�`���Z��
    Dim tmp_rowspacing As Double '�C���C�檺�Z��

    tmp_bardist = (G.pagewidth - G.LeftSpace - G.RightSpace) / G.bar
    tmp_rowspacing = (G.TrackToTrack * (G.Many - 1) + G.LineToLine)
    '�o�O�n�e�p�`�u
    'ROW_ALL_DEF �O�C�檺�`�������Ӧh�֡A�p �C�p�`2/4��A�C�榳 7 �Ӥp�`�A�h�C�檺 �`�������ӬO (240*2)*7=3360
    'row �ĴX��
    'colunm �ĴX��(�b�ĴX�檺�ĴX��)
    Dim ROW_ALL_DEF As Integer
    Dim tmp_modTempo As Integer
    Dim row As Integer

    Dim tmp_modCol As Integer
    Dim Col As Integer
    Dim col_b As Integer


    ROW_ALL_DEF = (G.bar * G.mete * PARTITION_DEF / (G.mete2 / 4))
    tmp_modTempo = the_alltempo Mod ROW_ALL_DEF '���������`���� ��A���٦��h�֪� ����
    row = (the_alltempo - tmp_modTempo) \ ROW_ALL_DEF

    tmp_modCol = tmp_modTempo Mod (PARTITION_DEF / (G.mete2 / 4))
    Col = (tmp_modTempo - tmp_modCol) / (PARTITION_DEF / (G.mete2 / 4)) '���o�C�檺�ĴX��
    col_b = (Col Mod G.mete)  '���o�C�p�`���ĴX��
    
    If the_track = 0 Then '�u���Ĥ@�y�~�n�e
        If Col = 0 And tmp_modCol = 0 Then '�o�O�b�Ĥ@�p�`

            startPt.x = the_pt.x + G.LeftSpace
            startPt.y = -((amt.LINE_PASE + amt.DROP_UP) * G.fontsize) + tmp_rowspacing * row
            startPt.y = -startPt.y + the_pt.y

            endPt.x = the_pt.x + G.LeftSpace
            endPt.y = G.TrackToTrack * (G.Many - 1) + tmp_rowspacing * row
            endPt.y = -endPt.y + the_pt.y

            ptlist.Clear
            ptlist.Add startPt
            ptlist.Add endPt
            Set tmp_pLWPoly = ThisDrawing.ModelSpace.AddPolyline(ptlist.ToXYZList)
            tmp_pLWPoly.ConstantWidth = amt.BAR_WITCH * G.fontsize / 4.6
            tmp_pLWPoly.Layer = "bar"
        End If

        If (Col Mod G.mete) = 0 And (tmp_modCol = 0) Then
            Dim j As Integer
            For j = 0 To G.Many - 1


                startPt.x = the_pt.x + G.LeftSpace + tmp_bardist + Col / G.mete * tmp_bardist

                startPt.y = -((amt.LINE_PASE + amt.DROP_UP) * G.fontsize) + (G.TrackToTrack * j) + (tmp_rowspacing * row)

                startPt.y = -startPt.y + the_pt.y


                endPt.x = the_pt.x + G.LeftSpace + tmp_bardist + Col / G.mete * tmp_bardist

                endPt.y = (G.TrackToTrack * j) + (tmp_rowspacing * row)

                endPt.y = -endPt.y + the_pt.y

                ptlist.Clear
                ptlist.Add startPt
                ptlist.Add endPt
                Set tmp_pLWPoly = ThisDrawing.ModelSpace.AddPolyline(ptlist.ToXYZList)
                tmp_pLWPoly.ConstantWidth = amt.BAR_WITCH * G.fontsize / 4.6
                tmp_pLWPoly.Layer = "bar"
            Next j
            'm_pLWPoly->setThickness(plineInfo.m_thick)
            'm_pLWPoly->setConstantWidth(plineInfo.m_width)

        End If


    End If

End Function





Private Sub UserForm_Initialize()
    Set rTime = New runTime
    AcadConnect

    Dim dd As Double
    Dim i As Integer
    '�r���]�w
    Me.cobFontName.AddItem "EUDC"
    Me.cobFontName.AddItem "�ө���"
    Me.cobFontName.AddItem "�رd�����"
    Me.cobFontName.AddItem "�з���"
    
    '�r���j�p�]�w
    For dd = 2 To 8 Step 0.5
        Me.cobFontSize.AddItem dd
    Next
    
    '�X�n���]�w
    For i = 1 To 8
        Me.cobMany.AddItem i
    Next
    
    '�X��]�w
    Me.cobMete.AddItem "1/4"
    Me.cobMete.AddItem "2/4"
    Me.cobMete.AddItem "3/4"
    Me.cobMete.AddItem "4/4"
    Me.cobMete.AddItem "1/8"
    Me.cobMete.AddItem "3/8"
    Me.cobMete.AddItem "5/8"
    Me.cobMete.AddItem "7/8"
    
    
    
    '�C��n�X�p�`�]�w
    For i = 2 To 12
        Me.cobBar.AddItem i
    Next
    
    Me.cobFontName.text = "EUDC"
    Me.cobFontSize.text = 4
    Me.cobMany.text = 1
    Me.cobMete.text = "4/4"
    Me.cobBar.text = 4

    Me.tbPageWidth = 210
    Me.tbLeftSpace = 12
    Me.tbRightSpace = 12
    Me.tbBarToNote = 2
    Me.tbTrackToTrack = 14
    Me.tbLineToLine = 19
    Me.tbMIN_X = 0.25
    
    AMT_LOAD '�o�ӭ��n�A�]�w��l���
End Sub

Sub setLayoutMusicItem()


            G.currMeasure = 0
            For tmp_track_item = 0 To m_buf.GetTrackBufferSize(tmp_track) - 1

                '�o�O�n�s���u���A�H m_Mete2 ���ɭ�
                If G.durationIndex >= (PARTITION_DEF / (G.mete2 / 4) * A_TEMPO_add) Then
                    If A_TEMPO_add = G.mete Then
                        A_TEMPO_add = 1
                    Else
                        A_TEMPO_add = A_TEMPO_add + 1
                    End If

                    If tmp_joinIds.Count >= 1 Then
                        Dim pp As Long
                        'ReDim tmp_joinApp(tmp_joinIds.Count - 1)
'                        For pp = 0 To tmp_joinIds.Count - 1
'                            Set tmp_joinApp(pp) = tmp_joinIds(pp)
'                        Next
                        tmp_joinApp.PushArray tmp_joinIds
                        MBG.addMusicJoin tmp_joinApp
                        tmp_joinApp.Clear
                    End If
                    tmp_joinIds.Clear

                End If

                If (G.durationIndex >= PARTITION_DEF * G.mete) Then
                    G.currMeasure = G.currMeasure + 1
                    G.durationIndex = 0

                    If G.currMeasure >= G.barsperstaff Then  '�ݨC��p�`�Ʀ��S���W�L
                        G.currMeasure = 0
                        G.currLine = G.currLine + 1
                    End If
                End If


                Set s1 = m_buf.GetData(tmp_track, tmp_track_item)

                Select Case s1.typs
                   Case Cg.bar:
                        G.barsperstaff = s1.barsperstaff
                        GoTo CallBackFor
                   Case Cg.meter:
                        G.mete = s1.mete
                        G.mete2 = s1.mete2
                        GoTo CallBackFor
                   Case Cg.Rest, Cg.note:
                   Case Else
                End Select


End Sub

Function layoutMusicItem(spacing As Double, musicGroup()) As Double

    Dim currentduration
    Dim durationIndex   As Double
    Const Epsilon As Double = 0.0000001
    Dim spacingunit As Double
    Dim spacingUnits As Double
    Dim minSpace As Double
    Dim x As Double
    Dim i As Integer, j As Integer, k As Integer
    
    
    minSpace = 1000
    '�o�j��O�]�w X �b�V
    Do While (finished(staffGroup.voices) = False)   ' Inner loop.
       Dim currVoice As VoiceElement
       Set currVoice = staffGroup.voices(1)
       Debug.Print currVoice.i
       
        
        '' ���n�b���n�����Կ�̤����G�m���Ĥ@�ӫ���ɶ��ŧO
        currentduration = Empty '' candidate smallest duration level
        For i = 0 To staffGroup.voices.Count - 1
            If currentduration = Empty Then
                currentduration = getDurationIndex(staffGroup.voices(i))
            Else
                If getDurationIndex(staffGroup.voices(i)) < currentduration Then
                    currentduration = getDurationIndex(staffGroup.voices(i))
                End If
            End If
        Next
        
        
              
        
        ''isolate voices at current duration level
        ''�j���ثe����ɶ����Ū��y��
        Dim currentvoices As New iArray ' VoiceElement[] = []
        Dim othervoices As New iArray ' VoiceElement[] = []
        currentvoices.Clear
        othervoices.Clear
        For i = 0 To staffGroup.voices.Count - 1
            durationIndex = getDurationIndex(staffGroup.voices(i))
            '' PER: Because of the inexactness of JS floating point math, we just get close.
            '' PER�G�ѩ� JS �B�I�ƾǪ�����T�ʡA�ڭ̥u�O����Ӥw�C
            If (durationIndex - currentduration > Epsilon) Then
                othervoices.Push staffGroup.voices(i)
                ''console.log("out: voice ",i)
             Else
                currentvoices.Push staffGroup.voices(i)
                ''if (debug) console.log("in: voice ",i)
            
            End If
        Next
        
         
        

        
        '' among the current duration level find the one which needs starting furthest right
        '' �b�ثe����ɶ��ŧO�����ݭn�q�̥k��}�l������ɶ��ŧO
        spacingunit = 0 '' number of spacingunits coming from the previously laid out element to this one
        Dim spacingduration As Double
        For i = 0 To currentvoices.Count - 1
            
            If (layoutVoiceElement.getNextX(currentvoices(i)) > x) Then
                x = layoutVoiceElement.getNextX(currentvoices(i))
                spacingunit = layoutVoiceElement.getSpacingUnits(currentvoices(i))
                spacingduration = currentvoices(i).spacingduration
            End If
        Next
        spacingUnits = spacingUnits + spacingunit
        minSpace = Math.min(Array(minSpace, spacingunit))
        

        Dim lastTopVoice
        For i = 0 To currentvoices.Count - 1
            Dim v As VoiceElement
            Dim topVoice As VoiceElement
            Dim voicechildx As Double
            Dim dx As Double
            Set v = currentvoices(i)
            If (v.voicenumber = 0) Then lastTopVoice = i
            If lastTopVoice <> Empty And currentvoices(lastTopVoice).voicenumber <> v.voicenumber Then
                Set topVoice = currentvoices(lastTopVoice)
            Else
                Set topVoice = Nothing
            End If
            ''line ������ if (~isSameStaff(v, topVoice)) then   Set topVoice = Empty
            voicechildx = layoutVoiceElement.layoutOneItem(x, spacing, v, 0, topVoice)
            dx = voicechildx - x
            ''�o�O�ݬO�_���e�ʭ�
            ''�p�G���A���������ŴN�b�[�e�ʭ����Z��
            If (dx > 0) Then
                x = voicechildx ''update x
                For j = 0 To i '' shift over all previously laid out elements
                    Call layoutVoiceElement.shiftRight(dx, currentvoices(j))
                Next
            End If
        Next

        '' remove the value of already counted spacing units in other voices
        '' (e.g.if a voice had planned to use up 5 spacing units but is not in line to be laid out at this duration level -
        '' where we've used 2 spacing units - then we must use up 3 spacing units, not 5)
        '' �R����L�y�����w�p�⪺���j��쪺��
        ''�]�Ҧp�A�p�G�@�ӻy���p���ϥ� 5 �Ӷ��j���A�����b������ɶ��ŧO�W�ƦC -
        '' �ڭ̨ϥΤF 2 �Ӷ��j��� - ����ڭ̥����Χ�3�Ӷ��Z���A�ӫD5 �ӡ^
        '' ���իᤣ�ݭn
        For i = 0 To othervoices.Count - 1
            othervoices(i).spacingduration = othervoices(i).spacingduration - spacingduration
            Call layoutVoiceElement.updateNextX(x, spacing, othervoices(i))   '' adjust other voices expectations
        Next
        
                    
              
        '' ��s�ثe�G������������
        For i = 0 To currentvoices.Count - 1
            Dim voice As VoiceElement
            Set voice = currentvoices(i)
            '' ��C�@�� voice.i �[ 1 ���U�@�Ӥl����
            '' �٦��ק� voice.durationindex �[�W�{�b�w�gŪ�� �����Ū���
            '' 4������=0.25 2������=0.5 ��������=1
            '' �C�@�p�`�`��(����)�� 1
            Call layoutVoiceElement.updateIndices(voice)
        Next
    Loop
    i = i + 1
End Function
