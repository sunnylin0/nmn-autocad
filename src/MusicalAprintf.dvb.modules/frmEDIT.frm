VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEDIT 
   Caption         =   "¹Ï¶ô¥Î EDIT"
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
'2024.11.28  isVirtualChar As Boolean '³]©wªÅ¤ß¦r²Å  v2.13
'2024.11.28  ¥[¤J I:setbar 5/3 ¦b¥xÀY³]©w "²Ä´X¤p¸`¶}©l/¨C¦æ´X¤p¸`"
'2024.11.27  v2.12 tuplet "{" "}" ¥[¤J 3 5 6 7 ³s­µªº¥\¯à
'2024.08.27  ­×§ï Å|­µªº°ª©³­µÂI¦ì¸m
'2024.08.23  "v2.1" ¥[¤J «eºó¤¸¯À '>'  «áºó¤¸¯À '<'
'2024.03.30  "v2.0" ­n¥[¤J ©ç¸¹ M:3/8 M:4/4 ¡A¤j§ï­µ²Å±Æ§Ç¤è¦¡
'2024.03.29  ­×§ï 4/4 3/8 ©ç¼Æ¹ï¦ì°ÝÃD
'2013.11.21  ­×§ï DataBuffer ¬°¤¸¯À
'            ¦h¥[ iAdd ¦X­µ¦æ
'2013.03.17  V3 ¥¿­n­×§ï ¤G­Jªºª©¥»¡A¦]µ{¦¡¤§«e¬O¥Î¥jºåªº«üªk¹Ï¡A²{¦b§ï¦¨¤G­Jªº«üªk¹Ï

Const version  As String = "v2.13" '³nÅé¸¹½X
Const c1 As Integer = 60   'C½Õ1ªºÁä¦W­È
'Const FOURPAINUM   As Integer = 64 '1/4­µ²Å­p¼Æ
'Const MIDICLOCK As Integer = 24   '¨C1/64­µ²ÅªºMIDICLOCK¼Æ
'Const TEMPO_DEF As Integer = 90   '¹w³]¨C¤ÀÄÁ90©ç
Const PARTITION_DEF As Integer = 384   '¹w³]¨C©ç¤À³Î¬°384­p®É³æ¤¸
Const VOLUME_DEF As Integer = 64
Const MAINLAYER As String = "MAIN"    '¥D­nªº¹Ï¼h
Dim partLayer As String     '°O¿ý²{¦b²Ä´X³¡



'¤G­J«üªk ª¬ºA
Private Type ErhuFing
    fing1 As Integer    '¢¹ ¢º ¢» ¢¼ "bªÅ²Í"
    fing2 As String  '¢¹ ¢º ¢» ¢¼ "bªÅ²Í"
    Push As String      '©Ô¡ä ±ÀV
    InOut As String     '¤º   ¥~
End Type

Dim m_buf As New DataBuffer
Dim TuneLines As New TuneLineList
Private constTE As New constructTuneElements

'1 ¦bvb¤uµ{¤¤¤Þ¥Îautocadªºƒ·
'2 ©w…óautocad†Á¶H
Private acadapp As AcadApplication
Private acadDoc As AcadDocument
'3 ‡À¥´…{‹×„¸autocadªº¨ç‡Û¡A¥H¤U¬O§Ú‡Àªº
'--------------------------------------------------------------
'„²±µCad
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
          MsgBox "¤£¯à†b¦æAutoCAD,ˆ[‰ä¬d¬O§_¦w†E¡I", vbOKCancel, "Äµ§i¡I"
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
    
    G.pagewidth = Me.tbPageWidth '¼e«×
    G.LeftSpace = Me.tbLeftSpace  '¥ªªÅ¥Õ
    G.RightSpace = Me.tbRightSpace  '¥kªÅ¥Õ
    G.BarToNoteSpace = Me.tbBarToNote    '¤p¸`¨ì­µ²Å
    G.TrackToTrack = Me.tbTrackToTrack  'Án³¡¶¡¶Z
    G.LineToLine = Me.tbLineToLine      '¨C¦æ¶¡¶Z
    G.check1 = True
    G.MIN_X = Me.tbMIN_X            '·L½Õ
    G.Beat_MIN_X = Me.tbBeat_min_x  '©ç·L½Õ
    G.IsBarAlign = Me.cbIsBarAlign
    G.isVirtualChar = Me.cbVirtualChar  'ªÅ¤ß¦r
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

    '¾ã²z ©ç¸¹¤Î¤p¸`¼Æ ¤À¦¨¤@¦æ¤@¦æªº TuneLines
    Set TuneLines = constTE.translate2Staffs(m_buf)
    
    'layout ±N x §¤¼Ð½Õ¾ã¨ì¨äµ´¹ï¦ì¸m
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
    'abc ´ú¸Õ

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
'³]©w½Õ¸¹ ­µ¦ì ©ç¸¹
    Dim staffEle As New Staff
    Dim cl As vClefProperties
    Dim ky As vKeySignature
    Dim mf As vMeter
    
    '¦s ½Õ¸¹
    Set cl = New vClefProperties
    cl.el_typs = "clef"
    cl.typs = "treble"
    cl.verticalPos = 0
    
    
    '¦s ­µ¦ì
    Dim vAcc As New vAccidental
    vAcc.acc = "sharp"
    vAcc.note = "f"
    vAcc.verticalPos = 10
    Set ky = New vKeySignature
    Set ky.accidentals = New iArray
    ky.el_typs = "keySignature"
    ky.root = "E"
    ky.accidentals.Push vAcc
    
    
    '©ç¸¹
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
    '³]©w --> "¹Ï¼h¦W¦r ÃC¦â"
    datalayer.PushArray Array( _
    "FIGE 6", _
    "TEXT 2", _
    "bar 181", _
    "¸Ë¹¢²Å¸¹ 4", _
    "¦ÛÃ¸½u 1", _
    "main 7", _
    "SimpErhu²Å¸¹ 151")

    
    '·s«Ø¹Ï¼h
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
    dataStyles(2, 1) = "¼ÐÅé"
    dataStyles(2, 2) = "KAIU.TTF"
    dataStyles(2, 3) = ""
    dataStyles(3, 1) = "MMP2005"
    dataStyles(3, 2) = ""
    dataStyles(3, 3) = ""
    dataStyles(4, 1) = "­µ²Å_¼Æ¦r"
    dataStyles(4, 2) = "SimSun.ttc"
    dataStyles(4, 3) = ""
    dataStyles(5, 1) = "«¬¦¡_²Ó©úÅé"
    dataStyles(5, 2) = "MingLiU.ttc"
    'dataStyles(5, 2) = "PMingLiU.ttf"
    dataStyles(5, 3) = ""
    dataStyles(6, 1) = "¦r²Å"
    dataStyles(6, 2) = "MAESTRO.TTF"
    dataStyles(6, 3) = ""
    dataStyles(7, 1) = "µØ±d²Ê¶Â"
    dataStyles(7, 2) = "DFFT_C7.ttc"
    'dataStyles(7, 2) = "DFLiHeiBold.ttf"
    dataStyles(7, 3) = ""
    dataStyles(8, 1) = "¤å¦r"
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
    '´¡¤J³]©w¸ê®Æ
    Dim mtxt As AcadMText
    Dim textObj As AcadText
    Dim textString As String
    Dim insertionPt As New point
    Dim height As Double
    
    ' Define the text object
    textString = version & vbCrLf
    textString = textString & "size " & the_G.fontsize & vbCrLf
    
    textString = textString & "¥ªªÅ¥Õ " & the_G.LeftSpace & "mm" & vbCrLf
    textString = textString & "¥kªÅ¥Õ " & the_G.RightSpace & "mm" & vbCrLf
    textString = textString & "¤p¸`¨ì­µ²Å " & the_G.BarToNoteSpace & "mm" & vbCrLf
    textString = textString & "Án³¡  " & the_G.TrackToTrack & "mm" & vbCrLf
    textString = textString & "¨C¦æ  " & the_G.LineToLine & "mm" & vbCrLf
    textString = textString & "·L½Õ  " & the_G.MIN_X & "mm" & vbCrLf
    textString = textString & "©ç·L½Õ " & the_G.Beat_MIN_X
    
    
    insertionPt.x = aPt.x - 30
    insertionPt.y = aPt.y
    insertionPt.Z = aPt.Z
    height = 3
    
    ' Create the text object in model space
    Set mtxt = ThisDrawing.ModelSpace.AddMText(insertionPt.ToDouble, height, textString)
    mtxt.width = 40
    mtxt.styleName = "Standard"
    
    '´¡¤J¸ê®Æ
    insertionPt.x = insertionPt.x - 3
    Set mtxt = ThisDrawing.ModelSpace.AddMText(insertionPt.ToDouble, 0.01, Me.TextBox8.text)
    mtxt.height = 0.01
    
End Sub
Private Sub drawLayoutStaff(abcTuneLines As TuneLineList)
    Dim MBG As New MusicBlockGraphics
    
    '©ñ¸m¤å¦r¥»
    Dim retPt As Variant
    Dim insPt As New point
    
    ' Return a point using a prompt
    retPt = ThisDrawing.Utility.GetPoint(, "\n¿ï¾Ü­n´¡¤JªºÂI ¡GEnter insertion point: ")
    insPt.a retPt
    '´¡¤J³]©w¸ê®Æ¤å¦r»¡©ú
    Call inst_G(G, insPt)
'***********************************************************************************
    'µe¥X©w¦ì½u-µe®Ø
    Dim plineObj As AcadPolyline
    Set plineObj = MBG.insterPositionBox(insPt, G)

'*********************************************************************************
    '´¡¤J³æ¤@¼ÐÃD
    Dim objText As AcadText
    Dim titlePT As New point
    Dim ooPt As New point
    
    titlePT.x = insPt.x + (G.pagewidth / 2)
    titlePT.y = insPt.y + G.fontsize * 5.5
    Set objText = ThisDrawing.ModelSpace.AddText(m_buf.getTITLE, titlePT.ToDouble, 6)
    ooPt.a objText.insertionPoint
    objText.Layer = "TEXT"
    objText.Alignment = acAlignmentCenter
    objText.styleName = "¤å¦r"
    ooPt.x = ooPt.x + objText.insertionPoint(0)
    ooPt.y = ooPt.y + objText.insertionPoint(1)
    
    objText.Move objText.insertionPoint, ooPt.ToDouble

'*********************************************************************************
'«Ø¥ß¥D­n ­µ²Å
    Dim layerObj As AcadLayer
    Set layerObj = ThisDrawing.Layers.Add(MAINLAYER)


    Dim tmp_joinApp As New iArray
    Dim tmp_joinIds As New iArray
    'double lastTemp
    Dim tmp_track As Integer
    Dim tmp_track_item As Long
    Dim A_TEMPO_add As Integer

    Dim tmp_delaytime As Double  '­pºâ©ç¤l ¬O­n­p¤W¤@­Ó¦rªºªø«×
    
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
    'Dim sumLine As Integer  '­pºâ²Ä´X¦æ
    'Dim sumMeasure As Integer  '­pºâ¨C¦æªº²Ä´X¦æ
    
''******Ã¸»s¤U©ç½u ÅÜ¼Æ
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
                        'Ã¸»s¤p¸`½u  ********************************
                        If j > 0 And k < 3 Then
                            '²Ä¤G³¡ªº²Ä¤@­Ó¤p¸`½u­n©Ôªø
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
                        
                        'Ã¸¤p¸`¸¹ ***********************************
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
                        'Ã¸»s©ç¸¹  ********************************
                        Dim metePt1 As New point
                        Dim metePt2 As New point
                        metePt2.x = currX + s1.x
                        metePt2.y = currY + 0.3 * G.fontsize
                        
                        metePt1.x = currX + s1.x
                        metePt1.y = currY + 1.3 * G.fontsize
                        
                        Set objText = ThisDrawing.ModelSpace.AddText(s1.mete2, metePt2.ToDouble, G.fontsize * 0.7)
                        objText.Layer = "¸Ë¹¢²Å¸¹":         objText.styleName = "­µ²Å_¼Æ¦r"
                        
                        Set objText = ThisDrawing.ModelSpace.AddText(s1.mete, metePt1.ToDouble, G.fontsize * 0.7)
                        objText.Layer = "¸Ë¹¢²Å¸¹":         objText.styleName = "­µ²Å_¼Æ¦r"
                        s1.oX = metePt2.x
                        s1.oY = metePt2.y
                        G.mete = s1.mete
                        G.mete2 = s1.mete2
                   Case Cg.Rest, Cg.note:
                        'Ã¸»s­µ²Å********************************
                        Dim N As Integer
                        startPt.x = currX + s1.x
                        startPt.y = currY
                        s1.oX = startPt.x
                        s1.oY = startPt.y
                        MBG.setDataText startPt, s1, G.fontsize
                        MBG.nowPartLayer = partLayer
                        Set BNewObj = MBG.InsterEnt '´¡¤J­µ²Å¤Î«üªk
                        
                        ''** ©ç½u­pºâ ***********************
                        nlast = s1.nflags
                        Temo1 = Fix(durationIndex / (Cg.BLEN / G.mete2)) ''¨ú±o¥¼Á`¥[¤§«e
                        durationIndex = durationIndex + s1.duration
                        
                        If (nlast > 0) Then
                            beamMuseList.Push s1
                             '³o¬O­n³sµ²½uªº¡A¥H m_Mete2 ¬°®É­È
                            Temo2 = durationIndex - (Temo1 * (Cg.BLEN / G.mete2))
                            
                            If Temo2 Mod (Cg.BLEN / G.mete2) = 0 Or Temo2 > (Cg.BLEN / G.mete2) Then
                                MBG.draw_dur beamMuseList
                                beamMuseList.Clear
                            End If
                         ElseIf beamMuseList.Count >= 1 Then
                            MBG.draw_dur beamMuseList
                            beamMuseList.Clear
                        End If
                        
                        '*******Ã¸»s¶ê·Æ½u'**************************************************************************************'
                        'AMT.iSlur = 7        ' ³s­µ²Å¦æ    (3456)
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
                '¶ê·Æ½u¬O§_¥¼µe§¹
                Set plineSlur = MBG.drawSlurToEnd(mt_slur_left, s1.oX + s1.w)
                Set mt_slur_left = Nothing
            End If
        

        Next
    Next
  


End Sub

Private Sub draw_many_text1()
  
  
    Dim MBG As New MusicBlockGraphics
    
    '©ñ¸m¤å¦r¥»
    Dim insPt As Variant
    Dim ipt As New point
    
    ' Return a point using a prompt
    insPt = ThisDrawing.Utility.GetPoint(, "\n¿ï¾Ü­n´¡¤JªºÂI ¡GEnter insertion point: ")
    '´¡¤J³]©w¸ê®Æ¤å¦r»¡©ú
    Call inst_G(G, insPt)
'***********************************************************************************
    'µe¥X©w¦ì½u-µe®Ø
    Dim plineObj As AcadPolyline
    ipt = insPt
    Set plineObj = MBG.insterPositionBox(ipt, G)

'*********************************************************************************
    '´¡¤J³æ¤@¼ÐÃD
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
    objText.styleName = "¤å¦r"
    ooPt(0) = ooPt(0) + objText.insertionPoint(0)
    ooPt(1) = ooPt(1) + objText.insertionPoint(1)
    
    objText.Move objText.insertionPoint, ooPt
    
'*********************************************************************************
'«Ø¥ß¥D­nªº¹Ï¼h
    Dim layerObj As AcadLayer
    Set layerObj = ThisDrawing.Layers.Add(MAINLAYER)


    Dim tmp_joinApp As New iArray
    Dim tmp_joinIds As New iArray
    'double lastTemp
    Dim tmp_track As Integer
    Dim tmp_track_item As Long
    Dim A_TEMPO_add As Integer

    Dim tmp_delaytime As Double  '­pºâ©ç¤l ¬O­n­p¤W¤@­Ó¦rªºªø«×
    
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
    
    'Dim sumLine As Integer  '­pºâ²Ä´X¦æ
    'Dim sumMeasure As Integer  '­pºâ¨C¦æªº²Ä´X¦æ
    
  
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

                '³o¬O­n³sµ²½uªº¡A¥H m_Mete2 ¬°®É­È
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

                    If G.currMeasure >= G.barsperstaff Then  '¬Ý¨C¦æ¤p¸`¼Æ¦³¨S¦³¶W¹L
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

'*******´¡¤J¨C¦æªº¤p¸`½u'**************************************************************************************'

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
'  ´¡¤JMusicText ª«¥ó
'**************************************************************************************'
                If Me.chkOption1 = True Then
                    '(¥jºå¥Î)
                    '¹Ï¶ô¥Îªþ¸¹)
                    ppnt.a atPt
                    MBG.setDataText ppnt, s1, G.fontsize
                    Set BNewObj = MBG.InsterEnt '´¡¤J­µ²Å¤Î«üªk



                ElseIf Me.chkOption2 = True Then

                '(¤G­J¥Î)
                '³o¬O¨S¦³«üªkªº
'                    AMT.iTONE = 1        ' * ¦æ
'                    AMT.iFinge = 2    '³o¬O«üªk¦æ     _+)(*&
'                    AMT.iScale = 3    '³o¬O°ª§C­µ¦æ   .,:
'                    AMT.iNote = 4     '³o¬°¥D¦æ       1234567.|l
'                    AMT.iTempo = 5    '³o¬O©ç¤l       -=368acefz
'                    AMT.iTowFinge = 6    '³o¬O«üªk¦æ²Ä¤G¦æ  _+)(*&
'                    AMT.iSlur = 7        ' ³s­µ²Å¦æ    (3456)


                    'Call NewObj.setData(atPt, cst_no_fing, G.fontName, G.FontSize)
                    'NewObj.Layer = "main"
                    tmp_erhu_fing.fing1 = 0

                    tmp_erhu_fing.Push = ""
                    tmp_erhu_fing.InOut = ""

                    '¨ú±oÃöÁä¦r -«üªk1
                    Select Case s1.notes(0).mfingering
                       Case "0": tmp_erhu_fing.fing1 = tmp_erhu_fing.fing1 + amt.fªÅ©¶ 'ªÅ©¶
                       Case "1": tmp_erhu_fing.fing1 = tmp_erhu_fing.fing1 + amt.f¤@«ü
                       Case "2": tmp_erhu_fing.fing1 = tmp_erhu_fing.fing1 + amt.f¤G«ü
                       Case "3": tmp_erhu_fing.fing1 = tmp_erhu_fing.fing1 + amt.f¤T«ü
                       Case "4": tmp_erhu_fing.fing1 = tmp_erhu_fing.fing1 + amt.f¥|«ü
                       Case "W", "w": tmp_erhu_fing.fing1 = tmp_erhu_fing.fing1 + amt.f´|©¶

                       Case "E", "e": tmp_erhu_fing.Push = "©Ô"
                       Case "V", "v": tmp_erhu_fing.Push = "±À"
                       Case "Q", "q": tmp_erhu_fing.InOut = "¤º"
                       Case "A", "a": tmp_erhu_fing.InOut = "¥~"
                       Case Else
                    End Select

                    '¨ú±oÃöÁä¦r -«üªk2
                    Select Case s1.notes(0).mtow_fingering
                       Case "0": tmp_erhu_fing.fing1 = tmp_erhu_fing.fing1 + amt.fªÅ©¶ 'ªÅ©¶
                       Case "1": tmp_erhu_fing.fing1 = tmp_erhu_fing.fing1 + amt.f¤@«ü
                       Case "2": tmp_erhu_fing.fing1 = tmp_erhu_fing.fing1 + amt.f¤G«ü
                       Case "3": tmp_erhu_fing.fing1 = tmp_erhu_fing.fing1 + amt.f¤T«ü
                       Case "4": tmp_erhu_fing.fing1 = tmp_erhu_fing.fing1 + amt.f¥|«ü
                       Case "W", "w": tmp_erhu_fing.fing1 = tmp_erhu_fing.fing1 + amt.f´|©¶

                       Case "E", "e": tmp_erhu_fing.Push = "©Ô"
                       Case "V", "v": tmp_erhu_fing.Push = "±À"
                       Case "Q", "q": tmp_erhu_fing.InOut = "¤º"
                       Case "A", "a": tmp_erhu_fing.InOut = "¥~"

                       Case Else
                    End Select

                    ppnt.x = atPt(0)
                    ppnt.y = atPt(1)
                    ppnt.Z = atPt(2)
                    MBG.setDataText ppnt, s1, G.fontsize
                    Set BNewObj = MBG.InsterEnt '´¡¤J­µ²Å¤Î«üªk


                    '´¡¤J«üªk ªþ¸¹(¤G­J¥Î)
                    InsertErhuFinge ppnt, tmp_erhu_fing, G.fontsize

                End If

'*******´¡¤J¶ê·Æ½u'**************************************************************************************'


                'AMT.iSlur = 7        ' ³s­µ²Å¦æ    (3456)
                If s1.slurStart = True Then
                    'Set mt_slur_left = New MusicBlockGraphics
                    Set mt_slur_left = MBG.copy

                ElseIf s1.slurEnd = True Then

                    Set plineSlur = MBG.drawSlur(mt_slur_left, MBG)

                End If
                
                
'*****************************************************************
                '³sµ²½u¥Î

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
'´¡¤J­µªþ
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
        '¨ú±oÃöÁä¦r  ¥D­µ AMT.iNote
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
    '©ñ¸m¤å¦r¥»
    Dim pt As Variant
    Dim insertionPoint(0 To 2) As Double
    
    ' Return a point using a prompt
    pt = ThisDrawing.Utility.GetPoint(, "\n¿ï¾Ü­n´¡¤JªºÂI ¡GEnter insertion point: ")
    
    insertionPoint(0) = pt(0)
    insertionPoint(1) = pt(1)
    insertionPoint(2) = pt(2)
    InsertMusicStar insertionPoint, 3.5
End Sub




Private Function InsertErhuFinge(midDownPt As point, this_ef As ErhuFing, size As Double)
'´¡¤J¤G­J«üªk
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
    yAdd = 0.7 '«üªkªº¦V¤W¼W¶q
    If this_ef.fing1 <> 0 Then
        
        '¬Ý¬O§_¬O¹Ï¶ô
        If this_ef.fing1 And amt.fªÅ©¶ Then
            'Call ThisDrawing.ModelSpace.InsertBlock(insertionPoint, "¤G­J_ªÅ", 0.75, 0.75, 0.75, 0)
            textString = "\U+5B80"
            height = size * 0.47
            alignmentPoint(0) = insertionPoint(0) + (amt.A_TEXT_WIDTH * size / 2)
            alignmentPoint(1) = insertionPoint(1) + size * 2.13 + (ipos * yAdd * size)
            alignmentPoint(2) = insertionPoint(2)
            
            Set textObj = ThisDrawing.ModelSpace.AddText(textString, insertionPoint, height)
            textObj.Alignment = acAlignmentCenter
            textObj.TextAlignmentPoint = alignmentPoint
            textObj.styleName = "­µ²Å_¼Æ¦r"
            textObj.Layer = "¸Ë¹¢²Å¸¹"
        ipos = ipos + 1
        End If
        If this_ef.fing1 And amt.f´|©¶ Then
            'textString = "\U+5B80"
            height = size * 0.47
            alignmentPoint(0) = insertionPoint(0) + (amt.A_TEXT_WIDTH * size / 2)
            alignmentPoint(1) = insertionPoint(1) + size * 2.13 + (ipos * yAdd * size)
            alignmentPoint(2) = insertionPoint(2)
            
            Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(alignmentPoint, "¤G­J_Ý¬´|", 1#, 1#, 1#, 0)
            blockRefObj.Layer = "¸Ë¹¢²Å¸¹"
        ipos = ipos + 1
        End If
        '¤£¬O ´N´¡¤J¤å¦r
        textString = ""
        If this_ef.fing1 And amt.f¤@«ü Then
            textString = "¢¹"
        ElseIf this_ef.fing1 And amt.f¤G«ü Then
            textString = "¢º"
        ElseIf this_ef.fing1 And amt.f¤T«ü Then
            textString = "¢»"
        ElseIf this_ef.fing1 And amt.f¥|«ü Then
            textString = "¢¼"
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
            textObj.styleName = "­µ²Å_¼Æ¦r"
            textObj.Layer = "SimpErhu²Å¸¹"
            ipos = ipos + 1
        End If
        
    End If
    
    If this_ef.InOut <> "" Then
    '¤º¥~
        textString = this_ef.InOut
        height = size * 0.47
        
        alignmentPoint(0) = insertionPoint(0) + (amt.A_TEXT_WIDTH * size / 2)
        alignmentPoint(1) = insertionPoint(1) + (size * 2.13) + (ipos * yAdd * size)
        alignmentPoint(2) = insertionPoint(2)
        
        
        ' Create the text object in model space
        Set textObj = ThisDrawing.ModelSpace.AddText(textString, insertionPoint, height)
        textObj.Alignment = acAlignmentCenter
        textObj.TextAlignmentPoint = alignmentPoint
        textObj.styleName = "­µ²Å_¼Æ¦r"
        textObj.Layer = "¸Ë¹¢²Å¸¹"
        ipos = ipos + 1
    End If
    
        


    If this_ef.Push <> "" Then

        alignmentPoint(0) = insertionPoint(0) + (amt.A_TEXT_WIDTH * size / 2)
        alignmentPoint(1) = insertionPoint(1) + (size * 2.13) + (ipos * yAdd * size)
        alignmentPoint(2) = insertionPoint(2)

        If this_ef.Push = "©Ô" Then
            textString = "b©Ô"
        ElseIf this_ef.Push = "±À" Then
            textString = "b±À"
        End If
        
        Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(alignmentPoint, textString, size, size, size, 0)
       
        'blockRefObj.styleName = "SimpErhu"
        blockRefObj.Layer = "SimpErhu²Å¸¹"
'
        ipos = ipos + 1
    End If
    
    
'    If this_ef.InOut <> "" Then
'    '¤º ¥~
'       '¤£¬O ´N´¡¤J¤å¦r
'            textString = this_ef.InOut
'            height = size * 0.47
'            alignmentPoint(0) = insertionPoint(0)
'            alignmentPoint(1) = insertionPoint(1) + (size * 2.13) + (ipos * yAdd * size)
'            alignmentPoint(2) = insertionPoint(2)
'            ' Create the text object in model space
'            Set textObj = ThisDrawing.ModelSpace.AddText(textString, insertionPoint, height)
'            textObj.Alignment = acAlignmentCenter
'            textObj.TextAlignmentPoint = alignmentPoint
'            textObj.styleName = "­µ²Å_¼Æ¦r"
'            textObj.Layer = "¸Ë¹¢²Å¸¹"
''
'        ipos = ipos + 1
'    End If
'



End Function


Private Function atBarXYpos(ByVal the_pt As point, ByVal the_track As Integer, _
 ByVal the_line As Integer, ByVal the_measure As Integer, ByVal the_alltempo As Double, ByVal the_isDorp As Boolean) As point
'¨ú±o²{¦bªºª«¥ó¦ì¸m
'atBarXYpos()
'*the_pt ­ìÂI
'*the_the_track ²{¦b¬O²Ä´X­y
'*the_line ²{¦b¬O²Ä´X¦æ
'*the_measure ²{¦b¬O²Ä´X¤p¸`
'*the_allTempo ²{¦bªº©çªø¬O¦h¤Ö
'*the_isDorp ²{¦b¬O§_¬O²ÅÂI­µ²Å
'


    
    
    'Dim ROW_ALL_DEF As Integer
    Dim MeasureDEF As Integer
    'Dim tmp_modTempo As Integer
    'Dim row As Integer


    Dim tmp_modCol As Integer
    Dim Col As Double
    Dim col_b As Integer
    Dim tmp_barSpaceWidth As Double '¤@¸`ªº¶ZÂ÷
    Dim tmp_rowspacing As Double '¨C¦æ¨ì¨C¦æªº¶ZÂ÷
    Dim tmp_NoteDist As Double   '¨C©ç­µ²Åªº¬Û¹ï¦ì¸m
    

    tmp_barSpaceWidth = (G.pagewidth - G.LeftSpace - G.RightSpace) / G.barsperstaff   'x ¶b¥Î
    tmp_NoteDist = (tmp_barSpaceWidth - G.BarToNoteSpace * 2) / G.mete
    tmp_rowspacing = (G.TrackToTrack * (G.Many - 1) + G.LineToLine)             'y ¶b¥Î

    MeasureDEF = (G.mete * (Cg.BLEN / G.mete2))   '­pºâ¨C¤p¸`¦³¦h¤Ö³æ¦ì (4©ç¥þ­µ²Å¦³¦h¤Ö³æ¦ì)
    'tmp_modTempo = the_alltempo Mod MeasureDEF '¥þ³¡°£±¼Á`³æ¦ì®É «á¡A¬ÝÁÙ¦³¦h¤Öªº ³æ¦ì®É
    '

    tmp_modCol = tmp_barSpaceWidth Mod G.mete2
    Col = the_alltempo / (Cg.BLEN / G.mete2)  '¨ú±o²{¦b¤p¸`ªº²Ä´X©ç
    col_b = (Col Mod G.mete)  '¨ú±o¨C¤p¸`ªº²Ä´X©ç

    
    
    

    Static lastPoint As New point
    Dim pp As New point
    'col ¬O¨C¤@¦æªº²Ä´X©ç¡A¬O¥H¤@©ç¬°³æ¦ì¨Ó¼Æ
    'tmp_modCol ¬O¨C¤@©çªº²Ä´X­Ó¦rªº¦ì¸m
    pp.x = G.LeftSpace + the_measure * tmp_barSpaceWidth + G.BarToNoteSpace  'ºâ¥X¤p¸`ªº¦ì¸m + ¤p¸`½u¨ì­µ²ÅªºªÅ¥Õ
    'pp.x = pp.x + (CDbl(tmp_modCol) / CDbl(PARTITION_DEF / (4))) * (amt.LINE_LEN * G.FontSize) '°£©ó 4¤À¤§¤@©ç ªº°ÝÃD
    pp.x = pp.x + CDbl(Col) * tmp_NoteDist  '°£©ó 4¤À¤§¤@©ç ªº°ÝÃD

    pp.y = (G.TrackToTrack) * the_track + ((G.TrackToTrack) * (G.Many - 1) + G.LineToLine) * the_line
    pp.y = -pp.y


    '¨C¸`³æ¦ì©ç¡Aªº·L½Õ
    '¨Ò  1 5    2 6             1 5  2 6
    '    ----   ----  ->«e¶i¦¨  ---- ----
    '    123456789AB            123456789AB
    pp.x = pp.x + CDbl(Col) * G.Beat_MIN_X '·L½Õ
    '¨C¤@³æ¦ì©ç¡Aªº·L½Õ
    '¨Ò  1  5 3   2  6 4             1 53     2 64
    '    ---====  ---====  ->«e¶i¦¨  --==     --==
    '    123456789ABCDEF             123456789ABCDEF
    Static ismodcol As Integer
'    If tmp_modCol > 0 Then
'        ismodcol = ismodcol + 1
'        pp.x = pp.x + ismodcol * G.MIN_X  '·L½Õ
'    Else
'        ismodcol = 0
'    End If

    If the_isDorp Then '¦pªG¬O²ÅÂI­µ²Å¡A´N«eªñ¤@¥b
        pp.x = (lastPoint.x + pp.x) / 2
    End If
    
    '¬O§_2­Ó¦r¤ÓªñÀ£¨ì¤F
    '²¾¦Ü1­Ó¦rªº¼e«×
    If Abs(lastPoint.x - pp.x) <= amt.A_TEXT_WIDTH * G.fontsize Then
        pp.x = lastPoint.x + amt.A_TEXT_WIDTH * G.fontsize * 1.05
    End If
    
    Set lastPoint = pp  '¦s¤J³Ì«áªºÂI
    
    
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
    Dim tmp_bardist As Double '¤@¸`ªº¶ZÂ÷
    Dim tmp_rowspacing As Double '¨C¦æ¨ì¨C¦æªº¶ZÂ÷
    
    If the_measure = 0 And the_alltempo = 0 Then
    '¨C¦æªº²Ä¤@¤p¸`¡Aµe¥X¨C¦æªº¤p¸`½u
        tmp_bardist = (G.pagewidth - G.LeftSpace - G.RightSpace) / G.barsperstaff
        tmp_rowspacing = (G.TrackToTrack * (G.Many - 1) + G.LineToLine)
        '³o¬O­nµe¤p¸`½u
        'ROW_ALL_DEF ¬O¨C¦æªºÁ`³æ¦ì¼ÆÀ³¸Ó¦h¤Ö¡A¦p ¨C¤p¸`2/4©ç¡A¨C¦æ¦³ 7 ­Ó¤p¸`¡A«h¨C¦æªº Á`³æ¦ì®ÉÀ³¸Ó¬O (240*2)*7=3360
        'row ²Ä´X¦æ
        'colunm ²Ä´XÄæ(¦b²Ä´X¦æªº²Ä´XÄæ)
        Dim ROW_ALL_DEF As Integer
        Dim tmp_modTempo As Integer
        'Dim row As Integer
    
        Dim tmp_modCol As Integer
        Dim Col As Integer
        Dim col_b As Integer
    
    
        ROW_ALL_DEF = (G.barsperstaff * G.mete * PARTITION_DEF / (G.mete2 / 4))
        tmp_modTempo = the_alltempo Mod ROW_ALL_DEF '¥þ³¡°£±¼Á`³æ¦ì®É «á¡A¬ÝÁÙ¦³¦h¤Öªº ³æ¦ì®É
        'row = (the_alltempo - tmp_modTempo) \ ROW_ALL_DEF
    
        tmp_modCol = tmp_modTempo Mod (PARTITION_DEF / (G.mete2 / 4))
        Col = (tmp_modTempo - tmp_modCol) / (PARTITION_DEF / (G.mete2 / 4)) '¨ú±o¨C¦æªº²Ä´X©ç
        col_b = (Col Mod G.mete)  '¨ú±o¨C¤p¸`ªº²Ä´X©ç
        
        If the_track = 0 Then '¥u¦³²Ä¤@­y¤~­nµe
            
            '³o¬O¦bµe²Ä¤@¤p¸`
    
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
'*the_pt ­ìÂI
'*the_pt ²{¦b¬O²Ä´X­y
'*the_allTempo ²{¦bªº©çªø¬O¦h¤Ö
'*the_isDorp ²{¦b¬O§_¬O²ÅÂI­µ²Å
'


    atTableDraw_bar the_pt, the_track, the_alltempo
    
    Dim ROW_ALL_DEF As Integer
    Dim tmp_modTempo As Integer
    Dim row As Integer


    Dim tmp_modCol As Integer
    Dim Col As Integer
    Dim col_b As Integer


    ROW_ALL_DEF = (G.bar * G.mete * PARTITION_DEF / (G.mete2 / 4))
    tmp_modTempo = the_alltempo Mod ROW_ALL_DEF '¥þ³¡°£±¼Á`³æ¦ì®É «á¡A¬ÝÁÙ¦³¦h¤Öªº ³æ¦ì®É
    row = (the_alltempo - tmp_modTempo) \ ROW_ALL_DEF

    tmp_modCol = tmp_modTempo Mod (PARTITION_DEF / (G.mete2 / 4))
    Col = (tmp_modTempo - tmp_modCol) / (PARTITION_DEF / (G.mete2 / 4)) '¨ú±o¨C¦æªº²Ä´X©ç
    col_b = (Col Mod G.mete)  '¨ú±o¨C¤p¸`ªº²Ä´X©ç

    '¶Ç¥X¨C¸`ªº¬Û¹ï¦ì¸m
    Dim tmp_xbarInterval As Double
    tmp_xbarInterval = (G.pagewidth - G.LeftSpace - G.RightSpace) / ((G.bar * G.mete))

    Static lastPoint As New point
    Dim pp As New point
    'col ¬O¨C¤@¦æªº²Ä´X©ç¡A¬O¥H¤@©ç¬°³æ¦ì¨Ó¼Æ
    'tmp_modCol ¬O¨C¤@©çªº²Ä´X­Ó¦rªº¦ì¸m
    pp.x = G.LeftSpace + G.BarToNoteSpace + Col * tmp_xbarInterval
    pp.x = pp.x + (CDbl(tmp_modCol) / CDbl(PARTITION_DEF / (4))) * (amt.LINE_LEN * G.fontsize) '°£©ó 4¤À¤§¤@©ç ªº°ÝÃD

    pp.y = (G.TrackToTrack) * the_track + ((G.TrackToTrack) * (G.Many - 1) + G.LineToLine) * row
    pp.y = -pp.y
    pp.x = pp.x + the_pt.x

    '¨C¸`³æ¦ì©ç¡Aªº·L½Õ
    '¨Ò  1 5    2 6             1 5  2 6
    '    ----   ----  ->«e¶i¦¨  ---- ----
    '    123456789AB            123456789AB
    pp.x = pp.x + CDbl(col_b) * G.Beat_MIN_X '·L½Õ
    '¨C¤@³æ¦ì©ç¡Aªº·L½Õ
    '¨Ò  1  5 3   2  6 4             1 53     2 64
    '    ---====  ---====  ->«e¶i¦¨  --==     --==
    '    123456789ABCDEF             123456789ABCDEF
    Static ismodcol As Integer
    If tmp_modCol > 0 Then
        ismodcol = ismodcol + 1
        pp.x = pp.x + ismodcol * G.MIN_X  '·L½Õ
    Else
        ismodcol = 0
    End If
    pp.y = pp.y + the_pt.y
    If the_isDorp Then '¦pªG¬O²ÅÂI­µ²Å¡A´N«eªñ¤@¥b
        pp.x = (lastPoint.x + pp.x) / 2
    End If
    
    '¬O§_2­Ó¦r¤ÓªñÀ£¨ì¤F
    '²¾¦Ü1­Ó¦rªº¼e«×
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
    Dim tmp_bardist As Double '¤@¸`ªº¶ZÂ÷
    Dim tmp_rowspacing As Double '¨C¦æ¨ì¨C¦æªº¶ZÂ÷

    tmp_bardist = (G.pagewidth - G.LeftSpace - G.RightSpace) / G.bar
    tmp_rowspacing = (G.TrackToTrack * (G.Many - 1) + G.LineToLine)
    '³o¬O­nµe¤p¸`½u
    'ROW_ALL_DEF ¬O¨C¦æªºÁ`³æ¦ì¼ÆÀ³¸Ó¦h¤Ö¡A¦p ¨C¤p¸`2/4©ç¡A¨C¦æ¦³ 7 ­Ó¤p¸`¡A«h¨C¦æªº Á`³æ¦ì®ÉÀ³¸Ó¬O (240*2)*7=3360
    'row ²Ä´X¦æ
    'colunm ²Ä´XÄæ(¦b²Ä´X¦æªº²Ä´XÄæ)
    Dim ROW_ALL_DEF As Integer
    Dim tmp_modTempo As Integer
    Dim row As Integer

    Dim tmp_modCol As Integer
    Dim Col As Integer
    Dim col_b As Integer


    ROW_ALL_DEF = (G.bar * G.mete * PARTITION_DEF / (G.mete2 / 4))
    tmp_modTempo = the_alltempo Mod ROW_ALL_DEF '¥þ³¡°£±¼Á`³æ¦ì®É «á¡A¬ÝÁÙ¦³¦h¤Öªº ³æ¦ì®É
    row = (the_alltempo - tmp_modTempo) \ ROW_ALL_DEF

    tmp_modCol = tmp_modTempo Mod (PARTITION_DEF / (G.mete2 / 4))
    Col = (tmp_modTempo - tmp_modCol) / (PARTITION_DEF / (G.mete2 / 4)) '¨ú±o¨C¦æªº²Ä´X©ç
    col_b = (Col Mod G.mete)  '¨ú±o¨C¤p¸`ªº²Ä´X©ç
    
    If the_track = 0 Then '¥u¦³²Ä¤@­y¤~­nµe
        If Col = 0 And tmp_modCol = 0 Then '³o¬O¦b²Ä¤@¤p¸`

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
    '¦r«¬³]©w
    Me.cobFontName.AddItem "EUDC"
    Me.cobFontName.AddItem "²Ó©úÅé"
    Me.cobFontName.AddItem "µØ±d¦æ®ÑÅé"
    Me.cobFontName.AddItem "¼Ð·¢Åé"
    
    '¦r«¬¤j¤p³]©w
    For dd = 2 To 8 Step 0.5
        Me.cobFontSize.AddItem dd
    Next
    
    '´XÁn³¡³]©w
    For i = 1 To 8
        Me.cobMany.AddItem i
    Next
    
    '´X©ç³]©w
    Me.cobMete.AddItem "1/4"
    Me.cobMete.AddItem "2/4"
    Me.cobMete.AddItem "3/4"
    Me.cobMete.AddItem "4/4"
    Me.cobMete.AddItem "1/8"
    Me.cobMete.AddItem "3/8"
    Me.cobMete.AddItem "5/8"
    Me.cobMete.AddItem "7/8"
    
    
    
    '¨C¦æ­n´X¤p¸`³]©w
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
    
    AMT_LOAD '³o­Ó­«­n¡A³]©wªì©l¸ê®Æ
End Sub

Sub setLayoutMusicItem()


            G.currMeasure = 0
            For tmp_track_item = 0 To m_buf.GetTrackBufferSize(tmp_track) - 1

                '³o¬O­n³sµ²½uªº¡A¥H m_Mete2 ¬°®É­È
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

                    If G.currMeasure >= G.barsperstaff Then  '¬Ý¨C¦æ¤p¸`¼Æ¦³¨S¦³¶W¹L
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
    '³o°j°é¬O³]©w X ¶b¦V
    Do While (finished(staffGroup.voices) = False)   ' Inner loop.
       Dim currVoice As VoiceElement
       Set currVoice = staffGroup.voices(1)
       Debug.Print currVoice.i
       
        
        '' §ä¨ì­n¦b¸óÁn­µªº­Ô¿ïªÌ¤§¶¡§G¸mªº²Ä¤@­Ó«ùÄò®É¶¡¯Å§O
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
        ''¹jÂ÷¥Ø«e«ùÄò®É¶¡µ¥¯Åªº»y­µ
        Dim currentvoices As New iArray ' VoiceElement[] = []
        Dim othervoices As New iArray ' VoiceElement[] = []
        currentvoices.Clear
        othervoices.Clear
        For i = 0 To staffGroup.voices.Count - 1
            durationIndex = getDurationIndex(staffGroup.voices(i))
            '' PER: Because of the inexactness of JS floating point math, we just get close.
            '' PER¡G¥Ñ©ó JS ¯BÂI¼Æ¾Çªº¤£ºë½T©Ê¡A§Ú­Ì¥u¬O±µªñ¦Ó¤w¡C
            If (durationIndex - currentduration > Epsilon) Then
                othervoices.Push staffGroup.voices(i)
                ''console.log("out: voice ",i)
             Else
                currentvoices.Push staffGroup.voices(i)
                ''if (debug) console.log("in: voice ",i)
            
            End If
        Next
        
         
        

        
        '' among the current duration level find the one which needs starting furthest right
        '' ¦b¥Ø«e«ùÄò®É¶¡¯Å§O¤¤§ä¨ì»Ý­n±q³Ì¥kÃä¶}©lªº«ùÄò®É¶¡¯Å§O
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
            ''line ¤£ª¾¨ì if (~isSameStaff(v, topVoice)) then   Set topVoice = Empty
            voicechildx = layoutVoiceElement.layoutOneItem(x, spacing, v, 0, topVoice)
            dx = voicechildx - x
            ''³o¬O¬Ý¬O§_¦³«e­Ê­µ
            ''¦pªG¦³¡A¥þ³¡ªº­µ²Å´N¦b¥[«e­Ê­µªº¶ZÂ÷
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
        '' §R°£¨ä¥L»y­µ¤¤¤w­pºâªº¶¡¹j³æ¦ìªº­È
        ''¡]¨Ò¦p¡A¦pªG¤@­Ó»y­µ­p¹º¨Ï¥Î 5 ­Ó¶¡¹j³æ¦ì¡A¦ý¥¼¦b¦¹«ùÄò®É¶¡¯Å§O¤W±Æ¦C -
        '' §Ú­Ì¨Ï¥Î¤F 2 ­Ó¶¡¹j³æ¦ì - ¨º»ò§Ú­Ì¥²¶·¥Î§¹3­Ó¶¡¶Z³æ¦ì¡A¦Ó«D5 ­Ó¡^
        '' ´ú¸Õ«á¤£»Ý­n
        For i = 0 To othervoices.Count - 1
            othervoices(i).spacingduration = othervoices(i).spacingduration - spacingduration
            Call layoutVoiceElement.updateNextX(x, spacing, othervoices(i))   '' adjust other voices expectations
        Next
        
                    
              
        '' §ó·s¥Ø«e§G§½¤¸¯Àªº¯Á¤Þ
        For i = 0 To currentvoices.Count - 1
            Dim voice As VoiceElement
            Set voice = currentvoices(i)
            '' §â¨C¤@­Ó voice.i ¥[ 1 ¬°¤U¤@­Ó¤l¤¸¯À
            '' ÁÙ¦³­×§ï voice.durationindex ¥[¤W²{¦b¤w¸gÅª¨ú ªº­µ²Åªø«×
            '' 4¤À­µªþ=0.25 2¤À­µªþ=0.5 ¥þ¤À­µªþ=1
            '' ¨C¤@¤p¸`Á`ªø(¤À¥À)¬° 1
            Call layoutVoiceElement.updateIndices(voice)
        Next
    Loop
    i = i + 1
End Function
