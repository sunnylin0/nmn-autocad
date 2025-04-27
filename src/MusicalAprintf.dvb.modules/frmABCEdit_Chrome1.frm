VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmABCEdit_Chrome1 
   Caption         =   "UserForm1"
   ClientHeight    =   5844
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   9648
   OleObjectBlob   =   "frmABCEdit_Chrome1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmABCEdit_Chrome1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'2023/11/12  ¨ú±o abc2svg ¸ê®Æ¡A¨ÓÃ¸¥X Â²ÃÐ
Option Base 0

Const version  As String = "v1.0" '³nÅé¸¹½X
Const c1 As Integer = 60   'C½Õ1ªºÁä¦W­È
Const FOURPAINUM   As Integer = 64 '1/4­µ²Å­p¼Æ
Const MIDICLOCK As Integer = 24   '¨C1/64­µ²ÅªºMIDICLOCK¼Æ
Const TEMPO_DEF As Integer = 90   '¹w³]¨C¤ÀÄÁ90©ç
Const PARTITION_DEF As Integer = 384   '¹w³]¨C©ç¤À³Î¬°384­p®É³æ¤¸
Const VOLUME_DEF As Integer = 64      '­µ¶q¤j¤p
Const MAINLAYER As String = "MAIN"    '¥D­nªº¹Ï¼h


'¤G­J«üªk ª¬ºA
Private Type ErhuFing
    fing1 As Integer    '¢¹ ¢º ¢» ¢¼ "bªÅ²Í"
    fing2 As String  '¢¹ ¢º ¢» ¢¼ "bªÅ²Í"
    Push As String      '©Ô¡ä ±ÀV
    InOut As String     '¤º   ¥~
End Type


Dim G As Glode
Dim m_buf As New DataBuffer

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




Private Sub cmOK_Click()

    G.fontName = Me.cobFontName.text
    G.fontsize = Me.cobFontSize.text
    
    Dim mete_mete As Variant
    mete_mete = Split(Me.cobMete, "/")
    G.mete = mete_mete(0)
    G.mete2 = mete_mete(1)
    
    G.Many = Me.cobMany
    G.bar = Me.cobBar.text
    
    G.pagewidth = Me.tbPageWidth '¼e«×
    G.LeftSpace = Me.tbLeftSpace  '¥ªªÅ¥Õ
    G.RightSpace = Me.tbRightSpace  '¥kªÅ¥Õ
    G.BarToNoteSpace = Me.tbBarToNote    '¤p¸`¨ì­µ²Å
    G.TrackToTrack = Me.tbTrackToTrack  'Án³¡¶¡¶Z
    G.LineToLine = Me.tbLineToLine      '¨C¦æ¶¡¶Z
    G.check1 = True
    G.MIN_X = Me.tbMIN_X            '·L½Õ
    G.Beat_MIN_X = Me.tbBeat_min_x  '©ç·L½Õ

    m_buf.Clear
    Call m_buf.LoadDataToBuf(Me.TextBox8.text)
    'MsgBox Me.TextBox8.text
    Me.Hide
    database
    'Call put_many_text3
    

    Call getABC_StaffGroupElement
    Call put_many_text4
End Sub


Private Sub database()
    '«Ø¥ß AutoCAD Àô¹Ò¸ê·½
    ' Create new layer
    Dim layerObj As AcadLayer
    Dim datalayer(0 To 20, 0 To 2) As String
    
    datalayer(0, 1) = "FIGE":
    datalayer(0, 2) = 6
    datalayer(1, 1) = "TEXT":
    datalayer(1, 2) = 2
    datalayer(2, 1) = "bar":
    datalayer(2, 2) = 181
    datalayer(3, 1) = "¸Ë¹¢²Å¸¹":
    datalayer(3, 2) = 4
    datalayer(4, 1) = "¦ÛÃ¸½u":
    datalayer(4, 2) = 1
    datalayer(5, 1) = "main":
    datalayer(5, 2) = 7
    datalayer(6, 1) = "TEMP":
    datalayer(6, 2) = 1
    datalayer(7, 1) = "SimpErhu²Å¸¹":
    datalayer(7, 2) = 151
    
    Dim i As Integer
    
    Dim color(0 To 8) As AcadAcCmColor
    
    For i = 0 To 7
        Set layerObj = ThisDrawing.Layers.Add(datalayer(i, 1))
        Set color(i) = AcadApplication.GetInterfaceObject("AutoCAD.AcCmColor." & ACAD_Ver)
        
        color(i).ColorIndex = datalayer(i, 2)
        layerObj.TrueColor = color(i) 'datalayer(i, 2)
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
Private Sub inst_G(the_G As Glode, pt As Variant)
    '´¡¤J³]©w¸ê®Æ
    Dim mtxt As AcadMText
    Dim textObj As AcadText
    Dim textString As String
    Dim insertionPoint(0 To 2) As Double
    Dim height As Double
    
    ' Define the text object
    textString = version & vbCrLf
    textString = textString & "size " & the_G.fontsize & vbCrLf
    
    textString = textString & "¥ªªÅ¥Õ " & the_G.LeftSpace & "mm" & vbCrLf
    textString = textString & "¥kªÅ¥Õ " & the_G.RightSpace & "mm" & vbCrLf
    textString = textString & "¤p¸`¨ì­µ²Å " & the_G.BarToNote & "mm" & vbCrLf
    textString = textString & "Án³¡  " & the_G.TrackToTrack & "mm" & vbCrLf
    textString = textString & "¨C¦æ  " & the_G.LineToLine & "mm" & vbCrLf
    textString = textString & "·L½Õ  " & the_G.MIN_X & "mm" & vbCrLf
    textString = textString & "©ç·L½Õ " & the_G.Beat_MIN_X
    
    
    insertionPoint(0) = pt(0) - 30: insertionPoint(1) = pt(1): insertionPoint(2) = pt(2)
    height = 3
    
    ' Create the text object in model space
    Set mtxt = ThisDrawing.ModelSpace.AddMText(insertionPoint, height, textString)
    mtxt.width = 40
    mtxt.styleName = "Standard"
    
End Sub
Function finished(voices As iArray) As Boolean
'¬Ý voices ¬O§_¤w¸g¨ì³Ì«á¤@­Ó
    Dim i As Integer
    Dim voice As VoiceElement
    For i = 0 To voices.Count - 1
        Set voice = voices(i)
        If (voice.i >= voice.children.Count) Then
            finished = True
            Exit Function
        End If
    Next
    finished = False
End Function

Function updateIndices(voices As iArray) As Boolean
' voices ¦U +1 ¡A²¾¸m¤U¤@­Ó¤¸¯À ¬O§_¤w¸g¨ì³Ì«á¤@­Ó
    Dim i As Integer
    Dim voice As VoiceElement
    For i = 0 To voices.Count - 1
        Set voice = voices(i)
        voice.durationIndex = voice.durationIndex + voice.children(voice.i).dur
        voice.i = voice.i + 1   '²¾¸m¤U¤@­Ó¤¸¯À
        If (voice.i >= voice.children.Count) Then
            updateIndices = True
            Exit Function
        End If
    Next
    updateIndices = False
End Function

Function getDurationIndex(element As VoiceElement) As Double
    '' if the ith element doesn't have a duration (is not a note), its duration index is fractionally before.
    '' This enables CLEF KEYSIG TIMESIG PART, etc.to be laid out before we get to the first note of other voices
    '' ¦pªG²Ä i ­Ó¤¸¯À¨S¦³«ùÄò®É¶¡¡]¤£¬O­µ²Å¡^¡A«h¨ä«ùÄò®É¶¡¯Á¤Þ¦b«e­±¡C
    '' ³o¨Ï±o CLEF KEYSIG TIMESIG PART µ¥¯à°÷¦b§Ú­Ì¨ì¹F¨ä¥LÁn­µªº²Ä¤@­Ó­µ²Å¤§«e¶i¦æ§G§½
    Dim getItemDuration As Double
    If TypeOf element.children(element.i) Is voiceItem Then
        If element.children(element.i).dur > 0 Then
            getItemDuration = element.durationIndex - 0
        Else
            getItemDuration = element.durationIndex - 0.0000005
        End If
    End If
    
    getDurationIndex = getItemDuration
End Function


Private Sub put_many_text4()

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

    
'*********************************************************************************
'­pºâ­µ²Åªº x ¶b
    Dim ret As Double
    Dim newspace As Double
    newspace = 10
    ret = layoutStaffGroup(newspace, Nothing, False, staffGroup, 0)
'*********************************************************************************
'Ã¸»s­µ²Å¶}©l

        '' -- draw_symbols --
        Dim currentVoice As VoiceElement
        Dim child As voiceItem
        Dim i, j
        For i = 0 To staffGroup.voices.Count - 1
            Set currentVoice = staffGroup.voices(i)
            For j = 0 To currentVoice.children.Count - 1
                
                Set child = currentVoice.children(j)
                Select Case child.typs    ' Evaluate Number.
                    Case Cg.bar:
                    Case Cg.meter:
                       'draw_meter child
                    Case Cg.note, Cg.Rest:
                        Dim dx As Double
                        Dim dy As Double
                        dx = insPt(0)
                        dy = insPt(1)
                        draw_note child, dx, dy
                    Case Cg.grace:
                        'for (var g = s.extra  g  g = g.next)
                        '     draw_note(g)
                        'Next
                    Case Else    ' Other values.
                       Debug.Print "Not between 1 and 10"
                End Select

            Next
        Next




End Sub



Function layoutStaffGroup(spacing As Double, renderer As RendererModule, debug_ As Boolean, staffGroup As StaffGroupElement, leftEdge As Double) As Double

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
Private Sub draw_note(child As voiceItem, dx As Double, dy As Double)

                Dim ppnt As New point
                Dim bbo As Boolean
                Dim MBG As New MusicBlockGraphics
                Dim BNewObj As Object
                Dim strcode As String
                Dim musItem As New MusicItem
                Dim noteItem As New MusicNoteItem
                ppnt.x = child.x + dx
                ppnt.y = dy + (child.v * 7)
                ppnt.Z = 0
                Dim jn As String
                If IsEmpty(child.notes) = False Then
                    jn = child.notes(0).jn
                    strcode = "   " + jn + "   "
                    
                    noteItem.mnote = jn
                    musItem.notes.Push noteItem
                    
                    MBG.setDataText ppnt, musItem, G.fontsize
                    Set BNewObj = MBG.InsterEnt '´¡¤J­µ²Å¤Î«üªk
                End If

End Sub

Private Sub put_many_text3()
  
    '©ñ¸m¤å¦r¥»
    Dim pt As Variant
    
    ' Return a point using a prompt
    pt = ThisDrawing.Utility.GetPoint(, "\n¿ï¾Ü­n´¡¤JªºÂI ¡GEnter insertion point: ")
    Call inst_G(G, pt)
'***********************************************************************************
    'µe¥X©w¦ì½u-µe®Ø
    
    Dim plineObj As AcadPolyline
    Dim Pnt As New PointList
    
    Call Pnt.Add(pt(0), pt(1) - 200, 0)
    Call Pnt.Add(pt(0), pt(1) + G.fontsize * 9, 0)
    Call Pnt.Add(pt(0) + G.pagewidth, pt(1) + G.fontsize * 9, 0)
    Call Pnt.Add(pt(0) + G.pagewidth, pt(1) - 200, 0)
    Set plineObj = ThisDrawing.ModelSpace.AddPolyline(Pnt.list())
    plineObj.Layer = "Defpoints"
'*********************************************************************************
    '´¡¤J³æ¤@¼ÐÃD
    Dim objText As AcadText
    Dim inPT As Variant
    Dim ooPt As Variant
    inPT = pt
    inPT(0) = inPT(0) + (G.pagewidth / 2)
    inPT(1) = inPT(1) + G.fontsize * 5.5
    Set objText = ThisDrawing.ModelSpace.AddText(m_buf.getTITLE, inPT, 6)
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

    
'*********************************************************************************
'Ã¸»s­µ²Å¶}©l

'lin    Dim tmp_joinApp As New MTJoinDequeApp()
'lin    Dim tmp_joinIds As New AcDbObjectIdArray()
    Dim tmp_joinApp As Variant
    Dim tmp_joinIds As New Collection
    'double lastTemp
    Dim tmp_track As Integer
    Dim tmp_track_item As Long
    Dim num As Integer
    Dim A_TEMPO_add As Integer

    Dim tmp_delaytime As Double  '­pºâ©ç¤l ¬O­n­p¤W¤@­Ó¦rªºªø«×
    Dim tmp_AllTempo As Long
    Dim tmp_xy As point
    'Dim NewObj As NnmText
    Dim BNewObj As AcadBlockReference
    Dim tmp_name As String
    
    Dim cst As String
    Dim cst_no_fing As String
    Dim ptGripMid As Variant
    
    tmp_delaytime = 0
    tmp_AllTempo = 0
    
    tmp_name = G.fontName
    Dim tmp_erhu_fing As ErhuFing
    Dim midDownPt(2) As Double
    Dim mt_slur_left As MusicBlockGraphics
    Dim plineSlur As AcadLWPolyline
    Dim MBG As New MusicBlockGraphics
  
        For tmp_track = 0 To m_buf.GetTrackSize() - 1
            'NewObj = Nothing
'lin            tmp_joinApp.clear()
            tmp_AllTempo = 0
            tmp_delaytime = 0
            num = 0
            A_TEMPO_add = 1

            For tmp_track_item = 0 To m_buf.GetTrackBufferSize(tmp_track)

                '³o¬O­n³sµ²½uªº¡A¥H m_Mete2 ¬°®É­È
                If num >= (PARTITION_DEF / (G.mete2 / 4) * A_TEMPO_add) Then
                    If A_TEMPO_add = G.mete Then
                        num = 0
                        A_TEMPO_add = 1
                    Else
                        A_TEMPO_add = A_TEMPO_add + 1
                    End If

                    If tmp_joinIds.Count >= 1 Then
                        Dim pp As Long
                        ReDim tmp_joinApp(tmp_joinIds.Count - 1)
                        For pp = 1 To tmp_joinIds.Count
                            Set tmp_joinApp(pp - 1) = tmp_joinIds.item(pp)
                        Next
                            MBG.addMusicJoin tmp_joinApp

                    End If
                    Set tmp_joinIds = Nothing

                End If


                cst = m_buf.GetData(tmp_track, tmp_track_item)
                If " " = Mid(cst, amt.iNote, 1) Or "" = Mid(cst, amt.iNote, 1) Then
                    Exit For
                End If

                tmp_AllTempo = tmp_AllTempo + tmp_delaytime
                Dim ppnt As New point
                Dim bbo As Boolean
                ppnt.x = pt(0)
                ppnt.y = pt(1)
                ppnt.Z = pt(2)
                If Mid(cst, amt.iNote, 1) = "." Then
                    Set tmp_xy = atTableDraw(ppnt, tmp_track, tmp_AllTempo, True)
                Else
                    Set tmp_xy = atTableDraw(ppnt, tmp_track, tmp_AllTempo, False)
                End If
                                
        
                
                Dim atPt As Variant
                atPt = tmp_xy.at
                atPt(0) = 0
                atPt(1) = 0
                atPt(2) = 0
                
                 atPt = tmp_xy.at
'**************************************************************************************'
'  ´¡¤JMusicText ª«¥ó
'**************************************************************************************'
                If Me.chkOption1 = True Then
                    '(¥jºå¥Î)
                    '¹Ï¶ô¥Îªþ¸¹)
                    
                    ppnt.x = atPt(0)
                    ppnt.y = atPt(1)
                    ppnt.Z = atPt(2)
                    MBG.setDataText ppnt, cst, G.fontsize
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
                    cst_no_fing = Mid(cst, amt.iTONE, 1) & _
                                    " " & _
                                    Mid(cst, amt.iScale, 1) & _
                                    Mid(cst, amt.iNote, 1) & _
                                    Mid(cst, amt.iTempo, 1)

                    'Call NewObj.setData(atPt, cst_no_fing, G.fontName, G.FontSize)
                    'NewObj.Layer = "main"
                    tmp_erhu_fing.fing1 = 0
                    
                    tmp_erhu_fing.Push = ""
                    tmp_erhu_fing.InOut = ""

                    '¨ú±oÃöÁä¦r -«üªk1
                    Select Case Mid(cst, amt.iFinge, 1)
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
                    Select Case Mid(cst, amt.iTowFinge, 1)
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
                    MBG.setDataText ppnt, cst, G.fontsize
                    Set BNewObj = MBG.InsterEnt '´¡¤J­µ²Å¤Î«üªk
                    
                    
                    '´¡¤J«üªk ªþ¸¹(¤G­J¥Î)
                    InsertErhuFinge ppnt, tmp_erhu_fing, G.fontsize
                    
                End If
                    
'*******´¡¤J¶ê·Æ½u'**************************************************************************************'

 
                'AMT.iSlur = 7        ' ³s­µ²Å¦æ    (3456)
                If Mid(cst, amt.iSlur, 1) = "[" Or Mid(cst, amt.iSlur, 1) = "(" Then
                    Set mt_slur_left = New MusicBlockGraphics
                    Set mt_slur_left = MBG.copy
                    
                ElseIf Mid(cst, amt.iSlur, 1) = "]" Or Mid(cst, amt.iSlur, 1) = ")" Then
'*******»E¦X½u µe©·'**************************************************************************************'
                    Dim points(0 To 7) As Double
                    
                    Dim points_s(0 To 5) As Double
                    Dim lenght As Double
                    ' Find the bulge of the third segment
                    Dim currentBulge As Double
                    Dim color As New AcadAcCmColor
                    
                    
                    Dim islurAddX As Double
                    Dim islurAddY As Double
                    Dim islurBx As Double
                    Dim islurBy As Double
                    
                    islurAddX = 0.6245 * MBG.TextSize
                    islurAddY = 0.4333 * MBG.TextSize
                    
                    islurBy = 0.45 * MBG.TextSize
                    
                    
                    '¬Ý¬O§_©·½uªº¶ZÂ÷¤Óªñ
                    
                    lenght = Abs(mt_slur_left.Grip.gptMid.x - MBG.Grip.gptMid.x)
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
                
                       
                       currentBulge = plineObj.GetBulge(3)
                       ' Change the bulge of the third segment
                       plineSlur.SetBulge 0, (-0.5858 / 1.4142)
                       plineSlur.SetBulge 2, (-0.5858 / 1.4142)
                       plineSlur.setWidth 0, 0, 0.1
                       plineSlur.setWidth 1, 0.1, 0.1
                       plineSlur.setWidth 2, 0.1, 0
                       plineSlur.Layer = "fige"
                       
                       color.ColorIndex = 3
                       plineSlur.TrueColor = color
                       plineSlur.Update
                                       
                    Else
                    '¶ZÂ÷¤Óªñªºµe½u
                        points_s(0) = mt_slur_left.Grip.gptMid.x
                        points_s(1) = mt_slur_left.Grip.gptMidUp.y + islurBy
                        points_s(2) = points_s(0) + (lenght / 2)
                        points_s(3) = points_s(1) + islurAddY
                        points_s(4) = points_s(2) + (lenght / 2)
                        points_s(5) = points_s(1)
                        ' Create a lightweight Polyline object in model space
                       Set plineSlur = ThisDrawing.ModelSpace.AddLightWeightPolyline(points_s)
                
                       
                       currentBulge = plineObj.GetBulge(3)
                       ' Change the bulge of the third segment
                       plineSlur.SetBulge 0, (-0.5858 / 1.4142)
                       plineSlur.SetBulge 1, (-0.5858 / 1.4142)
                       
                       plineSlur.setWidth 0, 0, 0.1
                       plineSlur.setWidth 1, 0.1, 0
                       plineSlur.Layer = "fige"
                       
                       color.ColorIndex = 3
                       plineSlur.TrueColor = color
                       plineSlur.Update
                    End If
                        

                    
                    
                    'Set mt_slur_left = Nothing
'*****************************************************************
                    
                End If
                
                '³sµ²½u¥Î
                tmp_joinIds.Add BNewObj
                Set BNewObj = Nothing

                Select Case Mid(cst, amt.iNote, 1)
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
                        If cn = Mid(cst, amt.iTempo, 1) Then
                            tmp_delaytime = PARTITION_DEF / tempo_ll(ii)
                            Exit For
                        Else
                            tmp_delaytime = 0
                        End If
                    Next ii
                End Select
'¤G­¿
'                Select Case Mid(cst, amt.iNote, 1)
'                Case "."
'
'                Case Else
'                        tmp_delaytime = tmp_delaytime * 2
'                End Select
'
                num = num + CInt(Fix(tmp_delaytime))
            Next tmp_track_item
        Next tmp_track

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




Private Function atTableDraw(ByVal the_pt As point, ByVal the_track As Integer, ByVal the_alltempo As Long, ByVal the_isDorp As Boolean) As point

'atTableDraw()
'*the_pt ­ìÂI
'*the_pt ²{¦b¬O²Ä´X­y
'*the_allTempo ²{¦bªº©çªø¬O¦h¤Ö
'*the_isDorp ²{¦b¬O§_¬O²ÅÂI­µ²Å
'

    'Ã¸¤p¸`½u
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

            ptlist.clean
            ptlist.addpt startPt
            ptlist.addpt endPt
            Set tmp_pLWPoly = ThisDrawing.ModelSpace.AddPolyline(ptlist.list)
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

                ptlist.clean
                ptlist.addpt startPt
                ptlist.addpt endPt
                Set tmp_pLWPoly = ThisDrawing.ModelSpace.AddPolyline(ptlist.list)
                tmp_pLWPoly.ConstantWidth = amt.BAR_WITCH * G.fontsize / 4.6
                tmp_pLWPoly.Layer = "bar"
            Next j
            'm_pLWPoly->setThickness(plineInfo.m_thick)
            'm_pLWPoly->setConstantWidth(plineInfo.m_width)

        End If


    End If

End Function


Private Sub UserForm_Initialize()
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
    Me.cobFontSize.text = 3.5
    Me.cobMany.text = 1
    Me.cobMete.text = "4/4"
    Me.cobBar.text = 4

    Me.tbPageWidth = 210
    Me.tbLeftSpace = 14
    Me.tbRightSpace = 14
    Me.tbBarToNote = 2
    Me.tbTrackToTrack = 14
    Me.tbLineToLine = 18
    Me.tbMIN_X = 0.25
    
    AMT_LOAD '³o­Ó­«­n¡A³]©wªì©l¸ê®Æ
End Sub




