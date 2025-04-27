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
'2013.03.17  V3 ���n�ק� �G�J�������A�]�{�����e�O�Υj�媺���k�ϡA�{�b�令�G�J�����k��
'2017.09.28  V8 �g�J�nmidi �ɪ�
Const VERSION  As String = "v1.65" '�n�鸹�X
Const C1 As Integer = 60   'C��1����W��
Const FOURPAINUM   As Integer = 64 '1/4���ŭp��
Const MIDICLOCK As Integer = 24   '�C1/64���Ū�MIDICLOCK��
Const TEMPO_DEF As Integer = 90   '�w�]�C����90��
Const PARTITION_DEF As Integer = 240   '�w�]�C����ά�120�p�ɳ椸
Const VOLUME_DEF As Integer = 64
Const MAINLAYER As String = "MAIN"    '�D�n���ϼh


Private Type Glode
    check1 As Boolean
    fontName As String
    FontSize As Double
    Many As Integer
    Bar As Integer
    Mete As Integer
    Mete2 As Integer
    
    PageWidth As Double
    LeftSpace As Double
    RightSpace As Double
    BarToNote As Double
    TrackToTrack As Double
    LineToLine As Double
    MIN_X As Double
    Beat_MIN_X As Double
End Type

'�G�J���k ���A
Private Type ErhuFing
    fing1 As Integer    '�� �� �� �� "b�Ų�"
    fing2 As String  '�� �� �� �� "b�Ų�"
    Push As String      '�ԡ� ��V
    InOut As String     '��   �~
End Type


Dim G As Glode
Dim m_Buf As New DataBuffer

'1 �bvb�u�{���ޥ�autocad����
'2 �w��autocad���H
Private acadApp As AcadApplication
Private acadDoc As AcadDocument
'3 �����{�ׄ�autocad����ۡA�H�U�O�ڇ���
'--------------------------------------------------------------
'����Cad
'-------------------------------------------------------------
Private Function AcadConnect() As Boolean
Dim flag As Boolean
On Error Resume Next
    Set acadApp = GetObject(, "AutoCAD.Application")
    flag = True
    If Err Then
       Err.clear
       Set acadApp = CreateObject("AutoCAD.Application")
       flag = True
       If Err Then
          flag = False
          MsgBox "�����b��AutoCAD,�[��d�O�_�w�E�I", vbOKCancel, "ĵ�i�I"
          Exit Function
       End If
    End If
    AcadConnect = flag
    Set acadDoc = acadApp.ActiveDocument
    'acadDoc.Close False
End Function


Private Sub cbCANCLE_Click()
    Me.Hide
End Sub


Private Sub cmOK_Click()
    G.fontName = Me.cobFontName.text
    G.FontSize = Me.cobFontSize.text
    
    Dim mete_mete As Variant
    mete_mete = Split(Me.cobMete, "/")
    G.Mete = mete_mete(0)
    G.Mete2 = mete_mete(1)
    
    G.Many = Me.cobMany
    G.Bar = Me.cobBar.text
    
    G.PageWidth = Me.tbPageWidth '�e��
    G.LeftSpace = Me.tbLeftSpace  '���ť�
    G.RightSpace = Me.tbRightSpace  '�k�ť�
    G.BarToNote = Me.tbBarToNote    '�p�`�쭵��
    G.TrackToTrack = Me.tbTrackToTrack  '�n�����Z
    G.LineToLine = Me.tbLineToLine      '�C�涡�Z
    G.check1 = True
    G.MIN_X = Me.tbMIN_X            '�L��
    G.Beat_MIN_X = Me.tbBeat_min_x  '��L��

    m_Buf.clear
    Call m_Buf.GetDataToBuf(Me.TextBox8.text)
    'MsgBox Me.TextBox8.text
    Me.Hide
    database
    Call put_many_text3
    
    
    
End Sub
Private Sub database()
    ' Create new layer
    Dim layerObj As AcadLayer
    Dim datalayer(20, 2) As String
    
    datalayer(1, 1) = "FIGE":
    datalayer(1, 2) = 6
    datalayer(2, 1) = "TEXT":
    datalayer(2, 2) = 2
    datalayer(3, 1) = "bar":
    datalayer(3, 2) = 181
    datalayer(4, 1) = "�˹��Ÿ�":
    datalayer(4, 2) = 4
    datalayer(5, 1) = "��ø�u":
    datalayer(5, 2) = 1
    datalayer(6, 1) = "main":
    datalayer(6, 2) = 7
    datalayer(7, 1) = "TEMP":
    datalayer(7, 2) = 1
    datalayer(8, 1) = "SimpErhu�Ÿ�":
    datalayer(8, 2) = 151
    
    Dim i As Integer
    
    Dim color(0 To 8) As AcadAcCmColor
    
    For i = 1 To 8
        Set layerObj = ThisDrawing.Layers.add(datalayer(i, 1))
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
Private Sub inst_G(the_G As Glode, pt As Variant)
    '���J�]�w���
    Dim mtxt As AcadMText
    Dim textObj As AcadText
    Dim textString As String
    Dim insertionPoint(0 To 2) As Double
    Dim height As Double
    
    ' Define the text object
    textString = VERSION & vbCrLf
    textString = textString & "size " & the_G.FontSize & vbCrLf
    
    textString = textString & "���ť� " & the_G.LeftSpace & "mm" & vbCrLf
    textString = textString & "�k�ť� " & the_G.RightSpace & "mm" & vbCrLf
    textString = textString & "�p�`�쭵�� " & the_G.BarToNote & "mm" & vbCrLf
    textString = textString & "�n��  " & the_G.TrackToTrack & "mm" & vbCrLf
    textString = textString & "�C��  " & the_G.LineToLine & "mm" & vbCrLf
    textString = textString & "�L��  " & the_G.MIN_X & "mm" & vbCrLf
    textString = textString & "��L�� " & the_G.Beat_MIN_X
    
    
    insertionPoint(0) = pt(0) - 30: insertionPoint(1) = pt(1): insertionPoint(2) = pt(2)
    height = 3
    
    ' Create the text object in model space
    Set mtxt = ThisDrawing.ModelSpace.AddMText(insertionPoint, height, textString)
    mtxt.width = 40
    mtxt.styleName = "Standard"
    
End Sub

Private Sub put_many_text3()
  
    '��m��r��
    Dim pt As Variant
    
    ' Return a point using a prompt
    pt = ThisDrawing.Utility.GetPoint(, "\n��ܭn���J���I �GEnter insertion point: ")
    Call inst_G(G, pt)
'***********************************************************************************
    '�e�X�w��u-�e��;
    
    Dim plineObj As AcadPolyline
    Dim Pnt As New PointList
    
    Call Pnt.add(pt(0), pt(1) - 200, 0)
    Call Pnt.add(pt(0), pt(1) + G.FontSize * 9, 0)
    Call Pnt.add(pt(0) + G.PageWidth, pt(1) + G.FontSize * 9, 0)
    Call Pnt.add(pt(0) + G.PageWidth, pt(1) - 200, 0)
    Set plineObj = ThisDrawing.ModelSpace.AddPolyline(Pnt.list())
    plineObj.Layer = "Defpoints"
'*********************************************************************************
    '���J��@���D
    Dim objText As AcadText
    Dim inPT As Variant
    Dim ooPt As Variant
    inPT = pt
    inPT(0) = inPT(0) + (G.PageWidth / 2)
    inPT(1) = inPT(1) + G.FontSize * 5.5
    Set objText = ThisDrawing.ModelSpace.AddText(m_Buf.getTITLE, inPT, 6)
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
    Set layerObj = ThisDrawing.Layers.add(MAINLAYER)

    
    
'lin    Dim tmp_joinApp As New MTJoinDequeApp()
'lin    Dim tmp_joinIds As New AcDbObjectIdArray()
    Dim tmp_joinApp As Variant
    Dim tmp_joinIds As New Collection
    'double lastTemp ;
    Dim tmp_track As Integer
    Dim tmp_track_item As Long
    Dim num As Integer
    Dim A_TEMPO_add As Integer

    Dim tmp_delaytime As Double  '�p���l �O�n�p�W�@�Ӧr������
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
  
        For tmp_track = 0 To m_Buf.GetTrackSize() - 1
            'NewObj = Nothing
'lin            tmp_joinApp.clear()
            tmp_AllTempo = 0
            tmp_delaytime = 0
            num = 0
            A_TEMPO_add = 1

            For tmp_track_item = 0 To m_Buf.GetTrackBufferSize(tmp_track)

                '�o�O�n�s���u���A�H m_Mete2 ���ɭ�
                If num >= (PARTITION_DEF / (G.Mete2 / 4) * A_TEMPO_add) Then
                    If A_TEMPO_add = G.Mete Then
                        num = 0
                        A_TEMPO_add = 1
                    Else
                        A_TEMPO_add = A_TEMPO_add + 1
                    End If

                    If tmp_joinIds.count >= 1 Then
                        Dim pp As Long
                        ReDim tmp_joinApp(tmp_joinIds.count - 1)
                        For pp = 1 To tmp_joinIds.count
                            Set tmp_joinApp(pp - 1) = tmp_joinIds.item(pp)
                        Next
                            MBG.addMusicJoin tmp_joinApp

                    End If
                    Set tmp_joinIds = Nothing

                End If


                cst = m_Buf.GetData(tmp_track, tmp_track_item)
                If " " = Mid(cst, amt.iNote, 1) Or "" = Mid(cst, amt.iNote, 1) Then
                    Exit For
                End If

                tmp_AllTempo = tmp_AllTempo + tmp_delaytime
                Dim ppnt As New point
                Dim bbo As Boolean
                ppnt.x = pt(0)
                ppnt.Y = pt(1)
                ppnt.z = pt(2)
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
'  ���JMusicText ����
'**************************************************************************************'
                If Me.chkOption1 = True Then
                    '(�j���)
                    '�϶��Ϊ���)
                    
                    ppnt.x = atPt(0)
                    ppnt.Y = atPt(1)
                    ppnt.z = atPt(2)
                    MBG.setDataText ppnt, cst, G.FontSize
                    Set BNewObj = MBG.InsterEnt '���J���ŤΫ��k

                    
                    
                ElseIf Me.chkOption2 = True Then

                '(�G�J��)
                '�o�O�S�����k��
'                    AMT.iTONE = 1        ' * ��
'                    AMT.iFinge = 2    '�o�O���k��     _+)(*&
'                    AMT.iScale = 3    '�o�O���C����   .,:;
'                    AMT.iNote = 4     '�o���D��       1234567.|l
'                    AMT.iTempo = 5    '�o�O��l       -=368acefz
'                    AMT.iTowFinge = 6    '�o�O���k��ĤG��  _+)(*&
'                    AMT.iSlur = 7        ' �s���Ŧ�    (3456)
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

                    '���o����r -���k1
                    Select Case Mid(cst, amt.iFinge, 1)
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
                    Select Case Mid(cst, amt.iTowFinge, 1)
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
                    ppnt.Y = atPt(1)
                    ppnt.z = atPt(2)
                    MBG.setDataText ppnt, cst, G.FontSize
                    Set BNewObj = MBG.InsterEnt '���J���ŤΫ��k
                    
                    
                    '���J���k ����(�G�J��)
                    InsertErhuFinge ppnt, tmp_erhu_fing, G.FontSize
                    
                End If
                    
'*******���J��ƽu'**************************************************************************************'

 
                'AMT.iSlur = 7        ' �s���Ŧ�    (3456)
                If Mid(cst, amt.iSlur, 1) = "[" Then
                    Set mt_slur_left = New MusicBlockGraphics
                    Set mt_slur_left = MBG.copy
                    
                ElseIf Mid(cst, amt.iSlur, 1) = "]" Then
'*******�E�X�u �e��'**************************************************************************************'
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
                    
                                        
                    'lenght = Abs(mt_slur_left.Grip.gptLeft.x - MBG.Grip.gptRight.x)

                    'points(0) = mt_slur_left.Grip.atPt.x
                    'points(1) = mt_slur_left.Grip.gptMidUp.Y
                    'points(2) = points(0) + islurAddX
                    'points(3) = points(1) + islurAddY
                    'points(4) = points(2) + lenght - (islurAddX * 2)
                    'points(5) = points(3)
                    'points(6) = points(4) + islurAddX
                    'points(7) = points(1)
                    
                    '�ݬO�_���u���Z���Ӫ�
                    
                    lenght = Abs(mt_slur_left.Grip.gptMid.x - MBG.Grip.gptMid.x)
                    If lenght >= islurAddX * 2 Then
                    
                        points(0) = mt_slur_left.Grip.gptMid.x
                        points(1) = mt_slur_left.Grip.gptMidUp.Y + islurBy
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
                       plineSlur.SetWidth 0, 0, 0.1
                       plineSlur.SetWidth 1, 0.1, 0.1
                       plineSlur.SetWidth 2, 0.1, 0
                       plineSlur.Layer = "fige"
                       
                       color.ColorIndex = 3
                       plineSlur.TrueColor = color
                       plineSlur.Update
                                       
                    Else
                    '�Z���Ӫ񪺵e�u
                        points_s(0) = mt_slur_left.Grip.gptMid.x
                        points_s(1) = mt_slur_left.Grip.gptMidUp.Y + islurBy
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
                       
                       plineSlur.SetWidth 0, 0, 0.1
                       plineSlur.SetWidth 1, 0.1, 0
                       plineSlur.Layer = "fige"
                       
                       color.ColorIndex = 3
                       plineSlur.TrueColor = color
                       plineSlur.Update
                    End If
                        

                    
                    
                    'Set mt_slur_left = Nothing
'*****************************************************************
                    
                End If
                
                '�s���u��
                tmp_joinIds.add BNewObj
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

                num = num + CInt(Fix(tmp_delaytime))
            Next tmp_track_item
        Next tmp_track

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
    insertionPoint(1) = midDownPt.Y
    insertionPoint(2) = midDownPt.z
    
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
    Dim col As Integer
    Dim col_b As Integer


    ROW_ALL_DEF = (G.Bar * G.Mete * PARTITION_DEF / (G.Mete2 / 4))
    tmp_modTempo = the_alltempo Mod ROW_ALL_DEF '���������`���� ��A���٦��h�֪� ����
    row = (the_alltempo - tmp_modTempo) \ ROW_ALL_DEF

    tmp_modCol = tmp_modTempo Mod (PARTITION_DEF / (G.Mete2 / 4))
    col = (tmp_modTempo - tmp_modCol) / (PARTITION_DEF / (G.Mete2 / 4)) '���o�C�檺�ĴX��
    col_b = (col Mod G.Mete)  '���o�C�p�`���ĴX��

    '�ǥX�C�`���۹��m
    Dim tmp_xbarInterval As Double
    tmp_xbarInterval = (G.PageWidth - G.LeftSpace - G.RightSpace) / ((G.Bar * G.Mete))

    Static lastPoint As New point
    Dim pp As New point
    'col �O�C�@�檺�ĴX��A�O�H�@�笰���Ӽ�
    'tmp_modCol �O�C�@�窺�ĴX�Ӧr����m
    pp.x = G.LeftSpace + G.BarToNote + col * tmp_xbarInterval
    pp.x = pp.x + (CDbl(tmp_modCol) / CDbl(PARTITION_DEF / (4))) * (amt.LINE_LEN * G.FontSize) '���� 4�����@�� �����D

    pp.Y = (G.TrackToTrack) * the_track + ((G.TrackToTrack) * (G.Many - 1) + G.LineToLine) * row
    pp.Y = -pp.Y
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
    pp.Y = pp.Y + the_pt.Y
    If the_isDorp Then '�p�G�O���I���šA�N�e��@�b
        pp.x = (lastPoint.x + pp.x) / 2
    End If
    
    '�O�_2�Ӧr�Ӫ�����F
    '����1�Ӧr���e��
    If Abs(lastPoint.x - pp.x) <= amt.A_TEXT_WIDTH * G.FontSize Then
        pp.x = lastPoint.x + amt.A_TEXT_WIDTH * G.FontSize * 1.05
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

    tmp_bardist = (G.PageWidth - G.LeftSpace - G.RightSpace) / G.Bar
    tmp_rowspacing = (G.TrackToTrack * (G.Many - 1) + G.LineToLine)
    '�o�O�n�e�p�`�u
    'ROW_ALL_DEF �O�C�檺�`�������Ӧh�֡A�p �C�p�`2/4��A�C�榳 7 �Ӥp�`�A�h�C�檺 �`�������ӬO (240*2)*7=3360
    'row �ĴX��
    'colunm �ĴX��(�b�ĴX�檺�ĴX��)
    Dim ROW_ALL_DEF As Integer
    Dim tmp_modTempo As Integer
    Dim row As Integer

    Dim tmp_modCol As Integer
    Dim col As Integer
    Dim col_b As Integer


    ROW_ALL_DEF = (G.Bar * G.Mete * PARTITION_DEF / (G.Mete2 / 4))
    tmp_modTempo = the_alltempo Mod ROW_ALL_DEF '���������`���� ��A���٦��h�֪� ����
    row = (the_alltempo - tmp_modTempo) \ ROW_ALL_DEF

    tmp_modCol = tmp_modTempo Mod (PARTITION_DEF / (G.Mete2 / 4))
    col = (tmp_modTempo - tmp_modCol) / (PARTITION_DEF / (G.Mete2 / 4)) '���o�C�檺�ĴX��
    col_b = (col Mod G.Mete)  '���o�C�p�`���ĴX��
    
    If the_track = 0 Then '�u���Ĥ@�y�~�n�e
        If col = 0 And tmp_modCol = 0 Then '�o�O�b�Ĥ@�p�`

            startPt.x = the_pt.x + G.LeftSpace
            startPt.Y = -((amt.LINE_PASE + amt.DROP_UP) * G.FontSize) + tmp_rowspacing * row
            startPt.Y = -startPt.Y + the_pt.Y

            endPt.x = the_pt.x + G.LeftSpace
            endPt.Y = G.TrackToTrack * (G.Many - 1) + tmp_rowspacing * row
            endPt.Y = -endPt.Y + the_pt.Y

            ptlist.clean
            ptlist.addpt startPt
            ptlist.addpt endPt
            Set tmp_pLWPoly = ThisDrawing.ModelSpace.AddPolyline(ptlist.list)
            tmp_pLWPoly.ConstantWidth = amt.BAR_WITCH * G.FontSize / 4.6
            tmp_pLWPoly.Layer = "bar"
        End If

        If (col Mod G.Mete) = 0 And (tmp_modCol = 0) Then
            Dim j As Integer
            For j = 0 To G.Many - 1


                startPt.x = the_pt.x + G.LeftSpace + tmp_bardist + col / G.Mete * tmp_bardist

                startPt.Y = -((amt.LINE_PASE + amt.DROP_UP) * G.FontSize) + (G.TrackToTrack * j) + (tmp_rowspacing * row)

                startPt.Y = -startPt.Y + the_pt.Y


                endPt.x = the_pt.x + G.LeftSpace + tmp_bardist + col / G.Mete * tmp_bardist

                endPt.Y = (G.TrackToTrack * j) + (tmp_rowspacing * row)

                endPt.Y = -endPt.Y + the_pt.Y

                ptlist.clean
                ptlist.addpt startPt
                ptlist.addpt endPt
                Set tmp_pLWPoly = ThisDrawing.ModelSpace.AddPolyline(ptlist.list)
                tmp_pLWPoly.ConstantWidth = amt.BAR_WITCH * G.FontSize / 4.6
                tmp_pLWPoly.Layer = "bar"
            Next j
            'm_pLWPoly->setThickness(plineInfo.m_thick);
            'm_pLWPoly->setConstantWidth(plineInfo.m_width);

        End If


    End If

End Function


Private Sub UserForm_Initialize()
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
    
    AMT_LOAD '�o�ӭ��n�A�]�w��l���
End Sub


