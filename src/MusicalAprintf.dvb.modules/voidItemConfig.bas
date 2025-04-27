Attribute VB_Name = "voidItemConfig"
Option Explicit

Public staffGroup As StaffGroupElement
Public G As Glode
Public Enum Cg
        BLEN = 1536   '預設全拍 4/4 為 1536 計時單元
        PARTITION_DEF = 384   '一拍 1/4 的計時單元

        '' symbol types
        bar = 0
        clef = 1
        CUSTOS = 2
        SM = 3    '' sequence marker (transient)
        grace = 4 ''這是裝飾音類別
        key = 5
        meter = 6
        MREST = 7
        note = 8
        part = 9
        Rest = 10
        space = 11
        staves = 12
        STBRK = 13
        tempo = 14
        Block = 16
        remark = 17
        Config = 18
        Other = 19

        '' note heads
        FULL = 0
        EMPTYe = 1
        OVAL = 2
        OVALBARS = 3
        SQUARE = 4

        '' position types
        SL_ABOVE = &H1       '' position (3 bits)
        SL_BELOW = &H2
        SL_AUTO = &H3
        SL_HIDDEN = &H4
        SL_DOTTED = &H8   '' modifiers
        SL_ALI_MSK = &H70 '' align
        SL_ALIGN = &H10
        SL_CENTER = &H20
        SL_CLOSE = &H40
        
        
    
End Enum

Enum AbcObject
    noteItem
    DecorationItme
    FontItem
    GchordItem
    LyrcsItem
    VoiecItem
    KMapItem
    Number
    Booling
End Enum
Type a_Meter
    top As String * 20
    bot As String * 20
End Type

'Type StaffGroupElement
'    line As Integer
'    startx As Integer
'    w As Integer
'    height As Integer
'    getTextSize As Integer
'    voices As New iArray 'VoiceElement[]
'    staffs As Object
'    brace As Variant
'    bracket As iArray 'BraceElem[]
'End Type
Function getABC_StaffGroupElement() As StaffGroupElement
    
    Dim ts As iArray
    Dim i As Integer
    Dim v As New VoiceElement
    Set staffGroup = New StaffGroupElement
    Set ts = jsABCgetJson.getAbcJson
    
    For i = 0 To ts.Count - 1
        Set v = New VoiceElement
        Set v.children = ts(i)
        staffGroup.voices.Push v
    Next
    
    Set getABC_StaffGroupElement = staffGroup
End Function

 Function calExtraw(s1 As MusicItem) As Double
    '前綴元素的寬度 是負值
    Dim i As Integer
    Dim mus As MusicItem
    Dim c As String
    If s1 Is Nothing Then
        Exit Function
    End If
    
    '找是否有升降記號
    For i = 0 To s1.notes.Count - 1
        c = s1.notes(i).mtone
        If c = "b" Or c = "#" Or c = "o" Then
            s1.extraw = s1.extraw - 0.3 * G.FONTSIZE
            i = s1.notes.Count '離開迴圈
        End If
    Next
    '計算 前墜字
    If Not s1.extraObjs Is Nothing Then
        s1.extraw = s1.extraw - (s1.extraObjs.Count * G.FONTSIZE * amt.extraScale * amt.wNote)
    End If
    
    '計算 後墜字
    
    If Not s1.rightObjs Is Nothing Then
        s1.w = s1.w + (s1.rightObjs.Count * G.FONTSIZE * amt.extraScale * amt.wNote)
    End If


    calExtraw = s1.extraw
 End Function
   
 Function calDuration(s1 As MusicItem, lastDuration As Double) As Double
 
    ' lastDuration  上一拍子長度
    Dim tmp_delaytime As Double
    Dim isDot As Boolean
    isDot = False
    Select Case (s1.notes(0).mnote)
        Case "|"
            tmp_delaytime = 0
        Case "-"
            tmp_delaytime = Cg.BLEN / G.mete2
        Case "."
            tmp_delaytime = lastDuration / 2
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
                    tmp_delaytime = Cg.PARTITION_DEF / tempo_ll(ii)
                    Exit For
                ElseIf s1.notes(0).mtempo = "" Then
                    tmp_delaytime = Cg.PARTITION_DEF
                    Exit For
                Else
                    tmp_delaytime = 0
                End If
            Next ii
    End Select
    
    For ii = 0 To s1.dots - 1
        tmp_delaytime = tmp_delaytime * 1.5
    Next

    calDuration = tmp_delaytime
End Function

 Function calDuration2(s1 As MusicItem, lastDuration As Double) As MusicItem
 
    ' lastDuration  上一拍子長度
    Dim tmp_delaytime As Double
    Dim isDot As Boolean
    Dim outcrop As Integer
    Dim tempo_hj As String
    Dim tempo_ll As Variant
    Dim tempo_flags As Variant
    
    tempo_hj = " -2=4536789aAbBcCdDeEfFgGzZ"
    'ttaaaaaa= array( , -  2  =  4  5  3  6  7  8   9  a   A   b   B   c   C   d   D  e   e    f   F   g   G   z   Z
    tempo_ll = Array(1, 2, 2, 4, 4, 5, 3, 6, 7, 8, 9, 10, 10, 11, 11, 12, 12, 13, 13, 14, 14, 15, 15, 16, 16, 32, 32)
    tempo_flags = Array(0, 1, 1, 2, 2, 2, 1, 2, 2, 3, 3, 3, 3, 3, 3, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 5)
    
    isDot = False
    Select Case (s1.notes(0).mnote)
        Case "|"
            tmp_delaytime = 0
        Case "-"
            tmp_delaytime = Cg.BLEN / G.mete2
        Case "."
            tmp_delaytime = lastDuration / 2
        Case " "
    
        Case Else
    

    
            Dim ii As Integer
            Dim cn As String
            For ii = 0 To Len(tempo_hj) - 1
                cn = Mid(tempo_hj, ii + 1, 1)
                If cn = s1.notes(0).mtempo Then
                    outcrop = ii
                    tmp_delaytime = Cg.PARTITION_DEF / tempo_ll(outcrop)
                    Exit For
                ElseIf s1.notes(0).mtempo = "" Then
                    tmp_delaytime = Cg.PARTITION_DEF
                    Exit For
                Else
                    tmp_delaytime = 0
                End If
            Next ii
    End Select
    '找3連音
    Select Case (s1.notes(0).mtempo)
        Case "3", "5", "6", "7", "9", "a", "A", "c", "C"
            s1.tripCount = tempo_ll(outcrop)
        Case Else
            s1.tripCount = 0
    End Select
    
    For ii = 0 To s1.dots - 1
        tmp_delaytime = tmp_delaytime * 1.5
    Next

    s1.duration = tmp_delaytime
    s1.nflags = Int(tempo_flags(outcrop))

    Set calDuration2 = s1
End Function


Public Sub debugVoices(staffGroup As StaffGroupElement)
'debug 資料行
        Dim i As Integer
        Dim t As Integer
        Dim sst As String
        Dim c1 As String
        Dim mus As MusicItem
        
        For i = 0 To staffGroup.voices(0).children.Count - 1
            Set mus = staffGroup.voices(0).children(i)
            
            If mus.typs = Cg.bar Then
                Debug.Print Format(CStr(mus.barNumber), "@@@@") & " _ "
                Exit For
            End If
        Next
        For t = 0 To staffGroup.voices.Count - 1
            sst = ""
            For i = 0 To staffGroup.voices(t).children.Count - 1
                Set mus = staffGroup.voices(t).children(i)
                If mus.typs = Cg.note Then
                    c1 = mus.notes(0).mnote
                ElseIf mus.typs = Cg.meter Then
                    c1 = "[" & mus.mete & "/" & mus.mete2 & "]"
                ElseIf mus.typs = Cg.bar Then
                    c1 = "|"
                Else
                    c1 = "*"
                End If
                sst = sst + c1
            Next
            Debug.Print Format("ch" & CStr(t) & "_" & sst, "@@@@")
        Next
End Sub


