Attribute VB_Name = "voidItemConfig"
Option Explicit

Public StaffGroup As StaffGroupElement
Public G As Glode
Public Enum Cg
        BLEN = 1536   '預設全拍 4/4 為 1536 計時單元
        PARTITION_DEF = 386   '一拍 1/4 的計時單元

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
        REMARK = 17
        Config = 18

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
    Dim V As New VoiceElement
    Set StaffGroup = New StaffGroupElement
    Set ts = jsABCgetJson.getAbcJson
    
    For i = 0 To ts.Count - 1
        Set V = New VoiceElement
        Set V.children = ts(i)
        StaffGroup.voices.Push V
    Next
    
    Set getABC_StaffGroupElement = StaffGroup
End Function

   
 Function calDuration(s1 As MusicItem, lastDuration As Double) As Double

     ' lastDuration  上一拍子長度
     Dim tmp_delaytime As Double
                Select Case (s1.notes(0).mnote)
                    Case "|"
                        tmp_delaytime = 0
                    Case "-"
                        tmp_delaytime = G.mete2
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
            calDuration = tmp_delaytime

End Function
