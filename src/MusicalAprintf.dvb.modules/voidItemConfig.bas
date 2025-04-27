Attribute VB_Name = "voidItemConfig"
Option Explicit

Public staffGroup As StaffGroupElement
Public Enum Cg
        BLEN = 1536

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

