VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MusicNoteItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

 Public typs As Cg
 Public configText As String
 Public mtone As String               '升降記號
 Public mfingering As String          '指法
 Public mnote As String               '高低音
 Public mscale As String              '音階
 Public mtempo As String              '拍子
 Public duration As Integer           '拍子的延遲長度
 Public mtow_fingering As String      '指法 第二個
 Public mslur As String               '連音符
 Public mand As String
 Public isAnd As Boolean              '記錄合音 true為跟後面一音合音 ,false為沒有
 Public what As Integer               '記錄一些不一樣的指法


Public Function ToOneMusicItem() As MusicItem
    
    Dim MuseIt As New MusicItem
    Dim lastDuration As Double
    Set MuseIt = New MusicItem
    Set MuseIt.notes = New MusicNoteItemList
    MuseIt.notes.Push Me
    
    lastDuration = calDuration(MuseIt, lastDuration)
    MuseIt.identify_Duration lastDuration
    calExtraw MuseIt
    
    Set ToOneMusicItem = MuseIt
End Function

