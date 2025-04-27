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


 Public mtone As String               '升降記號
 Public mfingering As String          '指法
 Public mnote As String               '高低音
 Public mscale As String              '音階
 Public mtempo As String              '拍子
 Public delay As Integer              '拍子的延遲長度
 Public mtow_fingering As String      '指法 第二個
 Public mslur As String               '連音符
 Public mand As Boolean               '記錄合音 true為跟後面一音合音 ,false為沒有
 Public what As Integer               '記錄一些不一樣的指法
