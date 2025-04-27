VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MusicItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

 Public typs As Cg
 Public notes As iArray      'noteFormat
    'mfingering As String  '指法
    'mnote As String           '高低音
    'mscale As String          '音階
 Public mtempo As String             '拍子
 Public delay As Integer                '拍子的延遲長度
    'mtow_fingering As String  '指法 第二個
 Public slurStart As Boolean                '連音符
 Public slurEnd As Boolean                '連音符
 Public tripStart As Boolean           '3 連音開始
 Public tripEnd As Boolean             '3 連音結束
 Public mand As Boolean                  '記錄合音 true為跟後面一音合音 ,false為沒有
 Public what As Integer                 '記錄一些不一樣的指法
 
 
 Public mete As Integer     '拍號設定改變
 Public mete2 As Integer     '拍號設定改變
 Public barsperstaff  As Integer    '設定每行幾小節
  
 Public ln As Integer       '要記錄讀進來的資料是在文件裡的第幾 "行"
 Public WD As Integer      '要記錄貢進來的資料是在文件要的第幾 "字"
 Public x As Double
 Public y As Double
 Public w As Double
 Public extce As Double '這是前置的位置


