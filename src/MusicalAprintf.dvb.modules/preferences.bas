Attribute VB_Name = "preferences"
Option Explicit


Type AMusicType
    DEFINE_FontFileName  As String          '預設字型的名字
    DEFINE_TEXT_SIZE As Double          '預設字型的大小

    
    A_TEXT_WIDTH   As Double   '預設一個字的寬度  字形的 百分比

    LINE_PASE As Double '''這是從原點 要高 字形的 百分比
    LINE_JUMP   As Double                   '每一個拍線和拍線的間格 字形的 百分比
    LINE_THICKNESS   As Double                   '這是拍線的厚度 字形的 百分比
    LINE_LEN   As Double                   '這是拍線的長度 字形的 百分比

    DROP_UP   As Double                   '高八度的位置  字形的 百分比
    DROP_DOWN   As Double                   '低八度的位置  字形的 百分比
    DROP_CIRCLE_RADIUS_SIZE   As Double                   '圓的 直徑 SIZE 字形的 百分比
    DROP_INTERVAL  As Double                   '每八度的間距 字形的 百分比
    DROP_ADD_LINE_INTERVAL  As Double                   '下點 多一條線點要下去多少  字形的 百分比

    BAR_WITCH As Double    '小節線的粗細

    'IsGiFill As Boolean                 '在閉合的圖形元素上產生填充區域
    IsGiFill As Boolean                 '是否填充區域
    IsDebugDialog As Boolean                 '是否叫出 DebugDialog

    IsDrawCursory As Boolean                 '是否 繪自定的點
    mHowData As Integer        '有多少行資料
    
    iTONE As Integer         ' * 行
    iFinge As Integer    '這是指法行     _+)(*&
    iScale  As Integer     '這是過低音行   .,:
    iNote  As Integer      '這為主行       1234567.|l
    iTempo  As Integer    '這是拍子       -=368acefz
    iTowFinge  As Integer     '這是指法行第二行  _+)(*&
    iSlur  As Integer         ' 連音符行    (3456)
    iAdd As Integer           '這是合音行 [2467]
    
    L_BUF_MD  As Integer   '這是看Buf_md要多少行
    L_WHAT  As Integer    '這是指法看要放在那一行
    
    
    
    f一指 As Integer
    f二指 As Integer
    f三指 As Integer
    f四指 As Integer
    f空弦 As Integer
    f揉弦 As Integer
    
    '2015/1/10 設定圖塊的資料，駐點的位置不一樣
    
    bkMidHight As Double    '物件中間下點的高度
    bkMidUpHight As Double    '物件中間上點的高度
    
    bkInster_Line_H As Double    '物件原點跟拍線的高度
    barInsterNumberSize As Double   '小節號字大小
    
    
    wBar As Double          '小節線的寬度
    wMete As Double         '拍號的寬度
    wNote As Double         '音符的寬度
    wOther As Double
    wNote2note As Double    '每拍到每拍的間距

    extraScale As Double    '前綴元素 比例
End Type

Public amt As AMusicType
Public rTime As runTime

'MusicTextBlock  畫圖塊用資料

Public Type MusicFing
    sTONE As String        ' 升降記號 行
    sFinge As String   '這是指法行     _+)(*&
    sScale As String   '這是高低音行   .,:
    sNote As String   '這為主行       1234567.|l
    sTempo As String   '這是拍子       -=368acefz
    sTowFinge As String   '這是指法行第二行  _+)(*&
    sSlur As String
    
    iScale As Integer     '這是高低音行   .,:
    iNote As Integer     '這為主行       1234567.|l
    iTempo As Integer    '這是拍子       -=368acefz
    mfPosition As point    '這是看這第一指法字上下移多少
    mfTowPosition As point '這是看這第二指法字上下移多少
End Type


'音附的駐點位置
Public Type GripPoints
    gptMid As point       '在正中間
    gptLeftDown As point     '在左下
    gptMidUp As point     '在中上
    gptLeft As point      '在左中
    gptRight As point      '在右中
    atPt As point           '備注點
End Type

    
Public Type Glode
    check1 As Boolean
    fontName As String
    FONTSIZE As Double
    Many As Integer     '幾聲部設定
    bar As Integer
    mete As Integer     '3/4 拍號的分子 --> 3
    mete2 As Integer    '3/4 拍號的分母 --> 4
    
    
    barsperstaff  As Integer '設定每行幾小節
    durationIndex As Double ''累記 現在已經讀取每小節的音符長度總合
    currLine As Integer     '現在是第幾行
    currMeasure As Integer     '現在是 這行的第幾小節
    
    pagewidth As Double
    LeftSpace As Double
    RightSpace As Double
    BarToNoteSpace As Double     '小節線到音符的空白
    TrackToTrack As Double
    LineToLine As Double
    MIN_X As Double
    Beat_MIN_X As Double
    IsBarAlign As Boolean       '小節是否對齊
    
    lastRightPoint As New point '計錄最後一個元素的 右邊點
End Type


    ''''''''''''''''''''''''''''''''''
Sub AMT_LOAD()
    Dim DEFINE_TEXT_SIZE  As Double
    amt.DEFINE_FontFileName = "細明體"              '預設字型的名字
    amt.DEFINE_TEXT_SIZE = 1000              '預設字型的大小
    DEFINE_TEXT_SIZE = 1000
    
    amt.A_TEXT_WIDTH = 726# / DEFINE_TEXT_SIZE               '預設一個字的寬度  字形的 百分比
    
    amt.LINE_PASE = 592.2 / DEFINE_TEXT_SIZE            '這是從原點 要高 字形的 百分比
    amt.LINE_JUMP = 155 / DEFINE_TEXT_SIZE            '每一個拍線和拍線的間格 字形的 百分比
    amt.LINE_THICKNESS = 65 / DEFINE_TEXT_SIZE            '這是拍線的厚度 字形的 百分比
    amt.LINE_LEN = 723.5 / DEFINE_TEXT_SIZE            '這是拍線的長度 字形的 百分比
    
    amt.DROP_UP = 1950# / DEFINE_TEXT_SIZE             '高八度的位置  字形的 百分比
    amt.DROP_DOWN = 500# / DEFINE_TEXT_SIZE             '低八度的位置  字形的 百分比
    amt.DROP_CIRCLE_RADIUS_SIZE = 200# / DEFINE_TEXT_SIZE / 2           '圓的 直徑 SIZE 字形的 百分比
    amt.DROP_INTERVAL = 290# / DEFINE_TEXT_SIZE             '每八度的間距 字形的 百分比
    amt.DROP_ADD_LINE_INTERVAL = 170# / DEFINE_TEXT_SIZE               '下點 多一條線點要下去多少  字形的 百分比
    
    amt.BAR_WITCH = 340 / DEFINE_TEXT_SIZE                '小節線的粗細

    '''bool  IsGiFill    =true ''在閉合的圖形元素上產生填充區域
    amt.IsGiFill = False '是否填充區域
    amt.IsDebugDialog = False      '是否叫出 DebugDialog

    amt.IsDrawCursory = True    '是否 繪自定的點

    amt.mHowData = 9    '有9行的資料

    amt.iTONE = 1        ' * 行
    amt.iFinge = 2    '這是指法行     _+)(*&
    amt.iScale = 3    '這是高低音行   .,:
    amt.iNote = 4     '這為主行       1234567.|l
    amt.iTempo = 5    '這是拍子       -=368acefz
    amt.iTowFinge = 6    '這是指法行第二行  _+)(*&
    amt.iSlur = 7        ' 連音符行    (3456)
    amt.iAdd = 8        ' 合音符行    [3456]
    
    amt.L_BUF_MD = 9  '這是看Buf_md要多少行
    amt.L_WHAT = 8    '這是指法看要放在那一行

    '二胡用
    amt.f一指 = 1
    amt.f二指 = 2
    amt.f三指 = 4
    amt.f四指 = 8
    amt.f空弦 = 16
    amt.f揉弦 = 32
    
    
    
    '2015/1/10 設定圖塊的資料，駐點的位置不一樣
    amt.bkMidHight = 690# / DEFINE_TEXT_SIZE      '物件中間下點的高度
    amt.bkMidUpHight = 1690# / DEFINE_TEXT_SIZE     '物件中間上點的高度
    amt.bkInster_Line_H = 97.8 / DEFINE_TEXT_SIZE    '物件原點跟拍線的高度
    amt.barInsterNumberSize = 0.6
    
    
    amt.wBar = 1    '小節線的寬度
    amt.wMete = 1      '拍號的寬度
    amt.wNote = 0.75        '音符的寬度
    amt.wOther = 1
    amt.wNote2note = 0.5 '每拍到每拍的間距
    
    amt.extraScale = 0.6 '
    
End Sub
    
    
