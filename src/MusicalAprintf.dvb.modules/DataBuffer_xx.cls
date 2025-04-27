VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataBuffer_xx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'2013.03.17  V3 正要修改 二胡的版本，因程式之前是用古箏的指法圖，現在改成二胡的指法圖

Const mChar As String = "*||||||"    ' 資料格式 '資料的前面字元有那些
Const TRACK_MAX As Integer = 5      '最大的音軌數

Dim mCharI() As Integer
Dim mMete As Integer        '每小節有幾拍
Dim mFontSize As Integer
Dim mTextNote() As String       '把每行資料放進來 這是有陣列的
Dim mmsa(0 To 1) As String
Dim mOpenATable_ As Boolean

Private Type t
    ln As Integer   '要記錄讀進來的資料是在文件裡的第幾 "行"
    WD As Integer   '要記錄貢進來的資料是在文件要的第幾 "字"
End Type

Private Type KeyFormat
    typs As Cg
    mfingering As String  '指法
    mnote As String           '高低音
    mscale As String          '音階
    mtempo As String         '拍子
    delay As Integer            '拍子的延遲長度
    mtow_fingering As String  '指法 第二個
    mslur As String           '連音符
    and As Boolean              '記錄合音 true為跟後面一音合音 ,false為沒有
    what As Integer             '記錄一些不一樣的指法
    
'    textdata As String
    
'    mfPosition As Double    '這是看這字上下移多少
'    mnPosition As Double    '這是看這字上下移多少
'    msPosition As Double     '這是看這字上下移多少
'    mtPosition As Double    '這是看這字上下移多少
'    mtow_fPosition As Double   '這是看 指法 第二個 這字上下移多少
'    mslPosition As Double    '連音符 字上下移多少
    
    ln As Integer   '要記錄讀進來的資料是在文件裡的第幾 "行"
    WD As Integer   '要記錄貢進來的資料是在文件要的第幾 "字"
End Type


Dim mMusic() As KeyFormat                '了解嗎
Dim mSize() As Long '現在的 mMusic() as KeyFormat 的物件數量
Dim mLogSize() As Long '現在的 mMusic() 的空白陣列 的數量
Dim author As String    '作者
Dim title As String     '標題


Property Get getTITLE() As String
    getTITLE = title
End Property




Private Sub Class_Initialize()
    ReDim mMusic(TRACK_MAX)
    ReDim mMusic(TRACK_MAX, 100)
    ReDim mSize(TRACK_MAX)
    ReDim mLogSize(TRACK_MAX)
    Dim i As Integer
    For i = 0 To TRACK_MAX
        mSize(i) = 0
        mLogSize(i) = UBound(mMusic, 2)
    Next
    
    ReDim mTextNote(amt.mHowData)
    ReDim mCharI(amt.mHowData)
End Sub
Sub Clear()
    Class_Initialize
End Sub
Sub GetDataToBuf__()
    Dim i As Integer
    Dim j As Integer
'
'    Dim tab_m As Integer
'    Dim size_i As Integer

     
    'tab_m = 1
    'numPClen = 1
    
    'Dim Buf_md(0 To L_BUF_MD) As MyClassT
    Dim c As String
    Dim cc As String
    Dim mcc As String
    'Dim bufcc(0 To L_BUF_MD) As MyBufT
    
    Call ttm.setLineChar(mChar, fMete)
    Dim valText As Variant
    Dim numPC As Integer
    
    If textSss = "" Then
        MsgBox ("沒有資料")
        Exit Sub
    End If
    
    valText = Split(textSss, vbCrLf)
    
    numPC = UBound(valText)
    
   
    For i = 0 To numPC
        'Set buf_Range = ActiveDocument.Paragraphs(i).Range
        c = Mid(valText(i), 1, 1)
        mcc = valText(i)
        'mcc = Mid(valText(i), 1, Len(valText(i)) )   ' 這是把最後一個字(vbCrLf)拿掉
        If c = ">" Then
            Exit For
        End If
        ttm.setSaveString (mcc)
        
    Next i
        ttm.ShowData
        
        Selection.EndOf Unit:=wdStory, Extend:=wdMove   '用 Selection.EndOf 移到文件的尾端
    
        Call ttm.setFontSize(fSize, fOpenATable)
        ttm.TextToMusic         '資料整理中
        
        Dim cNote As String
        Dim cScale As String
        Dim cFingering As String
        Dim cTowFingering As String
        Dim cTempo As String
        Dim cSlur As String
        
        
        Dim cposScale As Double
        Dim cposFing As Double
        Dim cposTowFing As Double
        Dim cposSlur As Double
        
        Dim drawTab As New DrawTable
        
        
        Dim tmp As String
        Dim pos As Integer
        
        Call drawTab.StartSet(fMany, fMete, fBar, 10, fOpenATable) '設定表格
        
        Selection.font.Size = fSize
        Selection.font.Name = UserFontName
        pos = 1
        
        '第 0 個位置是看有多少物件
        For i = 1 To ttm.getMTEXT(iNote, 0)
        
            cNote = ttm.getMTEXT(iNote, i)
            cScale = ttm.getMTEXT(iScale, i)
            cFingering = ttm.getMTEXT(iFinge, i)
            cTowFingering = ttm.getMTEXT(iTowFinge, i)
            cTempo = ttm.getMTEXT(iTempo, i)
            cSlur = ttm.getMTEXT(iSlur, i)
            
            cposScale = ttm.getMTEXTPosition(iScale, i)
            cposFing = ttm.getMTEXTPosition(iFinge, i)
            cposTowFing = ttm.getMTEXTPosition(iTowFinge, i)
            cposSlur = ttm.getMTEXTPosition(iSlur, i)
            
            If cNote = "|" Then
                Call drawTab.MoveNext(1, pos, False)
                pos = pos + 1
            ElseIf cNote = "!" Then
                If fOpenATable Then
                    Call drawTab.MoveNext(1, pos, True)
                End If
            ElseIf cNote <> cScale And cNote <> cTempo And cNote <> cFingering Then
                
                Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
                PreserveFormatting:=False                                   'Ctrl+F9
                
                Selection.MoveLeft Unit:=wdCharacter, Count:=1
                Selection.delete Unit:=wdCharacter, Count:=2

                Selection.typsText text:="eq \o(" & cNote & "," & cTempo
                
                If cScale <> "" Then  '高低音
                    Selection.typsText text:=","
                    Selection.font.position = cposScale '這是下移的指令
                    Selection.typsText text:=cScale
                    Selection.font.position = 0   '這是下移的指令
                End If
                
                If Not (cFingering = "" Or cFingering = " ") Then    '第一行指法
                    Selection.typsText text:=","
                    Selection.font.position = cposFing '這是下移的指令
                    Selection.typsText text:=cFingering
                    Selection.font.position = 0 '這是下移的指令
                
                End If
                
                If Not (cTowFingering = "" Or cTowFingering = " ") Then    '第二行指法
                    Selection.typsText text:=","
                    Selection.font.position = cposTowFing '這是下移的指令
                    Selection.typsText text:=cTowFingering
                    Selection.font.position = 0 '這是下移的指令
                End If
                
                If Not (cSlur = "" Or cSlur = " ") Then    '連音符行
                    Selection.typsText text:=","
                    Selection.font.position = cposSlur '這是下移的指令
                    Selection.typsText text:=cSlur
                    Selection.font.position = 0 '這是下移的指令
                End If
                
                
                Selection.typsText text:=")"
                'Selection.Fields.ToggleShowCodes 'Shift+F9
                'Selection.MoveRight Unit:=wdCharacter, Count:=1
                Selection.EndKey Unit:=wdLine
            End If
            
            ''這是看是否要加入空白  或  半拍的符號
            If cTempo = ChrW(&H5D39) Then
                '現在的音符 拍號 是半拍
                '看下一個音符的主音是否為 "!", "|" 小節符號
                If ttm.getMTEXT(iNote, i + 1) = "!" Or ttm.getMTEXT(iNote, i + 1) = "|" Then
                Else
                    '如不是加入一個 半拍 字
                    Selection.typsText text:=ChrW(&H5D39)
                End If
                
            ElseIf cTempo = "" Or cTempo = " " Then
                '現在的音符 拍號 是一拍
                '看現在及下一個音符的主音是否為 "!", "|" 小節符號
                
                If ttm.getMTEXT(iNote, i + 1) = "!" Or ttm.getMTEXT(iNote, i + 1) = "|" _
                    Or ttm.getMTEXT(iNote, i) = "!" Or ttm.getMTEXT(iNote, i) = "|" Then
                
                Else
                    '如不是加入2個 空白 字
                    Selection.typsText text:="  "
                End If
            End If
            
            
            
        Next i
        
End Sub

Private Sub SetMusicItem(ntrack As Integer, ln As Integer, WD As Integer)
'nTrack as integer '要轉換到第幾個音軌資料
'ln As Integer   '要記錄讀進來的資料是在文件裡的第幾 "行"
'wd As Integer   '要記錄貢進來的資料是在文件要的第幾 "字"
    Dim c() As String
    ReDim c(amt.mHowData)
    Dim i As Integer
    Dim j As Integer
    Dim lsup As Integer
    Dim ilo As Integer
    
    '看陣列夠不夠
    If mSize(ntrack) + Len(mTextNote(amt.iNote)) > UBound(mMusic, 2) Then
        ReDim Preserve mMusic(TRACK_MAX, UBound(mMusic, 2) * 2)
    End If
    
    '迴圈每個字
    Dim tmp_data As String
    Dim a As String
    
    Dim objMusic As KeyFormat
    
    Dim c_TONE As String
    Dim c_Note As String
    Dim c_Finge As String
    Dim c_TowFinge As String
    Dim c_Tempo As String
    Dim c_Scale As String
    Dim c_Slur As String
    
  
    
    
    For i = 1 To Len(mTextNote(amt.iNote))
        tmp_data = ""
        For j = amt.iTONE To Len(mChar)
            c(j) = Mid(mTextNote(j), i, 1)
            a = c(j)
            If a = "" Then
                a = " "
            End If
            tmp_data = tmp_data & a
        Next j
        
        
'        c_TONE = c(AMT.iTONE)      ' * 行
'        c_Note = c(AMT.iNote)      '這為主行       1234567.|l
'        c_Finge = c(AMT.iFinge)    '這是指法行     _+)(*&
'        c_TowFinge = c(AMT.iTowFinge) '這是指法行第二行  _+)(*&
'        c_Tempo = c(AMT.iTempo)    '這是拍子       -=368acefz
'        c_Scale = c(AMT.iScale)    '這是高低音行   .,:
'        c_Slur = c(AMT.iSlur)      '二胡連音線，(古箏是合音)
'
        
        ilo = mSize(ntrack)
        Select Case c(amt.iNote)     '這是音階
            Case "0", "1", "2", "3", "4", "5", "6", "7"
                c_Note = c(amt.iNote)
            Case "."
                c_Note = c(amt.iNote)
            Case "-"
                c_Note = c(amt.iNote)
            Case "|"
                If i = 1 Then
'lin'                    mSize(ntrack) = mSize(ntrack) - 1
                    GoTo nogood
                End If
                c_Note = c(amt.iNote)
            Case " ", ""
'lin'                mSize(ntrack) = mSize(ntrack) - 1
                GoTo nogood
            Case Else
                c_Note = c(amt.iNote)
        End Select
        
        
        Select Case c(amt.iFinge)     '這是指法
            Case " "
            Case "("
            Case "|"
            
            Case "a" To "z", "A" To "Z", "-", "="
            Case chr(21) To chr(125)
                c_Finge = c(amt.iFinge)
            Case Else
                c_Finge = c(amt.iFinge)
        End Select
            
        Select Case c(amt.iTowFinge)     '這是指法2 第二行
            Case "a" To "z", "A" To "Z", "-", "="
                c_TowFinge = c(amt.iTowFinge)
            Case "1" To "7"
                c_TowFinge = c(amt.iTowFinge)
            Case " ", "(", "|"
                c_TowFinge = c(amt.iTowFinge)
            Case Else
                c_TowFinge = c(amt.iTowFinge)
        End Select
        

        lsup = 0
        Select Case c(amt.iTempo)     '這是拍子
            Case " "
                c_Tempo = c(amt.iTempo)
            Case "-", "2"
                lsup = 1
                c_Tempo = c(amt.iTempo)
            Case "=", "4", "3"
                lsup = 2
                c_Tempo = c(amt.iTempo)
            Case "6", "8", "A"
                lsup = 3
                c_Tempo = c(amt.iTempo)
            Case "|"
                c_Tempo = c(amt.iTempo)
            Case Else
                c_Tempo = c(amt.iTempo)
        End Select
        
        
        c_Scale = c(amt.iScale)
        c_Slur = c(amt.iSlur)
        
        
        
        'objMusic.mTONE = c_TONE     ' * 行
        objMusic.mnote = c_Note      '這為主行       1234567.|l
        objMusic.mfingering = c_Finge    '這是指法行     _+)(*&
        objMusic.mtow_fingering = c_TowFinge  '這是指法行第二行  _+)(*&
        objMusic.mtempo = c_Tempo   '這是拍子       -=368acefz
        objMusic.mscale = c_Scale   '這是高低音行   .,:
        objMusic.mslur = c_Slur      '二胡連音線，(古箏是合音)
        
        
        objMusic.textdata = tmp_data
        objMusic.ln = ln  '寫入第幾行
        objMusic.WD = i   '寫入第幾字
        
        
        
         mMusic(ntrack, ilo) = objMusic
        
        
        mSize(ntrack) = mSize(ntrack) + 1

    
        
nogood:
    Next i
End Sub
 Public Sub GetDataToBuf(the_data As String)
' 從檔案中讀取一列字串放到特定位置,讀到'*'離開
    Dim ntrack As Integer
    Dim tmp_ntrack As Integer
    Dim buf_i(1000) As String
    Dim i As Integer, tab_m As Integer, vol As Integer
    Dim Name(256) As String
    Dim p As String
    
    
    Dim cc As String
    Dim mcc As String
    '一行一行的分開，放入 tmp_str
    Dim valText As Variant
    Dim numPC As Integer
    If the_data = "" Then
        MsgBox ("沒有資料")
        Exit Sub
    End If
    
    valText = Split(the_data, vbCrLf)
    ntrack = -1
    tmp_ntrack = -1
    numPC = UBound(valText)
    Dim pos As Integer
    Dim ko As Integer
    Dim findSt As Integer
    Dim findEnd As Integer
    
    For i = 0 To numPC
        cc = VBA.Mid(valText(i), 1, 1)
        mcc = valText(i)
        
        If cc = "<" Then
            If ntrack >= 0 Then
                Call SetMusicItem(ntrack, i, 1) '轉換成物件資料
                For ko = 0 To amt.mHowData
                    mTextNote(ko) = ""
                Next
            End If
            If ntrack = -1 Then
                findSt = InStr(1, mcc, "<")
                findEnd = InStr(1, mcc, ">")
                title = Mid(mcc, findSt + 1, findEnd - findSt - 1)
            End If
            ntrack = ntrack + 1
        ElseIf cc = "*" Then
            Call SetMusicItem(ntrack, i, 1) '轉換成物件資料
            
            For ko = 0 To amt.mHowData
                mTextNote(ko) = ""
            Next
            
            pos = amt.iTONE
            mTextNote(pos) = mcc
        ElseIf cc = "|" Then
            pos = pos + 1
            mTextNote(pos) = mcc
        End If
    Next
    
    Call SetMusicItem(ntrack, i, 1) '轉換成物件資料
    
    
    
    
End Sub
Function GetTrackSize() As Integer
'取得有多少軌
    Dim i As Integer
    For i = 0 To TRACK_MAX
        If mSize(i) < 3 Then
            GetTrackSize = i
            Exit Function
        End If
    Next
    GetTrackSize = TRACK_MAX

End Function

Function GetTrackBufferSize(the_track As Integer) As Long
'取得每軌有多少音符
    GetTrackBufferSize = mSize(the_track)
End Function
Function GetData(the_track As Integer, the_item As Long) As String
    'the_track 第幾部
    'the_item 第幾個
    Dim tmp_s As String
    Dim i As Integer
    If (mSize(the_track) > 1 And mSize(the_track) > the_item) Then
        GetData = mMusic(the_track, the_item).textdata
    End If
End Function
