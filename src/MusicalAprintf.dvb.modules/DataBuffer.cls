VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'2013.03.17  V3 正要修改 二胡的版本，因程式之前是用古箏的指法圖，現在改成二胡的指法圖

Const mChar As String = "*|||||||"    ' 資料格式 '資料的前面字元有那些
Const TRACK_MAX As Integer = 5      '最大的音軌數

Dim mCharI() As Integer
Dim mMete As Integer        '每小節有幾拍
Dim mFontSize As Integer
Dim mTextnote() As String       '把每行資料放進來 這是有陣列的
Dim mmsa(0 To 1) As String
Dim mOpenATable_ As Boolean

Dim currentTextNotePos As Integer
Private Type t
    ln As Integer   '要記錄讀進來的資料是在文件裡的第幾 "行"
    WD As Integer   '要記錄貢進來的資料是在文件要的第幾 "字"
End Type




Private Type KeyFormat
    typs As Cg
    mtone As String       '星指法
    mfingering As String  '指法
    mnote As String           '高低音
    mscale As String          '音階
    mtempo As String         '拍子
    duration As Integer            '拍子的延遲長度
    mtow_fingering As String  '指法 第二個
    mslur As String           '連音符
    mand As String               '記錄合音 true為跟後面一音合音 ,false為沒有
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


Dim mMusic() As MusicItemList  'voiceFormat                '了解嗎
'Dim mSize() As Long '現在的 mMusic() as KeyFormat 的物件數量
Dim pos(10) As Long '記錄現在到那個元素
Dim lastDuration As Double ''記錄上一次的長度
'Dim mLogSize() As Long '現在的 mMusic() 的空白陣列 的數量
Dim author As String    '作者
Dim title As String     '標題

Dim GroupABCLists(8)  As VoiceABCList

Property Get getTITLE() As String
    getTITLE = title
End Property

Public Sub init()
    Dim i As Integer
    For i = 0 To 10
        pos(i) = 0
    Next
End Sub


Private Sub Class_Initialize()
    'ReDim mMusic(TRACK_MAX)
    'ReDim mMusic(TRACK_MAX, 100)
    ReDim mSize(TRACK_MAX)
    ReDim mLogSize(TRACK_MAX)
    Dim i As Integer
    ReDim mMusic(0)
    
    Set mMusic(0) = New MusicItemList
    
    
    ReDim mTextnote(amt.mHowData)
    ReDim mCharI(amt.mHowData)
End Sub
Sub Clear()
    Class_Initialize
End Sub
Private Sub initMusicMain(nTrack As Integer)
    If UBound(mMusic) < nTrack Then
        ReDim Preserve mMusic(nTrack)
        
        Set mMusic(nTrack) = New MusicItemList
    
    End If
End Sub
Private Function joinFormation(mTextnote() As String) As MusicNoteItemList
    Dim moeList As New MusicNoteItemList
    Dim moe As MusicNoteItem
    Dim c() As String
    ReDim c(UBound(mTextnote))
    Dim j As Integer
    Dim pos As Integer
    '取得資料
    For pos = 1 To Len(mTextnote(amt.iNote))
        For j = amt.iTONE To Len(mChar)
            c(j) = Mid(mTextnote(j), pos, 1)
            If c(j) = "" Then c(j) = " "
        Next

        If c(amt.iNote) <> "" Then
            Set moe = New MusicNoteItem
            moe.typs = Cg.note
            moe.mtone = c(amt.iTONE)             ' * 行
            moe.mnote = c(amt.iNote)                  '這為主行       1234567.|l
            moe.mfingering = c(amt.iFinge)       '這是指法行     _+)(*&
            moe.mtow_fingering = c(amt.iTowFinge) '這是指法行第二行  _+)(*&
            moe.mtempo = c(amt.iTempo)            '這是拍子       -=368acefz
            moe.mscale = c(amt.iScale)            '這是高低音行   .,:
            moe.mslur = c(amt.iSlur)
            moe.mand = c(amt.iAdd)
            
            moeList.Push moe
        End If
        
    Next
    Set joinFormation = moeList

End Function

Private Sub SetMusicNoteFormation(moeList() As MusicNoteItemList)
'nTrack as integer '要轉換到第幾個音軌資料
'ln As Integer   '要記錄讀進來的資料是在文件裡的第幾 "行"
'wd As Integer   '要記錄貢進來的資料是在文件要的第幾 "字"
    Dim nTrack As Integer
    Dim pos As Integer
    Dim j As Integer
    Dim lsup As Integer
    Dim count_tuplet As Integer
    Dim flueTuplet As Boolean

    Dim objMusic As MusicItem
    Dim c As MusicNoteItem
    Dim sc As MusicNoteItem
    Dim isAdd As Boolean
    Dim isDots As Boolean
    Dim isRest As Boolean
    
    
    Dim extraList As MusicNoteItemList
    Dim isExtra As Boolean
    
    count_tuplet = 0
    flueTuplet = False
    For nTrack = 0 To UBound(moeList)
        initMusicMain nTrack
        
        For pos = 0 To moeList(nTrack).Count - 1
'            Debug.Print "sMNFOR  ", , nTrack, "  ", pos
'            If pos = 1503 Then
'                pos = pos
'            End If
            '取得資料
            Set c = moeList(nTrack)(pos)
            isDots = False

            If c.typs = Cg.Config Then
                noteConfigSet nTrack, c
                GoTo nogood
            End If
            Select Case c.mnote  '這是音階
                Case "."
                    isDots = True
                Case "Z"
                    '合併小節 Z
                    restConfigSet nTrack, c
                    GoTo nogood
                Case "|"
                        GoTo nogood
                Case " ", ""
                        GoTo nogood
            End Select
            
            If c.mand = "]" Then
                isAdd = False
                If isDots Then
                    objMusic.dots = 1
                Else
                    objMusic.typs = Cg.note
                    objMusic.notes.Push c
                End If
                
                Set objMusic = calDuration2(objMusic, lastDuration)
                mMusic(nTrack).Push objMusic
                
            ElseIf c.mand = "[" Or isAdd = True Then
                If c.mand = "[" Then
                    Set objMusic = New MusicItem
                    Set objMusic.notes = New MusicNoteItemList
                End If
                If isDots Then
                    objMusic.dots = 1
                Else
                    objMusic.typs = Cg.note
                    objMusic.notes.Push c
                End If
                
                isAdd = True
                Set objMusic = calDuration2(objMusic, lastDuration)
            ElseIf c.mtow_fingering = ">" Then
            '前墜字
                If isExtra = False Then
                    Set objMusic = Nothing
                    Set extraList = New MusicNoteItemList
                    isExtra = True
                End If
                extraList.Push c
                
            ElseIf c.mtow_fingering = "<" Then
            '後墜字
                    Set objMusic = mMusic(nTrack).Last
                    If objMusic.rightObjs Is Nothing Then
                        Set objMusic.rightObjs = New MusicNoteItemList
                    End If
                    objMusic.rightObjs.Push c
            Else
                If isDots Then
                    Set objMusic = mMusic(nTrack).Last
                    objMusic.dots = 1
                Set objMusic = calDuration2(objMusic, lastDuration)
                Else
                    
                    Set objMusic = New MusicItem
                    Set objMusic.notes = New MusicNoteItemList
                    
                    objMusic.typs = Cg.note
                    objMusic.notes.Push c
                    Set objMusic = calDuration2(objMusic, lastDuration)
                    mMusic(nTrack).Push objMusic
                End If
            End If
            
            If Not objMusic Is Nothing And isExtra = True Then
                '這是看如有前墜字物件後，有產生 MusicItem
                '就加入前墜字 給 MusicItem
                Set objMusic.extraObjs = extraList
                Set extraList = Nothing
                isExtra = False
                
            End If
            'calExtraw objMusic
        
            If c.mslur = "[" Or c.mslur = "(" Then
                objMusic.slurStart = True       '二胡連音線，(古箏是合音)
            ElseIf c.mslur = "]" Or c.mslur = ")" Then
                objMusic.slurEnd = True
            End If
            
            
            If c.mslur = "{" Then
                objMusic.tupletStart = True        '多連音線
                flueTuplet = True
                count_tuplet = 1
            ElseIf c.mslur = "}" Then
                objMusic.tupletEnd = True
                count_tuplet = count_tuplet + 1
                objMusic.tupletCount = count_tuplet
                count_tuplet = 0
                flueTuplet = False
            ElseIf flueTuplet = True Then
                count_tuplet = count_tuplet + 1
            End If
            
            
            'objMusic.ln = ln  '寫入第幾行
            'objMusic.WD = i   '寫入第幾字
            
nogood:
        Next pos
    Next nTrack
End Sub

Private Sub noteConfigSet(nTrack As Integer, c As MusicNoteItem)
    
    Dim kk As String
    Dim regEx As New RegExp
    regEx.Global = True
    
    If c.configText <> "" Or c.configText <> " " Then
        kk = Mid(c.configText, 1, 1)
        If kk = "M" Then
            '記錄拍號

            Dim meteMatches
            Dim musicIt As MusicItem
            Set musicIt = New MusicItem
            regEx.Pattern = "(\d+)/(\d+)"
            Set meteMatches = regEx.Execute(c.configText)

            If meteMatches.Count >= 1 Then
                musicIt.typs = Cg.meter
                musicIt.mete = meteMatches(0).SubMatches(0)
                musicIt.mete2 = meteMatches(0).SubMatches(1)
            End If
            mMusic(nTrack).Push musicIt
        ElseIf kk = "I" Then
            '設定每行幾小節
            
            Dim regMatches
            Dim rMusicItem As MusicItem
            Set rMusicItem = New MusicItem
            
            regEx.Pattern = ":setbar[\s]+(\d+)\s*?/\s*?(\d+)"
            Set regMatches = regEx.Execute(c.configText)
            If regMatches.Count = 1 Then
                rMusicItem.typs = Cg.Config
                '設定第x節開始，設定每行幾小節
                If regMatches(0).SubMatches.Count = 2 Then
                    rMusicItem.setbarstaffid = regMatches(0).SubMatches(0)
                    rMusicItem.setbarstaff = regMatches(0).SubMatches(1)
                End If
            End If
        
            regEx.Pattern = ":bar[\s]+(\d+)"
            Set regMatches = regEx.Execute(c.configText)
            If regMatches.Count = 1 Then
                rMusicItem.typs = Cg.Config
                rMusicItem.barsperstaff = regMatches(0).SubMatches(0) '設定每行幾小節
            End If
            
            '設定第 x 小節強制換行
            regEx.Pattern = ":linebreak[\s]+(\d+)"
            Set regMatches = regEx.Execute(c.configText)
            If regMatches.Count = 1 Then
                rMusicItem.typs = Cg.Config
                rMusicItem.barlinebreak = regMatches(0).SubMatches(0)
            End If
            
            '設定第 x 小節強制換頁，小節數重 1 開始
            regEx.Pattern = ":pagebreak[\s]+(\d+)"
            Set regMatches = regEx.Execute(c.configText)
            If regMatches.Count = 1 Then
                rMusicItem.typs = Cg.Config
                rMusicItem.barpagebreak = regMatches(0).SubMatches(0)
            End If
            
            
            
            
            mMusic(nTrack).Push rMusicItem
            
            

        End If
    End If


End Sub
Private Sub restConfigSet(nTrack As Integer, c As MusicNoteItem)
'休子符,小節合併
    Dim rMusicItem As MusicItem
    Set rMusicItem = New MusicItem
    
    
    If c.mnote = "Z" Then
        rMusicItem.typs = Cg.rest
        Set rMusicItem.rest = New aRestItem
        rMusicItem.rest.typs = "multimeasure"
        rMusicItem.rest.text = CharToValue(c.mtempo)
        rMusicItem.rest.barNumber = CharToValue(c.mtempo)
    End If
    
    mMusic(nTrack).Push rMusicItem
            
End Sub



Private Function getChars(mTextnote() As String, pos As Integer) As String()
    Dim c() As String
    ReDim c(UBound(mTextnote))
    Dim j As Integer
            '取得資料
        For j = amt.iTONE To Len(mChar)
            
            c(j) = Mid(mTextnote(j), pos, 1)
            If c(j) = "" Then c(j) = " "
        Next j
        getChars = c
End Function
Private Sub SetMusicItem(nTrack As Integer, ln As Integer, WD As Integer)
'nTrack as integer '要轉換到第幾個音軌資料
'ln As Integer   '要記錄讀進來的資料是在文件裡的第幾 "行"
'wd As Integer   '要記錄貢進來的資料是在文件要的第幾 "字"
    Dim c() As String
    ReDim c(amt.mHowData)
    Dim i As Integer
    Dim j As Integer
    Dim lsup As Integer
    

    
    '迴圈每個字
    Dim a As String
    
    'Dim spaceMusic As voiceFormat   '空白
    'Dim spaceNote(0) As noteFormat     '空白
    'ReDim Preserve spaceMusic.notes(0)
    Dim objMusic As MusicItem
    Dim objNote As MusicNoteItem
    
    Dim c_TONE As String
    Dim c_Note As String
    Dim c_Finge As String
    Dim c_TowFinge As String
    Dim c_Tempo As String
    Dim c_Scale As String
    Dim c_Slur As String
    Dim c_Add As String
    Dim isAdd As Boolean
    remMusic nTrack
    For i = 1 To Len(mTextnote(amt.iNote))
        '取得資料
        c = getChars(mTextnote, i)
        Debug.Print "mTextnote(amt.iNote) i: " & i
        If i = 3 Then
            i = i
        End If
        
'        c_TONE = c(AMT.iTONE)      ' * 行
'        c_Note = c(AMT.iNote)      '這為主行       1234567.|l
'        c_Finge = c(AMT.iFinge)    '這是指法行     _+)(*&
'        c_TowFinge = c(AMT.iTowFinge) '這是指法行第二行  _+)(*&
'        c_Tempo = c(AMT.iTempo)    '這是拍子       -=368acefz
'        c_Scale = c(AMT.iScale)    '這是高低音行   .,:
'        c_Slur = c(AMT.iSlur)      '二胡連音線，(古箏是合音)
'
        
        Select Case c(amt.iNote)     '這是音階
            Case "0", "1", "2", "3", "4", "5", "6", "7"
                c_Note = c(amt.iNote)
            Case "."
                'c_Note = c(amt.iNote)
                Dim uu As Integer
                For uu = mMusic(nTrack).Count - 1 To 0 Step -1
                    If mMusic(nTrack)(uu).duration > 0 Then
                        Set objMusic = mMusic(nTrack)(uu)
                        objMusic.identify_Duration objMusic.duration * 1.5
                        'objMusic.calMeWidth
                        GoTo nogood
                    End If
                Next

            Case "-"
                c_Note = c(amt.iNote)
            Case "|"
                If i = 1 Then
'lin'               mSize(ntrack) = mSize(ntrack) - 1
                    GoTo nogood
                End If
                c_Note = c(amt.iNote)
            Case " ", ""
'lin'               mSize(ntrack) = mSize(ntrack) - 1
                    GoTo nogood
            Case Else
                c_Note = c(amt.iNote)
        End Select
        
        
        Select Case c(amt.iFinge)     '這是指法
            Case " "
            Case "("
            Case "|"
            
            Case "a" To "z", "A" To "Z", "-", "="
            Case Chr(21) To Chr(125)
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
        

        Select Case c(amt.iTempo)     '這是拍子
            Case " "
                c_Tempo = c(amt.iTempo)
            Case "-", "2"
                c_Tempo = c(amt.iTempo)
            Case "=", "4", "3"
                c_Tempo = c(amt.iTempo)
            Case "6", "8", "A"
                c_Tempo = c(amt.iTempo)
            Case "|"
                c_Tempo = c(amt.iTempo)
            Case Else
                IsEmpty c(amt.iTempo)
                c_Tempo = c(amt.iTempo)
        End Select
        
        
        c_TONE = c(amt.iTONE)
        c_Scale = c(amt.iScale)
        c_Slur = c(amt.iSlur)   '滑音行
        c_Finge = c(amt.iFinge)
        c_Add = c(amt.iAdd)     '合音行
        
        Set objNote = New MusicNoteItem
        
        objNote.mtone = c_TONE     ' * 行
        objNote.mnote = c_Note      '這為主行       1234567.|l
        objNote.mfingering = c_Finge    '這是指法行     _+)(*&
        objNote.mtow_fingering = c_TowFinge  '這是指法行第二行  _+)(*&
        objNote.mtempo = c_Tempo   '這是拍子       -=368acefz
        objNote.mscale = c_Scale   '這是高低音行   .,:
        


        
        If c_Add = "]" Then
            isAdd = False
            objMusic.typs = Cg.note
            objMusic.notes.Push objNote
            objMusic.mtempo = c_Tempo   '這是拍子       -=368acefz
            lastDuration = calDuration(objMusic, lastDuration)
            objMusic.identify_Duration lastDuration
            mMusic(nTrack).Push objMusic
        ElseIf c_Add = "[" Or isAdd = True Then
            If c_Add = "[" Then
                Set objMusic = New MusicItem
                Set objMusic.notes = New MusicNoteItemList
            End If
            isAdd = True
            objMusic.typs = Cg.note
            objMusic.notes.Push objNote
            objMusic.mtempo = c_Tempo   '這是拍子       -=368acefz
            lastDuration = calDuration(objMusic, lastDuration)
            objMusic.identify_Duration lastDuration
        Else
            Set objMusic = New MusicItem
            Set objMusic.notes = New MusicNoteItemList
            
            objMusic.typs = Cg.note
            objMusic.notes.Push objNote
            objMusic.mtempo = c_Tempo   '這是拍子       -=368acefz
            lastDuration = calDuration(objMusic, lastDuration)
            objMusic.identify_Duration lastDuration
            mMusic(nTrack).Push objMusic
        End If
        ''calExtraw objMusic
    
        If c_Slur = "[" Or c_Slur = "(" Then
            objMusic.slurStart = True       '二胡連音線，(古箏是合音)
        ElseIf c_Slur = "]" Or c_Slur = ")" Then
            objMusic.slurEnd = True
        End If
        
        
        objMusic.ln = ln  '寫入第幾行
        objMusic.WD = i   '寫入第幾字
        
nogood:
    Next i
End Sub
Private Sub clearTextNote(mTextnote() As String)
    Dim i As Integer
    For i = 0 To UBound(mTextnote) - 1
        mTextnote(i) = ""
    Next
End Sub
 Public Sub LoadDataToBuf(the_data As String)
' 從檔案中讀取一列字串放到特定位置,讀到'*'離開
    Dim nTrack As Integer
    Dim tmp_ntrack As Integer
    
    
    Dim cc As String
    Dim mcc As String
    '一行一行的分開，放入 tmp_str
    Dim valText As Variant
    
    Dim ln As Integer
    
    
    Dim pos As Integer
    Dim ko As Integer
    Dim findSt As Integer
    Dim findEnd As Integer
    Dim regEx As New RegExp
    regEx.Global = True
    Dim MusicNoteFormation() As New MusicNoteItemList
    Dim moe As MusicNoteItem
    
    
    
    If the_data = "" Then
        MsgBox ("沒有資料")
        Exit Sub
    End If
    
    valText = Split(the_data, vbCrLf)
    nTrack = -1
    tmp_ntrack = -1
    
    For ln = 0 To UBound(valText)
        cc = VBA.Mid(valText(ln), 1, 1)
        mcc = valText(ln)
        If cc = "<" Then
            If nTrack >= 0 Then
                MusicNoteFormation(nTrack).PushArray joinFormation(mTextnote)
            End If
            If nTrack = -1 Then
                findSt = InStr(1, mcc, "<")
                findEnd = InStr(1, mcc, ">")
                title = Mid(mcc, findSt + 1, findEnd - findSt - 1)
                
            End If
            
            clearTextNote mTextnote
            nTrack = nTrack + 1
            ReDim Preserve MusicNoteFormation(nTrack)
            Set MusicNoteFormation(nTrack) = New MusicNoteItemList
        ElseIf cc = "*" Then
            MusicNoteFormation(nTrack).PushArray joinFormation(mTextnote)
            clearTextNote mTextnote
            
            pos = amt.iTONE
            mTextnote(pos) = mcc
        ElseIf cc = "|" Then
            pos = pos + 1
            mTextnote(pos) = mcc
                        
        ElseIf cc = "M" Or cc = "I" Then
            MusicNoteFormation(nTrack).PushArray joinFormation(mTextnote)
            clearTextNote mTextnote
            
            Set moe = New MusicNoteItem
            moe.typs = Cg.Config
            moe.configText = mcc
            '記錄拍號
            MusicNoteFormation(nTrack).Push moe
            
        End If
    Next
    MusicNoteFormation(nTrack).PushArray joinFormation(mTextnote)
    clearTextNote mTextnote
    SetMusicNoteFormation MusicNoteFormation
    'Call SetMusicItem(nTrack, ln, 1) '轉換成物件資料
    'Call SetMusicItemToABC(ntrack, i, 1)
    
    
    
    
End Sub
'
' Public Sub LoadDataToBuf(the_data As String)
'' 從檔案中讀取一列字串放到特定位置,讀到'*'離開
'    Dim ntrack As Integer
'    Dim tmp_ntrack As Integer
'    Dim buf_i(2000) As String
'    Dim tab_m As Integer, vol As Integer
'    Dim name(256) As String
'    Dim p As String
'
'
'    Dim cc As String
'    Dim mcc As String
'    '一行一行的分開，放入 tmp_str
'    Dim valText As Variant
'    Dim numPC As Integer
'    Dim ln As Integer
'    If the_data = "" Then
'        MsgBox ("沒有資料")
'        Exit Sub
'    End If
'
'    valText = Split(the_data, vbCrLf)
'    ntrack = -1
'    tmp_ntrack = -1
'    numPC = UBound(valText)
'    Dim pos As Integer
'    Dim ko As Integer
'    Dim findSt As Integer
'    Dim findEnd As Integer
'    Dim regEx As New RegExp
'    regEx.Global = True
'    Dim MusicNoteFormation As New MusicNoteItemList
'
'
'
'
'    For ln = 0 To numPC
'        cc = VBA.Mid(valText(ln), 1, 1)
'        mcc = valText(ln)
'        Debug.Print "ln: " & ln
'
'        If cc = "<" Then
'            If ntrack >= 0 Then
'                MusicNoteFormation.PushArray joinFormation(mTextnote)
'                'Call SetMusicItem(ntrack, ln, 1) '轉換成物件資料
'                'Call SetMusicItemToABC(ntrack, i, 1)
'            End If
'            If ntrack = -1 Then
'                findSt = InStr(1, mcc, "<")
'                findEnd = InStr(1, mcc, ">")
'                title = Mid(mcc, findSt + 1, findEnd - findSt - 1)
'            End If
'
'            clearTextNote mTextnote
'            ntrack = ntrack + 1
'        ElseIf cc = "*" Then
'            MusicNoteFormation.PushArray joinFormation(mTextnote)
'            'Call SetMusicItem(ntrack, ln, 1) '轉換成物件資料
'            'Call SetMusicItemToABC(ntrack, i, 1)
'            clearTextNote mTextnote
'
'            pos = amt.iTONE
'            mTextnote(pos) = mcc
'        ElseIf cc = "|" Then
'            pos = pos + 1
'            mTextnote(pos) = mcc
'
'        ElseIf cc = "M" Then
'            '記錄拍號
'            MusicNoteFormation.PushArray joinFormation(mTextnote)
'            'Call SetMusicItem(ntrack, ln, 1) '轉換成物件資料
'            'Call SetMusicItemToABC(ntrack, i, 1)
'            clearTextNote mTextnote
'
'            Dim meteMatches
'            Dim musicIt As MusicItem
'            Set musicIt = New MusicItem
'            regEx.Pattern = "(\d+)/(\d+)"
'            Set meteMatches = regEx.Execute(valText(ln))
'
'            If meteMatches.Count >= 1 Then
'                musicIt.typs = Cg.meter
'                musicIt.mete = meteMatches(0).SubMatches(0)
'                musicIt.mete2 = meteMatches(0).SubMatches(1)
'            End If
'            mMusic(ntrack).Push musicIt
'        ElseIf cc = "I" Then
'            '設定每行幾小節
'            MusicNoteFormation.PushArray joinFormation(mTextnote)
'            'Call SetMusicItem(ntrack, ln, 1) '轉換成物件資料
'            clearTextNote mTextnote
'            Dim regMatches
'            Dim rMusicItem As MusicItem
'            Set rMusicItem = New MusicItem
'            regEx.Pattern = "bar[\s]+(\d+)"
'            Set regMatches = regEx.Execute(valText(ln))
'
'            If regMatches.Count >= 1 Then
'                rMusicItem.typs = Cg.Config
'                rMusicItem.barsperstaff = regMatches(0).SubMatches(0) '設定每行幾小節
'
'            End If
'            mMusic(ntrack).Push rMusicItem
'        End If
'    Next
'
'    Call SetMusicItem(ntrack, ln, 1) '轉換成物件資料
'    'Call SetMusicItemToABC(ntrack, i, 1)
'
'
'
'
'End Sub


Function GetTrackSize() As Integer
'取得有多少軌
'    Dim i As Integer
'    For i = 0 To TRACK_MAX
'        If mSize(i) < 3 Then
'            GetTrackSize = i
'            Exit Function
'        End If
'    Next
'    GetTrackSize = TRACK_MAX
    GetTrackSize = UBound(mMusic) + 1
End Function

Public Function GetTrackBufferSize(the_track As Integer) As Long
'取得每軌有多少音符
    GetTrackBufferSize = mMusic(the_track).Count
End Function
Function GetData(the_track As Integer, the_item As Long) As MusicItem
    'the_track 第幾部
    'the_item 第幾個
    Dim tmp_s As String
    Dim i As Integer
    Dim musList As MusicItemList
    If UBound(mMusic) >= the_track Then
        'Set musList = mMusic(the_track)
        If mMusic(the_track).Count >= the_item Then
            Set GetData = mMusic(the_track)(the_item)
            Exit Function
        End If
    End If
    
    Set GetData = Nothing

End Function
Function GetDataPos(the_track As Integer) As MusicItem
    'the_track 第幾部
    'the_item 第幾個
    Dim tmp_s As String
    Dim i As Integer
    Dim musItem As MusicItemList
    If (mSize(the_track) > 1 And mSize(the_track) > pos(the_track)) Then
        Set musItem = mMusic(the_track)
        Set GetDataPos = musItem(pos(the_track))
        pos(the_track) = pos(the_track) + 1
    End If
End Function

Private Function ConverRegexList(arrCode As iArray, comparValue As String) As String
    Dim sarr() As String
    Dim i As Integer
    Dim rstr As String

        For i = 0 To arrCode.Count - 1
'            sarr = Split(arrCode(i), " ")
'            If iLe = sarr(0) Then
'                rstr = sarr(1)
'                Exit For
'            Else
'                rstr = cc
'            End If
        Next
    ConverRegexList = rstr
End Function
Private Function initTextNoteData() As Boolean
    currentTextNotePos = 1
End Function
Private Function getTextNoteNextKF() As KeyFormat
    Dim j As Integer
    Dim c() As String
    ReDim c(amt.mHowData)
        For j = amt.iTONE To Len(mChar)
            c(j) = Mid(mTextnote(j), currentTextNotePos, 1)
            If c(j) = "" Then c(j) = " "
        Next j
        Dim kg As KeyFormat
        kg.mtone = c(amt.iTONE)
        kg.mnote = c(amt.iNote)
        kg.mfingering = c(amt.iFinge)
        kg.mtow_fingering = c(amt.iTowFinge)
        kg.mscale = c(amt.iScale)
        kg.mtempo = c(amt.iTempo)
        kg.mand = c(amt.iAdd)
        kg.mslur = c(amt.iSlur)
        getTextNoteNextKF = kg
End Function
Private Function updateTextNoteData() As Boolean
        If currentTextNotePos <= Len(mTextnote(amt.iNote)) Then
            currentTextNotePos = currentTextNotePos + 1
            updateTextNoteData = True
        Else
            updateTextNoteData = False
        End If
End Function
Private Function getTempoConverDuration(str As String) As Double
    '拍子的 字元 轉 成  duration
    
    Dim tempo_hj As String
    Dim tempo_val As Variant
    Dim tempo_ll As Variant
    
    tempo_hj = " -2=45368aAcCfFgGzZ"
    tempo_ll = Array(1, 2, 2, 4, 4, 5, 3, 6, 8, 10, 10, 12, 12, 15, 15, 16, 32, 32)
    tempo_val = Array(0, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3, 4, 4, 4, 4, 4, 4, 4, 4)
    

    Dim ii As Integer
    Dim cn As String
    Dim val As Integer
    
    For ii = 0 To Len(tempo_hj) - 1
        cn = Mid(tempo_hj, ii + 1, 1)
        If cn = str Then
            val = tempo_ll(ii)
            getTempoConverDuration = 1 / (val * 4)
            Exit Function
        End If
    Next ii
    
End Function
Private Function getNoteConverPitch(str As String) As Double
    '高低音 的 字元 轉 成
    
    Dim tempo_hj As String
    Dim tempo_val As Variant
    Dim tempo_ll As Variant
    
    tempo_hj = "?;, .:!"
    tempo_ll = Array(-3, -2, -1, 0, 1, 2, 3)
    Dim ii As Integer
    Dim cn As String
    Dim val As Integer
        
    For ii = 0 To Len(tempo_hj) - 1
        cn = Mid(tempo_hj, ii + 1, 1)
        If cn = str Then
            val = tempo_ll(ii)
            getNoteConverPitch = 7 * val
            Exit Function
        End If
    Next ii
End Function
Private Function converMITABC(GroupABCs As VoiceABCList, vABC As VoiceABC, kf As KeyFormat) As VoiceABC
        Select Case kf.mnote     '這是音階
            Case "0"
                If (vABC.rest Is Nothing) Then
                    Set vABC.rest = New vrestNote
                End If
                vABC.el_typs = "note"
                vABC.rest.typs = "rest"
                vABC.duration = getTempoConverDuration(kf.mtempo)
                
            Case "1", "2", "3", "4", "5", "6", "7"
                Dim pitEle As vpitchesNote
                
                If (vABC.pitches Is Nothing) Then
                    Set vABC.pitches = New vpitchesNoteList
                End If
                Set pitEle = New vpitchesNote
                pitEle.Name = kf.mnote
                pitEle.pitch = CInt(kf.mnote) - 1 + getNoteConverPitch(kf.mscale)
                pitEle.verticalPos = pitEle.pitch
                pitEle.highestVert = pitEle.pitch + 6
                
                vABC.el_typs = "note"
                vABC.duration = getTempoConverDuration(kf.mtempo)
                vABC.pitches.Push pitEle
            Case "-"
                Set vABC = GroupABCs.Last
                vABC.duration = (1 / 4)
                vABC.pitches(0).Name = "-"
            Case "."
                Set vABC = GroupABCs.Last
                If (vABC.dots > 0) Then
                    Dim m As Integer
                    Dim d, f As Double
                        d = 1: f = 1
                    For m = 0 To vABC.dots - 1
                        f = f + d / 2
                        d = d / 2
                    Next
                    d = vABC.duration / f
                    vABC.dots = vABC.dots + 1
                    f = d
                    For m = 0 To vABC.dots
                        f = f + d / 2
                        d = d / 2
                    Next
                    vABC.duration = f
                Else
                    vABC.dots = 1
                    vABC.duration = vABC.duration + vABC.duration / 2
                End If
                Set vABC = Nothing

            Case "|", ":", "[", "]"
                Dim lastABC As VoiceABC
                Dim regEx As New RegExp
                
                ' 設定匹配規則
                regEx.Pattern = "[\[\]\|:]"
                Set lastABC = GroupABCs.Last
                If regEx.Test(lastABC.Name) Then
                    Set vABC = lastABC
                End If
                vABC.el_typs = "bar"
                vABC.Name = vABC.Name + kf.mnote
                
                Select Case vABC.Name
                    Case "|": vABC.typs = "bar_thin"
                    Case "||": vABC.typs = "bar_thin_thin"
                    Case ":|": vABC.typs = "bar_right_repeat"
                    Case "|:": vABC.typs = "bar_left_repeat"
                    Case ":|:": vABC.typs = "fasfasdfasf"
                    Case "[|": vABC.typs = "bar_thin_thick"
                    Case "|]": vABC.typs = "bar_thin_thick"
                    Case Else: vABC.typs = "bar_thin"
                End Select
                
            Case " ", ""
                Set vABC = Nothing
            Case Else
                Set vABC = Nothing
        End Select

        Set converMITABC = vABC
End Function
Private Function SetMusicItemToABC(nTrack As Integer, ln As Integer, WD As Integer) As VoiceABCList
'nTrack as integer '要轉換到第幾個音軌資料
'ln As Integer   '要記錄讀進來的資料是在文件裡的第幾 "行"
'wd As Integer   '要記錄貢進來的資料是在文件要的第幾 "字"
    Dim c() As String
    ReDim c(amt.mHowData)
    Dim i As Integer, j As Integer, k As Integer
    Dim lsup As Integer
    Dim ilo As Integer
    
    '看陣列夠不夠
'    If mSize(ntrack) + Len(mTextNote(amt.iNote)) > UBound(mMusic, 2) Then
'        ReDim Preserve mMusic(TRACK_MAX, UBound(mMusic, 2) * 2)
'    End If
    
    '迴圈每個字
    Dim a As String
    
    'Dim spaceMusic As voiceFormat   '空白
    'Dim spaceNote(0) As noteFormat     '空白
    'ReDim Preserve spaceMusic.notes(0)
    Dim objMusic As MusicItem
    Dim objNote As MusicNoteItem
    
    Dim c_TONE As String
    Dim c_Note As String
    Dim c_Finge As String
    Dim c_TowFinge As String
    Dim c_Tempo As String
    Dim c_Scale As String
    Dim c_Slur As String
    Dim c_Add As String
    Dim isAdd As Boolean

    Dim pit As New vpitchesNote
    Dim lastABC As VoiceABC       '記錄上一個音符元件
        
    
    Dim kf As KeyFormat
    Dim vABC As VoiceABC
    Dim pitEle As vpitchesNote
    Dim lastPitchs As vpitchesNoteList '記錄上一個音符元件
    Dim lastPitch As vpitchesNote
    Dim curABCList As VoiceABCList
    
    If GroupABCLists(nTrack) Is Nothing Then
        Set GroupABCLists(nTrack) = New VoiceABCList
    End If
    
        Set curABCList = GroupABCLists(nTrack)
    
    initTextNoteData
    Do
        kf = getTextNoteNextKF
        If curABCList.Count = 84 Then
             a = a
        End If
        
        Select Case kf.mand
            Case "["
                isAdd = True
                Set vABC = New VoiceABC
                Set vABC = converMITABC(curABCList, vABC, kf)
                If Not (vABC Is Nothing) Then
                    curABCList.Push vABC
                End If
                
            Case "]"
                isAdd = False
                Set vABC = curABCList.Last
                Set vABC = converMITABC(curABCList, vABC, kf)
            Case " "
                If isAdd = True Then
                    Set vABC = curABCList.Last
                    Set vABC = converMITABC(curABCList, vABC, kf)
                Else
                    Set vABC = New VoiceABC
                    Set vABC = converMITABC(curABCList, vABC, kf)
                    If Not (vABC Is Nothing) Then
                        curABCList.Push vABC
                    End If
                End If
            Case Else
        End Select
    
    Loop Until updateTextNoteData = False
    Set SetMusicItemToABC = curABCList
    i = 8
'
'    For i = 1 To Len(mTextNote(amt.iNote))
'
'        '取得資料
'        For j = amt.iTONE To Len(mChar)
'            C(j) = Mid(mTextNote(j), i, 1)
'            If C(j) = "" Then C(j) = " "
'        Next j
'
'
''        c_TONE = c(AMT.iTONE)      ' * 行
''        c_Note = c(AMT.iNote)      '這為主行       1234567.|l
''        c_Finge = c(AMT.iFinge)    '這是指法行     _+)(*&
''        c_TowFinge = c(AMT.iTowFinge) '這是指法行第二行  _+)(*&
''        c_Tempo = c(AMT.iTempo)    '這是拍子       -=368acefz
''        c_Scale = c(AMT.iScale)    '這是高低音行   .,:
''        c_Slur = c(AMT.iSlur)      '二胡連音線，(古箏是合音)
''
'        Dim vABC As New VoiceABC
'        Dim pit As New vpitchesNote
'        Dim lastABC As VoiceABC       '記錄上一個音符元件
'        Dim lastPitchs As vpitchesNoteList '記錄上一個音符元件
'
'        If Not (lastABC Is Nothing) Then Set lastPitch = lastABC.pitches
'
'        Select Case C(amt.iAdd)
'            Case "["
'                isAdd = True
'                For k = i To Len(mTextNote(amt.iNote))
'                    '取得資料
'                    For j = amt.iTONE To Len(mChar)
'                        C(j) = Mid(mTextNote(j), i, 1)
'                        If C(j) = "" Then C(j) = " "
'                    Next j
'                Next
'            Case "]"
'            Case " "
'            Case Else
'        End Select
'        ilo = mSize(ntrack)
'        Select Case C(amt.iNote)     '這是音階
'            Case "0"
'            Case "1", "2", "3", "4", "5", "6", "7"
'                pit.pitch = CInt(C(amt.iNote))
'            Case "."
'                lastPitch.duration = lastPitch.duration + lastPitch.duration / 2
'            Case "-"
'                lastPitch.duration = lastPitch.duration + 0.25
'            Case "|"
'                If i = 1 Then
''lin'               mSize(ntrack) = mSize(ntrack) - 1
'                    GoTo nogood
'                End If
'                c_Note = C(amt.iNote)
'            Case " ", ""
''lin'               mSize(ntrack) = mSize(ntrack) - 1
'                    GoTo nogood
'            Case Else
'                c_Note = C(amt.iNote)
'        End Select
'
'
'
'        lsup = 0
'        Select Case C(amt.iTempo)     '這是拍子
'            Case " "
'                c_Tempo = C(amt.iTempo)
'            Case "-", "2"
'                lsup = 1
'                c_Tempo = C(amt.iTempo)
'            Case "=", "4", "3"
'                lsup = 2
'                c_Tempo = C(amt.iTempo)
'            Case "6", "8", "A"
'                lsup = 3
'                c_Tempo = C(amt.iTempo)
'            Case "|"
'                c_Tempo = C(amt.iTempo)
'            Case Else
'                IsEmpty C(amt.iTempo)
'                c_Tempo = C(amt.iTempo)
'        End Select
'
'
'
'        Select Case C(amt.iFinge)     '這是指法
'            Case " "
'            Case "("
'            Case "|"
'
'            Case "a" To "z", "A" To "Z", "-", "="
'            Case chr(21) To chr(125)
'                c_Finge = C(amt.iFinge)
'            Case Else
'                c_Finge = C(amt.iFinge)
'        End Select
'
'        Select Case C(amt.iTowFinge)     '這是指法2 第二行
'            Case "a" To "z", "A" To "Z", "-", "="
'                c_TowFinge = C(amt.iTowFinge)
'            Case "1" To "7"
'                c_TowFinge = C(amt.iTowFinge)
'            Case " ", "(", "|"
'                c_TowFinge = C(amt.iTowFinge)
'            Case Else
'                c_TowFinge = C(amt.iTowFinge)
'        End Select
'
'
'
'
'        c_TONE = C(amt.iTONE)
'        c_Scale = C(amt.iScale)
'        c_Slur = C(amt.iSlur)   '滑音行
'        c_Finge = C(amt.iFinge)
'        c_Add = C(amt.iAdd)     '合音行
'
'        Set objNote = New MusicNoteItem
'
'        objNote.mtone = c_TONE     ' * 行
'        objNote.mnote = c_Note      '這為主行       1234567.|l
'        objNote.mfingering = c_Finge    '這是指法行     _+)(*&
'        objNote.mtow_fingering = c_TowFinge  '這是指法行第二行  _+)(*&
'        objNote.mtempo = c_Tempo   '這是拍子       -=368acefz
'        objNote.mscale = c_Scale   '這是高低音行   .,:
'
'
'
'
'        If c_Add = "]" Then
'            isAdd = False
'            objMusic.typs = Cg.note
'            objMusic.notes.Push objNote
'            objMusic.mtempo = c_Tempo   '這是拍子       -=368acefz
'
'            mMusic(ntrack).Push objMusic
'            mSize(ntrack) = mSize(ntrack) + 1
'        ElseIf c_Add = "[" Or isAdd = True Then
'            If c_Add = "[" Then
'                Set objMusic = New MusicItem
'                Set objMusic.notes = New iArray
'            End If
'            isAdd = True
'            objMusic.typs = Cg.note
'            objMusic.notes.Push objNote
'            objMusic.mtempo = c_Tempo   '這是拍子       -=368acefz
'        Else
'            Set objMusic = New MusicItem
'            Set objMusic.notes = New iArray
'
'            objMusic.typs = Cg.note
'            objMusic.notes.Push objNote
'            objMusic.mtempo = c_Tempo   '這是拍子       -=368acefz
'
'            mMusic(ntrack).Push objMusic
'            mSize(ntrack) = mSize(ntrack) + 1
'        End If
'
'        If c_Slur = "[" Or c_Slur = "(" Then
'            objMusic.slurStart = True       '二胡連音線，(古箏是合音)
'        ElseIf c_Slur = "]" Or c_Slur = ")" Then
'            objMusic.slurEnd = True
'        End If
'
'
'        objMusic.ln = ln  '寫入第幾行
'        objMusic.WD = i   '寫入第幾字
'
'nogood:
'    Next i
End Function
Function getVoiceListData(trackNum As Integer) As VoiceABCList
    If UBound(GroupABCLists) >= trackNum Then
        Set getVoiceListData = GroupABCLists(trackNum)
    End If

End Function
Function ConverToVoiceABCList(trackNum As Integer) As VoiceABCList

    Dim mItem As MusicItem
    Dim trackMaxLen As Integer
    Dim currenCount As Integer
    Dim sarr() As String
    Dim reg As New RegExp
    Dim cc As String
    Dim sNote As String
    Dim iTempo As Integer
    Dim sTempo As Integer
    Dim iScale As Integer
    Dim sScale As Integer
    Dim i As Integer
    Dim vABC As VoiceABC

'取得關鍵字  主音 AMT.iNote
    Dim noteListCode As New iArray
    Dim noteListType As iArray
            
        noteListCode.PushArray Array("0 M-NOTE0", _
            "1 M-NOTE1", _
            "2 M-NOTE2", _
            "3 M-NOTE3", _
            "4 M-NOTE4", _
            "5 M-NOTE5", _
            "6 M-NOTE6", _
            "7 M-NOTE7", _
            ". M-NOTE.", _
            "- M-NOTE-", _
            "o M-NOTEo", _
            "x M-NOTEx", _
            "X M-NOTEX2", _
            "e M-NOTEe", _
            "E M-NOTEE2")
            
    Dim iLe As Integer
    Dim cn As String
    Dim tempoListCode As New iArray

    tempoListCode.PushArray Array( _
            "1 M-TEMPO1", _
            "2 M-TEMPO2", _
            "3 M-TEMPO3", _
            "4 M-TEMPO4")
            
    
    Dim scaleListCode As New iArray
    scaleListCode.PushArray Array( _
            ". M-OCTAVE1 1", _
            ": M-OCTAVE2 2", _
            "! M-OCTAVE3 3", _
            ", M-OCTAVEa -1", _
            "  M-OCTAVEb -2", _
            "? M-OCTAVEc -3")
            
    '取得升降字 -特疏字法
    Dim toneListCode As New iArray
    toneListCode.PushArray Array( _
            "# MT-SHARP", _
            "b MT-FLAT", _
            "o MT-NATURAL")


    
    trackMaxLen = GetTrackBufferSize(trackNum)
    For i = 0 To trackMaxLen - 1
        Set mItem = GetData(trackNum, i)
                
        cc = ConverRegexList(noteListCode, "mItem.notes")
        
'        For j = 0 To noteListCode.Count - 1
'            sarr = Split(noteListCode(i), " ")
'            If cc = sarr(0) Then
'                MFS(currenCount).sNote = sarr(1)
'                Exit For
'            Else
'                MFS(currenCount).sNote = ""
'            End If
'        Next
        
        
    Next
    
    cc = ""
    
'
'
'    ReDim MFS(MData.notes.Count - 1)
'    For currenCount = 0 To MData.notes.Count - 1
'
'        cc = MData.notes(currenCount).mnote
'        For i = 0 To noteListCode.Count - 1
'            sarr = Split(noteListCode(i), " ")
'            If cc = sarr(0) Then
'                MFS(currenCount).sNote = sarr(1)
'                Exit For
'            Else
'                MFS(currenCount).sNote = ""
'            End If
'        Next
'
'
'        '取得關鍵字 拍子  -=368acefz amt.iTempo
'        iLe = Tempo_c2line(MData.notes(currenCount).mtempo)
'        cc = MData.notes(currenCount).mtempo
'        For i = 0 To tempoListCode.Count - 1
'            sarr = Split(tempoListCode(i), " ")
'            If iLe = sarr(0) Then
'                MFS(currenCount).sTempo = sarr(1)
'                MFS(currenCount).iTempo = iLe
'                Exit For
'            Else
'                MFS(currenCount).sTempo = cc
'                MFS(currenCount).iTempo = iLe
'            End If
'        Next
'
'
'        '取得關鍵字高低音行   .:! , ?
'        cc = MData.notes(currenCount).mscale
'        For i = 0 To scaleListCode.Count - 1
'            sarr = Split(scaleListCode(i), " ")
'            If cc = sarr(0) Then
'                MFS(currenCount).sScale = sarr(1)
'                MFS(currenCount).iScale = sarr(2)
'                Exit For
'            Else
'                MFS(currenCount).sScale = ""
'                MFS(currenCount).iScale = 0
'            End If
'        Next
'
'
'        '取得關鍵字 -指法1
'        MFS(currenCount).sFinge = ""
'        cc = MData.notes(currenCount).mfingering
'        Select Case cc
'            Case "a" To "z", "A" To "Z", "-", "=", "@"
'                MFS(currenCount).sFinge = kFing(Asc(cc))
'                Set MFS(currenCount).mfPosition = posFing(Asc(cc)).clone
'            Case Else
'            MFS(currenCount).sFinge = ""
'            Set MFS(currenCount).mfPosition = New point
'        End Select
'
'
'
'    '    '取得關鍵字 -指法2
'        cc = MData.notes(currenCount).mtow_fingering
'        MFS(currenCount).sTowFinge = ""
'        Select Case cc
'            Case "a" To "z", "A" To "Z", "-", "=", "@"
'                MFS(currenCount).sTowFinge = kTowFing(Asc(cc))
'                Set MFS(currenCount).mfTowPosition = posTowFing(Asc(cc)).clone
'            Case Else
'            MFS(currenCount).sTowFinge = ""
'            Set MFS(currenCount).mfTowPosition = New point
'        End Select
'
'        '取得升降字 -特疏字法
'        cc = MData.notes(currenCount).mtone
'        For i = 0 To toneListCode.Count - 1
'            sarr = Split(toneListCode(i), " ")
'            If cc = sarr(0) Then
'                MFS(currenCount).sTONE = sarr(1)
'                Exit For
'            Else
'                MFS(currenCount).sTONE = cc
'            End If
'        Next
'
'
'    Next

End Function
