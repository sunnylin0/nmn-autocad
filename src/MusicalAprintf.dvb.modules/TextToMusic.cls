VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextToMusic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mHowData As Integer      '有多少行資料
Dim mChar As String          '資料的前面字元有那些
Dim mCharI() As Integer
Dim mMete As Integer        '每小節有幾拍
Dim mFontSize As Integer
Dim mTextnote() As String       '把每行資料放進來 這是有陣列的
Dim mmsa(0 To 1) As String

Dim i As Integer
Dim j As Integer
Private Type KeyFormat
        mfingering As String  '指法
        mnote As String           '高低音
        mscale As String          '音階
        mtempo As String         '拍子
        mfPosition As Double    '這是看這字上下移多少
        mnPosition As Double    '這是看這字上下移多少
        msPosition As Double     '這是看這字上下移多少
        mtPosition As Double    '這是看這字上下移多少
End Type

Dim mMusic() As KeyFormat                '了解嗎
Public Sub setLineChar(Char As String, mete As Integer)
    'Char 設定資料的前面字元有那些
    'Mete 每小節有幾拍
    mChar = Char
    mMete = mete
    ReDim mCharI(Len(mChar))
    mHowData = Len(mChar)
    ReDim mTextnote(mHowData + 1)    ' 陳述式將重新調整陣列的大小，且將所有元素內容清除。
                                     ' 因為它從 0 開始 所以要加 1
End Sub

Public Sub setFontSize(FONTSIZE As Integer)
    mFontSize = FONTSIZE
End Sub

Public Sub TextToMusic()
    Dim kNUM1 As String, kNUM2 As String, kNUM3 As String, kNUM4 As String
    Dim kNUM5 As String, kNUM6 As String, kNUM7 As String, kNUM0 As String
    Dim k延長一拍 As String
    Dim kLIN1 As String: Dim kLIN2 As String: Dim kLIN3 As String
    Dim kUP1 As String
    Dim kUP2 As String
    Dim kDOW1 As String
    Dim kDOW2 As String
    Dim kAdot As String
    Dim kBarLine As String
    Dim kFing(30 To 255) As String
    Dim posLinDown As Variant
    Dim posFing(0 To 255) As Integer
    Const MAX200 As Integer = 200
    posLinDown = Array(0, 12, 26, 40)   '這是看低音下點要移多少下去，是當200字型時計算
    
    kNUM1 = ChrW(&H4E52)
    kNUM2 = ChrW(&H4E53)
    kNUM3 = ChrW(&H4E56)
    kNUM4 = ChrW(&H4E58)
    kNUM5 = ChrW(&H4E59)
    kNUM6 = ChrW(&H4E5C)
    kNUM7 = ChrW(&H4E5D)
    kNUM0 = ChrW(&H5D4C)
    k延長一拍 = ChrW(&H5D1F)
    kLIN1 = ChrW(&H5D39): kLIN2 = ChrW(&H5D3A): kLIN3 = ChrW(&H5D3C)
    kUP1 = ChrW(&H66E0): kUP2 = ChrW(&H66E3): kDOW1 = ChrW(&H66E6): kDOW2 = ChrW(&H66E8)
    kAdot = ChrW(&H5D3D)            '符點音符
    kBarLine = ChrW(&H601D)
    kFing(Asc("z")) = ChrW(&H6416)  ' z 大指
    kFing(Asc("x")) = ChrW(&H641A)  ' x 食指
    kFing(Asc("c")) = ChrW(&H641C)  ' c 中指
    kFing(Asc("v")) = ChrW(&H641F)  ' v 無名指
    kFing(Asc("a")) = ChrW(&H6421)  ' a 八度
    kFing(Asc("s")) = ChrW(&H66FC)  ' s 王字指法
    kFing(Asc("d")) = ChrW(&H6700)  ' d 搖指
    kFing(Asc("p")) = ChrW(&H6381)  ' p 卜指法
    kFing(Asc("Z")) = ChrW(&H6418)  ' 大寫Z 反大指法
    kFing(Asc("A")) = ChrW(&H6413)  ' 大寫A 反八度
    kFing(Asc("-")) = ChrW(&H66F7)  ' - 間一旋
    kFing(Asc("=")) = ChrW(&H66F8)  ' = 間二旋
    kFing(Asc("t")) = ChrW(&H670B)  ' t 打圓 托勾
    
    posFing(Asc("z")) = 157  ' z 大指
    posFing(Asc("x")) = 165  ' x 食指
    posFing(Asc("c")) = 157  ' c 中指
    posFing(Asc("v")) = 157  ' v 無名指
    posFing(Asc("a")) = 157  ' a 八度
    posFing(Asc("s")) = 157  ' s 王字指法
    posFing(Asc("d")) = 0  ' d 搖指
    posFing(Asc("p")) = -10  ' p 卜指法
    posFing(Asc("Z")) = 157  ' 大寫Z 反大指法
    posFing(Asc("A")) = 157  ' 大寫A 反八度
    posFing(Asc("-")) = 157  ' - 間一旋
    posFing(Asc("=")) = 157  ' = 間二旋
    posFing(Asc("t")) = 160  ' t 打圓 托勾

    
    Dim c() As String
    ReDim c(Len(mChar))
    ReDim mMusic(Len(mTextnote(MkDoc.iNote)) * 2)
    Dim lsup As Integer
    Dim pos As Integer
    pos = 1
    Call atImputer(mMete)  '這是要加入小節線 ,排拍子
    For i = 1 To Len(mTextnote(MkDoc.iNote))  ' 迴圈每個字
        For j = 1 To Len(mChar)
            c(j) = Mid(mTextnote(j), i, 1)
        Next j
        
        Select Case c(MkDoc.iFinge)     '這是指法
            Case "a" To "z", "A" To "Z", "-", "="
                mMusic(pos).mfingering = kFing(Asc(LCase(c(MkDoc.iFinge))))
                mMusic(pos).mfPosition = posFing(Asc(LCase(c(MkDoc.iFinge)))) _
                                        * mFontSize / MAX200
            Case " "
            Case "("
            Case "|"
            Case Else
                mMusic(pos).mfingering = c(MkDoc.iFinge)
                mMusic(pos).mfPosition = 0
        End Select
        
        lsup = 0
        Select Case c(MkDoc.iTempo)     '這是拍子
            Case " "
            Case "-", "2"
                lsup = 1
                mMusic(pos).mtempo = kLIN1
            Case "=", "4", "3"
                lsup = 2
                mMusic(pos).mtempo = kLIN2
            Case "6", "8", "A"
                lsup = 3
                mMusic(pos).mtempo = kLIN3
            Case "|"
            Case Else
                mMusic(pos).mtempo = c(MkDoc.iTempo)
        End Select
        
                
        Select Case c(MkDoc.iScale)     '這是 八度音
            Case " "
            Case "."        '上一點
                mMusic(pos).mscale = kUP1
                mMusic(pos).msPosition = 0
            Case ":"        '上二點
                mMusic(pos).mscale = kUP2
                mMusic(pos).msPosition = 0
            Case ","        '下一點
                mMusic(pos).mscale = kDOW1
                mMusic(pos).msPosition = -mFontSize * posLinDown(lsup) / MAX200   '這是下移的指令
            Case " "        '上二點
                mMusic(pos).mscale = kDOW2
                mMusic(pos).msPosition = -mFontSize * posLinDown(lsup) / MAX200   '這是下移的指令
            Case "|"
            Case Else
                mMusic(pos).mscale = c(MkDoc.iScale)
        End Select
        
        
        Select Case c(MkDoc.iNote)     '這是音階
            Case "0"
                mMusic(pos).mnote = kNUM0
            Case "1"
                mMusic(pos).mnote = kNUM1
            Case "2"
                mMusic(pos).mnote = kNUM2
            Case "3"
                mMusic(pos).mnote = kNUM3
            Case "4"
                mMusic(pos).mnote = kNUM4
            Case "5"
                mMusic(pos).mnote = kNUM5
            Case "6"
                mMusic(pos).mnote = kNUM6
            Case "7"
                mMusic(pos).mnote = kNUM7
            Case "."
                mMusic(pos).mnote = kAdot
            Case "-"
                mMusic(pos).mnote = k延長一拍
            Case "|"
                mMusic(pos).mnote = "|"
            Case Else
                mMusic(pos).mnote = c(MkDoc.iNote)
                
        End Select

        pos = pos + 1
    Next i
    mMusic(0).mnote = pos
    
End Sub
Private Sub atImputer(iMete As Integer)
    '排成有小節線的文字
    'iMete 每小節裡有幾拍
    Dim testTXT() As String
    ReDim testTXT(Len(Me.TextNote(MkDoc.iNote)))
    Dim c() As String
    ReDim c(Len(mChar))
    
    Dim i As Integer
    Dim j As Integer
    Dim num As Integer
    Dim nownum As Integer
    Dim chackFor As Boolean
    Const A_TEMPO As Integer = 240  '一拍的數值
    
    For i = 1 To Len(mTextnote(MkDoc.iNote))  ' 迴圈每個字
        For j = 1 To Len(mChar)
            c(j) = Mid(mTextnote(j), i, 1)
        Next j
        
        If num >= A_TEMPO * iMete Then
            If c(MkDoc.iNote) <> "|" Then
                num = 0
                For j = 1 To Len(mChar)
                    testTXT(j) = testTXT(j) & "|"
                Next j
            End If
        End If
        
        Select Case c(MkDoc.iNote)
            Case "|":
                num = 0
                nownum = 0
                chackFor = True
            Case "-":
                nownum = A_TEMPO
                chackFor = True
            Case ".":
                nownum = nownum / 2
                chackFor = True
                
            Case " ":
            Case Else:
                Select Case c(MkDoc.iTempo)
                    Case " ", "":
                        nownum = A_TEMPO
                    Case "-", "2":
                        nownum = A_TEMPO / 2
                    Case "=", "4":
                        nownum = A_TEMPO / 4
                    Case "3":
                        nownum = A_TEMPO / 3
                    Case "6":
                        nownum = A_TEMPO / 6
                    Case "8":
                        nownum = A_TEMPO / 8
                    Case "a":
                        nownum = A_TEMPO / 10
                    Case "":
                        
                End Select
                
                chackFor = True
            
        End Select
        If chackFor = True Then
            For j = 1 To Len(mChar)
                If c(j) = "" Then
                    testTXT(j) = testTXT(j) & " "
                Else
                    testTXT(j) = testTXT(j) & c(j)
                End If
            Next j
        End If
        chackFor = False
        num = num + nownum
        
    Next i
    
    For j = 1 To Len(mChar)
        mTextnote(j) = testTXT(j)
    Next j
End Sub
Private Function intComp(myInt1 As Integer, myInt2 As Integer) As Integer
    '比較數值
    '傳回值 = 較大的數值
    If myInt1 >= myInt2 Then
        intComp = myInt1
    Else
        intComp = myInt2
    End If
End Function
Private Sub mTextNoteLenAdd()
    '這函式是要 mTextNote(?) 中的每個長度一樣，不一樣補空白
    Dim x As Integer
    Dim str As String
    Dim BigInt As Integer
    BigInt = 0
    For x = 1 To mHowData
        BigInt = intComp(BigInt, Len(mTextnote(x)))
    Next
    If BigInt <> 0 Then
        
        For x = 1 To mHowData
            str = space(BigInt)
            LSet str = mTextnote(x)
            mTextnote(x) = str
        Next
    
    End If
End Sub

Public Sub setSaveString(tObject As String)
    '這是看進來什麼格式的資料 要放在 mTextNote 的第幾個
    Dim c As String     '記現在是什麼格式的資料
    Static cl As String    '記前一個是什麼格式的資料
    Static num As Integer   ' 記錄著一樣格式的標記 看以經有幾個了 如 |||||
    c = Mid(tObject, 1, 1)
    For i = 1 To mHowData
        If c = Mid(mChar, i, 1) Then
            If i = 1 Then
                ReDim mCharI(Len(mChar))
                Call mTextNoteLenAdd    '這函式是要 mTextNote(?) 中的每個長度一樣，不一樣補空白
            End If
            If mCharI(i) <> 1 Then
                mTextnote(i) = mTextnote(i) + tObject
                mCharI(i) = 1
                Exit For
            End If
        End If
    Next i
    
   
End Sub

Public Property Get TextNote(ii As Integer) As String
    TextNote = mTextnote(ii)
End Property

Public Sub Show()
    Dim ccc As String
    For i = 1 To mHowData
        ccc = ccc + vbCrLf
        ccc = ccc + Mid(mChar, i, 1)
    Next i
    MsgBox "有 " & mHowData & " 個資料   資料的格式是" & ccc
End Sub

Public Sub ShowData()
    Dim ccc As String
    
    For i = 1 To mHowData
        ccc = ccc + vbCrLf
        ccc = ccc & mTextnote(i)
    Next i
    MsgBox " 全部資料的 內容是" & ccc

End Sub
Public Function getMFingering(ii As Integer) As String
    getMFingering = mMusic(ii).mfingering
End Function
Public Function getMNote(ii As Integer) As String
    getMNote = mMusic(ii).mnote
End Function
Public Function getMScale(ii As Integer) As String
    getMScale = mMusic(ii).mscale
End Function
Public Function getMTempo(ii As Integer) As String
    getMTempo = mMusic(ii).mtempo
End Function
Public Function getMScalePosition(ii As Integer) As Double
    getMScalePosition = mMusic(ii).msPosition
End Function
Public Function getMFingPosition(ii As Integer) As Double
    getMFingPosition = mMusic(ii).mfPosition
End Function

Public Function getSaveString(ii As Integer) As String
    getSaveString = mTextnote(ii)
End Function

