VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "voiceItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Declare interface ()voiceBase {
Public TS_pos As String

Public typs As Integer
Public v As Integer        ''是第幾聲部
Public soln As Boolean     ''換行 一行結束
Public nl As Boolean       ''換行 新的一行開始
Public row As Integer      ''每聲部的第幾行
Public Col As Integer
Public minspacing As Double
Public extraw As Double
Public w As Double
Public barsperstaff As Variant  ''每行要幾小節
Public dur As Double
Public time As Double
Public fmt As String 'FormationInfo,
Public p_v As String 'PageVoiceTune
Public x As Double
Public y As Double
Public wr As Double
Public wl As Double
Public st As Double
Public a_dd As New iArray 'decorationItem[]
Public a_ly As New iArray 'lyricsItem[]
Public a_gch As New iArray 'gchordItem[]
Public nexs As String 'voiceItem
Public prev As String 'voiceItem
Public ts_prev As String 'voiceItem
Public ts_next As String 'voiceItem
Public err As Boolean



'Declare interface ()voiceBar extends voiceBase {
'Public type As  0 | Double
Public bar_type As String
Public bar_num As Double
Public bar_mrep As Double
Public bar_dotted As Boolean
Public text As String
Public fname As String
Public iStart As Double
Public multi As Double
Public iend As Double
Public invis As Boolean
Public pos As Object
Public seqst As Boolean
Public notes As iArray  'as aNoteItem()
Public nhd As Double
Public Mid As Double
Public ymx As Double
Public ymn As Double
Public shrink As Double
Public space As Double
Public rbstop As Double
Public rbstart As Double
Public xsh As Double
Public norepbra As Boolean

'Declare interface ()voiceClef extends voiceBase {
'   type as  1
Public clef_line As Double
Public clef_type As String
Public clef_auto As Boolean
'Public fname As String
'Public istart As Double
'Public iend As Double
'Public ymx As Double
'Public ymn As Double
'Public seqst As Boolean
'Public shrink As Double
'Public space As Double

'Declare interface ()voiceKey extends voiceBase {
'    type as  5 | Double
'Public istart As Double
'Public iend As Double
Public k_sf As Double
Public k_map As New iArray
Public k_mode As Double
Public k_b40 As Double
'Public seqst As Boolean
Public k_old_sf As Double
Public k_y_clef As Double
'Public ymx As Double
'Public ymn As Double
'Public shrink As Double
'Public space As Double

Public k_bagpipe
Public k_drum
Public k_none As Boolean
Public exp As Boolean
Public k_a_acc As New iArray 'noteItem[]

'Declare interface ()voiceMeter extends voiceBase {
'    type as  6 | Double
Public a_Meter As New iArray '{ top as  String, bot As String }[],
'Public fname As String
'Public istart As Double
'Public iend As Double
Public wmeasure As Double
Public x_meter As New iArray 'Double[]
'Public seqst As Boolean
'Public ymx As Double
'Public ymn As Double
'Public shrink As Double
'Public space As Double
'Public pos As Object
'Public notes As New iArray 'noteItem[]
'Public nhd As Double
'Public mid As Double


'Declare interface ()voiceNote extends voiceBase {
'    type as  8
'Public fname As String
Public stem As Double
'Public multi As Double
'Public nhd As Double
'Public istart As Double        ''這是記錄這物件在文字檔的位置 開始
'Public iend As Double          ''這是記錄這物件在文字檔的位置 結束
'Public notes As New iArray 'noteItem[]
Public dur_orig As Double
'Public pos As Object
Public head As Double
Public dots As Double      ''幾個 負點音符
Public nflags As Double    ''這是拍子 有幾條拍線
Public extra As New iArray 'voiceItem[]  ''裝飾音
Public acc As Double       '' 升降記號
Public beam_end As Boolean
Public beam_st As Boolean
'Public mid As Double
Public xmx As Double
Public ys As Double
'Public ymn As Double
'Public ymx As Double
Public in_tuplet As Boolean
Public tpe As Double
'Public seqst As Boolean
Public stemless As Boolean '' 音符是否無莖的
'Public shrink As Double
'Public space As Double
Public sls
'Public soln As Boolean
Public xs As Double
Public repeat_n As Double
Public repeat_k As Double
Public play As Boolean
'Public invis As Boolean
Public slurStart As iArray
Public slurEnd As iArray


'Declare interface ()voiceRest extends voiceBase {
'    type as  10
'Public fname As String
'Public stem As Double
'Public multi As Double
'Public nhd As Double
'Public xmx As Double
'Public istart As Double
'Public dur_orig As Double
Public fmr As Double
'Public notes As New iArray 'noteItem[]
'Public pos As Object
'Public iend As Double
'Public beam_end As Boolean
'Public head As Double
'Public dots As Double
'Public nflags As Double
'Public stemless As Boolean
'Public beam_st As Boolean
'Public mid As Double
'Public ymx As Double
'Public ymn As Double
'Public invis As Boolean
Public nmes As Double
'Public repeat_n As Double
'Public repeat_k As Double
Public rep_nb As Double
'Public seqst As Boolean
'Public shrink As Double
'Public space As Double
'Public soln As Boolean
Public noplay As Boolean



'declare interface voiceStaves extends voiceBase {
 'public   type as 12 | Double
 'Public fname As String
 'Public st As Double
 Public sY As Object
 '{
 '       voices as {
 '           st as Double
 '           range as Double
 '           sep as Double
 '       }[]
 '       staves as {
 '           stafflines as string
 '           staffscale as Double
 '           staffnonote as Double
 '           maxsep as Double
 '       }[]
 '       top_voice as Double
 '       nstaff as Double
 '   }
 'Public seqst As Boolean
 Public parts As String
 'Public notes As iArray  'noteItem
 'Public nhd As Double
 'Public ymx As Double
 'Public ymn As Double
 'Public shrink As Double
 'Public space As Double


'Declare interface ()voiceTempo extends voiceBase {
'    type as  14 | Double
'Public fname As String
'Public istart As Double
'Public iend As Double
Public tempo_str1 As String
Public tempo_notes As New iArray 'Double[]
Public tempo As Double
'Public notes As New iArray 'noteItem[]
'Public nhd As Double
'Public mid As Double
'Public ymx As Double
'Public ymn As Double
Public tempo_str As String
Public tempo_str2 As String
Public tempo_ca As String
Public new_beat As Double
Public tempo_wh As New iArray  '[Double, Double]
'Public seqst As Boolean
'Public shrink As Double
'Public space As Double
'Public invis As Boolean


'Declare interface ()voiceBlock extends voiceBase {
'    type as  16
Public subtype As String
'Public fname As String
'Public istart As Double
'Public iend As Double
'Public pos As Object
'Public invis As Double
'Public play As Double
Public chn As Double
'Public notes As iArray ' noteItem[]
'Public nhd As Double
'Public mid As Double
'Public ymx As Double
'Public ymn As Double
Public instr As Double

Sub setX(x As Double)
    Me.x = x
'    for (var i = 0  i < this.children.length  i++)
'            this.children[i].setX(x)
End Sub



Public Sub ConvertVarName(key, value)

        Select Case key    ' Evaluate Number.
            Case "type":        Me.typs = value     '這個有改名
            Case "next":        Me.nexs = value     '這個有改名
            
            '' 陣列
            Case "a_dd":        Me.a_dd = value 'a_dd As New iArray 'decorationItem[]
            Case "a_ly":        Me.a_ly = value 'a_ly As New iArray 'lyricsItem[]
            Case "a_gch":       Me.a_gch = value 'a_gch As New iArray 'gchordItem[]
            Case "extra":       Me.extra = value 'extra As New iArray 'voiceItem[]  ''裝飾音
            Case "k_map":       Me.k_map = value 'k_map As New iArray
            Case "a_meter":     Me.a_Meter = value 'a_meter As New iArray '{ top as  String, bot As String }[],
            Case "k_a_acc":     Me.k_a_acc = value 'k_a_acc As New iArray 'noteItem[]
            Case "x_meter":     Me.x_meter = value 'x_meter As New iArray 'Double[]
            Case "tempo_wh":    Me.tempo_wh = value 'tempo_wh As New iArray  '[Double, Double]
            Case "tempo_notes": Me.tempo_notes = value 'tempo_notes As New iArray 'Double[]

            ''物件
            Case "fmt":         Me.fmt = value 'fmt As String 'FormationInfo,
            Case "notes":   Set Me.notes = value 'notes As aNoteItem
            Case "nexs":        Me.nexs = value 'nexs As voiceItem
            Case "pos":     Set Me.pos = value 'pos As Object
            Case "p_v":         Me.p_v = value 'p_v As String 'PageVoiceTune
            Case "prev":        Me.prev = value 'prev As voiceItem
            Case "ts_next":     Me.ts_next = value 'ts_next As voiceItem
            Case "ts_prev":     Me.ts_prev = value 'ts_prev As voiceItem

            
            ''一般
            Case "acc":         Me.acc = value 'acc As Double       '' 升降記號
            Case "bar_dotted":  Me.bar_dotted = value 'bar_dotted As Boolean
            Case "bar_mrep":    Me.bar_mrep = value 'bar_mrep As Double
            Case "bar_num":     Me.bar_num = value 'bar_num As Double
            Case "bar_type":    Me.bar_type = value 'bar_type As String
            Case "beam_end":    Me.beam_end = value 'beam_end As Boolean
            Case "beam_st":     Me.beam_st = value 'beam_st As Boolean
            Case "chn":         Me.chn = value 'chn As Double
            Case "clef_line":   Me.clef_line = value 'clef_line As Double
            Case "clef_type":   Me.clef_type = value 'clef_type As String
            Case "dots":        Me.dots = value 'dots As Double      ''幾個 負點音符
            Case "dur":         Me.dur = value 'dur As Double
            Case "dur_orig":    Me.dur_orig = value 'dur_orig As Double
            Case "err":         Me.err = value 'err As Boolean
            Case "exp":         Me.exp = value 'exp As Boolean
            Case "fmr":         Me.fmr = value 'fmr As Double
            Case "fname":       Me.fname = value 'fname As String
            Case "head":        Me.head = value 'head As Double
            Case "iend":        Me.iend = value 'iend As Double
            Case "in_tuplet":   Me.in_tuplet = value 'in_tuplet As Boolean
            Case "instr":       Me.instr = value 'instr As Double
            Case "invis":       Me.invis = value 'invis As Boolean
            Case "istart":      Me.iStart = value 'istart As Double
            Case "k_b40":       Me.k_b40 = value 'k_b40 As Double
            Case "k_bagpipe":   Me.k_bagpipe = value 'k_bagpipe
            Case "k_drum":      Me.k_drum = value 'k_drum
            Case "k_mode":      Me.k_mode = value 'k_mode As Double
            Case "k_none":      Me.k_none = value 'k_none As Boolean
            Case "k_old_sf":    Me.k_old_sf = value 'k_old_sf As Double
            Case "k_sf":        Me.k_sf = value 'k_sf As Double
            Case "k_y_clef":    Me.k_y_clef = value 'k_y_clef As Double
            Case "mid":         Me.Mid = value 'mid As Double
            Case "multi":       Me.multi = value 'multi As Double
            Case "new_beat":    Me.new_beat = value 'new_beat As Double
            Case "nflags":      Me.nflags = value 'nflags As Double    ''這是拍子 有幾條拍線
            Case "nhd":         Me.nhd = value 'nhd As Double
            Case "nmes":        Me.nmes = value 'nmes As Double
            Case "norepbra":    Me.norepbra = value 'norepbra As Boolean
            Case "play":        Me.play = value 'play As Boolean
            Case "rbstart":     Me.rbstart = value 'rbstart As Double
            Case "rbstop":      Me.rbstop = value 'rbstop As Double
            Case "rep_nb":      Me.rep_nb = value 'rep_nb As Double
            Case "repeat_k":    Me.repeat_k = value 'repeat_k As Double
            Case "repeat_n":    Me.repeat_n = value 'repeat_n As Double
            Case "seqst":       Me.seqst = value 'seqst As Boolean
            Case "shrink":      Me.shrink = value 'shrink As Double
            Case "sls":         Me.sls = value 'sls
            Case "soln":        Me.soln = value 'soln As Boolean
            Case "space":       Me.space = value 'space As Double
            Case "st":          Me.st = value 'st As Double
            Case "stem":        Me.stem = value 'stem As Double
            Case "stemless":    Me.stemless = value 'stemless As Boolean '' 音符是否無莖的
            Case "subtype":     Me.subtype = value 'subtype As String
            Case "tempo":       Me.tempo = value 'tempo As Double
            Case "tempo_ca":    Me.tempo_ca = value 'tempo_ca As String
            Case "tempo_str":   Me.tempo_str = value 'tempo_str As String
            Case "tempo_str1":  Me.tempo_str1 = value 'tempo_str1 As String
            Case "tempo_str2":  Me.tempo_str2 = value 'tempo_str2 As String
            Case "text":        Me.text = value 'text As String
            Case "time":        Me.time = value 'time As Double
            Case "tpe":         Me.tpe = value 'tpe As Double
            Case "v":           Me.v = value 'v As Double
            Case "wl":          Me.wl = value 'wl As Double
            Case "wmeasure":    Me.wmeasure = value 'wmeasure As Double
            Case "wr":          Me.wr = value 'wr As Double
            Case "x":           Me.x = value 'x As Double
            Case "xmx":         Me.xmx = value 'xmx As Double
            Case "xs":          Me.xs = value 'xs As Double
            Case "xsh":         Me.xsh = value 'xsh As Double
            Case "y":           Me.y = value 'y As Double
            Case "ymn":         Me.ymn = value 'ymn As Double
            Case "ymx":         Me.ymx = value 'ymx As Double
            Case "ys":          Me.ys = value 'ys As Double
            Case Else
                Debug.Print "沒有變數 " & key & "名"
        End Select

End Sub

