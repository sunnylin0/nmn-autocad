Attribute VB_Name = "GlyphsUtility"
Option Explicit

Type SymbolFormat
    c As String * 20
    w As Double
    h As Double
    wLeft As Double
    wRight As Double
    hTob As Double
    hBottom As Double
End Type

Public Enum GLS
    asdferr
End Enum


    
    Sub printREF()
        
        Dim rr As Object
        Set rr = ThisDrawing.Blocks
        Dim blockRefObj As AcadBlockReference
        Dim comm As Collection
        Dim i
        Dim ccs As String
        Dim cName As String
        Dim w, h
        Dim midDownPt(0 To 2) As Double
        Dim height As Double
        Dim minExt As Variant
        Dim maxExt As Variant
        Dim sf As SymbolFormat

        midDownPt(0) = 0
        midDownPt(1) = 0
        midDownPt(2) = 0
        For i = 0 To rr.Count - 1
            cName = ThisDrawing.Blocks.item(i).Name
            ccs = Mid(ThisDrawing.Blocks.item(i).Name, 1, 1)
            
            'midDownPt(0) = (i \ 20) * 2
            'midDownPt(1) = i / 20 * 2
            If ccs = "M" Then
            
                Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock(midDownPt, ThisDrawing.Blocks.item(i).Name, 1, 1, 1, 0)
                blockRefObj.GetBoundingBox minExt, maxExt  ' 取的大小
                If minExt(0) < 0 Then
                    If maxExt(0) > 0 Then
                        sf.w = minExt(0)
                        sf.wLeft = minExt(0)
                        sf.wRight = maxExt(0)
                    Else
                        sf.w = minExt(0)
                        sf.wLeft = minExt(0)
                        sf.wRight = 0
                    End If
                Else
                        sf.w = maxExt(0)
                        sf.wLeft = 0
                        sf.wRight = maxExt(0)
                End If
                
                If minExt(1) <= 0 Then
                    If maxExt(1) > 0 Then
                        sf.h = Format(minExt(1), "0.00")
                        sf.hBottom = Format(minExt(1), "0.00")
                        sf.hTob = Format(maxExt(1), "0.00")
                    Else
                        sf.w = Format(minExt(1), "0.00")
                        sf.hBottom = Format(minExt(1), "0.00")
                        sf.hTob = 0
                    End If
                Else
                        sf.h = Format(maxExt(1), "0.00")
                        sf.hBottom = 0
                        sf.hTob = Format(maxExt(1), "0.00")
                End If
                
                Debug.Print cName & " w:" & Format(w, "0.00") & " " & Format(minExt(0), "0.00") & " " & Format(maxExt(0), "0.00") & _
                " h:" & Format(h, "0.00") & " " & Format(minExt(1), "0.00") & " " & Format(maxExt(1), "0.00")
                
                glyphs.Push cName & " " & Format(sf.w, "0.00") _
                    & " " & Format(sf.h, "0.00") _
                    & " " & Format(sf.wLeft, "0.00") _
                    & " " & Format(sf.wRight, "0.00") _
                    & " " & Format(sf.hBottom, "0.00") _
                    & " " & Format(sf.hTob, "0.00")
            End If

        Next
        
        Debug.Print glyphs.ToString(vbCrLf)
'    For i = 0 To glyphs.Count - 1
'        Set sf = glyphs(i)
'
'        Debug.Print sf.c & " " & sf.w & " " & sf.h & " " & sf.wLeft & " " & sf.wRight & " " & sf.hBottom & " " & sf.hTob
'    Next
        
      
    
        
    End Sub

Sub glyphsInit()


    Dim sf As SymbolFormat
    
    sf.c = "MT0122 " & _
"MT0120" & _
"MT0099" & _
"MT0118" & _
"MT0097" & _
"MT0115" & _
"MT0100" & _
"MT0112" & _
"MT0090" & _
"MT0065" & _
"MT0045" & _
"MT0061" & _
"MT0116" & _
"MT0119" & _
"MT0064"

    Dim arr As New iArray
    
    arr.Push "MT" & Format(Asc("z"), "0000") 'ASC_z 122 ChrW(&H6416)  ' z 大指
   arr.Push "MT" & Format(Asc("x"), "0000") 'ASC_x 120 ChrW(&H641A)  ' x 食指
arr.Push "MT" & Format(Asc("c"), "0000") 'ASC_c 99 ChrW(&H641C)  ' c 中指
arr.Push "MT" & Format(Asc("v"), "0000") 'ASC_v 118 ChrW(&H641F)  ' v 無名指
arr.Push "MT" & Format(Asc("a"), "0000") 'ASC_a 97 ChrW(&H6421)  ' a 八度
arr.Push "MT" & Format(Asc("s"), "0000") 'ASC_s 115 ChrW(&H66FC)  ' s 王字指法
arr.Push "MT" & Format(Asc("d"), "0000") 'ASC_d 100 ChrW(&H6700)  ' d 搖指
arr.Push "MT" & Format(Asc("p"), "0000")  'ASC_p 112 ChrW(&H6381)  ' p 卜指法
arr.Push "MT" & Format(Asc("Z"), "0000") 'ASC_Z 90 ChrW(&H6418)  ' 大寫Z 反大指法
arr.Push "MT" & Format(Asc("A"), "0000") 'ASC_A 65 ChrW(&H6413)  ' 大寫A 反八度
arr.Push "MT" & Format(Asc("-"), "0000") 'ASC_- 45 ChrW(&H66F7)  ' - 間一旋
arr.Push "MT" & Format(Asc("="), "0000") 'ASC_= 61 ChrW(&H66F8)  ' = 間二旋
arr.Push "MT" & Format(Asc("t"), "0000")  'ASC_t 116 ChrW(&H670B)  ' t 打圓 托勾
arr.Push "MT" & Format(Asc("w"), "0000") 'ASC_w 119 ChrW(&H670B)  ' w 柔音
arr.Push "MT" & Format(Asc("@"), "0000") 'ASC_w 119 ChrW(&H670B)  '
{"MT0122"
"MT0120"
"MT0099"
"MT0118"
"MT0097"
"MT0115"
"MT0100"
"MT0112"
"MT0090"
"MT0065"
"MT0045"
"MT0061"
"MT0116"
"MT0119"
"MT0064"}
Debug.Print arr.ToString(vbCrLf)
End Sub

