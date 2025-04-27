Attribute VB_Name = "Module_sort"
Option Explicit
Option Compare Binary


Option Base 1
'陣列排序----大
Dim MaxVal As Variant
    
Function SelectionSortMaxTow(TempArray As Variant, TempArrayEnt As Variant)
    Dim MaxIndex As Double
    Dim MaxObj As Object
    Dim i, j As Integer

    ' Step through the elements in the array starting with the
    ' last element in the array.
    For i = UBound(TempArray) To 0 Step -1

        ' Set MaxVal to the element in the array and save the
        ' index of this element as MaxIndex.
        MaxVal = TempArray(i)
        MaxIndex = i

        ' Loop through the remaining elements to see if any is
        ' larger than MaxVal. If it is then set this element
        ' to be the new MaxVal.
        For j = 0 To i
            If TempArray(j) > MaxVal Then
                MaxVal = TempArray(j)
                Set MaxObj = TempArrayEnt(j)
                MaxIndex = j
            End If
        Next j

        ' If the index of the largest element is not i, then
        ' exchange this element with element i.
        If MaxIndex < i Then
            TempArray(MaxIndex) = TempArray(i)
            TempArray(i) = MaxVal
            
            Set TempArrayEnt(MaxIndex) = TempArrayEnt(i)
            Set TempArrayEnt(i) = MaxObj
            
        End If
    Next i
End Function
    
    
    
    
    
    
    
Function SelectionSortMax(TempArray As Variant)
    Dim MaxIndex As Integer
    Dim i, j As Integer

    ' Step through the elements in the array starting with the
    ' last element in the array.
    For i = UBound(TempArray) To 1 Step -1

        ' Set MaxVal to the element in the array and save the
        ' index of this element as MaxIndex.
        MaxVal = TempArray(i)
        MaxIndex = i

        ' Loop through the remaining elements to see if any is
        ' larger than MaxVal. If it is then set this element
        ' to be the new MaxVal.
        For j = 1 To i
            If TempArray(j) > MaxVal Then
                MaxVal = TempArray(j)
                MaxIndex = j
            End If
        Next j

        ' If the index of the largest element is not i, then
        ' exchange this element with element i.
        If MaxIndex < i Then
            TempArray(MaxIndex) = TempArray(i)
            TempArray(i) = MaxVal
        End If
    Next i

End Function

'陣列排序----小
Function SelectionSortMin(TempArray As Variant)
    Dim MinVal As Variant
    Dim MinIndex As Integer
    Dim i, j As Integer

    ' Step through the elements in the array starting with the
    ' last element in the array.
    For i = 1 To UBound(TempArray)
    
        ' Set MaxVal to the element in the array and save the
        ' index of this element as MaxIndex.
        MinVal = TempArray(i)
        MinIndex = i

        ' Loop through the remaining elements to see if any is
        ' larger than MaxVal. If it is then set this element
        ' to be the new MaxVal.
        For j = 1 To UBound(TempArray)
            If TempArray(i) = "" Then
            
            ElseIf TempArray(j) > TempArray(i) Then
                    MinVal = TempArray(j)
                    TempArray(j) = TempArray(i)
                    TempArray(i) = MinVal
                
            ElseIf TempArray(i) = "" Then
                    MinVal = TempArray(j)
                    TempArray(j) = TempArray(i)
                    TempArray(i) = MinVal
            End If
        Next j

    Next
End Function

Function SelectionSort(TempArray As Variant)
    Dim MaxVal As Variant
    Dim MaxIndex As Integer
    Dim i, j As Integer

    ' Step through the elements in the array starting with the
    ' last element in the array.
    For i = UBound(TempArray) To 1 Step -1

        ' Set MaxVal to the element in the array and save the
        ' index of this element as MaxIndex.
        MaxVal = TempArray(i)
        MaxIndex = i

        ' Loop through the remaining elements to see if any is
        ' larger than MaxVal. If it is then set this element
        ' to be the new MaxVal.
        For j = 1 To i
            If TempArray(j) > MaxVal Then
                MaxVal = TempArray(j)
                MaxIndex = j
            End If
        Next j

        ' If the index of the largest element is not i, then
        ' exchange this element with element i.
        If MaxIndex < i Then
            TempArray(MaxIndex) = TempArray(i)
            TempArray(i) = MaxVal
        End If
    Next i

End Function

Sub SelectionSortMyArray()
    Dim TheArray As Variant
    Dim SSS As String
    Dim i As Integer
    ' Create the array.
    TheArray = Array("one", "two", "three", "four", "five", "six", _
        "seven", "eight", "nine", "ten")

    ' Sort the Array and display the values in order.
    SelectionSort TheArray
    For i = 1 To UBound(TheArray)
        SSS = SSS & TheArray(i) & vbCrLf
    Next i
MsgBox SSS
End Sub





