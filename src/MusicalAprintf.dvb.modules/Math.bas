Attribute VB_Name = "Math"


'Math module by arfu
Option Explicit
#If Win64 Then
    #If VBA7 Then
        Public Const MAX_INTEGER As LongLong = 2 ^ 63 - 1
    #Else
        Public Const MAX_INTEGER As Long = 2 ^ 63 - 1
    #End If
#Else
    Public Const MAX_INTEGER As Long = 2 ^ 31 - 1
#End If
Public Const PI As Double = 3.14159265359
Public Const E As Double = 2.71828182846
Public Const PI2 As Double = 1.57079632679
Public Const TAU As Double = 6.28318530718
Public Const GRatio As Double = 1.61803398875

Public Function IsPrime(ByVal x As Long) As Boolean
    Dim c As Integer
    IsPrime = True
    For c = 2 To Abs(x) - 1
        If Abs(x) Mod c = 0 Then
            IsPrime = False
            Exit Function
        End If
    Next
    If x = 0 Then IsPrime = False
End Function


Public Function Odd(ByVal Number As Long) As Long
    Odd = Number * 2 + 1
End Function


Public Function isDivisible(ByVal Number#, Optional ByVal DividedBy# = 2) As Boolean
    isDivisible = Number Mod DividedBy = 0
End Function


Public Function Evaluate(ByVal String1 As String) As Double
    On Error Resume Next
    Dim Excel As Object: Set Excel = CreateObject("Excel.Application")
    Evaluate = Excel.Evaluate(String1)
End Function

    
Public Function Pow(ByVal x#, Optional ByVal y# = 2) As Double
    Pow = (x ^ y)
End Function

Public Function Sqrt(ByVal x#) As Double
    
    If x > 0 Then
        Sqrt = Sqr(x)
    Else
        Sqrt = x
    End If
End Function


Public Function root(ByVal x#, Optional ByVal y As Double = 2) As Double
    root = Abs(x) ^ (1 / y)
End Function


Public Function RandomNum(Optional ByVal Minimum As Single, Optional ByVal Maximum As Single = 1, Optional ByVal Float As Integer, Optional RandomizeNumber As Variant) As Single
    If IsMissing(RandomizeNumber) Then
        Randomize
    Else
        Randomize RandomizeNumber
    End If
    RandomNum = Round((Maximum - Minimum) * Rnd + Minimum, Float)
End Function


Public Function Ceil(ByVal x#) As Long
    Ceil = IIf(Round(x, 0) >= x, Round(x, 0), Round(x, 0) + 1)
End Function


Public Function Trunc(ByVal x#) As Long
    Trunc = IIf(x > 0, Int(x), -Int(-x))
End Function


Public Function Floor(ByVal x#) As Long
    Floor = IIf(Round(x, 0) <= x, Round(x, 0), Round(x, 0) - 1)
End Function


Public Function delta(ByVal a#, Optional ByVal B# = 0, Optional ByVal c# = 0) As Double
    delta = B ^ 2 - 4 * a * c
End Function


Public Function Bhask(ByVal a#, Optional ByVal B# = 0, Optional ByVal c# = 0)
    If delta(a, B, c) < 0 Then Exit Function
    Bhask = Array((-B + Sqr(delta(a, B, c))) / (2 * a), (-B - Sqr(delta(a, B, c))) / (2 * a))
End Function


Public Function min(ParamArray x() As Variant) As Double
    Dim i%
    If UBound(x) = 0 Then      '只有一個
        min = x(0)(0)
        For i = LBound(x(0)) To UBound(x(0))
            If i = 0 Or x(0)(i) < min Then min = x(0)(i)
        Next
    ElseIf UBound(x) > 0 Then  '兩個以上
        min = x(0)
        For i = 0 To UBound(x)
            If x(i) < min Then
                min = x(i)
            End If
        Next
    End If
End Function


Public Function max(ParamArray x() As Variant) As Double
    Dim i%
    If UBound(x) = 0 Then      '只有一個
        For i = LBound(x(0)) To UBound(x(0))
            If i = 0 Or x(0)(i) > max Then
                max = x(0)(i)
            End If
        Next
    ElseIf UBound(x) > 0 Then  '兩個以上
        For i = 0 To UBound(x)
            If x(i) > max Then
                max = x(i)
            End If
        Next
    End If
End Function


Public Function GCD(ByVal a As Long, ByVal B As Long) As Long
    Dim remainder As Long
    If a = 0 Or B = 0 Then Exit Function
    Do
      remainder = Abs(a) Mod Abs(B)
      a = Abs(B)
      B = remainder
    Loop Until remainder = 0
    GCD = a
End Function


Public Function LCM(ByVal a As Long, ByVal B As Long) As Long
    If a = 0 Or B = 0 Then Exit Function
    LCM = (Abs(a) * Abs(B)) \ GCD(a, B)
End Function


Public Function Fact(ByVal N As Long, Optional ByVal StepValue As Long = 1) As Long
    Fact = 1
    For N = N To 1 Step -Abs(StepValue)
        Fact = Fact * N
    Next
End Function


Public Function Fibonacci(ByVal N As Long) As Long
    If N <= 0 Then Exit Function
    Fibonacci = IIf(N = 1, 1, Fibonacci(N - 1) + Fibonacci(N - 2))
End Function


Public Function Mean(ParamArray x() As Variant) As Double
    Dim i%
    For i = LBound(x) To UBound(x)
        Mean = Mean + x(i)
    Next
    Mean = Mean / (UBound(x) + 1)
End Function


Public Function Median(ParamArray x() As Variant) As Double
    Median = x(0)
    If UBound(x) = 0 Then Exit Function
    Median = IIf(UBound(x) Mod 2, (x(UBound(x) \ 2) + x(UBound(x) \ 2 + 1)) / 2, x(UBound(x) \ 2))
End Function


Public Function Variance(ByVal N1#, ByVal n2#) As Double
    Variance = (Mean(N1, n2) - N1) ^ 2 + (Mean(N1, n2) - n2) ^ 2
End Function


'Public Function Mid(ByVal X1#, ByVal X2#) As Double
'    Mid = (X1 + X2) / 2
'End Function


Public Function FindA(ByVal x1#, ByVal x2#, ByVal y1#, ByVal y2#) As Double
    If x1 = x2 Then Exit Function
    FindA = (y1 - y2) / (x1 - x2)
End Function


Public Function Lerp(ByVal x1#, ByVal x2#, ByVal y1#, ByVal y2#, ByVal x#) As Double
    If x1 = x2 Then Exit Function
    Lerp = y1 + (x - x1) * (y2 - y1) / (x2 - x1)
End Function


Public Function LineLineIntersect(ByVal x1#, ByVal y1#, ByVal x2#, ByVal y2#, ByVal x3#, ByVal y3#, ByVal x4#, ByVal y4#)
    Dim x As Double, y As Double
    If (x1 - x2) * (y3 - y4) = (y1 - y2) * (x3 - x4) Then Exit Function
    x = ((x1 * y2 - y1 * x2) * (x3 - x4) - (x1 - x2) * (x3 * y4 - y3 * x4)) / ((x1 - x2) * (y3 - y4) - (y1 - y2) * (x3 - x4))
    y = ((x1 * y2 - y1 * x2) * (y3 - y4) - (y1 - y2) * (x3 * y4 - y3 * x4)) / ((x1 - x2) * (y3 - y4) - (y1 - y2) * (x3 - x4))
    LineLineIntersect = Array(x, y)
End Function


Public Function Distance(ByVal x1#, ByVal x2#, ByVal y1#, ByVal y2#, Optional ByVal Sqrt As Boolean = True) As Double
    Distance = IIf(Sqrt, Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2), (x2 - x1) ^ 2 + (y2 - y1) ^ 2)
End Function


Public Function Distance2(ByVal x1#, ByVal x2#, ByVal y1#, ByVal y2#, ByVal Z1#, ByVal Z2#, Optional ByVal Sqrt As Boolean = True) As Double
    Distance2 = IIf(Sqrt, Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2 + (Z2 - Z1) ^ 2), (x2 - x1) ^ 2 + (y2 - y1) ^ 2 + (Z2 - Z1) ^ 2)
End Function


Public Function Hypot(ByVal x#, ByVal y#) As Double
    Hypot = Sqr(x ^ 2 + y ^ 2)
End Function


Public Function LogN(ByVal x#, ByVal y#) As Double
    LogN = Log(x) / Log(y)
End Function


Public Function ATn2(ByVal x#, ByVal y#) As Double
    ATn2 = IIf(x > 0, Atn(y / x), IIf(x < 0, Atn(y / x) + PI * Sgn(y) + IIf(y = 0, PI, 0), PI / 2 * Sgn(y)))
End Function


Public Function Sec(ByVal x#) As Double
    Sec = 1 / Cos(x)
End Function


Public Function Cosec(ByVal x#) As Double
    Cosec = 1 / Sin(x)
End Function


Public Function Cotan(ByVal x#) As Double
    Cotan = 1 / Tan(x)
End Function


Public Function Radians(ByVal Degrees#) As Double
    Radians = Degrees * 180 / PI
End Function


Public Function Degrees(ByVal Radians#) As Double
    Radians = Radians * PI / 180
End Function


Public Function ASin(ByVal x#) As Double
    ASin = Atn(x / Sqr(-x * x + 1))
End Function


Public Function ACos(ByVal x#) As Double
    ACos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
End Function


Public Function ASec(ByVal x#) As Double
    ASec = Atn(x / Sqr(x * x - 1)) + Sgn((x) - 1) * (2 * Atn(1))
End Function


Public Function ACosec(ByVal x#) As Double
    ACosec = Atn(x / Sqr(x * x - 1)) + (Sgn(x) - 1) * (2 * Atn(1))
End Function


Public Function ACotan(ByVal x#) As Double
    ACotan = Atn(x) + 2 * Atn(1)
End Function


Public Function HSin(ByVal x#) As Double
    HSin = (exp(x) - exp(-x)) / 2
End Function


Public Function HCos(ByVal x#) As Double
    HCos = (exp(x) + exp(-x)) / 2
End Function


Public Function HTan(ByVal x#) As Double
    HTan = (exp(x) - exp(-x)) / (exp(x) + exp(-x))
End Function


Public Function HSec(ByVal x#) As Double
    HSec = 2 / (exp(x) + exp(-x))
End Function


Public Function HCosec(ByVal x#) As Double
    HCosec = 2 / (exp(x) - exp(-x))
End Function


Public Function HCotan(ByVal x#) As Double
    HCotan = (exp(x) + exp(-x)) / (exp(x) - exp(-x))
End Function


Public Function HASin(ByVal x#) As Double
    HASin = Log(x + Sqr(x * x + 1))
End Function


Public Function HACos(ByVal x#) As Double
    HACos = Log(x + Sqr(x * x - 1))
End Function


Public Function HATan(ByVal x#) As Double
    HATan = Log((1 + x) / (1 - x)) / 2
End Function


Public Function HASec(ByVal x#) As Double
    HASec = Log((Sqr(-x * x + 1) + 1) / x)
End Function


Public Function HACosec(ByVal x#) As Double
    HACosec = Log((Sgn(x) * Sqr(x * x + 1) + 1) / x)
End Function


Public Function HACotan(ByVal x#) As Double
    HACotan = Log((x + 1) / (x - 1)) / 2
End Function

Public Function LawCos(ByVal B As Double, ByVal c As Double, ByVal Angle As Double) As Double
    LawCos = B ^ 2 + c ^ 2 - 2 * c * Cos(Angle)
End Function


Function BitMoveLeft(ByRef v As Long, num As Long) As Long
'左移位元 , 需要注意乘2的時候是否會溢出:
    Dim i As Long
    Dim flag As Boolean '是否要把第32位轉換為1
    
    For i = 1 To num
        '判斷第31位是否=1
        If v >= &H40000000 Then
            flag = True
            '把第31位置換為 0
            v = v And &H3FFFFFFF
        Else
            flag = False
        End If
        
        v = v * 2
    Next
    
    If flag Then
        v = v Or &H80000000
    End If
    
    BitMoveLeft = v
End Function




Function BitMoveRight(ByRef B As Long, num As Long) As Long
'右移位元 , 需要注意負數的情況:
    Dim iStart As Long
    
    iStart = 1
    If B < 0 Then
        '第32位置換為 0
        B = B And &H7FFFFFFF
        B = B \ 2
        
        '第31位置換為 1
        B = B Or &H40000000
        
        iStart = 2
    End If
    
    Dim i As Long
    For i = iStart To num
        B = B \ 2
    Next
    
    BitMoveRight = B
End Function
