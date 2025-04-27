Attribute VB_Name = "ModuleTextStyle"
Option Explicit

'�r��T�|����
Public Const LF_FACESIZE = 32

Public Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName(LF_FACESIZE) As Byte
End Type

Public Type NEWTEXTMETRIC
        tmHeight As Long
        tmAscent As Long
        tmDescent As Long
        tmInternalLeading As Long
        tmExternalLeading As Long
        tmAveCharWidth As Long
        tmMaxCharWidth As Long
        tmWeight As Long
        tmOverhang As Long
        tmDigitizedAspectX As Long
        tmDigitizedAspectY As Long
        tmFirstChar As Byte
        tmLastChar As Byte
        tmDefaultChar As Byte
        tmBreakChar As Byte
        tmItalic As Byte
        tmUnderlined As Byte
        tmStruckOut As Byte
        tmPitchAndFamily As Byte
        tmCharSet As Byte
        ntmFlags As Long
        ntmSizeEM As Long
        ntmCellHeight As Long
        ntmAveWidth As Long
End Type

#If VBA7 Then
Public Declare PtrSafe Function EnumFontFamilies Lib "gdi32" Alias "EnumFontFamiliesA" (ByVal hDC As Long, ByVal lpszFamily As String, ByVal lpEnumFontFamProc As Long, lParam As Any) As Long
Public Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
    
'�ƻs�H�U�N�X��@�Ҷ���
Public Declare PtrSafe Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare PtrSafe Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

#Else

Public Declare Function EnumFontFamilies Lib "gdi32" Alias "EnumFontFamiliesA" (ByVal hDC As Long, ByVal lpszFamily As String, ByVal lpEnumFontFamProc As Long, lParam As Any) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
    
'�ƻs�H�U�N�X��@�Ҷ���
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

#End If

Public Enum FontType
    ShxFont = 0
    BigFont = 1
    TrueTypeFont = 2
End Enum

Public Function EnumFontFamProc(lpNLF As LOGFONT, lpNTM As NEWTEXTMETRIC, _
    ByVal FontType As Long, fonts As Collection) As Long
    
    Dim FaceName As String
    Dim FullName As String
    FaceName = StrConv(lpNLF.lfFaceName, vbUnicode)
    fonts.Add left$(FaceName, InStr(FaceName, vbNullChar) - 1)
    EnumFontFamProc = 1
End Function

'�۩w�q��Ө�ơG
'����t�Φr��
Public Sub FillListWithFonts(ByVal hWnd As Long, ByRef fonts As Collection)
    Dim hDC As Long
    hDC = GetDC(hWnd)
    EnumFontFamilies hDC, vbNullString, AddressOf EnumFontFamProc, fonts
    ReleaseDC hWnd, hDC
End Sub

Public Function GetShxFont(ByVal bBigFont As Boolean) As Variant
    Dim strFontFileName() As String     ' �Ҧ��r��W�٪��Ʋ�
    Dim strFontPath() As String     ' AutoCAD���r������|
    
    ' ��o�Ҧ�����������|
    strFontPath = Split(ThisDrawing.Application.preferences.Files, " ")
    
    ' �M���Ҧ�����������|
    Dim i As Integer
    Dim bFirst As Boolean       ' �O�_�O�Ĥ@���d��Ӥ��
    Dim strFont As String       ' �r����W��
    Dim strTemp As String       ' Ū���쪺�r���󪺤@��
    Dim intCount As Integer     ' �r��Ʋժ�����
    Dim strFontFile As String   ' �r���󪺦�m
    intCount = -1
    For i = 0 To UBound(strFontPath)
        bFirst = True
        ' �T�O�̫�@�Ӧr�ŬO"\"
        strFontPath(i) = IIf(right(strFontPath(i), 1) = "\", strFontPath(i), strFontPath(i) & "\")
        
        Do
            If bFirst Then
                strFont = dir(strFontPath(i) & "*.shx")
                bFirst = False
            Else
                strFont = dir
            End If
            
            If Len(strFont) <> 0 Then
                ' ���}�r����
                strFontFile = strFontPath(i) & strFont
                Open strFontFile For Input As #1
                Line Input #1, strTemp
                Close #1
                
                ' �P�_�r�骺����
                If bBigFont Then
                    If Mid(strTemp, 12, 7) = "bigfont" Then
                        intCount = intCount + 1
                        ReDim Preserve strFontFileName(intCount)
                        strFontFileName(intCount) = strFont
                    End If
                Else
                    If Mid(strTemp, 12, 7) = "unifont" Or Mid(strTemp, 12, 6) = "shapes" Then
                        intCount = intCount + 1
                        ReDim Preserve strFontFileName(intCount)
                        strFontFileName(intCount) = strFont
                    End If
                End If
            Else
                Exit Do
            End If
        Loop
    Next i
    
    GetShxFont = strFontFileName
End Function
 '�b�{�Ǥ��ե�

Public Function GetWindowsFonts() As Variant
    Dim bBigFont As Boolean
    Dim WindowsDirectory As String, SystemDirectory As String, x As Long
    WindowsDirectory = space(255)       'Windows���w�˥ؿ�
    SystemDirectory = space(255)        '�t�Υؿ��O
    x = GetWindowsDirectory(WindowsDirectory, 255)
    x = GetSystemDirectory(SystemDirectory, 255)
    
    
    Dim strFontFileName() As String     ' �Ҧ��r��W�٪��Ʋ�
    Dim WinFontPath As String     ' Windows ���r������|
    
    ' ��o�Ҧ�����������|
    WinFontPath = left(WindowsDirectory, InStr(WindowsDirectory, chr(0)) - 1)
    WinFontPath = Trim(WinFontPath) & "\Fonts\"
    
    ' �M���Ҧ�����������|
    Dim i As Integer
    Dim bFirst As Boolean       ' �O�_�O�Ĥ@���d��Ӥ��
    Dim strFont As String       ' �r����W��
    Dim strTemp As String       ' Ū���쪺�r���󪺤@��
    Dim intCount As Integer     ' �r��Ʋժ�����
    Dim strFontFile As String   ' �r���󪺦�m
    intCount = -1
    For i = 0 To 0
        bFirst = True
        ' �T�O�̫�@�Ӧr�ŬO"\"
        WinFontPath = IIf(right(WinFontPath, 1) = "\", WinFontPath, WinFontPath & "\")
        
        Do
            If bFirst Then
                strFont = dir(WinFontPath & "*")
                bFirst = False
            Else
                strFont = dir
            End If
            
            If Len(strFont) <> 0 Then
                ' ���}�r����
                strFontFile = WinFontPath & strFont
                
                
                
                ' �P�_�r�骺����
                If InStr(1, UCase(strFontFile), "TTF") <> 0 Then
                        intCount = intCount + 1
                        ReDim Preserve strFontFileName(intCount)
                        strFontFileName(intCount) = strFontFile
                        
                ElseIf InStr(1, UCase(strFontFile), "TTE") <> 0 Then
                        intCount = intCount + 1
                        ReDim Preserve strFontFileName(intCount)
                        strFontFileName(intCount) = strFontFile
                ElseIf InStr(1, UCase(strFontFile), "TTC") <> 0 Then
                        intCount = intCount + 1
                        ReDim Preserve strFontFileName(intCount)
                        strFontFileName(intCount) = strFontFile
                End If
            Else
                Exit Do
            End If
        Loop
    Next i
    
    GetWindowsFonts = strFontFileName
    
    
End Function

Public Function CreateTextStyle(ByVal fontName As String, ByVal styleName As String, ByVal font As FontType) As AcadTextStyle
    Dim objTextStyle As AcadTextStyle
    Set objTextStyle = ThisDrawing.TextStyles.Add(styleName)
    
    Dim WindowsDirectory As String, SystemDirectory As String, x As Long
    Dim fontTemp As String
    
    If font = ShxFont Then
        objTextStyle.fontFile = fontName
    ElseIf font = BigFont Then
    
        WindowsDirectory = space(255)       'Windows���w�˥ؿ�
        x = GetWindowsDirectory(WindowsDirectory, 255)
        
        
        Dim strFontFileName() As String     ' �Ҧ��r��W�٪��Ʋ�
        Dim WinFontPath As String     ' Windows ���r������|
        
        ' ��o�Ҧ�����������|
        WinFontPath = left(WindowsDirectory, InStr(WindowsDirectory, chr(0)) - 1)
        WinFontPath = Trim(WinFontPath) & "\Fonts\"
        ' �P�_�r�骺����
        If InStr(1, UCase(fontName), "TTF") <> 0 Then
            fontTemp = """" & WinFontPath & fontName & """"
            objTextStyle.BigFontFile = fontTemp
            
        ElseIf InStr(1, UCase(fontName), "TTE") <> 0 Then
            fontTemp = """" & WinFontPath & fontName & """"
            objTextStyle.BigFontFile = fontTemp
        End If
    ElseIf font = TrueTypeFont Then
        ' ��o��e�r�骺�˦�
        Dim typeFace As String
        Dim Bold As Boolean
        Dim Italic As Boolean
        Dim charSet As Long
        Dim PitchandFamily As Long
        ThisDrawing.ActiveTextStyle.GetFont typeFace, Bold, Italic, charSet, PitchandFamily
        
        objTextStyle.SetFont fontName, Bold, Italic, charSet, PitchandFamily
    End If
    
    objTextStyle.width = 1#
    objTextStyle.ObliqueAngle = 0#
    
    Set CreateTextStyle = objTextStyle
End Function



