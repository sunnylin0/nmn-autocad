Attribute VB_Name = "Module1"
Option Explicit

#If VBA7 Then
Public Declare PtrSafe Function GetActiveWindow Lib "user32" () As Long '��o���ʵ������y�`
Declare PtrSafe Function getPT Lib "CaiqsVBApinvoke.arx" Alias "getpt" (ByRef x As Double, ByRef Y As Double, ByRef z As Double) As Integer

'�bWindows�t�Τ��[�J�@�ئr�θ귽�C
Public Declare PtrSafe Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long
Public Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const HWND_BROADCAST = &HFFFF&
Public Const WM_FONTCHANGE = &H1D

#Else

Public Declare Function GetActiveWindow Lib "user32" () As Long '��o���ʵ������y�`
Declare Function getPT Lib "CaiqsVBApinvoke.arx" Alias "getpt" (ByRef x As Double, ByRef y As Double, ByRef Z As Double) As Integer

'�bWindows�t�Τ��[�J�@�ئr�θ귽�C
Public Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const HWND_BROADCAST = &HFFFF&
Public Const WM_FONTCHANGE = &H1D

#End If
