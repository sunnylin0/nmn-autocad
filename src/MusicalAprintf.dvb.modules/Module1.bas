Attribute VB_Name = "Module1"
Option Explicit

#If VBA7 Then
Public Declare PtrSafe Function GetActiveWindow Lib "user32" () As Long '獲得活動視窗的句柄
Declare PtrSafe Function getPT Lib "CaiqsVBApinvoke.arx" Alias "getpt" (ByRef x As Double, ByRef Y As Double, ByRef z As Double) As Integer

'在Windows系統中加入一種字形資源。
Public Declare PtrSafe Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long
Public Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const HWND_BROADCAST = &HFFFF&
Public Const WM_FONTCHANGE = &H1D

#Else

Public Declare Function GetActiveWindow Lib "user32" () As Long '獲得活動視窗的句柄
Declare Function getPT Lib "CaiqsVBApinvoke.arx" Alias "getpt" (ByRef x As Double, ByRef y As Double, ByRef Z As Double) As Integer

'在Windows系統中加入一種字形資源。
Public Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const HWND_BROADCAST = &HFFFF&
Public Const WM_FONTCHANGE = &H1D

#End If
