Attribute VB_Name = "OpenExcel"
Option Explicit

' Declare necessary API routines:
Private Declare Function FindWindow Lib "user32" Alias _
"FindWindowA" (ByVal lpClassName As String, _
                    ByVal lpWindowName As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias _
"SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
                    ByVal wParam As Long, _
                    ByVal lParam As Long) As Long
Sub GetExcel()
    Dim MyXL As Object    ' Variable to hold reference
                                ' to Microsoft Excel.
    Dim ExcelWasNotRunning As Boolean    ' Flag for final release.
' Test to see if there is a copy of Microsoft Excel already running.
    On Error Resume Next    ' Defer error trapping.
' GetObject function called without the first argument returns a
' reference to an instance of the application. If the application isn't
' running, an error occurs.
    Set MyXL = GetObject(, "Excel.Application")
    If err.Number <> 0 Then ExcelWasNotRunning = True
    err.Clear    ' Clear Err object in case error occurred.
' Check for Microsoft Excel. If Microsoft Excel is running,
' enter it into the Running Object table.
    DetectExcel
' Set the object variable to reference the file you want to see.
    Set MyXL = GetObject("C:\GitHub\muse\app����.xls")
' Show Microsoft Excel through its Application property. Then
' show the actual window containing the file using the Windows
' collection of the MyXL object reference.
    MyXL.Application.Visible = True
    MyXL.parent.Windows(1).Visible = True
     'Do manipulations of your  file here.
    ' ...
' If this copy of Microsoft Excel was not running when you
' started, close it using the Application property's Quit method.
' Note that when you try to quit Microsoft Excel, the
' title bar blinks and a message is displayed asking if you
' want to save any loaded files.
    If ExcelWasNotRunning = True Then
        MyXL.Application.Quit
    End If
    Set MyXL = Nothing    ' Release reference to the
                                ' application and spreadsheet.
End Sub
Sub DetectExcel()
' Procedure dectects a running Excel and registers it.
    Const WM_USER = 1024
    Dim hWnd As Long
' If Excel is running this API call returns its handle.
    hWnd = FindWindow("XLMAIN", 0)
    If hWnd = 0 Then    ' 0 means Excel not running.
        Exit Sub
    Else
    ' Excel is running so use the SendMessage API
    ' function to enter it in the Running Object Table.
        SendMessage hWnd, WM_USER + 18, 0, 0
    End If
End Sub

