Attribute VB_Name = "CommonFileDialog"


'CommonFileDialog 模組內容
'Description:
'
'The FileDialog class allows you to use the common dialogs "Open" and "Save" from VBA. Properties for the class include
'Multiselect, title, filter, initial directory, and hide read only.
'Also included is simple methods for finding the window to make the parent of the dialog (the one that it will center on).
'Snippet Code: - Common File Dialog Class

Option Explicit
#If VBA7 Then
'//The Win32 API Functions///
Private Declare PtrSafe Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#Else
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#End If


'//A few of the available Flags///
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_HIDEREADONLY = &H4 '隱蔽只讀複選框



'This one keeps your dialog from turning into
'A browse by folder dialog if multiselect is true!
'Not sure what I mean? Remove it from the flags
'In the "Show Open" & "Show Save" methods.
Private Const OFN_EXPLORER As Long = &H80000
'//The Structure
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long '擁有對話框的窗口
    hInstance As Long
    lpstrFilter As String '裝載文件過濾器的緩衝區
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String    '對話框的標題
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private lngHwnd As Long
Private strFilter As String
Private strTitle As String
Private strDir As String
Private blnHideReadOnly As Boolean
Private blnAllowMulti As Boolean
Private blnMustExist As Boolean
Private Sub Class_Initialize()
    'Set default values when
    'class is first created
    strDir = CurDir
    strTitle = "Llamas Rule"
    strFilter = "All Files" _
    & Chr$(0) & "*.*" & Chr$(0)
    lngHwnd = FindWindow(vbNullString, Application.Caption)
    'None of the flags are set here!
End Sub
Public Function FindUserForm(objForm As UserForm) As Long
    Dim lngTemp As Long
    Dim strCaption As String
    strCaption = objForm.Caption
    lngTemp = FindWindow(vbNullString, strCaption)
    If lngTemp <> 0 Then
        FindUserForm = lngTemp
    End If
End Function
Public Property Let OwnerHwnd(WindowHandle As Long)
    '//FOR YOU TODO//
    'Use the API to validate this handle
    lngHwnd = WindowHandle
    'This value is set at startup to the handle of the
    'AutoCAD Application window, if you want the owner
    'to be a user form you will need to obtain its
    'Handle by using the "FindUserForm" function in
    'This class.
End Property
Public Property Get OwnerHwnd() As Long
    OwnerHwnd = lngHwnd
End Property
Public Property Let title(Caption As String)
    'don't allow null strings
    If Not Caption = vbNullString Then
        strTitle = Caption
    End If
End Property
Public Property Get title() As String
    title = strTitle
End Property
Public Property Let Filter(ByVal FilterString As String)
    'Filters change the type of files that are
    'displayed in the dialog. I have designed this
    'validation to use the same filter format the
    'Common dialog OCX uses:
    '"All Files (*.*)|*.*"
    Dim intPos As Integer
    Do While InStr(FilterString, "|") > 0
    intPos = InStr(FilterString, "|")
    If intPos > 0 Then
        FilterString = Left$(FilterString, intPos - 1) _
        & Chr$(0) & Right$(FilterString, _
        Len(FilterString) - intPos)
    End If
    Loop
    If Right$(FilterString, 2) <> Chr$(0) & Chr$(0) Then
        FilterString = FilterString & Chr$(0)
    End If
    strFilter = FilterString
End Property
Public Property Get Filter() As String
    'Here we reverse the process and return
    'the Filter in the same format the it was
    'entered
    Dim intPos As Integer
    Dim strTemp As String
    strTemp = strFilter
    Do While InStr(strTemp, Chr$(0)) > 0
    intPos = InStr(strTemp, Chr$(0))
    If intPos > 0 Then
        strTemp = Left$(strTemp, intPos - 1) _
        & "|" & Right$(strTemp, _
        Len(strTemp) - intPos)
    End If
    Loop
    If Right$(strTemp, 1) = "|" Then
        strTemp = Left$(strTemp, Len(strTemp) - 1)
    End If
    Filter = strTemp
End Property
Public Property Let InitialDir(strFolder As String)
    'Sets the directory the dialog displays when called
    If Len(dir(strFolder)) > 0 Then
        strDir = strFolder
        Else
        Err.Raise 514, "FileDialog", "Invalid Initial Directory"
    End If
End Property
Public Property Let HideReadOnly(blnVal As Boolean)
    blnHideReadOnly = blnVal
End Property
Public Property Let MultiSelect(blnVal As Boolean)
    'allow users to select more than one file using
    'The Shift or CTRL keys during selection
    blnAllowMulti = blnVal
End Property
Public Property Let FileMustExist(blnVal As Boolean)
    blnMustExist = blnVal
End Property
'@~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~@
' Display and use the File open dialog
'@~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~@
Public Function ShowOpen() As String
    Dim strTemp As String
    Dim udtStruct As OPENFILENAME
    udtStruct.lStructSize = Len(udtStruct)
    'Use our private variable
    udtStruct.hwndOwner = lngHwnd
    'Use our private variable
    udtStruct.lpstrFilter = strFilter
    udtStruct.lpstrFile = space$(254)
    udtStruct.nMaxFile = 255
    udtStruct.lpstrFileTitle = space$(254)
    udtStruct.nMaxFileTitle = 255
    'Use our private variable
    udtStruct.lpstrInitialDir = strDir
    'Use our private variable
    udtStruct.lpstrTitle = strTitle
    'Ok, here we test our booleans to
    'set the flag
    If blnHideReadOnly And blnAllowMulti And blnMustExist Then
        udtStruct.flags = OFN_HIDEREADONLY Or _
        OFN_ALLOWMULTISELECT Or OFN_EXPLORER Or OFN_FILEMUSTEXIST
        ElseIf blnHideReadOnly And blnAllowMulti Then
        udtStruct.flags = OFN_ALLOWMULTISELECT _
        Or OFN_EXPLORER Or OFN_HIDEREADONLY
        ElseIf blnHideReadOnly And blnMustExist Then
        udtStruct.flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST
        ElseIf blnAllowMulti And blnMustExist Then
        udtStruct.flags = OFN_ALLOWMULTISELECT Or _
        OFN_EXPLORER Or OFN_FILEMUSTEXIST
        ElseIf blnHideReadOnly Then
        udtStruct.flags = OFN_HIDEREADONLY
        ElseIf blnAllowMulti Then
        udtStruct.flags = OFN_ALLOWMULTISELECT _
        Or OFN_EXPLORER
        ElseIf blnMustExist Then
        udtStruct.flags = OFN_FILEMUSTEXIST
    End If
    If GetOpenFileName(udtStruct) Then
        strTemp = (Trim(udtStruct.lpstrFile))
        ShowOpen = Mid(strTemp, 1, Len(strTemp) - 1)
    End If
End Function
'@~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~@
' Display and use the File Save dialog
'@~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~@
Public Function ShowSave() As String
    Dim strTemp As String
    Dim udtStruct As OPENFILENAME
    udtStruct.lStructSize = Len(udtStruct)
    'Use our private variable
    udtStruct.hwndOwner = lngHwnd
    'Use our private variable
    udtStruct.lpstrFilter = strFilter
    udtStruct.lpstrFile = space$(254)
    udtStruct.nMaxFile = 255
    udtStruct.lpstrFileTitle = space$(254)
    udtStruct.nMaxFileTitle = 255
    'Use our private variable
    udtStruct.lpstrInitialDir = strDir
    'Use our private variable
    udtStruct.lpstrTitle = strTitle
    If blnMustExist Then
        udtStruct.flags = OFN_FILEMUSTEXIST
    End If
    If GetSaveFileName(udtStruct) Then
        strTemp = (Trim(udtStruct.lpstrFile))
        ShowSave = Mid(strTemp, 1, Len(strTemp) - 1)
    End If
End Function


Function GetFile(strTitle As String, strFilter As String, Optional strIniDir As String) As String
 
    On Error Resume Next
    Dim lntFile As Integer
    Dim FileName As String
    Dim OFileBox As OPENFILENAME
    With OFileBox
        .lpstrTitle = strTitle '對話框標題
        .lpstrInitialDir = strIniDir '初始目錄
        .lStructSize = Len(OFileBox)
        .hwndOwner = ThisDrawing.hWnd
        .flags = OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY
        .lpstrFile = String$(255, 0)
        .nMaxFile = 255
        .lpstrFileTitle = String$(255, 0)
        .nMaxFileTitle = 255
        .lpstrFilter = strFilter  '過濾器
        .nFilterIndex = 1
    End With
     
    lntFile = GetOpenFileName(OFileBox) '執行打開對話框
    If lntFile <> 0 Then
        FileName = Left(OFileBox.lpstrFile, InStr(OFileBox.lpstrFile, vbNullChar) - 1)
        GetFile = FileName
    Else
        GetFile = ""
    End If
 
End Function
Function saveFile(strTitle As String, strFilter As String, Optional strIniDir As String) As String
 
    On Error Resume Next
    Dim lntFile As Integer
    Dim FileName As String
    Dim OFileBox As OPENFILENAME
    With OFileBox
        .lpstrTitle = strTitle '對話框標題
        .lpstrInitialDir = strIniDir '初始目錄
        .lStructSize = Len(OFileBox)
        .hwndOwner = ThisDrawing.hWnd
        .flags = OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY
        .lpstrFile = String$(255, 0)
        .nMaxFile = 255
        .lpstrFileTitle = String$(255, 0)
        .nMaxFileTitle = 255
        .lpstrFilter = strFilter  '過濾器
        .nFilterIndex = 1
    End With
     
    lntFile = GetSaveFileName(OFileBox) '執行打開對話框
    If lntFile <> 0 Then
        FileName = Left(OFileBox.lpstrFile, InStr(OFileBox.lpstrFile, vbNullChar) - 1)
        saveFile = FileName
    Else
        saveFile = ""
    End If
 
End Function
