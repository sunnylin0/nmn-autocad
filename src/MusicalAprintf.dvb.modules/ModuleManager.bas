Attribute VB_Name = "ModuleManager"
Option Explicit
Option Private Module

Private Const MY_NAME = "ModuleManager"
Private Const MY_NAMENOT = "ModuleManagerNOT"
Private Const ERR_SUPPORTED_APPS = MY_NAME & " currently only supports Microsoft Word and Excel."

#Const MANAGING_WORD = 0
#Const MANAGING_EXCEL = 0
#Const MANAGING_POWERPOINT = 0
#Const MANAGING_ACCESS = 0
#Const MANAGING_OUTLOOK = 0
#Const MANAGING_AUTOCAD = 1


Dim allComponents As VBComponents
Dim fileSys As Object
Dim alreadySaved As Boolean

Public Sub ImportModules(FromDirectory As String, Optional ShowMsgBox As Boolean = True)
    Set fileSys = CreateObject("Scripting.FileSystemObject")
    'Cache some references
    'If the given directory does not exist then show an error dialog and exit
    Dim fromPath As String: fromPath = FromDirectory
    If Not fileSys.FolderExists(fromPath) Then
        fromPath = getFilePath() & "\" & fromPath
        If Not fileSys.FolderExists(fromPath) Then
            MsgBox "Could not locate import directory:  " & FromDirectory
            Exit Sub
        End If
    End If
    Dim dir As Object:     Set dir = fileSys.GetFolder(fromPath)
                
    'Import all VB code files from the given directory if any)
    Dim f As Object
    Dim arrThisMain As New Collection
    Dim imports As Object 'Must be qualified to distinguish it from an MS Word Dictionary
    Set imports = CreateObject("Scripting.Dictionary")
    Dim numFiles As Integer: numFiles = 0
    For Each f In dir.Files
        Dim dotIndex As String: dotIndex = InStrRev(f.Name, ".")
        Dim ext As String: ext = UCase(Right(f.Name, Len(f.Name) - dotIndex))
        Dim correctType As Boolean: correctType = (ext = "BAS" Or ext = "CLS" Or ext = "FRM")
        Dim allowedName As Boolean: allowedName = Left(f.Name, InStrRev(f.Name, ".") - 1) <> MY_NAME
        If correctType And allowedName Then
            numFiles = numFiles + 1
            Dim replaced As Boolean: replaced = doImport(f)
            Dim replacedStr As String
            replacedStr = IIf(replaced, " (replaced)", " (new)")
            imports.add f.Name, replacedStr
        End If

        If ext = "DOCCLS" Then
            arrThisMain.add f
        End If
    Next f
    '注： ThisMain 要最後載入
    Dim thisDoc As Object
    For Each thisDoc In arrThisMain
        Dim fc As Object
        Set fc = thisDoc
        numFiles = numFiles + 1
        replaced = InsertThisMainCode(fc)
        replacedStr = IIf(replaced, " (replaced)", " (Failed)")
        imports.add thisDoc.Name, replacedStr
    Next thisDoc

    'Show a success message box, if requested
    If ShowMsgBox Then
        Dim i As Integer
        Dim msg As String: msg = numFiles & " modules imported:" & vbCrLf & vbCrLf
        For i = 0 To imports.count - 1
            msg = msg & "    " & imports.Keys()(i) & imports.Items()(i) & vbCrLf
        Next i
        Dim result As VbMsgBoxResult: result = MsgBox(msg, vbOKOnly)
    End If
End Sub
Public Sub ExportModules(ToDirectory As String)
    Set fileSys = CreateObject("Scripting.FileSystemObject")
    'Cache some references
    'If the given directory does not exist then show an error dialog and exit
    Dim toPath As String: toPath = ToDirectory
    If Not fileSys.FolderExists(toPath) Then

#If MANAGING_AUTOCAD Then
            toPath = ToDirectory
            fileSys.CreateFolder (toPath)
#Else
        toPath = getFilePath() & "\" & toPath
        If Not fileSys.FolderExists(toPath) Then _
            fileSys.CreateFolder (toPath)
#End If
    End If
    Dim dir As Object
    Set dir = fileSys.GetFolder(toPath)
    
    'Export all modules from this file (except default MS Office modules)
    Dim vbc As VBComponent
    Dim allComponents As VBComponents
    Set allComponents = getAllComponents()
    For Each vbc In allComponents
        Dim correctType As Boolean
        correctType = (vbc.Type = vbext_ct_StdModule Or vbc.Type = vbext_ct_ClassModule Or vbc.Type = vbext_ct_MSForm Or vbc.Type = vbext_ct_Document)
        If correctType And vbc.Name <> MY_NAMENOT Then
            Call doExport(vbc, dir.Path)
        End If
    Next vbc
End Sub
Public Sub RemoveModules(Optional ShowMsgBox As Boolean = True)
    'Check the saved flag to prevent a save event loop
    If alreadySaved Then
        alreadySaved = False
        Exit Sub
    End If

    'Remove all modules from this file (except default MS Office modules obviously)
    Dim removals As New Collection
    Dim vbc As VBComponent
    Dim numModules As Integer: numModules = 0
    Dim allComponents As VBComponents:     Set allComponents = getAllComponents()
    For Each vbc In allComponents
        Dim correctType As Boolean: correctType = (vbc.Type = vbext_ct_StdModule Or vbc.Type = vbext_ct_ClassModule Or vbc.Type = vbext_ct_MSForm)
        If correctType And vbc.Name <> MY_NAME Then
            numModules = numModules + 1
            removals.add vbc.Name
            allComponents.Remove vbc
        End If
    Next vbc

    'Set the saved flag to prevent a save event loop
    'Save file again now that all modules have been removed
    alreadySaved = True
    Call saveFile

    'Show a success message box
    If ShowMsgBox Then
        Dim item As Variant
        Dim msg As String: msg = numModules & " modules successfully removed:" & vbCrLf & vbCrLf
        For Each item In removals
            msg = msg & "    " & item & vbCrLf
        Next item
        msg = msg & vbCrLf & "Don't forget to remove any empty lines after the Attribute lines in .frm files..." _
                  & vbCrLf & "ModuleManager will never be re-imported or exported.  You must do this manually if desired." _
                  & vbCrLf & "NEVER edit code in the VBE and a separate editor at the same time!"
        Dim result As VbMsgBoxResult: result = MsgBox(msg, vbOKOnly)
    End If
End Sub

Private Function getFilePath() As String
#If MANAGING_WORD Then
    getFilePath = ThisDocument.Path
#ElseIf MANAGING_EXCEL Then
    getFilePath = ThisWorkbook.Path
#ElseIf MANAGING_AUTOCAD Then
        getFilePath = ThisDrawing.Path
#Else
        Call raiseUnsupportedAppError
#End If
End Function

Private Function getAllComponents() As VBComponents
    #If MANAGING_WORD Then
        Set getAllComponents = ThisDocument.VBProject.VBComponents
    #ElseIf MANAGING_EXCEL Then
        Set getAllComponents = ThisWorkbook.VBProject.VBComponents
    #ElseIf MANAGING_AUTOCAD Then
        Set getAllComponents = ThisDrawing.Application.vbe.ActiveVBProject.VBComponents
    #Else
        Call raiseUnsupportedAppError
    #End If
End Function



Public Sub Export()
    '您可以使用此程式碼匯出所有文件:
    Dim vbe As vbe
    Set vbe = ThisDrawing.Application.vbe
    Dim comp As VBComponent
    Dim outDir As String
    outDir = "C:\\Temp\\VbaOutput"
    If dir(outDir, vbDirectory) = "" Then
        MkDir outDir
    End If
    For Each comp In vbe.ActiveVBProject.VBComponents
        Select Case comp.Type
            Case vbext_ct_StdModule
                comp.Export outDir & "\" & comp.Name & ".bas"
            Case vbext_ct_Document, vbext_ct_ClassModule
                comp.Export outDir & "\" & comp.Name & ".cls"
            Case vbext_ct_MSForm
                comp.Export outDir & "\" & comp.Name & ".frm"
            Case Else
                comp.Export outDir & "\" & comp.Name
        End Select
    Next comp

    MsgBox "VBA files were exported to : " & outDir
End Sub


Private Sub saveFile()
#If MANAGING_WORD Then
    ThisDocument.save
#ElseIf MANAGING_EXCEL Then
    ThisWorkbook.save
#ElseIf MANAGING_AUTOCAD Then
        ThisDrawing.save
#Else
        Call raiseUnsupportedAppError
#End If
End Sub

Private Sub raiseUnsupportedAppError()
    Err.Raise Number:=vbObjectError + 1, Description:=ERR_SUPPORTED_APPS
End Sub

Private Function doImport(ByRef codeFile As Object) As Boolean
    On Error Resume Next

    'Determine whether a module with this name already exists
    Dim Name As String: Name = Left(codeFile.Name, Len(codeFile.Name) - 4)
    Dim allComponents As VBComponents:   Set allComponents = getAllComponents()
    Dim m As VBComponent:     Set m = allComponents.item(Name)
    If Err.Number <> 0 Then
        Set m = Nothing
    End If
    On Error GoTo 0

    'If so, remove it
    Dim alreadyExists As Boolean: alreadyExists = Not (m Is Nothing)
    If alreadyExists Then
        allComponents.Remove m
    End If
    'Then import the new module
    allComponents.Import (codeFile.Path)
    doImport = alreadyExists
End Function
Private Function InsertThisMainCode(ByRef codeFile As Object) As Boolean
    On Error Resume Next

    'Determine whether a module with this name already exists
    Dim Name As String: Name = Left(codeFile.Name, Len(codeFile.Name) - 7)
    Dim allComponents As VBComponents:     Set allComponents = getAllComponents()
    Dim ThisMainModule As VBComponent:     Set ThisMainModule = allComponents.item(Name)
    If Err.Number <> 0 Then
        Set ThisMainModule = Nothing
        InsertThisMainCode = False
        Exit Function
    End If

    'If so, remove it
    Dim alreadyExists As Boolean: alreadyExists = fileSys.FileExists(codeFile.Path)

    'If so, remove it (even if its ReadOnly)
    If alreadyExists Then

        Dim LnAllCode As String
        Dim stream As Object
        Set stream = fileSys.OpenTextFile(codeFile.Path, 1) '1 means ForReading
        Do Until stream.AtEndOfStream
            LnAllCode = LnAllCode + vbCrLf + stream.ReadLine
        Loop
        If ThisMainModule.CodeModule.CountOfLines > 1 Then
            ThisMainModule.CodeModule.DeleteLines 1, ThisMainModule.CodeModule.CountOfLines - 1
        End If
        ThisMainModule.CodeModule.InsertLines 1, LnAllCode
        ThisMainModule.CodeModule.DeleteLines 1, 1 '刪除第一行空白
    End If

    InsertThisMainCode = alreadyExists
End Function
Private Function doExport(ByRef module As VBComponent, dirPath As String) As Boolean
    'Determine whether a file with this component's name already exists
    Dim ext As String
    Select Case module.Type
        Case vbext_ct_MSForm
            ext = "frm"
        Case vbext_ct_ClassModule
            ext = "cls"
        Case vbext_ct_StdModule
            ext = "bas"
        Case vbext_ct_Document
            ext = "doccls"
    End Select
    Dim filePath As String: filePath = dirPath & "\" & module.Name & "." & ext
    Dim alreadyExists As Boolean: alreadyExists = fileSys.FileExists(filePath)

    'If so, remove it (even if its ReadOnly)
    If alreadyExists Then
        Dim f As Object:     Set f = fileSys.GetFile(filePath)
        If (f.Attributes And 1) Then _
            f.Attributes = f.Attributes - 1 'The bitmask for ReadOnly file attribute
        fileSys.DeleteFile (filePath)
    End If
    If module.Type = vbext_ct_Document Then
        ' ThisWork 匯出是不一樣的
        doExportThisMainCode module, filePath
    Else
        'Then export the module
        'Remove it also, so that the file stays small (and unchanged according to version control)
        module.Export (filePath)
    End If
    doExport = alreadyExists
End Function
Private Function doExportThisMainCode(ByRef module As VBComponent, filePath As String) As Boolean

    'Dim alreadyExists As Boolean: alreadyExists = fileSys.FileExists(filePath)
    Dim stream As Object
    Set stream = fileSys.OpenTextFile(filePath, 2, True) '如果不存在test1.xls將自動建立。
    If Not (stream Is Nothing) Then

        Dim LnAllCode As String


        If module.CodeModule.CountOfLines > 1 Then
            LnAllCode = module.CodeModule.Lines(1, module.CodeModule.CountOfLines - 1)
        Else
            LnAllCode = module.CodeModule.Lines(1, 1)
        End If

        stream.WriteLine LnAllCode

    End If

End Function



'Private Sub Workbook_Open()
'    'Provide the path to a directory with VBA code files. Paths may be relative or absolute.
'    'You can add additional ImportModules() statements to import from multiple locations.
'    'The boolean argument specifies whether or not to show a Message Box on completion.
'    'Remove or comment out these statements when you are ready to provide this workbook to end users,
'    'so that they don't get confused by message boxes about import errors.
'    Call ImportModules(ThisWorkbook.Name & ".modules", ShowMsgBox:=True)
'End Sub
'
'Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, ByRef Cancel As Boolean)
'    'Provide the path to a directory where you want to export modules. Paths may be relative or absolute.
'    'You can add additional ExportModules() statements to export to multiple locations.
'    'Remove or comment out these statements when you are ready to provide this workbook to end users,
'    'so that they don't get confused by the appearance of a bunch of mysterious code files upon saving!
'    Call ExportModules(ThisWorkbook.Name & ".modules")
'End Sub

'Private Sub Workbook_BeforeClose__xxx(ByRef Cancel As Boolean)
'The boolean argument specifies whether or not to show a Message Box on completion.
'Remove or comment out this statement when you are ready to provide this workbook to end users,
'so that modules are not removed, and functionality is not broken the next time the workbook is opened.
'    Call RemoveModules(ShowMsgBox:=True)
'End Sub


Public Sub ExportExcelCode()
    'Excel Vba Code 匯出
    Call ExportModules(ThisWorkbook.Name & ".modules")

End Sub

Public Sub ImportExcelCode()
    'Excel Vba Code 匯入
    Call ImportModules(ThisWorkbook.Name & ".modules", ShowMsgBox:=True)

End Sub


Public Sub ThisImportModules()
    'AutoCAD Vba Code 匯入
    Dim objVBE As vbe
    Set objVBE = ThisDrawing.Application.vbe
    'MsgBox objVBE.ActiveVBProject.FileName
    Call ImportModules(objVBE.ActiveVBProject.FileName & ".modules", ShowMsgBox:=True)

End Sub
Public Sub ThisExportModules()
    'AutoCAD Vba Code 匯出
    Dim objVBE As vbe
    Set objVBE = ThisDrawing.Application.vbe
    'MsgBox objVBE.ActiveVBProject.Name

    Call ExportModules(objVBE.ActiveVBProject.FileName & ".modules")
End Sub






