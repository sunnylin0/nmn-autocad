Attribute VB_Name = "ToolsSet"
Option Explicit

Sub SetMusicTextToolsBar()
    ' This example uses MenuGroups to obtain a reference to the AutoCAD main menu.
    ' It then creates a new Toolbar (TestMenu) and inserts a ToolBarButton
    ' with a custom icon. The menu is automatically shown.
    '
    ' * NOTE: The paths of the icons for the new toolbar should be updated
    ' before running this example.
        
    Dim currMenuGroup As AcadMenuGroup
    Dim newToolBar As AcadToolbar, newToolBarButton As AcadToolbarItem
    Dim loadMusic As String
    Dim unloadMusic As String
    Dim loadMusicdvb As String
    Dim ShowMusic As String
    Dim ShowMusicEdit As String
    Dim ShowABCEdit As String
    Dim SetWindowSaveWmf As String
    Dim ToBWMFWmf As String
    Dim SmallBitmapName  As String, LargeBitmapName  As String
    
    On Error GoTo ERRORTRAP
    
    ' Use MenuGroups property to obtain reference to main AutoCAD menu
    Set currMenuGroup = ThisDrawing.Application.MenuGroups.item("ACAD")
    
    If IsFindArrayName("MenuToolsBar", currMenuGroup.Toolbars) > 0 Then
            Set newToolBar = currMenuGroup.Toolbars.item("MenuToolsBar")
    Else
        ' Create the new Toolbar in this group
        Set newToolBar = currMenuGroup.Toolbars.Add("MenuToolsBar")
    End If

    ShowMusicEdit = Chr(3) & Chr(3) & Chr(95) & "-vbarun frmMusicEdit" & vbCr
    Set newToolBarButton = newToolBar.AddToolbarButton(newToolBar.Count + 1, "開起 MusicEdit 介面", "開起 MusicEdit 介面", ShowMusicEdit, False)
    

    ShowABCEdit = Chr(3) & Chr(3) & Chr(95) & "-vbarun Baidu_appition" & vbCr
    Set newToolBarButton = newToolBar.AddToolbarButton(newToolBar.Count + 1, "ABC 介面 Chrome", "ABC 介面 Chrome", ShowABCEdit, False)
    
   
    ' Read icon paths for this Toolbar button
'    GoSub READPATHS
    
'    ' Change the default icon (smile face) for the new toolbar button
'    SmallBitmapName = "c:\images\16x16.bmp"     ' Use a 16x16 pixel .BMP image
'    LargeBitmapName = "c:\images\32x32.bmp"     ' Use a 32x32 pixel .BMP image
'    newToolBarButton.SetBitmaps SmallBitmapName, LargeBitmapName
'
    ' Read icon paths for this Toolbar button
'    GoSub READPATHS
    
    Exit Sub
    
READPATHS:
    ' Read icon paths for this Toolbar button
    newToolBarButton.GetBitmaps SmallBitmapName, LargeBitmapName
'    MsgBox "The new Toolbar uses the following icon files: " & _
'           vbCrLf & vbCrLf & "Small Bitmap: " & SmallBitmapName & vbCrLf & _
'           "Large Bitmap: " & LargeBitmapName
    Return

ERRORTRAP:
  '  MsgBox "The following error has occurred: " & Err.Description
End Sub

Public Function IsFindArrayName( _
  findName As String, _
  arr As Variant, _
  Optional nthOccurrence As Long = 1 _
  ) As Long

    IsFindArrayName = -1

    Dim i As Long
    For i = 1 To arr.Count - 1
        If arr(i).Name = findName Then
            If nthOccurrence > 1 Then
                nthOccurrence = nthOccurrence - 1
                GoTo continue
            End If
            IsFindArrayName = i
            Exit Function
        End If
continue:
    Next i

End Function


Sub Example_SetTools()
    ' This example uses MenuGroups to obtain a reference to the AutoCAD main menu.
    ' It then creates a new Toolbar (TestMenu) and inserts a ToolBarButton
    ' with a custom icon. The menu is automatically shown.
    '
    ' * NOTE: The paths of the icons for the new toolbar should be updated
    ' before running this example.
        
    Dim currMenuGroup As AcadMenuGroup
    Dim newToolBar As AcadToolbar, newToolBarButton As AcadToolbarItem
    Dim loadMusic As String
    Dim unloadMusic As String
    Dim loadMusicdvb As String
    Dim ShowMusic As String
    Dim ShowMusicEdit As String
    Dim SetWindowSaveWmf As String
    Dim ToBWMFWmf As String
    Dim SmallBitmapName  As String, LargeBitmapName  As String
    
    On Error GoTo ERRORTRAP
    
    ' Use MenuGroups property to obtain reference to main AutoCAD menu
    Set currMenuGroup = ThisDrawing.Application.MenuGroups.item("ACAD")
    
    ' Create the new Toolbar in this group
    Set newToolBar = currMenuGroup.Toolbars.Add("TestMenu")
    
    ' Add an item to the new Toolbar and assign an Open macro
    ' (VBA equivalent of: "ESC ESC _open ")
    loadMusic = Chr(3) & Chr(3) & Chr(95) & "_arx l " & """Z:/download/備份簡譜/Musical/debug/AsdkMusicalDb.dbx""" & Chr(32)

    Set newToolBarButton = newToolBar.AddToolbarButton(newToolBar.Count + 1, "載入 Musical.dbx", "載入 Musical.dbx", loadMusic, False)
   
    unloadMusic = Chr(3) & Chr(3) & Chr(95) & "_arx U asdkmusicaldb.dbx" & vbCr
    Set newToolBarButton = newToolBar.AddToolbarButton(newToolBar.Count + 1, "釋放 Musical.dbx", "釋放 Musical.dbx", unloadMusic, False)
    
    loadMusicdvb = Chr(3) & Chr(3) & Chr(95) & "-vbaload " & """Z:/download/備份簡譜/Musical/debug/MusicalAprintf.dvb""" & Chr(32)
    Set newToolBarButton = newToolBar.AddToolbarButton(newToolBar.Count + 1, "截入 MusicalAprintf.dvb", "截入 MusicalAprintf.dvb", loadMusicdvb, False)
    
    ShowMusic = Chr(3) & Chr(3) & Chr(95) & "-vbarun frmMusic" & vbCr
    Set newToolBarButton = newToolBar.AddToolbarButton(newToolBar.Count + 1, "開起 Music 介面", "開起 Music 介面", ShowMusic, False)
    
    ShowMusicEdit = Chr(3) & Chr(3) & Chr(95) & "-vbarun frmMusicEdit" & vbCr
    Set newToolBarButton = newToolBar.AddToolbarButton(newToolBar.Count + 1, "開起 MusicEdit 介面", "開起 MusicEdit 介面", ShowMusicEdit, False)
    
    SetWindowSaveWmf = Chr(3) & Chr(3) & Chr(95) & "-vbarun Ch4_MaximizeApplicationWindow" & vbCr
    Set newToolBarButton = newToolBar.AddToolbarButton(newToolBar.Count + 1, "將當前圖形縮放至兩點定義的窗口", "將當前圖形縮放至兩點定義的窗口", SetWindowSaveWmf, False)
   
    ToBWMFWmf = Chr(3) & Chr(3) & Chr(95) & "bwmfout" & vbCr
    Set newToolBarButton = newToolBar.AddToolbarButton(newToolBar.Count + 1, "BWMF 兩點定義的窗口", "BWMF 兩點定義的窗口", ToBWMFWmf, False)
   
   
   
    ' Read icon paths for this Toolbar button
'    GoSub READPATHS
    
'    ' Change the default icon (smile face) for the new toolbar button
'    SmallBitmapName = "c:\images\16x16.bmp"     ' Use a 16x16 pixel .BMP image
'    LargeBitmapName = "c:\images\32x32.bmp"     ' Use a 32x32 pixel .BMP image
'    newToolBarButton.SetBitmaps SmallBitmapName, LargeBitmapName
'
    ' Read icon paths for this Toolbar button
'    GoSub READPATHS
    
    Exit Sub
    
READPATHS:
    ' Read icon paths for this Toolbar button
    newToolBarButton.GetBitmaps SmallBitmapName, LargeBitmapName
'    MsgBox "The new Toolbar uses the following icon files: " & _
'           vbCrLf & vbCrLf & "Small Bitmap: " & SmallBitmapName & vbCrLf & _
'           "Large Bitmap: " & LargeBitmapName
    Return

ERRORTRAP:
'    MsgBox "The following error has occurred: " & Err.Description
End Sub

Sub Example_exTools()
'匯出工具列
    ' This example uses MenuGroups to obtain a reference to the AutoCAD main menu.
    ' It then creates a new Toolbar (TestMenu) and inserts a ToolBarButton
    ' with a custom icon. The menu is automatically shown.
    '
    ' * NOTE: The paths of the icons for the new toolbar should be updated
    ' before running this example.
        
    Dim currMenuGroup As AcadMenuGroup
    Dim newToolBar As AcadToolbar, newToolBarButton As AcadToolbarItem
    Dim loadMusic As String
    Dim unloadMusic As String
    Dim ShowMusic As String
    Dim ShowMusicEdit As String
    Dim SmallBitmapName  As String, LargeBitmapName  As String
    
    'On Error GoTo ERRORTRAP
    On Error Resume Next
    
    ' Use MenuGroups property to obtain reference to main AutoCAD menu
    Set currMenuGroup = ThisDrawing.Application.MenuGroups.item("ACAD")
    
    ' Create the new Toolbar in this group
    Set newToolBar = currMenuGroup.Toolbars.item("MusicText")
    
    ' Add an item to the new Toolbar and assign an Open macro
    ' (VBA equivalent of: "ESC ESC _open ")
    
    Dim help1 As String
    Dim code As String
    Dim i As Integer
    For i = 1 To newToolBar.Count - 1
        help1 = newToolBar.item(i).Name
        code = newToolBar.item(i).Macro
        ThisDrawing.Utility.Prompt "\n第" & i & "個," & help1 & "," & code
    Next
'
'    loadMusic = Chr(3) & Chr(3) & Chr(95) & "_arx l " & """E:/games/AutoCad_Arx/samples/entity/Musical/debug/AsdkMusicalDb.dbx""" & Chr(32)
'    Set newToolBarButton = newToolBar.AddToolbarButton(newToolBar.count + 1, "載入 Musical.dbx", "載入 Musical.dbx", loadMusic, False)
'
'    'unloadMusic = Chr(3) & Chr(3) & Chr(95) & "_arx U " & "AsdkMusicalDb.dbx "
'    unloadMusic = Chr(3) & Chr(3) & Chr(95) & "_arx U asdkmusicaldb.dbx" & vbCr
'    Set newToolBarButton = newToolBar.AddToolbarButton(newToolBar.count + 1, "釋放 Musical.dbx", "釋放 Musical.dbx", unloadMusic, False)
'
'    ShowMusic = Chr(3) & Chr(3) & Chr(95) & "-vbarun frmMusic" & vbCr
'    Set newToolBarButton = newToolBar.AddToolbarButton(newToolBar.count + 1, "開起 Music 介面", "開起 Music 介面", ShowMusic, False)
'
'    ShowMusicEdit = Chr(3) & Chr(3) & Chr(95) & "-vbarun frmMusicEdit" & vbCr
'    Set newToolBarButton = newToolBar.AddToolbarButton(newToolBar.count + 1, "開起 MusicEdit 介面", "開起 MusicEdit 介面", ShowMusicEdit, False)
'
    ' Read icon paths for this Toolbar button
'    GoSub READPATHS
    
'    ' Change the default icon (smile face) for the new toolbar button
'    SmallBitmapName = "c:\images\16x16.bmp"     ' Use a 16x16 pixel .BMP image
'    LargeBitmapName = "c:\images\32x32.bmp"     ' Use a 32x32 pixel .BMP image
'    newToolBarButton.SetBitmaps SmallBitmapName, LargeBitmapName
'
    ' Read icon paths for this Toolbar button
'    GoSub READPATHS
    
    Exit Sub
    
READPATHS:
    ' Read icon paths for this Toolbar button
    newToolBarButton.GetBitmaps SmallBitmapName, LargeBitmapName
'    MsgBox "The new Toolbar uses the following icon files: " & _
'           vbCrLf & vbCrLf & "Small Bitmap: " & SmallBitmapName & vbCrLf & _
'           "Large Bitmap: " & LargeBitmapName
    Return

ERRORTRAP:
    'MsgBox "The following error has occurred: " & Err.Description
End Sub


Sub Example_SetMusicTextTools()
    ' This example uses MenuGroups to obtain a reference to the AutoCAD main menu.
    ' It then creates a new Toolbar (TestMenu) and inserts a ToolBarButton
    ' with a custom icon. The menu is automatically shown.
    '
    ' * NOTE: The paths of the icons for the new toolbar should be updated
    ' before running this example.
        
    Dim currMenuGroup As AcadMenuGroup
    Dim newToolBar As AcadToolbar, newToolBarButton As AcadToolbarItem
    Dim loadMusic As String
    Dim unloadMusic As String
    Dim ShowMusic As String
    Dim ShowMusicEdit As String
    Dim SmallBitmapName  As String, LargeBitmapName  As String
    
    On Error GoTo ERRORTRAP
    
    ' Use MenuGroups property to obtain reference to main AutoCAD menu
    Set currMenuGroup = ThisDrawing.Application.MenuGroups.item("ACAD")
    
    ' Create the new Toolbar in this group
    Set newToolBar = currMenuGroup.Toolbars.Add("MusicText")
    Dim i As Integer
    Dim Marco(10, 2) As String
    Dim vbaPATH As String
    Dim imagePATH As String
    Dim pos As Integer
Marco(0, 0) = "Unload 釋放AsdkMusicText"
Marco(1, 0) = "Load 載入AsdkMusicText"
Marco(2, 0) = "開起 Edit 介面"
Marco(3, 0) = "加入符號"
Marco(4, 0) = "開起 PageMusic 頁面"
Marco(5, 0) = "連結拍子"
Marco(6, 0) = "移除連結"
Marco(7, 0) = "IsSet"
Marco(8, 0) = "SelectDraw"

Marco(0, 1) = Chr(3) & Chr(3) & Chr(95) & "_arx u AsdkMusicText.arx" & vbCr
Marco(1, 1) = Chr(3) & Chr(3) & Chr(95) & "_arx l " & """AsdkMusicText.arx""" & Chr(32)
Marco(2, 1) = Chr(3) & Chr(3) & Chr(95) & "_musicsample "
Marco(3, 1) = Chr(3) & Chr(3) & Chr(95) & "_addmusictext "
Marco(4, 1) = Chr(3) & Chr(3) & Chr(95) & "_addMusicTextPage "
Marco(5, 1) = Chr(3) & Chr(3) & Chr(95) & "_addMusicJoinMany "
Marco(6, 1) = Chr(3) & Chr(3) & Chr(95) & "_addMusicEraseMany "
Marco(7, 1) = Chr(3) & Chr(3) & Chr(95) & "_IsSet "
Marco(8, 1) = Chr(3) & Chr(3) & Chr(95) & "_-vbarun ThisDrawing.Example_SelectOnScreen "

    vbaPATH = ThisDrawing.Application.vbe.ActiveVBProject.fileName
    pos = InStrRev(vbaPATH, "\")
    imagePATH = Mid(vbaPATH, 1, pos) & "image\"
Marco(0, 2) = imagePATH & "UnLoad.BMP"
Marco(1, 2) = imagePATH & "Load.BMP"
Marco(2, 2) = imagePATH & "SEdit.BMP"
Marco(3, 2) = imagePATH & "AddMusicText.BMP"
Marco(4, 2) = imagePATH & "PageMusic.BMP"
Marco(5, 2) = imagePATH & "JoinMany.BMP"
Marco(6, 2) = imagePATH & "EraseMany.BMP"
Marco(7, 2) = imagePATH & "IsSet.BMP"
Marco(8, 2) = imagePATH & "IsSet.BMP"
    
    For i = 0 To 8
        Set newToolBarButton = newToolBar.AddToolbarButton(newToolBar.Count + 1, Marco(i, 0), Marco(i, 0), Marco(i, 1), False)
        
        newToolBarButton.SetBitmaps Marco(i, 2), Marco(i, 2)
    Next
    ' Add an item to the new Toolbar and assign an Open macro
    ' (VBA equivalent of: "ESC ESC _open ")
    'loadMusic = Chr(3) & Chr(3) & Chr(95) & "_arx l " & """E:/games/AutoCad_Arx/samples/entity/Musical/debug/AsdkMusicalDb.dbx""" & Chr(32)
    'Set newToolBarButton = newToolBar.AddToolbarButton(newToolBar.count + 1, "載入 Musical.dbx", "載入 Musical.dbx", loadMusic, False)
    
'    ' Change the default icon (smile face) for the new toolbar button
    'SmallBitmapName = "c:\images\16x16.bmp"     ' Use a 16x16 pixel .BMP image
    'LargeBitmapName = "c:\images\32x32.bmp"     ' Use a 32x32 pixel .BMP image
    'newToolBarButton.SetBitmaps SmallBitmapName, LargeBitmapName
'
    ' Read icon paths for this Toolbar button
'    GoSub READPATHS
    
    Exit Sub
    
READPATHS:
    ' Read icon paths for this Toolbar button
    newToolBarButton.GetBitmaps SmallBitmapName, LargeBitmapName
'    MsgBox "The new Toolbar uses the following icon files: " & _
'           vbCrLf & vbCrLf & "Small Bitmap: " & SmallBitmapName & vbCrLf & _
'           "Large Bitmap: " & LargeBitmapName
    Return

ERRORTRAP:
    'MsgBox "The following error has occurred: " & Err.Description
End Sub


