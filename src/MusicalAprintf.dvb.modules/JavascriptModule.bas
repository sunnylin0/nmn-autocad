Attribute VB_Name = "JavascriptModule"
Option Explicit
'�i�H�� ������ javascript
Public WD As SeleniumBasic.IWebDriver
Dim abcJSON As String
Sub Baidu_Chrome()
    On Error GoTo Err1
    Dim Service As SeleniumBasic.ChromeDriverService
    Dim options As SeleniumBasic.ChromeOptions
    Set WD = New SeleniumBasic.IWebDriver
    Set Service = New SeleniumBasic.ChromeDriverService
    With Service
        .CreateDefaultService driverPath:="C:\GitHub\SeleniumBasic\Drivers"
        .HideCommandPromptWindow = True
    End With
    Set options = New SeleniumBasic.ChromeOptions
    With options
        .BinaryLocation = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
        .AddExcludedArgument "enable-automation"
        .AddArgument "--start-maximized"
        '.DebuggerAddress = "127.0.0.1:9999" '���n�O��L�L?�V��
    End With
    WD.New_ChromeDriver Service:=Service, options:=options
    'WD.URL = "https:''www.baidu.com"
    WD.url = "file:''/C:/GitHub/abc2svg/zabc/edit-lin.html"
    Dim form As SeleniumBasic.IWebElement
    Dim Keyword As SeleniumBasic.IWebElement
    Dim button As SeleniumBasic.IWebElement
    Set form = WD.FindElementById("form")
    Set Keyword = form.FindElementById("kw")
    Keyword.Clear
    Keyword.SendKeys "�n�ݪ��H"
    Set button = form.FindElementById("su")
    button.Click
    Debug.Print WD.title, WD.url
    Debug.Print WD.PageSource
    'MsgBox "�U���h�X Web�C"
    'WD.Quit
    Exit Sub
Err1:
    MsgBox err.Description, vbCritical
End Sub

Sub aas()
'WD.URL = "file:''/C:/GitHub/abc2svg/zabc/edit-lin.html"
Baidu_appition
Call WD.ExecuteScript("alert('Hello,ryueifu�I')")

End Sub
Sub ttm55()
Dim prod As Variant
Baidu_appition
prod = WD.ExecuteScript("return ttm.ui")
Debug.Print prod
prod = WD.ExecuteScript("return toAcadMusicJSON()")
Debug.Print prod

End Sub
Sub tt123()
Dim prod As Variant
Dim aas As SeleniumBasic.IOptions
prod = WD.ExecuteScript("return tt123")
Debug.Print prod

End Sub

Function getAbcJson()
    Dim prod As Variant
    Baidu_appition
    
    prod = WD.ExecuteScript("return toAcadMusicJSON()")
    Debug.Print prod

End Function
Function testJSON()
    Dim JsonString As String
    Dim JsonObject
    Dim Json As Dictionary
    Set Json = JsonConverter.ParseJson("{""a"":123,""b"":[1,2,3,4],""c"":{""d"":456}}")
    
    ' Json("a") -> 123
    ' Json("b")(2) -> 2
    ' Json("c")("d") -> 456
    Json("c")("e") = 789
    
    Debug.Print JsonConverter.ConvertToJson(Json)
    ' -> "{"a":123,"b":[1,2,3,4],"c":{"d":456,"e":789}}"
    
    Debug.Print JsonConverter.ConvertToJson(Json, Whitespace:=2)
    Debug.Print Json.Count

End Function

Function voieMusic()
    Dim vvbase As New voiceBase
    Debug.Print vvbase.typs
    vvbase.typs = 3
    Debug.Print vvbase.typs
    
    Dim vvbar As New voiceBar
    Debug.Print vvbar.typs
End Function

Sub Baidu_Opera()
    On Error GoTo Err1
    Dim Service As SeleniumBasic.OperaDriverService
    Dim options As SeleniumBasic.OperaOptions
    Set WD = New SeleniumBasic.IWebDriver
    Set Service = New SeleniumBasic.OperaDriverService
    With Service
        .CreateDefaultService driverPath:="C:\GitHub\SeleniumBasic\Drivers", driverexecutablefilename:="chromedriver.exe"
        .HideCommandPromptWindow = True
    End With
    Set options = New SeleniumBasic.OperaOptions
    With options
        '.BinaryLocation = "C:\Users\Administrator\AppData\Local\Programs\Opera\71.0.3770.148\opera.exe"
        .BinaryLocation = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    End With
    WD.New_OperaDriver Service:=Service, options:=options
    WD.Navigate.GoToUrl "https:''www.baidu.com"
    Dim form As SeleniumBasic.IWebElement
    Dim Keyword As SeleniumBasic.IWebElement
    Dim button As SeleniumBasic.IWebElement
    Set form = WD.FindElementById("form")
    Set Keyword = form.FindElementById("kw")
    Keyword.Clear
    Keyword.SendKeys "VBA Selenium"
    Set button = form.FindElementById("su")
    button.Click
    Debug.Print WD.title, WD.url
    Debug.Print WD.PageSource
    MsgBox "�����˳��������"
    WD.Quit
    Exit Sub
Err1:
    MsgBox err.Description, vbCritical
End Sub

Sub Baidu2()
    On Error GoTo Err1
    Set WD = New SeleniumBasic.IWebDriver
    WD.New_ChromeDriver
    WD.url = "https:''www.baidu.com"
    MsgBox "�U���h�X??���C"
    WD.Quit
    Exit Sub
Err1:
    MsgBox err.Description, vbCritical
End Sub


Sub Baidu_appition()
    Dim Service As SeleniumBasic.ChromeDriverService
    Dim options As SeleniumBasic.ChromeOptions
    Set WD = New SeleniumBasic.IWebDriver
    Set Service = New SeleniumBasic.ChromeDriverService
    With Service
        .CreateDefaultService driverPath:="C:\GitHub\SeleniumBasic\Drivers"
        .HideCommandPromptWindow = True
    End With
    Set options = New SeleniumBasic.ChromeOptions
    With options
        .BinaryLocation = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
        '.AddExcludedArgument "enable-automation"
        '.AddArgument "--start-maximized"
        .DebuggerAddress = "127.0.0.1:9999" '���n�O��L�X�ӲV��
    End With

    WD.New_ChromeDriver Service:=Service, options:=options
    'WD.URL = "file:''/C:/GitHub/abc2svg/zabc/edit-lin.html"
End Sub


