VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpecDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' SpecDefinition v1.4.0
' (c) Tim Hall - https:''github.com/timhall/Excel-TDD
'
' Provides helpers and acts as workbook proxy
'
' @dependencies
' @author tim.hall.engr@gmail.com
' @license: MIT (http:''www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Private pExpectations As Collection
Private pFailedExpectations As Collection

Public Enum SpecResult
    Pass
    Fail
    Pending
End Enum


' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' Properties
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Public Description As String
Public ID As String

Public Property Get Expectations() As Collection
    If pExpectations Is Nothing Then
        Set pExpectations = New Collection
    End If
    Set Expectations = pExpectations
End Property
Private Property Let Expectations(value As Collection)
    Set pExpectations = value
End Property

Public Property Get FailedExpectations() As Collection
    If pFailedExpectations Is Nothing Then
        Set pFailedExpectations = New Collection
    End If
    Set FailedExpectations = pFailedExpectations
End Property
Private Property Let FailedExpectations(value As Collection)
    Set pFailedExpectations = value
End Property


' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' Public Methods
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

''
' Create a new expectation to test the defined value
'
' @param {Variant} value Value to be tested by expectation
' @returns {Expectation}
' --------------------------------------------- '

Public Function Expect(Optional value As Variant) As SpecExpectation
    Dim exp As New SpecExpectation
    
    If VarType(value) = vbObject Then
        Set exp.Actual = value
    Else
        exp.Actual = value
    End If
    Me.Expectations.Add exp
    
    Set Expect = exp
End Function

''
' Run each expectation, store failed expectations, and return result
'
' @returns {SpecResult} Pass/Fail/Pending
' --------------------------------------------- '

Public Function result() As SpecResult
    Dim exp As SpecExpectation
    
    ' Reset failed expectations
    FailedExpectations = New Collection
    
    ' If no expectations have been defined, return pending
    If Me.Expectations.Count < 1 Then
        result = Pending
    Else
        ' Loop through all expectations
        For Each exp In Me.Expectations
            ' If expectation fails, store it
            If exp.result = Fail Then
                FailedExpectations.Add exp
            End If
        Next exp
        
        ' If no expectations failed, spec passes
        If Me.FailedExpectations.Count > 0 Then
            result = Fail
        Else
            result = Pass
        End If
    End If
End Function

''
' Helper to get result name (i.e. "Pass", "Fail", "Pending")
'
' @returns {String}
' --------------------------------------------- '

Public Function ResultName() As String
    Select Case Me.result
        Case Pass: ResultName = "Pass"
        Case Fail: ResultName = "Fail"
        Case Pending: ResultName = "Pending"
    End Select
End Function
