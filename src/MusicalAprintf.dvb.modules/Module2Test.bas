Attribute VB_Name = "Module2Test"
Option Explicit

Type sNote
    aty As String * 20
    ajo As Integer
    ajn As Integer
End Type

Type asFormat
    typs As Cg
    notes() As sNote
    dur As Integer
    slur As Integer
End Type
Function typeStest()
    'ด๚ธี type
     Dim exl As Excel.Application
     exl.FindFormat
     
    Dim ff() As asFormat
    Dim f1 As asFormat
    Dim spaceAsF As asFormat
    Dim spaceNote As sNote
    Dim sn() As sNote
    Dim se() As sNote
    Dim str As String
    Dim aiir
    
    Debug.Print IsEmpty(aiir)
    
    ReDim Preserve ff(3)
    f1 = ff(0)
    f1.typs = Cg.note
    f1.dur = 100
    f1.slur = 1
    
    Debug.Print ff(0).typs & "  " & ff(0).dur & "  " & ff(0).slur
    Debug.Print f1.typs & "  " & f1.dur & "  " & f1.slur
    ff(0) = f1
    Debug.Print ff(0).typs & "  " & ff(0).dur & "  " & ff(0).slur
    
    ReDim Preserve sn(2)
    sn(0).aty = "TTYA":    sn(0).ajn = 6:    sn(0).ajo = -4
    sn(1).aty = "aase":    sn(1).ajn = 1:    sn(1).ajo = 5
    
    ff(0).notes = sn
    Debug.Print ff(0).typs & "  " & ff(0).dur & "  " & ff(0).slur
    Debug.Print ff(0).notes(0).aty & "  " & ff(0).notes(0).ajn & "  " & ff(0).notes(0).ajo
    Debug.Print ff(0).notes(1).aty & "  " & ff(0).notes(1).ajn & "  " & ff(0).notes(1).ajo
    Debug.Print ff(0).notes(2).aty & "  " & ff(0).notes(2).ajn & "  " & ff(0).notes(2).ajo
        On Error Resume Next    ' Defer error trapping.
    ReDim Preserve se(0)
    ff(0).notes = se
    Debug.Print ff(0).typs & "  " & ff(0).dur & "  " & ff(0).slur
    If UBound(ff(0).notes) Then
        Debug.Print ff(0).notes(0).aty & "  " & ff(0).notes(0).ajn & "  " & ff(0).notes(0).ajo
    End If
    

    Dim msg
    If err.Number <> 0 Then
        ' Tell user what happened. Then clear the Err object.
        msg = "code: " & err.Number & vbCrLf & err.Description
        MsgBox msg, , "Deferred Error Test"
        err.Clear    ' Clear Err object fields
    End If
    
    Debug.Print ff(0).notes(0).aty & "  " & ff(0).notes(0).ajn & "  " & ff(0).notes(0).ajo
    Debug.Print ff(0).notes(1).aty & "  " & ff(0).notes(1).ajn & "  " & ff(0).notes(1).ajo
    Debug.Print ff(0).notes(2).aty & "  " & ff(0).notes(2).ajn & "  " & ff(0).notes(2).ajo
    
    
'
'    s1.aai = 4
'    arrc(0) = 3
'    arrc(1) = 2
'    arrc(2) = 1
'    s1.aac = arrc
'    MsgBox UBound(s1.aac)
'    ReDim s1.aac(3)
'    s1.aac(0) = 14
'    MsgBox s1.aai
'
'    MsgBox s1.aac(0)
End Function
