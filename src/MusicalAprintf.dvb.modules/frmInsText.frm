VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInsText 
   Caption         =   "���k��J"
   ClientHeight    =   3255
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   2568
   OleObjectBlob   =   "frmInsText.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmInsText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    ' This example prompts for user input of a point. By using the
    ' InitializeUserInput method to define a keyword list, the example can also
    ' return keywords entered by the user.
    
    
    'Me.Hide
    ' Define the valid keywords
    Dim keywordList As String
    keywordList = "1 2 3 4 0"
    Dim strT As String
    
    Dim textObj As AcadText
    Dim textString As String
    Dim insertionPoint(0 To 2) As Double, alignmentPoint(0 To 2) As Double

    Dim height As Double
    Me.Hide
GoToLook:
'On Error GoTo ErrorHandler
On Error Resume Next

    ' Call InitializeUserInput to set up the keywords
    ThisDrawing.Utility.InitializeUserInput 129, keywordList
    ' Get the user input
    Dim returnPnt As Variant
    returnPnt = ThisDrawing.Utility.GetPoint(, "0 1 2 3 4 E(��) V(��) Q(��) W(�~) C(�M��): ")
    
    If Err Then
        '**Description �o�ӭn�諸
         If StrComp(Err.Description, "�ϥΪ̿�J���O����r", 1) = 0 Then
         ' One of the keywords was entered
             Dim inputString As String
             Err.clear
             inputString = ThisDrawing.Utility.GetInput
             '���o����r
             Select Case inputString
                Case "0": strT = "b�Ų�"  '�ũ�
                Case "1": strT = "��"
                Case "2": strT = "��"
                Case "3": strT = "��"
                Case "4": strT = "��"
                Case Else
                    Dim ss As String
                    Dim data As Double
                    ss = Mid(inputString, 1, 1)
                    Select Case ss
                        Case "E", "e": lbPush_d ("��")
                        Case "V", "v": lbPush_d ("��")
                        Case "Q", "q": lbInOut_d ("��")
                        Case "W", "w": lbInOut_d ("�~")
                        Case "T", "t":
                        Case "S", "s":
                            data = Mid(inputString, 2, 10)
                            Me.tbSize = data
                        Case "X", "x":
                            data = Mid(inputString, 2, 10)
                            Me.tbX = data
                        Case "Y", "y":
                            data = Mid(inputString, 2, 10)
                            Me.tbY = data
                        Case "C", "c":
                        '�M������
                            Me.lbInOut = ""
                            Me.lbInOut.SpecialEffect = fmSpecialEffectEtched
                            
                            Me.lbPush = ""
                            Me.lbPush.SpecialEffect = fmSpecialEffectEtched
                            
                            Me.tbT.text = ""
                            strT = ""
                        Case Else
                            Exit Sub
                    End Select
             End Select
             
             Me.tbT = strT
             
             GoTo GoToLook
         Else
             MsgBox "Error selecting the point: " & Err.Description
             'Err.clear
         End If
    Else
        ' Display point coordinates
        'MsgBox "The WCS of the point is: " & returnPnt(0) & ", " & returnPnt(1) & ", " & returnPnt(2), , "GetInput Example"
         
        ' Define the text object
        insertionPoint(0) = returnPnt(0) + Me.tbX
        insertionPoint(1) = returnPnt(1) + Me.tbY
        insertionPoint(2) = returnPnt(2)
        
        alignmentPoint(0) = insertionPoint(0)
        alignmentPoint(1) = insertionPoint(1)
        alignmentPoint(2) = insertionPoint(2)
        
        Dim ipos As Integer
        Dim yAdd As Double
        ipos = 0
        yAdd = 3 '���k���V�W�W�q
        If Me.tbT <> "" Then
            If Mid(Me.tbT, 1, 1) = "b" Then
            '�ݬO�_�O�϶�
                Select Case Me.tbT
                    Case "b�Ų�"
                        'Call ThisDrawing.ModelSpace.InsertBlock(insertionPoint, "�G�J_��", 0.75, 0.75, 0.75, 0)
                        textString = "\U+5B80"
                        height = Me.tbSize
                        Set textObj = ThisDrawing.ModelSpace.AddText(textString, insertionPoint, height)
                        textObj.Alignment = acAlignmentCenter
                        textObj.TextAlignmentPoint = alignmentPoint
                        textObj.styleName = "����_�Ʀr"
                        textObj.Layer = "�˹��Ÿ�"
                        

                    Case Else
                End Select
            Else
            '���O �N���J��r
                textString = Me.tbT
                height = Me.tbSize
                ' Create the text object in model space
                Set textObj = ThisDrawing.ModelSpace.AddText(textString, insertionPoint, height)
                textObj.Alignment = acAlignmentCenter
                textObj.TextAlignmentPoint = alignmentPoint
                textObj.styleName = "����_�Ʀr"
                textObj.Layer = "�˹��Ÿ�"
            End If
            
            ipos = ipos + 1
        End If
        
        If Me.lbInOut.Caption <> "" Then
        
            textString = lbInOut.Caption
            height = Me.tbSize
            
                
            insertionPoint(0) = returnPnt(0) + Me.tbX
            insertionPoint(1) = returnPnt(1) + Me.tbY + (yAdd * ipos)
            insertionPoint(2) = returnPnt(2)
            
            alignmentPoint(0) = insertionPoint(0)
            alignmentPoint(1) = insertionPoint(1)
            alignmentPoint(2) = insertionPoint(2)
            
            
            ' Create the text object in model space
            Set textObj = ThisDrawing.ModelSpace.AddText(textString, insertionPoint, height)
            textObj.Alignment = acAlignmentCenter
            textObj.TextAlignmentPoint = alignmentPoint
            textObj.styleName = "����_�Ʀr"
            textObj.Layer = "�˹��Ÿ�"
            ipos = ipos + 1
        End If
        
        Dim MtextObj As AcadMText
        If Me.lbPush.Caption <> "" Then

'            If Me.lbPush.Caption = "��" Then
'                textString = "��"
'            ElseIf Me.lbPush.Caption = "��" Then
'                textString = "�w"
'            End If
''
'            height = Me.tbSize * 2
'
'            insertionPoint(0) = returnPnt(0) + Me.tbX
'            insertionPoint(1) = returnPnt(1) + Me.tbY + (yAdd * ipos)
'            insertionPoint(2) = returnPnt(2)
'
'            alignmentPoint(0) = insertionPoint(0)
'            alignmentPoint(1) = insertionPoint(1)
'            alignmentPoint(2) = insertionPoint(2)
'
'            Set textObj = ThisDrawing.ModelSpace.AddText(textString, insertionPoint, height)
'            textObj.StyleName = "MMP2005"
'            textObj.Layer = "MMP2005�Ÿ�"
'            textObj.Alignment = acAlignmentCenter
'            textObj.TextAlignmentPoint = alignmentPoint
'


            If Me.lbPush.Caption = "��" Then
                textString = "\U+020C"
            ElseIf Me.lbPush.Caption = "��" Then
                textString = "\U+020A"
            End If
            
            height = Me.tbSize * 3.6
            
                
            insertionPoint(0) = returnPnt(0) + Me.tbX - 0.31
            insertionPoint(1) = returnPnt(1) + Me.tbY + (yAdd * ipos) - (Me.tbSize * 2.8)
            insertionPoint(2) = returnPnt(2)
            
            alignmentPoint(0) = insertionPoint(0)
            alignmentPoint(1) = insertionPoint(1)
            alignmentPoint(2) = insertionPoint(2)
            
            Set textObj = ThisDrawing.ModelSpace.AddText(textString, insertionPoint, height)
            textObj.styleName = "SimpErhu"
            textObj.Layer = "SimpErhu�Ÿ�"
            textObj.Alignment = acAlignmentCenter
            textObj.TextAlignmentPoint = alignmentPoint
            
            ipos = ipos + 1
        End If
        
        
        
        'Err.clear
        GoTo GoToLook
    End If
ErrorHandler:

    
End Sub


Private Sub CommandButton2_Click()
'�m�� ���Ԥ�r�A�אּ �϶�
On Error Resume Next
    Dim Object As Object
    Dim PickedPoint As Variant, TransMatrix As Variant, ContextData As Variant
    
        
TRYAGAIN:
        Me.Hide
While 1
    ' Get information about selected object
    ThisDrawing.Utility.GetSubEntity Object, PickedPoint, TransMatrix, ContextData


    Dim objText As AcadText
    Dim objBlock As AcadBlockReference
    Dim pt As Variant
    Dim SStr As String
    Dim insPt(2) As Double
        If (Object.ObjectName = "AcDbText") Then
            Set objText = Object
            
            If objText.textString = "\U+020A" Then
                insPt(0) = objText.insertionPoint(0) + 1.51
                insPt(1) = objText.insertionPoint(1) + 4.7
                insPt(2) = objText.insertionPoint(2)
                Set objBlock = ThisDrawing.ModelSpace.InsertBlock(insPt, "b��", 4, 4, 4, 0)
                objBlock.Layer = "�˹��Ÿ�"
                objText.Delete
            ElseIf objText.textString = "\U+020C" Then
                insPt(0) = objText.insertionPoint(0) + 1.51
                insPt(1) = objText.insertionPoint(1) + 4.65
                insPt(2) = objText.insertionPoint(2)
                Set objBlock = ThisDrawing.ModelSpace.InsertBlock(insPt, "b��", 4, 4, 4, 0)
                objBlock.Layer = "�˹��Ÿ�"
                objText.Delete
            End If
        Else
            Exit Sub
        End If
Wend
    
End Sub

Private Sub lbPush_Click()
    
    
    If Me.lbPush.Caption = "��" Then
        Me.lbPush.Caption = "��"
        Me.lbPush.SpecialEffect = fmSpecialEffectFlat
    ElseIf Me.lbPush.Caption = "��" Then
        Me.lbPush.Caption = ""
        Me.lbPush.SpecialEffect = fmSpecialEffectEtched
    Else
        Me.lbPush.Caption = "��"
        Me.lbPush.SpecialEffect = fmSpecialEffectFlat
    End If
End Sub
Private Sub lbPush_d(data As String)
    If data = "��" Then
        Me.lbPush.Caption = "��"
        Me.lbPush.SpecialEffect = fmSpecialEffectFlat
    ElseIf data = "��" Then
        Me.lbPush.Caption = "��"
        Me.lbPush.SpecialEffect = fmSpecialEffectFlat
    End If
End Sub
Private Sub lbInOut_Click()
    If Me.lbInOut.Caption = "�~" Then
        Me.lbInOut.Caption = "��"
        Me.lbInOut.SpecialEffect = fmSpecialEffectFlat
    ElseIf Me.lbInOut.Caption = "��" Then
        Me.lbInOut.Caption = ""
        Me.lbInOut.SpecialEffect = fmSpecialEffectEtched
    Else
        Me.lbInOut.Caption = "�~"
        Me.lbInOut.SpecialEffect = fmSpecialEffectFlat
    End If
End Sub

Private Sub lbInOut_d(data As String)
    If data = "��" Then
        Me.lbInOut.Caption = "��"
        Me.lbInOut.SpecialEffect = fmSpecialEffectFlat
    ElseIf data = "�~" Then
        Me.lbInOut.Caption = "�~"
        Me.lbInOut.SpecialEffect = fmSpecialEffectFlat
    End If
End Sub

