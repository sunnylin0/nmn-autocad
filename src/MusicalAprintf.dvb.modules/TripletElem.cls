VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TripletElem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public typs As String     '"TripletElem"
Public anchor1 As RelativeElement
Public anchor2 As RelativeElement
Public Number As Integer
Public durationClass As String
Public middleElems As RelativeElementList
Public flatBeams As Boolean
Public endingHeightAbove As Double
Public Sub init(Number As Double, anchor1 As RelativeElement, options As Dictionary)
        Me.typs = "TripletElem"
        Set Me.anchor1 = anchor1  '' must have a .x and a .parent property or be null (means starts at the "beginning" of the line - after key signature)
        Me.Number = Number
        Me.durationClass = ("d" + (Math.Round(anchor1.parent.durationClass * 1000) / 1000))
        Me.durationClass = Replace(Me.durationClass, ".", "-")
        Set Me.middleElems = New RelativeElementList  '' This is to calculate the highest interior pitch. It is used to make sure that the drawn bracket never crosses a really high middle note.
        Me.flatBeams = options("flatBeams")
End Sub

Public Function isClosed() As Boolean
        isClosed = IIf(Not (Me.anchor2 Is Nothing), True, False)
End Function

Public Sub middleNote(elem As RelativeElement)
        Me.middleElems.Push elem
End Sub

Public Sub setCloseAnchor(anchor2 As RelativeElement)
        Set Me.anchor2 = anchor2
        '' TODO-PER: This used to be just for beamed triplets but it looks like bracketed triplets need extra room, too. The only one that doesn't is stem down and beamed
        ''if (me.anchor1.parent.beam)
        If (Not Me.anchor1.parent.Beam Or Me.anchor1.stemdir = "up") Then
            Me.endingHeightAbove = 4
        End If
End Sub

Public Function Clone() As TripletElem
    Dim cope As New TripletElem
    
    cope.typs = Me.typs
Set cope.anchor1 = Me.anchor1
Set cope.anchor2 = Me.anchor2
    cope.Number = Me.Number
    cope.durationClass = Me.durationClass
Set cope.middleElems = Me.middleElems
    cope.flatBeams = Me.flatBeams
    cope.endingHeightAbove = Me.endingHeightAbove

End Function
