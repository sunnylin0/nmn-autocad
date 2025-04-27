VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TestDll_Print 
   Caption         =   "UserForm1"
   ClientHeight    =   3225
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4704
   OleObjectBlob   =   "TestDll_Print.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TestDll_Print"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CommandButton1_Click()
    Dim inss As INS_Object
    Dim ii As Double
    Dim hWnd As Long
    Set inss = New INS_Object
    ii = inss.AddLine(33, 99, 77, 44)
    hWnd = GetActiveWindow
    MsgBox "VBA is GetActiveWindow  : " & hWnd
    
    inss.aPaint hWnd
    
    
End Sub

Private Sub CommandButton2_Click()
    On Error Resume Next
    Dim oApp As Object  ' AcadApplication
    ' Try to connect to a running instance of AutoCAD.
    
    Set oApp = GetObject(, "AutoCAD.Application")
    If err Then
        ' Failed to get AutoCAD.
        MsgBox ("Could not connect to AutoCAD.")
        err.Clear
        Exit Sub
    End If
     
    
    Dim myCom As myCustomCom
    Dim vvv As Variant
    'myCom = oApp.GetInterfaceObject("comArxProject1Lib.myCustomCom1")
    
    Set myCom = New myCustomCom
    Dim x As Single
    Dim y As Single
    Dim Z As Single
    Call myCom.getPosition(x, y, Z)
    myCom.addMuiscText
    MsgBox x & ":" & y & ":" & Z & ":"
  
End Sub
'Imports asdkcomServerFromArxLib
'Imports Autodesk.AutoCAD.Interop
'Imports System.Windows.Forms
'Public Class Form1
'
'Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
'Dim oApp As AcadApplication
'' Try to connect to a running instance of AutoCAD.
'Try
'oApp = GetObject(, "AutoCAD.Application")
'' Failed to get AutoCAD.
'MsgBox ("Could not connect to AutoCAD.")
'Exit Sub
'
'
'End Try
'
'Dim myCom As myCustomCom
'myCom = oApp.GetInterfaceObject("comServerFromArx.myCustomCom")
'
'Dim x As Single
'Dim y As Single
'Dim z As Single
'myCom.getPosition(x, y, z)
'MessageBox.Show ("(" + x.ToString() + "," + y.ToString() + "," + z.ToString() + ")")
'End Sub
'End Class
'
Private Sub CommandButton3_Click()

    Dim returnObj As Object
    Dim basePnt As Variant
    Dim i As Integer
    Dim j As Integer
  
    Dim SEOBJ As SmileyEntity
    
    ThisDrawing.Utility.GetEntity returnObj, basePnt, "選擇要插入有屬性的圖元："
    
    returnObj.center = basePnt
End Sub

Private Sub CommandButton4_Click()
    
    Dim SEOBJ As SmileyEntity
    Dim basePnt(0 To 2) As Double
    Dim cenPnt(0 To 2) As Double
    Dim returnPnt As Variant
    Dim rad As Integer
    
    Me.Hide
    returnPnt = ThisDrawing.Utility.GetPoint(basePnt, "\n選擇要插入的點")
    

    'Set SEOBJ = New SmileyEntity
    Dim SSS As New SmileyApplication
    Dim ss As SmileyUi.SmileyApplication
    
    Set SEOBJ = ThisDrawing.ModelSpace.AddCustomObject("ASDKSMILEY")
    Call SEOBJ.setNewObject(returnPnt, 255)
    
    'sss.CreateSmiley
    
    'Dim aaa As Variant
    'aaa = SEOBJ.MouthCenter
    Me.Show
End Sub








