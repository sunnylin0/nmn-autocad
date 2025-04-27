VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vStaffInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public top As Double
Public bottom As Double
Public lines As Double
Public voices As iArray
Public specialY As Dictionary
Public absoluteY As Double
Public bottomLine As Double
Public topLine As Double

Private Sub Class_Initialize()
    Set voices = New iArray
    Set specialY = New Dictionary
End Sub
