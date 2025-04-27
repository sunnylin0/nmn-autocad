VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vrestNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public typs As String  'invisible' | 'spacer' | 'rest' | 'multimeasure' | 'invisible-multimeasure' | 'whole'
Public text As Double
Public startTie As vSlur
Public endTie As Boolean
