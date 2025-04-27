VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TuneLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public columns As Dictionary     ' as  { formatting as  any, lines as  any }
Public image As String
Public newpage As Integer
Public staffbreak As Integer
    '' Only one of separator, subtitle, text, or staff will be present
Public separator As Dictionary
'    {
'        endChar As Number
'        lineLength As Number
'        spaceAbove As Number
'        spaceBelow As Number
'        startChar As Number
'    }
Public subtitle As Dictionary
'    {
'        endChar As Number
'        startChar As Number
'                text As String
'    }
Public text As Dictionary
'    {
'        endChar As Number
'        startChar As Number
'        text As TextFieldProperties
'    }
Public Staffs      As StaffList  ' Array<Staff>
Public staffGroup As StaffGroupElement
Public vskip As Double
