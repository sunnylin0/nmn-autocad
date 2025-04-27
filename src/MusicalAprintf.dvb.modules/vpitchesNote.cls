VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vpitchesNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public acciaccatura As Boolean     ''是否為碎音（acciaccatura） as 這種小字體、加上斜線的音符，意思是很快地彈過去，只占非常短的時間。
Public accidental As String 'AccidentalName //臨時記號（accidental）
Public duration As Double
Public pitch As Integer
Public Name As String
Public midipitch As Double
Public startSlur As vSlurList
'{
'    label As Integer
'    direction as 'up', 'down'
'    style as 'dotted'
'}[]
Public endSlur As iArray 'Array<double>
Public startTie As vSlur
'{
'    direction as 'up', 'down'
'    style as 'dotted'
'}

Public endTie As Boolean
Public dottedSlur As Boolean
Public verticalPos As Double
Public highestVert As Double
Public style As String

Public finge As String      '第一個指法
Public towfinge As String   '第二個指法
