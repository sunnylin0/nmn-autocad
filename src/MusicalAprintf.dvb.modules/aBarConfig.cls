VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aBarConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public barId As Integer     '實既小節數
Public barNumber As Integer '小節數修改號碼
Public mete As Integer
Public mete2 As Integer
Public keyMain As String
Public barLineQuantity As Integer '一行的小節數量
Public barlinebreak As Boolean  'true 小節後換行
Public barpagebreak As Boolean
