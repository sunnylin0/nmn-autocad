VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMIDIFILE 
   Caption         =   "UserForm1"
   ClientHeight    =   3204
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   4932
   OleObjectBlob   =   "frmMIDIFILE.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMIDIFILE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim ptn0(10) As Byte
    Dim ptn1(4) As Byte
    Dim ptn2(3) As Byte
    Dim ptn_len(3) As Byte  '這是記 ptn0,ptn1,ptn2 的長度
Sub setMain()
    'MIDI檔標誌MThd
     ptn0(1) = AscW("M")
     ptn0(2) = AscW("T")
     ptn0(3) = AscW("h")
     ptn0(4) = AscW("d")
     ptn0(5) = 0
     ptn0(6) = 0
     ptn0(7) = 0
     ptn0(8) = 6
     ptn0(9) = 0
     ptn0(10) = 1
     
     '磁軌標誌MTrk
     ptn1(1) = AscW("M")
     ptn1(2) = AscW("T")
     ptn1(3) = AscW("r")
     ptn1(4) = AscW("k")
     
     '磁軌結束資訊
     ptn2(1) = &HFF
     ptn2(2) = &H2F
     ptn2(3) = &H0
     
     '這是記 ptn0,ptn1,ptn2 的長度
     ptn_len(1) = 10: ptn_len(2) = 4: ptn_len(3) = 3:
    'Public Shared ptn_tbl()() As Byte = {ptn0, ptn1, ptn2}
    
End Sub
Private Sub CommandButton1_Click()
    setMain
    
    Dim vbaPATH As String
    Dim imagePATH As String
    Dim pos As Integer
    '取得 dvb 的目錄
    vbaPATH = ThisDrawing.Application.vbe.ActiveVBProject.FileName
    pos = InStrRev(vbaPATH, "\")
    imagePATH = Mid(vbaPATH, 1, pos)

    Dim MIDI_FILE As Integer
    Dim path_1 As String
    Dim i
    MIDI_FILE = 1
    path_1 = imagePATH & "test3.txt"
    
    Open path_1 For Binary Access Write As #MIDI_FILE
    For i = 0 To ptn_len(1)
        Put #MIDI_FILE, , ptn0(i)
    Next
    Put #MIDI_FILE, , 12
    Put #MIDI_FILE, , 130
    Put #MIDI_FILE, , 12
    Put #MIDI_FILE, , 13
    Put #MIDI_FILE, , "Apple"
    Put #MIDI_FILE, , "Banana"
    Put #MIDI_FILE, , "Cat"
    Put #MIDI_FILE, , "Dog"
    Put #MIDI_FILE, , "Erase"
    Put #MIDI_FILE, , "Foolish"
    Put #MIDI_FILE, , "Hot"
    Close #MIDI_FILE
End Sub
