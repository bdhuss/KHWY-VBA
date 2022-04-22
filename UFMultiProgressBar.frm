VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UFMultiProgressBar 
   Caption         =   "Running..."
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4320
   OleObjectBlob   =   "UFMultiProgressBar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UFMultiProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub LBar2_Click()

End Sub

' INIT
Private Sub UserForm_Initialize()
    LDescription1.caption = ""
    LBar1.Width = 0
    LDescription2.caption = ""
    LPercentage2.caption = "0% Complete"
    LBar2.Width = 0
End Sub

' Set caption for top label
Public Sub SetDescription1(str As String)
    LDescription1.caption = str
End Sub

' Set caption for bottom label
Public Sub SetDescription2(str As String)
    LDescription2.caption = str
End Sub

' Progress update for top progress bar
Public Sub Progress1(current As Long, total As Long)
    Dim p1 As Long
    p1 = Round((current / total) * 100, 0)
    LBar1.Width = p1 * 2
    
    DoEvents
End Sub

' Progress update for bottom progress bar and percentage label
Public Sub Progress2(current As Long, total As Long)
    Dim p2 As Long
    p2 = Round((current / total) * 100, 0)
    LPercentage2.caption = p2 & "% Complete"
    LBar2.Width = p2 * 2
    
    DoEvents
End Sub

' Resets both bars and percentage string, sets description labels to passed strings
Public Sub ResetAll(str1 As String, str2 As String)
    LDescription1.caption = str1
    LBar1.Width = 0
    LDescription2.caption = str2
    LPercentage2.caption = "0% Complete"
    LBar2.Width = 0
End Sub

' Resets top progress bar and description label to passed string
Public Sub Reset1(str As String)
    LDescription1.caption = str
    LBar1.Width = 0
End Sub

' Resets bottom progress bar and percentae label and description label
Public Sub Reset2(str As String)
    LDescription2.caption = str
    LPercentage2.caption = "0% Complete"
    LBar2.Width = 0
End Sub
