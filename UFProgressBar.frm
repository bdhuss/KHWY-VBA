VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UFProgressBar 
   Caption         =   "Progress Indicator"
   ClientHeight    =   1080
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4335
   OleObjectBlob   =   "UFProgressBar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UFProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' User Form 'ProgressBar' Initialization
Private Sub UserForm_Initialize()
    LPercentage.caption = "0% Complete"
    LBar.Width = 0
End Sub

' Progress percentage indicator for drawing on User Form 'Progress'
Public Sub Progress(current As Long, total As Long)
    Dim percCompl As Long
    percCompl = Round((current / total) * 100, 0)
    LPercentage = percCompl & "% Complete"
    LBar.Width = percCompl * 2
    
    DoEvents
End Sub

Public Sub SetCaption(str As String)
    LPercentage = str
End Sub

Public Sub Reset(str As String)
    LBar.Width = 0
    LPercentage = str
End Sub
