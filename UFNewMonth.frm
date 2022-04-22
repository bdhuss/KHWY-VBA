VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UFNewMonth 
   Caption         =   "UserForm1"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3345
   OleObjectBlob   =   "UFNewMonth.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UFNewMonth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' ***** UFNewMonth *****
' * AUTHOR: Bryan Huss, 8/16/18
' **********************
Private Sub cbCancel_Click()
    Unload Me
End Sub

Private Sub cbCreate_Click()
    If (Len(Trim(tbName.Text)) = 0) Or Not IsDate(tbName.Text) Then
        MalformedInput
    Else
        Dim str() As String
        str = Split(tbName.Text)
        
        If UBound(str) <> 1 Then
            MalformedInput
        Else
            If Not Len(str(1)) = 4 Then
                MalformedInput
            Else
                Me.Hide
                CreateNewMonth tbName.Text
                Unload Me
            End If
        End If
    End If
End Sub

Private Sub UserForm_Terminate()
    Unload Me
End Sub

Private Sub MalformedInput()
    MsgBox "MONTH YYYY input required.", vbInformation, "Malformed Input"
    With tbName
        .Text = ""
        .SetFocus
    End With
End Sub
