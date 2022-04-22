VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UFTicketDump 
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4110
   OleObjectBlob   =   "UFTicketDump.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UFTicketDump"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' ***** UFTicketDump *****
' * User Form for selection of which sheets to dump
' * AUTHOR: Bryan Huss, 8/30/18
' **********************

' Initialize
Private Sub UserForm_Initialize()
    ' Defaulting options on user form
    With ListBox
        .MultiSelect = fmMultiSelectMulti: .ListStyle = fmListStyleOption
    End With
    
    ' Get sheets from active workbook that have month names
    Dim months As Collection
    Set months = GetMonthSheets(ActiveWorkbook)
    
    ' Add names of sheets to listbox
    For Each s In months
        ListBox.AddItem s.name
    Next
    
    ' Memory management
    Set months = Nothing
End Sub

' On close, ensure unload
Private Sub UserForm_Terminate()
    Unload Me
End Sub

' Command Buttons
Private Sub cbSelectAll_Click()
    ' If button says "Select All"
    If cbSelectAll.caption = "Select All" Then
        For x = 0 To ListBox.ListCount - 1
            ListBox.Selected(x) = True
        Next x
        
        cbSelectAll.caption = "Deselect All"
        Repaint
    ' Else if button says "Deselect All"
    ElseIf cbSelectAll.caption = "Deselect All" Then
        For x = 0 To ListBox.ListCount - 1
            ListBox.Selected(x) = False
        Next x
        
        cbSelectAll.caption = "Select All"
        Repaint
    End If
End Sub

Private Sub cbCancel_Click()
    Unload Me
End Sub

Private Sub cbRun_Click()
    Dim sheetName As String
    If Len(Trim(tbSheetName.value)) = 0 Then
        MsgBox "Enter a name for output sheet.", vbExclamation, "Herp Derp!"
        With tbSheetName
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    Else
        Dim nameTaken As Boolean
        nameTaken = False
        For s = 1 To ActiveWorkbook.Sheets.count
            If UCase(tbSheetName.value) = UCase(ActiveWorkbook.Sheets(s).name) Then
                nameTaken = True
            End If
        Next s
        
        If nameTaken Then
            MsgBox "Sheet name already taken. Try a different one.", vbInformation, "Herp Derp!"
            With tbSheetName
                .SetFocus
                .SelStart = 0
                .SelLength = Len(.Text)
            End With
        Else
            ' Adds selected sheets to collection
            Dim selectedSheets As New Collection
            For x = 0 To ListBox.ListCount - 1
                If ListBox.Selected(x) Then
                    selectedSheets.Add ActiveWorkbook.Worksheets(ListBox.List(x))
                End If
            Next x
        
            ' Nothing selected to search from
            If selectedSheets.count < 1 Then
                MsgBox "No sheets have been selected!", vbExclamation, "Herp Derp!"
            Else
                Me.Hide
            
                Ticket_Dump tbSheetName.value, selectedSheets, True
            
                ' Memory management
                Set selectedSheets = Nothing
                Unload Me
            End If
        End If
    End If
End Sub
