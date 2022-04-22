VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UFConsolidate 
   Caption         =   "Consolidate"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4110
   OleObjectBlob   =   "UFConsolidate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UFConsolidate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' Initialize
Private Sub UserForm_Initialize()
    ' Defaulting options on user form
    obName = True
    obTail = False
    lSearchCriteria = "Name:"
    tbSearchCriteria = vbNullString
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
    
    tbSearchCriteria.SetFocus
    
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

Private Sub cbSearch_Click()
    ' Search criteria is empty
    If Len(tbSearchCriteria.Text) < 1 Then
        MsgBox "Did you want to search for something?" & vbNewLine & _
                     "Search criteria is empty.", vbExclamation, "I can't search for nothing silly!"
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
            MsgBox "No sheets have been selected!", vbExclamation, "Search where for what now?"
        Else
            Me.Hide
            
            ' Search by name
            If obName = True Then
                Consolidate_Tickets 4, tbSearchCriteria.Text, selectedSheets, True
            ' Search by tail #
            Else
                Consolidate_Tickets 3, tbSearchCriteria.Text, selectedSheets, True
            End If
            
            ' Memory management
            Set selectedSheets = Nothing
            Unload Me
        End If
    End If
End Sub

' Option Boxes
Private Sub obName_Click()
    lSearchCriteria.caption = "Name:"
    Repaint
End Sub

Private Sub obTail_Click()
    lSearchCriteria.caption = "Tail#:"
    Repaint
End Sub
