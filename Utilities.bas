Attribute VB_Name = "Utilities"
' ***** Utilities *****
' * - Common functions used throughout
' * - AUTHOR: Bryan Huss, 8/15/18
' *****************

' ***************
' * GetMonthSheets
' * - Returns a collection of Sheets from passed Workbook that names begin with months
' ***************
Public Function GetMonthSheets(wb As Workbook) As Collection
    Dim months As New Collection
    months.Add "Jan*"
    months.Add "Feb*"
    months.Add "Mar*"
    months.Add "Apr*"
    months.Add "May*"
    months.Add "Jun*"
    months.Add "Jul*"
    months.Add "Aug*"
    months.Add "Sep*"
    months.Add "Oct*"
    months.Add "Nov*"
    months.Add "Dec*"

    Dim sheetsCol As New Collection
    For x = 1 To wb.Sheets.count
        For m = 1 To months.count
            If wb.Sheets(x).name Like months(m) Then
                sheetsCol.Add wb.Sheets(x)
            End If
        Next m
    Next x
    
    Set GetMonthSheets = sheetsCol
End Function

' ***************
' * GetLastRow
' * - Returns the row number of the last row with a border style on the bottom
' ***************
Function GetLastRow(ws As Worksheet) As Long
    Dim row As Long
    ' Data starts on row 5
    row = 5
    
    Dim borderTest As Boolean
    ' True while there is a bottom border on the cell analyzed
    borderTest = True
    
    While borderTest
        ' Lack of bottom border style signals the end of the spreadsheet
        If ws.Cells(row, 1).Borders(xlEdgeBottom).LineStyle = xlNone Then
            borderTest = False
            row = row - 1
        Else
            row = row + 1
        End If
    Wend
    
    GetLastRow = row
End Function

