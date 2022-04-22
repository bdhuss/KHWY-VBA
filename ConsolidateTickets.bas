Attribute VB_Name = "ConsolidateTickets"
' ***** Consolidate *****
' * - Consolidates ticket entries based on input
' * - searchType: Integer - row to search for matching criteria (name or N#)
' * - searchCriteria: String - criteria to search for (name or N#)
' * - searchSheets: Collection - collection of worksheets to search for matching criteria (name or N#)
' * - showProgressBar: Boolean - shows or hides progress bar while performing search and consolidation
' * - AUTHOR: Bryan Huss, 8/16/18
' *********************

Public Sub Consolidate_Tickets(searchType As Integer, searchCriteria As String, searchSheets As Collection, Optional showProgressBar As Boolean)
On Error GoTo ErrorHandler:
    If showProgressBar Then
        UFProgressBar.Show xlModeless
        UFProgressBar.SetCaption "Initial setup..."
    End If
    
    ' Variables
    Dim tickets As New Collection
    Dim progressTotal As Long
    Dim progressCurrent As Long
    progressTotal = 0: progressCurrent = 0
    
    ' Get total number of rows to analyze
    For w = 1 To searchSheets.count
        progressTotal = progressTotal + GetLastRow(searchSheets(w))
    Next w
    
    ' Cycle through selected/passed worksheets searching for searchCriteria.
    ' Creates a CTicketEntry object for each matching criteria then adds it to
    ' a Collection of objects.
    For ws = 1 To searchSheets.count
        ' Update progress bar
        If showProgressBar Then
            progressCurrent = progressCurrent + 4
        End If
        
        ' Cycle through the sheet starting at first valid ticket row
        For r = 5 To GetLastRow(searchSheets(ws))
            If searchSheets(ws).Cells(r, searchType).value = searchCriteria Then
                Dim paid As Double
                Select Case searchSheets(ws).Cells(r, 17).value
                    Case 1
                        paid = searchSheets(ws).Cells(r, 18).value
                    Case 2
                        paid = searchSheets(ws).Cells(r, 19).value
                    Case 3
                        paid = searchSheets(ws).Cells(r, 20).value
                    Case 4
                        paid = searchSheets(ws).Cells(r, 21).value
                    Case Else
                        paid = 0#
                End Select
                
                Dim cte As CTicketEntry
                Set cte = New CTicketEntry
                cte.TicketEntry TicketNum:=searchSheets(ws).Cells(r, 1).value, _
                                        PurchaseDate:=searchSheets(ws).Cells(r, 2).value, _
                                        TailNum:=searchSheets(ws).Cells(r, 3).value, _
                                        name:=searchSheets(ws).Cells(r, 4).value, _
                                        AVGASMeterStart:=Abs(searchSheets(ws).Cells(r, 5).value), _
                                        AVGASMeterStop:=Abs(searchSheets(ws).Cells(r, 6).value), _
                                        AVGASMeterDiffManual:=Abs(searchSheets(ws).Cells(r, 7).value), _
                                        AVGASMeterDiffAuto:=Abs(searchSheets(ws).Cells(r, 8).value), _
                                        AVGASDiffDiff:=Abs(searchSheets(ws).Cells(r, 9).value), _
                                        JETMeterStart:=Abs(searchSheets(ws).Cells(r, 10).value), _
                                        JETMeterStop:=Abs(searchSheets(ws).Cells(r, 11).value), _
                                        JETMeterDiffManual:=Abs(searchSheets(ws).Cells(r, 12).value), _
                                        JETMeterDiffAuto:=Abs(searchSheets(ws).Cells(r, 13).value), _
                                        JETDiffDiff:=Abs(searchSheets(ws).Cells(r, 14).value), _
                                        FuelPPG:=searchSheets(ws).Cells(r, 15).value, _
                                        NFPT:=searchSheets(ws).Cells(r, 16).value, _
                                        PayCode:=searchSheets(ws).Cells(r, 17).value, _
                                        AmountPaid:=paid, _
                                        Comments:=searchSheets(ws).Cells(r, 22).value
                
                tickets.Add cte
                
            End If
            
            If showProgressBar Then
                progressCurrent = progressCurrent + 1
                UFProgressBar.Progress progressCurrent, progressTotal
            End If
            
        Next r
    Next ws
    
    ' If no matches were found for searchCriteria on searchType
    If tickets.count < 1 Then
        If showProgressBar Then
            Unload UFProgressBar
        End If
        
        MsgBox "No matches found for search criteria", vbOKOnly, "No matches"
        Exit Sub
    End If
    
    ' Reset progress bar for creation and consolidation of data
    If showProgressBar Then
        UFProgressBar.Reset "Setting up new spreadsheet..."
    End If
    
    ' Setup of new worksheet for data display
    Dim newWS As Worksheet
    Set newWS = ActiveWorkbook.Worksheets.Add
    Dim d As Date
    d = Date
    newWS.name = Month(d) & "." & Day(d) & "." & Year(d) & " " & searchCriteria
    
    With newWS
        .Range("A1:I1").Font.FontStyle = "Bold"
        .Range("A1:I1").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("A1:I1").Borders(xlEdgeBottom).Weight = xlThick
        .Range("A1:I1").WrapText = True
        .Range("A1:I1").AutoFilter
        .Range("A:D").HorizontalAlignment = xlLeft
        .Range("E:I").HorizontalAlignment = xlRight
        .Range("A1:I1").HorizontalAlignment = xlCenter
        .Range("A1:I1").VerticalAlignment = xlCenter
        .Columns("A").ColumnWidth = 7
        .Range("A1").value = "TICKET#"
        .Columns("B").ColumnWidth = 10
        .Range("B1").value = "DATE"
        .Columns("C").ColumnWidth = 7
        .Range("C1").value = "TAIL#"
        .Columns("D").ColumnWidth = 10
        .Range("D1").value = "NAME"
        .Columns("E").ColumnWidth = 8
        .Range("E1").value = "AVGAS (gal)"
        .Columns("F").ColumnWidth = 7
        .Range("F1").value = "JET (gal)"
        .Columns("G").ColumnWidth = 5.5
        .Range("G1").value = "PAY CODE"
        .Columns("H").ColumnWidth = 7.5
        .Range("H1").value = "Price / gal"
        .Columns("I").ColumnWidth = 10
        .Range("I1").value = "TOTAL"
    End With
    
    ' Set progress bar to zero
    If showProgressBar Then
        UFProgressBar.Progress CLng(0), CLng(tickets.count)
    End If
    
    ' Variables
    Dim totalPaid As Double
    Dim avgasTotal As Double
    Dim jetTotal As Double
    totalPaid = 0#: avgasTotal = 0#: jetTotal = 0#
    
    ' Print collected data to new worksheet
    For t = 1 To tickets.count
        newWS.Cells(t + 1, 1).value = tickets(t).TicketNum
        newWS.Cells(t + 1, 2).value = tickets(t).PurchaseDate
        newWS.Cells(t + 1, 3).value = tickets(t).TailNum
        newWS.Cells(t + 1, 4).value = tickets(t).name
        newWS.Cells(t + 1, 5).value = tickets(t).AVGASMeterDiffAuto
        avgasTotal = avgasTotal + tickets(t).AVGASMeterDiffAuto
        newWS.Cells(t + 1, 6).value = tickets(t).JETMeterDiffAuto
        jetTotal = jetTotal + tickets(t).JETMeterDiffAuto
        newWS.Cells(t + 1, 7).value = tickets(t).PayCode
        newWS.Cells(t + 1, 8).value = tickets(t).FuelPPG
        newWS.Cells(t + 1, 9).value = tickets(t).AmountPaid
        totalPaid = totalPaid + tickets(t).AmountPaid
        
        ' Update progress bar
        If showProgressBar Then
            UFProgressBar.Progress CLng(t), CLng(tickets.count)
        End If
    Next t
    
    ' List totals below respective cells
    ' Add borders for totals and ticket volume
    ' Sets number formats for columns/cells
    With newWS
        .Cells(tickets.count + 2, 1).value = "TOTALS"
        .Columns("E").NumberFormat = "###,##0.0"
        .Cells(tickets.count + 2, 5).value = Round(avgasTotal, 1)
        '.Cells(tickets.count + 2, 5).NumberFormat = "###,##0.0"
        .Cells(tickets.count + 2, 6).value = jetTotal
        .Cells(tickets.count + 2, 6).NumberFormat = "###,##0"
        .Cells(tickets.count + 2, 9).value = totalPaid
        .Range(.Cells(2, 1), .Cells(tickets.count + 1, 9)).BorderAround xlContinuous, xlThin
        .Range(.Cells(tickets.count + 2, 1), .Cells(tickets.count + 2, 9)).BorderAround xlContinuous, xlThick
        .Columns("H").NumberFormat = "$0.00"
        .Columns("I").NumberFormat = "$###,##0.00"
    End With
    
    ' Freeze top row
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True
    End With
    
    ' Sort on date, Oldest --> Newest
    newWS.AutoFilter.Sort.SortFields.Add Key:=Range("B1"), _
                                                                SortOn:=xlSortOnValues, _
                                                                Order:=xlDescending, _
                                                                DataOption:=xlSortNormal
    With newWS.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
' Memory management on close
BeforeExit:
    If showProgressBar Then
        Unload UFProgressBar
    End If
    Set tickets = Nothing
    Set ticket = Nothing
    Set newWS = Nothing
Exit Sub

' Error Handler
ErrorHandler:
'    MsgBox Err.Description, vbCritical, "ERROR"
    Debug.Print Err.Description
Resume Next
End Sub
