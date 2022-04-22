Attribute VB_Name = "TicketDump"
' ***** TicketDump *****
' * - Collects and dumps all tickets from selected sheets to a new sheet
' * - searchSheets: Collection - collection of worksheets to collect tickets from
' * - showProgressBar: Boolean - shows or hides progress bar
' * - AUTHOR: Bryan Huss, 8/30/18
' ********************

Public Sub Ticket_Dump(sheetName As String, searchSheets As Collection, showProgressBar As Boolean)
On Error GoTo ErrorHandler:
    Application.ScreenUpdating = False
    
    If showProgressBar Then
        UFMultiProgressBar.Show xlModeless
        UFMultiProgressBar.SetDescription1 "Initial setup..."
    End If
    
    ' Variables
    Dim tickets As New Collection
    Dim sheetRows As Long
    
    ' Cycle through selected/passed worksheets searching for any and all tickets.
    ' Creates a CTicketEntry object for any ticket it finds and adds it to a collection
    ' of tickets.
    For ws = 1 To searchSheets.count
        If showProgressBar Then
            UFMultiProgressBar.SetDescription1 searchSheets(ws).name
        End If
        
        sheetRows = GetLastRow(searchSheets(ws))
        
        ' Cycle through each sheet starting at row 5
        For r = 5 To sheetRows
            If Not IsEmpty(searchSheets(ws).Cells(r, 2)) Then
                ' Gets the amount paid cell based on payment type
                Dim paid As Double
                Select Case searchSheets(ws).Cells(r, 17).value
                    Case 1
                        paid = Round(searchSheets(ws).Cells(r, 18).value, 2)
                    Case 2
                        paid = Round(searchSheets(ws).Cells(r, 19).value, 2)
                    Case 3
                        paid = Round(searchSheets(ws).Cells(r, 20).value, 2)
                    Case 4
                        paid = Round(searchSheets(ws).Cells(r, 21).value, 2)
                    Case Else
                        paid = 0#
                End Select

                ' Create CTicketEntry object for data in sheet
                Dim cte As CTicketEntry
                Set cte = New CTicketEntry
                cte.TicketEntry TicketNum:=searchSheets(ws).Cells(r, 1).value, _
                                        PurchaseDate:=searchSheets(ws).Cells(r, 2).value, _
                                        TailNum:=searchSheets(ws).Cells(r, 3).value, _
                                        name:=searchSheets(ws).Cells(r, 4).value, _
                                        AVGASMeterStart:=Abs(Round(searchSheets(ws).Cells(r, 5).value, 1)), _
                                        AVGASMeterStop:=Abs(Round(searchSheets(ws).Cells(r, 6).value, 1)), _
                                        AVGASMeterDiffManual:=Abs(Round(searchSheets(ws).Cells(r, 7).value, 1)), _
                                        AVGASMeterDiffAuto:=Abs(Round(searchSheets(ws).Cells(r, 8).value, 1)), _
                                        AVGASDiffDiff:=Abs(Round(searchSheets(ws).Cells(r, 9).value, 1)), _
                                        JETMeterStart:=Abs(Round(searchSheets(ws).Cells(r, 10).value, 1)), _
                                        JETMeterStop:=Abs(Round(searchSheets(ws).Cells(r, 11).value, 1)), _
                                        JETMeterDiffManual:=Abs(Round(searchSheets(ws).Cells(r, 12).value, 1)), _
                                        JETMeterDiffAuto:=Abs(Round(searchSheets(ws).Cells(r, 13).value, 1)), _
                                        JETDiffDiff:=Abs(Round(searchSheets(ws).Cells(r, 14).value, 1)), _
                                        FuelPPG:=searchSheets(ws).Cells(r, 15).value, _
                                        NFPT:=searchSheets(ws).Cells(r, 16).value, _
                                        PayCode:=searchSheets(ws).Cells(r, 17).value, _
                                        AmountPaid:=paid, _
                                        Comments:=searchSheets(ws).Cells(r, 22).value
                
                ' Stores CTicketEntry object in collection
                tickets.Add cte
                    
            End If

            If showProgressBar Then
                UFMultiProgressBar.Progress2 CLng(r), CLng(sheetRows)
                UFMultiProgressBar.SetDescription2 ("Row: " & r & " of " & sheetRows)
            End If
        Next r
        
        If showProgressBar Then
            UFMultiProgressBar.Progress1 CLng(ws), CLng(searchSheets.count + 1)
        End If
        
    Next ws
    
    If showProgressBar Then
        UFMultiProgressBar.SetDescription1 "Dumping collection to new sheet..."
        UFMultiProgressBar.Reset2 "Setting up output sheet"
        UFMultiProgressBar.Progress1 CLng(searchSheets.count + 1), CLng(searchSheets.count + 1)
    End If
    
    ' Create new worksheet to dump all tickets to
    Dim newWS As Worksheet
    Set newWS = ActiveWorkbook.Worksheets.Add
    newWS.name = sheetName
    
    ' Dump all CTicketEntry object from collection to new worksheet
    For t = 1 To tickets.count
        newWS.Cells(t, 1).value = tickets(t).TicketNum
        newWS.Cells(t, 2).value = tickets(t).PurchaseDate
        newWS.Cells(t, 3).value = tickets(t).TailNum
        newWS.Cells(t, 4).value = tickets(t).name
        newWS.Cells(t, 5).value = tickets(t).AVGASMeterStart
        newWS.Cells(t, 6).value = tickets(t).AVGASMeterStop
        newWS.Cells(t, 7).value = tickets(t).AVGASMeterDiffManual
        newWS.Cells(t, 8).value = tickets(t).AVGASMeterDiffAuto
        newWS.Cells(t, 9).value = tickets(t).AVGASDiffDiff
        newWS.Cells(t, 10).value = tickets(t).JETMeterStart
        newWS.Cells(t, 11).value = tickets(t).JETMeterStop
        newWS.Cells(t, 12).value = tickets(t).JETMeterDiffManual
        newWS.Cells(t, 13).value = tickets(t).JETMeterDiffAuto
        newWS.Cells(t, 14).value = tickets(t).JETDiffDiff
        newWS.Cells(t, 15).value = tickets(t).FuelPPG
        newWS.Cells(t, 16).value = tickets(t).NFPT
        newWS.Cells(t, 17).value = tickets(t).PayCode
        newWS.Cells(t, 18).value = tickets(t).AmountPaid
        newWS.Cells(t, 19).value = tickets(t).Comments
        
        If showProgressBar Then
            UFMultiProgressBar.SetDescription2 ("Ticket " & t & " of " & tickets.count)
            UFMultiProgressBar.Progress2 CLng(t), CLng(tickets.count)
        End If
    Next t

' Memory management
BeforeExit:
    If showProgressBar Then
        Unload UFMultiProgressBar
    End If
    Set tickets = Nothing
    Application.ScreenUpdating = True
Exit Sub

' Lazy error handling. Build robust when time calls.
ErrorHandler:
    ' For line numbers, use Erl. Must number lines to work
    Debug.Print Err.Number & " : " & Err.Description
Resume Next
End Sub


