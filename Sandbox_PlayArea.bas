Attribute VB_Name = "Sandbox"
Public Sub sandbox_DateTime()
    
End Sub

Public Sub sandbox_AccessingCollectionsTest()
    Dim parentCol As New Collection
    Dim ws As Worksheet
    Set ws = Sheets("DATA")
    
    Run "Utilities.DataToCollections", ws, parentCol
    
    For pc = 1 To parentCol.Count
        For ic = 1 To parentCol(pc).Count
            Debug.Print pc & ":" & ic & ":" & parentCol(pc)(ic).Count
        Next ic
    Next pc
    
    Dim avgasTotal As Double: avgasTotal = 0#
    For x = 1 To parentCol(1)(7).Count
        avgasTotal = avgasTotal + parentCol(1)(7)(x).AvgasMeterDiffManual
    Next x
    
    Dim avgasTotal1 As Double: avgasTotal1 = 0#
    For x = 1 To parentCol(1)(8).Count
        avgasTotal1 = avgasTotal1 + parentCol(1)(8)(x).AvgasMeterDiffManual
    Next x
    
    Debug.Print "parentCol(1)(7) AvGas Total = " & avgasTotal
    Debug.Print "parentCol(1)(8) AvGas Total = " & avgasTotal1
End Sub

Public Sub sandbox_LastRowColumn()
    Debug.Print ActiveSheet.Cells(rows.Count, 2).End(xlUp).Row
    Dim ws As Worksheet
    Set ws = Sheets("DATA")
    Debug.Print ws.Cells(rows.Count, 2).End(xlUp).Row
    Debug.Print ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column
End Sub

Public Sub sandbox_MultidimensionalArrays()
    Dim arr() As Integer
    
    ReDim arr(1 To 10, 1 To 3)
    
    For x = 1 To 10
        For y = 1 To 3
            arr(x, y) = x
        Next y
    Next x
    
    For Row = 1 To 10
        For Col = 1 To 3
            Cells(Row, Col).Value = arr(Row, Col)
        Next Col
    Next Row
    
    ReDim Preserve arr(1 To 12, 1 To 3)
    
    For x = 1 To 12
        For y = 1 To 3
            arr(x, y) = y
        Next y
    Next x
    
    
    For r = 1 To 12
        For c = 1 To 3
            Cells(r + 5, c + 12).Value = arr(r, c)
        Next c
    Next r
End Sub


Public Sub sandbox_NestedCollections()
    Dim parentCol As New Collection
    
    ' Inital parent collection size
    For pc = 1 To Cells(1, 1).Value
        parentCol.Add New Collection
    Next pc
    
    For pc = 1 To parentCol.Count
        ' 12 months in a year
        For ic = 1 To 12
            parentCol(pc).Add New Collection
        Next ic
    Next pc
    
    ' Input a blank CTicketItem based on sheet input
    For x = 1 To 3
        parentCol(x)(Cells(x, 2).Value).Add New CTicketItem
    Next x
    
    ' Cycle through nested collections and print item counts
    Debug.Print "pc.count = " & parentCol.Count
    For pc = 1 To parentCol.Count
        Debug.Print "parentCol(" & pc & ").count = " & parentCol(pc).Count
        For m = 1 To parentCol(pc).Count
            Debug.Print "parentCol(" & pc & ")(" & m & ").count = " & parentCol(pc)(m).Count
        Next m
    Next pc
End Sub


Public Sub sandbox()
End Sub
