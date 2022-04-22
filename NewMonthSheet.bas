Attribute VB_Name = "NewMonthSheet"
' ***** CreateNewMonth *****
' * - Takes user input for MONTH YEAR and creates a new
' * fuel sales monthly spreadsheet with all automated
' * fields, dates, and formulas correctly editted and
' * updated.
' * - AUTHOR: Bryan Huss, 8/9/18
' *************************
Public Sub CreateNewMonth(sheetName As String)
    ' Copies master sheet then renames based on user input
    Sheets("MASTER - DO NOT USE").Copy Before:=Sheets(1)
    Sheets("MASTER - DO NOT USE (2)").name = sheetName
    
    ' Enters user input MONTH YEAR into spreadsheet, colors it red,
    ' and clears automated date from offset cell
    With Sheets(sheetName).Range("A3")
        .value = "Total Gallons: " + sheetName
        .Characters(1, 14).Font.ColorIndex = 1
        .Offset(0, 1).ClearContents
    End With
    
    ' Turns user input string into CDate object
    startingDate = CDate(sheetName)
    ' Gets number of days in month indicated by user input
    numDays = Day(DateSerial(Year(startingDate), Month(startingDate) + 1, 1) - 1)
    ' First summation row
    rowNum = 30
    
    ' Formula strings
    LLMeterTotal = "=SUM("
    LLMeterDiff = "=SUM("
    LLDiff = "=SUM("
    JetMeterTotal = "=SUM("
    JetMeterDiff = "=SUM("
    JetDiff = "=SUM("
    cash = "=SUM("
    check = "=SUM("
    credit = "=SUM("
    tabT = "=SUM("
    
    ' Cycles through days of month for summation rows
    For d = 0 To (numDays - 1)
        currentDate = CDate(DateSerial(Year(startingDate), _
                                          Month(startingDate), _
                                          Day(startingDate)) _
                                          + d)
        With Sheets(sheetName)
            .Cells(rowNum, 1).value = (currentDate & " Daily Subtotal:")
            .Cells(rowNum, 2).ClearContents
        End With
        
        ' First day of month
        If d = 0 Then
            LLMeterTotal = LLMeterTotal & "G" & rowNum
            LLMeterDiff = LLMeterDiff & "H" & rowNum
            LLDiff = LLDiff & "I" & rowNum
            JetMeterTotal = JetMeterTotal & "L" & rowNum
            JetMeterDiff = JetMeterDiff & "M" & rowNum
            JetDiff = JetDiff & "N" & rowNum
            cash = cash & "R" & rowNum
            check = check & "S" & rowNum
            credit = credit & "T" & rowNum
            tabT = tabT & "U" & rowNum
        ' Any other day besides first of month
        Else
            LLMeterTotal = LLMeterTotal & ",G" & rowNum
            LLMeterDiff = LLMeterDiff & ",H" & rowNum
            LLDiff = LLDiff & ",I" & rowNum
            JetMeterTotal = JetMeterTotal & ",L" & rowNum
            JetMeterDiff = JetMeterDiff & ",M" & rowNum
            JetDiff = JetDiff & ",N" & rowNum
            cash = cash & ",R" & rowNum
            check = check & ",S" & rowNum
            credit = credit & ",T" & rowNum
            tabT = tabT & ",U" & rowNum
        End If
        
        ' Adds non-summation rows to int
        rowNum = rowNum + 27
    Next d
    
    ' Closes formula strings
    LLMeterTotal = LLMeterTotal & ")"
    LLMeterDiff = LLMeterDiff & ")"
    LLDiff = LLDiff & ")"
    JetMeterTotal = JetMeterTotal & ")"
    JetMeterDiff = JetMeterDiff & ")"
    JetDiff = JetDiff & ")"
    cash = cash & ")"
    check = check & ")"
    credit = credit & ")"
    tabT = tabT & ")"

    ' Inserts summation strings into appropriate cells
    With Sheets(sheetName)
        .Range("G3").Formula = LLMeterTotal
        .Range("H3").Formula = LLMeterDiff
        .Range("I3").Formula = LLDiff
        .Range("L3").Formula = JetMeterTotal
        .Range("M3").Formula = JetMeterDiff
        .Range("N3").Formula = JetDiff
        .Range("R3").Formula = cash
        .Range("S3").Formula = check
        .Range("T3").Formula = credit
        .Range("U3").Formula = tabT
    End With
    
    ' Deletes unneeded rows from sheet
    If numDays = 28 Then
        Sheets(sheetName).Rows("761:841").Delete
        Sheets(sheetName).Cells(761, 2).ClearContents
    ElseIf numDays = 29 Then
        Sheets(sheetName).Rows("788:841").Delete
        Sheets(sheetName).Cells(788, 2).ClearContents
    ElseIf numDays = 30 Then
        Sheets(sheetName).Rows("815:841").Delete
        Sheets(sheetName).Cells(815, 2).ClearContents
    End If
    
    ' Year to date formula strings
    Dim ytdAvgas As String
    Dim ytdJet As String
    Dim ytd As String
    
    ' Cycles through month spreadsheets and adds them to YTD forumla strings
    For s = 1 To (ActiveWorkbook.Sheets.count - 2)
        If s = 1 Then
            ytdAvgas = "'" & Sheets(s).name & "'!G3"
            ytdJet = "'" & Sheets(s).name & "'!L3"
        Else
            ytdAvgas = ytdAvgas & ",'" & Sheets(s).name & "'!G3"
            ytdJet = ytdJet & ",'" & Sheets(s).name & "'!L3"
        End If
    Next s
    
    ' Creates and inputs YTD formula
    ytd = "=CONCATENATE(ROUND(SUM(" & ytdAvgas & "), 1), "" 100LL || "", ROUND(SUM(" & ytdJet & "), 0), "" JET-A"")"
    Sheets(sheetName).Range("V3").Formula = ytd
    
    ' Selects first cell of sheet
    Sheets(sheetName).Range("A1").Select
End Sub
