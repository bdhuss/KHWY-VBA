'*** Worksheet Change Listener ***
'
Private Const ticket As Integer = 1
Private Const DATE_COLUMN As Integer = 2
Private Const TAIL_NUMBER As Integer = 3
Private Const NAME_COLUMN As Integer = 4
Private Const AVGAS_METER_START As Integer = 5
Private Const AVGAS_METER_STOP As Integer = 6
Private Const AVGAS_METER_MANUAL As Integer = 7
Private Const AVGAS_METER_AUTO As Integer = 8
Private Const AVGAS_METER_DIFF As Integer = 9
Private Const JET_METER_START As Integer = 10
Private Const JET_METER_STOP As Integer = 11
Private Const JET_METER_MANUAL As Integer = 12
Private Const JET_METER_AUTO As Integer = 13
Private Const JET_METER_DIFF As Integer = 14
Private Const FUEL_PRICE As Integer = 15
Private Const PAYMENT_OPTION As Integer = 17
Private Const CASH_AMOUNT As Integer = 18
Private Const CHECK_AMOUNT As Integer = 19
Private Const CREDIT_AMOUNT As Integer = 20
Private Const TAB_AMOUNT As Integer = 21

Private Sub Worksheet_Change(ByVal Target As Range)
    
    ' Date input
    If Target.column = ticket Then
        Date_Input Target
    End If
    
    ' Tail number lookup
    If Target.column = TAIL_NUMBER And Not IsEmpty(Target) Then
        TailNumber_Lookup Target
    End If
   
   ' Name column changer
    If Target.column = NAME_COLUMN Then
        If IsEmpty(Target) Then
            Application.Cells(Target.row, PAYMENT_OPTION) = 0
        ElseIf Not IsEmpty(Target) And Not IsEmpty(Application.Cells(Target.row, TAIL_NUMBER)) Then
            PaymentOption_Input Application.Cells(Target.row, TAIL_NUMBER)
        End If
    End If
    
    ' AVGAS meter change
    If Target.column = AVGAS_METER_START Or Target.column = AVGAS_METER_STOP Or Target.column = AVGAS_METER_MANUAL Then
        If IsEmpty(Application.Cells(Target.row, AVGAS_METER_START)) Or IsEmpty(Application.Cells(Target.row, AVGAS_METER_STOP)) Or IsEmpty(Cells(Target.row, AVGAS_METER_MANUAL)) Then
            If Not IsEmpty(Application.Cells(Target.row, AVGAS_METER_AUTO)) Then
                Application.Cells(Target.row, AVGAS_METER_AUTO).ClearContents
                Application.Cells(Target.row, AVGAS_METER_DIFF).ClearContents
                Cells(Target.row, FUEL_PRICE).ClearContents
            End If
        Else
            AvgasCalc Target
        End If
        
        If Not IsEmpty(Application.Cells(Target.row, AVGAS_METER_START)) And Not IsEmpty(Application.Cells(Target.row, AVGAS_METER_STOP)) And Not IsEmpty(Application.Cells(Target.row, AVGAS_METER_MANUAL)) Then
            If IsEmpty(Application.Cells(Target.row, FUEL_PRICE)) Then
                Price_Input Target, "AVGAS"
            End If
        End If
    End If
    
    ' JET-A meter change
    If Target.column = JET_METER_START Or Target.column = JET_METER_STOP Or Target.column = JET_METER_MANUAL Then
        If IsEmpty(Application.Cells(Target.row, JET_METER_START)) Or IsEmpty(Application.Cells(Target.row, JET_METER_STOP)) Or IsEmpty(Cells(Target.row, JET_METER_MANUAL)) Then
            If Not IsEmpty(Application.Cells(Target.row, JET_METER_AUTO)) Then
                Application.Cells(Target.row, JET_METER_AUTO).ClearContents
                Application.Cells(Target.row, JET_METER_DIFF).ClearContents
                Cells(Target.row, FUEL_PRICE).ClearContents
            End If
        Else
            JetCalc Target
        End If
        
        If Not IsEmpty(Application.Cells(Target.row, JET_METER_START)) And Not IsEmpty(Application.Cells(Target.row, JET_METER_STOP)) And Not IsEmpty(Application.Cells(Target.row, JET_METER_MANUAL)) Then
            If IsEmpty(Application.Cells(Target.row, FUEL_PRICE)) Then
                Price_Input Target, "JET"
            End If
        End If
    End If
    
    ' Payment Option Input/Change
    If Target.column = PAYMENT_OPTION Then
        If Application.Cells(Target.row, AVGAS_METER_MANUAL).value > 0 Then
            Price_Input Target, "AVGAS"
        ElseIf Application.Cells(Target.row, JET_METER_MANUAL).value > 0 Then
            Price_Input Target, "JET"
        End If
    End If
    
    ' Price Cell Change
    If Target.column = FUEL_PRICE Then
        If IsEmpty(Application.Cells(Target.row, FUEL_PRICE)) And Not IsEmpty(Application.Cells(Target.row, PAYMENT_OPTION)) Then
            ClearTotals Target
        Else
            If Not IsEmpty(Cells(Target.row, AVGAS_METER_MANUAL)) And IsEmpty(Cells(Target.row, JET_METER_MANUAL)) Then
                ClearTotals Target
                CalcTotal Target, Cells(Target.row, AVGAS_METER_MANUAL), Cells(Target.row, PAYMENT_OPTION)
            ElseIf Not IsEmpty(Cells(Target.row, JET_METER_MANUAL)) And IsEmpty(Cells(Target.row, AVGAS_METER_MANUAL)) Then
                ClearTotals Target
                CalcTotal Target, Cells(Target.row, JET_METER_MANUAL), Cells(Target.row, PAYMENT_OPTION)
            End If
        End If
    End If
End Sub
    
    ' *** Date_Input ***
Private Sub Date_Input(ByVal Target As Range)
     If Not IsEmpty(Target) And IsEmpty(Application.Cells(Target.row, DATE_COLUMN)) Then
        Application.Cells(Target.row, DATE_COLUMN) = Date
    End If
End Sub

    ' *** TailNumber_Lookup ***
Private Sub TailNumber_Lookup(ByVal Target As Range)
    Dim lookup As Variant
    lookup = Application.VLookup(Target.value, Worksheets("TNLU").Range("A:B"), 2, False)
    
    If IsError(lookup) Then
        Application.Cells(Target.row, NAME_COLUMN).ClearContents
    Else
        Application.Cells(Target.row, NAME_COLUMN) = lookup
        PaymentOption_Input Target
    End If
End Sub

    ' *** PaymentOption_Input ***
Private Sub PaymentOption_Input(ByVal Target As Range)
    Dim lookupPaymentOption As Variant
    lookupPaymentOption = Application.VLookup(Target.value, Worksheets("TNLU").Range("A:C"), 3, False)
    
    If IsError(lookupPaymentOption) Then
        Application.Cells(Target.row, PAYMENT_OPTION).ClearContents
    Else
        Application.Cells(Target.row, PAYMENT_OPTION) = lookupPaymentOption
    End If
End Sub

    ' *** AvgasCalc ***
Private Sub AvgasCalc(ByVal Target As Range)
    If Not IsEmpty(Application.Cells(Target.row, AVGAS_METER_START)) And Not IsEmpty(Application.Cells(Target.row, AVGAS_METER_STOP)) Then
        If Application.Cells(Target.row, AVGAS_METER_START).value > 0 And Application.Cells(Target.row, AVGAS_METER_STOP) > 0 Then
            Application.Cells(Target.row, AVGAS_METER_AUTO) = Round((Application.Cells(Target.row, AVGAS_METER_STOP).value - Application.Cells(Target.row, AVGAS_METER_START).value), 1)
            Application.Cells(Target.row, AVGAS_METER_DIFF) = Round((Application.Cells(Target.row, AVGAS_METER_AUTO).value - Application.Cells(Target.row, AVGAS_METER_MANUAL).value), 1)
        End If
    End If
End Sub

    ' *** JetCalc ***
Private Sub JetCalc(ByVal Target As Range)
    If Not IsEmpty(Application.Cells(Target.row, JET_METER_START)) And Not IsEmpty(Application.Cells(Target.row, JET_METER_STOP)) Then
        If Application.Cells(Target.row, JET_METER_STOP).value > 0 And Application.Cells(Target.row, JET_METER_STOP).value > 0 Then
            Application.Cells(Target.row, JET_METER_AUTO) = Round((Application.Cells(Target.row, JET_METER_STOP).value - Application.Cells(Target.row, JET_METER_START).value), 0)
            Application.Cells(Target.row, JET_METER_DIFF) = Round((Application.Cells(Target.row, JET_METER_AUTO).value - Application.Cells(Target.row, JET_METER_MANUAL).value), 0)
        End If
    End If
End Sub

    ' *** Price_Input ***
Private Sub Price_Input(ByVal Target As Range, fuelType As String)
    If fuelType = "AVGAS" Then
        If Not IsEmpty(Application.Cells(Target.row, PAYMENT_OPTION)) And Application.Cells(Target.row, AVGAS_METER_AUTO) > 0 Then
            If (Application.Cells(Target.row, PAYMENT_OPTION).value = 3) Or (Application.Cells(Target.row, PAYMENT_OPTION).value = 0) Then
                Application.Cells(Target.row, FUEL_PRICE) = Sheets("TNLU").Range("I4").value
            Else
                Application.Cells(Target.row, FUEL_PRICE) = Sheets("TNLU").Range("H4").value
            End If
        End If
    ElseIf fuelType = "JET" Then
        If Not IsEmpty(Application.Cells(Target.row, PAYMENT_OPTION)) And Application.Cells(Target.row, JET_METER_AUTO) > 0 Then
            Dim lookupTenant As Variant
            lookupTenant = Application.VLookup(Application.Cells(Target.row, TAIL_NUMBER).value, Worksheets("TNLU").Range("A:D"), 4, False)
            
            If IsError(lookupTenant) Then
                Application.Cells(Target.row, FUEL_PRICE) = Sheets("TNLU").Range("J4").value
            Else
                If lookupTenant = 1 Then
                    Application.Cells(Target.row, FUEL_PRICE) = Sheets("TNLU").Range("L4").value
                Else
                    Application.Cells(Target.row, FUEL_PRICE) = Sheets("TNLU").Range("J4").value
                End If
            End If
        End If
    End If
End Sub

    ' *** CalcTotal ***
Private Sub CalcTotal(ByVal Target, meter As Variant, method As Integer)
    If method = 1 Then
        Cells(Target.row, CASH_AMOUNT) = Cells(Target.row, FUEL_PRICE).value * meter
    ElseIf method = 2 Then
        Cells(Target.row, CHECK_AMOUNT) = Cells(Target.row, FUEL_PRICE).value * meter
    ElseIf method = 3 Then
        Cells(Target.row, CREDIT_AMOUNT) = Cells(Target.row, FUEL_PRICE).value * meter
    ElseIf method = 4 Then
        Cells(Target.row, TAB_AMOUNT) = Cells(Target.row, FUEL_PRICE).value * meter
    End If
End Sub

    ' *** ClearTotal ***
Private Sub ClearTotals(ByVal Target)
    Dim col As Integer
    
    For col = CASH_AMOUNT To TAB_AMOUNT
        Cells(Target.row, col).ClearContents
    Next col
End Sub
