VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTicketEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ***** CTicketEntry *****
' * - Class-Module object for Fuel Sales spreadsheet
' * - AUTHOR: Bryan Huss, 8/15/18
' *********************

' VARIABLES
Private dPurchaseDate As Date
Private sTicketNum As String, _
            sTailNum As String, _
            sName As String, _
            sNFPT As String, _
            sComments As String
Private dAVGASMeterStart As Double, _
            dAVGASMeterStop As Double, _
            dAVGASMeterDiffManual As Double, _
            dAVGASMeterDiffAuto As Double, _
            dAVGASDiffDiff As Double, _
            dJETMeterStart As Double, _
            dJETMeterStop As Double, _
            dJETMeterDiffManual As Double, _
            dJETMeterDiffAuto As Double, _
            dJETDiffDiff As Double, _
            dFuelPPG As Double, _
            dAmountPaid As Double
Private bPayCode As Byte

' CONSTRUCTOR (of sorts)
Public Sub TicketEntry(TicketNum As String, _
                                    PurchaseDate As Date, _
                                    TailNum As String, _
                                    name As String, _
                                    AVGASMeterStart As Double, _
                                    AVGASMeterStop As Double, _
                                    AVGASMeterDiffManual As Double, _
                                    AVGASMeterDiffAuto As Double, _
                                    AVGASDiffDiff As Double, _
                                    JETMeterStart As Double, _
                                    JETMeterStop As Double, _
                                    JETMeterDiffManual As Double, _
                                    JETMeterDiffAuto As Double, _
                                    JETDiffDiff As Double, _
                                    FuelPPG As Double, _
                                    NFPT As String, _
                                    PayCode As Byte, _
                                    AmountPaid As Double, _
                                    Comments As String)

    sTicketNum = TicketNum
    dPurchaseDate = PurchaseDate
    sTailNum = TailNum
    sName = name
    dAVGASMeterStart = AVGASMeterStart
    dAVGASMeterStop = AVGASMeterStop
    dAVGASMeterDiffManual = AVGASMeterDiffManual
    dAVGASMeterDiffAuto = AVGASMeterDiffAuto
    dAVGASDiffDiff = AVGASDiffDiff
    dJETMeterStart = JETMeterStart
    dJETMeterStop = JETMeterStop
    dJETMeterDiffManual = JETMeterDiffManual
    dJETMeterDiffAuto = JETMeterDiffAuto
    dJETDiffDiff = JETDiffDiff
    dFuelPPG = FuelPPG
    sNFPT = NFPT
    bPayCode = PayCode
    dAmountPaid = AmountPaid
    sComments = Comments
End Sub

' GET/LET METHODS
' Ticket Number
Public Property Get TicketNum() As String
    TicketNum = sTicketNum
End Property

Public Property Let TicketNum(value As String)
    sTicketNum = value
End Property

' Puchase Date
Public Property Get PurchaseDate() As Date
    PurchaseDate = dPurchaseDate
End Property

Public Property Let PurchaseDate(value As Date)
    dPurchaseDate = value
End Property

' Tail Number
Public Property Get TailNum() As String
    TailNum = sTailNum
End Property

Public Property Let TailNum(value As String)
    sTailNum = value
End Property

' Name
Public Property Get name() As String
    name = sName
End Property

Public Property Let name(value As String)
    sName = value
End Property

' Non-Fuel Purchase Ticket
Public Property Get NFPT() As String
    NFPT = sNFPT
End Property

Public Property Let NFPT(value As String)
    sNFPT = value
End Property

' Comments
Public Property Get Comments() As String
    Comments = sComments
End Property

Public Property Let Comments(value As String)
    sComments = value
End Property

' AVGAS Meter Start
Public Property Get AVGASMeterStart() As Double
    AVGASMeterStart = dAVGASMeterStart
End Property

Public Property Let AVGASMeterStart(value As Double)
    dAVGASMeterStart = value
End Property

' AVGAS Meter Stop
Public Property Get AVGASMeterStop() As Double
    AVGASMeterStop = dAVGASMeterStop
End Property

Public Property Let AVGASMeterStop(value As Double)
    dAVGASMeterStop = value
End Property

' AVGAS Meter Difference Manual
Public Property Get AVGASMeterDiffManual() As Double
    AVGASMeterDiffManual = dAVGASMeterDiffManual
End Property

Public Property Let AVGASMeterDiffManual(value As Double)
    dAVGASMeterManual = value
End Property

' AVGAS Meter Difference Auto
Public Property Get AVGASMeterDiffAuto() As Double
    AVGASMeterDiffAuto = dAVGASMeterDiffAuto
End Property

Public Property Let AVGASMeterDiffAuto(value As Double)
    dAVGASMeterDiffAuto = value
End Property

' AVGAS Difference Difference
Public Property Get AVGASDiffDiff() As Double
    AVGASDiffDiff = dAVGASDiffDiff
End Property

Public Property Let AVGASDiffDiff(value As Double)
    dAVGASDiffDiff = value
End Property

' JET Meter Start
Public Property Get JETMeterStart() As Double
    JETMeterStart = dJETMeterStart
End Property

Public Property Let JETMeterStart(value As Double)
    dJETMeterStart = value
End Property

' JET Meter Stop
Public Property Get JETMeterStop() As Double
    JETMeterStop = dJETMeterStop
End Property

Public Property Let JETMeterStop(value As Double)
    dJETMeterStop = value
End Property

' JET Meter Difference Manual
Public Property Get JETMeterDiffManual() As Double
    JETMeterDiffManual = dJETMeterDiffManual
End Property

Public Property Let JETMeterDiffManual(value As Double)
    dJETMeterDiffManual = value
End Property

' JET Meter Difference Auto
Public Property Get JETMeterDiffAuto() As Double
    JETMeterDiffAuto = dJETMeterDiffAuto
End Property

Public Property Let JETMeterDiffAuto(value As Double)
    dJETMeterDiffAuto = value
End Property

' JET Difference Difference
Public Property Get JETDiffDiff() As Double
    JETDiffDiff = dJETDiffDiff
End Property

Public Property Let JETDiffDiff(value As Double)
    dJETDiffDiff = value
End Property

' Fuel Price (per gallon)
Public Property Get FuelPPG() As Double
    FuelPPG = dFuelPPG
End Property

Public Property Let FuelPPG(value As Double)
    dFuelPPG = value
End Property

' Amount Paid Total
Public Property Get AmountPaid() As Double
    AmountPaid = dAmountPaid
End Property

Public Property Let AmountPaid(value As Double)
    dAmountPaid = value
End Property

' Pay Code
Public Property Get PayCode() As Byte
    PayCode = bPayCode
End Property

Public Property Let PayCode(value As Byte)
    bPayCode = value
End Property

' DEBUG PRINT
Public Sub PrintTicket()
    Debug.Print "Ticket#: " & sTicketNum
    Debug.Print "Purchase Date: " & dPurchaseDate
    Debug.Print "N#: " & sTailNum
    Debug.Print "Name: " & sName
    Debug.Print "AVGAS Start: " & dAVGASMeterStart
    Debug.Print "AVGAS Stop: " & dAVGASMeterStop
    Debug.Print "AVGAS Meter Diff Manual: " & dAVGASMeterDiffManual
    Debug.Print "AVGAS Meter Diff Auto: " & dAVGASMeterDiffAuto
    Debug.Print "AVGAS Diff Diff: " & dAVGASDiffDiff
    Debug.Print "JET Start: " & dJETMeterStart
    Debug.Print "JET Stop: " & dJETMeterStop
    Debug.Print "JET Meter Diff Manual: " & dJETMeterDiffManual
    Debug.Print "JET Meter Diff Auto: " & dJETMeterDiffAuto
    Debug.Print "JET Diff Diff: " & dJETDiffDiff
    Debug.Print "Fuel PPG: " & dFuelPPG
    Debug.Print "NFPT: " & sNFPT
    Debug.Print "Pay Code: " & bPayCode
    Debug.Print "Amount Paid: " & dAmountPaid
    Debug.Print "Comments: " & sComments
End Sub

' TOSTRING
Public Property Get ToString()
    Dim toS As String
    toS = "T#: " & sTicketNum & _
              ", PD: " & dPurchaseDate & _
              ", N#: " & sTailNum & _
              ", Name: " & sName & _
              ", AStart: " & dAVGASMeterStart & _
              ", AStop: " & dAVGASMeterStop & _
              ", AMDM: " & dAVGASMeterDiffManual & _
              ", AMDA: " & dAVGASMeterDiffAuto & _
              ", ADD: " & dAVGASDiffDiff & _
              ", JStart: " & dJETMeterStart & _
              ", JStop: " & dJETMeterStop & _
              ", JMDM: " & dJETMeterDiffManual & _
              ", JMDA: " & dJETMeterDiffAuto & _
              ", JDD: " & dJETDiffDiff & _
              ", PPG: $" & dFuelPPG & _
              ", NFPT: " & sNFPT & _
              ", PC: " & bPayCode & _
              ", Paid: $" & dAmountPaid & _
              ", Com: " & sComments
              
    ToString = toS
End Property
