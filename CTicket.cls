VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ''''''''''''''''''''''''''''''''''''''''
' Class CTicket
' Ticket object used when evaluating data from Fuel Sales Summary spreadsheet
' ''''''''''''''''''''''''''''''''''''''''
' ''''''''''''''''''''''
' Attributes
' ''''''''''''''''''''''
Private lTicketNum As Long
Private iPayCode As Integer
Private dPurchaseDate As Date
Private sTail As String, sName As String
Private dAvgas As Double, dJet As Double, dPricePerGal As Double, dPriceTotal As Double

' '''''''''''''''''''''''''''''''''''''''''''''
' Constructor (of sorts)
' '''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init(TicketNum As Long, PurchaseDate As Date, tail As String, name As String, avgas As Double, jet As Double, PayCode As Integer, pricePerGal As Double, priceTotal As Double)
    lTicketNum = TicketNum
    dPurchaseDate = PurchaseDate
    sTail = tail
    sName = name
    dAvgas = avgas
    dJet = jet
    iPayCode = PayCode
    dPricePerGal = pricePerGal
    dPriceTotal = priceTotal
End Sub

' ''''''''''''''''''''''''''''''''''''
' Get/Let Methods
' ''''''''''''''''''''''''''''''''''''
' Ticket Number
Public Property Get TicketNum() As Long
    TicketNum = lTicketNum
End Property

Public Property Let TicketNum(value As Long)
    lTicketNum
End Property

' Pay Code
Public Property Get PayCode() As Integer
    PayCode = iPayCode
End Property

Public Property Let PayCode(value As Integer)
    iPayCode = value
End Property

' Purchase Date
Public Property Get PurchaseDate() As Date
    PurchaseDate = dPurchaseDate
End Property

Public Property Let PurchaseDate(value As Date)
    dPurchaseDate = value
End Property

' Tail #
Public Property Get tail() As String
    tail = sTail
End Property

Public Property Let tail(value As String)
    sTail = value
End Property

' Name
Public Property Get name() As String
    name = sName
End Property

Public Property Let name(value As String)
    sName = value
End Property

' Avgas meter toal
Public Property Get avgas() As Double
    avgas = dAvgas
End Property

Public Property Let avgas(value As Double)
    dAvgas = value
End Property

' Jet meter total
Public Property Get jet() As Double
    jet = dJet
End Property

Public Property Let jet(value As Double)
    dJet = value
End Property

' Price per gallon
Public Property Get pricePerGallon() As Double
    pricePerGallon = dPricePerGal
End Property

Public Property Let pricePerGallon(value As Double)
    dPricePerGal = value
End Property

' Price total
Public Property Get priceTotal() As Double
    priceTotal = dPriceTotal
End Property

Public Property Let priceTotal(value As Double)
    dPriceTotal = value
End Property


' '''''''''''''''''''''''''''''''''''''
' Auxillary Methods
' '''''''''''''''''''''''''''''''''''''
Public Sub PrintTicket()
    Debug.Print "[Ticket#] " & lTicketNum & _
                        ", [Date] " & dPurchaseDate & _
                        ", [Tail#] " & sTail & _
                        ", [Name] " & sName & _
                        ", [AvGas] " & dAvgas & _
                        ", [Jet] " & dJet & _
                        ", [PayCode] " & iPayCode & _
                        ", [PricePerGal] " & dPricePerGal & _
                        ", [PriceTotal] " & dPriceTotal
End Sub
