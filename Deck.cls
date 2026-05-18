VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Deck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const MAX_CARD_VALUE As Integer = 13
Const SUITS_COUNT As Integer = 4
Const CARDS_COUNT As Integer = 52

'Dim pCards(CARDS_COUNT - 1) As Card
Dim pCards As Collection

Sub Deal(Target As Player, Optional Amount As Integer = 1, Optional Visible As Boolean = True)
    ' Deal from start of collection (top of deck)
    Dim Card As Card
    
    Dim Count As Integer
    For Count = 1 To Amount
        Set Card = pCards(1)
        Card.Visible = Visible
        pCards.Remove 1
        Target.Cards.Add Card
    Next
End Sub

Sub Add(Value As Card)
    pCards.Add Value
End Sub

Sub Shuffle()

    Dim Idx1 As Integer
    Dim Idx2 As Integer
    Dim SwapCount As Integer
    SwapCount = 10000
    Dim CardsCount As Integer
    CardsCount = pCards.Count
    
    Call Randomize
    
    Dim Counter As Integer
    For Counter = 1 To SwapCount
        Idx1 = Int(Rnd * CARDS_COUNT) + 1
        Idx2 = Int(Rnd * CARDS_COUNT) + 1
            
        Debug.Assert Idx1 > 0 And Idx1 <= CardsCount
        Debug.Assert Idx2 > 0 And Idx2 <= CardsCount

        Dim Card1 As Card
        Set Card1 = pCards(Idx1)
        Dim Card2 As Card
        Set Card2 = pCards(Idx2)
'        Dim Temp As Card
'        'Set Temp = pCards(Idx2)
'        Set Temp = pCards(Idx2)
'
'        Set Card2 = Card1
'        Set Card1 = Temp
        
        pCards.Remove Idx1
        If Idx1 = 1 Then
            pCards.Add Card2, , 1
        Else
            pCards.Add Card2, , , (Idx1 - 1)
        End If
        
        pCards.Remove Idx2
        If Idx2 = 1 Then
            pCards.Add Card1, , 1
        Else
            pCards.Add Card1, , , (Idx2 - 1)
        End If
'        Set pCards(Idx2) = pCards(Idx1)
'        Set pCards(Idx1) = Temp
    Next
    
End Sub

Function ToString() As String
    Dim Card As Card
    Dim Idx As Long
    Idx = 1
    Dim Cards() As String
    ReDim Cards(pCards.Count - 1)
    For Each Card In pCards
        Cards(Idx - 1) = Card.ToString
        Idx = Idx + 1
    Next
    ToString = Join(Cards, ",")
End Function

Sub Make()
    Dim Value As Integer
    Dim SuitValue As Integer
    Dim Suit As Suit
    For SuitValue = 0 To SUITS_COUNT - 1
        Set Suit = New Suit
        Suit.ThisType = SuitValue
        For Value = 1 To MAX_CARD_VALUE
'            Dim Idx As Integer
'            Idx = ((Value - 1) + (MAX_CARD_VALUE * SuitValue))
'            Set pCards(Idx) = New Card
'            pCards(Idx).MakeFromValue Value, Suit
            Dim Card As Card
            Set Card = New Card
            Card.MakeFromValue Value, Suit
            pCards.Add Card, Card.Name & "_" & Suit.Name
        Next
    Next
End Sub

Private Sub Class_Initialize()
    Set pCards = New Collection
End Sub

Private Sub Class_Terminate()
'
End Sub

Public Property Get Cards() As Collection
    Set Cards = pCards
End Property
