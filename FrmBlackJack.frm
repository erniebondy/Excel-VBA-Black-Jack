VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmBlackJack 
   Caption         =   "Black Jack"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   OleObjectBlob   =   "FrmBlackJack.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmBlackJack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Deck As Deck
Dim Player As Player
Dim Dealer As Player

Private Sub BtnDeal_Click()
    Set Deck = New Deck
    Deck.Make
    Deck.Shuffle
    
    Set Player = New Player
    Set Dealer = New Player
    
    Deck.Deal Player
    Deck.Deal Dealer
    Deck.Deal Player
    Deck.Deal Dealer, , False
    
    Call WriteCards(Dealer.Cards, LblDealerCards)
    Call WriteCards(Player.Cards, LblPlayerCards)
    
    If Not Turn(Player) Then
        Exit Sub
    End If
    
    Call UpdateDeckLabel
    Call UpdateValueLabels
End Sub

Private Sub BtnHit_Click()
    Deck.Deal Player
    Call WriteCards(Player.Cards, LblPlayerCards)
    Call UpdateDeckLabel
    Call UpdateValueLabels
    If Not Turn(Player) Then
        Exit Sub
    End If
End Sub

Private Sub BtnStand_Click()
    Dealer.RevealCards
    Do
        Deck.Deal Dealer
        Call WriteCards(Dealer.Cards, LblDealerCards)
        Call UpdateDeckLabel
        Call UpdateValueLabels
    Loop While Turn(Dealer)
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RevealBoard()
    Dealer.RevealCards
    Call WriteCards(Dealer.Cards, LblDealerCards)
    Call UpdateValueLabels
End Sub

Private Function Turn(ActivePlayer As Player) As Boolean
    
    Dim PlayerValue As Long
    PlayerValue = GetTotalCardsValue(Player.Cards)
    
    Dim DealerValue As Long
    DealerValue = GetTotalCardsValue(Dealer.Cards)
    
    If ActivePlayer Is Player Then
        If PlayerValue = 21 Then
            Call RevealBoard
            MsgBox "Player WINS!"
            Turn = False: Exit Function ' Return
        End If
        
        If PlayerValue > 21 Then
            Call RevealBoard
            MsgBox "Player LOSES!"
            Turn = False: Exit Function ' Return
        End If
    End If
    
    If ActivePlayer Is Dealer Then
        Call RevealBoard
        If PlayerValue = 21 And DealerValue = 21 Then
            MsgBox "PUSH!"
            Turn = False: Exit Function ' Return
        End If
        
        If DealerValue = 21 Then
            MsgBox "Player LOSES!"
            Turn = False: Exit Function ' Return
        End If
        
        If DealerValue > 21 Then
            MsgBox "Player WINS!"
            Turn = False: Exit Function ' Return
        End If
        
        If DealerValue > PlayerValue Then
            MsgBox "Player LOSES!"
            Turn = False: Exit Function ' Return
        End If
    End If
    
    Turn = True
End Function

Private Sub UpdateValueLabels()
    LblPlayerValue.Caption = "(" & GetTotalCardsValue(Player.Cards) & ")"
    LblDealerValue.Caption = "(" & GetTotalCardsValue(Dealer.Cards) & ")"
End Sub

Private Function GetTotalCardsValue(Cards As Deck) As Long
    Dim Sum As Long
    Dim Card As Card
    For Each Card In Cards.Cards
        Sum = Sum + GetBlackJackValue(Card)
    Next
    
    ' Reduce Ace(s) to 1
    If Sum > 21 Then
        For Each Card In Cards.Cards
            If Card.Name = "Ace" Then
                Sum = Sum - 11 + 1
                If Sum <= 21 Then
                    Exit For
                End If
            End If
        Next
    End If
    
    GetTotalCardsValue = Sum
End Function

Private Function GetBlackJackValue(Card As Card) As Integer
    If Not Card.Visible Then
        GetBlackJackValue = 0: Exit Function ' Return
    End If
    
    If Card.Name = "Ace" Then
        GetBlackJackValue = 11: Exit Function ' Return
    End If
    
    GetBlackJackValue = IIf(Card.CardType = enFace, 10, Card.Value)
End Function

Private Sub WriteCards(Cards As Deck, Lbl As MSForms.Label)
    Dim Card As Card
    Dim CardsArr() As String
    ReDim CardsArr(Cards.Cards.Count - 1)
    Dim Idx As Long
    Idx = 0
    Dim Value As Integer
    For Each Card In Cards.Cards
        Value = GetBlackJackValue(Card)
        CardsArr(Idx) = IIf(Card.Visible, Card.Name & " of " & Card.Suit.ToString & " (" & Value & ")", "###")
        Idx = Idx + 1
    Next
    Lbl.Caption = Join(CardsArr, vbNewLine)
End Sub

Private Sub UpdateDeckLabel()
    LblDeck.Caption = Deck.Cards.Count & " / 52"
End Sub
