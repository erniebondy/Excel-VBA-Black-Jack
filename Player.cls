VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Player"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum PlayerState
    enWaiting
    enActive
End Enum

Dim pId As Long
Dim pName As String
Dim pState As PlayerState
Dim pCards As Deck

Sub RevealCards()
    Dim Card As Card
    For Each Card In pCards.Cards
        Card.Visible = True
    Next
End Sub

Private Sub Class_Initialize()
    pId = -1
    pState = enWaiting
    Set pCards = New Deck
End Sub

Public Property Get Cards() As Deck
    Set Cards = pCards
End Property

Public Property Get Id() As Long
    Id = pId
End Property

Public Property Let Name(Value As String)
    pName = Value
End Property

Public Property Get Name() As String
    Name = pName
End Property

Public Property Let PlayerState(Value As PlayerState)
    pState = Value
End Property

Public Property Get PlayerState() As PlayerState
    PlayerState = pState
End Property

