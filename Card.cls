VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Card"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Enum CardType
    enPip
    enFace
End Enum

Dim pName As String
Dim pValue As Integer
Dim pSuit As Suit
Dim pCardType As CardType
Dim pVisible As Boolean

Dim pCardValues As Collection
Dim pCardNames() As String

Sub MakeFromName(Name As String, Suit As Suit)
'    If Trim(Name) = vbNullString And Value = 0 Then
'        Err.Raise vbObjectError + 100, "Card.Make()", "No Name or Value was provided!"
'        Exit Sub
'    End If
    Call Make(Name, GetValueFromName(Name), Suit)
End Sub

Sub MakeFromValue(Value As Integer, Suit As Suit)
    Call Make(GetNameFromValue(Value), Value, Suit)
End Sub

Function GetValueFromName(Name As String) As Integer
    GetValueFromName = pCardValues(Name)
End Function

Function GetNameFromValue(Value As Integer) As String
    GetNameFromValue = pCardNames(Value - 1)
End Function

Private Sub Make(Name As String, Value As Integer, Suit As Suit)
    pName = Name
    pValue = Value
    Set pSuit = Suit
    If Value > 10 Then
        pCardType = enFace
    Else
        pCardType = enPip
    End If
End Sub

Private Sub Class_Initialize()
    Set pCardValues = New Collection
    pCardValues.Add 1, "Ace"
    pCardValues.Add 2, "Two"
    pCardValues.Add 3, "Three"
    pCardValues.Add 4, "Four"
    pCardValues.Add 5, "Five"
    pCardValues.Add 6, "Six"
    pCardValues.Add 7, "Seven"
    pCardValues.Add 8, "Eight"
    pCardValues.Add 9, "Nine"
    pCardValues.Add 10, "Ten"
    pCardValues.Add 11, "Jack"
    pCardValues.Add 12, "Queen"
    pCardValues.Add 13, "King"
    
    pCardNames = ModUtils.ArrayStr("Ace", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Jack", "Queen", "King")
End Sub

Private Sub Class_Terminate()
    Set pCardValues = Nothing
End Sub

Public Function ToString() As String
    ToString = IIf(Visible, pName & " of " & pSuit.ToString & " (" & pValue & ")", "~XXX~")
End Function

Public Property Let Name(Value As String)
    pName = Value
End Property

Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Value(Value As Integer)
    pValue = Value
End Property

Public Property Get Value() As Integer
    Value = pValue
End Property

Public Property Set Suit(Value As Suit)
    Set pSuit = Value
End Property

Public Property Get Suit() As Suit
    Set Suit = pSuit
End Property

Public Property Get CardType() As CardType
    CardType = pCardType
End Property

Public Property Let Visible(Value As Boolean)
    pVisible = Value
End Property

Public Property Get Visible() As Boolean
    Visible = pVisible
End Property
