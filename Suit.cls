VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Suit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Enum SuitType
    enHeart
    enSpade
    enDiamond
    enClub
End Enum

Dim pSuitType As SuitType
Dim pName As String
Dim pColor As Long
Dim pColorName As String

Sub Copy(From As Suit)
    Me.Make From.ThisType
End Sub

Function ToString() As String
    ToString = pName & "s"
End Function

Sub Make(SuitType As SuitType)
    pSuitType = SuitType
    Call SetClass
End Sub

Private Sub SetClass()
    Select Case pSuitType
        Case SuitType.enHeart
            pName = "Heart"
        Case SuitType.enDiamond
            pName = "Diamond"
        Case SuitType.enSpade
            pName = "Spade"
        Case SuitType.enClub
            pName = "Club"
    End Select
    
    Select Case pSuitType
        Case SuitType.enHeart, SuitType.enDiamond
            pColor = vbRed
            pColorName = "Red"
        Case SuitType.enSpade, SuitType.enClub
            pColor = vbBlack
            pColorName = "Black"
    End Select
End Sub

' Naming the Property 'SuitType' does not cause an issue
' It does not show in the intellisense though
Property Let ThisType(Value As SuitType)
    pSuitType = Value
    Call SetClass
End Property

Property Get ThisType() As SuitType
    ThisType = pSuitType
End Property

Public Property Get Name() As String
    Name = pName
End Property
