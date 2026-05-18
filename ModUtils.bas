Attribute VB_Name = "ModUtils"
Option Explicit

'Function CollectionFromArray(ParamArray Values()) As Collection
'    Dim Coll As Collection
'    Set Coll = New Collection
'
'End Function

Function ArrayStr(ParamArray Args()) As String()

    Dim Arr() As String
    ReDim Arr(UBound(Args))
    
    Dim Idx As Long
    For Idx = LBound(Args) To UBound(Args)
        Arr(Idx) = Args(Idx)
    Next
    
    ArrayStr = Arr ' Return
End Function
