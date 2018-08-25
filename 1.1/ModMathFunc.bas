Attribute VB_Name = "ModMathFunc"
Public Function Ceil(n As Double)
    Ceil = -(CInt(-n))
End Function

Public Function Min(a As Integer, b As Integer)
    Min = IIf(a < b, a, b)
End Function

Public Function Max(a As Integer, b As Integer)
    Max = IIf(a > b, a, b)
End Function
