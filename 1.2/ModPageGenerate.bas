Attribute VB_Name = "ModPageGenerate"
Public Function PageGenerate(st As String)
    If st = "" Then
        PageGenerate = Array("1")
        Exit Function
    End If
    PageGenerate = Split(st, ",")
End Function
