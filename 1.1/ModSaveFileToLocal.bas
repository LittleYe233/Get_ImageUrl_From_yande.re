Attribute VB_Name = "ModSaveFileToLocal"
Public Function SaveFileToLocal(path As String, txt As String)
    Open path For Output As #1
    Write #1, txt
    Close #1
End Function
