Attribute VB_Name = "ModGetUrl"
Public Function GetUrl(Url As String)
    Set xmlHttp = CreateObject("Microsoft.XMLHTTP")
    xmlHttp.Open "GET", Url, True
    xmlHttp.send (Null)
    While xmlHttp.ReadyState <> 4
        DoEvents
    Wend
    GetUrl = xmlHttp.responseText
End Function
