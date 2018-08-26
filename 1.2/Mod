Attribute VB_Name = "ModAnalyzeRegExp"
Public Function AnalyzeRegExp(ByVal pat As String, ByVal txt As String)
    Dim mRegExp As RegExp
    Dim mMatches As MatchCollection   '匹配字符串集合对象
    Dim mMatch As Match     '匹配字符串
    Set mRegExp = New RegExp
    With mRegExp
        .Global = True      'True表示匹配所有, False表示仅匹配第一个符合项
        .IgnoreCase = False     'True表示不区分大小写, False表示区分大小写
        .Pattern = pat   '匹配字符模式
        Set mMatches = .Execute(txt)    '执行正则查找，返回所有匹配结果的集合，若未找到，则为空
    End With
    Set AnalyzeRegExp = mMatches
End Function
