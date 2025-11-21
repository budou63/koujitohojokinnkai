Option Explicit

Public Function NormalizeAccountName(ByVal inputText As String) As String
    Dim s As String
    s = CStr(inputText)
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")

    Dim originalHasSpace As Boolean
    originalHasSpace = (InStr(s, " ") > 0 Or InStr(s, "　") > 0)

    s = ToKatakanaWide(s)
    s = ApplyAbbreviations(s)
    s = ConvertMarks(s)
    s = StrConv(s, vbNarrow)
    s = ExpandSmallKana(s)
    s = NormalizeSpaces(s, originalHasSpace)

    NormalizeAccountName = s
End Function

Private Function ToKatakanaWide(ByVal s As String) As String
    s = StrConv(s, vbKatakana)
    s = StrConv(s, vbWide)
    ToKatakanaWide = s
End Function

Private Function ApplyAbbreviations(ByVal s As String) As String
    Dim pairs As Collection
    Set pairs = New Collection

    pairs.Add Array("（株）", "ｶ)")
    pairs.Add Array("(株)", "ｶ)")
    pairs.Add Array("㈱", "ｶ)")
    pairs.Add Array("株式会社", "ｶ)")
    pairs.Add Array("（有）", "ﾕ)")
    pairs.Add Array("(有)", "ﾕ)")
    pairs.Add Array("㈲", "ﾕ)")
    pairs.Add Array("有限会社", "ﾕ)")
    pairs.Add Array("（同）", "ﾄﾞ)")
    pairs.Add Array("合同会社", "ﾄﾞ)")
    pairs.Add Array("合名会社", "ｺﾞﾒｲ)")
    pairs.Add Array("合資会社", "ｺﾞｼ)")
    pairs.Add Array("学校法人", "ｶﾞｸ)")
    pairs.Add Array("医療法人", "ｲﾘ)")
    pairs.Add Array("社会福祉法人", "ｼﾔ)")
    pairs.Add Array("宗教法人", "ｼﾕｳ)")
    pairs.Add Array("一般社団法人", "ｼﾔﾀﾞﾝ)")
    pairs.Add Array("公益社団法人", "ｺｳｴｷｼﾔ)")
    pairs.Add Array("一般財団法人", "ｻﾞｲ)")
    pairs.Add Array("公益財団法人", "ｺｳｴｷｻﾞ)")
    pairs.Add Array("特定非営利活動法人", "NPO)")
    pairs.Add Array("ＮＰＯ法人", "NPO)")
    pairs.Add Array("NPO法人", "NPO)")
    pairs.Add Array("農業協同組合", "ﾉｳｷﾖｳ)")
    pairs.Add Array("農協", "ﾉｳｷﾖｳ)")
    pairs.Add Array("農業共済組合", "ﾉｳｷﾖｳｻｲ)")
    pairs.Add Array("信用金庫", "ｼﾝｷﾝ)")
    pairs.Add Array("信金", "ｼﾝｷﾝ)")
    pairs.Add Array("信用組合", "ｼﾝｸﾐ)")
    pairs.Add Array("信組", "ｼﾝｸﾐ)")
    pairs.Add Array("労働金庫", "ﾛｳｷﾝ)")
    pairs.Add Array("労金", "ﾛｳｷﾝ)")
    pairs.Add Array("漁業協同組合", "ｷﾞﾖｷﾖｳ)")
    pairs.Add Array("漁協", "ｷﾞﾖｷﾖｳ)")
    pairs.Add Array("生活協同組合", "ｾｲｷﾖｳ)")
    pairs.Add Array("生協", "ｾｲｷﾖｳ)")
    pairs.Add Array("森林組合", "ｼﾝﾘﾝ)")
    pairs.Add Array("協同組合", "ｷﾖｳﾄﾞｳ)")
    pairs.Add Array("事業協同組合", "ｼﾞｷﾞﾖｳ)")
    pairs.Add Array("商工会", "ｼﾖｳｺｳｶｲ)")
    pairs.Add Array("商工会議所", "ｼﾖｳｺｳｶｲ)")
    pairs.Add Array("振興会", "ｼﾝｺｳｶｲ)")
    pairs.Add Array("営業所", "ｴｲ)")

    Dim itm As Variant
    For Each itm In pairs
        s = Replace(s, itm(0), itm(1))
    Next itm
    ApplyAbbreviations = s
End Function

Private Function ConvertMarks(ByVal s As String) As String
    s = Replace(s, "ｰ", "-")
    s = Replace(s, "－", "-")
    s = Replace(s, "ー", "-")
    s = Replace(s, "・", ".")
    s = Replace(s, "･", ".")
    ConvertMarks = s
End Function

Private Function ExpandSmallKana(ByVal s As String) As String
    s = Replace(s, "ｧ", "ｱ")
    s = Replace(s, "ｨ", "ｲ")
    s = Replace(s, "ｩ", "ｳ")
    s = Replace(s, "ｪ", "ｴ")
    s = Replace(s, "ｫ", "ｵ")
    s = Replace(s, "ｯ", "ﾂ")
    s = Replace(s, "ｬ", "ﾔ")
    s = Replace(s, "ｭ", "ﾕ")
    s = Replace(s, "ｮ", "ﾖ")
    s = Replace(s, "ｧﾟ", "ｱﾟ")
    s = Replace(s, "ｩﾟ", "ｳﾟ")
    s = Replace(s, "ｧﾞ", "ｱﾞ")
    ExpandSmallKana = s
End Function

Private Function NormalizeSpaces(ByVal s As String, ByVal originalHasSpace As Boolean) As String
    s = Replace(s, "　", " ")
    s = Replace(s, vbTab, " ")
    While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Wend
    s = Trim(s)

    Dim isCorp As Boolean
    isCorp = IsCorporateName(s)

    If isCorp Then
        s = Replace(s, " ", "")
    Else
        If originalHasSpace Then
            s = Replace(s, " ", " ")
        Else
            s = Replace(s, " ", "")
        End If
    End If

    NormalizeSpaces = s
End Function

Private Function IsCorporateName(ByVal s As String) As Boolean
    Dim keywords As Variant
    keywords = Array("ｶ)", "ﾕ)", "ﾄﾞ)", "ｺﾞﾒｲ)", "ｺﾞｼ)", "ｶﾞｸ)", "ｲﾘ)", "ｼﾔ)", "ｼﾕｳ)", "ｼﾔﾀﾞﾝ)", "ｺｳｴｷｼﾔ)", "ｻﾞｲ)", "ｺｳｴｷｻﾞ)", "NPO)", "ﾉｳｷﾖｳ)", "ｼﾝｷﾝ)", "ｼﾝｸﾐ)", "ﾛｳｷﾝ)", "ｷﾞﾖｷﾖｳ)", "ｾｲｷﾖｳ)", "ｷﾖｳﾄﾞｳ)", "ｼﾞｷﾞﾖｳ)", "ｼﾖｳｺｳｶｲ)", "ｼﾝｺｳｶｲ)", "ｴｲ)")

    Dim k As Variant
    For Each k In keywords
        If InStr(s, CStr(k)) > 0 Then
            IsCorporateName = True
            Exit Function
        End If
    Next k
    If InStr(s, "法人") > 0 Or InStr(s, "組合") > 0 Then
        IsCorporateName = True
    End If
End Function
