'
' 使い方；
' (1) セルAからEまで入力して、
' (2) makeAddBookmarkUrl()を実行。
' (3) ワークシートをウェブページとして保存して、
' (4) 保存したウェブページをブラウザで開く。
' (5) 登録！リンクを新しいウィンドウで開いて、アノテーションをサブミット。

Sub makeAddBookmarkUrl()
    Dim baseUrl, target As String
    Dim maxRow, i As Integer
    maxRow = Range("C1").End(xlDown).Row ' 表の最終行
    For i = 2 To maxRow
        target = AddBookmarkUrl(Cells(i, "A").Value, Cells(i, "B").Value, Cells(i, "C").Value, Cells(i, "D").Value, Cells(i, "E").Value)
        Cells(i, "G").Value = target
        Cells(i, "F").Value = "=HYPERLINK(""" + target + """, ""登録！"")"
    Next i
End Sub

Function AddBookmarkUrl(dataset As String, genome As String, gene As String, tags As String, comment As String) As String
    Dim targetUrl, commentStr, tagStr As String
    tagStr = Replace(tags, ", ", "][")
    tagStr = Replace(tags, ",", "][")
    tagStr = Replace(tagStr, "][ ", "][")
    commentStr = "[" + tagStr + "] " + comment
    targetUrl = genomeUrl(dataset, genome, gene)
    AddBookmarkUrl = "http://a.kazusa.or.jp/bookmarks/add?comment=" + commentStr + "&uri=" + targetUrl
End Function

Function genomeUrl(dataset As String, genome As String, gene As String) As String
    genomeUrl = "http://genome.kazusa.or.jp/" + dataset + "/" + genome + "/genes/" + gene
End Function

Sub clearLinks()
    Dim maxRow, i As Integer
    maxRow = Range("C1").End(xlDown).Row
    For i = 2 To maxRow
        Cells(i, "F").Value = ""
        Cells(i, "G").Value = ""
    Next i
End Sub

Sub reset()
    Dim maxRow, i As Integer
    maxRow = Range("C1").End(xlDown).Row
    Range(Cells(2, "A"), Cells(maxRow, "G")).ClearContents
    Range("A2").Select
End Sub

Sub sampleInput()
    Cells(2, "A").Value = "cyanobase"
    Cells(2, "B").Value = "Synechocystis"
    Cells(2, "C").Value = "slr1234"
    Cells(2, "D").Value = "test,gnam:test"
    Cells(2, "E").Value = "tesuto"
End Sub

