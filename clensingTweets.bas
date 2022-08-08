Option Explicit

Sub clensingTweets()
'ツイートはここから取得：https://www.allmytweets.net/
'正規表現を利用した整形
'以下条件が含まれるセルを空白にし、最後に空白セルだけ削除
'条件：文頭に＠かRTがある、「@YouTubeより」という文字列がある
'URLとハッシュタグと数字が含まれる場合は、それらのみを削除（ただしここの部分は精度が低い…）

Application.ScreenUpdating = False
    
    Dim row As Long: row = 2175
    Dim i As Long, j As Long
    Dim arrPat As Variant: arrPat = Array("^@.*", "^RT", "@YouTubeより")

    Dim strText As String
    Dim strPattern As String

    For i = LBound(arrPat) To UBound(arrPat) 'パターンごと
        For j = 1 To row '行ごと
            strPattern = arrPat(i)
            strText = Cells(j, 1).Value
            If hasPattern(strPattern, strText) = True Then
                Cells(j, 1).Value = ""
            End If
        Next j
    Next i

    arrPat = Array("\d", "#[^#\s]*", "http(s)?://([\w-]+\.)+[\w-]+(/[\w- ./?%&=]*)?")
    
    For i = LBound(arrPat) To UBound(arrPat) 'パターンごと
        For j = 1 To row
            strPattern = arrPat(i)
            strText = Cells(j, 1).Value
            Cells(j, 1).Value = removePattern(strPattern, strText)
        Next j
        Range("A1:A2175").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    Next i

    
    Range("A1:A2175").SpecialCells(xlCellTypeBlanks).EntireRow.Delete

    MsgBox "完了"
    
    Application.ScreenUpdating = True
End Sub


Function hasPattern(ByVal strPattern As String, ByVal strText As String)
'パターンにマッチしたらTrueを返す

    'RegExpを使えるようにオブジェクト宣言
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")

        With reg
            .Pattern = strPattern
            .Global = True '文字列全体を見る

            If .Test(strText) = True Then
                hasPattern = True
            Else
                hasPattern = False
            End If
        End With

End Function

Function removePattern(ByVal strPattern As String, ByVal strText As String)
'パターンにマッチするものがあったら削除する
    'RegExpを使えるようにオブジェクト宣言
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")

        With reg
            .Pattern = strPattern
            .IgnoreCase = True '大文字と小文字を区別しない
            .Global = True '文字列全体を見る
            removePattern = .Replace(strText, "")
        End With
End Function
