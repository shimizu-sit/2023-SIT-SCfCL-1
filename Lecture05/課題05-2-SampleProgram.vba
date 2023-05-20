Option Explicit

Sub Prac05_2_1_Ans()
    ' 変数を準備
    Dim 学籍番号 As String
    
    ' 変数に代入する
    学籍番号 = Range("B3").Value
  
    '学科判定
    If 学籍番号 Like "??A1???" Then
        Range("D3").Value = "機械工学科"
    ElseIf 学籍番号 Like "??A2???" Then
        Range("D3").Value = "電気電子工学科"
    ElseIf 学籍番号 Like "??A3???" Then
        Range("D3").Value = "情報工学科"
    ElseIf 学籍番号 Like "??A6???" Then
        Range("D3").Value = "コンピュータ応用学科"
    ElseIf 学籍番号 Like "??A7???" Then
        Range("D3").Value = "総合デザイン学科"
    ElseIf 学籍番号 Like "??A8???" Then
        Range("D3").Value = "人間環境学科"
    Else
        MsgBox "入力に誤りがあります"
    End If
End Sub

Sub prac05_2_2_Ans()
    ' 変数を準備
    Dim 学籍番号 As String
    
    ' 変数に代入する
    学籍番号 = Range("B3").Value
    
    '学年判定
    If 学籍番号 Like "23A????" Then
        Range("E3").Value = "1年"
    ElseIf 学籍番号 Like "22A????" Then
        Range("E3").Value = "2年"
    ElseIf 学籍番号 Like "21A????" Then
        Range("E3").Value = "3年"
    ElseIf 学籍番号 Like "20A????" Then
        Range("E3").Value = "4年"
    Else
        MsgBox "入力に誤りがあります"
    End If
End Sub