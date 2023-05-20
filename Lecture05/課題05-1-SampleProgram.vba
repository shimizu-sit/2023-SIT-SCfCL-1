Option Explicit

Sub Prac05_1_1_Ans()
    ' 変数を準備
    Dim 点数 As Long
    Dim 評価 As String
    
    ' 変数「点数」に点数を代入
    点数 = Range("C4").Value
    
    ' 科目Aの評価
    If 点数 >= 90 Then
        Range("D4").Value = "S"
    ElseIf 点数 >= 80 And 点数 <= 89 Then
        Range("D4").Value = "A"
    ElseIf 点数 >= 70 And 点数 <= 79 Then
        Range("D4").Value = "B"
    ElseIf 点数 >= 60 And 点数 <= 69 Then
        Range("D4").Value = "C"
    ElseIf 点数 >= 0 And 点数 <= 59 Then
        Range("D4").Value = "D"
    Else
        MsgBox "入力値が不正です"
    End If
    
    ' 変数「評価」に評価を代入
    評価 = Range("D4").Value
    
    ' 科目Aの合否
    If 評価 = "S" Or 評価 = "A" Or 評価 = "b" Or 評価 = "C" Then
        Range("E4").Value = "合格"
    ElseIf 評価 = "D" Then
        Range("E4").Value = "不合格"
    Else
        MsgBox "評価値が不正です"
    End If
End Sub

Sub Prac05_1_2_Ans()
    ' 変数の準備
    Dim 点数 As Long
    Dim 評価 As String

    ' 変数「点数」に点数を代入
    点数 = Range("F4").Value
    
    ' 科目Bの評価
    If 点数 >= 90 Then
        Range("G4").Value = "S"
    ElseIf 点数 >= 80 And 点数 <= 89 Then
        Range("G4").Value = "A"
    ElseIf 点数 >= 70 And 点数 <= 79 Then
        Range("G4").Value = "B"
    ElseIf 点数 >= 60 And 点数 <= 69 Then
        Range("G4").Value = "C"
    ElseIf 点数 >= 0 And 点数 <= 59 Then
        Range("G4").Value = "D"
    Else
        MsgBox "入力値が不正です"
    End If
    
    ' 変数「評価」に評価を代入
    評価 = Range("G4").Value
    
    ' 科目Bの合否
    If 評価 = "S" Or 評価 = "A" Or 評価 = "B" Or 評価 = "C" Then
        Range("H4").Value = "合格"
    ElseIf 評価 = "D" Then
        Range("H4").Value = "不合格"
    Else
        MsgBox "評価値が不正です．"
    End If
End Sub

Sub Prac05_1_3_Ans()
    If Range("E4").Value = "合格" And Range("H4").Value = "合格" Then
        Range("I4").Value = "合格"
    ElseIf Range("E4").Value = "不合格" Or Range("H4").Value = "不合格" Then
        Range("I4").Value = "不合格"
    Else
        MsgBox "合否が不正です"
    End If
End Sub
