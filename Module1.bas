Attribute VB_Name = "Module1"
Option Explicit
Sub 非アクティブシートへの代入エラー例()
    Dim 配列()
    Sheet1.Activate
    配列 = Range(Cells(1, 1), Cells(5, 5))
    Sheet2.Range(Cells(1, 1), Cells(5, 5)) = 配列
End Sub
Sub OS視点から見たエラー例()
    Dim 配列()
    Sheet1.Activate
    配列 = Range(Cells(1, 1), Cells(5, 5))
    'Cellsの親オブジェクトが省略されている → ActiveSheet(= Sheet1)が親オブジェクトだと推定
    Sheet2.Range(ActiveSheet.Cells(1, 1), ActiveSheet.Cells(5, 5)) = 配列
End Sub
Sub OK例1()
    Dim 配列()
    Sheet1.Activate
    配列 = Range(Cells(1, 1), Cells(5, 5))
    Range(Sheet2.Cells(1, 1), Sheet2.Cells(5, 5)) = 配列
    'Rangeの親オブジェクトが省略されており、ActiveSheet(= Sheet1)が親オブジェクトだと推定されそうにも見えるが、
    '入れ子構造は内部から順に特定されていく(Cells特定 → Cellsが入れ子されたRange特定の順)ため、
    '問題無い模様。(Sheet2.Cellsの入れ子によって表現されるRange → Sheet2が親オブジェクトだと推定)
    '※実際のコーディングではWith構文による親オブジェクト記述省略を使うと楽で可読性も上がります。
    End With
End Sub
Sub OK例2()
    Dim 配列()
    Sheet1.Activate
    配列 = Range(Cells(1, 1), Cells(5, 5))
    Sheet2.Cells(1, 1).Resize(5, 5) = 配列
End Sub
