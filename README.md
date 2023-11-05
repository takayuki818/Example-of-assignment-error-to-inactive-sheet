## 非アクティブシートへの代入エラーについて（親オブジェクト省略の落とし穴）
VBA初心者がつまづきがちな「**別のシートへの代入時エラー**」について、エラーコード例を交えて解説します。（本稿掲載の各コードは同じ処理の書き換えです）

```vba
Sub 非アクティブシートへの代入時エラー例()
    Dim 配列()
    Sheet1.Activate
    配列 = Range(Cells(1, 1), Cells(5, 5))
    Sheet2.Range(Cells(1, 1), Cells(5, 5)) = 配列
End Sub
```
↑のサンプルコードは、実行すると`Sheet2.Range(Cells(1, 1), Cells(5, 5)) = 配列`の行が原因で実行時エラーが発生します。

このコードは「`Sheet1`のセル範囲の値を動的配列に格納し、`Sheet2`のセル範囲に代入する」という意図のものですが、問題の箇所では**Cellsの親オブジェクトが省略**されています。
これにより、OSは**Cellsの親オブジェクトをActiveSheet(= Sheet1)だと認識する**ため、
**Cellsとその入れ子先であるRangeで親オブジェクトが異なる状態になってしまう**ことがエラーの原因です。

問題の箇所はOS側には↓のように認識されています。
`Sheet2.Range(ActiveSheet.Cells(1, 1), ActiveSheet.Cells(5, 5)) = 配列`
※例えるなら「目的地は千葉県にある東京都千葉市」というような誤ったコーディングとなっています。


このコードを正常に動作するよう修正すると、
```vba
Sub 非アクティブシートへの代入()
    Dim 配列()
    Sheet1.Activate
    配列 = Range(Cells(1, 1), Cells(5, 5))
    Range(Sheet2.Cells(1, 1), Sheet2.Cells(5, 5)) = 配列
End Sub
```
となります。

Rangeの親オブジェクトを省略しているので、OS側に
`ActiveSheet.Range(Sheet2.Cells(1, 1), Sheet2.Cells(5, 5)) = 配列`
のように認識されてしまいそうな所ですが、
入れ子構造は内部から順に特定されていく(`Cells`を特定 → `Cells`が入れ子された`Range`を特定の順)ため、問題は生じない模様です。(`Sheet2.Cells`の入れ子によって表現された`Range` → `Sheet2`が親オブジェクトだと認識される)

なお、実際のコーディングでは「With構文による親オブジェクト記述の省略」を用いると、「どのシートで何をしているか」が明白になり、可読性が上がりエラー防止にも繋がります。
```vba
Sub 別シートへの代入()
    Dim 配列()
    With Sheet1
        配列 = Range(.Cells(1, 1), .Cells(5, 5))
    End With
    With Sheet2
        Range(.Cells(1, 1), .Cells(5, 5)) = 配列
    End With
End Sub
```
※`Sheet1.Activate`は代入動作上不要になるので削除しました。

なお、「別のシートへの代入」には`Resize`を使ったやり方もあります。
```vba
Sub Resizeによる別解()
    Dim 配列()
    With Sheet1
        配列 = Range(.Cells(1, 1), .Cells(5, 5))
    End With
    With Sheet2
        .Cells(1, 1).Resize(5, 5) = 配列
    End With
End Sub
```
どちらを採用するかはケースバイケースですね。
