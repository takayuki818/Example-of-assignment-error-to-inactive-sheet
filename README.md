# Example-of-assignment-error-to-inactive-sheet
## 非アクティブシートへの代入エラーについて（親オブジェクト省略の落とし穴）
VBA初心者がつまづく「別のシートへの代入エラー」について、エラーコード例を交えて解説しています。

「Sub 非アクティブシートへの代入エラー例()」は、実行すると最後の行が原因でエラーになります。  
このコードは「動的配列にSheet1のセル範囲の値を格納しSheet2に代入する」意図のものですが、  
代入を実行する最後の行の「Sheet2.Range(Cells(1, 1), Cells(5, 5)) = 配列」では、  
Cellsの**親オブジェクトが省略**されています。  
これにより**OSはCellsの親オブジェクトをActiveSheet(= Sheet1)だと推定する**ため、  
**Cellsとその入れ子先であるRangeで親オブジェクトが異なる状態になってしまう**ことがエラーの原因です。  
※例えるなら「目的地は千葉県にある東京都千葉市」というような誤ったコーディングとなっています。

この行を正常に動作するよう修正すると「Range(Sheet2.Cells(1, 1), Sheet2.Cells(5, 5)) = 配列」になります。  
Rangeの親オブジェクトを省略しているので、OS側に  
「ActiveSheet.Range(Sheet2.Cells(1, 1), Sheet2.Cells(5, 5)) = 配列」  
のように推定されてしまいそうな所ですが、  
入れ子構造は内部から順に特定されていく(Cells特定 → Cellsが入れ子されたRange特定の順)ため、  
問題は生じない模様です。(Sheet2.Cellsの入れ子によって表現されるRange → Sheet2が親オブジェクトだと推定)

※実際のコーディングではWith構文による親オブジェクト記述省略を使うと楽で可読性も上がります。  
※別解として「Sheet2.Cells(1, 1).Resize(5, 5) = 配列」という書き方もあります。
