# 期間スライダーボタンの作り方

表示期間を前月・翌月にスライドさせる機能

![完成イメージ](./screenshots/slidercontrol.png)

>💡ˎˊ˗ 前月ボタン／翌月ボタンをクリックすると、グラフの表示期間が1ヶ月ずつ移動し、データの端では警告が出る


---
### 仕組み

| 要素 | 内容 |
|------|------|
| シート1 'Sheet1'　| 前月・翌月ボタンを置くシート |
| シート2 'Sheet2'　| 日付などが入ったデータテーブル |
| セル1 'B1'　| 表示期間の**開始日**（VBAが書き換える） |
| セル2 'D1'　| 表示期間の**終了日** |
| ボタン2個 | 前月 / 翌月（シェイプまたはフォームコントロール） |
| 標準モジュール | 全プロシージャをここに記述 |

---

### STEP 1 シートとテーブルの準備

#### シート2にテーブルを作る

1. 新しいシートを追加し、VBEのプロパティで `(Name)` を `Sheet2` に設定
2. 以下の列構成でデータを入力する

```
例）
A列: 日付　B列: 費目1　C列: 費目2　D列: メモ
```

3. 範囲を選択して「挿入」→「テーブル」→ テーブル名を `シート2テーブル` に設定
4. データは**日付昇順**で入力する（`LBound` で最古日、`UBound` で最新日を取るため）

#### 推移シートを準備する

1. 新しいシートを追加し、VBEのプロパティで `(Name)` を `Sheet1` に設定
2. B1・D1 に表示期間の開始日・終了日を入力し、日付型として書式設定する

```
B1: 2024/01/01   （表示開始日）
D1: 2024/12/31   （表示終了日）
```

---

### STEP 2 Enum（列番号の定数）を定義

標準モジュールを挿入し、列番号を名前で管理できるよう Enum を定義します。  
`Sheet2(i, 1)` と書く代わりに `Sheet2(i, Sheet2Index.Hiduke)` と書けて可読性が上がる

```vba
Option Explicit

Enum Sheet2Index
    Hiduke  = 1   ' A列: 日付
    Himoku1 = 2   ' B列: 費目1
    Himoku2 = 3   ' C列: 費目2
    Memo    = 4   ' D列: メモ
End Enum
```

---

### STEP 3 ユーティリティを定義

```vba
Sub DisableUpdating()
    With Application
        .Calculation    = xlCalculationManual  ' 自動計算をOFF
        .ScreenUpdating = False                ' 画面更新をOFF
        .EnableEvents   = False                ' イベント発火をOFF
    End With
End Sub

Sub EnableUpdating()
    With Application
        .Calculation    = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents   = True
    End With
End Sub
```

---

### STEP 4 — 共通処理 `MoveSuiiMonth` を作る

前月・翌月どちらのボタンからも呼ばれる処理です。引数 `move` に `-1` または `1` を渡すだけで両方に対応します。

```vba
Sub MoveSuiiMonth(ByVal move As Integer)

    Dim dtStart As Date
    Dim dtEnd   As Date
    dtStart = Sheet2.Range("B1").Value
    dtEnd   = Sheet2.Range("D1").Value

    ' move ヶ月分ずらす
    dtStart = DateSerial(Year(dtStart), Month(dtStart) + move, 1)
    dtEnd   = DateSerial(Year(dtEnd),   Month(dtEnd) + 1 + move, 0)
    '   終了日: 「ずらした月の翌月の0日目」= ずらした月の末日

    ' セルに書き戻す（グラフが参照しているので自動更新される）
    Sheet2.Range("B1").Value = dtStart
    Sheet2.Range("D1").Value = dtEnd

    ' タイムラインスライサーを使っている場合は以下も追加
    ' ThisWorkbook.SlicerCaches("NativeTimeline_日付1") _
    '     .TimelineState.SetFilterDateRange CStr(dtStart), CStr(dtEnd)

End Sub
```

> 💡ˎˊ˗ スライサーキャッシュ名は「スライサーの設定」ダイアログで確認できます。  
> 💡ˎˊ˗ スライサーを使わない場合（SUMIFS で集計している場合）は不要

---

### STEP 5 前月ボタンの処理を作る

「まだ前のデータがあるか？」を確認してから移動します。

```vba
Sub MoveSuiiPreMonth()
    DisableUpdating

    ' テーブルを2次元配列に一括読み込み（行ループより高速）
    Dim tblSheet1Work As ListObject
    Set tblSheet1Work = wsSheet1Work.ListObjects("家計簿ワークテーブル")

    Dim Sheet1() As Variant
    Sheet1 = tblSheet1Work.DataBodyRange.Value

    ' 配列の先頭行（= 最古のデータ）から日付を取得し、月初に丸める
    Dim dtMin As Date
    dtMin = Sheet1(LBound(Sheet1), Sheet1Index.Hiduke)
    dtMin = DateSerial(Year(dtMin), Month(dtMin), 1)
    '   月初に丸める理由:
    '   データが 1/5 始まりでも「1月のグラフはまだ表示できる」と正しく判定するため

    ' 現在の表示開始日より前にデータがあれば移動
    If dtMin < Sheet2.Range("B1").Value Then
        MoveSuiiMonth -1
    Else
        MsgBox "これより過去のデータがありません"
    End If

    EnableUpdating
End Sub
```

---

### STEP 6 翌月ボタンの処理を作る

```vba
Sub MoveSuiiNextMonth()
    DisableUpdating

    Dim tblSheet1Work As ListObject
    Set tblSheet1Work = wsSheet1Work.ListObjects("家計簿ワークテーブル")

    Dim Sheet1() As Variant
    Sheet1 = tblSheet1Work.DataBodyRange.Value

    ' 配列の末尾行（= 最新のデータ）から日付を取得し、月末に丸める
    Dim dtMax As Date
    dtMax = Sheet1(UBound(Sheet1), Sheet1Index.Hiduke)
    dtMax = DateSerial(Year(dtMax), Month(dtMax) + 1, 0)
    '   月末に丸める理由:
    '   データが 6/15 で終わっても「6月のグラフはまだ表示できる」と正しく判定するため

    ' 現在の表示終了日より後にデータがあれば移動
    If Sheet2.Range("M1").Value < dtMax Then
        MoveSuiiMonth 1
    Else
        MsgBox "これより未来のデータがありません"
    End If

    EnableUpdating
End Sub
```

---

### STEP 7 — シートにボタンを配置してマクロを割り当てる

1. 推移シートを開き「挿入」→「図形」→「角丸四角形」でボタンを作成
2. テキストを `◀ 前月` に設定
3. 図形を右クリック →「マクロの登録」→ `MoveSuiiPreMonth` を選択
4. 同様に翌月ボタン（`翌月 ▶`）を作り `MoveSuiiNextMonth` を割り当て

---

### その他. グラフを B1・D1 と連動させる

グラフのデータ範囲を B1・D1 で絞り込む方法は2通りあります。

**方法A — SUMIFS で月ごとの集計列を作る（シンプル）**

推移シートに補助列を作り、グラフはその列を参照します。B1 が変わると SUMIFS の結果が変わり、グラフが自動更新されます。


**方法B — ピボットテーブル＋タイムラインスライサー（本家の実装）**

1. 家計簿ワークテーブルからピボットテーブルを作成
2. 日付フィールドにタイムラインスライサーを追加
3. STEP 4 の `MoveSuiiMonth` のコメントアウト部分（`SlicerCaches`）を有効化
