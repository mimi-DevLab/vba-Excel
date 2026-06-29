# 期間スライダーボタンの作り方

表示期間を**前月・翌月へ1か月ずつ移動**するボタンを作成します。

![完成イメージ](./screenshots/slidercontrol.png)

> 💡 前月／翌月ボタンをクリックすると表示期間が1か月移動し、データの端まで到達するとメッセージを表示します。

---

# 全体の仕組み

| 要素           | 役割                   |
| ------------ | -------------------- |
| `Sheet1`     | 元データを保存するシート         |
| `Sheet1テーブル` | 日付・費目などを格納するテーブル     |
| `Sheet2`     | グラフ・スライダーボタンを配置するシート |
| `Sheet2!B1`  | 表示開始日（VBAが更新）        |
| `Sheet2!D1`  | 表示終了日（VBAが更新）        |
| 前月・翌月ボタン     | 表示期間を1か月ずつ移動する       |
| 標準モジュール      | すべてのVBAを記述する         |

---

# STEP1 シートとテーブルを準備する

## ① Sheet1（元データ）

新しいシートを作成し、VBEのプロパティで `(Name)` を **Sheet1** に設定します。

次のようなデータを作成します。

| 列  | 内容  |
| -- | --- |
| A列 | 日付  |
| B列 | 費目1 |
| C列 | 費目2 |
| D列 | メモ  |

範囲をテーブル化し、テーブル名を

```
Sheet1テーブル
```

に設定します。

> **データは日付の昇順で並べてください。**
> 最古日・最新日の判定に先頭行・末尾行を利用します。

---

## ② Sheet2（グラフ・操作画面）

新しいシートを作成し、VBEのプロパティで `(Name)` を **Sheet2** に設定します。

表示期間を管理するセルを用意します。

| セル | 内容    |
| -- | ----- |
| B1 | 表示開始日 |
| D1 | 表示終了日 |

例

```
B1：2024/01/01
D1：2024/12/31
```

どちらも日付形式に設定してください。

---

# STEP2 Enum（列番号）を定義する

列番号を名前で管理できるようにします。

```vba
Option Explicit

Enum Sheet1Index
    Hiduke  = 1   ' A列：日付
    Himoku1 = 2   ' B列：費目1
    Himoku2 = 3   ' C列：費目2
    Memo    = 4   ' D列：メモ
End Enum
```

これにより、

```vba
arr(i, Sheet1Index.Hiduke)
```

のように書けるため、可読性が向上します。

---

# STEP3 共通ユーティリティを作成する

画面更新・自動計算・イベントを一時停止する処理です。

```vba
Sub DisableUpdating()
    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
    End With
End Sub

Sub EnableUpdating()
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
    End With
End Sub
```

---

# STEP4 共通処理 `MoveSuiiMonth` を作成する

前月・翌月ボタンの両方から呼び出される処理です。

引数

* `-1`：前月
* `1`：翌月

を指定します。

```vba
Sub MoveSuiiMonth(ByVal move As Integer)

    Dim dtStart As Date
    Dim dtEnd As Date

    dtStart = Sheet2.Range("B1").Value
    dtEnd = Sheet2.Range("D1").Value

    dtStart = DateSerial(Year(dtStart), Month(dtStart) + move, 1)
    dtEnd = DateSerial(Year(dtEnd), Month(dtEnd) + move + 1, 0)

    Sheet2.Range("B1").Value = dtStart
    Sheet2.Range("D1").Value = dtEnd

    ' タイムラインスライサーを使用する場合
    ' ThisWorkbook.SlicerCaches("NativeTimeline_日付1") _
    '     .TimelineState.SetFilterDateRange CStr(dtStart), CStr(dtEnd)

End Sub
```

### この処理で行うこと

1. `Sheet2!B1`・`Sheet2!D1`から現在の表示期間を取得
2. 指定した月数だけ移動
3. `Sheet2!B1`・`Sheet2!D1`へ書き戻す
4. 必要に応じてタイムラインスライサーを更新

---

# STEP5 前月ボタンを作成する

過去データが存在する場合のみ表示期間を移動します。

```vba
Sub MoveSuiiPreMonth()

    DisableUpdating

    Dim tbl As ListObject
    Set tbl = Sheet1.ListObjects("Sheet1テーブル")

    Dim arr As Variant
    arr = tbl.DataBodyRange.Value

    Dim dtMin As Date
    dtMin = arr(LBound(arr), Sheet1Index.Hiduke)

    dtMin = DateSerial(Year(dtMin), Month(dtMin), 1)

    If dtMin < Sheet2.Range("B1").Value Then
        MoveSuiiMonth -1
    Else
        MsgBox "これより過去のデータがありません"
    End If

    EnableUpdating

End Sub
```

### 判定方法

* テーブルの最古の日付を取得
* 月初へ丸める
* 表示開始日より前のデータがある場合のみ移動

---

# STEP6 翌月ボタンを作成する

未来データが存在する場合のみ表示期間を移動します。

```vba
Sub MoveSuiiNextMonth()

    DisableUpdating

    Dim tbl As ListObject
    Set tbl = Sheet1.ListObjects("Sheet1テーブル")

    Dim arr As Variant
    arr = tbl.DataBodyRange.Value

    Dim dtMax As Date
    dtMax = arr(UBound(arr), Sheet1Index.Hiduke)

    dtMax = DateSerial(Year(dtMax), Month(dtMax) + 1, 0)

    If Sheet2.Range("D1").Value < dtMax Then
        MoveSuiiMonth 1
    Else
        MsgBox "これより未来のデータがありません"
    End If

    EnableUpdating

End Sub
```

### 判定方法

* テーブルの最新日を取得
* 月末へ丸める
* 表示終了日より後のデータがある場合のみ移動

---

# STEP7 ボタンを配置する

1. `Sheet2` を開く
2. **挿入 → 図形** を選択
3. 「角丸四角形」を配置
4. 「◀ 前月」と入力
5. `MoveSuiiPreMonth` を割り当てる
6. 同様に「翌月 ▶」を作成し、`MoveSuiiNextMonth` を割り当てる

---

# グラフと連動させる

表示期間の変更をグラフへ反映する方法は2通りあります。

## 方法1：SUMIFSを使用する

補助列で `Sheet2!B1` ～ `Sheet2!D1` の期間だけを集計し、その結果をグラフ化します。

シンプルで実装しやすい方法です。

---

## 方法2：ピボットテーブル＋タイムラインスライサー

1. `Sheet1テーブル` からピボットテーブルを作成する
2. 日付フィールドにタイムラインスライサーを追加する
3. `MoveSuiiMonth` の `SlicerCaches` を有効にする

こちらはタイムラインスライサーを利用して表示期間を更新する方法です。
