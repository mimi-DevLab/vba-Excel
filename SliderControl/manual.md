# 期間スライダーボタン（`MoveSuiiPreMonth` / `MoveSuiiNextMonth`）

表示期間を前月・翌月にスライドさせる機能
「データの端まで来たら移動を止める」という**ガード付き月移動**の実装パターン

! 完成イメージ

---

### 処理の全体像

```
ボタンクリック
    │
    ├─ 家計簿ワークテーブルから「最古の日付」または「最新の日付」を取得
    │
    ├─ 現在の表示開始日（K1）または表示終了日（M1）と比較
    │
    ├─ まだ移動できる → MoveSuiiMonth(-1 または 1) を呼ぶ
    └─ 端まで来ている → MsgBox でお知らせして終了
```

---

### `MoveSuiiPreMonth` — 前月へ移動

```vba
Sub MoveSuiiPreMonth()
    DisableUpdating   ' 画面更新・イベントを一時停止（処理を速くするおまじない）

    ' --- ① 家計簿ワークテーブルを配列に読み込む ---
    Dim tblKakeiboWork As ListObject
    Set tblKakeiboWork = wsKakeiboWork.ListObjects("家計簿ワークテーブル")

    Dim kakeibo() As Variant
    kakeibo = tblKakeiboWork.DataBodyRange.Value
    '   DataBodyRange = テーブルのヘッダーを除いたデータ部分
    '   .Value で2次元配列に一括取得（行ごとのループより高速）

    ' --- ② データの最古日（月初）を求める ---
    Dim dtMin As Date
    dtMin = kakeibo(LBound(kakeibo), KakeiboIndex.Hiduke)
    '   LBound(kakeibo) = 配列の先頭インデックス（通常は1）
    '   KakeiboIndex.Hiduke = 日付列の列番号を示すEnum定数

    dtMin = DateSerial(Year(dtMin), Month(dtMin), 1)
    '   日付を「その月の1日」に丸める
    '   例: 2021/01/05 → 2021/01/01

    ' --- ③ 現在の表示開始日と比較 ---
    Dim dtStart As Date
    dtStart = wsSuii.Range("K1").Value
    '   K1セルには「現在グラフに表示中の期間の開始日」が入っている

    If dtMin < dtStart Then
        MoveSuiiMonth -1   ' まだ過去データがある → 1ヶ月前へ
    Else
        MsgBox "これより過去のデータがありません"   ' 先頭に達した → 止める
    End If

    EnableUpdating   ' 画面更新・イベントを再開
End Sub
```

#### ポイント解説

**`LBound(kakeibo)` で先頭行を取得する理由**

配列のインデックスは環境によって 0 始まりにも 1 始まりにもなりえます。`LBound()` を使うことで、どちらでも正しく先頭行を指せます。

```vba
' 危険（インデックスが1でない場合に誤動作する）
dtMin = kakeibo(1, KakeiboIndex.Hiduke)

' 安全（常に先頭行を指す）
dtMin = kakeibo(LBound(kakeibo), KakeiboIndex.Hiduke)
```

**`DateSerial(Year(dtMin), Month(dtMin), 1)` で月初に丸める理由**

データの最初の行が月中（例: 1月5日）の場合、そのまま比較すると「1月5日より前のデータがない」と判定されてしまいます。月初（1月1日）に丸めることで、「この月のグラフはまだ表示できる」と正しく判定できます。

```
データ最古日: 2021/01/05
月初に丸める: 2021/01/01

表示開始日（K1）: 2021/01/01 の場合
  dtMin(2021/01/01) < dtStart(2021/01/01) → False → 「過去データなし」と表示 ✅

表示開始日（K1）: 2021/02/01 の場合
  dtMin(2021/01/01) < dtStart(2021/02/01) → True  → 前月へ移動 ✅
```

---

### `MoveSuiiNextMonth` — 翌月へ移動

```vba
Sub MoveSuiiNextMonth()
    DisableUpdating

    ' --- ① 家計簿ワークテーブルを配列に読み込む ---
    Dim tblKakeiboWork As ListObject
    Set tblKakeiboWork = wsKakeiboWork.ListObjects("家計簿ワークテーブル")

    Dim kakeibo() As Variant
    kakeibo = tblKakeiboWork.DataBodyRange.Value

    ' --- ② データの最新日（月末）を求める ---
    Dim dtMax As Date
    dtMax = kakeibo(UBound(kakeibo), KakeiboIndex.Hiduke)
    '   UBound(kakeibo) = 配列の末尾インデックス = データの最終行

    dtMax = DateSerial(Year(dtMax), Month(dtMax) + 1, 0)
    '   「翌月の0日目」= その月の末日
    '   例: 2024/06/15 → DateSerial(2024, 7, 0) → 2024/06/30

    ' --- ③ 現在の表示終了日と比較 ---
    Dim dtEnd As Date
    dtEnd = wsSuii.Range("M1").Value
    '   M1セルには「現在グラフに表示中の期間の終了日」が入っている

    If dtEnd < dtMax Then
        MoveSuiiMonth 1   ' まだ未来データがある → 1ヶ月後へ
    Else
        MsgBox "これより未来のデータがありません"   ' 末尾に達した → 止める
    End If

    EnableUpdating
End Sub
```

#### ポイント解説

**`UBound(kakeibo)` で末尾行を取得する**

`LBound` の逆で、配列の最終インデックスを返します。テーブルが日付順にソートされていることを前提に、末尾行 = 最新データとして扱います。

**`DateSerial(Year(dtMax), Month(dtMax) + 1, 0)` で月末に丸める**

前月ボタンと対称的に、最新日を「その月の末日」に丸めます。これにより「その月のグラフがまだ表示できるか」を正しく判定できます。

```
データ最新日: 2024/06/15
月末に丸める: DateSerial(2024, 7, 0) = 2024/06/30

表示終了日（M1）: 2024/06/30 の場合
  dtEnd(2024/06/30) < dtMax(2024/06/30) → False → 「未来データなし」と表示 ✅

表示終了日（M1）: 2024/05/31 の場合
  dtEnd(2024/05/31) < dtMax(2024/06/30) → True  → 翌月へ移動 ✅
```

---

### `MoveSuiiMonth` — 実際に表示期間をずらす（共通処理）

前月・翌月ボタンの両方から呼ばれる共通処理です。`-1` または `1` を受け取って K1・M1 を書き換え、グラフのスライサーも連動させます。

```vba
Sub MoveSuiiMonth(ByVal move As Integer)

    ' K1（開始日）・M1（終了日）を読む
    Dim dtStart As Date
    Dim dtEnd   As Date
    dtStart = wsSuii.Range("K1").Value
    dtEnd   = wsSuii.Range("M1").Value

    ' move ヶ月分ずらす
    dtStart = DateSerial(Year(dtStart), Month(dtStart) + move, 1)
    dtEnd   = DateSerial(Year(dtEnd),   Month(dtEnd)   + 1 + move, 0)
    '   終了日は「ずらした月の末日」になるよう計算

    ' K1・M1 を更新（これでグラフの集計範囲が変わる）
    wsSuii.Range("K1").Value = dtStart
    wsSuii.Range("M1").Value = dtEnd

    ' タイムラインスライサーも連動させて同期
    ThisWorkbook.SlicerCaches("NativeTimeline_日付1") _
        .TimelineState.SetFilterDateRange CStr(dtStart), CStr(dtEnd)

End Sub
```

---

### 3つのプロシージャの関係図

```
前月ボタンクリック              翌月ボタンクリック
       │                              │
MoveSuiiPreMonth()           MoveSuiiNextMonth()
       │                              │
  ① データ最古日を取得          ① データ最新日を取得
  ② 月初に丸める               ② 月末に丸める
  ③ K1（表示開始）と比較        ③ M1（表示終了）と比較
       │                              │
  移動できる？                  移動できる？
  Yes → MoveSuiiMonth(-1)      Yes → MoveSuiiMonth(1)
  No  → MsgBox で通知          No  → MsgBox で通知
                │
        MoveSuiiMonth(move)
                │
        K1・M1 を書き換え
        スライサーを同期
        グラフが自動更新
```

---

### `DisableUpdating` / `EnableUpdating` とは

コード内で `DisableUpdating` と `EnableUpdating` が呼ばれていますが、これは標準モジュールで定義されたユーティリティです。

```vba
Sub DisableUpdating()
    With Application
        .Calculation   = xlCalculationManual  ' 自動計算をOFF
        .ScreenUpdating = False                ' 画面更新をOFF（チラつき防止）
        .EnableEvents   = False                ' イベント発火をOFF（無限ループ防止）
    End With
End Sub

Sub EnableUpdating()
    With Application
        .Calculation   = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents   = True
    End With
End Sub
```

処理の最初に止めて、最後に必ず再開する、というセットで使います。`EnableUpdating` を忘れると Excelが計算しなくなるので注意してください。
