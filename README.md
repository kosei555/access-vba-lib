# Access VBA Library

Microsoft Access VBAの開発を効率化するための再利用可能なクラスライブラリ。

---

## 特徴

- モジュール化された再利用可能な設計
- データベース操作の簡素化
- トランザクションを考慮した安全な処理
- 可読性・保守性の向上

---

## コンポーネント

### QueryManager
#### 機能

- クエリ実行(パラメータクエリ可能)
- 選択クエリ実行結果のDAOレコードセット取得(パラメータクエリ可能)
- トランザクション制御
- クエリ実行時エラーハンドリング

#### メソッド
- RegisterQuery: 実行対象となるSQL文を登録します。
- SetParam: クエリのパラメータに対して値を設定します。
- ExecQuery: 登録されたクエリを実行します。
- GetQueryRecordSet: クエリを実行し、レコードセットを取得します。
- BeginTrans: トランザクションを開始します。
- CommitTrans: トランザクションを確定します。
- WithTarget: 操作対象となるオブジェクトを指定します。
- WSetParam: WithTargetメソッドで指定したクエリのパラメータを設定したい場合に使用。
- WExecQuery: WithTargetメソッドで指定したクエリを実行したい場合に使用。
#### 例
- パラメータクエリ実行
```
Option Compare Database
Option Explicit
Public Sub test()
    Dim qm As IQueryManager
    Set qm = New QueryManager
    'トランザクション開始
    qm.BeginTrans
    'クエリ名を渡す
    qm.RegisterQuery ("QueryManager_Test_I_P")
    'パラメータの設定
    qm.SetParam "QueryManager_Test_I_P", "p_id", "5"
    qm.SetParam "QueryManager_Test_I_P", "p_name", "6"
    qm.SetParam "QueryManager_Test_I_P", "p_birth_day", "7"
    'クエリ実行
    qm.ExecQuery ("QueryManager_Test_I_P")
    'トランザクションコミット
    qm.CommitTrans
    
End Sub

```

- WithTargetメソッドを使い、クエリ名を省略する場合
```
Option Compare Database
Option Explicit
Public Sub test()
    Dim qm As IQueryManager
    Set qm = New QueryManager
    With qm
    'トランザクション開始
        .BeginTrans
    'クエリ名を渡す
        .RegisterQuery ("QueryManager_Test_I_P")
        .WithTarget ("QueryManager_Test_I_P")
    'パラメータの設定
        .WSetParam "p_id", "5"
        .WSetParam "p_name", "6"
        .WSetParam "p_birth_day", "7"
    'クエリ実行
        .WExecQuery
    'トランザクションコミット
        .CommitTrans
    End With
End Sub

```

- 複数クエリの制御

```
Option Compare Database
Option Explicit
Public Sub test()
    Dim qm As IQueryManager
    Set qm = New QueryManager
    With qm
        .BeginTrans
        .RegisterQuery ("QueryManager_Test_C")
        .RegisterQuery ("QueryManager_Test_I_P")
        .WithTarget ("QueryManager_Test_I_P")
        .WSetParam "p_id", "5"
        .WSetParam "p_name", "6"
        .WSetParam "p_birth_day", "7"
        .ExecQuery ("QueryManager_Test_C")
        .ExecQuery ("QueryManager_Test_I_P")
        .CommitTrans
    End With
    
End Sub

```
- エラーハンドリングの挙動
```
Option Compare Database
Option Explicit
Public Sub test()
    Dim qm As IQueryManager
    Set qm = New QueryManager
    With qm
        .BeginTrans
        'テーブル作成クエリを登録
        .RegisterQuery ("QueryManager_Test_C")
        'テーブル作成クエリを実行
        .ExecQuery ("QueryManager_Test_C")
        '既にテーブルが作成されているためエラーを表示し、直近のBeginTransまで自動的にロールバックが行われる
        .ExecQuery ("QueryManager_Test_C")
        .CommitTrans
    End With
End Sub
```
- 選択クエリ実行結果のレコードセットを取得したい場合
```
Option Compare Database
Option Explicit
Public Sub test()
    Dim qm As IQueryManager
    Set qm = New QueryManager
    Dim rec As Recordset
    
    '選択クエリ登録
    qm.RegisterQuery ("QueryManager_Test_S_P")
    '選択クエリ実行結果のレコードセット取得
    Set rec = qm.GetQueryRecordSet("QueryManager_Test_S_P")
    rec.Close
End Sub
---
```
## ディレクトリ構造

* `src/services` — 主要なクラスファイルを保存
* `src/adapters` — service内にあるクラスを実行するために必要な依存クラス
* `src/interfaces` — 抽象クラス

---


## 現状

開発中

---

## License

MIT License
