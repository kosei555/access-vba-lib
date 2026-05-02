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

#### 例

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

---

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
