
# 背景

Accessで作成したクエリを安全に実行するため、クエリ実行と自動ロールバック機能を兼ね備えたクラスを作成した

---

# 機能概要



---

# メンバ変数

| アクセス修飾子 | 定数/変数 | 型 | 変数名 | 概要 |
| ---- | ---- |---- | ---- |---- |
| Private | 変数 | Object | m_queryDic | クエリ名をキーとしてIQueryObの参照情報を格納する |
| Private | 変数 | DAO.database | m_db | 本コードが実行されるAccessデータベースのDatabaseオブジェクトの参照情報を格納する |
| Private | 変数 | DAO.Workspace | m_ws | トランザクション制御に使用するWorkspaceオブジェクトの参照情報を格納する |
| Private | 変数 | String | m_execHistory | 過去に実行したクエリの履歴を格納する |
| Private | 変数 | Boolean | m_isBeginTrans | トランザクションが開始されているかをBoolean値で表現する |
| Private | 変数 | String | m_targetQuery | W系メソッドでターゲットにされるクエリ名を格納する |


---

# メソッド

---

# 機能詳細

---