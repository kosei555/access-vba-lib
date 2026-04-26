VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QueryManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'-------------------------------------------------------------------------------
'クラス名   :   QueryManager
'概要       :   クエリの管理およびトランザクション制御
'役割       :   複数のQueryObクラスを保有し、各QueryObインスタンスを管理する
'               クエリ実行時にエラーが出た場合はロールバックを行う
'依存関係   :   IQueryManager(インターフェース), IQueryOb(クエリ実体インターフェイス),QueryOb(クエリ実体)
'実行環境   :   64bit-AccessVba
'参照設定   :   Microsoft Office xx.x database engine object library
'使い方     :   RegisterQueryメソッドで使いたいクエリを登録
'                               ↓
'               パラメータクエリの場合はSetParamメソッドでパラメータを設定
'                               ↓
'               BeginTransメソッドでトランザクション開始
'                               ↓
'               ExecQueryメソッドでクエリ実行 (Select文のレコードセットが欲しい場合はGetQueryRecordSetメソッドを使用)
'                               ↓
'               Commitメソッドでクエリ結果をコミット
'
'注意点     :   BeginTransを使用しないとクエリ実行でエラーが発生した場合、ロールバックされない。
'               GetQueryRecordSetではオプションのisDynasetにTrueをいれなければDynasetでレコードセットを取得できない。
'確認状況   :   非パラメータクエリの実行〇
'           :   パラメータクエリの実行
'           :   非パラメータクエリのレコードセット取得
'           :   パラメータクエリのレコードセット取得
'           :   ロールバックが正常にされるか
'-------------------------------------------------------------------------------



Option Compare Database
Implements IQueryManager

Const QUERY_ERROR_MESSAGE As String = "クエリ実行時にエラーが発生しました。"
Const ERR_RAISE_NUMBER As Integer = 513
Const ERR_RAISE_SOURCE As String = "QueryManager"
Const ERR_RAISE_DESCRITION As String = "ロールバックが完了しました。コードの実行を強制停止します。"
Const ROLLBACK_MESSAGE As String = "次のクエリ実行結果がロールバックされます。"

Private m_queryDic As Object
Private m_db As DAO.Database
Private m_ws As DAO.Workspace
Private m_execHistory As String
Private m_isBeginTrans As Boolean


Public Sub Class_Initialize()
    Set m_db = CurrentDb
    Set m_queryDic = CreateObject("Scripting.Dictionary")
    Set m_ws = DBEngine.Workspaces(0)
    m_isBeginTrans = False
End Sub
Public Sub IQueryManager_RegisterQuery(qName As String)
    Dim query As IQueryOb
    Set query = New QueryOb
    m_queryDic.Add qName, query
    Call query.QueryObInit(m_db, qName)
End Sub
Public Sub IQueryManager_SetParam(qName As String, pName As String, pData As Variant)

    Dim query As IQueryOb
    Set query = m_queryDic.Item(qName)
    query.SetParam pName, pData
End Sub
Public Sub IQueryManager_ExecQuery(qName As String)
    Dim query As IQueryOb
    Dim isError As Boolean
    Set query = m_queryDic.Item(qName)
    isError = query.ExecQuery
    m_execHistory = m_execHistory & vbCrLf & qName
    If (isError) Then
        MsgBox qName & QUERY_ERROR_MESSAGE
        If (m_isBeginTrans) Then
            Call Rollback
        End If
    End If
End Sub
Public Function IQueryManager_GetQueryRecordSet(qName As String, Optional isDynaset As Boolean = False) As DAO.Recordset
    Dim query As IQueryOb
    Set query = m_queryDic.Item(qName)
    Set IQueryManager_GetQueryRecordSet = query.GetQueryRecordSet(isDynaset)
    m_execHistory = m_execHistory & vbCrLf & qName
    If (IQueryManager_GetQueryRecordSet Is Nothing) Then
        MsgBox qName & QUERY_ERROR_MESSAGE
        If (m_isBeginTrans) Then
            Call Rollback
        End If
    End If
End Function
Public Sub IQueryManager_BeginTrans()
    m_ws.BeginTrans
    m_execHistory = ""
    m_isBeginTrans = True
End Sub
Public Sub IQueryManager_CommitTrans()
    m_ws.CommitTrans
    m_isBeginTrans = False
End Sub
Private Sub Rollback()
    MsgBox ROLLBACK_MESSAGE & vbCrLf & m_execHistory
    m_ws.Rollback
    Err.Raise Number:=ERR_RAISE_NUMBER, Source:=ERR_RAISE_SOURCE, Description:=ERR_RAISE_DESCRITION
End Sub


