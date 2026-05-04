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
'依存関係   :   IQueryManager(インターフェース), IQueryOb(クエリ実体インターフェイス),QueryOb(クエリ実体),IErrBuf(エラー情報保存クラスインターフェイス),ErrBuf(エラー情報保存クラス)
'実行環境   :   64bit-AccessVba
'参照設定   :   Microsoft Office xx.x database engine object library
'               Microsoft Scripting Runtime
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
'           :   パラメータクエリの実行〇
'           :   非パラメータクエリのレコードセット取得
'           :   パラメータクエリのレコードセット取得〇
'           :   ロールバックが正常にされるか〇
'-------------------------------------------------------------------------------



Option Compare Database
Implements IQueryManager

Private Const ROLLBACK_MESSAGE As String = "以下のクエリ実行をロールバックします"

Private m_queryDic As Object
Private m_db As DAO.database
Private m_ws As DAO.Workspace
Private m_execHistory As String
Private m_isBeginTrans As Boolean
Private m_targetQuery As String
Private m_curData As CurrentData


Public Sub Class_Initialize()
    Set m_db = CurrentDb
    Set m_curData = CurrentData
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
Public Sub IQueryManager_WSetParam(pName As String, pData As Variant)
    Call IQueryManager_SetParam(m_targetQuery, pName, pData)
End Sub
Public Function IQueryManager_WithTarget(qName As String) As IQueryManager
    m_targetQuery = qName
End Function
Public Sub IQueryManager_ExecQuery(qName As String)
    Dim query As IQueryOb
    Dim errorInfo As IErrBuf
    Set query = m_queryDic.Item(qName)
    Set errorInfo = query.ExecQuery
    m_execHistory = m_execHistory & vbCrLf & qName
    If (Not errorInfo Is Nothing) Then
        Call IQueryManager_ShowErrorAndRollback(errorInfo)
    End If
End Sub
Public Sub IQueryManager_WExecQuery()
    Call IQueryManager_ExecQuery(m_targetQuery)
End Sub

Public Sub IQueryManager_RegisterAllQuery()
    Dim accessOb As AccessObject
    Dim queryName As String
    Dim query As IQueryOb
    For Each accessOb In m_curData.AllQueries
        Set query = New QueryOb
        queryName = accessOb.Name
        m_queryDic.Add queryName, query
        Call query.QueryObInit(m_db, queryName)
        Set query = Nothing
    Next
End Sub


Public Function IQueryManager_GetQueryRecordSet(qName As String, Optional isDynaset As Boolean = False) As DAO.Recordset
    Dim query As IQueryOb
    Dim rec As DAO.Recordset
    Dim errorInfo As Object
    Set query = m_queryDic.Item(qName)
    Set errorInfo = query.GetQueryRecordSet(rec, isDynaset)
    m_execHistory = m_execHistory & vbCrLf & qName
    If (Not errorInfo Is Nothing) Then
        Call IQueryManager_ShowErrorAndRollback(errorInfo)
    End If
    Set IQueryManager_GetQueryRecordSet = rec
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
Private Sub IQueryManager_ShowErrorAndRollback(errorInfo As IErrBuf)
    If (m_isBeginTrans) Then
        Call IQueryManager_Rollback
    End If
    Err.Raise Number:=errorInfo.Number, Source:=errorInfo.Source, Description:=errorInfo.Description
End Sub
Private Sub IQueryManager_Rollback()
    MsgBox ROLLBACK_MESSAGE & vbCrLf & m_execHistory
    m_ws.Rollback
    m_isBeginTrans = False
End Sub
