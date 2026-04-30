VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QueryOb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'-------------------------------------------------------------------------------
'クラス名   :   QueryOb
'概要       :   クエリの実行、パラメータクエリのパラメータ設定
'役割       :   一つのクエリに対するパラメータの設定
'               クエリの実行
'               選択クエリ実行結果のレコードセットの取得
'依存関係   :   IQueryOb(クエリ実体インターフェイス),IQueryManager(QueryMangaerインターフェース), QueryMangaer(QueryObの管理)
'実行環境   :   64bit-AccessVba
'参照設定   :   Microsoft Office xx.x database engine object library
'               Microsoft Scripting Runtime
'注意点     :   標準モジュールからQueryObを直接使用することを想定していない。必ずQueryMangaerを使うこと。
'-------------------------------------------------------------------------------



Option Compare Database
Implements IQueryOb
Const RS_TYPE_SNAPSHOT As Integer = 4
Const RS_TYPE_DYNASET As Integer = 2
Private m_queryName As String
Private m_paramDic As Object
Private m_queryDef As DAO.queryDef
Private m_db As DAO.database


Public Sub Class_Initialize()
    Set m_paramDic = CreateObject("Scripting.Dictionary")
End Sub

Public Sub IQueryOb_QueryObInit(db As DAO.database, qName As String)
    Set m_db = db
    Set m_queryDef = db.QueryDefs(qName)
End Sub



Public Sub IQueryOb_SetParam(pName As String, pData As Variant)
    m_paramDic.Add pName, pData
End Sub

Public Function IQueryOb_GetQueryCode() As String
    IQueryOb_GetQueryCode = m_queryDef.sql
End Function

Public Function IQueryOb_ExecQuery() As IErrBuf
    On Error GoTo Catch
    Dim error As IErrBuf
    DoCmd.SetWarnings False
    For Each pName In m_paramDic.Keys
        m_queryDef(pName).Value = m_paramDic.Item(pName)
    Next
    m_queryDef.Execute dbFailOnError
    Set IQueryOb_ExecQuery = Nothing
    GoTo Finally
Catch:
    Set error = New ErrBuf
    With error
        .Description = Err.Description
        .Number = Err.Number
        .Source = Err.Source
    End With
    Set IQueryOb_ExecQuery = error
Finally:
    m_paramDic.RemoveAll
    DoCmd.SetWarnings True
End Function
Public Function IQueryOb_GetQueryRecordSet(rec As DAO.Recordset, isDynaset As Boolean) As IErrBuf
    On Error GoTo Catch
    Dim error As IErrBuf
    DoCmd.SetWarnings False
    Dim rsType As Integer
    rsType = RS_TYPE_SNAPSHOT
    If (isDynaset) Then
        rsType = RS_TYPE_DYNASET
    End If
    For Each pName In m_paramDic
        m_queryDef(pName).Value = m_paramDic.Item(pName)
    Next
    Set rec = m_queryDef.OpenRecordset(rsType)
    Set IQueryOb_GetQueryRecordSet = Nothing
    GoTo Finally
Catch:
    Set error = New ErrBuf
    With error
        .Description = Err.Description
        .Number = Err.Number
        .Source = Err.Source
    End With
    Set IQueryOb_GetQueryRecordSet = error
Finally:
    m_paramDic.RemoveAll
    DoCmd.SetWarnings True
End Function


