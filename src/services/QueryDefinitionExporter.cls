VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QueryDefinitionExporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements IObjectDefinitionExporter

Const QUERY_TYPE_COLUMN_NUM As Integer = 0
Const QUERY_NAME_COLUMN_NUM As Integer = 1
Const SQL_COLUMN_NUM As Integer = 2
Const PARAM_TYPE_COLUMN_NUM As Integer = 3
Const PARAM_NAME_COLUMN_NUM As Integer = 4

Private m_curData As CurrentData
Private m_database As DAO.database
Private m_queryTypeToStr As Object
Private m_paramTypeToStr As Object

Public Sub Class_Initialize()
    Set m_curData = CurrentData
    Set m_database = CurrentDb
    SetQueryTypeToStr
    SetDataTypeToStr
End Sub

Public Sub IObjectDefinitionExporter_ExportToCsv(fullPath As String)
    Dim qName As String
    Dim queryDef As DAO.queryDef
    Dim pm As Object
    Dim accessOb As AccessObject
    Dim csvW As ICsvWriter
    Dim csvD(4) As Variant
    
    Set csvW = New CsvWriter
    
    csvW.CreateFile (fullPath)
    
    csvD(QUERY_TYPE_COLUMN_NUM) = "タイプ"
    csvD(QUERY_NAME_COLUMN_NUM) = "クエリ名"
    csvD(SQL_COLUMN_NUM) = "SQL"
    csvD(PARAM_TYPE_COLUMN_NUM) = "パラメータ型"
    csvD(PARAM_NAME_COLUMN_NUM) = "パラメータ名"
    
    csvW.AddLine csvD()
    
    For Each accessOb In m_curData.AllQueries
        qName = accessOb.Name
        Set queryDef = m_database.QueryDefs(qName)
        csvD(QUERY_TYPE_COLUMN_NUM) = m_queryTypeToStr(queryDef.Type)
        csvD(QUERY_NAME_COLUMN_NUM) = qName
        csvD(SQL_COLUMN_NUM) = Replace(queryDef.sql, vbCrLf, " ")
        For Each pm In queryDef.Parameters
            csvD(PARAM_TYPE_COLUMN_NUM) = pm.Name
            csvD(PARAM_NAME_COLUMN_NUM) = m_paramTypeToStr(pm.Type)
            csvW.AddLine csvD()
            csvD(QUERY_TYPE_COLUMN_NUM) = ""
            csvD(QUERY_NAME_COLUMN_NUM) = ""
            csvD(SQL_COLUMN_NUM) = ""
        Next
        If (queryDef.Parameters.Count = 0) Then
            csvD(PARAM_TYPE_COLUMN_NUM) = ""
            csvD(PARAM_NAME_COLUMN_NUM) = ""
            csvW.AddLine csvD()
        End If
    Next
    
    csvW.Complete
End Sub

Private Sub SetQueryTypeToStr()
    Set m_queryTypeToStr = CreateObject("Scripting.Dictionary")

    With m_queryTypeToStr
        .Add 0, "選択クエリ"
        .Add 16, "クロス集計クエリ"
        .Add 32, "削除クエリ"
        .Add 48, "更新クエリ"
        .Add 64, "追加クエリ"
        .Add 80, "テーブル作成クエリ"
        .Add 96, "DDL(データ定義言語)クエリ"
        .Add 112, "SQLパススルークエリ"
        .Add 128, "結合クエリ"
        .Add 144, "一括操作クエリ"
        .Add 160, "複合クエリ"
        .Add 224, "ストアドプロシージャを実行するSQLプロシージャ"
        .Add 240, "アクションクエリ"
    End With
End Sub

Private Sub SetDataTypeToStr()
    Set m_paramTypeToStr = CreateObject("Scripting.Dictionary")

    With m_paramTypeToStr
        .Add 8192, "adArray"
        .Add 20, "adBigInt"
        .Add 128, "adBinary"
        .Add 11, "adBoolean"
        .Add 8, "adBSTR"
        .Add 136, "adChapter"
        .Add 129, "adChar"
        .Add 6, "adCurrency"
        .Add 7, "adDate"
        .Add 133, "adDBDate"
        .Add 134, "adDBTime"
        .Add 135, "adDBTimeStamp"
        .Add 14, "adDecimal"
        .Add 5, "adDouble"
        .Add 0, "adEmpty"
        .Add 10, "adError"
        .Add 64, "adFileTime"
        .Add 72, "adGUID"
        .Add 9, "adIDispatch"
        .Add 3, "adInteger"
        .Add 13, "adIUnknown"
        .Add 205, "adLongVarBinary"
        .Add 201, "adLongVarChar"
        .Add 203, "adLongVarWChar"
        .Add 131, "adNumeric"
        .Add 138, "adPropVariant"
        .Add 4, "adSingle"
        .Add 2, "adSmallInt"
        .Add 16, "adTinyInt"
        .Add 21, "adUnsignedBigInt"
        .Add 19, "adUnsignedInt"
        .Add 18, "adUnsignedSmallInt"
        .Add 17, "adUnsignedTinyInt"
        .Add 132, "adUserDefined"
        .Add 204, "adVarBinary"
        .Add 200, "adVarChar"
        .Add 12, "adVariant"
        .Add 139, "adVarNumeric"
        .Add 202, "adVarWChar"
        .Add 130, "adWChar"
    End With
End Sub
