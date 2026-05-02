VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableDefinitionExporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Implements IObjectDefinitionExporter

Const TABLE_LINK_ATTRIBUTE_COLUMN_NUM As Integer = 0
Const TABLE_EXCLUSIVE_ATTRIBUTE_COLUMN_NUM As Integer = 1
Const TABLE_TABLE_NAME_COLUMN_NUM As Integer = 2
Const TABLE_PRIMARYKEY_COLUMN_NUM As Integer = 3
Const TABLE_FIELD_NAME_COLUMN_NUM As Integer = 4
Const TABLE_DATA_TYPE_COLUMN_NUM As Integer = 5
Const TABLE_SIZE_COLUMN_NUM As Integer = 6
Const TABLE_NULL_ALLOW_COLUMN_NUM As Integer = 7
Const IDX_TABLE_NAME_COLUMN_NUM As Integer = 0
Const IDX_NAME_COLUMN_NUM As Integer = 1
Const IDX_FIELD_NAME_COLUMN_NUM As Integer = 2
Const IDX_SORT_COLUMN_NUM As Integer = 3
Const TABLE_LINKED_ODBC As String = "ODBC接続リンクテーブル"
Const TABLE_LINKED_NOT_ODBC As String = "ODBC接続でないリンクテーブル"
Const TABLE_LOCAL As String = "ローカルテーブル"

Private m_curData As CurrentData
Private m_database As DAO.database
Private m_dbDataTypeToStr As Object


Public Sub Class_Initialize()
    Set m_curData = CurrentData
    Set m_database = CurrentDb
    SetDbDataTypeToStr
End Sub

Public Sub IObjectDefinitionExporter_ExportToCsv(fullPath As String)
    Dim csvW As ICsvWriter
    Set csvW = New CsvWriter
    csvW.CreateFile (fullPath)
    TableExport csvW
    IdxExport csvW
    csvW.Complete
    
End Sub

Private Sub TableExport(csvW As ICsvWriter)
    Dim tableDef As tableDef
    Dim accessOb As AccessObject
    Dim primary As Object
    Dim idx As Object
    Dim field As field
    Dim csvD(8) As Variant
    Dim fieldName As String
    Dim tableAttribute As Long
    Set primary = CreateObject("Scripting.Dictionary")
    csvD(TABLE_LINK_ATTRIBUTE_COLUMN_NUM) = "テーブルリンク属性"
    csvD(TABLE_EXCLUSIVE_ATTRIBUTE_COLUMN_NUM) = "テーブル排他設定"
    csvD(TABLE_TABLE_NAME_COLUMN_NUM) = "テーブル名"
    csvD(TABLE_PRIMARYKEY_COLUMN_NUM) = "主キー"
    csvD(TABLE_FIELD_NAME_COLUMN_NUM) = "フィールド名"
    csvD(TABLE_DATA_TYPE_COLUMN_NUM) = "データ型"
    csvD(TABLE_SIZE_COLUMN_NUM) = "サイズ"
    csvD(TABLE_NULL_ALLOW_COLUMN_NUM) = "Null許可"

    csvW.AddLine csvD()
    
    For Each accessOb In m_curData.AllTables
        primary.RemoveAll
        Set tableDef = m_database.TableDefs(accessOb.Name)
        If (IsSystemTable(tableDef.Attributes)) Then
            GoTo SystemTableSkip
        End If
        tableAttribute = tableDef.Attributes
        csvD(TABLE_LINK_ATTRIBUTE_COLUMN_NUM) = GetTableLinkAttribute(tableAttribute)
        If (IsExclusiveTable(tableAttribute)) Then
            csvD(TABLE_EXCLUSIVE_ATTRIBUTE_COLUMN_NUM) = "〇"
        Else
            csvD(TABLE_EXCLUSIVE_ATTRIBUTE_COLUMN_NUM) = "×"
        End If
        csvD(TABLE_TABLE_NAME_COLUMN_NUM) = tableDef.Name
        
        '主キーとなるフィールドをIndexから取得
        For Each idx In tableDef.Indexes
            If (idx.primary) Then
                For Each field In idx.Fields
                    primary.Add field.Name, True
                Next
            End If
        Next
        
        For Each field In tableDef.Fields
            fieldName = field.Name
            csvD(TABLE_FIELD_NAME_COLUMN_NUM) = fieldName
            csvD(TABLE_DATA_TYPE_COLUMN_NUM) = m_dbDataTypeToStr(field.Type)
            csvD(TABLE_SIZE_COLUMN_NUM) = field.Size
            If (primary.Exists(fieldName)) Then
                csvD(TABLE_PRIMARYKEY_COLUMN_NUM) = "〇"
            Else
                csvD(TABLE_PRIMARYKEY_COLUMN_NUM) = "×"
            End If
            
            If (field.Required) Then
                csvD(TABLE_NULL_ALLOW_COLUMN_NUM) = "×"
            Else
                csvD(TABLE_NULL_ALLOW_COLUMN_NUM) = "〇"
            End If
            csvW.AddLine csvD()
            csvD(TABLE_LINK_ATTRIBUTE_COLUMN_NUM) = ""
            csvD(TABLE_EXCLUSIVE_ATTRIBUTE_COLUMN_NUM) = ""
            csvD(TABLE_TABLE_NAME_COLUMN_NUM) = ""

        Next
SystemTableSkip:
    
    Next
    
    
End Sub

Private Sub IdxExport(csvW As ICsvWriter)
    Dim accessOb As AccessObject
    Dim tDef As tableDef
    Dim idx As Object
    Dim field As Object
    Dim csvD(3) As Variant
    Dim tableName As String
    
    
    csvD(IDX_TABLE_NAME_COLUMN_NUM) = "テーブル名"
    csvD(IDX_NAME_COLUMN_NUM) = "インデックス名"
    csvD(IDX_FIELD_NAME_COLUMN_NUM) = "フィールド名"
    csvD(IDX_SORT_COLUMN_NUM) = "ソート"
    
    csvW.AddLine csvD()
    For Each accessOb In m_curData.AllTables
        tableName = accessOb.Name
        csvD(IDX_TABLE_NAME_COLUMN_NUM) = tableName
        Set tDef = m_database.TableDefs(tableName)
        If (IsSystemTable(tDef.Attributes)) Then
            GoTo SystemTableSkip
        End If
        For Each idx In tDef.Indexes
            csvD(IDX_NAME_COLUMN_NUM) = idx.Name
            For Each field In idx.Fields
                csvD(IDX_FIELD_NAME_COLUMN_NUM) = field.Name
                If (field.Attributes = 1) Then
                    csvD(IDX_SORT_COLUMN_NUM) = "DESC"
                Else
                    csvD(IDX_SORT_COLUMN_NUM) = "ASC"
                End If
                csvW.AddLine csvD()
                csvD(IDX_TABLE_NAME_COLUMN_NUM) = ""
                csvD(IDX_NAME_COLUMN_NUM) = ""
                
            Next
        Next
SystemTableSkip:
    Next
End Sub


Private Function GetTableLinkAttribute(value As Long) As String
    If ((value And &H2000000) = &H2000000) Then
        GetTableLinkAttribute = TABLE_LINKED_ODBC
    ElseIf ((value And &H4000000) = &H4000000) Then
        GetTableLinkAttribute = TABLE_LINKED_NOT_ODBC
    Else
        GetTableLinkAttribute = TABLE_LOCAL
    End If
End Function

Private Function IsExclusiveTable(value As Long) As Boolean
        If ((value And &H200000) = &H200000) Then
            IsExclusiveTable = True
        Else
            IsExclusiveTable = False
        End If
End Function
Private Function IsSystemTable(value As Long) As Boolean
    If (((value And &H80000000) = &H80000000) Or value = 2) Then
        IsSystemTable = True
    Else
        IsSystemTable = False
    End If
End Function

Private Sub SetDbDataTypeToStr()
    Set m_dbDataTypeToStr = New Scripting.Dictionary
    With m_dbDataTypeToStr
        .Add 1, "dbBoolean"
        .Add 2, "dbByte"
        .Add 3, "dbInteger"
        .Add 4, "dbLong"
        .Add 5, "dbCurrency"
        .Add 6, "dbSingle"
        .Add 7, "dbDouble"
        .Add 8, "dbDate"
        .Add 9, "dbBinary"
        .Add 10, "dbText"
        .Add 11, "dbLongBinary"
        .Add 12, "dbMemo"
        .Add 15, "dbGUID"
        .Add 16, "dbBigInt"
        .Add 17, "dbVarBinary"
        .Add 18, "dbChar"
        .Add 19, "dbNumeric"
        .Add 20, "dbDecimal"
        .Add 21, "dbFloat"
        .Add 22, "dbTime"
        .Add 23, "dbTimeStamp"
    End With
End Sub
