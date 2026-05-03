VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IdxDefinitionExporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Implements IObjectDefinitionExporter

Const IDX_TABLE_NAME_COLUMN_NUM As Integer = 0
Const IDX_NAME_COLUMN_NUM As Integer = 1
Const IDX_FIELD_NAME_COLUMN_NUM As Integer = 2
Const IDX_SORT_COLUMN_NUM As Integer = 3

Private m_curData As CurrentData
Private m_database As DAO.database

Public Sub Class_Initialize()
    Set m_curData = CurrentData
    Set m_database = CurrentDb
End Sub

Public Sub IObjectDefinitionExporter_ExportToCsv(fullPath As String)
    Dim csvW As ICsvWriter
    Set csvW = New CsvWriter
    csvW.CreateFile (fullPath)
    IdxExport csvW
    csvW.Complete
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

Private Function IsSystemTable(value As Long) As Boolean
    If (((value And &H80000000) = &H80000000) Or value = 2) Then
        IsSystemTable = True
    Else
        IsSystemTable = False
    End If
End Function
