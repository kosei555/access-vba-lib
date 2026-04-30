VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IQueryManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Public Sub RegisterQuery(qName As String)
End Sub
Public Sub SetParam(qName As String, pName As String, pData As Variant)
End Sub
Public Sub ExecQuery(qName As String)
End Sub
Public Function GetQueryRecordSet(qName As String, Optional isDynaset As Boolean = False) As DAO.Recordset
End Function
Public Sub BeginTrans()
End Sub
Public Sub CommitTrans()
End Sub
Public Sub WExecQuery()
End Sub
Public Function WithTarget(qName As String) As IQueryManager
End Function
Public Sub WSetParam(pName As String, pData As Variant)
End Sub
Private Sub ShowErrorAndRollback(errorInfo As IErrBuf)
End Sub
Private Sub Rollback()
End Sub
